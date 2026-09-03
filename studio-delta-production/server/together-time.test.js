const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-together-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const { callShopFunction } = require("./gas");

initWorkbook();
const book = getBook();
book.getSheetByName("Users").appendRow(["Thabo", "Welding", "1234", "Welding", "Production", "No"]);
persistWorkbook();

function order(num) {
  return db.upsertOrder({
    order_number: num,
    status: "Ready for Welding",
    type: "Gate",
    category: "Driveway",
    product: "Slider",
    price_excl_vat: "100.00"
  });
}

function openLogs(orderNum) {
  const sheet = book.getSheetByName("Production_Log");
  const last = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 13);
  const grid = sheet.getRange(1, 1, last, lastCol).getValues();
  return grid
    .map((row, i) => ({ row, i }))
    .filter((x) => x.i > 0 && String(x.row[1]) === orderNum && !x.row[6]);
}

(async function main() {
  const a = order("S-2001");
  const b = order("S-2002");
  const c = order("S-2003");
  const d = order("S-2004");

  const first = await callShopFunction("startOrder", [a.id, "Thabo", "Welding", [], "", false]);
    assert.strictEqual(first.success, true, JSON.stringify(first));

  const countsAfterStart = await callShopFunction("getFloorTaskCounts", []);
  assert.ok(countsAfterStart && countsAfterStart.Welding, "Home cards need Welding counts");
  assert.ok(countsAfterStart.Welding.active >= 1, "started order is Active: " + JSON.stringify(countsAfterStart.Welding));
  assert.ok(countsAfterStart.Welding.ready >= 3, "unstarted welding orders are Ready: " + JSON.stringify(countsAfterStart.Welding));

  const blocked = await callShopFunction("startOrder", [b.id, "Thabo", "Welding", [], "", false]);
  assert.ok(blocked.needsSwitchReason, "must ask work-together or switch: " + JSON.stringify(blocked));
  assert.ok(blocked.runningOrders && blocked.runningOrders.length, "running orders listed");

  const together = await callShopFunction("startOrder", [b.id, "Thabo", "Welding", [], "", true]);
  assert.strictEqual(together.success, true, JSON.stringify(together));

  const logsA = openLogs("S-2001");
  const logsB = openLogs("S-2002");
  assert.strictEqual(logsA.length, 1, "S-2001 still running together");
  assert.strictEqual(logsB.length, 1, "S-2002 running together");
  const metaA = JSON.parse(String(logsA[0].row[12] || "{}"));
  const metaB = JSON.parse(String(logsB[0].row[12] || "{}"));
  assert.ok(Number(metaA.batchShare) >= 2, "S-2001 time split: " + logsA[0].row[12]);
  assert.ok(Number(metaB.batchShare) >= 2, "S-2002 time split: " + logsB[0].row[12]);
  assert.strictEqual(String(metaA.batchId), String(metaB.batchId));

  const pollTogether = await callShopFunction("pollFloor", ["Welding", "Thabo"]);
  const cardA = pollTogether.orders.find((o) => String(o.order) === "S-2001");
  const cardB = pollTogether.orders.find((o) => String(o.order) === "S-2002");
  assert.ok(cardA && cardB, "both orders on floor");
  assert.ok(cardA.isBatched && cardB.isBatched, "together badge");
  assert.ok(cardA.batchShare >= 2 && cardB.batchShare >= 2, "batchShare on floor");
  assert.ok(!cardA.isPaused && !cardB.isPaused, "together does not pause");

  const finishA = await callShopFunction("finishOrder", [
    a.id, logsA[0].row[0], null, "", [], "Thabo", [], "S-2001", []
  ]);
  assert.strictEqual(finishA.success, true, JSON.stringify(finishA));
  const finishB = await callShopFunction("finishOrder", [
    b.id, logsB[0].row[0], null, "", [], "Thabo", [], "S-2002", []
  ]);
  assert.strictEqual(finishB.success, true, JSON.stringify(finishB));

  const closed = book.getSheetByName("Production_Log");
  const closedGrid = closed.getRange(1, 1, closed.getLastRow(), 13).getValues();
  const doneA = closedGrid.filter((r, i) => i > 0 && String(r[1]) === "S-2001" && r[6]).pop();
  const doneB = closedGrid.filter((r, i) => i > 0 && String(r[1]) === "S-2002" && r[6]).pop();
  assert.ok(doneA && doneB, "both together logs closed");
  assert.ok(Number(JSON.parse(String(doneA[12] || "{}")).batchShare) >= 2, "finished S-2001 keeps split share");
  assert.ok(Number(JSON.parse(String(doneB[12] || "{}")).batchShare) >= 2, "finished S-2002 keeps split share");

  const startC = await callShopFunction("startOrder", [c.id, "Thabo", "Welding", [], "", false]);
  assert.strictEqual(startC.success, true, JSON.stringify(startC));
  const switchToD = await callShopFunction("startOrder", [d.id, "Thabo", "Welding", [], "No materials", false]);
  assert.strictEqual(switchToD.success, true, JSON.stringify(switchToD));

  const pollSwitch = await callShopFunction("pollFloor", ["Welding", "Thabo"]);
  const cardC = pollSwitch.orders.find((o) => String(o.order) === "S-2003");
  const cardD = pollSwitch.orders.find((o) => String(o.order) === "S-2004");
  assert.ok(cardC && cardD);
  assert.ok(cardC.isPaused, "switch auto-pauses the first order");
  assert.ok(!cardD.isPaused, "new order is running");
  assert.ok(cardC.pausedAt, "paused countdown freezes");

  const countsAfterSwitch = await callShopFunction("getFloorTaskCounts", []);
  assert.ok(countsAfterSwitch.Welding.active >= 1, "running order stays Active: " + JSON.stringify(countsAfterSwitch.Welding));
  assert.ok(countsAfterSwitch.Welding.paused >= 1, "switched-away order is Paused: " + JSON.stringify(countsAfterSwitch.Welding));

  console.log("together-time.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
