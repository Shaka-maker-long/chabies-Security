const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-wb-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook, workbookPath } = require("./workbook-store");
const db = require("./db");
const { callShopFunction } = require("./gas");

initWorkbook();
const book = getBook();
book.getSheetByName("Users").appendRow(["Sipho", "Profile Cutting", "1234", "Profile Cutting"]);
book.getSheetByName("Users").appendRow(["Lerato", "Assembly", "1234", "Assembly"]);
persistWorkbook();

const saved = db.upsertOrder({
  order_number: "S-1001",
  status: "Not Yet Started",
  type: "Gate",
  category: "Driveway",
  product: "Slider",
  client_name: "Test Client",
  price_excl_vat: "1000.00"
});
assert.ok(saved.id >= 2);
assert.strictEqual(saved.price_incl_vat, "1150.00");

const assemblyOrder = db.upsertOrder({
  order_number: "S-1002",
  status: "Ready for Assembly",
  type: "Gate",
  category: "Driveway",
  product: "Slider",
  price_excl_vat: "500.00"
});

(async function main() {
  const login = await callShopFunction("verifyGlobalLogin", ["Sipho", "1234"]);
  assert.strictEqual(login.success, true, JSON.stringify(login));
  assert.ok(login.tasks.indexOf("Profile Cutting") !== -1);

  const start = await callShopFunction("startOrder", [saved.id, "Sipho", "Profile Cutting", [], "", false]);
  assert.strictEqual(start.success, true, JSON.stringify(start));
  assert.strictEqual(start.newStatus, "Profile Cutting");
  assert.ok(start.logId);

  const afterStart = db.listOrders().find((o) => o.order_number === "S-1001");
  assert.ok(afterStart, "office listOrders must see the floor order");
  assert.strictEqual(afterStart.status, "Profile Cutting");
  assert.strictEqual(afterStart.assigned_operator, "Sipho");

  const pause = await callShopFunction("workerPauseOrder", [afterStart.id, "S-1001", "Sipho", "No materials"]);
  assert.strictEqual(pause.success, true, JSON.stringify(pause));

  const logSheet = book.getSheetByName("Production_Log");
  const logGrid = logSheet.getRange(1, 1, logSheet.getLastRow(), logSheet.getLastColumn()).getValues();
  const openLog = logGrid.find((row, i) => i > 0 && String(row[1]) === "S-1001");
  assert.ok(openLog, "production log row");
  const meta = JSON.parse(String(openLog[12] || "{}"));
  assert.ok(meta.pauses && meta.pauses.length, "pause recorded in log meta");

  const resume = await callShopFunction("workerResumeOrder", [afterStart.id, "S-1001", "Sipho", "", false]);
  assert.strictEqual(resume.success, true, JSON.stringify(resume));

  const finish = await callShopFunction("finishOrder", [
    afterStart.id,
    start.logId,
    null,
    "",
    [],
    "Sipho",
    [{ type: "40x40", size: "6m", category: "Square" }],
    "S-1001",
    []
  ]);
  assert.strictEqual(finish.success, true, JSON.stringify(finish));

  const steel = book.getSheetByName("Steel_Usage");
  assert.ok(steel.getLastRow() >= 2, "steel usage row");
  const steelRow = steel.getRange(2, 1, 1, 6).getValues()[0];
  assert.strictEqual(String(steelRow[1]), "S-1001");
  assert.ok(String(steelRow[4]).indexOf("40x40") !== -1);

  const afterFinish = db.listOrders().find((o) => o.order_number === "S-1001");
  assert.strictEqual(afterFinish.status, "Ready for Tagging");
  assert.strictEqual(afterFinish.assigned_operator, "");

  const asmStart = await callShopFunction("startOrder", [assemblyOrder.id, "Lerato", "Assembly", [], "", false]);
  assert.strictEqual(asmStart.success, true, JSON.stringify(asmStart));
  const asmFinish = await callShopFunction("finishOrder", [
    assemblyOrder.id,
    asmStart.logId,
    null,
    "",
    [],
    "Lerato",
    [],
    "S-1002",
    [{ type: "18mm shutter", size: "1", category: "Board" }]
  ]);
  assert.strictEqual(asmFinish.success, true, JSON.stringify(asmFinish));
  const boards = book.getSheetByName("Backboard_Usage");
  assert.ok(boards.getLastRow() >= 2, "backboard usage row");

  const paid = db.recordPayment("S-1001", "150", "deposit");
  assert.strictEqual(paid.paid, "150.00");
  assert.ok(Number(paid.owing) > 0);
  const debtors = db.listDebtors();
  assert.ok(debtors.some((o) => o.order_number === "S-1001"));

  persistWorkbook();
  const raw = JSON.parse(fs.readFileSync(workbookPath(), "utf8"));
  assert.ok(raw.sheets.ORDERS);
  assert.ok(raw.sheets.Production_Log.grid.length >= 3);
  assert.ok(raw.sheets.Steel_Usage.grid.length >= 2);
  assert.ok(raw.sheets.Backboard_Usage.grid.length >= 2);
  const orderRow = raw.sheets.ORDERS.grid.find((r) => String(r[1]) === "S-1001");
  assert.ok(orderRow);
  assert.strictEqual(String(orderRow[2]), "Ready for Tagging");

  console.log("railway-db.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
