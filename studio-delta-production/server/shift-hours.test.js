const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-shift-"));
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

const order = db.upsertOrder({
  order_number: "S-SHIFT-1",
  status: "Ready for Welding",
  type: "Gate",
  category: "Driveway",
  product: "Slider",
  price_excl_vat: "100.00"
});

function openLog() {
  const sheet = book.getSheetByName("Production_Log");
  const last = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 13);
  const grid = sheet.getRange(1, 1, last, lastCol).getValues();
  for (let i = 1; i < grid.length; i++) {
    if (String(grid[i][1]) === "S-SHIFT-1" && !grid[i][6]) {
      return { sheet, row: i + 1, values: grid[i] };
    }
  }
  return null;
}

function writeMeta(next) {
  const hit = openLog();
  assert.ok(hit, "open log");
  hit.sheet.getRange(hit.row, 10, 1, 4).setValues([["", 0, "", JSON.stringify(next)]]);
  persistWorkbook();
}

function readMeta() {
  const hit = openLog();
  assert.ok(hit, "open log");
  return JSON.parse(String(hit.values[12] || "{}"));
}

(async function main() {
  const started = await callShopFunction("startOrder", [order.id, "Thabo", "Welding", [], "", false]);
  assert.strictEqual(started.success, true, JSON.stringify(started));

  writeMeta({
    pauses: [],
    batchId: "",
    batchShare: 1,
    batchSplitAt: null,
    entryType: "production",
    overtimeContinue: false
  });

  const lunch = await callShopFunction("enforceShiftHours", [new Date("2026-09-07T12:10:00+02:00")]);
  assert.strictEqual(lunch.kind, "lunch");
  assert.strictEqual(lunch.paused, 1, JSON.stringify(lunch));
  assert.strictEqual(readMeta().pauses[0].reason, "Lunch");

  const afterLunch = await callShopFunction("enforceShiftHours", [new Date("2026-09-07T12:35:00+02:00")]);
  assert.strictEqual(afterLunch.kind, "paid");
  assert.strictEqual(afterLunch.resumed, 1, JSON.stringify(afterLunch));
  assert.ok(readMeta().pauses[0].end, "lunch pause closed");

  const endShift = await callShopFunction("enforceShiftHours", [new Date("2026-09-07T16:00:00+02:00")]);
  assert.strictEqual(endShift.kind, "end");
  assert.strictEqual(endShift.paused, 1, JSON.stringify(endShift));
  assert.strictEqual(readMeta().pauses[readMeta().pauses.length - 1].reason, "End of shift");

  const resumed = await callShopFunction("workerResumeOrder", [order.id, "S-SHIFT-1", "Thabo", "", false]);
  assert.strictEqual(resumed.success, true, JSON.stringify(resumed));
  const afterResume = readMeta();
  afterResume.overtimeContinue = true;
  writeMeta(afterResume);

  const stillOt = await callShopFunction("enforceShiftHours", [new Date("2026-09-07T16:10:00+02:00")]);
  assert.strictEqual(stillOt.kind, "end");
  assert.strictEqual(stillOt.paused, 0, "overtime stays running: " + JSON.stringify(stillOt));
  const last = readMeta().pauses[readMeta().pauses.length - 1];
  assert.ok(last.end, "end-of-shift pause stayed closed");

  console.log("shift-hours.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
