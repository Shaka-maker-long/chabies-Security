const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-duration-"));
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

const CONFIRM = { understood: true, highlights: [] };

function openLog(orderNum) {
  const sheet = book.getSheetByName("Production_Log");
  const last = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 13);
  const grid = sheet.getRange(1, 1, last, lastCol).getValues();
  for (let i = 1; i < grid.length; i++) {
    if (String(grid[i][1]) === orderNum && !grid[i][6]) {
      return { sheet, row: i + 1, values: grid[i] };
    }
  }
  return null;
}

function writeOpenLog(orderNum, start, meta) {
  const hit = openLog(orderNum);
  assert.ok(hit, "open log for " + orderNum);
  hit.sheet.getRange(hit.row, 6).setValue(start);
  hit.sheet.getRange(hit.row, 10, 1, 4).setValues([["", 0, "", JSON.stringify(meta)]]);
  persistWorkbook();
}

function durationRows() {
  const sheet = book.getSheetByName("Task_Durations");
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues()
    .filter((r) => String(r[0] || "").trim());
}

function seedDuration(product, process, minutes) {
  const sheet = book.getSheetByName("Task_Durations");
  sheet.appendRow([product, process, minutes]);
  persistWorkbook();
}

(async function main() {
  await callShopFunction("grantOvertime", ["Thabo", "", "Admin", "test"]);

  const lunchCross = await callShopFunction("getTaskTimeEstimate", ["2026-09-07T11:00:00+02:00", 120]);
  assert.strictEqual(Date.parse(lunchCross.etaAt), Date.parse("2026-09-07T13:30:00+02:00"), JSON.stringify(lunchCross));
  assert.ok(String(lunchCross.etaLabel).indexOf("13:30") !== -1, JSON.stringify(lunchCross));

  const threeAndHalf = await callShopFunction("getTaskTimeEstimate", ["2026-09-07T11:00:00+02:00", 210]);
  assert.strictEqual(Date.parse(threeAndHalf.etaAt), Date.parse("2026-09-07T15:00:00+02:00"), JSON.stringify(threeAndHalf));

  const noLunch = await callShopFunction("getTaskTimeEstimate", ["2026-09-07T08:00:00+02:00", 60]);
  assert.strictEqual(Date.parse(noLunch.etaAt), Date.parse("2026-09-07T09:00:00+02:00"), JSON.stringify(noLunch));

  const duringLunch = await callShopFunction("getTaskTimeEstimate", ["2026-09-07T12:10:00+02:00", 30]);
  assert.strictEqual(Date.parse(duringLunch.etaAt), Date.parse("2026-09-07T13:00:00+02:00"), JSON.stringify(duringLunch));

  const spoken = await callShopFunction("getTaskDuration", ["Slider", "Welding"]);
  assert.strictEqual(spoken.minutes, 0);
  assert.strictEqual(spoken.durationLabel, "");

  const none = db.upsertOrder({
    order_number: "S-DUR-1",
    status: "Ready for Welding",
    type: "Gate",
    product: "Slider",
    price_excl_vat: "100.00"
  });
  const briefEmpty = await callShopFunction("getOrderJobBrief", ["S-DUR-1", "Welding"]);
  assert.strictEqual(briefEmpty.targetMinutes, 0);
  assert.strictEqual(briefEmpty.durationLabel, "");

  const started = await callShopFunction("startOrder", [none.id, "Thabo", "Welding", [], "", false, null, CONFIRM]);
  assert.strictEqual(started.success, true, JSON.stringify(started));
  assert.strictEqual(started.targetMinutes, 0);
  const startedLog = openLog("S-DUR-1");
  const startedMeta = JSON.parse((startedLog && startedLog.values[12]) || "{}");
  const openPause = (startedMeta.pauses || []).some((p) => p && !p.end);
  assert.strictEqual(openPause, false, "missing task time must not auto-pause the job");

  const pauseStart = new Date(Date.now() - 90 * 60 * 1000);
  const pauseEnd = new Date(pauseStart.getTime() + 30 * 60 * 1000);
  writeOpenLog("S-DUR-1", pauseStart, {
    pauses: [{ start: pauseStart.toISOString(), end: pauseEnd.toISOString(), reason: "No materials" }],
    batchId: "",
    batchShare: 1,
    batchSplitAt: null,
    entryType: "production",
    overtimeContinue: true,
    targetMinutes: 0
  });

  const finished = await callShopFunction("finishOrder", [
    none.id, started.logId, null, "", [], "Thabo", [], "S-DUR-1", []
  ]);
  assert.strictEqual(finished.success, true, JSON.stringify(finished));
  assert.ok(finished.savedDuration, JSON.stringify(finished));
  assert.ok(finished.durationMins >= 55 && finished.durationMins <= 70, "pause must be ignored: " + JSON.stringify(finished));
  assert.ok(finished.targetMinutes >= 55 && finished.targetMinutes <= 70, JSON.stringify(finished));

  const saved = durationRows();
  assert.strictEqual(saved.length, 1, JSON.stringify(saved));
  assert.strictEqual(String(saved[0][0]), "Slider");
  assert.strictEqual(String(saved[0][1]), "Welding");
  assert.ok(Number(saved[0][2]) >= 0.9 && Number(saved[0][2]) <= 1.2, "first finish must save hours: " + JSON.stringify(saved));

  const briefNow = await callShopFunction("getOrderJobBrief", ["S-DUR-1", "Welding"]);
  assert.ok(briefNow.targetMinutes >= 55, JSON.stringify(briefNow));
  assert.ok(/hour|minute/i.test(briefNow.durationLabel), JSON.stringify(briefNow));

  const again = db.upsertOrder({
    order_number: "S-DUR-2",
    status: "Ready for Welding",
    type: "Gate",
    product: "Slider",
    price_excl_vat: "100.00"
  });
  const start2 = await callShopFunction("startOrder", [again.id, "Thabo", "Welding", [], "", false, null, CONFIRM]);
  assert.ok(start2.targetMinutes >= 55, JSON.stringify(start2));
  assert.ok(start2.durationLabel, JSON.stringify(start2));
  writeOpenLog("S-DUR-2", new Date(Date.now() - 20 * 60 * 1000), {
    pauses: [],
    batchId: "",
    batchShare: 1,
    batchSplitAt: null,
    entryType: "production",
    overtimeContinue: true,
    targetMinutes: start2.targetMinutes
  });
  const finish2 = await callShopFunction("finishOrder", [
    again.id, start2.logId, null, "", [], "Thabo", [], "S-DUR-2", []
  ]);
  assert.strictEqual(finish2.success, true, JSON.stringify(finish2));
  assert.strictEqual(finish2.savedDuration, false, "must not overwrite a saved duration");
  assert.strictEqual(finish2.onTime, true, JSON.stringify(finish2));
  assert.strictEqual(durationRows().length, 1, "still one duration row");

  seedDuration("Talitha", "Welding", 0.5);
  const late = db.upsertOrder({
    order_number: "S-DUR-3",
    status: "Ready for Welding",
    type: "Gate",
    product: "Talitha",
    price_excl_vat: "100.00"
  });
  const start3 = await callShopFunction("startOrder", [late.id, "Thabo", "Welding", [], "", false, null, CONFIRM]);
  assert.strictEqual(start3.targetMinutes, 30, JSON.stringify(start3));
  assert.strictEqual(start3.durationLabel, "30 minutes");
  writeOpenLog("S-DUR-3", new Date(Date.now() - 80 * 60 * 1000), {
    pauses: [],
    batchId: "",
    batchShare: 1,
    batchSplitAt: null,
    entryType: "production",
    overtimeContinue: true,
    targetMinutes: 30
  });
  const finish3 = await callShopFunction("finishOrder", [
    late.id, start3.logId, null, "", [], "Thabo", [], "S-DUR-3", []
  ]);
  assert.strictEqual(finish3.success, true, JSON.stringify(finish3));
  assert.strictEqual(finish3.overtime, true, JSON.stringify(finish3));
  assert.strictEqual(finish3.onTime, false);

  seedDuration("Long Gate", "Welding", 3.5);
  const long = db.upsertOrder({
    order_number: "S-DUR-4",
    status: "Ready for Welding",
    type: "Gate",
    product: "Long Gate",
    price_excl_vat: "100.00"
  });
  const start4 = await callShopFunction("startOrder", [long.id, "Thabo", "Welding", [], "", false, null, CONFIRM]);
  assert.strictEqual(start4.targetMinutes, 210, JSON.stringify(start4));
  assert.strictEqual(start4.durationLabel, "3 hours 30 minutes");
  assert.ok(start4.etaLabel, JSON.stringify(start4));
  assert.ok(start4.etaAt, JSON.stringify(start4));

  const mine = await callShopFunction("getMyCompletedWork", ["Thabo"]);
  assert.ok(mine.items && mine.items.length >= 3, JSON.stringify(mine));
  const first = mine.items.find((x) => x.order === "S-DUR-1");
  const second = mine.items.find((x) => x.order === "S-DUR-2");
  const third = mine.items.find((x) => x.order === "S-DUR-3");
  assert.ok(first && first.onTime, JSON.stringify(first));
  assert.ok(second && second.onTime, JSON.stringify(second));
  assert.ok(third && third.overtime, JSON.stringify(third));

  console.log("task-duration.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
