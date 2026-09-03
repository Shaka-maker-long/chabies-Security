const assert = require("assert");
const { Sheet, coerceRead, coerceWrite } = require("./sheets");
const { jsonSafe, ALLOWED } = require("./gas");

const book = { pendingHides: [] };
const sheet = new Sheet(book, { title: "Logs", sheetId: 1 });
sheet.grid = [
  ["id", "order", "worker"],
  ["1", "A-1", "Sipho"],
  ["2", "A-2", "Thabo"],
  ["3", "A-3", "Lerato"]
];
sheet.lastRow = 4;
sheet.lastCol = 3;
sheet.deleteRow(2);
assert.strictEqual(sheet.lastRow, 3);
assert.strictEqual(sheet.grid[1][1], "A-2");
assert.strictEqual(sheet.grid[2][1], "A-3");

sheet.getRange(2, 2, 1, 1).setValue("B-9");
assert.strictEqual(sheet.getRange(2, 2).getValue(), "B-9");
sheet.appendRow(["4", "A-4", "Naledi"]);
assert.strictEqual(sheet.lastRow, 4);
assert.strictEqual(sheet.getRange(4, 3).getValue(), "Naledi");

const serial = 45931; // ~2025-10-01
const asDate = coerceRead(serial);
assert.ok(asDate instanceof Date);
assert.ok(!isNaN(asDate.getTime()));
const written = coerceWrite(new Date("2026-08-27T07:45:00+02:00"));
assert.ok(String(written).indexOf("2026-08-27") === 0);

const safe = jsonSafe({ start: new Date("2026-08-27T07:45:00+02:00"), nested: [new Date(0)] });
assert.strictEqual(typeof safe.start, "number");
assert.strictEqual(safe.nested[0], 0);

[
  "verifyGlobalLogin", "pollFloor", "getFloorTaskCounts", "startOrder", "getAdminDashboardData",
  "getScheduleBoard", "generateWorkerSchedule", "checkIdleWorkers", "getTaskDuration"
].forEach((fn) => assert.ok(ALLOWED.has(fn), fn));
assert.ok(!ALLOWED.has("eval"));

console.log("sheets.test.js ok");
