const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-idle-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
process.env.WORK_LOCKS_DISABLED = "false";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const staff = require("./staff");
const { callShopFunction, clearShopCache } = require("./gas");
const { coerceRead } = require("./sheets");

initWorkbook();

staff.upsertUser({
  name: "Siya",
  access: "Admin",
  role: "Admin",
  password: "siya",
  seeDebtors: "Yes"
});
staff.upsertUser({
  name: "Site Manager",
  access: "Admin",
  role: "Manager",
  password: "mgr",
  seeDebtors: "Yes"
});
staff.upsertUser({
  name: "Uriah",
  access: "Production",
  role: "Assembler",
  password: "1234",
  tasks: ["Assembly"]
});
staff.upsertUser({
  name: "Willard",
  access: "Production",
  role: "Welder",
  password: "1234",
  tasks: ["Welding"]
});
staff.upsertUser({
  name: "Sipho",
  access: "Admin",
  role: "Production Manager",
  password: "sipho",
  seeDebtors: "Yes"
});

function todayStamp() {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Africa/Johannesburg",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).format(new Date());
}

function idleSheet() {
  return getBook().getSheetByName("Idle_Alerts");
}

function openHole(worker, dateCell) {
  const sheet = idleSheet();
  sheet.appendRow([
    dateCell,
    worker,
    "Assembler",
    new Date("2026-09-14T06:03:00.000Z"),
    new Date("2026-09-14T06:18:00.000Z"),
    "Open",
    "",
    "",
    ""
  ]);
  persistWorkbook();
}

(async function main() {
  const today = todayStamp();
  openHole("Uriah", today);

  const blocked = await callShopFunction("assignIndirectTask", [
    "Uriah", "Other", "Willard", "wrapping order S260200"
  ]);
  assert.strictEqual(blocked.success, false);
  assert.ok(/Siya or the Manager/i.test(blocked.message), JSON.stringify(blocked));

  const missingNote = await callShopFunction("assignIndirectTask", [
    "Uriah", "Other", "Siya", ""
  ]);
  assert.strictEqual(missingNote.success, false);

  const listed = await callShopFunction("getIdleWorkers", []);
  assert.ok((listed.workers || []).some((w) => w.worker === "Uriah"), JSON.stringify(listed));

  const saved = await callShopFunction("assignIndirectTask", [
    "Uriah", "Other", "Siya", "wrapping order S260200"
  ]);
  assert.strictEqual(saved.success, true, JSON.stringify(saved));
  assert.ok(String(saved.task).indexOf("wrapping order S260200") !== -1);

  const after = await callShopFunction("getIdleWorkers", []);
  assert.ok(!(after.workers || []).some((w) => w.worker === "Uriah"), "assigned idle hole must leave the list");

  const logs = getBook().getSheetByName("Production_Log").getDataRange().getValues();
  const indirect = logs.find((row) => String(row[1]) === "INDIRECT" && String(row[2]) === "Uriah");
  assert.ok(indirect, "Other task must write an indirect production log");
  assert.ok(String(indirect[4]).indexOf("wrapping order S260200") !== -1);

  clearShopCache();
  openHole("Uriah", coerceRead(today));
  const listedDate = await callShopFunction("getIdleWorkers", []);
  assert.ok((listedDate.workers || []).some((w) => w.worker === "Uriah"), "Date cells must still count as today");
  const savedDate = await callShopFunction("assignIndirectTask", [
    "Uriah", "Cleaning", "Site Manager", ""
  ]);
  assert.strictEqual(savedDate.success, true, JSON.stringify(savedDate));
  const gone = await callShopFunction("getIdleWorkers", []);
  assert.ok(!(gone.workers || []).some((w) => w.worker === "Uriah"));

  clearShopCache();
  openHole("Uriah", today);
  const listedPm = await callShopFunction("getIdleWorkers", []);
  const hole = (listedPm.workers || []).find((w) => w.worker === "Uriah");
  assert.ok(hole && hole.row, JSON.stringify(listedPm));
  const savedPm = await callShopFunction("assignIndirectTask", [
    "Uriah", "Other", "Sipho", "wrapping order S260200", hole.row
  ]);
  assert.strictEqual(savedPm.success, true, JSON.stringify(savedPm));
  const gonePm = await callShopFunction("getIdleWorkers", []);
  assert.ok(!(gonePm.workers || []).some((w) => w.worker === "Uriah"), "Production Manager must be able to assign idle tasks");

  const pauseAt = new Date("2026-09-15T10:15:00+02:00");
  const checkAt = new Date("2026-09-15T10:40:00+02:00");
  const liveStart = new Date("2026-09-15T12:30:00+02:00");
  const logsSheet = getBook().getSheetByName("Production_Log");
  logsSheet.appendRow([
    "log-willard-pause",
    "S-IDLE-1",
    "Willard",
    "Welding",
    "Welding",
    new Date("2026-09-15T07:45:00+02:00"),
    "",
    "",
    "",
    pauseAt,
    0,
    "No materials",
    JSON.stringify({
      pauses: [{ start: pauseAt.getTime(), end: null, reason: "No materials" }]
    })
  ]);
  persistWorkbook();
  clearShopCache();
  await callShopFunction("checkIdleWorkers", [checkAt]);
  const willardHole = (idleSheet().getDataRange().getValues() || []).find((row, i) => i > 0 && String(row[1]) === "Willard" && String(row[5] || "").toLowerCase() === "open");
  assert.ok(willardHole, "paused job must open an idle hole");
  const holeStart = new Date(willardHole[3]).getTime();
  assert.ok(Math.abs(holeStart - pauseAt.getTime()) < 60 * 1000, "Other hole starts at the pause, not the job start: " + willardHole[3]);

  logsSheet.appendRow([
    "log-willard-live",
    "S-IDLE-2",
    "Willard",
    "Welding",
    "Welding",
    liveStart,
    "",
    "",
    "",
    "",
    0,
    "",
    JSON.stringify({ pauses: [] })
  ]);
  persistWorkbook();
  clearShopCache();
  const savedLive = await callShopFunction("assignIndirectTask", [
    "Willard", "Other", "Siya", "wrapping order S260200"
  ]);
  assert.strictEqual(savedLive.success, true, JSON.stringify(savedLive));
  const willardIndirect = getBook().getSheetByName("Production_Log").getDataRange().getValues()
    .find((row) => String(row[1]) === "INDIRECT" && String(row[2]) === "Willard");
  assert.ok(willardIndirect, "Other must write a log");
  const otherEnd = willardIndirect[6] ? new Date(willardIndirect[6]).getTime() : 0;
  assert.ok(otherEnd, "Other over a live job must close");
  assert.ok(Math.abs(otherEnd - liveStart.getTime()) < 60 * 1000, "Other must stop when the next job starts, not keep counting: " + willardIndirect[6]);

  console.log("idle-assign.test.js ok");
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
