"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-close-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const staff = require("./staff");
const closing = require("./closing-plan");

initWorkbook();

staff.upsertUser({
  name: "Sam",
  access: "Production",
  role: "Cutter",
  password: "1234",
  tasks: ["Profile Cutting", "Tagging", "Plate Cutting", "Welding", "Grinding", "Assembly", "Quality Control"]
});
staff.setDurations([
  { product: "Air Chair", process: "Profile Cutting", hours: 1 },
  { product: "Air Chair", process: "Tagging", hours: 1 },
  { product: "Air Chair", process: "Plate Cutting", hours: 1 },
  { product: "Air Chair", process: "Welding", hours: 8 },
  { product: "Air Chair", process: "Grinding", hours: 1 },
  { product: "Air Chair", process: "Pre-Powder Coating", hours: 1 },
  { product: "Air Chair", process: "Assembly", hours: 1 },
  { product: "Air Chair", process: "Final QC", hours: 1 }
]);

db.upsertOrder({
  order_number: "S269001",
  status: "Delivered",
  product: "Air Chair",
  client_name: "Done",
  price_incl_vat: "1000.00",
  amount_paid: "1000.00"
});
db.upsertOrder({
  order_number: "S269002",
  status: "Ready for Steelwork",
  product: "Air Chair",
  client_name: "Open",
  price_incl_vat: "1000.00",
  amount_paid: "1000.00"
});

const book = getBook();
const log = book.getSheetByName("Production_Log");
function addClock(order, worker, process, start, end) {
  log.appendRow([
    "L" + Math.random().toString(16).slice(2, 8),
    order,
    worker,
    process,
    process,
    start,
    end,
    "OK",
    "",
    "",
    0,
    "",
    "{}"
  ]);
}
addClock("S269001", "Sam", "Welding", "2026-09-01T08:00:00+02:00", "2026-09-01T10:00:00+02:00");
addClock("S269001", "Sam", "Welding", "2026-09-02T08:00:00+02:00", "2026-09-02T09:30:00+02:00");
addClock("S269001", "Sam", "Profile Cutting", "2026-08-31T08:00:00+02:00", "2026-08-31T08:45:00+02:00");
persistWorkbook();

assert.strictEqual(closing.median([10, 30, 20]), 20);
assert.ok(closing.netMinutes("2026-09-01T08:00:00+02:00", "2026-09-01T10:00:00+02:00", []) >= 119);

const emptyDate = closing.build({ now: "2026-11-09" });
assert.strictEqual(emptyDate.openOrderCount, 1);
assert.ok(emptyDate.clockCount >= 2);
assert.ok(/last shop day/i.test(emptyDate.verdict), emptyDate.verdict);

closing.saveSettings({ closingDate: "2027-06-01", weekdayOtHours: 2, saturday: true });
const far = closing.build({ now: "2026-11-09", closingDate: "2027-06-01" });
assert.strictEqual(far.settings.closingDate, "2027-06-01");
assert.ok(far.jobs.length === 1);
const weld = far.jobs[0].steps.find((s) => s.process === "Welding");
assert.ok(weld, JSON.stringify(far.jobs[0]));
assert.strictEqual(weld.worker, "Sam");
assert.ok(weld.source.indexOf("person") !== -1 || weld.samples >= 1, weld.source);
assert.ok(far.canFinishPaid, far.verdict);
assert.ok(far.learned.some((row) => row.process === "Welding" && row.jobs >= 2));
const samDay = (far.daily || []).find((row) => row.worker === "Sam");
assert.ok(samDay, "Sam has a day-to-day rate from clocks, not a month average");
assert.ok(samDay.dayCount >= 2);
assert.ok(samDay.typicalMin > 0 && samDay.typicalMin < 7.5 * 60);
assert.ok(far.note.indexOf("median") !== -1);

const tight = closing.build({ now: "2026-11-09", closingDate: "2026-11-10", settings: { weekdayOtHours: 2, saturday: true } });
assert.ok(tight.openOrderCount === 1);
assert.ok(!tight.canFinishPaid, "one day cannot finish a full Air Chair: " + tight.verdict);
assert.ok(tight.overtime.length >= 1, "should suggest overtime");

console.log("closing-plan.test.js ok");
