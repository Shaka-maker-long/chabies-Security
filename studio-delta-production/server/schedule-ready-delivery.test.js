"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-rfd-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const db = require("./db");
const plan = require("./floor-planning");
const live = require("./schedule-live");
const sched = require("./office-schedule");

initWorkbook();

db.upsertOrder({
  order_number: "S260226 C",
  status: "Ready for Delivery",
  product: "Penelope Display Cabinet",
  type: "Standard",
  category: "Cabinet",
  province: "Western Cape",
  payment_date: "03/09/2026"
});

const rows = db.listSchedule("2026-09-01", "2026-10-10");
const row = rows.find((r) => r.order_number === "S260226 C");
assert.ok(row, "schedule row exists");

// Stale auto letters like the live bug: A then later PQC on a Ready for Delivery job.
db.setScheduleCell(row.id, "2026-09-22", "LC", true, { skipPlan: true, source: "manual" });
db.setScheduleCell(row.id, "2026-09-08", "A", true, { skipPlan: true, source: "auto" });
db.setScheduleCell(row.id, "2026-09-09", "PQC", true, { skipPlan: true, source: "auto" });

// Leftover planning blocks must not keep painting.
plan.save({
  blocks: [
    {
      id: "b1",
      orderId: "S260226 C",
      product: "Penelope Display Cabinet",
      process: "Assembly",
      workerId: "Admire",
      workerName: "Admire",
      start: "2026-09-08T07:45:00+02:00",
      end: "2026-09-08T10:00:00+02:00"
    },
    {
      id: "b2",
      orderId: "S260226 C",
      product: "Penelope Display Cabinet",
      process: "Pre-Powder Coating",
      workerId: "Siya",
      workerName: "Siya",
      start: "2026-09-09T07:45:00+02:00",
      end: "2026-09-09T09:00:00+02:00"
    }
  ],
  assignments: {}
});

const sync = live.syncLiveScheduleCodes();
assert.ok(sync.cleared >= 2, "stale auto A and PQC are cleared");

const after = db.listSchedule("2026-09-01", "2026-10-10").find((r) => r.order_number === "S260226 C");
assert.strictEqual(after.cells["2026-09-22"], "LC", "LC stays");
assert.ok(!after.cells["2026-09-08"], "Assembly letter gone on Ready for Delivery");
assert.ok(!after.cells["2026-09-09"], "PQC letter gone on Ready for Delivery");
assert.ok(!(plan.load().blocks || []).some((b) => b.orderId === "S260226 C"),
  "leftover planning blocks for Ready for Delivery are dropped");

// Chain day order: PQC must not sit before an earlier CS day is respected —
// if CS is Mon and someone offers PQC on a prior day, drop PQC.
const marks = new Map([
  ["S1|2026-09-08", "CS"],
  ["S1|2026-09-07", "PQC"],
  ["S1|2026-09-09", "A"]
]);
live.enforceChainDayOrder(marks);
assert.ok(!marks.has("S1|2026-09-07"), "PQC before CS day is dropped");
assert.strictEqual(marks.get("S1|2026-09-08"), "CS");
assert.strictEqual(marks.get("S1|2026-09-09"), "A");
assert.deepStrictEqual(sched.AUTO_SCHEDULE_ORDER.indexOf("PQC") < sched.AUTO_SCHEDULE_ORDER.indexOf("A"), true);

console.log("schedule-ready-delivery.test.js ok");
