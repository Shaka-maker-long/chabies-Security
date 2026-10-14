"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-align-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const staff = require("./staff");
const plan = require("./floor-planning");
const live = require("./schedule-live");

initWorkbook();

staff.upsertUser({
  name: "Willard",
  access: "Production",
  role: "Welding",
  password: "1234",
  tasks: ["Profile Cutting", "Tagging", "Plate Cutting", "Welding", "Grinding", "Assembly", "Quality Control"]
});
staff.setDurations([
  { product: "Air Chair", process: "Profile Cutting", hours: 1 },
  { product: "Air Chair", process: "Tagging", hours: 0.5 },
  { product: "Air Chair", process: "Plate Cutting", hours: 1 },
  { product: "Air Chair", process: "Welding", hours: 2 },
  { product: "Air Chair", process: "Grinding", hours: 1 },
  { product: "Air Chair", process: "Assembly", hours: 1 },
  { product: "Air Chair", process: "Final QC", hours: 0.5 }
]);

db.upsertOrder({
  order_number: "S-ALIGN-1",
  status: "Not Yet Started",
  product: "Air Chair",
  client_name: "Client"
});

const schedRows = db.listSchedule("2026-09-01", "2026-11-30");
const row = schedRows.find((r) => r.order_number === "S-ALIGN-1");
db.setScheduleCell(row.id, "2026-10-02", "LD", true, { skipPlan: true });

plan.save({ blocks: [], assignments: {} });
plan.scheduleSelected({
  orderIds: ["S-ALIGN-1"],
  assignments: {
    "S-ALIGN-1": {
      "Profile Cutting": "Willard",
      Tagging: "Willard",
      "Plate Cutting": "Willard",
      Welding: "Willard",
      Grinding: "Willard",
      Assembly: "Willard",
      "Final QC": "Willard"
    }
  },
  from: "2026-09-08T07:45:00+02:00"
});
persistWorkbook();

function planDayMap(orderId) {
  const journey = plan.buildJourney(plan.load().blocks || []);
  const orderJourney = journey.orders.find((o) => o.orderId === orderId);
  assert.ok(orderJourney, "planning journey has the order");
  const planDays = {};
  orderJourney.rows.forEach((r) => {
    const letter = plan.processCode(r.process);
    (r.days || []).forEach((iso) => {
      planDays[iso] = planDays[iso] ? planDays[iso] + "," + letter : letter;
    });
  });
  return planDays;
}

function assertMirrors(planDays, painted, label) {
  Object.keys(planDays).forEach((iso) => {
    const cell = painted.cells[iso] || "";
    planDays[iso].split(",").forEach((letter) => {
      assert.ok(String(cell).split(",").indexOf(letter) !== -1,
        label + ": schedule day " + iso + " should include Planning letter " + letter + " (cell=" + cell + ")");
    });
  });
  Object.keys(painted.cells || {}).forEach((iso) => {
    const cell = painted.cells[iso];
    if (!cell || cell === "LD" || cell === "LC" || cell === "C" || cell === "PD") return;
    assert.ok(planDays[iso], label + ": schedule has " + cell + " on " + iso + " but Planning has no plan that day");
  });
}

const planDays = planDayMap("S-ALIGN-1");
assert.ok(Object.keys(planDays).length > 0, "planning has plan days");
assert.ok(Object.values(planDays).join(" ").indexOf("CS") !== -1, "planning booked CS");

live.syncLiveScheduleCodes();
let painted = db.listSchedule("2026-09-01", "2026-11-30").find((r) => r.order_number === "S-ALIGN-1");
assertMirrors(planDays, painted, "full plan");

// Even after shop status moves past cutting, schedule still mirrors the Planning plan days
// (no hollow gap where CS/T were cleared).
db.upsertOrder({
  order_number: "S-ALIGN-1",
  status: "Ready for Welding",
  product: "Air Chair",
  client_name: "Client"
});
live.syncLiveScheduleCodes();
painted = db.listSchedule("2026-09-01", "2026-11-30").find((r) => r.order_number === "S-ALIGN-1");
assertMirrors(planDays, painted, "after status advanced");
assert.ok(Object.values(painted.cells || {}).some((c) => String(c).split(",").indexOf("CS") !== -1),
  "CS from Planning plan remains on the schedule after status advanced");

console.log("schedule-plan-align.test.js ok", { planDays: Object.keys(planDays).length });
