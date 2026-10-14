"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-siya-del-"));
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
const { callShopFunction, clearShopCache } = require("./gas");

initWorkbook();
staff.upsertUser({
  name: "Siya",
  access: "Admin",
  role: "Production Manager",
  password: "s",
  tasks: ["Quality Control"]
});
staff.upsertUser({
  name: "Willard",
  access: "Production",
  role: "Welding",
  password: "1234",
  tasks: ["Profile Cutting", "Tagging", "Welding", "Grinding", "Assembly", "Quality Control"]
});
staff.setDurations([
  { product: "Air Chair", process: "Profile Cutting", hours: 1 },
  { product: "Air Chair", process: "Tagging", hours: 0.5 },
  { product: "Air Chair", process: "Welding", hours: 2 },
  { product: "Air Chair", process: "Grinding", hours: 1 },
  { product: "Air Chair", process: "Assembly", hours: 1 }
]);

const outOrder = db.upsertOrder({
  order_number: "S-DEL-1",
  status: "Out for Delivery",
  product: "Air Chair",
  assigned_operator: "Willard",
  client_name: "Client"
});
const weldOrder = db.upsertOrder({
  order_number: "S-DEL-2",
  status: "Not Yet Started",
  product: "Air Chair",
  client_name: "Client"
});

const rows = db.listSchedule("2026-09-21", "2026-10-10");
const schedRow = rows.find((r) => r.order_number === "S-DEL-2");
assert.ok(schedRow, "schedule row exists");
db.setScheduleCell(schedRow.id, "2026-09-30", "LD", true, { skipPlan: true });

plan.save({ blocks: [], assignments: {} });
plan.scheduleSelected({
  orderIds: ["S-DEL-2"],
  assignments: {
    "S-DEL-2": {
      "Profile Cutting": "Willard",
      Tagging: "Willard",
      Welding: "Willard",
      Grinding: "Willard",
      Assembly: "Willard"
    }
  },
  from: "2026-09-08T07:45:00+02:00"
});
persistWorkbook();
clearShopCache();

(async function main() {
  const assigned = await callShopFunction("getFloorTaskCounts", []);
  assert.ok(assigned.Delivery, "Delivery pile is counted");
  const after = db.listOrders().find((o) => o.order_number === "S-DEL-1");
  assert.strictEqual(String(after.assigned_operator || ""), "Willard", "counts do not steal Out for Delivery for Siya");

  const sync = live.syncLiveScheduleCodes();
  assert.ok(sync.written >= 0);
  const painted = db.listSchedule("2026-09-08", "2026-10-10").find((r) => r.order_number === "S-DEL-2");
  const letters = Object.values(painted.cells || {});
  assert.ok(letters.some((c) => /^(CS|T|P|W|G|PQC|PC|U|A|FQC)(,(CS|T|P|W|G|PQC|PC|U|A|FQC))*$/.test(c) || c === "LD"),
    "auto-plan / shop letters paint onto the schedule: " + letters.join(","));
  assert.ok(letters.indexOf("LD") !== -1, "LD stays on the schedule");
  assert.ok(letters.every((c) => {
    const parts = String(c).split(",");
    return parts.indexOf("PD") === -1 && parts.indexOf("C") === -1;
  }), "auto-plan never writes PD or Courier C: " + letters.join(","));
  assert.ok(letters.some((c) => c === "CS" || String(c).split(",").indexOf("CS") !== -1),
    "Profile Cutting CS appears on the schedule: " + letters.join(","));
  assert.ok(letters.some((c) => c === "W" || c === "G" || c === "A" || c.indexOf("W") !== -1 || c.indexOf(",") !== -1),
    "planned station letters appear: " + letters.join(","));

  // Chain order: first CS day must be on or before first T day.
  const days = Object.keys(painted.cells || {}).sort();
  const firstCs = days.find((d) => String(painted.cells[d]).split(",").indexOf("CS") !== -1) || "";
  const firstT = days.find((d) => String(painted.cells[d]).split(",").indexOf("T") !== -1) || "";
  if (firstCs && firstT) {
    assert.ok(firstCs <= firstT, "CS must not come after T: CS=" + firstCs + " T=" + firstT);
  }

  // Manual letter must not be wiped by auto sync.
  const day = Object.keys(painted.cells).find((d) => painted.cells[d] === "M" || painted.cells[d] === "Gr" || painted.cells[d] === "A")
    || "2026-09-09";
  db.setScheduleCell(painted.id, "2026-09-11", "U", true, { skipPlan: true, source: "manual" });
  db.setScheduleCell(painted.id, "2026-09-12", "C", true, { skipPlan: true, source: "manual" });
  live.syncLiveScheduleCodes();
  const again = db.listSchedule("2026-09-08", "2026-10-10").find((r) => r.order_number === "S-DEL-2");
  assert.strictEqual(again.cells["2026-09-11"], "U", "hand-typed Upholstery stays");
  assert.strictEqual(again.cells["2026-09-12"], "C", "hand-typed Courier stays");
  assert.strictEqual(again.cells["2026-09-30"], "LD", "LD is never cleared");

  // Ready for Delivery / Delivered orders are not auto-planned and lose stale production letters.
  db.upsertOrder(Object.assign({}, weldOrder, { status: "Ready for Delivery" }));
  const readyPlan = plan.autoPlanFromDeliveries({ from: "2026-09-08T07:45:00+02:00" });
  assert.ok((readyPlan.orders || 0) === 0 || (readyPlan.clearedFinished || 0) >= 1,
    "Ready for Delivery orders are not re-planned");
  live.syncLiveScheduleCodes();
  const afterReady = db.listSchedule("2026-09-08", "2026-10-10").find((r) => r.order_number === "S-DEL-2");
  assert.strictEqual(afterReady.cells["2026-09-30"], "LD", "LD survives after Ready for Delivery");
  assert.strictEqual(afterReady.cells["2026-09-12"], "C", "Courier survives after Ready for Delivery");
  assert.strictEqual(afterReady.cells["2026-09-11"], "U", "manual U stays");
  const prodLetters = Object.keys(afterReady.cells || {}).filter((d) => {
    const v = afterReady.cells[d];
    return v && v !== "LD" && v !== "C" && v !== "U";
  });
  assert.strictEqual(prodLetters.length, 0, "auto production letters clear once Ready for Delivery: " + prodLetters.map((d) => afterReady.cells[d]).join(","));

  db.upsertOrder(Object.assign({}, weldOrder, { status: "Delivered" }));
  const deliveredPlan = plan.autoPlanFromDeliveries({ from: "2026-09-08T07:45:00+02:00" });
  assert.ok((deliveredPlan.orders || 0) === 0 || (deliveredPlan.clearedFinished || 0) >= 1,
    "delivered orders are not re-planned");

  console.log("delivery-siya-schedule.test.js ok", { outOrder: outOrder.order_number, weldOrder: weldOrder.order_number, day });
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
