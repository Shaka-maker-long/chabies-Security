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
  tasks: ["Welding", "Grinding", "Assembly", "Quality Control"]
});
staff.setDurations([
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
  status: "Ready for Welding",
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
    "S-DEL-2": { Welding: "Willard", Grinding: "Willard", Assembly: "Willard" }
  },
  from: "2026-09-08T07:45:00+02:00"
});
persistWorkbook();
clearShopCache();

(async function main() {
  const assigned = await callShopFunction("getFloorTaskCounts", []);
  assert.ok(assigned.Delivery, "Delivery pile is counted");
  const after = db.listOrders().find((o) => o.order_number === "S-DEL-1");
  assert.strictEqual(String(after.assigned_operator || "").toLowerCase(), "siya", "Out for Delivery is assigned to Siya");

  const sync = live.syncLiveScheduleCodes();
  assert.ok(sync.written >= 0);
  const painted = db.listSchedule("2026-09-08", "2026-10-10").find((r) => r.order_number === "S-DEL-2");
  const letters = Object.values(painted.cells || {});
  assert.ok(letters.indexOf("LD") !== -1, "LD stays on the schedule");
  assert.ok(letters.some((c) => c === "M" || c === "Gr" || c === "A" || c === "CS" || c === "PD"),
    "auto-plan / shop letters paint onto the schedule: " + letters.join(","));
  assert.ok(!letters.some((c) => c === "LD" && Object.keys(painted.cells).filter((d) => painted.cells[d] === "LD").length > 1));

  // Manual letter must not be wiped by auto sync.
  const day = Object.keys(painted.cells).find((d) => painted.cells[d] === "M" || painted.cells[d] === "Gr" || painted.cells[d] === "A")
    || "2026-09-09";
  db.setScheduleCell(painted.id, "2026-09-11", "U", true, { skipPlan: true, source: "manual" });
  live.syncLiveScheduleCodes();
  const again = db.listSchedule("2026-09-08", "2026-10-10").find((r) => r.order_number === "S-DEL-2");
  assert.strictEqual(again.cells["2026-09-11"], "U", "hand-typed Upholstery stays");

  console.log("delivery-siya-schedule.test.js ok", { outOrder: outOrder.order_number, weldOrder: weldOrder.order_number, day });
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
