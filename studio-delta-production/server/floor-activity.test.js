"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-floor-act-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const staff = require("./staff");
const plan = require("./floor-planning");
const { callShopFunction, clearShopCache } = require("./gas");

initWorkbook();
staff.upsertUser({
  name: "Willard",
  access: "Production",
  role: "Welding",
  password: "1234",
  tasks: ["Welding"]
});
staff.setDurations([
  { product: "Air Chair", process: "Welding", hours: 1 }
]);

db.upsertOrder({
  order_number: "S260801",
  status: "Ready for Welding",
  product: "Air Chair",
  client_name: "Late"
});
db.upsertOrder({
  order_number: "S260802",
  status: "Welding",
  product: "Air Chair",
  client_name: "Early"
});

const rows = db.listSchedule("2026-09-21", "2026-10-02");
const late = rows.find((r) => r.order_number === "S260801");
const early = rows.find((r) => r.order_number === "S260802");
db.setScheduleCell(late.id, "2026-09-30", "LC", true, { skipPlan: true });
db.setScheduleCell(early.id, "2026-09-23", "LD", true, { skipPlan: true });

plan.save({ blocks: [], assignments: {} });
plan.scheduleSelected({
  orderIds: ["S260801", "S260802"],
  assignments: {
    S260801: { Welding: "Willard" },
    S260802: { Welding: "Willard" }
  },
  from: "2026-09-08T07:45:00+02:00"
});
persistWorkbook();
clearShopCache();

(async function main() {
  const poll = await callShopFunction("pollFloor", ["Welding", "Willard"]);
  const orders = poll.orders || [];
  const ids = orders.map((o) => String(o.order));
  assert.ok(ids.indexOf("S260802") !== -1, "early due order stays on Welding");
  assert.ok(ids.indexOf("S260801") !== -1, "later due order stays on Welding");
  assert.ok(ids.indexOf("S260802") < ids.indexOf("S260801"), "due-first order is listed first: " + ids.join(","));
  const earlyCard = orders.find((o) => String(o.order) === "S260802");
  const lateCard = orders.find((o) => String(o.order) === "S260801");
  assert.ok(!earlyCard.assigned, "Ready/Welding without a live clock stays off In Progress");
  assert.ok(!earlyCard.logId, "planned weld does not open a Production_Log clock");
  assert.ok(!lateCard.assigned, "Ready for Welding stays on Available");
  assert.ok(!lateCard.logId, "Ready for Welding has no live clock");
  assert.strictEqual(earlyCard.plannedWorker, "Willard");
  assert.strictEqual(earlyCard.delivery_code, "LD");
  console.log("floor-activity.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
