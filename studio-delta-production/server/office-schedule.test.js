const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-sched-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const db = require("./db");
const sched = require("./office-schedule");

initWorkbook();

assert.strictEqual(sched.isoWeekInfo("2026-09-21").week, 39);
assert.strictEqual(sched.isoWeekInfo("2026-09-21").year, 2026);
assert.strictEqual(sched.weekKey(sched.isoWeekInfo("2026-09-21")), "2026-W39");
assert.strictEqual(sched.mondayOf("2026-09-23"), "2026-09-21");
assert.strictEqual(sched.weekdayLong("2026-09-21"), "Monday");
assert.ok(sched.formatDayLabel("2026-09-21").indexOf("Monday") !== -1);
assert.strictEqual(sched.gridWeekLabel("2026-09-21"), "Week39");
assert.deepStrictEqual(sched.workdays("2026-09-21", 5), [
  "2026-09-21", "2026-09-22", "2026-09-23", "2026-09-24", "2026-09-25"
]);
assert.ok(sched.DELIVERY_CODES.indexOf("LD") !== -1);
assert.ok(sched.DELIVERY_CODES.indexOf("LC") !== -1);
assert.ok(sched.SCHEDULE_CODES.some((c) => c.code === "LD" && c.label === "Latest Delivery"));
assert.ok(sched.SCHEDULE_CODES.some((c) => c.code === "P" && c.label === "Photograpy"));

db.upsertOrder({
  order_number: "S260186",
  type: "Standard",
  category: "Cabinet",
  product: "Vivienne Arched Cabinet",
  province: "Gauteng",
  payment_date: "01/09/2026",
  status: "Not Yet Started"
});
db.upsertOrder({
  order_number: "S260207",
  type: "Custom",
  category: "Sideboard",
  product: "Violet Sideboard 3-Door",
  province: "Western Cape",
  status: "At Couriers"
});

const synced = db.syncScheduleFromOrders();
assert.ok(synced.synced >= 2);
const rows = db.listSchedule("2026-09-21", "2026-09-25");
assert.ok(rows.some((r) => r.order_number === "S260186"));
assert.ok(rows.some((r) => r.order_number === "S260207"));
const cabinet = rows.find((r) => r.order_number === "S260186");
assert.strictEqual(cabinet.item_type, "Standard");
assert.strictEqual(cabinet.product, "Vivienne Arched Cabinet");
assert.strictEqual(cabinet.category, "Cabinet");

db.setScheduleCell(cabinet.id, "2026-09-21", "LD");
const side = rows.find((r) => r.order_number === "S260207");
db.setScheduleCell(side.id, "2026-09-22", "LC");
db.setScheduleCell(side.id, "2026-09-22", "LC");
db.setScheduleCell(cabinet.id, "2026-09-24", "QC");

const delivery = db.listDeliveryItems();
assert.strictEqual(delivery.length, 2);
assert.ok(delivery.some((i) => i.order_number === "S260186" && i.code === "LD" && i.week === 39));
assert.ok(delivery.some((i) => i.order_number === "S260207" && i.code === "LC" && i.weekday === "Tuesday"));
assert.ok(!delivery.some((i) => i.code === "QC"));

const again = db.upsertScheduleRow({
  order_number: "S260186",
  courier: "CAMPOS",
  waybill: "WB-1"
});
assert.strictEqual(again.id, cabinet.id);
assert.strictEqual(again.courier, "CAMPOS");

db.upsertOrder({
  order_number: "S260186",
  type: "Standard",
  category: "Cabinet",
  product: "Vivienne Arched Cabinet — oak",
  province: "Gauteng",
  status: "In production"
});
const refreshed = db.listSchedule("2026-09-21", "2026-09-25").find((r) => r.order_number === "S260186");
assert.strictEqual(refreshed.product, "Vivienne Arched Cabinet — oak");
assert.strictEqual(refreshed.status, "In production");
assert.strictEqual(refreshed.courier, "CAMPOS");
assert.strictEqual(refreshed.cells["2026-09-21"], "LD");

db.deleteOrder("S260207");
assert.ok(!db.listSchedule("2026-09-21", "2026-09-25").some((r) => r.order_number === "S260207"));
assert.ok(!db.listDeliveryItems().some((i) => i.order_number === "S260207"));

const weeks = sched.weekOptions(delivery, "2026-09-21");
assert.ok(weeks.some((w) => w.key === "2026-W39"));

console.log("office-schedule.test.js ok");
