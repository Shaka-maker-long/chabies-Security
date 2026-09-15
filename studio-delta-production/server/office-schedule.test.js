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
assert.strictEqual(sched.formatOrderDate("07/01/2026"), "07-Jan");
assert.strictEqual(sched.formatOrderDate("28/07/2026"), "28-Jul");
assert.strictEqual(sched.formatOrderDate("2026-09-21"), "21-Sep");
assert.strictEqual(sched.formatOrderDate("07-Jan"), "07-Jan");
assert.strictEqual(sched.formatOrderDate("7 Jan 2026"), "07-Jan");
assert.deepStrictEqual(sched.workdays("2026-09-21", 5), [
  "2026-09-21", "2026-09-22", "2026-09-23", "2026-09-24", "2026-09-25"
]);
assert.strictEqual(sched.SCHEDULE_WORKDAYS, 180);
const horizon = sched.workdays("2026-09-07", sched.SCHEDULE_WORKDAYS);
assert.strictEqual(horizon.length, 180);
assert.ok(horizon.indexOf("2026-11-30") !== -1, "grid must pass week 48");
assert.ok(horizon[horizon.length - 1] >= "2027-05-01", "grid must reach the following year");
assert.ok(sched.DELIVERY_CODES.indexOf("LD") !== -1);
assert.ok(sched.DELIVERY_CODES.indexOf("LC") !== -1);
assert.ok(sched.SCHEDULE_CODES.some((c) => c.code === "LD*" && c.moved));
assert.ok(sched.SCHEDULE_CODES.some((c) => c.code === "LC*" && c.moved));
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

assert.strictEqual(cabinet.delivery_planned, false);
db.setScheduleCell(cabinet.id, "2026-09-21", "LD");
const planned = db.listSchedule("2026-09-21", "2026-09-25").find((r) => r.order_number === "S260186");
assert.strictEqual(planned.delivery_planned, true);
assert.ok(planned.delivery_days.indexOf("2026-09-21") !== -1);
assert.strictEqual(planned.order_date_label, "01-Sep");
const side = db.listSchedule("2026-09-21", "2026-09-25").find((r) => r.order_number === "S260207");
db.setScheduleCell(side.id, "2026-09-22", "LC");
db.setScheduleCell(side.id, "2026-09-22", "LC");
assert.throws(() => db.setScheduleCell(cabinet.id, "2026-09-24", "LD"), /reason/);
db.setScheduleCell(cabinet.id, "2026-09-24", "LD", true, { reason: "Client asked for later", skipPlan: true });
const moved = db.listSchedule("2026-09-21", "2026-09-25").find((r) => r.order_number === "S260186");
assert.strictEqual(moved.cells["2026-09-21"], "LD*");
assert.strictEqual(moved.cells["2026-09-24"], "LD");
assert.deepStrictEqual(moved.delivery_days, ["2026-09-24"]);
assert.ok(!db.listDeliveryItems().items.some((i) => i.order_number === "S260186" && i.day === "2026-09-21"));
assert.ok(db.listDeliveryItems().items.some((i) => i.order_number === "S260186" && i.day === "2026-09-24" && i.code === "LD"));
db.setScheduleCell(cabinet.id, "2026-09-21", "LD*", true, { skipPlan: true });
db.setScheduleCell(cabinet.id, "2026-09-24", "", true, { skipPlan: true });
db.setScheduleCell(cabinet.id, "2026-09-21", "LD", true, { skipPlan: true });
db.setScheduleCell(cabinet.id, "2026-09-24", "QC");

const delivery = db.listDeliveryItems();
assert.strictEqual(delivery.items.length, 2);
assert.ok(delivery.items.some((i) => i.order_number === "S260186" && i.code === "LD" && i.week === 39));
assert.ok(delivery.items.some((i) => i.order_number === "S260207" && i.code === "LC" && i.weekday === "Tuesday"));
assert.strictEqual(delivery.items.find((i) => i.order_number === "S260186").status, "Not Yet Started");
assert.strictEqual(delivery.items.find((i) => i.order_number === "S260207").status, "At Couriers");
assert.ok(!delivery.items.some((i) => i.code === "QC"));
assert.ok(delivery.categories.indexOf("Cabinet") !== -1);

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
assert.strictEqual(db.listDeliveryItems().items.find((i) => i.order_number === "S260186").status, "In production");
assert.strictEqual(refreshed.cells["2026-09-21"], "LD");

db.deleteOrder("S260207");
assert.ok(!db.listSchedule("2026-09-21", "2026-09-25").some((r) => r.order_number === "S260207"));
assert.ok(!db.listDeliveryItems().items.some((i) => i.order_number === "S260207"));

const weeks = sched.weekOptions(delivery.items, "2026-09-21");
assert.ok(weeks.some((w) => w.key === "2026-W39"));

console.log("office-schedule.test.js ok");
