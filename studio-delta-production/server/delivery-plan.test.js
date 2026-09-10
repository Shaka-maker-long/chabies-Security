const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-deliv-plan-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const db = require("./db");
const staff = require("./staff");
const plan = require("./floor-planning");

initWorkbook();

[
  ["Sam", ["Profile Cutting"]],
  ["John", ["Tagging"]],
  ["Muruba", ["Welding"]],
  ["Willard", ["Welding"]],
  ["Thabile", ["Plate Cutting"]],
  ["Admire", ["Assembly"]],
  ["Uriah", ["Assembly"]],
  ["Nomsa", ["Assembly"]],
  ["Thabo", ["Grinding"]]
].forEach(([name, tasks]) => {
  staff.upsertUser({
    name,
    access: "Production",
    role: "Shop",
    password: "1234",
    tasks
  });
});

staff.setDurations([
  { product: "Air Chair", process: "Profile Cutting", hours: 1 },
  { product: "Air Chair", process: "Tagging", hours: 1 },
  { product: "Air Chair", process: "Plate Cutting", hours: 1 },
  { product: "Air Chair", process: "Welding", hours: 2 },
  { product: "Air Chair", process: "Grinding", hours: 1 },
  { product: "Air Chair", process: "Assembly", hours: 1 }
]);

db.upsertOrder({
  order_number: "S260701",
  status: "Not Yet Started",
  product: "Air Chair",
  client_name: "Late",
  type: "Standard"
});
db.upsertOrder({
  order_number: "S260702",
  status: "Not Yet Started",
  product: "Air Chair",
  client_name: "Early",
  type: "Standard"
});

const rows = db.listSchedule("2026-09-21", "2026-10-02");
const late = rows.find((r) => r.order_number === "S260701");
const early = rows.find((r) => r.order_number === "S260702");
db.setScheduleCell(late.id, "2026-09-30", "LC", true, { skipPlan: true });
db.setScheduleCell(early.id, "2026-09-23", "LD", true, { skipPlan: true });

plan.save({ blocks: [], assignments: {} });
const fromTue = plan.sastMs(2026, 9, 8, 7, 45);
const auto = plan.autoPlanFromDeliveries({ from: fromTue });
assert.ok(auto.count > 0, auto.error || "auto plan should place work");

const q = plan.queueOrders();
assert.ok(q[0].order_number === "S260702", "earliest LD is First Out");
assert.strictEqual(q[0].delivery_code, "LD");
assert.ok(!q.some((row) => row.processes.some((p) => p.auto)));

const store = plan.load();
assert.ok(!store.blocks.some((b) => b.process === "Grinding"), "auto plan leaves grinding for the user");
const cutEarly = store.blocks.find((b) => b.orderId === "S260702" && b.process === "Profile Cutting");
const cutLate = store.blocks.find((b) => b.orderId === "S260701" && b.process === "Profile Cutting");
assert.ok(cutEarly && cutLate);
assert.ok(cutEarly.start <= cutLate.start, "earliest LD uses the first profile-cutting slot");
assert.strictEqual(cutEarly.workerName, "Sam");
assert.strictEqual(store.blocks.find((b) => b.process === "Tagging").workerName, "John");
assert.strictEqual(store.blocks.find((b) => b.process === "Plate Cutting").workerName, "Thabile");
assert.ok(["Admire", "Uriah"].indexOf(store.blocks.find((b) => b.process === "Assembly").workerName) !== -1);

const welds = store.blocks.filter((b) => b.process === "Welding");
assert.ok(welds.length);
welds.forEach((b) => {
  const day = new Date(b.start).getDay();
  if (b.workerName === "Muruba") {
    assert.ok([1, 3, 5].indexOf(day) !== -1, "Muruba only welds Mon Wed Fri");
  }
});
const tueWeld = welds.find((b) => new Date(b.start).getDay() === 2);
if (tueWeld) assert.strictEqual(tueWeld.workerName, "Willard");

const grind = plan.scheduleGrinding({ orderId: "S260702", worker: "Thabo" });
assert.strictEqual(grind.blocks[0].workerName, "Thabo");
assert.ok(plan.load().blocks.some((b) => b.orderId === "S260702" && b.process === "Grinding"));

console.log("delivery-plan.test.js ok");
