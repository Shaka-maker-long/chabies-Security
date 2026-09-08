const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-planning-"));
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

staff.upsertUser({
  name: "Willard",
  access: "Production",
  role: "Cutter",
  password: "1234",
  tasks: ["Profile Cutting", "Plate Cutting"]
});
staff.upsertUser({
  name: "Sipho",
  access: "Production",
  role: "Tagger",
  password: "1234",
  tasks: ["Tagging"]
});
staff.upsertUser({
  name: "Thabo",
  access: "Production",
  role: "Welder",
  password: "1234",
  tasks: ["Welding", "Grinding"]
});
staff.upsertUser({
  name: "Nomsa",
  access: "Production",
  role: "Assembler",
  password: "1234",
  tasks: ["Assembly"]
});
staff.upsertUser({
  name: "Office Only",
  access: "Admin",
  role: "Quoting",
  password: "office",
  tasks: []
});
staff.upsertUser({
  name: "QC Pat",
  access: "Production",
  role: "Quality Control",
  password: "1234",
  tasks: ["Quality Control"]
});

staff.setDurations([
  { product: "Product A", process: "Profile Cutting", hours: 1 },
  { product: "Product A", process: "Tagging", hours: 2 },
  { product: "Product A", process: "Plate Cutting", hours: 1 },
  { product: "Product A", process: "Welding", hours: 3 },
  { product: "Product A", process: "Grinding", hours: 1.5 },
  { product: "Product A", process: "Quality Control", hours: 9 },
  { product: "Product A", process: "Paint Preparation", hours: 9 },
  { product: "Product A", process: "Painting", hours: 9 },
  { product: "Product A", process: "Assembly", hours: 4 },
  { product: "No Tag", process: "Profile Cutting", hours: 1 },
  { product: "No Tag", process: "Plate Cutting", hours: 1 },
  { product: "No Tag", process: "Welding", hours: 1 },
  { product: "No Plate", process: "Profile Cutting", hours: 1 },
  { product: "No Plate", process: "Tagging", hours: 1 },
  { product: "No Plate", process: "Welding", hours: 1 }
]);

assert.deepStrictEqual(plan.weekDays("2026-09-10").map((d) => d.iso), [
  "2026-09-07", "2026-09-08", "2026-09-09", "2026-09-10", "2026-09-11"
]);
assert.strictEqual(plan.weekMondayIso("2026-09-10"), "2026-09-07");
plan.weekDays("2026-09-07").forEach((d) => {
  assert.ok(["Saturday", "Sunday", "Sat", "Sun"].indexOf(d.weekday) === -1, "week view is work days only");
  assert.ok(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"].indexOf(d.weekday) !== -1);
});
assert.strictEqual(plan.weekDays("2026-09-07").length, 5);

const tue1100 = plan.sastMs(2026, 9, 8, 11, 0);
const weld = plan.placeTask(tue1100, 180, []);
assert.strictEqual(weld.segments.length, 2, "3h from 11:00 spans lunch");
assert.strictEqual(weld.segments[0].start, "2026-09-08T11:00:00+02:00");
assert.strictEqual(weld.segments[0].end, "2026-09-08T12:00:00+02:00");
assert.strictEqual(weld.segments[1].start, "2026-09-08T12:30:00+02:00");
assert.strictEqual(weld.segments[1].end, "2026-09-08T14:30:00+02:00");

const fri1400 = plan.sastMs(2026, 9, 11, 14, 0);
const weekend = plan.placeTask(fri1400, 180, []);
assert.strictEqual(weekend.segments[0].start, "2026-09-11T14:00:00+02:00");
assert.strictEqual(weekend.segments[0].end, "2026-09-11T15:45:00+02:00");
assert.strictEqual(weekend.segments[1].start, "2026-09-14T07:45:00+02:00");
assert.strictEqual(weekend.segments[1].end, "2026-09-14T09:00:00+02:00");

const busy = [{ start: "2026-09-08T08:00:00+02:00", end: "2026-09-08T10:00:00+02:00" }];
const queued = plan.placeTask(plan.sastMs(2026, 9, 8, 7, 45), 60, busy);
assert.strictEqual(queued.segments[0].start, "2026-09-08T07:45:00+02:00");
assert.strictEqual(queued.segments[1].start, "2026-09-08T10:00:00+02:00");
const hole = plan.placeContiguousTask(plan.sastMs(2026, 9, 8, 7, 45), 60, busy);
assert.strictEqual(hole.segments.length, 1, "grinding waits for a real idle hole, not 15 minutes before the next job");
assert.strictEqual(hole.segments[0].start, "2026-09-08T10:00:00+02:00");
assert.strictEqual(hole.segments[0].end, "2026-09-08T11:00:00+02:00");

const sat = plan.nextWorkInstant(plan.sastMs(2026, 9, 12, 10, 0));
assert.strictEqual(plan.isoFromMs(sat), "2026-09-14T07:45:00+02:00");

assert.strictEqual(
  plan.isoFromMs(plan.earliestPaintMonday(plan.sastMs(2026, 9, 11, 15, 0))),
  "2026-09-14T07:45:00+02:00",
  "Friday grind drops next Monday"
);
assert.strictEqual(
  plan.isoFromMs(plan.earliestPaintMonday(plan.sastMs(2026, 9, 14, 10, 0))),
  "2026-09-14T07:45:00+02:00",
  "Monday grind drops that Monday"
);

db.upsertOrder({
  order_number: "S260100 A",
  status: "Not Yet Started",
  product: "Product A",
  type: "Standard",
  category: "Chair"
});
db.upsertOrder({
  order_number: "S260100 B",
  status: "Ready for Steelwork",
  product: "Product A",
  type: "Standard",
  category: "Chair"
});
db.upsertOrder({
  order_number: "S260199",
  status: "Welding",
  product: "Product A"
});
db.upsertOrder({
  order_number: "S260101",
  status: "Not Yet Started",
  product: "No Tag"
});
db.upsertOrder({
  order_number: "S260102",
  status: "Not Yet Started",
  product: "No Plate"
});

const fromTue = plan.sastMs(2026, 9, 8, 7, 45);
const assignA = {
  "Profile Cutting": "Willard",
  "Tagging": "Sipho",
  "Plate Cutting": "Willard",
  "Welding": "Thabo",
  "Assembly": "Nomsa"
};

const first = plan.scheduleSelected({
  orderIds: ["S260100 A"],
  assignments: { "S260100 A": assignA },
  from: plan.isoFromMs(fromTue)
});

const byProcess = {};
first.blocks.forEach((b) => {
  if (!byProcess[b.process]) byProcess[b.process] = [];
  byProcess[b.process].push(b);
});

assert.ok(!byProcess["Quality Control"] && !byProcess["Paint Preparation"] && !byProcess.Painting, "skip QC and in-house paint");
assert.ok(byProcess["Profile Cutting"]);
assert.strictEqual(byProcess["Profile Cutting"][0].start, "2026-09-08T07:45:00+02:00");
assert.strictEqual(byProcess["Profile Cutting"][0].end, "2026-09-08T08:45:00+02:00");
assert.strictEqual(byProcess.Tagging[0].start, "2026-09-08T08:45:00+02:00");
assert.strictEqual(byProcess.Tagging[0].end, "2026-09-08T10:45:00+02:00");

assert.ok(byProcess.Welding[0].start >= byProcess.Tagging[0].end, "welding waits for tagging");
assert.strictEqual(byProcess.Welding[0].start, "2026-09-08T10:45:00+02:00");
assert.strictEqual(byProcess["Plate Cutting"][0].start, "2026-09-08T10:45:00+02:00");
assert.ok(byProcess["Plate Cutting"][0].end <= byProcess.Welding[byProcess.Welding.length - 1].end, "plate may overlap welding");
assert.ok(byProcess.Welding.some((b) => b.start < byProcess["Plate Cutting"][0].end), "welding does not wait for plate");

const grindStart = byProcess.Grinding[0].start;
const weldEnd = byProcess.Welding[byProcess.Welding.length - 1].end;
assert.ok(grindStart >= weldEnd, "grinding after welding");
assert.strictEqual(byProcess.Grinding[0].workerName, "Thabo", "only Thabo is in the grinding pool");
assert.ok(!assignA.Grinding, "grinding is not assigned by the user");

const paint = byProcess["Powder coating"][0];
assert.strictEqual(paint.workerId, plan.PAINT_WORKER_ID);
assert.strictEqual(paint.start, "2026-09-14T07:45:00+02:00");
assert.strictEqual(paint.end, "2026-09-19T07:45:00+02:00");
assert.ok(byProcess.Assembly[0].start >= "2026-09-21T07:45:00+02:00", "assembly after 5 calendar days and next work slot");
assert.strictEqual(byProcess.Assembly[0].workerName, "Nomsa");

assert.throws(
  () => plan.scheduleSelected({
    orderIds: ["S260100 A"],
    assignments: { "S260100 A": Object.assign({}, assignA, { Welding: "Willard" }) },
    from: plan.isoFromMs(fromTue)
  }),
  /not ticked for Welding/
);

const second = plan.scheduleSelected({
  orderIds: ["S260100 B"],
  assignments: { "S260100 B": assignA },
  from: plan.isoFromMs(fromTue)
});
const bProfile = second.blocks.find((b) => b.process === "Profile Cutting");
assert.ok(bProfile.start >= byProcess["Profile Cutting"][0].end, "second order waits for Willard's free slot");
assert.ok(
  bProfile.end <= byProcess["Plate Cutting"][0].start || bProfile.start >= byProcess["Plate Cutting"][0].end,
  "Willard cannot overlap the first order's plate cutting"
);

const willardBlocks = plan.load().blocks
  .filter((b) => b.workerId === "Willard")
  .sort((a, b) => a.start.localeCompare(b.start));
for (let i = 1; i < willardBlocks.length; i++) {
  assert.ok(willardBlocks[i].start >= willardBlocks[i - 1].end, "same person never overlaps");
}

plan.save({ blocks: [] });
const noTag = plan.scheduleSelected({
  orderIds: ["S260101"],
  assignments: {
    S260101: { "Profile Cutting": "Willard", "Plate Cutting": "Willard", Welding: "Thabo" }
  },
  from: plan.isoFromMs(fromTue)
});
const noTagProfile = noTag.blocks.find((b) => b.process === "Profile Cutting");
const noTagPlate = noTag.blocks.find((b) => b.process === "Plate Cutting");
assert.strictEqual(noTagPlate.start, noTagProfile.end, "if tagging hours are 0, plates start after profile cutting");

plan.save({ blocks: [] });
const noPlate = plan.scheduleSelected({
  orderIds: ["S260102"],
  assignments: {
    S260102: { "Profile Cutting": "Willard", Tagging: "Sipho", Welding: "Thabo" }
  },
  from: plan.isoFromMs(fromTue)
});
assert.ok(!noPlate.blocks.some((b) => b.process === "Plate Cutting"), "skip plate when hours are 0");
assert.ok(noPlate.blocks.some((b) => b.process === "Welding"));

plan.save({ blocks: [] });
plan.scheduleSelected({
  orderIds: ["S260100 A"],
  assignments: { "S260100 A": assignA },
  from: plan.isoFromMs(fromTue)
});
const replaced = plan.scheduleSelected({
  orderIds: ["S260100 A"],
  assignments: { "S260100 A": assignA },
  from: plan.isoFromMs(fromTue)
});
const remainingA = plan.load().blocks.filter((b) => b.orderId === "S260100 A");
assert.strictEqual(remainingA.length, replaced.blocks.length, "reschedule replaces the order's blocks");

const board = plan.getBoard("2026-09-08");
assert.strictEqual(board.weekStart, "2026-09-07");
assert.strictEqual(board.weekDays.length, 5);
assert.ok(board.workers.some((w) => w.name === "Willard"));
assert.ok(board.workers.some((w) => w.name === "Sipho"));
assert.ok(board.workers.some((w) => w.name === "Thabo"));
assert.ok(board.workers.some((w) => w.name === "Nomsa"));
assert.ok(board.workers.some((w) => w.name === "QC Pat"), "every shop worker has a calendar");
assert.ok(!board.workers.some((w) => w.name === "Office Only"), "people without floor tasks are not calendars");
assert.strictEqual(board.paintShop.id, plan.PAINT_WORKER_ID);
assert.ok(board.queue.some((q) => q.order_number === "S260100 A" && q.scheduled));
assert.ok(board.queue.some((q) => q.order_number === "S260100 B"));
assert.ok(!board.queue.some((q) => q.order_number === "S260199"), "only NYS and Ready for Steelwork");
const productA = board.queue.find((q) => q.order_number === "S260100 A");
assert.ok(!productA.processes.some((p) => p.process === "Quality Control"));
assert.ok(productA.processes.find((p) => p.process === "Welding").hours === 3);
assert.ok(productA.processes.find((p) => p.process === "Welding").workers.some((w) => w.name === "Thabo"));
const grindRow = productA.processes.find((p) => p.process === "Grinding");
assert.ok(grindRow && grindRow.auto, "grinding is auto-assigned");
assert.deepStrictEqual(grindRow.workers, []);

const journey = board.journey;
assert.ok(journey);
assert.ok(journey.days.length > 0, "journey has work-day columns");
journey.days.forEach((d) => {
  assert.ok(["Saturday", "Sunday"].indexOf(d.weekday) === -1, "journey dates are work days");
  assert.ok(/^\d{2}-[A-Z][a-z]{2}$/.test(d.label), "date header like 08-Sep");
});
assert.strictEqual(plan.formatDayHeader("2026-09-08"), "08-Sep");
const trip = journey.orders.find((o) => o.orderId === "S260100 A");
assert.ok(trip, "scheduled order has a journey");
assert.strictEqual(trip.rows.find((r) => r.process === "Profile Cutting").workerName, "Willard");
assert.strictEqual(trip.rows.find((r) => r.process === "Tagging").workerName, "Sipho");
assert.strictEqual(trip.rows.find((r) => r.process === "Welding").workerName, "Thabo");
assert.ok(trip.rows.find((r) => r.process === "Tagging").days.indexOf("2026-09-08") !== -1);
assert.ok(trip.rows.find((r) => r.process === "Powder coating").days.indexOf("2026-09-14") !== -1);
assert.ok(trip.rows.find((r) => r.process === "Assembly").days.indexOf("2026-09-21") !== -1);
assert.ok(journey.days.some((d) => d.iso === "2026-09-21"));
assert.ok(!journey.days.some((d) => d.iso === "2026-09-12" || d.iso === "2026-09-13"), "weekends stay off the journey");

const removed = plan.unscheduleOrder("S260100 A");
assert.ok(removed.removed > 0);
assert.ok(!plan.load().blocks.some((b) => b.orderId === "S260100 A"));

staff.upsertUser({
  name: "Sipho",
  access: "Production",
  role: "Tagger",
  password: "1234",
  tasks: ["Tagging", "Grinding"]
});
staff.upsertUser({
  name: "Sam",
  access: "Production",
  role: "Metal",
  password: "1234",
  tasks: ["Profile Cutting", "Tagging", "Welding", "Grinding"]
});
plan.save({ blocks: [] });
const batch = plan.scheduleSelected({
  orderIds: ["S260100 A", "S260100 B"],
  assignments: {
    "S260100 A": assignA,
    "S260100 B": assignA
  },
  from: plan.isoFromMs(fromTue)
});
const grindA = batch.blocks.filter((b) => b.orderId === "S260100 A" && b.process === "Grinding");
const grindB = batch.blocks.filter((b) => b.orderId === "S260100 B" && b.process === "Grinding");
const weldAEnd = batch.blocks.filter((b) => b.orderId === "S260100 A" && b.process === "Welding").pop().end;
const weldB = batch.blocks.filter((b) => b.orderId === "S260100 B" && b.process === "Welding");
assert.ok(grindA.length, "batch still auto-places grinding");
assert.ok(grindA[0].start >= weldAEnd, "grind waits for that order's welding");
assert.strictEqual(grindA[0].workerName, "Sam", "idle Sam gets grinding, not the tagger who already has work");
assert.ok(!batch.blocks.some((b) => b.process === "Grinding" && b.workerName === "Sipho"), "do not pile grinding onto a busy tagger");
assert.ok(grindB.length);
assert.ok(["Sam", "Thabo"].indexOf(grindB[0].workerName) !== -1);
assert.ok(!batch.blocks.some((b) => b.process === "Grinding" && b.workerName === "Nomsa"));
assert.ok(!batch.blocks.some((b) => b.process === "Grinding" && b.workerName === "Willard"));

console.log("floor-planning.test.js ok");
