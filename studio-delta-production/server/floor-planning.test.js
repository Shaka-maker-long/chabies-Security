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

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
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

assert.ok(!byProcess.Grinding, "the system does not book grinding");
assert.ok(!assignA.Grinding, "grinding is not assigned by the user");
const grindLater = plan.scheduleGrinding({
  orderId: "S260100 A",
  worker: "Thabo",
  from: byProcess.Welding[byProcess.Welding.length - 1].end
});
assert.ok(grindLater.blocks[0].start >= byProcess.Welding[byProcess.Welding.length - 1].end, "grinding after welding");
assert.strictEqual(grindLater.blocks[0].workerName, "Thabo");

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
const weldQ = board.queue.find((q) => q.order_number === "S260199");
assert.ok(weldQ, "in-progress orders stay in the queue for remaining work");
assert.ok(!weldQ.processes.some((p) => p.process === "Profile Cutting" || p.process === "Tagging" || p.process === "Welding"));
assert.ok(weldQ.processes.some((p) => p.process === "Grinding"));
const productA = board.queue.find((q) => q.order_number === "S260100 A");
assert.ok(!productA.processes.some((p) => p.process === "Quality Control"));
assert.ok(productA.processes.find((p) => p.process === "Welding").hours === 3);
assert.ok(productA.processes.find((p) => p.process === "Welding").workers.some((w) => w.name === "Thabo"));
const grindRow = productA.processes.find((p) => p.process === "Grinding");
assert.ok(grindRow && !grindRow.auto, "grinding is booked by the user");
assert.ok(grindRow.workers.some((w) => w.name === "Thabo"));

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
assert.ok(trip.start && trip.start.indexOf("2026-09-08") === 0);
assert.ok(board.weekStarting.some((o) => o.orderId === "S260100 A"));
assert.ok(board.bands.some((b) => b.label === "Production Meeting"));
assert.ok(board.bands.some((b) => b.label === "Lunch"));
assert.ok(board.bands.some((b) => b.label === "Cleaning"));
assert.strictEqual(trip.rows.find((r) => r.process === "Profile Cutting").workerName, "Willard");
assert.strictEqual(trip.rows.find((r) => r.process === "Tagging").workerName, "Sipho");
assert.strictEqual(trip.rows.find((r) => r.process === "Welding").workerName, "Thabo");
assert.ok(trip.rows.find((r) => r.process === "Profile Cutting").segments.length >= 1);
assert.ok(trip.rows.find((r) => r.process === "Tagging").days.indexOf("2026-09-08") !== -1);
assert.ok(trip.rows.find((r) => r.process === "Powder coating").days.indexOf("2026-09-14") !== -1);
assert.ok(trip.rows.find((r) => r.process === "Assembly").days.indexOf("2026-09-21") !== -1);
assert.deepStrictEqual(trip.rows.find((r) => r.process === "Profile Cutting").actual.days, [], "no clock yet means an empty Actual row");
assert.ok(journey.days.some((d) => d.iso === "2026-09-21"));
assert.ok(!journey.days.some((d) => d.iso === "2026-09-12" || d.iso === "2026-09-13"), "weekends stay off the journey");

const profileA = plan.load().blocks.find((b) => b.orderId === "S260100 A" && b.process === "Profile Cutting");
assert.ok(profileA);
const delayed = plan.moveBlock(profileA.id, "2026-09-08T09:00:00+02:00");
assert.ok(delayed.blocks.length);
const profileLater = plan.load().blocks.find((b) => b.orderId === "S260100 A" && b.process === "Profile Cutting");
const tagLater = plan.load().blocks.find((b) => b.orderId === "S260100 A" && b.process === "Tagging");
const weldLater = plan.load().blocks.find((b) => b.orderId === "S260100 A" && b.process === "Welding");
assert.ok(profileLater.start >= "2026-09-08T09:00:00+02:00", "drag later delays that job");
assert.ok(tagLater.start >= profileLater.end, "tagging follows the new profile end");
assert.ok(weldLater.start >= tagLater.end, "welding follows tagging after the shift");
assert.ok(board.journeyWeeks && board.journeyWeeks.length >= plan.JOURNEY_WEEK_COUNT, "journey shows many week columns so the order column can freeze while you scroll");
assert.strictEqual(board.journeyWeeks[0].start, board.weekStart, "journey weeks start from the current shop week");
assert.ok(/Sep/.test(board.journeyWeeks[0].label), "week headers are real dates, not Week 1");
assert.ok(board.journeyWeeks.every((w) => !/^Week \d+$/.test(w.label)), "week headers are date ranges, not Week 1 / 2 / 3");
assert.ok(board.firstPlannedWeek, "board can jump to the first week that has planned work");
assert.strictEqual(board.firstPlannedWeek, plan.weekMondayIso(trip.start));
assert.strictEqual(plan.processCode("Profile Cutting"), "C");
assert.strictEqual(plan.processCode("Powder coating"), "PC");
assert.ok(trip.rows.find((r) => r.process === "Profile Cutting").code === "C");

const removed = plan.unscheduleOrder("S260100 A");
assert.ok(removed.removed > 0);
assert.ok(!plan.load().blocks.some((b) => b.orderId === "S260100 A"));

plan.save({ blocks: [] });
plan.scheduleSelected({
  orderIds: ["S260100 A"],
  assignments: { "S260100 A": assignA },
  from: plan.isoFromMs(fromTue)
});
assert.throws(() => plan.insertOtherTask({
  workerId: "Willard",
  title: "Material",
  minutes: 60,
  start: "2026-09-08T07:45:00+02:00"
}), /already has work/);
const profileFrozen = plan.load().blocks.find((b) => b.orderId === "S260100 A" && b.process === "Profile Cutting");
const profileStartBefore = profileFrozen.start;
const other = plan.insertOtherTask({
  workerId: "Willard",
  title: "Material",
  minutes: 60,
  start: "2026-09-08T14:00:00+02:00"
});
assert.ok(other.jobId);
const otherBlock = plan.load().blocks.find((b) => b.kind === "other" && b.title === "Material");
assert.ok(otherBlock);
assert.ok(otherBlock.start >= "2026-09-08T14:00:00+02:00");
const profileAfterOther = plan.load().blocks.find((b) => b.orderId === "S260100 A" && b.process === "Profile Cutting");
assert.strictEqual(profileAfterOther.start, profileStartBefore, "an other task must not shove existing shop work");
assert.ok(plan.getBoard("2026-09-08").otherTasks.indexOf("Material") !== -1);

plan.removeBlock(otherBlock.id);
assert.ok(!plan.load().blocks.some((b) => b.kind === "other"));

plan.save({ blocks: [] });
plan.scheduleSelected({
  orderIds: ["S260100 A"],
  assignments: { "S260100 A": assignA },
  from: plan.isoFromMs(fromTue)
});
plan.scheduleSelected({
  orderIds: ["S260100 B"],
  assignments: { "S260100 B": assignA },
  from: plan.isoFromMs(fromTue)
});
const profileB = plan.load().blocks.find((b) => b.orderId === "S260100 B" && b.process === "Profile Cutting");
const profileAKept = plan.load().blocks.find((b) => b.orderId === "S260100 A" && b.process === "Profile Cutting");
assert.ok(profileB);
assert.throws(() => plan.moveBlock(profileB.id, "2026-09-08T07:45:00+02:00"), /already has work/);
const stillA = plan.load().blocks.find((b) => b.orderId === "S260100 A" && b.process === "Profile Cutting");
const stillB = plan.load().blocks.find((b) => b.orderId === "S260100 B" && b.process === "Profile Cutting");
assert.strictEqual(stillA.start, profileAKept.start, "overlapping a move must leave the other order where it was");
assert.strictEqual(stillB.start, profileB.start, "rejected move must not move the dragged order");
const laterB = plan.moveBlock(profileB.id, stillB.end);
assert.ok(laterB.blocks.length);
const movedB = plan.load().blocks.find((b) => b.orderId === "S260100 B" && b.process === "Profile Cutting");
const afterA = plan.load().blocks.find((b) => b.orderId === "S260100 A" && b.process === "Profile Cutting");
assert.ok(movedB.start >= stillB.end, "a free-slot move is allowed");
assert.strictEqual(afterA.start, stillA.start, "moving one order must not drop the other order");

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
assert.ok(!batch.blocks.some((b) => b.process === "Grinding"), "batch does not auto-place grinding");
const weldAEnd = batch.blocks.filter((b) => b.orderId === "S260100 A" && b.process === "Welding").pop().end;
const grindA = plan.scheduleGrinding({ orderId: "S260100 A", worker: "Sam", from: weldAEnd });
assert.ok(grindA.blocks[0].start >= weldAEnd, "grind waits for that order's welding");
assert.strictEqual(grindA.blocks[0].workerName, "Sam");

const priorBlocks = plan.load();
plan.save({
  blocks: [
    {
      id: "late-a",
      orderId: "S260001 A",
      studioNo: "S260001 A",
      product: "Product A",
      process: "Tagging",
      workerId: "Sipho",
      workerName: "Sipho",
      start: "2026-09-08T12:45:00+02:00",
      end: "2026-09-08T13:15:00+02:00",
      kind: "work"
    },
    {
      id: "early-f",
      orderId: "S260001 F",
      studioNo: "S260001 F",
      product: "Product A",
      process: "Profile Cutting",
      workerId: "Willard",
      workerName: "Willard",
      start: "2026-09-08T07:45:00+02:00",
      end: "2026-09-08T09:15:00+02:00",
      kind: "work"
    }
  ]
});
const byStart = plan.getBoard("2026-09-08").journey.orders.map((o) => o.orderId);
assert.strictEqual(byStart[0], "S260001 F", "the order that starts first is listed first, not A–Z");
assert.ok(byStart.indexOf("S260001 F") < byStart.indexOf("S260001 A"));
plan.save(priorBlocks);

assert.strictEqual(plan.matchJourneyProcess("Profile Cutter"), "Profile Cutting");
getBook().getSheetByName("Production_Log").appendRow([
  "log_cut_week2",
  "S260100 A",
  "Willard",
  "Profile Cutting",
  "Done",
  new Date("2026-09-15T05:45:00.000Z"),
  new Date("2026-09-15T06:45:00.000Z"),
  "",
  "",
  "",
  0,
  "",
  ""
]);
const compared = plan.getBoard("2026-09-08").journey.orders.find((o) => o.orderId === "S260100 A");
const cutPlan = compared.rows.find((r) => r.process === "Profile Cutting");
assert.ok(cutPlan.days.indexOf("2026-09-08") !== -1, "plan stays on the booked Monday");
assert.ok(cutPlan.actual.days.indexOf("2026-09-15") !== -1, "actual Cutting lands in the week it was clocked");
assert.ok(String(cutPlan.actual.start).indexOf("2026-09-15T07:45") === 0);
assert.ok(compared.rows.find((r) => r.process === "Tagging").actual.days.length === 0, "other processes stay empty until clocked");

const catalog = require("./product-catalog");
const talithaUrl = catalog.lookupProduct("Talitha Bookshelf").imageUrl;
assert.ok(talithaUrl && /TALITHA/i.test(talithaUrl));
const withPhoto = plan.buildJourney([{
  id: "photo-talitha",
  orderId: "S260193",
  product: "Talitha Bookshelf",
  process: "Tagging",
  workerId: "Sipho",
  workerName: "Sipho",
  start: "2026-09-10T07:45:00+02:00",
  end: "2026-09-10T08:45:00+02:00",
  kind: "work"
}]);
assert.strictEqual(withPhoto.orders[0].product, "Talitha Bookshelf");
assert.strictEqual(withPhoto.orders[0].imageUrl, talithaUrl, "Journey carries the catalog photo for hover");
assert.strictEqual(compared.imageUrl || "", "", "unknown products have no photo");

const idleSheet = getBook().getSheetByName("Idle_Alerts");
idleSheet.appendRow([
  "2026-09-10",
  "Willard",
  "Cutter",
  new Date("2026-09-10T06:00:00.000Z"),
  new Date("2026-09-10T06:15:00.000Z"),
  "Assigned",
  "Cleaning",
  new Date("2026-09-10T06:45:00.000Z"),
  ""
]);
idleSheet.appendRow([
  "2026-09-10",
  "Sipho",
  "Tagger",
  new Date("2026-09-10T07:00:00.000Z"),
  new Date("2026-09-10T07:15:00.000Z"),
  "Open",
  "",
  "",
  ""
]);
persistWorkbook();
const idleJourney = plan.buildJourney([]);
assert.ok((idleJourney.otherActuals || []).some((row) => row.code === "O" && row.title === "Cleaning"), JSON.stringify(idleJourney.otherActuals));
assert.ok((idleJourney.otherActuals || []).every((row) => String(row.workerName) !== "Sipho"), "open idle holes are not Journey actuals yet");
assert.strictEqual(plan.PROCESS_CODES.Other, "O");

const split = plan.buildJourney([
  {
    id: "cut-am",
    orderId: "S260214 B",
    product: "Naomi Arched Cabinet",
    process: "Profile Cutting",
    workerId: "Sam",
    workerName: "Sam",
    start: "2026-09-10T08:45:00+02:00",
    end: "2026-09-10T09:00:00+02:00",
    kind: "work"
  },
  {
    id: "cut-pm",
    orderId: "S260214 B",
    product: "Naomi Arched Cabinet",
    process: "Profile Cutting",
    workerId: "Sam",
    workerName: "Sam",
    start: "2026-09-10T12:30:00+02:00",
    end: "2026-09-10T13:15:00+02:00",
    kind: "work"
  }
]);
const splitCut = split.orders[0].rows.find((r) => r.process === "Profile Cutting");
assert.strictEqual(splitCut.segments.length, 2, "split cutting stays as separate booked slots");
assert.strictEqual(splitCut.start, "2026-09-10T08:45:00+02:00");
assert.strictEqual(splitCut.end, "2026-09-10T13:15:00+02:00");
assert.strictEqual(splitCut.segments[0].end, "2026-09-10T09:00:00+02:00");
assert.strictEqual(splitCut.segments[1].start, "2026-09-10T12:30:00+02:00");

plan.save({
  blocks: [
    {
      id: "keep-me",
      orderId: "S260888",
      process: "Profile Cutting",
      workerId: "Sam",
      workerName: "Sam",
      start: "2026-09-10T07:45:00+02:00",
      end: "2026-09-10T08:45:00+02:00",
      kind: "work"
    }
  ],
  assignments: { "S260888": { "Profile Cutting": "Sam" } }
});
const wipedPlan = plan.clearAllPlanning();
assert.ok(wipedPlan.removed >= 1);
assert.strictEqual(plan.load().blocks.length, 0);
assert.deepStrictEqual(plan.load().assignments, {});
plan.save({
  blocks: [{
    id: "gone-with-order",
    orderId: "S260100 A",
    process: "Tagging",
    workerId: "Sipho",
    workerName: "Sipho",
    start: "2026-09-10T07:45:00+02:00",
    end: "2026-09-10T08:45:00+02:00",
    kind: "work"
  }],
  assignments: { "S260100 A": { Tagging: "Sipho" } }
});
db.deleteOrder("S260100 A");
assert.ok(!plan.load().blocks.some((b) => b.orderId === "S260100 A"), "deleting an order drops its planning slots");
db.upsertOrder({
  order_number: "S260900",
  status: "Not Yet Started",
  product: "Air Chair"
});
plan.save({
  blocks: [{
    id: "wipe-all",
    orderId: "S260900",
    process: "Assembly",
    workerId: "Nomsa",
    workerName: "Nomsa",
    start: "2026-09-10T07:45:00+02:00",
    end: "2026-09-10T08:45:00+02:00",
    kind: "work"
  }],
  assignments: { S260900: { Assembly: "Nomsa" } }
});
db.deleteAllOrders();
assert.strictEqual(plan.load().blocks.length, 0, "clearing orders also clears planning");
assert.deepStrictEqual(plan.load().assignments, {});

console.log("floor-planning.test.js ok");
