"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-shop-month-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const staff = require("./staff");
const hub = require("./shop-month");

initWorkbook();
staff.upsertUser({
  name: "Thabo",
  access: "Production",
  role: "Welding",
  password: "1234",
  tasks: ["Welding"]
});
staff.upsertUser({
  name: "Sipho",
  access: "Production",
  role: "Assembly",
  password: "1234",
  tasks: ["Assembly"]
});
staff.upsertUser({
  name: "Lerato",
  access: "Marketing",
  role: "Marketing",
  password: "1234"
});
staff.upsertUser({
  name: "Erin",
  access: "Admin",
  role: "Interior Designer",
  password: "1234",
  seeDebtors: "Yes"
});
staff.upsertUser({
  name: "Shabana",
  access: "Admin",
  role: "Clients Relations Manager",
  password: "1234",
  seeDebtors: "Yes"
});

db.upsertOrder({
  order_number: "S260901",
  status: "Welding",
  product: "Giselle Chair",
  client_name: "Ada"
});
db.upsertOrder({
  order_number: "S260902",
  status: "Assembly",
  product: "Charlotte Chair",
  client_name: "Ben"
});

const book = getBook();
const log = book.getSheetByName("Production_Log");
const meta = JSON.stringify({ entryType: "production", pauses: [], overtimeContinue: false, batchShare: 1 });
log.appendRow([
  "log-weld", "S260901", "Thabo", "Welding", "Done",
  new Date("2026-10-05T07:45:00+02:00"),
  new Date("2026-10-05T12:00:00+02:00"),
  "", "", "", 0, "", meta
]);
log.appendRow([
  "log-asm", "S260902", "Thabo", "Assembly", "Done",
  new Date("2026-10-06T07:45:00+02:00"),
  new Date("2026-10-06T11:45:00+02:00"),
  "", "", "", 0, "", meta
]);
log.appendRow([
  "log-weld-2", "S260901", "Thabo", "Welding", "Done",
  new Date("2026-10-05T12:30:00+02:00"),
  new Date("2026-10-05T13:00:00+02:00"),
  "", "", "", 0, "", meta
]);
log.appendRow([
  "log-weld-a", "S260901 A", "Thabo", "Welding", "Done",
  new Date("2026-10-07T07:45:00+02:00"),
  new Date("2026-10-07T09:00:00+02:00"),
  "", "", "", 0, "", meta
]);
log.appendRow([
  "log-weld-b", "S260901B", "Thabo", "Welding", "Done",
  new Date("2026-10-07T09:15:00+02:00"),
  new Date("2026-10-07T10:15:00+02:00"),
  "", "", "", 0, "", meta
]);
log.appendRow([
  "log-indirect", "", "Thabo", "Indirect", "Sweeping",
  new Date("2026-10-08T07:45:00+02:00"),
  new Date("2026-10-08T08:15:00+02:00"),
  "", "", "", 0, "", JSON.stringify({ entryType: "indirect", pauses: [], overtimeContinue: false, batchShare: 1 })
]);
log.appendRow([
  "log-sep", "S260800", "Thabo", "Welding", "Done",
  new Date("2026-09-10T07:45:00+02:00"),
  new Date("2026-09-10T12:00:00+02:00"),
  "", "", "", 0, "", meta
]);
persistWorkbook();

const enquiry = db.upsertEnquiry({
  date_enquired: "06/10/2026",
  client_name: "Ada",
  product: "Giselle Chair",
  category: "Chair",
  status: "New"
}, { actor: "Shabana" });
const rawEnq = db.getEnquiryRaw(enquiry.enquiry_no);
rawEnq.events = [
  { id: "ev1", at: "2026-10-06T08:00:00.000Z", kind: "created", actor: "Shabana", label: "Enquiry captured", status: "New" },
  { id: "ev2a", at: "2026-10-06T09:00:00.000Z", kind: "complete_cost_sheet", actor: "Shabana", label: "Cost sheet uploaded", status: "Costed" },
  { id: "ev2", at: "2026-10-06T10:00:00.000Z", kind: "complete_quote", actor: "Shabana", label: "Quote PDF issued", status: "Quoted" },
  { id: "ev3", at: "2026-10-08T10:00:00.000Z", kind: "complete_order", actor: "Shabana", label: "POP saved", status: "Ordered" },
  { id: "ev4", at: "2026-10-09T09:00:00.000Z", kind: "complete_drawing", actor: "Erin", label: "Drawing uploaded", status: "Ordered" },
  { id: "ev5", at: "2026-10-09T11:00:00.000Z", kind: "complete_drawing_approval", actor: "Shabana", label: "Client approved drawing", status: "Ordered" }
];
rawEnq.status = "Ordered";
rawEnq.product = "Giselle Chair";
db.saveEnquiryRecord(rawEnq);
require("./order-created").record([{
  order_number: "S260910",
  enquiry_no: enquiry.enquiry_no,
  product: "Giselle Chair",
  client_name: "Ada"
}], {
  actor: "Shabana",
  at: "2026-10-08T12:00:00.000Z",
  readyAt: "2026-10-08T10:00:00.000Z"
});

assert.strictEqual(hub.parseMonth("2026-10"), "2026-10");
assert.strictEqual(hub.monthTitle("2026-10"), "October 2026");
assert.ok(hub.isShopUser({ name: "Thabo", access: "Production", tasks: ["Welding"] }));
assert.ok(!hub.isShopUser({ name: "Lerato", access: "Marketing", tasks: [] }));

(async function main() {
  const data = await hub.build({ month: "2026-10" });
  assert.strictEqual(data.month, "2026-10");
  assert.strictEqual(data.title, "October 2026");
  const thabo = data.people.find((p) => p.worker === "Thabo");
  const sipho = data.people.find((p) => p.worker === "Sipho");
  const lerato = data.people.find((p) => p.worker === "Lerato");
  assert.ok(thabo, "Thabo is on the pack");
  assert.ok(thabo.clocked);
  assert.ok(thabo.totalMinutes > 0, "October clocks count");
  assert.ok(thabo.processes.some((p) => p.name === "Welding"));
  assert.ok(thabo.processes.some((p) => p.name === "Assembly"));
  assert.strictEqual(thabo.orderCount, 4, "repeat clocks collapse; S260901, A and B are three lines plus S260902");
  const welding = thabo.processes.find((p) => p.name === "Welding");
  assert.ok(welding);
  assert.strictEqual(welding.orderCount, 3, "S260901, S260901 A and S260901 B are separate welding orders");
  assert.ok(welding.clocks > welding.orderCount, "repeat sessions of the same line do not add extra orders");
  const indirect = thabo.processes.find((p) => p.name === "Indirect");
  assert.ok(indirect);
  assert.strictEqual(indirect.taskCount, 1);
  assert.ok(!(thabo.orders && thabo.orders.length), "order lines stay off the card payload");
  const weld = (thabo.tasks || []).find((t) => t.order === "S260901");
  assert.ok(weld);
  assert.strictEqual(weld.product, "Giselle Chair");
  assert.ok(weld.imageUrl, "catalog photo is attached for hover");
  assert.ok(sipho, "Sipho is listed even with no clocks");
  assert.ok(!sipho.clocked);
  assert.ok(!lerato, "Marketing is not a shop employee on this pack");
  assert.ok(!hub.isSalesUser({ name: "Lerato", access: "Marketing" }));
  assert.ok(hub.isSalesUser({ name: "Erin", access: "Admin", role: "Interior Designer" }));
  assert.ok(hub.isSalesUser({ name: "Shabana Patel", access: "Admin" }));
  const sales = data.sales;
  assert.ok(sales, "sales pack is attached");
  assert.ok(!sales.people.some((p) => p.worker === "Lerato"), "Marketing is not on the Shabana/Erin pack");
  const salesShabana = sales.people.find((p) => p.worker === "Shabana");
  const salesErin = sales.people.find((p) => p.worker === "Erin");
  assert.ok(salesShabana, "Shabana is on the sales pack");
  assert.ok(salesErin, "Erin is on the sales pack");
  assert.ok(salesShabana.worked);
  assert.ok(salesErin.worked);
  assert.strictEqual(salesShabana.captured, 1);
  assert.strictEqual(salesShabana.costing, 1);
  assert.strictEqual(salesShabana.quoted, 1);
  assert.strictEqual(salesShabana.ordered, 1);
  assert.strictEqual(salesShabana.approval, 1);
  assert.strictEqual(salesErin.drawing, 1);
  assert.strictEqual(sales.addedCount, 1);
  assert.strictEqual(sales.medianLagLabel, "2h");
  assert.strictEqual(sales.added[0].order, "S260910");
  assert.strictEqual(sales.added[0].lagLabel, "2h");
  assert.ok(data.clockedCount >= 1);
  assert.ok(data.absentCount >= 1);
  assert.ok(data.stations.some((s) => s.name === "Welding"));

  staff.setDurations([
    { product: "Giselle Chair", process: "Welding", hours: 2 }
  ]);
  const withPred = await hub.build({ month: "2026-10" });
  const thaboPred = withPred.people.find((p) => p.worker === "Thabo");
  assert.strictEqual(hub.weekdayCount("2026-10"), 22);
  assert.strictEqual(thaboPred.workDays, 22);
  assert.strictEqual(thaboPred.shiftHours, 7.5);
  assert.ok(thaboPred.completedMins > 0);
  assert.strictEqual(thaboPred.predictedMins, 360, "three Giselle welding lines × 2h");
  assert.ok(thaboPred.productivityPct > 0);
  assert.ok(thaboPred.actualProductivityPct <= thaboPred.productivityPct);
  assert.ok(thaboPred.trophies + thaboPred.overtimeJobs >= 1);
  assert.ok(thaboPred.byDay.length === 22);
  assert.ok(thaboPred.byDay.some((d) => d.completedMins > 0));
  assert.ok(withPred.productivityPct != null);
  assert.ok(withPred.productivityPct <= thaboPred.productivityPct + 0.05, "shop figure is the average, not all hours vs one person's month");
  assert.ok(withPred.trophies === thaboPred.trophies);

  const html = fs.readFileSync(path.join(__dirname, "../public/shop-month.html"), "utf8");
  assert.ok(html.indexOf("/api/office/shop-month") !== -1);
  assert.ok(html.indexOf("Print pack") !== -1);
  assert.ok(html.indexOf("Orders they worked") === -1);
  assert.ok(html.indexOf("function openOrders") !== -1);
  assert.ok(html.indexOf("data-product-shot") !== -1);
  assert.ok(html.indexOf("id=\"productShot\"") !== -1);
  assert.ok(html.indexOf("tr.station") !== -1);
  assert.ok(html.indexOf("function groupOrderRows") !== -1);
  assert.ok(html.indexOf("function groupIndirectRows") !== -1);
  assert.ok(html.indexOf("finished clock") === -1);
  assert.ok(html.indexOf("No clocks this month") === -1);
  assert.ok(html.indexOf("No orders this month") !== -1);
  assert.ok(html.indexOf("A and B lines are separate orders") !== -1);
  assert.ok(html.indexOf(">Sales<") !== -1 || html.indexOf("Sales</h2>") !== -1);
  assert.ok(html.indexOf("Shabana and Erin") === -1);
  assert.ok(html.indexOf("Costing") !== -1);
  assert.ok(html.indexOf("marketing employee") === -1);
  assert.ok(html.indexOf("Median time to add") !== -1);
  assert.ok(html.indexOf("function openAdded") !== -1);
  assert.ok(html.indexOf("Completed vs predicted") !== -1);
  assert.ok(html.indexOf("Completed over time") !== -1);
  assert.ok(html.indexOf("Actual productivity") !== -1);
  assert.ok(html.indexOf("7.5") !== -1);
  assert.ok(html.indexOf("shopScoreCards") !== -1);
  assert.ok(html.indexOf("Total completed vs predicted") !== -1);
  assert.ok(html.indexOf("Total completed over time") !== -1);
  assert.ok(html.indexOf("Trophies vs overtime") !== -1);
  assert.ok(html.indexOf("badge trophy") !== -1);
  assert.ok(html.indexOf("Average of each person") !== -1);
  assert.ok(withPred.byDay && withPred.byDay.length === 22);
  assert.strictEqual(hub.baseOrderNo("S260277 A"), "S260277");
  assert.strictEqual(hub.baseOrderNo("S260277B"), "S260277");
  assert.strictEqual(hub.orderLineId("S260300 A"), "S260300 A");
  assert.strictEqual(hub.orderLineId("S260300B"), "S260300 B");
  assert.notStrictEqual(hub.orderLineId("S260300 A"), hub.orderLineId("S260300 B"));
  console.log("shop-month.test.js ok");
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
