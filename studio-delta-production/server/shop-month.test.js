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
}, { actor: "Lerato" });
const rawEnq = db.getEnquiryRaw(enquiry.enquiry_no);
rawEnq.events = [
  { id: "ev1", at: "2026-10-06T08:00:00.000Z", kind: "created", actor: "Lerato", label: "Enquiry captured", status: "New" },
  { id: "ev2", at: "2026-10-06T10:00:00.000Z", kind: "complete_quote", actor: "Lerato", label: "Quote PDF issued", status: "Quoted" },
  { id: "ev3", at: "2026-10-08T10:00:00.000Z", kind: "complete_order", actor: "Lerato", label: "POP saved", status: "Ordered" }
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
  actor: "Lerato",
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
  assert.ok(hub.isSalesUser({ name: "Lerato", access: "Marketing" }));
  const sales = data.sales;
  assert.ok(sales, "sales pack is attached");
  const salesLerato = sales.people.find((p) => p.worker === "Lerato");
  assert.ok(salesLerato, "Lerato is on the sales pack");
  assert.ok(salesLerato.worked);
  assert.strictEqual(salesLerato.captured, 1);
  assert.strictEqual(salesLerato.quoted, 1);
  assert.strictEqual(salesLerato.ordered, 1);
  assert.strictEqual(sales.addedCount, 1);
  assert.strictEqual(sales.medianLagLabel, "2h");
  assert.strictEqual(sales.added[0].order, "S260910");
  assert.strictEqual(sales.added[0].lagLabel, "2h");
  assert.ok(data.clockedCount >= 1);
  assert.ok(data.absentCount >= 1);
  assert.ok(data.stations.some((s) => s.name === "Welding"));

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
  assert.ok(html.indexOf("Sales team") !== -1);
  assert.ok(html.indexOf("Median time to add") !== -1);
  assert.ok(html.indexOf("function openAdded") !== -1);
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
