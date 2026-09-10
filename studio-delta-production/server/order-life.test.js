"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-life-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook } = require("./workbook-store");
const db = require("./db");
const life = require("./order-life");

initWorkbook();

const enquiry = db.upsertEnquiry({
  date_enquired: "03/03/2026",
  enquiry_source: "Website",
  enquiry_type: "Standard",
  client_name: "Journey Client",
  client_email: "journey@example.com",
  client_number: "0820000000",
  product: "Air Chair",
  status: "Ordered",
  province: "Gauteng"
});
assert.ok(enquiry.events.some((ev) => /captured|created/i.test(ev.label || ev.kind)));

const order = db.upsertOrder({
  order_number: "S260900",
  status: "Profile Cutting",
  client_name: "Journey Client",
  product: "Air Chair",
  enquiry_no: enquiry.enquiry_no,
  payment_date: "2026-03-10",
  type: "Standard",
  category: "Chair"
});
assert.strictEqual(order.order_number, "S260900");

fs.writeFileSync(path.join(dir, "job-cards.json"), JSON.stringify({
  S260900: {
    order_number: "S260900",
    created_at: "2026-03-11T06:00:00.000Z",
    created_date: "2026-03-11",
    product: "Air Chair"
  }
}));

getBook().getSheetByName("Production_Log").appendRow([
  "log_cut_life",
  "S260900",
  "Willard",
  "Profile Cutting",
  "Done",
  new Date("2026-03-12T05:45:00.000Z"),
  new Date("2026-03-12T06:45:00.000Z"),
  "",
  "",
  "",
  0,
  "",
  JSON.stringify({
    pauses: [{ start: "2026-03-12T10:00:00+02:00", end: "2026-03-12T10:30:00+02:00", reason: "Lunch" }]
  })
]);

getBook().getSheetByName("Wood_To_Order").appendRow([
  "wood_1",
  new Date("2026-03-13T08:00:00.000Z"),
  "S260900",
  "Thabo",
  "Seat",
  "Oak",
  "16",
  400,
  400,
  1,
  "To order"
]);

const snap = life.collectOrderLife("S260900");
assert.strictEqual(snap.order_number, "S260900");
assert.strictEqual(snap.enquiry_no, enquiry.enquiry_no);
assert.ok(snap.events.length >= 4, "enquiry, order, job card, and shop clock should all appear");

const titles = snap.events.map((ev) => ev.title);
assert.ok(titles.some((t) => /enquiry captured/i.test(t)), "starts with the enquiry");
assert.ok(titles.some((t) => /order opened/i.test(t)), "order opened is on the path");
assert.ok(titles.some((t) => /job card generated/i.test(t)), "job card is on the path");
assert.ok(titles.some((t) => /started profile cutting/i.test(t)), "shop start is on the path");
assert.ok(titles.some((t) => /paused profile cutting/i.test(t)), "pauses are on the path");
assert.ok(titles.some((t) => /finished profile cutting/i.test(t)), "shop finish is on the path");
assert.ok(titles.some((t) => /wood logged/i.test(t)), "wood logged is on the path");

const start = snap.events.find((ev) => /started profile cutting/i.test(ev.title));
assert.strictEqual(start.actor, "Willard");
assert.strictEqual(start.stage, "shop");
assert.ok(start.at);

const first = snap.events[0];
const last = snap.events[snap.events.length - 1];
assert.ok(String(first.at) <= String(last.at), "events are chronological");

const fromEnquiry = life.collectEnquiryLife(enquiry.enquiry_no);
assert.ok(fromEnquiry.order_numbers.indexOf("S260900") !== -1);
assert.ok(fromEnquiry.events.some((ev) => /started profile cutting/i.test(ev.title)));

assert.throws(() => life.collectOrderLife("S999999"), /not found/i);
assert.throws(() => life.collectEnquiryLife("#1"), /not found/i);

const missing = db.upsertOrder({
  order_number: "S260901",
  status: "Not Yet Started",
  client_name: "Empty Path",
  product: "Air Chair"
});
const empty = life.collectOrderLife(missing.order_number);
assert.ok(Array.isArray(empty.events));

console.log("order-life.test.js ok");
