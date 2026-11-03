"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-ops-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const db = require("./db");
const hub = require("./ops-hub");

initWorkbook();

assert.strictEqual(hub.quoteToken("soq 88"), "SOQ88");
assert.strictEqual(hub.quoteToken("22300"), "22300");
assert.deepStrictEqual(hub.quoteTokens("SOQ10 / SOQ11 & SOQ12"), ["SOQ10", "SOQ11", "SOQ12"]);
assert.strictEqual(hub.monthTag("January 2026"), "JAN 2026");
assert.strictEqual(hub.monthTag("2026-03-15"), "MAR 2026");

db.upsertEnquiry({
  enquiry_no: "E1001",
  date_enquired: "2026-01-10",
  month_enquired: "JAN 2026",
  client_name: "Ada",
  source: "Google",
  status: "Quoted",
  date_quoted: "2026-01-12",
  quote_no: "SOQ88"
});
db.upsertEnquiry({
  enquiry_no: "E1002",
  date_enquired: "2026-01-11",
  client_name: "Ben",
  source: "Instagram",
  status: "New"
});
db.upsertOrder({
  order_number: "S260001",
  quote_number: "SOQ88",
  client_name: "Ada",
  status: "Not Yet Started",
  price_excl_vat: "1000",
  payment_date: "2026-01-20",
  month_of_sale: "JAN 2026",
  source: "Google"
});
db.upsertOrder({
  order_number: "S260002",
  quote_number: "22300",
  client_name: "Web",
  status: "Not Yet Started",
  price_excl_vat: "500",
  payment_date: "2026-01-22",
  month_of_sale: "JAN 2026",
  source: "Website"
});
db.upsertOrder({
  order_number: "S260003",
  quote_number: "SOQ99",
  client_name: "Direct",
  status: "Cancelled",
  price_excl_vat: "9000",
  payment_date: "2026-01-23",
  month_of_sale: "JAN 2026"
});

const data = hub.build({ year: 2026 });
assert.strictEqual(data.year, 2026);
assert.strictEqual(data.enquiries.length, 2);
assert.strictEqual(data.orders.length, 2, "cancelled orders stay off the hub");
const ada = data.enquiries.find((e) => e.no === "E1001" || e.client === "Ada");
assert.ok(ada);
assert.ok(ada.is_q);
assert.ok(ada.isLinked);
const linked = data.orders.find((o) => o.no === "S260001");
assert.ok(linked.isLinked);
const web = data.orders.find((o) => o.no === "S260002");
assert.ok(!web.isLinked, "website quote with no enquiry is online/direct");
assert.strictEqual(linked.rev, 1000);

const html = fs.readFileSync(path.join(__dirname, "../public/ops-hub.html"), "utf8");
assert.ok(html.indexOf("2026 Linked conversion funnel") !== -1);
assert.ok(html.indexOf("MARKETING_STATS") === -1);
assert.ok(html.indexOf("/api/office/ops-hub") !== -1);
assert.ok(html.indexOf("Total Ad Spend") === -1);

console.log("ops-hub.test.js ok");
