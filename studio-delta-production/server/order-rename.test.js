"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-order-rename-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
process.env.SD_TRUST_DEVICES = "0";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const staff = require("./staff");
const comments = require("./order-cell-comments");
const noPlates = require("./no-plates");
const plan = require("./floor-planning");

initWorkbook();
staff.upsertUser({
  name: "Office Ada",
  access: "Admin",
  role: "Admin",
  password: "ada",
  seeDebtors: "Yes"
});

db.upsertOrder({
  order_number: "S260501",
  client_name: "Rename Me",
  product: "Air Chair",
  status: "Not Yet Started",
  amount_paid: "100.00",
  price_incl_vat: "1150.00",
  payment_date: "01/10/2026"
});

db.recordPayment("S260501", "50.00", "Deposit", {
  dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  filename: "proof.png",
  mime: "image/png"
});

comments.createComment({
  order_number: "S260501",
  field: "status",
  body: "Check status",
  assignee: "Office Ada"
}, "Office Ada");
noPlates.markNoPlate("S260501", "Plate");

const book = getBook();
book.getSheetByName("Production_Log").appendRow([
  "log1", "S260501", "Worker", "Welding", "Running", new Date(), "", "", "", "", "", "", "{}"
]);
persistWorkbook();

const planningPath = path.join(dir, "floor-planning.json");
fs.writeFileSync(planningPath, JSON.stringify({
  blocks: [{
    id: "b1",
    kind: "work",
    orderId: "S260501",
    studioNo: "S260501",
    process: "Welding",
    workerId: "Worker",
    workerName: "Worker"
  }],
  assignments: { S260501: { Welding: "Worker" } }
}));

const renamed = db.upsertOrder({
  order_number: "S260599",
  previous_order_number: "S260501",
  client_name: "Rename Me",
  product: "Air Chair",
  status: "Not Yet Started",
  amount_paid: "150.00",
  price_incl_vat: "1150.00",
  payment_date: "01/10/2026"
});

assert.strictEqual(renamed.order_number, "S260599");
assert.ok(!db.listOrders().some((o) => o.order_number === "S260501"), "old number leaves Orders");
const live = db.listOrders().find((o) => o.order_number === "S260599");
assert.ok(live, "new number is on Orders");
assert.strictEqual(live.amount_paid, "150.00");
assert.ok((live.payments || []).length >= 1, "payment history moves with the rename");
assert.ok(!comments.listComments({ order: "S260501" }).rows.length);
assert.ok(comments.listComments({ order: "S260599" }).rows.some((c) => c.field === "status"));
assert.ok(!noPlates.isNoPlate("S260501"));
assert.ok(noPlates.isNoPlate("S260599"));

const logVals = getBook().getSheetByName("Production_Log").getDataRange().getValues();
assert.ok(logVals.some((row) => String(row[1]) === "S260599"), "Production_Log order # updates");
assert.ok(!logVals.some((row, i) => i > 0 && String(row[1]) === "S260501"), "old log number is gone");

const planning = JSON.parse(fs.readFileSync(planningPath, "utf8"));
assert.ok(planning.blocks.some((b) => b.orderId === "S260599"), "planning blocks retarget");
assert.ok(planning.assignments.S260599, "planning assignments retarget");
assert.ok(!planning.assignments.S260501, "old planning assignment key is gone");
assert.strictEqual(plan.renameOrder("S260599", "S260599").changed, 0);

db.upsertOrder({
  order_number: "S260700",
  client_name: "Taken",
  status: "Not Yet Started"
});
assert.throws(
  () => db.renameOrderNumber("S260599", "S260700"),
  /already/
);

const same = db.renameOrderNumber("S260599", "S260599");
assert.strictEqual(same.order_number, "S260599");

const ordersHtml = fs.readFileSync(path.join(__dirname, "../public/orders.html"), "utf8");
assert.ok(ordersHtml.indexOf("previous_order_number") !== -1, "edit must send the old number so linked data can move");
assert.ok(ordersHtml.indexOf("readOnly = !!existing") === -1, "order number is no longer locked for every edit");
assert.ok(ordersHtml.indexOf("marketing && existing") !== -1, "marketing still cannot change order numbers");

console.log("order-rename.test.js: ok");
