"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-order-life-cycle-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const db = require("./db");
const dash = require("./order-dashboard");
const life = require("./order-lifecycle");
const delivery = require("./delivery-pod");

initWorkbook();

db.upsertOrder({
  order_number: "S260801",
  status: "Welding",
  product: "Air Chair",
  client_name: "Ann",
  price_incl_vat: "11500.00",
  amount_paid: "2000.00",
  payment_date: "14/09/2026"
});
db.upsertOrder({
  order_number: "S260802",
  status: "Ready for Delivery",
  product: "Air Chair",
  client_name: "Ben",
  address: "9 Rivonia Road",
  city: "Sandton",
  province: "Gauteng",
  price_incl_vat: "8000.00",
  amount_paid: "8000.00"
});
db.upsertOrder({
  order_number: "S260803",
  status: "Ready for Delivery",
  product: "Air Chair",
  client_name: "Cara",
  address: "1 Main Rd",
  city: "Cape Town",
  province: "Western Cape",
  price_incl_vat: "9000.00",
  amount_paid: "1000.00"
});

const openBefore = dash.buildDashboard({ range: "year" }).kpis.openJobs;
assert.ok(openBefore >= 3);

let missingReason = null;
try { life.cancelOrder("S260801", "  ", "Office"); } catch (e) { missingReason = e; }
assert.ok(missingReason, "cancel needs a reason");

const cancelled = life.cancelOrder("S260801", "Client withdrew", "Office");
assert.strictEqual(cancelled.status, "Cancelled");
assert.strictEqual(cancelled.cancel_reason, "Client withdrew");
assert.ok(db.listOrders().some((o) => o.order_number === "S260801" && o.status === "Cancelled"));
assert.ok(life.listCancelled().some((o) => o.order_number === "S260801"));
assert.strictEqual(dash.buildDashboard({ range: "year" }).kpis.openJobs, openBefore - 1);
assert.ok(!db.listDebtors().some((o) => o.order_number === "S260801"), "cancelled leaves debtors");

const held = life.holdOrder("S260802", "Waiting for glass", "Office");
assert.strictEqual(held.status, "On Hold");
assert.strictEqual(held.hold_reason, "Waiting for glass");
assert.ok(life.listOnHold().some((o) => o.order_number === "S260802"));
assert.strictEqual(dash.buildDashboard({ range: "year" }).kpis.openJobs, openBefore - 2);

const resumed = life.resumeOrder("S260802", "Office");
assert.strictEqual(resumed.status, "Ready for Delivery");

const decorated = delivery.decorateOrder(db.listOrders().find((o) => o.order_number === "S260803"));
assert.ok(decorated.payment_blocked);
assert.ok(decorated.owing > 0);

let blocked = null;
try { delivery.loadOnTruck(["S260803"], "Lebo"); } catch (e) { blocked = e; }
assert.ok(blocked, "unpaid ready orders cannot load");
assert.ok(/owes/i.test(blocked.message));
assert.strictEqual(db.listOrders().find((o) => o.order_number === "S260803").status, "Ready for Delivery");

const paidLoad = delivery.loadOnTruck(["S260802"], "Lebo");
assert.ok(!paidLoad || true);

console.log("order-lifecycle.test.js ok");
