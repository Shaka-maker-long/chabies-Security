"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-inprog-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
process.env.WORK_LOCKS_DISABLED = "true";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const staff = require("./staff");
const { callShopFunction } = require("./gas");
const { waitingStatusIfProfileCuttingIdle } = require("./shop-status");
const inProgress = require("./in-progress-status");

assert.strictEqual(waitingStatusIfProfileCuttingIdle("Profile Cutting", false), "Ready for Steelwork");
assert.strictEqual(waitingStatusIfProfileCuttingIdle("Profile Cutting", true), "Profile Cutting");
assert.strictEqual(waitingStatusIfProfileCuttingIdle("Not Yet Started", false), "Not Yet Started");

initWorkbook();
staff.upsertUser({
  name: "Sam",
  access: "Production",
  role: "Profile Cutting",
  password: "1234",
  tasks: ["Profile Cutting"]
});

db.upsertOrder({
  order_number: "S-WAIT-1",
  status: "Profile Cutting",
  product: "Air Chair",
  client_name: "Waiting"
});
db.upsertOrder({
  order_number: "S-LIVE-1",
  status: "Ready for Steelwork",
  product: "Air Chair",
  client_name: "Live"
});
persistWorkbook();

const CONFIRM = { understood: true, highlights: [] };

(async function main() {
  const first = inProgress.reconcile();
  assert.ok(first.updated >= 1, JSON.stringify(first));
  assert.strictEqual(db.listOrders().find((o) => o.order_number === "S-WAIT-1").status, "Ready for Steelwork");

  await callShopFunction("grantOvertime", ["Sam", "", "Admin", "test"]);
  const live = db.listOrders().find((o) => o.order_number === "S-LIVE-1");
  const start = await callShopFunction("startOrder", [live.id, "Sam", "Profile Cutting", [], "", false, null, CONFIRM]);
  assert.strictEqual(start.success, true, JSON.stringify(start));
  assert.strictEqual(start.newStatus, "Profile Cutting");
  const afterStart = inProgress.reconcile();
  assert.strictEqual(afterStart.updated, 0, "a live clock must keep Profile Cutting");
  assert.strictEqual(db.listOrders().find((o) => o.order_number === "S-LIVE-1").status, "Profile Cutting");

  const poll = await callShopFunction("pollFloor", ["Profile Cutting", "Sam"]);
  const liveCard = (poll.orders || []).find((o) => o.order === "S-LIVE-1");
  const waitCard = (poll.orders || []).find((o) => o.order === "S-WAIT-1");
  assert.ok(liveCard && liveCard.logId && liveCard.assigned === "Sam", JSON.stringify(liveCard));
  assert.strictEqual(liveCard.status, "Profile Cutting");
  assert.ok(waitCard, "waiting steelwork stays on the cutter list");
  assert.strictEqual(waitCard.status, "Ready for Steelwork");
  assert.ok(!waitCard.logId, "waiting job has no clock");

  const floor = fs.readFileSync(path.join(__dirname, "../index.html"), "utf8");
  assert.ok(floor.indexOf("function displayShopStatus") !== -1);
  assert.ok(floor.indexOf("new orders") !== -1);

  console.log("in-progress-status.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
