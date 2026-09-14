const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const assert = require("assert");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-correct-ord-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const staff = require("./staff");
const db = require("./db");
const plan = require("./floor-planning");
const jobCard = require("./job-card");
const correct = require("./order-correct");
const { mountOffice } = require("./office");

initWorkbook();
staff.upsertUser({
  name: "Office Boss",
  access: "Admin",
  role: "Manager",
  password: "admin",
  seeDebtors: "Yes",
  canManageUsers: true
});
staff.upsertUser({
  name: "Willard",
  access: "Production",
  role: "Welder",
  password: "1234",
  tasks: ["Welding"]
});
staff.upsertUser({
  name: "Sam",
  access: "Production",
  role: "Cutter",
  password: "1234",
  tasks: ["Profile Cutting"]
});

db.upsertOrder({
  order_number: "S-FIX-1",
  product: "Air Chair",
  status: "Welding",
  assigned_operator: "Willard",
  client_name: "Wrong Start",
  type: "Standard",
  price_excl_vat: "1000.00"
});

assert.throws(() => correct.correctOrderShop({
  order_number: "S-FIX-1",
  status: "Ready for Steelwork",
  assigned_operator: "Sam"
}), /Tick that this is where the order is/);

const book = getBook();
const start = new Date("2026-09-14T06:00:00.000Z");
book.getSheetByName("Production_Log").appendRow([
  "log_fix_1", "S-FIX-1", "Willard", "Welding", "Welding", start, "", "", "", "", "", "", "{}"
]);
book.getSheetByName("Overview").appendRow([
  "log_fix_1", "S-FIX-1", "Willard", "Welding", start, "", ""
]);
persistWorkbook();

plan.save({
  blocks: [{ orderId: "S-FIX-1", process: "Welding", person: "Willard", start: "2026-09-14T07:45:00+02:00" }]
});

const moved = correct.correctOrderShop({
  order_number: "S-FIX-1",
  status: "Ready for Steelwork",
  assigned_operator: "Sam",
  confirmed: true
});
assert.strictEqual(moved.row.status, "Ready for Steelwork");
assert.strictEqual(moved.row.assigned_operator, "Sam");
assert.ok(moved.closedLogs >= 1, "open clock must close");
assert.ok(moved.planningRemoved >= 1, "wrong planning slots must drop");

const after = db.listOrders().find((o) => o.order_number === "S-FIX-1");
assert.strictEqual(after.status, "Ready for Steelwork");
assert.strictEqual(after.assigned_operator, "Sam");

const logEnd = book.getSheetByName("Production_Log").getRange(2, 7).getValue();
assert.ok(logEnd, "production log end must be written");
const overviewEnd = book.getSheetByName("Overview").getRange(2, 6).getValue();
assert.ok(overviewEnd, "overview end must be written");

const leftoverPlan = (plan.load().blocks || []).some((b) => b.orderId === "S-FIX-1");
assert.ok(!leftoverPlan, "planning for the corrected order must be empty");

const idleCut = correct.correctOrderShop({
  order_number: "S-FIX-1",
  status: "Profile Cutting",
  assigned_operator: "Sam",
  confirmed: true
});
assert.strictEqual(idleCut.row.status, "Ready for Steelwork", "Profile Cutting without a clock waits as Ready for Steelwork");

const locked = jobCard.applyOfficeOrderStatusLock({
  order_number: "S-FIX-1",
  status: "Delivered",
  assigned_operator: "Nobody"
}, after);
assert.strictEqual(locked.status, "Ready for Steelwork", "ordinary office edits still do not move shop status");

(async function main() {
  const app = express();
  app.use(express.json({ limit: "2mb" }));
  mountOffice(app);
  const server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const base = "http://127.0.0.1:" + server.address().port;
  const login = await fetch(base + "/api/office/login", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: "Office Boss", password: "admin" })
  });
  const session = await login.json();
  assert.ok(session.ok, JSON.stringify(session));

  const list = await fetch(base + "/api/office/orders", {
    headers: { "x-sd-token": session.token }
  });
  const listJson = await list.json();
  assert.ok(listJson.ok);
  assert.ok((listJson.shopStatuses || []).indexOf("Welding") !== -1);

  const refuse = await fetch(base + "/api/office/orders/correct-status", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({
      order_number: "S-FIX-1",
      status: "Assembly",
      assigned_operator: "Willard"
    })
  });
  assert.strictEqual(refuse.status, 400);

  const ok = await fetch(base + "/api/office/orders/correct-status", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({
      order_number: "S-FIX-1",
      status: "Ready for Assembly",
      assigned_operator: "",
      confirmed: true
    })
  });
  const okJson = await ok.json();
  assert.ok(okJson.ok, JSON.stringify(okJson));
  assert.strictEqual(okJson.row.status, "Ready for Assembly");
  assert.strictEqual(okJson.row.assigned_operator, "");

  const put = await fetch(base + "/api/office/orders", {
    method: "PUT",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({
      order_number: "S-FIX-1",
      status: "Delivered",
      assigned_operator: "Willard",
      product: "Air Chair",
      client_name: "Wrong Start"
    })
  });
  const putJson = await put.json();
  assert.ok(putJson.ok);
  assert.strictEqual(putJson.row.status, "Ready for Assembly");
  assert.strictEqual(putJson.row.assigned_operator, "Willard");

  server.close();
  console.log("order-correct.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
