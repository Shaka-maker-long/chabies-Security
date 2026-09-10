const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const assert = require("assert");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-onboard-ord-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const staff = require("./staff");
const db = require("./db");
const plan = require("./floor-planning");
const { remainingPlanForStatus } = require("./shop-status");
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
  role: "Cutter",
  password: "1234",
  tasks: ["Profile Cutting", "Plate Cutting"]
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
staff.setDurations([
  { product: "Air Chair", process: "Profile Cutting", hours: 1 },
  { product: "Air Chair", process: "Tagging", hours: 1 },
  { product: "Air Chair", process: "Plate Cutting", hours: 1 },
  { product: "Air Chair", process: "Welding", hours: 2 },
  { product: "Air Chair", process: "Grinding", hours: 1 },
  { product: "Air Chair", process: "Assembly", hours: 2 }
]);

const weld = remainingPlanForStatus("Welding");
assert.deepStrictEqual(weld.processes, ["Plate Cutting", "Grinding", "Assembly"]);
assert.strictEqual(weld.paintWait, true);
assert.ok(remainingPlanForStatus("Ready for Welding").processes.indexOf("Welding") !== -1);
assert.ok(remainingPlanForStatus("Ready for Assembly").processes.indexOf("Assembly") !== -1);
assert.strictEqual(remainingPlanForStatus("Ready for Assembly").paintWait, false);
assert.deepStrictEqual(remainingPlanForStatus("Final QC").processes, []);

assert.throws(() => db.onboardExistingOrder({
  order_number: "S260400",
  status: "Welding",
  product: "Air Chair",
  client_name: "Live Client",
  type: "Standard"
}), /Tick that this is where the order is/);

const created = db.onboardExistingOrder({
  order_number: "S260400",
  status: "Welding",
  product: "Air Chair",
  client_name: "Live Client",
  type: "Standard",
  province: "Gauteng",
  delivery_date: "2026-10-15",
  confirmed: true
});
assert.strictEqual(created.row.order_number, "S260400");
assert.strictEqual(created.row.status, "Welding");
assert.strictEqual(created.row.client_name, "Live Client");

const noTimes = db.onboardExistingOrder({
  order_number: "S260410",
  status: "Welding",
  product: "Unknown Gate",
  client_name: "No Times",
  type: "Standard",
  confirmed: true
});
assert.strictEqual(noTimes.row.status, "Welding");
const bare = plan.queueOrders().find((row) => row.order_number === "S260410");
assert.ok(bare, "in-progress orders stay in Planning even before Task times exist");
assert.ok(bare.processes.some((p) => p.process === "Grinding"));
assert.ok(bare.processes.every((p) => !(p.minutes > 0)));

const locked = require("./job-card").applyOfficeOrderStatusLock({
  order_number: "S260400",
  status: "Delivered",
  client_name: "Hack"
}, created.row);
assert.strictEqual(locked.status, "Welding", "ordinary office edits must not move shop status");

const queue = plan.queueOrders();
const q = queue.find((row) => row.order_number === "S260400");
assert.ok(q, "in-progress orders stay in Planning for remaining work");
assert.ok(!q.processes.some((p) => p.process === "Profile Cutting" || p.process === "Welding"));
assert.ok(q.processes.some((p) => p.process === "Grinding"));

plan.save({ blocks: [] });
const booked = plan.scheduleSelected({
  orderIds: ["S260400"],
  assignments: {
    S260400: { "Plate Cutting": "Willard", Assembly: "Nomsa" }
  },
  from: "2026-09-08T07:45:00+02:00"
});
assert.ok(!booked.blocks.some((b) => b.process === "Profile Cutting" || b.process === "Tagging" || b.process === "Welding"));
assert.ok(booked.blocks.some((b) => b.process === "Grinding"), "remaining grinding is auto-booked");
assert.ok(booked.blocks.some((b) => b.process === "Powder coating"));
assert.ok(booked.blocks.some((b) => b.process === "Assembly"));

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
  const meta = await fetch(base + "/api/office/orders/onboard", {
    headers: { "x-sd-token": session.token }
  });
  const metaJson = await meta.json();
  assert.ok(metaJson.ok);
  assert.ok(metaJson.statuses.indexOf("Welding") !== -1);
  assert.ok(metaJson.remaining.Welding.processes.indexOf("Grinding") !== -1);

  const refuse = await fetch(base + "/api/office/orders/onboard", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({
      order_number: "S260400",
      status: "Welding",
      product: "Air Chair",
      client_name: "Dup",
      type: "Standard",
      confirmed: true
    })
  });
  assert.strictEqual(refuse.status, 400);

  const again = await fetch(base + "/api/office/orders/onboard", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({
      order_number: "S260401",
      status: "Ready for Assembly",
      product: "Air Chair",
      client_name: "Paint Back",
      type: "Standard",
      confirmed: true
    })
  });
  const againJson = await again.json();
  assert.ok(againJson.ok, JSON.stringify(againJson));
  assert.deepStrictEqual(againJson.remaining.processes, ["Assembly"]);
  assert.strictEqual(againJson.remaining.paintWait, false);

  const put = await fetch(base + "/api/office/orders", {
    method: "PUT",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({
      order_number: "S260401",
      status: "Delivered",
      product: "Air Chair",
      client_name: "Paint Back"
    })
  });
  const putJson = await put.json();
  assert.ok(putJson.ok);
  assert.strictEqual(putJson.row.status, "Ready for Assembly");

  server.close();
  console.log("order-onboard.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
