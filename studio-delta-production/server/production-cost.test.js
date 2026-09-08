const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const assert = require("assert");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-cost-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const staff = require("./staff");
const db = require("./db");
const steelRates = require("./steel-rates");
const cost = require("./production-cost");
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

db.upsertOrder({
  order_number: "S-COST-1",
  product: "Slider",
  status: "Welding",
  price_excl_vat: "2000.00"
});

cost.upsertLabourRate({ process: "Welding", ratePerHour: "200" });
steelRates.upsertRate({ type: "25x25x2", ratePerM: "80" });
assert.strictEqual(steelRates.costUsage("Tube - 25x25x2", 3).cost, 240);

const book = getBook();
const start = new Date("2026-09-09T06:00:00.000Z"); // 08:00 SAST
const end = new Date("2026-09-09T08:00:00.000Z"); // 10:00 SAST
book.getSheetByName("Production_Log").appendRow([
  "log_cost_1", "S-COST-1", "Willard", "Welding", "Welding", start, end, "", "", "", 0, "", ""
]);
book.getSheetByName("Steel_Usage").appendRow([
  start, "S-COST-1", "Willard", "Welding", "Tube - 25x25x2", 3
]);
persistWorkbook();

assert.strictEqual(cost.matchTask("Welding"), "Welding");
assert.strictEqual(cost.matchTask("Profile Cutter"), "Profile Cutting");
assert.strictEqual(cost.matchTask("Final QC"), "");

(async function main() {
  const data = await cost.getAppData({ mode: "all" });
  const order = data.orders.find((o) => o.orderNum === "S-COST-1");
  assert.ok(order, JSON.stringify(data.orders));
  assert.ok(Math.abs(order.totalHours - 2) < 0.05, "hours " + order.totalHours);
  assert.ok(Math.abs(order.laborCost - 400) < 0.05, "labour " + order.laborCost);
  assert.strictEqual(order.materialCost, 240);
  assert.ok(order.staff.Willard);
  assert.ok(Math.abs(order.tasks.Welding.h - 2) < 0.05);
  assert.ok(Math.abs(order.overheadCost - 800) < 0.05, "overhead " + order.overheadCost);
  assert.strictEqual(order.sellingPrice, 2000);

  const sept = await cost.getAppData({ mode: "month", from: "2026-09" });
  assert.ok(sept.orders.some((o) => o.orderNum === "S-COST-1"));
  const jan = await cost.getAppData({ mode: "month", from: "2026-01" });
  assert.ok(!jan.orders.some((o) => o.laborCost > 0 && o.orderNum === "S-COST-1") || jan.orders.length === 0);

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
  const snap = await fetch(base + "/api/office/production-cost?mode=all", {
    headers: { "x-sd-token": session.token }
  });
  const snapJson = await snap.json();
  assert.ok(snapJson.ok, JSON.stringify(snapJson));
  assert.ok((snapJson.orders || []).some((o) => o.orderNum === "S-COST-1" && o.materialCost === 240));
  const saved = await fetch(base + "/api/office/labour-rates", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({ process: "Profile Cutting", ratePerHour: "150" })
  });
  const savedJson = await saved.json();
  assert.ok(savedJson.ok, JSON.stringify(savedJson));
  assert.ok((savedJson.rates || []).some((r) => r.process === "Profile Cutting"));
  const steel = await fetch(base + "/api/office/steel-rates", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({ type: "38x2", ratePerM: "95" })
  });
  const steelJson = await steel.json();
  assert.ok(steelJson.ok, JSON.stringify(steelJson));

  const runningStart = new Date("2026-09-09T09:00:00.000Z");
  book.getSheetByName("Production_Log").appendRow([
    "log_open_1", "S-COST-OPEN", "Willard", "Plate Cutting", "Plate Cutting", runningStart, "", "", "", "", 0, "", ""
  ]);
  book.getSheetByName("Production_Log").appendRow([
    "log_zero_1", "S-COST-ZERO", "Willard", "Profile Cutting", "Profile Cutting", runningStart, runningStart, "", "", "", 0, "", ""
  ]);
  persistWorkbook();
  const openSnap = await cost.getAppData({ mode: "all" });
  assert.ok(!(openSnap.orders || []).some((o) => o.orderNum === "S-COST-OPEN"), "open jobs must not stay on the cost matrix");
  assert.ok(!(openSnap.orders || []).some((o) => o.orderNum === "S-COST-ZERO"), "zero-minute starts must not stay on the cost matrix");

  const refuseSteel = await fetch(base + "/api/office/steel-usage/clear", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({ confirm: "no" })
  });
  assert.strictEqual(refuseSteel.status, 400);

  const wipeSteel = await fetch(base + "/api/office/steel-usage/clear", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({ confirm: "CLEAR" })
  });
  const wipeSteelJson = await wipeSteel.json();
  assert.ok(wipeSteelJson.ok, JSON.stringify(wipeSteelJson));
  assert.ok(wipeSteelJson.removed >= 1);
  assert.strictEqual(book.getSheetByName("Steel_Usage").getLastRow(), 1);

  const wipeLogs = await fetch(base + "/api/office/production-log/clear", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({ confirm: "CLEAR" })
  });
  const wipeLogsJson = await wipeLogs.json();
  assert.ok(wipeLogsJson.ok, JSON.stringify(wipeLogsJson));
  assert.ok(wipeLogsJson.removed >= 1);
  assert.strictEqual(book.getSheetByName("Production_Log").getLastRow(), 1, "every production log row must go");

  const after = await fetch(base + "/api/office/production-cost?mode=all", {
    headers: { "x-sd-token": session.token }
  });
  const afterJson = await after.json();
  assert.ok(afterJson.ok);
  const leftover = (afterJson.orders || []).find((o) => o.orderNum === "S-COST-1");
  assert.ok(!leftover, "cleared labour and steel must leave the cost matrix");

  server.close();
  console.log("production-cost.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
