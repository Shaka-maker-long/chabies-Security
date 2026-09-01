const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const assert = require("assert");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-staff-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const staff = require("./staff");
const { mountOffice } = require("./office");
const { callShopFunction } = require("./gas");

initWorkbook();

staff.upsertUser({
  name: "Floor Only",
  access: "Production",
  role: "Welding",
  password: "floor",
  tasks: ["Welding"]
});
staff.upsertUser({
  name: "Office Boss",
  access: "Admin",
  role: "Admin",
  password: "admin",
  seeDebtors: "Yes"
});
staff.upsertUser({
  name: "Office Quiet",
  access: "Admin",
  role: "Admin",
  password: "quiet",
  seeDebtors: "No"
});
staff.setDurations([{ product: "Slider", process: "Welding", minutes: 45 }]);

const floor = staff.verifyUser("Floor Only", "floor");
assert.ok(floor);
assert.strictEqual(floor.canSeeOffice, false);
assert.strictEqual(floor.canSeeDebtors, false);
assert.ok(floor.tasks.indexOf("Welding") !== -1);

const productionTitledAdmin = staff.upsertUser({
  name: "Named Admin Job",
  access: "Production",
  role: "Admin",
  password: "x"
});
assert.strictEqual(productionTitledAdmin.access, "Production");
assert.strictEqual(productionTitledAdmin.canSeeOffice, false);

const quiet = staff.verifyUser("Office Quiet", "quiet");
assert.ok(quiet);
assert.strictEqual(quiet.canSeeOffice, true);
assert.strictEqual(quiet.canSeeDebtors, false);

assert.strictEqual(staff.durationMinutes("Slider", "Welding"), 45);
assert.strictEqual(staff.durationMinutes("Slider", "Painting"), 0);

const now = Date.parse("2026-08-27T10:00:00+02:00");
const started = Date.parse("2026-08-27T09:30:00+02:00");
assert.strictEqual(
  staff.countdownRemainingMs({ targetMinutes: 45, startedAt: started, pauseMs: 0, isPaused: false }, now),
  15 * 60 * 1000
);
const pausedAt = Date.parse("2026-08-27T09:40:00+02:00");
assert.strictEqual(
  staff.countdownRemainingMs({
    targetMinutes: 45,
    startedAt: started,
    pauseMs: 0,
    isPaused: true,
    pausedAt
  }, now),
  35 * 60 * 1000
);
assert.strictEqual(staff.countdownRemainingMs({ targetMinutes: 0, startedAt: started }, now), null);
assert.strictEqual(staff.countdownRemainingMs({ targetMinutes: 10 }, now), null);

(async function main() {
  const login = await callShopFunction("verifyGlobalLogin", ["Floor Only", "floor"]);
  assert.strictEqual(login.success, true, JSON.stringify(login));
  assert.strictEqual(login.canSeeOffice, false);
  assert.strictEqual(login.canSeeDebtors, false);
  assert.ok(login.tasks.indexOf("Welding") !== -1);

  const officeLogin = await callShopFunction("verifyGlobalLogin", ["Office Quiet", "quiet"]);
  assert.strictEqual(officeLogin.success, true);
  assert.strictEqual(officeLogin.canSeeOffice, true);
  assert.strictEqual(officeLogin.canSeeDebtors, false);

  const dur = await callShopFunction("getTaskDuration", ["Slider", "Welding"]);
  assert.strictEqual(dur.minutes, 45);

  const app = express();
  app.use(express.json());
  mountOffice(app);
  const server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const base = "http://127.0.0.1:" + server.address().port;

  const blocked = await fetch(base + "/api/office/login", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: "Floor Only", password: "floor" })
  });
  assert.strictEqual(blocked.status, 403);

  const noAuth = await fetch(base + "/api/office/orders");
  assert.strictEqual(noAuth.status, 401);

  const bossLogin = await fetch(base + "/api/office/login", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: "Office Boss", password: "admin" })
  });
  const boss = await bossLogin.json();
  assert.ok(boss.ok && boss.token, JSON.stringify(boss));
  assert.strictEqual(boss.canSeeDebtors, true);

  const quietLogin = await fetch(base + "/api/office/login", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: "Office Quiet", password: "quiet" })
  });
  const q = await quietLogin.json();
  assert.ok(q.ok, JSON.stringify(q));
  assert.strictEqual(q.canSeeDebtors, false);

  const debtorsQuiet = await fetch(base + "/api/office/debtors", { headers: { "x-sd-token": q.token } });
  assert.strictEqual(debtorsQuiet.status, 403);

  const debtorsBoss = await fetch(base + "/api/office/debtors", { headers: { "x-sd-token": boss.token } });
  assert.strictEqual(debtorsBoss.status, 200);

  const ordersBoss = await fetch(base + "/api/office/orders", { headers: { "x-sd-token": boss.token } });
  assert.strictEqual(ordersBoss.status, 200);

  server.close();
  console.log("staff-access.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
