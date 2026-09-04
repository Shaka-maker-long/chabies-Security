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
assert.deepStrictEqual(quiet.enquiryRoles || [], []);

const coster = staff.upsertUser({
  name: "Office Coster",
  access: "Admin",
  role: "Admin",
  password: "cost",
  seeDebtors: "Yes",
  enquiryRoles: ["Costing", "Quoting"]
});
assert.deepStrictEqual(coster.enquiryRoles, ["Costing", "Quoting"]);
assert.strictEqual(staff.defaultEnquiryAssignee("Costing"), "Office Coster");
assert.strictEqual(staff.defaultEnquiryAssignee("Quoting"), "Office Coster");
assert.strictEqual(staff.defaultEnquiryAssignee("Approval"), "");
const floorRoles = staff.upsertUser({
  name: "Floor Only",
  access: "Production",
  role: "Welding",
  password: "floor",
  tasks: ["Welding"],
  enquiryRoles: ["Costing"]
});
assert.deepStrictEqual(floorRoles.enquiryRoles, []);

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
  assert.strictEqual(login.jobTitle, "Welding");

  const officeLogin = await callShopFunction("verifyGlobalLogin", ["Office Quiet", "quiet"]);
  assert.strictEqual(officeLogin.success, true);
  assert.strictEqual(officeLogin.canSeeOffice, true);
  assert.strictEqual(officeLogin.canSeeDebtors, false);
  assert.strictEqual(officeLogin.jobTitle, "Admin");

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
  assert.strictEqual(boss.canManageUsers, true);
  assert.strictEqual(q.canManageUsers, false);

  const quietUsers = await fetch(base + "/api/office/users", { headers: { "x-sd-token": q.token } });
  const quietUsersJson = await quietUsers.json();
  assert.strictEqual(quietUsers.status, 200);
  assert.strictEqual(quietUsersJson.canManageUsers, false);
  assert.ok((quietUsersJson.rows || []).every((u) => u.name === "Office Quiet"));

  const quietPut = await fetch(base + "/api/office/users", {
    method: "PUT",
    headers: { "Content-Type": "application/json", "x-sd-token": q.token },
    body: JSON.stringify({ name: "Office Boss", access: "Admin", role: "Admin", password: "hacked" })
  });
  assert.strictEqual(quietPut.status, 403);

  const quietDel = await fetch(base + "/api/office/users/" + encodeURIComponent("Floor Only"), {
    method: "DELETE",
    headers: { "x-sd-token": q.token }
  });
  assert.strictEqual(quietDel.status, 403);

  const quietPass = await fetch(base + "/api/office/password", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": q.token },
    body: JSON.stringify({ current_password: "quiet", new_password: "quiet2" })
  });
  const quietPassJson = await quietPass.json();
  assert.ok(quietPassJson.ok, JSON.stringify(quietPassJson));
  assert.ok(staff.verifyUser("Office Quiet", "quiet2"));
  assert.ok(!staff.verifyUser("Office Quiet", "quiet"));

  const floorPass = await fetch(base + "/api/office/password", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: "Floor Only", current_password: "floor", new_password: "floor2" })
  });
  const floorPassJson = await floorPass.json();
  assert.ok(floorPassJson.ok, JSON.stringify(floorPassJson));
  assert.ok(staff.verifyUser("Floor Only", "floor2"));
  assert.ok(!staff.verifyUser("Floor Only", "floor"), "old access code must stop working");
  assert.ok(!staff.verifyUser("Floor Only", "admin"), "Manager access code must not log in someone else");

  const floorOldLogin = await callShopFunction("verifyGlobalLogin", ["Floor Only", "floor"]);
  assert.strictEqual(floorOldLogin.success, false);
  const floorMasterLogin = await callShopFunction("verifyGlobalLogin", ["Floor Only", "admin"]);
  assert.strictEqual(floorMasterLogin.success, false, JSON.stringify(floorMasterLogin));
  const floorNewLogin = await callShopFunction("verifyGlobalLogin", ["Floor Only", "floor2"]);
  assert.strictEqual(floorNewLogin.success, true);

  const managerReset = await fetch(base + "/api/office/password", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": boss.token },
    body: JSON.stringify({ name: "Floor Only", new_password: "floor3" })
  });
  const managerResetJson = await managerReset.json();
  assert.ok(managerResetJson.ok, JSON.stringify(managerResetJson));
  assert.ok(staff.verifyUser("Floor Only", "floor3"));
  assert.ok(!staff.verifyUser("Floor Only", "floor2"), "Manager reset must retire the previous code");
  assert.ok(staff.verifyUser("Office Boss", "admin"), "Manager reset must not change the Manager's own code");

  const quietResetOther = await fetch(base + "/api/office/password", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": q.token },
    body: JSON.stringify({ name: "Floor Only", new_password: "hacked" })
  });
  assert.strictEqual(quietResetOther.status, 403);
  assert.ok(staff.verifyUser("Floor Only", "floor3"));

  const bossUsers = await fetch(base + "/api/office/users", { headers: { "x-sd-token": boss.token } });
  const bossUsersJson = await bossUsers.json();
  assert.ok(bossUsersJson.canManageUsers);
  assert.ok((bossUsersJson.rows || []).some((u) => u.name === "Floor Only"));

  const viaCookie = staff.readSession({
    headers: { "x-sd-token": "dead-token", cookie: "sd_office=" + boss.token }
  });
  assert.ok(viaCookie, "a stale header must not hide a valid office cookie");
  assert.strictEqual(viaCookie.name, "Office Boss");

  const backup = await fetch(base + "/api/office/backup", { headers: { "x-sd-token": boss.token } });
  assert.strictEqual(backup.status, 200);
  const pack = await backup.json();
  assert.strictEqual(pack.database, "railway");
  assert.ok(pack.workbook && pack.workbook.sheets);

  const sqliteDl = await fetch(base + "/api/office/backup.db", { headers: { "x-sd-token": boss.token } });
  assert.strictEqual(sqliteDl.status, 200);
  const sqliteBuf = Buffer.from(await sqliteDl.arrayBuffer());
  assert.ok(sqliteBuf.length > 100);
  assert.strictEqual(sqliteBuf.slice(0, 6).toString("utf8"), "SQLite");

  const dbInfo = await fetch(base + "/api/office/database", { headers: { "x-sd-token": boss.token } });
  const info = await dbInfo.json();
  assert.strictEqual(info.live, "railway");
  assert.strictEqual(info.sheetsLive, false);
  assert.strictEqual(info.migrateAvailable, false);

  const migrate = await fetch(base + "/api/office/migrate-from-google", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": boss.token }
  });
  assert.strictEqual(migrate.status, 400);

  const sessFile = path.join(dir, "office-sessions.json");
  assert.ok(fs.existsSync(sessFile), "office sessions must be saved on disk");
  const sess = JSON.parse(fs.readFileSync(sessFile, "utf8"));
  assert.ok(sess[boss.token] && sess[boss.token].name === "Office Boss");

  const loggedOut = await fetch(base + "/api/office/logout", {
    method: "POST",
    headers: { "x-sd-token": boss.token }
  });
  assert.strictEqual(loggedOut.status, 200);
  const afterLogout = JSON.parse(fs.readFileSync(sessFile, "utf8"));
  assert.ok(!afterLogout[boss.token], "logout must drop the server session");
  const meGone = await fetch(base + "/api/office/me", { headers: { "x-sd-token": boss.token } });
  assert.strictEqual(meGone.status, 401);
  const ordersGone = await fetch(base + "/api/office/orders", { headers: { "x-sd-token": boss.token } });
  assert.strictEqual(ordersGone.status, 401);

  const bossAgain = await fetch(base + "/api/office/login", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: "Office Boss", password: "admin" })
  });
  const boss2 = await bossAgain.json();
  assert.ok(boss2.ok, JSON.stringify(boss2));
  assert.strictEqual(boss2.canManageUsers, true);
  assert.strictEqual(boss2.jobTitle, "Admin");

  const meBoss = await fetch(base + "/api/office/me", { headers: { "x-sd-token": boss2.token } });
  const meBossJson = await meBoss.json();
  assert.ok(meBossJson.ok);
  assert.strictEqual(meBossJson.profile.name, "Office Boss");
  assert.strictEqual(meBossJson.profile.jobTitle, "Admin");

  const promote = await fetch(base + "/api/office/users", {
    method: "PUT",
    headers: { "Content-Type": "application/json", "x-sd-token": boss2.token },
    body: JSON.stringify({ name: "Site Manager", access: "Production", role: "Manager", password: "mgr", seeDebtors: "Yes" })
  });
  const promoted = await promote.json();
  assert.ok(promoted.ok, JSON.stringify(promoted));
  assert.strictEqual(promoted.row.access, "Admin");
  assert.strictEqual(promoted.row.role, "Manager");
  assert.strictEqual(promoted.row.canManageUsers, true);

  const bossAfter = staff.readSession({ headers: { "x-sd-token": boss2.token } });
  assert.strictEqual(bossAfter.canManageUsers, false);

  const mgrLogin = await fetch(base + "/api/office/login", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: "Site Manager", password: "mgr" })
  });
  const mgr = await mgrLogin.json();
  assert.ok(mgr.ok, JSON.stringify(mgr));
  assert.strictEqual(mgr.canManageUsers, true);
  assert.strictEqual(mgr.jobTitle, "Manager");
  assert.strictEqual(mgr.access, "Admin");

  const bossPut = await fetch(base + "/api/office/users", {
    method: "PUT",
    headers: { "Content-Type": "application/json", "x-sd-token": boss2.token },
    body: JSON.stringify({ name: "Floor Only", access: "Production", role: "Welding" })
  });
  assert.strictEqual(bossPut.status, 403);

  const mgrPut = await fetch(base + "/api/office/users", {
    method: "PUT",
    headers: { "Content-Type": "application/json", "x-sd-token": mgr.token },
    body: JSON.stringify({ name: "Floor Only", access: "Production", role: "Welding" })
  });
  const mgrPutJson = await mgrPut.json();
  assert.ok(mgrPutJson.ok, JSON.stringify(mgrPutJson));

  server.close();
  console.log("staff-access.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
