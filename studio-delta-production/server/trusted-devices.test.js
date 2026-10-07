const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");
const http = require("http");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-devices-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
process.env.DEVICE_BOOTSTRAP_CODE = "first-device-unlock";
delete process.env.SD_TRUST_DEVICES;
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const staff = require("./staff");
const { mountOffice } = require("./office");
const devices = require("./trusted-devices");

initWorkbook();
staff.upsertUser({
  name: "Office Boss",
  access: "Admin",
  role: "Manager",
  password: "admin",
  seeDebtors: "Yes"
});
staff.upsertUser({
  name: "Sam Floor",
  access: "Admin",
  role: "Admin",
  password: "sam1",
  seeDebtors: "No"
});

// No auto-approve: first login without bootstrap is blocked
const blocked = devices.assertLoginAllowed({
  deviceId: "d_laptopabc1234567890",
  userName: "Office Boss",
  userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
  ip: "127.0.0.1",
  canManageUsers: true
});
assert.strictEqual(blocked.ok, false);
assert.strictEqual(blocked.pending, true);
assert.strictEqual(blocked.needsBootstrap, true);

// Wrong bootstrap code is rejected
const wrongBoot = devices.assertLoginAllowed({
  deviceId: "d_laptopabc1234567890",
  userName: "Office Boss",
  userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
  ip: "127.0.0.1",
  canManageUsers: true,
  bootstrapCode: "nope"
});
assert.strictEqual(wrongBoot.ok, false);

// Manager unlocks the first device with DEVICE_BOOTSTRAP_CODE
const first = devices.assertLoginAllowed({
  deviceId: "d_laptopabc1234567890",
  userName: "Office Boss",
  userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
  ip: "127.0.0.1",
  canManageUsers: true,
  bootstrapCode: "first-device-unlock"
});
assert.strictEqual(first.ok, true);
assert.strictEqual(first.bootstrapped, true);
assert.strictEqual(first.device.assignedTo, "Office Boss");

// Sam already has the app installed but is not linked → blocked until Manager approves
const samFirst = devices.assertLoginAllowed({
  deviceId: "d_phonexyz12345678901",
  userName: "Sam Floor",
  userAgent: "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 Safari/604.1",
  ip: "10.0.0.2"
});
assert.strictEqual(samFirst.ok, false);
assert.strictEqual(samFirst.pending, true);
assert.ok((samFirst.device.pendingUsers || []).some((p) => p.name === "Sam Floor")
  || samFirst.device.requestedBy === "Sam Floor");

// Non-manager cannot bootstrap after a device already exists
const samBoot = devices.assertLoginAllowed({
  deviceId: "d_phonexyz12345678901",
  userName: "Sam Floor",
  userAgent: "Mozilla/5.0 (iPhone)",
  ip: "10.0.0.2",
  canManageUsers: false,
  bootstrapCode: "first-device-unlock"
});
assert.strictEqual(samBoot.ok, false);

// Second Manager with no personal device can still bootstrap even though Sam's phone exists
staff.upsertUser({
  name: "New Manager",
  access: "Admin",
  role: "Manager",
  password: "mgr1",
  seeDebtors: "Yes"
});
const mgrLockout = devices.assertLoginAllowed({
  deviceId: "d_newmgrphoneabcdefghij",
  userName: "New Manager",
  userAgent: "Mozilla/5.0 (iPhone)",
  ip: "10.0.0.9",
  canManageUsers: true
});
assert.strictEqual(mgrLockout.ok, false);
assert.strictEqual(mgrLockout.needsBootstrap, true);

const mgrUnlock = devices.assertLoginAllowed({
  deviceId: "d_newmgrphoneabcdefghij",
  userName: "New Manager",
  userAgent: "Mozilla/5.0 (iPhone)",
  ip: "10.0.0.9",
  canManageUsers: true,
  bootstrapCode: "first-device-unlock"
});
assert.strictEqual(mgrUnlock.ok, true);
assert.strictEqual(mgrUnlock.bootstrapped, true);

let snap = devices.approveDevice("d_phonexyz12345678901", "Office Boss", {
  assignedTo: "Sam Floor",
  nickname: "Sam iPhone"
});
assert.ok(snap.approved.some((row) => row.id === "d_phonexyz12345678901" && row.assignedTo === "Sam Floor"));

const samOk = devices.assertLoginAllowed({
  deviceId: "d_phonexyz12345678901",
  userName: "Sam Floor",
  userAgent: "Mozilla/5.0 (iPhone)",
  ip: "10.0.0.2"
});
assert.strictEqual(samOk.ok, true);

// Sam tries Boss laptop → Manager must approve Sam on that device
const cross = devices.assertLoginAllowed({
  deviceId: "d_laptopabc1234567890",
  userName: "Sam Floor",
  userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
  ip: "10.0.0.3"
});
assert.strictEqual(cross.ok, false);
assert.strictEqual(cross.pending, true);
assert.ok((cross.device.pendingUsers || []).some((p) => p.name === "Sam Floor"));

// Boss cannot use Sam phone without approval
const bossOnSamPhone = devices.assertLoginAllowed({
  deviceId: "d_phonexyz12345678901",
  userName: "Office Boss",
  userAgent: "Mozilla/5.0 (iPhone)",
  ip: "127.0.0.1"
});
assert.strictEqual(bossOnSamPhone.ok, false);

snap = devices.approveDevice("d_laptopabc1234567890", "Office Boss", {
  assignedTo: "Sam Floor",
  nickname: "Office laptop"
});
assert.ok(snap.approved.some((row) => {
  return row.id === "d_laptopabc1234567890"
    && row.nickname === "Office laptop"
    && row.assignedUsers.indexOf("Office Boss") !== -1
    && row.assignedUsers.indexOf("Sam Floor") !== -1;
}));

const samOnLaptop = devices.assertLoginAllowed({
  deviceId: "d_laptopabc1234567890",
  userName: "Sam Floor",
  userAgent: "Mozilla/5.0 (Windows NT 10.0)",
  ip: "10.0.0.3"
});
assert.strictEqual(samOnLaptop.ok, true);

// Sam's second device needs approval
const secondPhone = devices.assertLoginAllowed({
  deviceId: "d_phone2abcdefghijklmn",
  userName: "Sam Floor",
  userAgent: "Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) Safari/604.1",
  ip: "10.0.0.8"
});
assert.strictEqual(secondPhone.ok, false);
assert.strictEqual(secondPhone.pending, true);

snap = devices.approveDevice("d_phone2abcdefghijklmn", "Office Boss", {
  assignedTo: "Sam Floor",
  nickname: "Sam spare iPhone"
});
assert.ok(snap.approved.some((row) => row.id === "d_phone2abcdefghijklmn" && row.assignedTo === "Sam Floor"));

assert.strictEqual(devices.sessionStillValid("Sam Floor", "d_phone2abcdefghijklmn").ok, true);

snap = devices.revokeDevice("d_phone2abcdefghijklmn", "Office Boss");
assert.ok(snap.revoked.some((row) => row.id === "d_phone2abcdefghijklmn"));
assert.strictEqual(devices.assertLoginAllowed({
  deviceId: "d_phone2abcdefghijklmn",
  userName: "Sam Floor",
  userAgent: "Mozilla/5.0 (iPhone)",
  ip: "10.0.0.8"
}).ok, false);
assert.strictEqual(devices.sessionStillValid("Sam Floor", "d_phone2abcdefghijklmn").ok, false);

(async function main() {
  const app = express();
  app.use(express.json({ limit: "2mb" }));
  mountOffice(app);
  const server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;

  async function api(pathname, opts) {
    const r = await fetch("http://127.0.0.1:" + port + pathname, opts);
    const j = await r.json();
    return { status: r.status, json: j };
  }

  const login = await api("/api/office/login", {
    method: "POST",
    headers: { "content-type": "application/json", "x-sd-device-id": "d_newmgrphoneabcdefghij" },
    body: JSON.stringify({
      name: "New Manager",
      password: "mgr1",
      deviceId: "d_newmgrphoneabcdefghij"
    })
  });
  assert.strictEqual(login.json.ok, true, JSON.stringify(login.json));
  assert.ok(login.json.deviceId || (login.json.device && login.json.device.id));
  const token = login.json.token;
  const headers = { "content-type": "application/json", "x-sd-token": token, "x-sd-device-id": "d_newmgrphoneabcdefghij" };

  const pendingLogin = await api("/api/office/login", {
    method: "POST",
    headers: { "content-type": "application/json", "x-sd-device-id": "d_tablet111222333444" },
    body: JSON.stringify({ name: "Sam Floor", password: "sam1", deviceId: "d_tablet111222333444" })
  });
  assert.strictEqual(pendingLogin.json.ok, false, "Sam tablet needs Manager approval");
  assert.strictEqual(pendingLogin.json.pendingDevice, true);

  const listed = await api("/api/office/devices", { headers });
  assert.strictEqual(listed.json.ok, true);
  assert.ok((listed.json.people || []).indexOf("Sam Floor") !== -1);

  const approved = await api("/api/office/devices/d_tablet111222333444/approve", {
    method: "POST",
    headers,
    body: JSON.stringify({ assignedTo: "Sam Floor", nickname: "Sam tablet" })
  });
  assert.strictEqual(approved.json.ok, true, JSON.stringify(approved.json));
  const tablet = (approved.json.approved || []).find((row) => row.id === "d_tablet111222333444");
  assert.strictEqual(tablet.assignedTo, "Sam Floor");
  assert.strictEqual(tablet.nickname, "Sam tablet");

  const second = await api("/api/office/login", {
    method: "POST",
    headers: { "content-type": "application/json", "x-sd-device-id": "d_tablet111222333444" },
    body: JSON.stringify({ name: "Sam Floor", password: "sam1", deviceId: "d_tablet111222333444" })
  });
  assert.strictEqual(second.json.ok, true, JSON.stringify(second.json));

  const pin = await api("/api/office/devices/d_tablet111222333444/driver-phone", {
    method: "POST",
    headers,
    body: JSON.stringify({ on: true })
  });
  assert.strictEqual(pin.json.ok, true, JSON.stringify(pin.json));
  const pinnedRow = (pin.json.approved || []).find((row) => row.id === "d_tablet111222333444");
  assert.ok(pinnedRow && pinnedRow.driverPhone, "Manager can pin an approved device as the driver phone");
  assert.ok(devices.driverPhoneByDeviceId("d_tablet111222333444"));

  const unpin = await api("/api/office/devices/d_tablet111222333444/driver-phone", {
    method: "POST",
    headers,
    body: JSON.stringify({ on: false })
  });
  assert.strictEqual(unpin.json.ok, true);
  assert.ok(!devices.driverPhoneByDeviceId("d_tablet111222333444"));

  const bossOnTablet = await api("/api/office/login", {
    method: "POST",
    headers: { "content-type": "application/json", "x-sd-device-id": "d_tablet111222333444" },
    body: JSON.stringify({ name: "New Manager", password: "mgr1", deviceId: "d_tablet111222333444" })
  });
  assert.strictEqual(bossOnTablet.json.ok, false);
  assert.strictEqual(bossOnTablet.json.pendingDevice, true);

  const audit = await api("/api/office/audit", { headers });
  assert.strictEqual(audit.json.ok, true);
  assert.ok((audit.json.rows || []).length >= 1);

  server.close();
  console.log("trusted-devices.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
