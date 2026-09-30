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
  password: "sam",
  seeDebtors: "No"
});

const first = devices.assertLoginAllowed({
  deviceId: "d_laptopabc1234567890",
  userName: "Office Boss",
  userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
  ip: "127.0.0.1"
});
assert.strictEqual(first.ok, true);
assert.strictEqual(first.bootstrapped, true);
assert.strictEqual(first.device.assignedTo, "Office Boss");

// Sam's first phone auto-assigns to Sam
const samFirst = devices.assertLoginAllowed({
  deviceId: "d_phonexyz12345678901",
  userName: "Sam Floor",
  userAgent: "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 Safari/604.1",
  ip: "10.0.0.2"
});
assert.strictEqual(samFirst.ok, true);
assert.strictEqual(samFirst.bootstrapped, true);
assert.strictEqual(samFirst.device.assignedTo, "Sam Floor");

// Sam tries Boss laptop (different device) → Manager must approve
const cross = devices.assertLoginAllowed({
  deviceId: "d_laptopabc1234567890",
  userName: "Sam Floor",
  userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
  ip: "10.0.0.3"
});
assert.strictEqual(cross.ok, false);
assert.strictEqual(cross.pending, true);
assert.ok((cross.device.pendingUsers || []).some((p) => p.name === "Sam Floor"));
assert.strictEqual(cross.device.status, "approved");

// Boss cannot use Sam phone without approval
const bossOnSamPhone = devices.assertLoginAllowed({
  deviceId: "d_phonexyz12345678901",
  userName: "Office Boss",
  userAgent: "Mozilla/5.0 (iPhone)",
  ip: "127.0.0.1"
});
assert.strictEqual(bossOnSamPhone.ok, false);

let snap = devices.approveDevice("d_laptopabc1234567890", "Office Boss", {
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

snap = devices.revokeDevice("d_phone2abcdefghijklmn", "Office Boss");
assert.ok(snap.revoked.some((row) => row.id === "d_phone2abcdefghijklmn"));
assert.strictEqual(devices.assertLoginAllowed({
  deviceId: "d_phone2abcdefghijklmn",
  userName: "Sam Floor",
  userAgent: "Mozilla/5.0 (iPhone)",
  ip: "10.0.0.8"
}).ok, false);

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
    headers: { "content-type": "application/json", "x-sd-device-id": "d_laptopabc1234567890" },
    body: JSON.stringify({ name: "Office Boss", password: "admin", deviceId: "d_laptopabc1234567890" })
  });
  assert.strictEqual(login.json.ok, true, JSON.stringify(login.json));
  const token = login.json.token;
  const headers = { "content-type": "application/json", "x-sd-token": token };

  const pendingLogin = await api("/api/office/login", {
    method: "POST",
    headers: { "content-type": "application/json", "x-sd-device-id": "d_tablet111222333444" },
    body: JSON.stringify({ name: "Sam Floor", password: "sam", deviceId: "d_tablet111222333444" })
  });
  assert.strictEqual(pendingLogin.json.ok, false, "Sam already has a phone, tablet needs approval");
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
    body: JSON.stringify({ name: "Sam Floor", password: "sam", deviceId: "d_tablet111222333444" })
  });
  assert.strictEqual(second.json.ok, true, JSON.stringify(second.json));

  const bossOnTablet = await api("/api/office/login", {
    method: "POST",
    headers: { "content-type": "application/json", "x-sd-device-id": "d_tablet111222333444" },
    body: JSON.stringify({ name: "Office Boss", password: "admin", deviceId: "d_tablet111222333444" })
  });
  assert.strictEqual(bossOnTablet.json.ok, false);
  assert.strictEqual(bossOnTablet.json.pendingDevice, true);

  server.close();
  console.log("trusted-devices.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
