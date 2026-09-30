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

assert.strictEqual(devices.approvedCount(), 0);
const first = devices.assertLoginAllowed({
  deviceId: "d_laptopabc1234567890",
  userName: "Office Boss",
  userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
  ip: "127.0.0.1"
});
assert.strictEqual(first.ok, true);
assert.strictEqual(first.bootstrapped, true);
assert.strictEqual(first.device.status, "approved");
assert.ok(/Chrome on Windows/.test(first.device.label));
assert.strictEqual(devices.approvedCount(), 1);

const blocked = devices.assertLoginAllowed({
  deviceId: "d_phonexyz12345678901",
  userName: "Sam Floor",
  userAgent: "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 Safari/604.1",
  ip: "10.0.0.2"
});
assert.strictEqual(blocked.ok, false);
assert.strictEqual(blocked.pending, true);
assert.strictEqual(blocked.device.status, "pending");
assert.ok(/Safari on iPhone/.test(blocked.device.label));

let snap = devices.snapshot();
assert.strictEqual(snap.pendingCount, 1);
assert.strictEqual(snap.approvedCount, 1);

snap = devices.approveDevice("d_phonexyz12345678901", "Office Boss");
assert.strictEqual(snap.pendingCount, 0);
assert.strictEqual(snap.approvedCount, 2);

const allowed = devices.assertLoginAllowed({
  deviceId: "d_phonexyz12345678901",
  userName: "Sam Floor",
  userAgent: "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15",
  ip: "10.0.0.2"
});
assert.strictEqual(allowed.ok, true);
assert.strictEqual(allowed.device.status, "approved");

snap = devices.revokeDevice("d_phonexyz12345678901", "Office Boss");
assert.strictEqual(snap.approvedCount, 1);
assert.ok(snap.revoked.some((row) => row.id === "d_phonexyz12345678901"));

const again = devices.assertLoginAllowed({
  deviceId: "d_phonexyz12345678901",
  userName: "Sam Floor",
  userAgent: "Mozilla/5.0 (iPhone)",
  ip: "10.0.0.2"
});
assert.strictEqual(again.ok, false);
assert.strictEqual(again.pending, true);

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
  assert.strictEqual(pendingLogin.json.ok, false);
  assert.strictEqual(pendingLogin.json.pendingDevice, true);

  const listed = await api("/api/office/devices", { headers });
  assert.strictEqual(listed.json.ok, true);
  assert.ok((listed.json.pending || []).some((row) => row.id === "d_tablet111222333444"));

  const approved = await api("/api/office/devices/d_tablet111222333444/approve", {
    method: "POST",
    headers,
    body: "{}"
  });
  assert.strictEqual(approved.json.ok, true, JSON.stringify(approved.json));
  assert.ok((approved.json.approved || []).some((row) => row.id === "d_tablet111222333444"));

  const second = await api("/api/office/login", {
    method: "POST",
    headers: { "content-type": "application/json", "x-sd-device-id": "d_tablet111222333444" },
    body: JSON.stringify({ name: "Sam Floor", password: "sam", deviceId: "d_tablet111222333444" })
  });
  assert.strictEqual(second.json.ok, true, JSON.stringify(second.json));

  server.close();
  console.log("trusted-devices.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
