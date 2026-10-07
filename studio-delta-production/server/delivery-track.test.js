"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-delivery-track-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";

const { initWorkbook } = require("./workbook-store");
initWorkbook();
const track = require("./delivery-track");
const trusted = require("./trusted-devices");
const staff = require("./staff");

staff.upsertUser({
  name: "Lebo",
  access: "Admin",
  role: "Driver",
  password: "lebo1",
  seeDebtors: "No",
  jobTitle: "Driver",
  tasks: ["Delivery"]
});

let bad = null;
try {
  track.saveLocation("Lebo", { sharing: true });
} catch (e) {
  bad = e;
}
assert.ok(bad, "lat/lng required while sharing");

const live = track.saveLocation("Lebo", {
  lat: -25.74,
  lng: 28.21,
  accuracy: 18,
  sharing: true
});
assert.strictEqual(live.name, "Lebo");
assert.ok(live.live);
assert.strictEqual(live.lat, -25.74);
assert.ok(String(track.FACTORY.address || "").indexOf("Derdepoort") !== -1);

const listed = track.listLocations();
assert.strictEqual(listed.length, 1);
assert.strictEqual(listed[0].name, "Lebo");

const stopped = track.saveLocation("Lebo", { sharing: false });
assert.strictEqual(stopped.sharing, false);
assert.ok(stopped.stale);

const again = track.saveLocation("Lebo", {
  lat: -26.0674,
  lng: 28.1112,
  sharing: true
});
assert.ok(again.live);
assert.ok(Math.abs(again.lat + 26.0674) < 0.0001);

const phoneId = "d_driverphonepin123456789";
fs.writeFileSync(path.join(dir, "trusted-devices.json"), JSON.stringify({
  devices: [{
    id: phoneId,
    status: "approved",
    label: "Safari on iPhone",
    nickname: "Lebo van phone",
    assignedUsers: ["Lebo"],
    assignedTo: "Lebo",
    driverPhone: false,
    approvedAt: new Date().toISOString(),
    approvedBy: "Manager",
    pendingUsers: []
  }]
}));

trusted.setDriverPhone(phoneId, true);
const pinned = trusted.driverPhoneByDeviceId(phoneId);
assert.ok(pinned);
assert.strictEqual(pinned.driverPhone, true);

const fromPinned = track.saveLocation("Lebo", {
  lat: -25.75,
  lng: 28.22,
  sharing: true,
  deviceId: phoneId
}, { headers: { "x-sd-device-id": phoneId } });
assert.ok(fromPinned.pinned_phone);
assert.ok(String(fromPinned.device_label || "").indexOf("Lebo") !== -1 || fromPinned.device_label);
assert.ok(Math.abs(fromPinned.lat + 25.75) < 0.0001);

const otherPhone = track.saveLocation("Lebo", {
  lat: -26.1,
  lng: 28.2,
  sharing: true,
  deviceId: "d_otherphoneabcdef123456"
}, { headers: { "x-sd-device-id": "d_otherphoneabcdef123456" } });
assert.ok(Math.abs(otherPhone.lat + 25.75) < 0.0001, "non-pinned phone must not move the driver pin");
assert.ok(otherPhone.pinned_phone);

const status = track.trackerStatus(
  { name: "Lebo", jobTitle: "Driver", tasks: ["Delivery"] },
  { headers: { "x-sd-device-id": phoneId } }
);
assert.strictEqual(status.driverPhone, true);
assert.strictEqual(status.alwaysShare, true);
assert.strictEqual(status.shouldShare, true);

const statusOther = track.trackerStatus(
  { name: "Lebo", jobTitle: "Driver", tasks: ["Delivery"] },
  { headers: { "x-sd-device-id": "d_otherphoneabcdef123456" } }
);
assert.strictEqual(statusOther.driverPhone, false);
assert.strictEqual(statusOther.pinnedForMe, true);
assert.strictEqual(statusOther.shouldShare, false);
assert.strictEqual(statusOther.alwaysShare, false);

console.log("delivery-track.test.js ok");
