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

console.log("delivery-track.test.js ok");
