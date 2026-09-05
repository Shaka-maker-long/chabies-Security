const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-persist-"));
const volume = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-vol-"));
process.env.DATA_DIR = dir;
process.env.TZ = "Africa/Johannesburg";
delete process.env.RAILWAY_ENVIRONMENT;
delete process.env.RAILWAY_PROJECT_ID;
delete process.env.RAILWAY_SERVICE_ID;
delete process.env.RAILWAY_VOLUME_MOUNT_PATH;
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;

const store = require("./workbook-store");

assert.strictEqual(store.dataDir(), dir);
assert.strictEqual(store.storageInfo().usingEphemeralDisk, false);
assert.ok(/SQLite|JSON files/.test(store.storageInfo().database));

process.env.RAILWAY_ENVIRONMENT = "production";
assert.strictEqual(store.storageInfo().onRailway, true);
assert.strictEqual(store.storageInfo().usingEphemeralDisk, true);
assert.ok(store.storageInfo().warning && /volume/i.test(store.storageInfo().warning));
assert.strictEqual(store.dataDir(), dir, "without a volume, DATA_DIR is still used");

process.env.RAILWAY_VOLUME_MOUNT_PATH = volume;
assert.strictEqual(store.dataDir(), volume, "Railway volume wins over DATA_DIR");
assert.strictEqual(store.storageInfo().usingEphemeralDisk, false);
assert.strictEqual(store.storageInfo().warning, null);
assert.strictEqual(store.storageInfo().volumeMount, volume);

const healthJs = fs.readFileSync(path.join(__dirname, "index.js"), "utf8");
assert.ok(healthJs.indexOf("usingEphemeralDisk") !== -1);
assert.ok(healthJs.indexOf("persistenceInfo") !== -1);

console.log("persistence.test.js ok");
