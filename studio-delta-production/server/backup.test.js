const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-bak-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;
delete process.env.GMAIL_SENDER;
delete process.env.BACKUP_DRIVE_FOLDER_ID;
delete process.env.BACKUP_EMAIL;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const backup = require("./backup");

initWorkbook();
const book = getBook();
book.getSheetByName("Users").appendRow(["Backup User", "Admin", "x", "", "Admin", "Yes"]);
persistWorkbook();
db.upsertEnquiry({
  date_enquired: "02/09/2026",
  client_name: "Backup Client",
  product: "Air Chair",
  status: "New"
});

const status = backup.runBackup("test");
assert.ok(status.ok, status.error || "backup should succeed locally");
assert.ok(status.localDb);
assert.strictEqual(status.offsite, false);
assert.ok(fs.existsSync(path.join(backup.backupsDir(), status.localDb)));
assert.ok(status.bytes > 0);
assert.ok(status.sha256 && status.sha256.length === 64);

const listed = backup.listLocalSnapshots();
assert.ok(listed.length >= 1);
assert.ok(!backup.isDue(backup.loadStatus()), "a successful backup today is not due again");

const info = backup.info();
assert.strictEqual(info.backupOk, true);
assert.strictEqual(info.backupOffsite, false);
assert.ok(info.backupLocalCount >= 1);

const safe = backup.safeBackupName(status.localDb);
assert.ok(safe && fs.existsSync(safe));
assert.strictEqual(backup.safeBackupName("../studio-delta.db"), null);
assert.strictEqual(backup.safeBackupName("last.json"), null);

console.log("backup.test.js ok");
