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
const proofDir = path.join(dir, "debtor-payments", "S-BACKUP");
fs.mkdirSync(proofDir, { recursive: true });
fs.writeFileSync(path.join(proofDir, "proof.txt"), "proof-of-payment");

const status = backup.runBackup("test");
assert.ok(status.ok, status.error || "backup should succeed locally");
assert.ok(status.localDb);
assert.ok(status.complete, "complete .tgz must be written");
assert.ok(status.verified, "snapshot must pass integrity_check");
assert.strictEqual(status.integrity, "ok");
assert.strictEqual(status.offsite, false);
assert.ok(fs.existsSync(path.join(backup.backupsDir(), status.localDb)));
assert.ok(fs.existsSync(path.join(backup.backupsDir(), status.complete)));
assert.ok(status.bytes > 0);
assert.ok(status.sha256 && status.sha256.length === 64);
assert.ok(status.counts && status.counts.enquiries >= 1);

const listed = backup.listLocalSnapshots();
assert.ok(listed.length >= 1);
assert.ok(listed[0].complete || listed[0].db);
assert.ok(!backup.isDue(backup.loadStatus()), "a successful backup today is not due again");

const info = backup.info();
assert.strictEqual(info.backupOk, true);
assert.strictEqual(info.backupVerified, true);
assert.strictEqual(info.backupOffsite, false);
assert.ok(info.backupLocalCount >= 1);

const complete = path.join(backup.backupsDir(), status.complete);
const inspected = backup.inspectSource(complete);
assert.strictEqual(inspected.kind, "complete");
assert.strictEqual(inspected.verified.integrity, "ok");

assert.throws(() => backup.restoreFromPath(complete, { confirm: "yes" }), /RESTORE/);

db.deleteAllEnquiries();
fs.rmSync(path.join(dir, "debtor-payments"), { recursive: true, force: true });
assert.strictEqual(db.listEnquiries().length, 0, "live shop must be empty before restore");

const restored = backup.restoreFromPath(complete, { confirm: "RESTORE" });
assert.ok(restored.ok, restored.error || "restore should succeed");
assert.ok(restored.safety, "a safety snapshot must be taken before restore");
const after = db.listEnquiries();
assert.ok(after.some((row) => String(row.client_name) === "Backup Client"), JSON.stringify(after));
assert.ok(fs.existsSync(path.join(dir, "debtor-payments", "S-BACKUP", "proof.txt")), "proof files must come back");
assert.strictEqual(fs.readFileSync(path.join(dir, "debtor-payments", "S-BACKUP", "proof.txt"), "utf8"), "proof-of-payment");

const junk = path.join(backup.backupsDir(), "studio-delta-20990101-000000.db");
fs.writeFileSync(junk, "not a sqlite database");
const keepCount = db.listEnquiries().length;
assert.throws(() => backup.restoreFromPath(junk, { confirm: "RESTORE" }));
assert.strictEqual(db.listEnquiries().length, keepCount, "corrupt restore must not wipe the live shop");

const safe = backup.safeBackupName(status.localDb);
assert.ok(safe && fs.existsSync(safe));
assert.strictEqual(backup.safeBackupName("../studio-delta.db"), null);
assert.strictEqual(backup.safeBackupName("last.json"), null);

console.log("backup.test.js ok");
