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
fs.writeFileSync(path.join(dir, "email-replies.json"), JSON.stringify({ subject: "Thanks" }));
fs.writeFileSync(path.join(dir, "showroom-bookings.json"), JSON.stringify({ bookings: [{ who: "Nomsa" }] }));
fs.writeFileSync(path.join(dir, "floor-layout.json"), JSON.stringify({ stations: ["weld"] }));
fs.writeFileSync(path.join(dir, "standard-cutting-lists.json"), JSON.stringify({ lists: ["std"] }));
fs.mkdirSync(path.join(dir, "job-card-images"), { recursive: true });
fs.writeFileSync(path.join(dir, "job-card-images", "chair.png"), "fake-png");
fs.mkdirSync(path.join(dir, "pdf-images"), { recursive: true });
fs.writeFileSync(path.join(dir, "pdf-images", "glass.png"), "fake-glass");
fs.writeFileSync(path.join(dir, "future-module.json"), JSON.stringify({ keep: "everything" }));

const status = backup.runBackup("test");
assert.ok(status.ok, status.error || "backup should succeed locally");
assert.ok(status.localDb);
assert.ok(status.complete, "complete .tgz must be written");
assert.ok(status.verified, "snapshot must pass integrity_check");
assert.strictEqual(status.integrity, "ok");
assert.strictEqual(status.offsite, false);
assert.ok(fs.existsSync(path.join(backup.backupsDir(), status.localDb)));
assert.ok(fs.existsSync(path.join(backup.backupsDir(), status.complete)));
assert.ok(status.packedAll, "backup must pack every live data file");
assert.ok(Array.isArray(status.packedFiles) && status.packedFiles.indexOf("email-replies.json") !== -1);
assert.ok(status.packedFiles.indexOf("job-card-images") !== -1);
assert.ok(status.packedFiles.indexOf("future-module.json") !== -1);

const sqlite = require("./sqlite-store");
const logBefore = sqlite.open().prepare("SELECT COUNT(*) AS n FROM sheet_rows WHERE title = 'Production_Log'").get();
assert.ok(logBefore && logBefore.n >= 1, "production logs must be in SQLite before backup");
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
assert.strictEqual(info.backupPackedAll, true);
assert.ok(info.backupLocalCount >= 1);

const complete = path.join(backup.backupsDir(), status.complete);
const inspected = backup.inspectSource(complete);
assert.strictEqual(inspected.kind, "complete");
assert.strictEqual(inspected.verified.integrity, "ok");

assert.throws(() => backup.restoreFromPath(complete, { confirm: "yes" }), /RESTORE/);

db.deleteAllEnquiries();
fs.rmSync(path.join(dir, "debtor-payments"), { recursive: true, force: true });
fs.unlinkSync(path.join(dir, "email-replies.json"));
fs.unlinkSync(path.join(dir, "showroom-bookings.json"));
fs.unlinkSync(path.join(dir, "floor-layout.json"));
fs.unlinkSync(path.join(dir, "standard-cutting-lists.json"));
fs.unlinkSync(path.join(dir, "future-module.json"));
fs.rmSync(path.join(dir, "job-card-images"), { recursive: true, force: true });
fs.rmSync(path.join(dir, "pdf-images"), { recursive: true, force: true });
assert.strictEqual(db.listEnquiries().length, 0, "live shop must be empty before restore");

const restored = backup.restoreFromPath(complete, { confirm: "RESTORE" });
assert.ok(restored.ok, restored.error || "restore should succeed");
assert.ok(restored.safety, "a safety snapshot must be taken before restore");
const after = db.listEnquiries();
assert.ok(after.some((row) => String(row.client_name) === "Backup Client"), JSON.stringify(after));
assert.ok(fs.existsSync(path.join(dir, "debtor-payments", "S-BACKUP", "proof.txt")), "proof files must come back");
assert.strictEqual(fs.readFileSync(path.join(dir, "debtor-payments", "S-BACKUP", "proof.txt"), "utf8"), "proof-of-payment");
assert.strictEqual(JSON.parse(fs.readFileSync(path.join(dir, "email-replies.json"), "utf8")).subject, "Thanks");
assert.strictEqual(JSON.parse(fs.readFileSync(path.join(dir, "showroom-bookings.json"), "utf8")).bookings[0].who, "Nomsa");
assert.strictEqual(JSON.parse(fs.readFileSync(path.join(dir, "floor-layout.json"), "utf8")).stations[0], "weld");
assert.ok(fs.existsSync(path.join(dir, "job-card-images", "chair.png")));
assert.ok(fs.existsSync(path.join(dir, "pdf-images", "glass.png")));
assert.strictEqual(JSON.parse(fs.readFileSync(path.join(dir, "future-module.json"), "utf8")).keep, "everything");
const logAfter = sqlite.open().prepare("SELECT COUNT(*) AS n FROM sheet_rows WHERE title = 'Production_Log'").get();
assert.ok(logAfter && logAfter.n >= 1, "production logs must come back from the .tgz");

const junk = path.join(backup.backupsDir(), "studio-delta-20990101-000000.db");
fs.writeFileSync(junk, "not a sqlite database");
const keepCount = db.listEnquiries().length;
assert.throws(() => backup.restoreFromPath(junk, { confirm: "RESTORE" }));
assert.strictEqual(db.listEnquiries().length, keepCount, "corrupt restore must not wipe the live shop");

const safe = backup.safeBackupName(status.localDb);
assert.ok(safe && fs.existsSync(safe));
assert.strictEqual(backup.safeBackupName("../studio-delta.db"), null);
assert.strictEqual(backup.safeBackupName("last.json"), null);

const staff = require("./staff");
const actor = { name: "Backup User" };
assert.throws(() => staff.checkRestoreSecrets(actor, { confirm: "RESTORE", password: "x" }), /twice/);
assert.throws(() => staff.checkRestoreSecrets(actor, { confirm: "RESTORE", password: "x", confirmPassword: "nope" }), /match/);
assert.throws(() => staff.checkRestoreSecrets(actor, { confirm: "RESTORE", password: "wrong", confirmPassword: "wrong" }), /wrong/i);
assert.throws(() => staff.checkRestoreSecrets(actor, { confirm: "yes", password: "x", confirmPassword: "x" }), /RESTORE/);
staff.checkRestoreSecrets(actor, { confirm: "RESTORE", password: "x", confirmPassword: "x" });

assert.strictEqual(backup.normalizeFolderId("https://drive.google.com/drive/folders/AbC123_xYz?usp=sharing"), "AbC123_xYz");
assert.strictEqual(backup.normalizeFolderId("AbC123_xYz"), "AbC123_xYz");

const { parseGoogleCredentials } = require("./workbook-store");
process.env.GOOGLE_SERVICE_ACCOUNT_JSON = JSON.stringify(JSON.stringify({
  type: "service_account",
  client_email: "bot@x.iam.gserviceaccount.com"
}));
assert.strictEqual(parseGoogleCredentials().email, "bot@x.iam.gserviceaccount.com");

process.env.GOOGLE_SERVICE_ACCOUNT_JSON = JSON.stringify({
  type: "service_account",
  client_email: "bot@x.iam.gserviceaccount.com",
  private_key: "x"
});
delete process.env.BACKUP_DRIVE_FOLDER_ID;
const missingFolder = backup.runBackup("drive-folder-missing");
assert.ok(missingFolder.ok, missingFolder.error || "local backup still runs without a Drive folder");
assert.strictEqual(missingFolder.offsite, false);
assert.ok(/BACKUP_DRIVE_FOLDER_ID/.test(missingFolder.offsiteError || ""), missingFolder.offsiteError);
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;

console.log("backup.test.js ok");
