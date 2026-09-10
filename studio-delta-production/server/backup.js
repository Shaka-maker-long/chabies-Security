const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { spawnSync } = require("child_process");
const { dataDir, persistWorkbook, hasGoogleAuth, parseGoogleCredentials, googleServiceAccountEmail } = require("./workbook-store");
const sqlite = require("./sqlite-store");

const KEEP_LOCAL = 14;
const KEEP_DRIVE = 14;
const SAST_OFFSET_MS = 2 * 60 * 60 * 1000;
const MAIL_ATTACH_MAX = 12 * 1024 * 1024;
const SAFE_NAME = /^studio-delta-[A-Za-z0-9._-]+$/;
const CONFIRM_WORD = "RESTORE";
const SKIP_LIVE = new Set([
  "backups",
  "studio-delta.db",
  "studio-delta.db-wal",
  "studio-delta.db-shm",
  "studio-delta.db-journal",
  "manifest.json"
]);
const FILE_DIRS = [
  "enquiry-quotes",
  "enquiry-files",
  "debtor-payments",
  "paint-shop-invoices",
  "glass-po-invoices",
  "job-cards",
  "job-card-images",
  "pdf-images"
];

let running = false;
let restoring = false;

function sastDate(d) {
  const sast = new Date((d || new Date()).getTime() + SAST_OFFSET_MS);
  const p = (n) => String(n).padStart(2, "0");
  return {
    y: sast.getUTCFullYear(),
    m: sast.getUTCMonth() + 1,
    day: sast.getUTCDate(),
    h: sast.getUTCHours(),
    min: sast.getUTCMinutes(),
    date: sast.getUTCFullYear() + "-" + p(sast.getUTCMonth() + 1) + "-" + p(sast.getUTCDate()),
    stamp: sast.getUTCFullYear() + p(sast.getUTCMonth() + 1) + p(sast.getUTCDate()) + "-" + p(sast.getUTCHours()) + p(sast.getUTCMinutes()) + p(sast.getUTCSeconds())
  };
}

function backupsDir() {
  const dir = path.join(dataDir(), "backups");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function statusPath() {
  return path.join(backupsDir(), "last.json");
}

function restoreLogPath() {
  return path.join(backupsDir(), "last-restore.json");
}

function readJson(file, fallback) {
  try { return JSON.parse(fs.readFileSync(file, "utf8")); } catch (e) { return fallback; }
}

function writeJson(file, value) {
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(value, null, 2));
  fs.renameSync(tmp, file);
}

function loadStatus() {
  return readJson(statusPath(), null);
}

function loadRestoreStatus() {
  return readJson(restoreLogPath(), null);
}

function sha256File(file) {
  const hash = crypto.createHash("sha256");
  hash.update(fs.readFileSync(file));
  return hash.digest("hex");
}

function fileSize(file) {
  try { return fs.statSync(file).size; } catch (e) { return 0; }
}

function copyDir(src, dest) {
  fs.mkdirSync(dest, { recursive: true });
  fs.readdirSync(src, { withFileTypes: true }).forEach((ent) => {
    const from = path.join(src, ent.name);
    const to = path.join(dest, ent.name);
    if (ent.isDirectory()) copyDir(from, to);
    else if (ent.isFile()) fs.copyFileSync(from, to);
  });
}

function replaceDir(src, dest) {
  if (fs.existsSync(dest)) fs.rmSync(dest, { recursive: true, force: true });
  if (fs.existsSync(src)) copyDir(src, dest);
}

function replaceFile(src, dest) {
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  const tmp = dest + ".new";
  fs.copyFileSync(src, tmp);
  fs.renameSync(tmp, dest);
}

function rmQuiet(file) {
  try { fs.unlinkSync(file); } catch (e) {}
}

function skipLiveName(name) {
  if (!name || name.charAt(0) === ".") return true;
  if (SKIP_LIVE.has(name)) return true;
  if (name.indexOf("backup-staging-") === 0) return true;
  if (name.indexOf("restore-inspect-") === 0) return true;
  if (name.indexOf("restore-apply-") === 0) return true;
  return false;
}

function listLiveDataEntries() {
  const dir = dataDir();
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir, { withFileTypes: true }).filter((ent) => !skipLiveName(ent.name));
}

function packLiveData(stageDir) {
  listLiveDataEntries().forEach((ent) => {
    const src = path.join(dataDir(), ent.name);
    const dest = path.join(stageDir, ent.name);
    if (ent.isDirectory()) copyDir(src, dest);
    else if (ent.isFile()) fs.copyFileSync(src, dest);
  });
}

function uniqueStamp() {
  const base = sastDate().stamp;
  let stamp = base;
  let n = 1;
  while (
    fs.existsSync(path.join(backupsDir(), "studio-delta-" + stamp + ".tgz")) ||
    fs.existsSync(path.join(backupsDir(), "studio-delta-" + stamp + ".db"))
  ) {
    n += 1;
    stamp = base + "-" + n;
  }
  return stamp;
}

function stampFromName(name) {
  const m = String(name || "").match(/^studio-delta-(?:uploaded-)?(\d{8}-\d{4,6}(?:-\d+)?)/);
  return m ? m[1] : "";
}

function snapshotFiles(stamp) {
  const dir = backupsDir();
  const base = "studio-delta-" + stamp;
  return {
    stamp,
    complete: path.join(dir, base + ".tgz"),
    db: path.join(dir, base + ".db"),
    json: path.join(dir, base + ".json"),
    files: path.join(dir, base + "-files.tgz"),
    manifest: path.join(dir, base + "-manifest.json")
  };
}

function verifySqliteFile(file) {
  if (!file || !fs.existsSync(file) || fileSize(file) < 100) {
    throw new Error("SQLite snapshot is missing or empty");
  }
  const { DatabaseSync } = require("node:sqlite");
  let db;
  try {
    db = new DatabaseSync(file, { readOnly: true });
  } catch (e) {
    db = new DatabaseSync(file);
  }
  try {
    const row = db.prepare("PRAGMA integrity_check").get();
    const msg = String((row && (row.integrity_check != null ? row.integrity_check : Object.values(row)[0])) || "");
    if (msg.toLowerCase() !== "ok") {
      throw new Error("SQLite snapshot failed integrity_check: " + msg);
    }
    const one = (sql) => {
      try {
        const r = db.prepare(sql).get();
        return Number(r && (r.n != null ? r.n : Object.values(r)[0])) || 0;
      } catch (e) {
        return 0;
      }
    };
    return {
      ok: true,
      integrity: "ok",
      counts: {
        users: one("SELECT COUNT(*) AS n FROM users"),
        orders: one("SELECT COUNT(*) AS n FROM orders"),
        enquiries: one("SELECT COUNT(*) AS n FROM enquiries"),
        payments: one("SELECT COUNT(*) AS n FROM payments")
      }
    };
  } finally {
    try { db.close(); } catch (e) {}
  }
}

function listLocalSnapshots() {
  const dir = backupsDir();
  const stamps = new Set();
  fs.readdirSync(dir).forEach((name) => {
    const stamp = stampFromName(name);
    if (stamp) stamps.add(stamp);
  });
  return Array.from(stamps).map((stamp) => {
    const files = snapshotFiles(stamp);
    const complete = fs.existsSync(files.complete);
    const db = fs.existsSync(files.db);
    const chosen = complete ? files.complete : (db ? files.db : files.json);
    const st = chosen && fs.existsSync(chosen) ? fs.statSync(chosen) : null;
    const manifest = readJson(files.manifest, null);
    return {
      stamp,
      name: path.basename(chosen || files.db),
      complete: complete ? path.basename(files.complete) : null,
      db: db ? path.basename(files.db) : null,
      json: fs.existsSync(files.json) ? path.basename(files.json) : null,
      files: fs.existsSync(files.files) ? path.basename(files.files) : null,
      bytes: st ? st.size : 0,
      at: st ? st.mtime.toISOString() : null,
      verified: !!(manifest && manifest.ok && manifest.integrity === "ok"),
      integrity: manifest && manifest.integrity ? manifest.integrity : (db ? "unchecked" : null),
      counts: manifest && manifest.counts ? manifest.counts : null,
      reason: manifest && manifest.reason ? manifest.reason : null
    };
  }).filter((row) => row.bytes > 0).sort((a, b) => String(b.at || "").localeCompare(String(a.at || "")));
}

function pruneLocal() {
  const keep = listLocalSnapshots().slice(KEEP_LOCAL);
  keep.forEach((row) => {
    const files = snapshotFiles(row.stamp);
    [files.complete, files.db, files.json, files.files, files.manifest, files.db + "-wal", files.db + "-shm"].forEach(rmQuiet);
  });
}

function snapshotSqlite(dest) {
  try { persistWorkbook(); } catch (e) {}
  try { require("./db").persist(); } catch (e) {}
  sqlite.checkpoint();
  if (fs.existsSync(dest)) fs.unlinkSync(dest);
  const db = sqlite.open();
  if (!db) throw new Error("SQLite is not available");
  const live = sqlite.sqlitePath();
  try {
    const escaped = dest.replace(/'/g, "''");
    db.exec("VACUUM INTO '" + escaped + "'");
  } catch (e) {
    fs.copyFileSync(live, dest);
    try { fs.copyFileSync(live + "-wal", dest + "-wal"); } catch (err) {}
    try { fs.copyFileSync(live + "-shm", dest + "-shm"); } catch (err) {}
  }
  if (!fs.existsSync(dest)) throw new Error("SQLite snapshot was not created");
  return verifySqliteFile(dest);
}

function copyOfficeJson(dest) {
  try {
    writeJson(dest, require("./db").railwayBackup());
  } catch (e) {
    const office = path.join(dataDir(), "studio-delta.json");
    if (fs.existsSync(office)) fs.copyFileSync(office, dest);
  }
}

function tarCreate(dest, cwd, parts) {
  const r = spawnSync("tar", ["-czf", dest, "-C", cwd].concat(parts), { encoding: "utf8" });
  if (r.status !== 0) throw new Error(r.stderr || "Could not write the backup archive");
  return fileSize(dest);
}

function tarExtract(archive, dest) {
  fs.mkdirSync(dest, { recursive: true });
  const r = spawnSync("tar", ["-xzf", archive, "-C", dest], { encoding: "utf8" });
  if (r.status !== 0) throw new Error(r.stderr || "Could not read the backup archive");
}

function buildCompleteBundle(stamp, localDb, localJson, verified) {
  const staging = path.join(dataDir(), "backup-staging-" + stamp);
  try {
    if (fs.existsSync(staging)) fs.rmSync(staging, { recursive: true, force: true });
    fs.mkdirSync(staging, { recursive: true });
    fs.copyFileSync(localDb, path.join(staging, "studio-delta.db"));
    packLiveData(staging);
    if (fs.existsSync(localJson)) fs.copyFileSync(localJson, path.join(staging, "studio-delta.json"));
    FILE_DIRS.forEach((name) => {
      const dest = path.join(staging, name);
      if (!fs.existsSync(dest)) fs.mkdirSync(dest, { recursive: true });
    });
    const parts = fs.readdirSync(staging);
    const manifest = {
      version: 3,
      kind: "studio-delta-complete",
      at: new Date().toISOString(),
      stamp,
      ok: true,
      packedAll: true,
      integrity: verified.integrity,
      counts: verified.counts,
      sha256Db: sha256File(localDb),
      files: parts.slice()
    };
    writeJson(path.join(staging, "manifest.json"), manifest);
    const complete = snapshotFiles(stamp).complete;
    tarCreate(complete, staging, fs.readdirSync(staging));
    writeJson(snapshotFiles(stamp).manifest, Object.assign({}, manifest, {
      bytes: fileSize(complete),
      complete: path.basename(complete)
    }));
    const filesTar = snapshotFiles(stamp).files;
    const fileParts = fs.readdirSync(staging, { withFileTypes: true })
      .filter((ent) => ent.isDirectory())
      .map((ent) => ent.name);
    if (fileParts.length) tarCreate(filesTar, staging, fileParts);
    return {
      complete: path.basename(complete),
      filesArchive: fs.existsSync(filesTar) ? path.basename(filesTar) : null,
      manifest
    };
  } finally {
    try { fs.rmSync(staging, { recursive: true, force: true }); } catch (e) {}
  }
}

function impersonateEmail() {
  return String(process.env.BACKUP_DRIVE_IMPERSONATE || process.env.GMAIL_SENDER || "").trim();
}

function explainDriveError(raw, email) {
  const msg = String(raw || "");
  const bot = email || googleServiceAccountEmail() || "the service account";
  if (/storage quota|do not have storage quota/i.test(msg)) {
    return "Google will not let " + bot +
      " own files in an ordinary My Drive folder (service accounts have no storage). Put the backup folder in a Shared drive and add that email as Content manager. Or set BACKUP_DRIVE_IMPERSONATE to a Workspace mailbox after an admin enables domain-wide delegation.";
  }
  return msg;
}

function driveRpc(payload) {
  const body = Object.assign({}, payload || {});
  const who = impersonateEmail();
  if (who && !body.impersonate) body.impersonate = who;
  const r = spawnSync(process.execPath, [path.join(__dirname, "drive-cli.js")], {
    input: JSON.stringify(body),
    encoding: "utf8",
    env: process.env,
    maxBuffer: 8 * 1024 * 1024
  });
  if (!r.stdout) throw new Error(explainDriveError(r.stderr || "Drive helper failed"));
  const parsed = JSON.parse(r.stdout);
  if (!parsed.ok) throw new Error(explainDriveError(parsed.error || "Drive helper error"));
  return parsed;
}

function shareEmail() {
  return String(process.env.BACKUP_SHARE_EMAIL || process.env.BACKUP_EMAIL || process.env.GMAIL_SENDER || "").trim();
}

function mailTo() {
  return String(process.env.BACKUP_EMAIL || process.env.GMAIL_SENDER || "").trim();
}

function normalizeFolderId(raw) {
  let s = String(raw || "").trim();
  if ((s.charAt(0) === '"' && s.slice(-1) === '"') || (s.charAt(0) === "'" && s.slice(-1) === "'")) {
    s = s.slice(1, -1).trim();
  }
  const folder = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (folder) return folder[1];
  const byId = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (byId) return byId[1];
  return s;
}

function configuredFolderId() {
  return normalizeFolderId(process.env.BACKUP_DRIVE_FOLDER_ID);
}

function probeDriveFolder(folderId, email) {
  let meta;
  try {
    meta = driveRpc({ op: "getFile", fileId: folderId });
  } catch (e) {
    throw new Error(
      "The service account " + (email || "(unknown)") +
      " cannot open that Drive folder. Open the folder → Share → add that email as Editor (not Viewer). Enable the Google Drive API on the Google Cloud project. " +
      (e.message || String(e))
    );
  }
  if (meta.mimeType && meta.mimeType !== "application/vnd.google-apps.folder") {
    throw new Error("BACKUP_DRIVE_FOLDER_ID must be a folder ID, not a file.");
  }
  if (!impersonateEmail() && !meta.driveId) {
    throw new Error(
      "This folder is in ordinary My Drive. Google will not let " + (email || "the service account") +
      " store files there (no storage quota). In Drive: New → Shared drive, add that email as Content manager, put the backup folder inside the Shared drive, then set BACKUP_DRIVE_FOLDER_ID to that folder. Or set BACKUP_DRIVE_IMPERSONATE to a Workspace mailbox."
    );
  }
  return meta;
}

function shareUploaded(fileId) {
  const email = shareEmail();
  if (!email || !fileId) return;
  try { driveRpc({ op: "shareWithEmail", fileId, email, role: "writer" }); } catch (e) {
    console.warn("[backup] could not share uploaded Drive file", e.message || e);
  }
}

function pruneDrive(folderId) {
  const listed = driveRpc({ op: "listFiles", folderId });
  const archives = (listed.files || [])
    .filter((f) => {
      const n = String(f.name || "");
      return /^studio-delta-/.test(n) && n.slice(-4) === ".tgz" && n.indexOf("-files.tgz") === -1;
    })
    .sort((a, b) => String(b.createdTime || "").localeCompare(String(a.createdTime || "")));
  archives.slice(KEEP_DRIVE).forEach((f) => {
    try { driveRpc({ op: "trashFile", fileId: f.id }); } catch (e) {}
  });
}

function driveReady() {
  if (!hasGoogleAuth()) {
    return { ok: false, offsite: false, reason: "Google Drive is not configured on Railway" };
  }
  const creds = parseGoogleCredentials();
  if (!creds.ok) {
    return { ok: false, offsite: false, offsiteError: creds.error, serviceAccount: null };
  }
  const folderId = configuredFolderId();
  if (!folderId) {
    return {
      ok: false,
      offsite: false,
      offsiteError: "Set BACKUP_DRIVE_FOLDER_ID on Railway to the folder ID from the Drive URL, then share that folder with " + creds.email + " as Editor.",
      serviceAccount: creds.email
    };
  }
  return { ok: true, email: creds.email, folderId, impersonate: impersonateEmail() || null };
}

function uploadOffsite(localDb, localJson, archivePath, completePath) {
  const ready = driveReady();
  if (!ready.ok) return ready;
  if (!completePath || !fs.existsSync(completePath)) {
    throw new Error("Complete restore archive was not built, so nothing was sent to Drive.");
  }
  probeDriveFolder(ready.folderId, ready.email);
  const zipUp = driveRpc({
    op: "uploadFile",
    path: completePath,
    name: path.basename(completePath),
    folderId: ready.folderId,
    mimeType: "application/gzip"
  });
  shareUploaded(zipUp.id);
  let driveId = zipUp.id;
  let driveUrl = zipUp.url || null;
  if (localDb && fs.existsSync(localDb)) {
    try {
      const dbUp = driveRpc({
        op: "uploadFile",
        path: localDb,
        name: path.basename(localDb),
        folderId: ready.folderId,
        mimeType: "application/vnd.sqlite3"
      });
      shareUploaded(dbUp.id);
    } catch (e) {
      console.warn("[backup] Drive SQLite upload failed", e.message || e);
    }
  }
  try { pruneDrive(ready.folderId); } catch (e) {
    console.warn("[backup] Drive prune failed", e.message || e);
  }
  return {
    offsite: true,
    driveId,
    driveUrl,
    filesUrl: driveUrl,
    folderId: ready.folderId,
    serviceAccount: ready.email,
    impersonate: ready.impersonate || null,
    completeDrive: path.basename(completePath)
  };
}

function testDriveUpload() {
  const ready = driveReady();
  if (!ready.ok) {
    const err = new Error(ready.offsiteError || ready.reason || "Google Drive is not configured");
    err.detail = ready;
    throw err;
  }
  probeDriveFolder(ready.folderId, ready.email);
  const tmp = path.join(backupsDir(), "studio-delta-drive-check.txt");
  fs.writeFileSync(tmp, "Studio Delta Drive check " + new Date().toISOString() + "\nShare this folder with " + ready.email + " as Editor.\n");
  try {
    const up = driveRpc({
      op: "uploadFile",
      path: tmp,
      name: "studio-delta-drive-check.txt",
      folderId: ready.folderId,
      mimeType: "text/plain"
    });
    shareUploaded(up.id);
    return {
      ok: true,
      offsite: true,
      name: "studio-delta-drive-check.txt",
      driveId: up.id || null,
      driveUrl: up.url || null,
      folderId: ready.folderId,
      serviceAccount: ready.email
    };
  } finally {
    rmQuiet(tmp);
  }
}

function retryOffsite() {
  const last = loadStatus();
  if (!last || !last.ok || last.offsite) return last;
  const ready = driveReady();
  if (!ready.ok) return last;
  const complete = last.complete ? path.join(backupsDir(), last.complete) : latestCompletePath();
  const dbFile = last.localDb ? path.join(backupsDir(), last.localDb) : null;
  if (!complete || !fs.existsSync(complete)) return last;
  try {
    const off = uploadOffsite(dbFile, null, null, complete);
    const next = Object.assign({}, last, off, { offsiteError: null, reason: last.reason || "retry-offsite" });
    writeJson(statusPath(), next);
    console.log("[backup] off-site retry ok", next.completeDrive || next.complete);
    return next;
  } catch (e) {
    const next = Object.assign({}, last, { offsite: false, offsiteError: e.message || String(e), serviceAccount: ready.email });
    writeJson(statusPath(), next);
    console.warn("[backup] off-site retry failed", next.offsiteError);
    return next;
  }
}

function sendBackupMail(status, localDb) {
  const to = mailTo();
  if (!to || !hasGoogleAuth() || !process.env.GMAIL_SENDER) return { emailedTo: null };
  const html = "<p>Studio Delta backup " + (status.ok ? "succeeded" : "failed") + ".</p>" +
    "<p>When: " + (status.at || "") + " (Africa/Johannesburg day " + (status.sastDate || "") + ")</p>" +
    "<p>SQLite: " + (status.bytes || 0) + " bytes, sha256 " + (status.sha256 || "") + "</p>" +
    "<p>Integrity: " + (status.integrity || "n/a") + "</p>" +
    "<p>Off-site Drive: " + (status.offsite ? "yes" : "no") + (status.driveUrl ? " — " + status.driveUrl : "") + "</p>" +
    (status.error ? "<p>Error: " + String(status.error) + "</p>" : "") +
    "<p>Restore from Users → Backup. Type RESTORE and enter the Manager access code twice. Upload the complete .tgz from Drive to restore everything.</p>";
  const attachments = [];
  if (status.ok && localDb && fileSize(localDb) <= MAIL_ATTACH_MAX) {
    attachments.push({
      name: path.basename(localDb),
      mime: "application/vnd.sqlite3",
      base64: fs.readFileSync(localDb).toString("base64")
    });
  }
  driveRpc({
    op: "sendMail",
    to,
    subject: (status.ok ? "Studio Delta backup OK" : "Studio Delta backup FAILED") + " " + (status.sastDate || ""),
    html,
    attachments
  });
  return { emailedTo: to };
}

function runBackup(reason) {
  if (running) return loadStatus() || { ok: false, error: "A backup is already running" };
  running = true;
  const when = sastDate();
  const stamp = uniqueStamp();
  const base = "studio-delta-" + stamp;
  const files = snapshotFiles(stamp);
  const localDb = files.db;
  const localJson = files.json;
  const status = {
    ok: false,
    at: new Date().toISOString(),
    sastDate: when.date,
    reason: reason || "scheduled",
    offsite: false,
    verified: false
  };
  try {
    const verified = snapshotSqlite(localDb);
    copyOfficeJson(localJson);
    const bundle = buildCompleteBundle(stamp, localDb, localJson, Object.assign({ reason: status.reason }, verified));
    status.localDb = path.basename(localDb);
    status.complete = bundle.complete;
    status.filesArchive = bundle.filesArchive;
    status.bytes = fileSize(files.complete) || fileSize(localDb);
    status.sha256 = sha256File(localDb);
    status.integrity = verified.integrity;
    status.counts = verified.counts;
    status.verified = true;
    status.packedAll = true;
    status.packedFiles = (bundle.manifest && bundle.manifest.files) || [];
    let off = { offsite: false };
    try {
      off = uploadOffsite(localDb, localJson, files.files, files.complete);
    } catch (e) {
      off = { offsite: false, offsiteError: e.message || String(e) };
      console.warn("[backup] off-site upload failed", off.offsiteError);
    }
    Object.assign(status, off);
    pruneLocal();
    status.ok = true;
    try { Object.assign(status, sendBackupMail(status, localDb)); } catch (e) {
      status.mailError = e.message || String(e);
    }
    console.log("[backup] ok", status.complete || status.localDb, "offsite", !!status.offsite);
  } catch (e) {
    status.ok = false;
    status.error = e.message || String(e);
    console.error("[backup] failed", status.error);
    try { sendBackupMail(status, null); } catch (err) {}
  } finally {
    writeJson(statusPath(), status);
    running = false;
  }
  return status;
}

function lastSuccessSastDate(status) {
  if (!status || !status.ok || !status.at) return "";
  if (status.sastDate) return status.sastDate;
  return sastDate(new Date(status.at)).date;
}

function isDue(status) {
  const now = sastDate();
  if (!status || !status.ok) return true;
  if (lastSuccessSastDate(status) === now.date) return false;
  if (now.h >= 2) return true;
  const age = Date.now() - Date.parse(status.at);
  return Number.isFinite(age) && age > 26 * 60 * 60 * 1000;
}

function tick() {
  if (running || restoring) return null;
  try {
    if (isDue(loadStatus())) return runBackup("scheduled");
    return retryOffsite();
  } catch (e) {
    console.error("[backup] tick failed", e && e.message ? e.message : e);
    return null;
  }
}

function latestCompletePath() {
  const row = listLocalSnapshots().find((s) => s.complete);
  return row ? path.join(backupsDir(), row.complete) : null;
}

function info() {
  const last = loadStatus();
  const restored = loadRestoreStatus();
  const stale = !last || !last.ok || isDue(last);
  const persist = (() => {
    try { return require("./workbook-store").storageInfo(); } catch (e) { return {}; }
  })();
  return {
    backupAt: last && last.at ? last.at : null,
    backupOk: !!(last && last.ok),
    backupOffsite: !!(last && last.offsite),
    backupOffsiteError: last && last.offsiteError ? last.offsiteError : null,
    backupStale: stale,
    backupError: last && last.error ? last.error : null,
    backupDriveUrl: last && last.driveUrl ? last.driveUrl : null,
    backupLocalCount: listLocalSnapshots().length,
    backupKeepDays: KEEP_LOCAL,
    backupGoogleConfigured: hasGoogleAuth(),
    backupServiceAccount: googleServiceAccountEmail(),
    backupDriveFolderSet: !!configuredFolderId(),
    backupDriveFolderId: configuredFolderId() || null,
    backupDriveImpersonate: impersonateEmail() || null,
    backupEmail: mailTo() || null,
    backupVerified: !!(last && last.ok && last.verified),
    backupComplete: last && last.complete ? last.complete : null,
    backupIntegrity: last && last.integrity ? last.integrity : null,
    backupCounts: last && last.counts ? last.counts : null,
    backupPackedAll: !!(last && last.packedAll),
    backupPackedFiles: last && last.packedFiles ? last.packedFiles : null,
    restoreAt: restored && restored.at ? restored.at : null,
    restoreOk: restored ? !!restored.ok : null,
    restoreError: restored && restored.error ? restored.error : null,
    usingEphemeralDisk: !!persist.usingEphemeralDisk,
    volumeWarning: persist.warning || null,
    dataDir: persist.dataDir || dataDir()
  };
}

function safeBackupName(name) {
  const base = path.basename(String(name || ""));
  if (!SAFE_NAME.test(base)) return null;
  const full = path.join(backupsDir(), base);
  if (!fs.existsSync(full)) return null;
  return full;
}

function requireConfirm(confirm) {
  if (String(confirm || "").trim().toUpperCase() !== CONFIRM_WORD) {
    throw new Error("Type RESTORE to restore this backup.");
  }
}

function inspectSource(sourcePath) {
  if (!sourcePath || !fs.existsSync(sourcePath)) throw new Error("That backup file is not on this volume.");
  const lower = String(sourcePath).toLowerCase();
  if (lower.endsWith(".db")) {
    const verified = verifySqliteFile(sourcePath);
    return { kind: "sqlite", verified, sourcePath };
  }
  if (lower.endsWith(".tgz") || lower.endsWith(".tar.gz") || lower.endsWith(".gz")) {
    const staging = path.join(dataDir(), "restore-inspect-" + Date.now());
    try {
      tarExtract(sourcePath, staging);
      const dbFile = fs.existsSync(path.join(staging, "studio-delta.db"))
        ? path.join(staging, "studio-delta.db")
        : null;
      if (!dbFile) throw new Error("That archive has no studio-delta.db. It is not a Studio Delta backup.");
      const verified = verifySqliteFile(dbFile);
      const manifest = readJson(path.join(staging, "manifest.json"), null);
      return { kind: "complete", verified, manifest, sourcePath };
    } finally {
      try { fs.rmSync(staging, { recursive: true, force: true }); } catch (e) {}
    }
  }
  throw new Error("Restore a complete .tgz backup, or a .db snapshot.");
}

function applyExtracted(staging) {
  const dbSrc = path.join(staging, "studio-delta.db");
  if (!fs.existsSync(dbSrc)) throw new Error("The backup has no studio-delta.db");
  verifySqliteFile(dbSrc);
  sqlite.close();
  const liveDb = sqlite.sqlitePath();
  replaceFile(dbSrc, liveDb);
  rmQuiet(liveDb + "-wal");
  rmQuiet(liveDb + "-shm");
  fs.readdirSync(staging, { withFileTypes: true }).forEach((ent) => {
    if (ent.name === "studio-delta.db" || ent.name === "manifest.json") return;
    const src = path.join(staging, ent.name);
    const dest = path.join(dataDir(), ent.name);
    if (ent.isDirectory()) replaceDir(src, dest);
    else if (ent.isFile()) replaceFile(src, dest);
  });
  const liveCheck = verifySqliteFile(liveDb);
  sqlite.reopen();
  require("./workbook-store").reloadWorkbook();
  require("./db").reloadOfficeState();
  try { require("./gas").clearShopCache(); } catch (e) {}
  return liveCheck;
}

function applySqliteOnly(dbFile) {
  verifySqliteFile(dbFile);
  sqlite.close();
  const liveDb = sqlite.sqlitePath();
  replaceFile(dbFile, liveDb);
  rmQuiet(liveDb + "-wal");
  rmQuiet(liveDb + "-shm");
  const liveCheck = verifySqliteFile(liveDb);
  sqlite.reopen();
  require("./workbook-store").reloadWorkbook();
  require("./db").reloadOfficeState();
  try { require("./gas").clearShopCache(); } catch (e) {}
  return liveCheck;
}

function reloadKeptSession(kept) {
  try { require("./staff").reloadSessionsKeeping(kept || null); } catch (e) {}
}

function restoreFromPath(sourcePath, opts) {
  opts = opts || {};
  requireConfirm(opts.confirm);
  if (restoring) throw new Error("A restore is already running");
  if (running) throw new Error("Wait for the backup that is already running, then restore.");
  const inspected = inspectSource(sourcePath);
  let safety = null;
  if (!opts.skipSafety) {
    safety = runBackup("pre-restore");
    if (!safety.ok) {
      throw new Error("Could not snapshot the live shop before restore. Restore was not started. " + (safety.error || ""));
    }
  }
  restoring = true;
  const result = {
    ok: false,
    at: new Date().toISOString(),
    source: path.basename(sourcePath),
    safety: safety && (safety.complete || safety.localDb) ? (safety.complete || safety.localDb) : null,
    rolledBack: false
  };
  const staging = path.join(dataDir(), "restore-apply-" + Date.now());
  try {
    let liveCheck;
    if (inspected.kind === "sqlite") {
      liveCheck = applySqliteOnly(sourcePath);
    } else {
      tarExtract(sourcePath, staging);
      liveCheck = applyExtracted(staging);
    }
    reloadKeptSession(opts.keepSession);
    result.ok = true;
    result.integrity = liveCheck.integrity;
    result.counts = liveCheck.counts;
    console.log("[restore] ok from", result.source);
  } catch (e) {
    result.ok = false;
    result.error = e.message || String(e);
    console.error("[restore] failed", result.error);
    if (safety && !opts.skipSafety) {
      const safetyPath = safety.complete
        ? path.join(backupsDir(), safety.complete)
        : (safety.localDb ? path.join(backupsDir(), safety.localDb) : null);
      if (safetyPath && fs.existsSync(safetyPath)) {
        try {
          restoring = false;
          restoreFromPath(safetyPath, { confirm: CONFIRM_WORD, skipSafety: true, keepSession: opts.keepSession });
          result.rolledBack = true;
          result.error = (result.error || "Restore failed") + " Live shop was rolled back to the safety copy taken just before restore.";
        } catch (err) {
          result.rollbackError = err.message || String(err);
        }
      }
    }
    if (!result.ok) throw new Error(result.error);
  } finally {
    restoring = false;
    try { fs.rmSync(staging, { recursive: true, force: true }); } catch (e) {}
    writeJson(restoreLogPath(), result);
  }
  return result;
}

function restoreNamed(name, opts) {
  const full = safeBackupName(name);
  if (!full) throw new Error("That backup file is not on this volume.");
  return restoreFromPath(full, opts);
}

function saveUploadedBackup(buffer, filename) {
  const raw = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer || []);
  if (!raw.length) throw new Error("No backup file was uploaded.");
  const stamp = uniqueStamp();
  const lower = String(filename || "").toLowerCase();
  const ext = lower.endsWith(".db") ? ".db" : ".tgz";
  const dest = path.join(backupsDir(), "studio-delta-uploaded-" + stamp + ext);
  fs.writeFileSync(dest, raw);
  inspectSource(dest);
  return dest;
}

module.exports = {
  backupsDir,
  runBackup,
  tick,
  isDue,
  info,
  loadStatus,
  loadRestoreStatus,
  listLocalSnapshots,
  safeBackupName,
  sastDate,
  verifySqliteFile,
  inspectSource,
  restoreFromPath,
  restoreNamed,
  saveUploadedBackup,
  latestCompletePath,
  testDriveUpload,
  retryOffsite,
  normalizeFolderId,
  impersonateEmail,
  explainDriveError,
  CONFIRM_WORD
};
