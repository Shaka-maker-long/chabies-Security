const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { spawnSync } = require("child_process");
const { dataDir, persistWorkbook, hasGoogleAuth } = require("./workbook-store");
const sqlite = require("./sqlite-store");

const KEEP_LOCAL = 14;
const KEEP_DRIVE = 14;
const SAST_OFFSET_MS = 2 * 60 * 60 * 1000;
const FILES_ARCHIVE_MAX = 250 * 1024 * 1024;
const MAIL_ATTACH_MAX = 12 * 1024 * 1024;
const FOLDER_NAME = "Studio Delta ERP backups";
const SAFE_NAME = /^studio-delta-[A-Za-z0-9._-]+$/;

let running = false;

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
    stamp: sast.getUTCFullYear() + p(sast.getUTCMonth() + 1) + p(sast.getUTCDate()) + "-" + p(sast.getUTCHours()) + p(sast.getUTCMinutes())
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

function folderStatePath() {
  return path.join(backupsDir(), "drive-folder.json");
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

function sha256File(file) {
  const hash = crypto.createHash("sha256");
  hash.update(fs.readFileSync(file));
  return hash.digest("hex");
}

function fileSize(file) {
  try { return fs.statSync(file).size; } catch (e) { return 0; }
}

function listLocalSnapshots() {
  const dir = backupsDir();
  return fs.readdirSync(dir)
    .filter((name) => name.startsWith("studio-delta-") && name.endsWith(".db"))
    .map((name) => {
      const full = path.join(dir, name);
      const st = fs.statSync(full);
      return { name, bytes: st.size, at: st.mtime.toISOString() };
    })
    .sort((a, b) => b.at.localeCompare(a.at));
}

function pruneLocal() {
  const keep = listLocalSnapshots().slice(KEEP_LOCAL);
  keep.forEach((row) => {
    const stamp = row.name.replace(/^studio-delta-/, "").replace(/\.db$/, "");
    ["studio-delta-" + stamp + ".db", "studio-delta-" + stamp + ".json", "studio-delta-" + stamp + "-files.tgz"].forEach((name) => {
      const full = path.join(backupsDir(), name);
      try { fs.unlinkSync(full); } catch (e) {}
    });
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
}

function copyOfficeJson(dest) {
  try {
    writeJson(dest, require("./db").railwayBackup());
  } catch (e) {
    const office = path.join(dataDir(), "studio-delta.json");
    if (fs.existsSync(office)) fs.copyFileSync(office, dest);
  }
}

function archiveEnquiryFiles(dest) {
  const quotes = path.join(dataDir(), "enquiry-quotes");
  const files = path.join(dataDir(), "enquiry-files");
  const parts = [];
  if (fs.existsSync(quotes)) parts.push("enquiry-quotes");
  if (fs.existsSync(files)) parts.push("enquiry-files");
  if (!parts.length) return { path: null, bytes: 0, skipped: "no enquiry files" };
  const r = spawnSync("tar", ["-czf", dest, "-C", dataDir()].concat(parts), { encoding: "utf8" });
  if (r.status !== 0) throw new Error(r.stderr || "Could not archive enquiry files");
  const bytes = fileSize(dest);
  if (bytes > FILES_ARCHIVE_MAX) {
    try { fs.unlinkSync(dest); } catch (e) {}
    return { path: null, bytes, skipped: "enquiry files larger than 250 MB" };
  }
  return { path: dest, bytes, skipped: null };
}

function driveRpc(payload) {
  const r = spawnSync(process.execPath, [path.join(__dirname, "drive-cli.js")], {
    input: JSON.stringify(payload),
    encoding: "utf8",
    env: process.env,
    maxBuffer: 8 * 1024 * 1024
  });
  if (!r.stdout) throw new Error(r.stderr || "Drive helper failed");
  const parsed = JSON.parse(r.stdout);
  if (!parsed.ok) throw new Error(parsed.error || "Drive helper error");
  return parsed;
}

function shareEmail() {
  return String(process.env.BACKUP_SHARE_EMAIL || process.env.BACKUP_EMAIL || process.env.GMAIL_SENDER || "").trim();
}

function mailTo() {
  return String(process.env.BACKUP_EMAIL || process.env.GMAIL_SENDER || "").trim();
}

function ensureDriveFolder() {
  const configured = String(process.env.BACKUP_DRIVE_FOLDER_ID || "").trim();
  if (configured) return { id: configured, created: false };
  const saved = readJson(folderStatePath(), null);
  if (saved && saved.id) return { id: saved.id, created: false };
  const listed = driveRpc({ op: "listFoldersByName", name: FOLDER_NAME });
  const existing = (listed.files || [])[0];
  if (existing && existing.id) {
    writeJson(folderStatePath(), { id: existing.id, name: FOLDER_NAME });
    return { id: existing.id, created: false };
  }
  const created = driveRpc({ op: "createFolder", name: FOLDER_NAME });
  const email = shareEmail();
  if (email) {
    try { driveRpc({ op: "shareWithEmail", fileId: created.id, email, role: "writer" }); } catch (e) {
      console.warn("[backup] could not share Drive folder", e.message || e);
    }
  }
  writeJson(folderStatePath(), { id: created.id, name: FOLDER_NAME, url: created.url || null });
  return { id: created.id, created: true, url: created.url || null };
}

function pruneDrive(folderId) {
  const listed = driveRpc({ op: "listFiles", folderId });
  const ours = (listed.files || [])
    .filter((f) => String(f.name || "").indexOf("studio-delta-") === 0 && String(f.name).slice(-3) === ".db")
    .sort((a, b) => String(b.createdTime || "").localeCompare(String(a.createdTime || "")));
  ours.slice(KEEP_DRIVE).forEach((f) => {
    try { driveRpc({ op: "trashFile", fileId: f.id }); } catch (e) {}
  });
}

function uploadOffsite(localDb, localJson, archive) {
  if (!hasGoogleAuth()) return { offsite: false, reason: "Google Drive is not configured on Railway" };
  const folder = ensureDriveFolder();
  const dbUp = driveRpc({
    op: "uploadFile",
    path: localDb,
    name: path.basename(localDb),
    folderId: folder.id,
    mimeType: "application/vnd.sqlite3"
  });
  let jsonUrl = null;
  if (localJson && fs.existsSync(localJson)) {
    try {
      const jsonUp = driveRpc({
        op: "uploadFile",
        path: localJson,
        name: path.basename(localJson),
        folderId: folder.id,
        mimeType: "application/json"
      });
      jsonUrl = jsonUp.url || null;
    } catch (e) {
      console.warn("[backup] Drive JSON upload failed", e.message || e);
    }
  }
  let filesUrl = null;
  if (archive && archive.path) {
    try {
      const zipUp = driveRpc({
        op: "uploadFile",
        path: archive.path,
        name: path.basename(archive.path),
        folderId: folder.id,
        mimeType: "application/gzip"
      });
      filesUrl = zipUp.url || null;
    } catch (e) {
      console.warn("[backup] Drive files archive upload failed", e.message || e);
    }
  }
  try { pruneDrive(folder.id); } catch (e) {
    console.warn("[backup] Drive prune failed", e.message || e);
  }
  return {
    offsite: true,
    driveId: dbUp.id,
    driveUrl: dbUp.url || null,
    jsonUrl,
    filesUrl,
    folderId: folder.id
  };
}

function sendBackupMail(status, localDb) {
  const to = mailTo();
  if (!to || !hasGoogleAuth() || !process.env.GMAIL_SENDER) return { emailedTo: null };
  const html = "<p>Studio Delta backup " + (status.ok ? "succeeded" : "failed") + ".</p>" +
    "<p>When: " + (status.at || "") + " (Africa/Johannesburg day " + (status.sastDate || "") + ")</p>" +
    "<p>SQLite: " + (status.bytes || 0) + " bytes, sha256 " + (status.sha256 || "") + "</p>" +
    "<p>Off-site Drive: " + (status.offsite ? "yes" : "no") + (status.driveUrl ? " — " + status.driveUrl : "") + "</p>" +
    (status.error ? "<p>Error: " + String(status.error) + "</p>" : "") +
    "<p>This is an automatic copy. Keep this email. Restore by replacing studio-delta.db on the Railway volume from a downloaded copy.</p>";
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
  const base = "studio-delta-" + when.stamp;
  const localDb = path.join(backupsDir(), base + ".db");
  const localJson = path.join(backupsDir(), base + ".json");
  const localTar = path.join(backupsDir(), base + "-files.tgz");
  const status = {
    ok: false,
    at: new Date().toISOString(),
    sastDate: when.date,
    reason: reason || "scheduled",
    offsite: false
  };
  try {
    snapshotSqlite(localDb);
    copyOfficeJson(localJson);
    const archive = archiveEnquiryFiles(localTar);
    status.localDb = path.basename(localDb);
    status.bytes = fileSize(localDb);
    status.sha256 = sha256File(localDb);
    status.filesArchive = archive.path ? path.basename(archive.path) : null;
    status.filesSkipped = archive.skipped || null;
    let off = { offsite: false };
    try {
      off = uploadOffsite(localDb, localJson, archive);
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
    console.log("[backup] ok", status.localDb, "offsite", !!status.offsite);
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
  if (running) return null;
  try {
    if (!isDue(loadStatus())) return null;
    return runBackup("scheduled");
  } catch (e) {
    console.error("[backup] tick failed", e && e.message ? e.message : e);
    return null;
  }
}

function info() {
  const last = loadStatus();
  const stale = !last || !last.ok || isDue(last);
  return {
    backupAt: last && last.at ? last.at : null,
    backupOk: !!(last && last.ok),
    backupOffsite: !!(last && last.offsite),
    backupStale: stale,
    backupError: last && last.error ? last.error : null,
    backupDriveUrl: last && last.driveUrl ? last.driveUrl : null,
    backupLocalCount: listLocalSnapshots().length,
    backupKeepDays: KEEP_LOCAL,
    backupGoogleConfigured: hasGoogleAuth(),
    backupEmail: mailTo() || null
  };
}

function safeBackupName(name) {
  const base = path.basename(String(name || ""));
  if (!SAFE_NAME.test(base)) return null;
  const full = path.join(backupsDir(), base);
  if (!fs.existsSync(full)) return null;
  return full;
}

module.exports = {
  backupsDir,
  runBackup,
  tick,
  isDue,
  info,
  loadStatus,
  listLocalSnapshots,
  safeBackupName,
  sastDate
};
