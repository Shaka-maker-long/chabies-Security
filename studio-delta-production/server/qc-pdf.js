"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { spawnSync } = require("child_process");
const { dataDir, getBook, persistWorkbook, hasGoogleAuth } = require("./workbook-store");

const PRE_PHOTOS = [
  "Front", "Left Side", "Right Side", "Back", "Open", "Top (Optional)", "Level 1", "Level 2 (Optional)"
];
const FINAL_PHOTOS = [
  "Front", "Level 1", "Back", "Left Side", "Right Side", "Job Card", "Open", "Top (Optional)", "Level 2 (Optional)"
];
const QC_DRIVE_FOLDER_ID = process.env.QC_PDF_DRIVE_FOLDER_ID || "1pyzJ-jcgltJlrCOwjR7c8AIFxwcd2YcK";
const QC_QUEUE_FOLDER_ID = process.env.QC_QUEUE_FOLDER_ID || "1MRl3nX7-4d8dmrjQU0UrCbzCf6Ilymub";

function storePath() {
  return path.join(dataDir(), "qc-pdfs.json");
}

function pdfDir() {
  const dir = path.join(dataDir(), "qc-pdfs");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function photosDir(id) {
  const dir = path.join(pdfDir(), String(id || ""));
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function loadStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(storePath(), "utf8"));
    if (parsed && Array.isArray(parsed.records)) return parsed;
  } catch (e) {
    if (e && e.code !== "ENOENT") console.error("[qc-pdf] read", e.message || e);
  }
  return { records: [] };
}

function saveStore(store) {
  const file = storePath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(store));
  fs.renameSync(tmp, file);
  return store;
}

function newId() {
  return crypto.randomBytes(12).toString("hex");
}

function pdfUrlFor(id) {
  return "/api/qc-pdfs/" + encodeURIComponent(id) + "/pdf";
}

function qcKind(processName) {
  const s = String(processName || "").toLowerCase();
  if (s.indexOf("final") !== -1) return "Final QC";
  return "Pre-powder coating QC";
}

function qcLabel(processName) {
  const s = String(processName || "").toLowerCase();
  if (s.indexOf("final") !== -1) return "Final QC PDF";
  if (s.indexOf("pre-powder") !== -1 || s.indexOf("pre powder") !== -1) return "Pre-powder QC PDF";
  return "QC PDF";
}

function photoLabels(processName) {
  return qcKind(processName) === "Final QC" ? FINAL_PHOTOS : PRE_PHOTOS;
}

function filePayload(file) {
  if (!file) return "";
  if (Buffer.isBuffer(file)) return file;
  if (typeof file === "string") return file;
  if (file.data != null) return file.data;
  if (file.dataUrl) return file.dataUrl;
  if (file.base64) return file.base64;
  return "";
}

function isJpeg(buf) {
  return Buffer.isBuffer(buf) && buf.length >= 3 && buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff;
}

function isPng(buf) {
  return Buffer.isBuffer(buf) && buf.length >= 4 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47;
}

function decodeImage(raw) {
  if (Buffer.isBuffer(raw)) {
    if (!raw.length) return null;
    return { mime: isPng(raw) ? "image/png" : "image/jpeg", buf: raw };
  }
  const s = String(raw || "").trim();
  if (!s) return null;
  if (/^https?:\/\//i.test(s) || /^\/api\//i.test(s)) return null;
  const dataUrl = s.match(/^data:([^;]+);base64,([\s\S]+)$/i);
  if (dataUrl) {
    const buf = Buffer.from(String(dataUrl[2]).replace(/\s/g, ""), "base64");
    if (!buf.length) return null;
    return { mime: dataUrl[1], buf };
  }
  const compact = s.replace(/\s/g, "");
  if (compact.length < 32 || /[^A-Za-z0-9+/=]/.test(compact)) return null;
  const buf = Buffer.from(compact, "base64");
  if (!buf.length) return null;
  return { mime: "image/jpeg", buf };
}

async function toJpeg(raw) {
  const decoded = decodeImage(raw);
  if (!decoded || !decoded.buf || decoded.buf.length < 24) return null;
  try {
    const sharp = require("sharp");
    return await sharp(decoded.buf)
      .rotate()
      .resize({ width: 1600, height: 1600, fit: "inside", withoutEnlargement: true })
      .flatten({ background: "#ffffff" })
      .jpeg({ quality: 82 })
      .toBuffer();
  } catch (e) {
    if (isJpeg(decoded.buf) || isPng(decoded.buf)) return decoded.buf;
    return null;
  }
}

function formatWhen(iso) {
  try {
    return new Intl.DateTimeFormat("en-GB", {
      timeZone: process.env.TZ || "Africa/Johannesburg",
      day: "2-digit",
      month: "short",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false
    }).format(iso ? new Date(iso) : new Date());
  } catch (e) {
    return String(iso || "");
  }
}

async function packPhotos(processName, filesData) {
  const labels = photoLabels(processName);
  const out = [];
  const files = filesData || [];
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    if (!file) continue;
    const jpeg = await toJpeg(filePayload(file));
    if (!jpeg) continue;
    const name = labels[i] || (file && file.name) || ("Photo " + (i + 1));
    out.push({ name, buf: jpeg });
  }
  return out;
}

function writePhotoFiles(id, photos, signatureBuf) {
  const dir = photosDir(id);
  (photos || []).forEach((photo, i) => {
    if (!photo || !photo.buf) return;
    const safe = String(photo.name || "photo").replace(/[^\w.-]+/g, "_").slice(0, 40);
    const name = String(i + 1).padStart(2, "0") + "-" + (safe || "photo") + ".jpg";
    fs.writeFileSync(path.join(dir, name), photo.buf);
  });
  if (signatureBuf && signatureBuf.length) {
    fs.writeFileSync(path.join(dir, "signature.jpg"), signatureBuf);
  }
}

async function renderPdf(record, dest) {
  const PDFDocument = require("pdfkit");
  await new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: "A4", margin: 36, info: {
      Title: (record.order_number || "Order") + " " + (record.kind || "QC"),
      Author: "Studio Delta"
    } });
    const stream = fs.createWriteStream(dest);
    doc.pipe(stream);
    stream.on("finish", resolve);
    stream.on("error", reject);
    doc.on("error", reject);

    doc.font("Helvetica-Bold").fontSize(18).text("Studio Delta");
    doc.moveDown(0.2);
    doc.font("Helvetica-Bold").fontSize(14).text(record.kind || "QC Report");
    doc.moveDown(0.6);
    doc.font("Helvetica").fontSize(11);
    doc.text("Order: " + (record.order_number || ""));
    if (record.product) doc.text("Product: " + record.product);
    doc.text("Completed by: " + (record.worker || ""));
    doc.text("Date: " + formatWhen(record.created_at));
    doc.moveDown();
    doc.font("Helvetica-Bold").fontSize(12).text("Checklist");
    doc.moveDown(0.3);
    (record.answers || []).forEach((row, i) => {
      const q = String(row.q || "Q" + (i + 1));
      let a = String(row.a || "");
      if (a === "Y") a = "Yes";
      if (a === "N") a = "No";
      doc.font("Helvetica").fontSize(10).text((i + 1) + ". " + q);
      doc.font("Helvetica-Bold").text(a || "—", { indent: 14 });
      doc.moveDown(0.25);
    });

    if (record.signatureBuf) {
      doc.addPage();
      doc.font("Helvetica-Bold").fontSize(12).text("Signature");
      doc.moveDown(0.4);
      try {
        doc.image(record.signatureBuf, { fit: [240, 120] });
      } catch (e) {
        doc.font("Helvetica").fontSize(10).text("Signature could not be placed.");
      }
    }

    const photos = record.photos || [];
    photos.forEach((photo) => {
      if (!photo || !photo.buf) return;
      doc.addPage();
      doc.font("Helvetica-Bold").fontSize(12).text(photo.name || "Photo");
      doc.moveDown(0.4);
      try {
        doc.image(photo.buf, { fit: [520, 680] });
      } catch (e) {
        doc.font("Helvetica").fontSize(10).text("This photo could not be placed.");
      }
    });

    if (record.photoAttempted && !photos.length) {
      doc.addPage();
      doc.font("Helvetica-Bold").fontSize(12).text("Photos");
      doc.moveDown(0.4);
      doc.font("Helvetica").fontSize(10).text("Photos were sent with this QC but could not be placed on the PDF.");
    }

    doc.end();
  });
}

function displayName(rec) {
  return (rec.order_number || "Order") + " · " + (rec.kind || "QC") + (rec.worker ? " · " + rec.worker : "");
}

function isoFromLog(value) {
  if (!value) return new Date().toISOString();
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return new Date().toISOString();
  return d.toISOString();
}

async function saveFromFinish(job) {
  const order = String((job && job.orderNum) || "").trim();
  if (!order) throw new Error("Order number is required.");
  const id = newId();
  const filename = id + ".pdf";
  const pdfPath = path.join(pdfDir(), filename);
  const rec = {
    id,
    order_number: order,
    worker: String((job && job.workerName) || "").trim(),
    process: String((job && job.processName) || "").trim(),
    kind: qcKind(job && job.processName),
    product: String((job && job.productName) || "").trim(),
    answers: Array.isArray(job && job.qcData) ? job.qcData.map((row) => ({ q: row.q, a: row.a })) : [],
    created_at: (job && job.created_at) ? isoFromLog(job.created_at) : new Date().toISOString(),
    pdf_path: pdfPath,
    log_id: String((job && job.logId) || ""),
    row: Number(job && job.rowToUpdate) || 0,
    photo_count: 0,
    imported: false
  };
  const files = (job && job.filesData) || [];
  const photoAttempted = files.some((file) => !!file);
  const photos = await packPhotos(rec.process, files);
  const signatureBuf = await toJpeg(job && job.signatureUrl);
  rec.photo_count = photos.length;
  rec.has_signature = !!signatureBuf;
  writePhotoFiles(id, photos, signatureBuf);
  await renderPdf({ ...rec, photos, signatureBuf, photoAttempted }, pdfPath);
  const store = loadStore();
  store.records.unshift(rec);
  saveStore(store);
  return {
    id,
    url: pdfUrlFor(id),
    name: displayName(rec),
    label: qcLabel(rec.process),
    order_number: order,
    dateCreated: Date.parse(rec.created_at) || Date.now(),
    photo_count: rec.photo_count
  };
}

function writeUrlOnLog(rowToUpdate, url) {
  const row = Number(rowToUpdate) || 0;
  if (row < 2 || !url) return false;
  const book = getBook();
  const sheet = book.getSheetByName("Production_Log");
  if (!sheet) return false;
  const current = String(sheet.getRange(row, 8).getValue() || "");
  if (current.indexOf(url) !== -1) return true;
  const next = current ? (current.replace(/\s+$/, "") + "\n\nQC PDF: " + url) : ("QC PDF: " + url);
  sheet.getRange(row, 8).setValue(next);
  persistWorkbook();
  return true;
}

function listReports() {
  return loadStore().records
    .filter((rec) => rec && rec.id && rec.pdf_path && fs.existsSync(rec.pdf_path))
    .map((rec) => ({
      id: rec.id,
      name: displayName(rec),
      dateCreated: Date.parse(rec.created_at) || 0,
      order_number: rec.order_number || "",
      label: qcLabel(rec.process),
      worker: rec.worker || "",
      kind: rec.kind || "",
      log_id: rec.log_id || "",
      process: rec.process || "",
      photo_count: Number(rec.photo_count) || 0,
      imported: !!rec.imported,
      drive_url: rec.drive_url || "",
      url: (!(Number(rec.photo_count) > 0 || rec.imported) && rec.drive_url) ? rec.drive_url : pdfUrlFor(rec.id)
    }))
    .sort((a, b) => (b.dateCreated || 0) - (a.dateCreated || 0));
}

function readPdf(id) {
  const rec = loadStore().records.find((row) => row && row.id === String(id || ""));
  if (!rec || !rec.pdf_path || !fs.existsSync(rec.pdf_path)) return null;
  return {
    filename: (rec.order_number || "QC") + "_" + String(rec.kind || "QC").replace(/\s+/g, "_") + ".pdf",
    buffer: fs.readFileSync(rec.pdf_path)
  };
}

function pdfsForOrder(orderNumber) {
  const want = String(orderNumber || "").trim().toLowerCase();
  if (!want) return [];
  return listReports().filter((row) => String(row.order_number || "").trim().toLowerCase() === want);
}

function parseAnswersFromNotes(notes) {
  const out = [];
  String(notes || "").split(/\r?\n/).forEach((line) => {
    const t = line.trim();
    if (!t) return;
    if (/^QC PDF:/i.test(t) || /^Error adding to PDF Queue/i.test(t)) return;
    const idx = t.lastIndexOf(": ");
    if (idx < 1) return;
    const q = t.slice(0, idx).trim();
    const a = t.slice(idx + 2).trim();
    if (!q || !a) return;
    out.push({ q, a });
  });
  return out;
}

function notesHaveLocalPdf(notes) {
  return /QC PDF:\s*\/api\/qc-pdfs\//i.test(String(notes || ""));
}

function parseDriveFileId(url) {
  const s = String(url || "");
  const file = s.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (file) return file[1];
  const id = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return id ? id[1] : "";
}

function driveIdsFromNotes(notes) {
  const ids = [];
  String(notes || "").replace(/https?:\/\/[^\s]+/gi, (url) => {
    const id = parseDriveFileId(url);
    if (id && ids.indexOf(id) === -1) ids.push(id);
    return url;
  });
  return ids;
}

function needsPhotoRecover(rec) {
  if (!rec || rec.imported || rec.drive_import_tried) return false;
  return !(Number(rec.photo_count) > 0);
}

function attachDriveUrl(rec, urlOrId) {
  if (!rec) return false;
  const raw = String(urlOrId || "").trim();
  const parsed = parseDriveFileId(raw);
  const id = parsed || (/^[a-zA-Z0-9_-]{6,}$/.test(raw) && !/^https?:/i.test(raw) ? raw : "");
  const url = parsed
    ? ("https://drive.google.com/file/d/" + parsed + "/view")
    : (/^https?:\/\//i.test(raw) ? raw : (id ? ("https://drive.google.com/file/d/" + id + "/view") : ""));
  if (!url) return false;
  if (rec.drive_url === url) return false;
  rec.drive_url = url;
  const store = loadStore();
  const idx = store.records.findIndex((row) => row && row.id === rec.id);
  if (idx >= 0) {
    store.records[idx].drive_url = url;
    saveStore(store);
  }
  return true;
}

function attachDriveUrlsFromLogs() {
  const book = getBook();
  const sheet = book.getSheetByName("Production_Log");
  if (!sheet || sheet.getLastRow() < 2) return { saved: 0 };
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  let saved = 0;
  for (let i = 0; i < values.length; i++) {
    const notes = String(values[i][7] || "");
    const ids = driveIdsFromNotes(notes);
    if (!ids.length) continue;
    const rec = existingForLog(String(values[i][0] || ""), i + 2);
    if (!rec) continue;
    if ((Number(rec.photo_count) || 0) > 0 || rec.imported) continue;
    if (attachDriveUrl(rec, ids[0])) saved += 1;
  }
  return { saved };
}

function existingForLog(logId, row) {
  const wantLog = String(logId || "").trim();
  const wantRow = Number(row) || 0;
  return loadStore().records.find((rec) => {
    if (!rec || !rec.id || !rec.pdf_path || !fs.existsSync(rec.pdf_path)) return false;
    if (wantLog && rec.log_id && String(rec.log_id) === wantLog) return true;
    if (wantRow >= 2 && Number(rec.row) === wantRow) return true;
    return false;
  }) || null;
}

function tryDrive(payload) {
  if (!hasGoogleAuth()) return null;
  try {
    const r = spawnSync(process.execPath, [path.join(__dirname, "drive-cli.js")], {
      input: JSON.stringify(payload),
      encoding: "utf8",
      env: process.env,
      maxBuffer: 40 * 1024 * 1024
    });
    if (!r.stdout) return null;
    const parsed = JSON.parse(r.stdout);
    if (!parsed || !parsed.ok) return null;
    return parsed;
  } catch (e) {
    return null;
  }
}

function matchDriveName(name, orderNum, kind) {
  const n = String(name || "").toLowerCase().replace(/\.pdf$/i, "");
  const o = String(orderNum || "").toLowerCase();
  if (!o || n.indexOf(o) === -1) return false;
  const k = String(kind || "").toLowerCase();
  if (k.indexOf("final") !== -1) return n.indexOf("final") !== -1;
  if (k.indexOf("pre") !== -1) return n.indexOf("pre") !== -1;
  return true;
}

let drivePdfList = undefined;
function listDriveQcPdfs() {
  if (drivePdfList !== undefined) return drivePdfList;
  const listed = tryDrive({ op: "listFiles", folderId: QC_DRIVE_FOLDER_ID });
  drivePdfList = (listed && listed.files) || [];
  return drivePdfList;
}

function pdfBufferFromDrive(got) {
  if (!got || !got.base64) return null;
  const buf = Buffer.from(got.base64, "base64");
  if (buf.length < 100 || buf.slice(0, 4).toString() !== "%PDF") return null;
  return buf;
}

function downloadDriveFilePdf(fileId) {
  const id = String(fileId || "").trim();
  if (!id) return null;
  return pdfBufferFromDrive(tryDrive({ op: "downloadFile", fileId: id }));
}

function downloadDriveQcPdf(orderNum, kind) {
  const files = listDriveQcPdfs();
  const hit = files.find((file) => matchDriveName(file && file.name, orderNum, kind));
  if (!hit || !hit.id) return null;
  return downloadDriveFilePdf(hit.id);
}

function logNotesForRecord(rec) {
  const empty = { notes: "", signature: "" };
  try {
    const book = getBook();
    const sheet = book.getSheetByName("Production_Log");
    if (!sheet || sheet.getLastRow() < 2) return empty;
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    const wantLog = String((rec && rec.log_id) || "").trim();
    const wantRow = Number(rec && rec.row) || 0;
    const wantOrder = String((rec && rec.order_number) || "").trim().toLowerCase();
    const wantKind = String((rec && rec.kind) || (rec && rec.process) || "").toLowerCase();
    let fallback = empty;
    for (let i = 0; i < values.length; i++) {
      const row = i + 2;
      const logId = String(values[i][0] || "");
      const order = String(values[i][1] || "").trim().toLowerCase();
      const process = String(values[i][4] || "").toLowerCase();
      const notes = String(values[i][7] || "");
      const signature = String(values[i][8] || "");
      if (wantLog && logId === wantLog) return { notes, signature };
      if (wantRow >= 2 && row === wantRow) return { notes, signature };
      if (wantOrder && order === wantOrder) {
        const wantFinal = wantKind.indexOf("final") !== -1;
        const isFinal = process.indexOf("final") !== -1;
        if (!wantKind || wantFinal === isFinal) fallback = { notes, signature };
      }
    }
    return fallback;
  } catch (e) {
    return empty;
  }
}

let queueList = undefined;
function listQueueJobs() {
  if (queueList !== undefined) return queueList;
  const listed = tryDrive({ op: "listFiles", folderId: QC_QUEUE_FOLDER_ID, mimeType: "text/plain" });
  queueList = (listed && listed.files) || [];
  return queueList;
}

function queuePhotosForOrder(orderNum) {
  const want = String(orderNum || "").trim().toLowerCase();
  if (!want) return [];
  const files = listQueueJobs();
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const name = String((file && file.name) || "").toLowerCase();
    if (name.indexOf(want) === -1) continue;
    const got = tryDrive({ op: "getFileText", fileId: file.id });
    if (!got || !got.text) continue;
    try {
      const job = JSON.parse(got.text);
      if (String((job && job.orderNum) || "").trim().toLowerCase() !== want) continue;
      if (Array.isArray(job.filesData) && job.filesData.some((row) => !!row)) return job.filesData;
    } catch (e) {}
  }
  return [];
}

function replacePdfBytes(rec, buf, meta) {
  if (!rec || !rec.pdf_path || !buf) return false;
  fs.writeFileSync(rec.pdf_path, buf);
  Object.assign(rec, meta || {});
  const store = loadStore();
  const idx = store.records.findIndex((row) => row && row.id === rec.id);
  if (idx >= 0) store.records[idx] = { ...store.records[idx], ...rec };
  saveStore(store);
  return true;
}

async function rebuildRecordFromPhotos(rec, filesData, signatureUrl) {
  const photos = await packPhotos(rec.process, filesData);
  const signatureBuf = await toJpeg(signatureUrl);
  if (!photos.length && !signatureBuf) return false;
  writePhotoFiles(rec.id, photos, signatureBuf);
  await renderPdf({
    ...rec,
    photos,
    signatureBuf,
    photoAttempted: !!(filesData && filesData.some((row) => !!row))
  }, rec.pdf_path);
  rec.photo_count = photos.length;
  rec.has_signature = !!signatureBuf;
  rec.imported = false;
  rec.drive_import_tried = true;
  const store = loadStore();
  const idx = store.records.findIndex((row) => row && row.id === rec.id);
  if (idx >= 0) store.records[idx] = rec;
  saveStore(store);
  return photos.length > 0;
}

async function recoverPhotosForRecord(rec, signatureUrl, notes) {
  if (!needsPhotoRecover(rec)) return false;
  if (!hasGoogleAuth()) return false;
  const log = logNotesForRecord(rec);
  const sig = signatureUrl || log.signature;
  const noteText = notes || log.notes;
  rec.drive_import_tried = true;
  const queued = queuePhotosForOrder(rec.order_number);
  if (queued.length) {
    try {
      if (await rebuildRecordFromPhotos(rec, queued, sig)) return true;
    } catch (e) {
      console.error("[qc-pdf] queue recover", rec.order_number, e.message || e);
    }
  }
  const noteIds = driveIdsFromNotes(noteText);
  if (noteIds.length) attachDriveUrl(rec, noteIds[0]);
  for (let i = 0; i < noteIds.length; i++) {
    const drivePdf = downloadDriveFilePdf(noteIds[i]);
    if (drivePdf) {
      rec.imported = true;
      rec.photo_count = -1;
      rec.drive_import_tried = true;
      return replacePdfBytes(rec, drivePdf, rec);
    }
  }
  const drivePdf = downloadDriveQcPdf(rec.order_number, rec.kind || rec.process);
  if (drivePdf) {
    rec.imported = true;
    rec.photo_count = -1;
    rec.drive_import_tried = true;
    return replacePdfBytes(rec, drivePdf, rec);
  }
  const store = loadStore();
  const idx = store.records.findIndex((row) => row && row.id === rec.id);
  if (idx >= 0) {
    store.records[idx].drive_import_tried = true;
    saveStore(store);
  }
  return false;
}

async function recoverMissingPhotos() {
  if (!hasGoogleAuth()) return { saved: 0 };
  const store = loadStore();
  let saved = 0;
  for (const rec of store.records) {
    if (!rec || !rec.id || !rec.pdf_path || !fs.existsSync(rec.pdf_path)) continue;
    if (!needsPhotoRecover(rec)) continue;
    const ok = await recoverPhotosForRecord(rec);
    if (ok) saved += 1;
  }
  return { saved };
}

async function backfillFromLogs() {
  const book = getBook();
  const sheet = book.getSheetByName("Production_Log");
  if (!sheet || sheet.getLastRow() < 2) {
    const recovered = await recoverMissingPhotos();
    return { saved: recovered.saved || 0 };
  }
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  let saved = 0;
  for (let i = 0; i < values.length; i++) {
    const role = String(values[i][3] || "").trim();
    const process = String(values[i][4] || "").trim();
    const ended = values[i][6];
    const notes = String(values[i][7] || "");
    const signature = String(values[i][8] || "");
    const blob = (role + " " + process).toLowerCase();
    if (blob.indexOf("powder coating") !== -1 && blob.indexOf("pre") === -1) continue;
    if (blob.indexOf("out for delivery") !== -1) continue;
    const isQc = role === "Quality Control" || blob.indexOf("qc") !== -1 || blob.indexOf("pre-powder") !== -1 || blob.indexOf("pre powder") !== -1;
    if (!isQc) continue;
    if (!ended) continue;
    const rowToUpdate = i + 2;
    const logId = String(values[i][0] || "");
    const existing = existingForLog(logId, rowToUpdate);
    if (existing) {
      if (!notesHaveLocalPdf(notes)) writeUrlOnLog(rowToUpdate, pdfUrlFor(existing.id));
      const recovered = await recoverPhotosForRecord(existing, signature);
      if (recovered || !notesHaveLocalPdf(notes)) saved += 1;
      continue;
    }
    if (notesHaveLocalPdf(notes)) continue;
    const answers = parseAnswersFromNotes(notes);
    if (!answers.length && !signature) continue;
    const orderNum = String(values[i][1] || "").trim();
    const workerName = String(values[i][2] || "").trim();
    try {
      const queued = queuePhotosForOrder(orderNum);
      const made = await saveFromFinish({
        logId,
        rowToUpdate,
        orderNum,
        workerName,
        processName: process,
        qcData: answers,
        signatureUrl: signature,
        filesData: queued,
        created_at: ended
      });
      if (made && made.photo_count === 0) {
        const rec = loadStore().records.find((row) => row && row.id === made.id);
        if (rec) await recoverPhotosForRecord(rec, signature);
      }
      if (made && made.url) {
        writeUrlOnLog(rowToUpdate, made.url);
        saved += 1;
      }
    } catch (e) {
      console.error("[qc-pdf] backfill", orderNum, e.message || e);
    }
  }
  const recovered = await recoverMissingPhotos();
  const linked = attachDriveUrlsFromLogs();
  return { saved: saved + (recovered.saved || 0) + (linked.saved || 0) };
}

module.exports = {
  saveFromFinish,
  writeUrlOnLog,
  listReports,
  readPdf,
  pdfsForOrder,
  pdfUrlFor,
  qcLabel,
  backfillFromLogs,
  parseAnswersFromNotes,
  recoverMissingPhotos,
  rebuildRecordFromPhotos,
  parseDriveFileId,
  driveIdsFromNotes,
  attachDriveUrlsFromLogs,
  toJpeg,
  packPhotos
};
