"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir, getBook, persistWorkbook } = require("./workbook-store");

const PRE_PHOTOS = [
  "Front", "Left Side", "Right Side", "Back", "Open", "Top (Optional)", "Level 1", "Level 2 (Optional)"
];
const FINAL_PHOTOS = [
  "Front", "Level 1", "Back", "Left Side", "Right Side", "Job Card", "Open", "Top (Optional)", "Level 2 (Optional)"
];

function storePath() {
  return path.join(dataDir(), "qc-pdfs.json");
}

function pdfDir() {
  const dir = path.join(dataDir(), "qc-pdfs");
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

function decodeImage(raw) {
  const s = String(raw || "").trim();
  if (!s) return null;
  const m = s.match(/^data:([^;]+);base64,(.+)$/);
  if (m) return { mime: m[1], buf: Buffer.from(m[2], "base64") };
  try {
    return { mime: "image/jpeg", buf: Buffer.from(s, "base64") };
  } catch (e) {
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

    const sig = decodeImage(record.signature);
    if (sig) {
      doc.addPage();
      doc.font("Helvetica-Bold").fontSize(12).text("Signature");
      doc.moveDown(0.4);
      try {
        doc.image(sig.buf, { fit: [240, 120] });
      } catch (e) {
        doc.font("Helvetica").fontSize(10).text("Signature could not be placed.");
      }
    }

    (record.photos || []).forEach((photo) => {
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

    doc.end();
  });
}

function packPhotos(processName, filesData) {
  const labels = photoLabels(processName);
  const out = [];
  (filesData || []).forEach((file, i) => {
    if (!file) return;
    const img = decodeImage(file.data);
    if (!img) return;
    out.push({ name: labels[i] || file.name || ("Photo " + (i + 1)), buf: img.buf });
  });
  return out;
}

function displayName(rec) {
  return (rec.order_number || "Order") + " · " + (rec.kind || "QC") + (rec.worker ? " · " + rec.worker : "");
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
    signature: String((job && job.signatureUrl) || ""),
    created_at: new Date().toISOString(),
    pdf_path: pdfPath,
    log_id: String((job && job.logId) || ""),
    row: Number(job && job.rowToUpdate) || 0
  };
  const photos = packPhotos(rec.process, job && job.filesData);
  await renderPdf({ ...rec, photos }, pdfPath);
  const store = loadStore();
  const slim = { ...rec };
  delete slim.signature;
  store.records.unshift(slim);
  saveStore(store);
  return {
    id,
    url: pdfUrlFor(id),
    name: displayName(rec),
    label: qcLabel(rec.process),
    order_number: order,
    dateCreated: Date.parse(rec.created_at) || Date.now()
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
      url: pdfUrlFor(rec.id),
      dateCreated: Date.parse(rec.created_at) || 0,
      order_number: rec.order_number || "",
      label: qcLabel(rec.process),
      worker: rec.worker || "",
      kind: rec.kind || "",
      log_id: rec.log_id || "",
      process: rec.process || ""
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

async function backfillFromLogs() {
  const book = getBook();
  const sheet = book.getSheetByName("Production_Log");
  if (!sheet || sheet.getLastRow() < 2) return { saved: 0 };
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
    if (notesHaveLocalPdf(notes)) continue;
    const rowToUpdate = i + 2;
    const logId = String(values[i][0] || "");
    const existing = existingForLog(logId, rowToUpdate);
    if (existing) {
      writeUrlOnLog(rowToUpdate, pdfUrlFor(existing.id));
      saved += 1;
      continue;
    }
    const answers = parseAnswersFromNotes(notes);
    if (!answers.length && !signature) continue;
    const orderNum = String(values[i][1] || "").trim();
    const workerName = String(values[i][2] || "").trim();
    try {
      const made = await saveFromFinish({
        logId,
        rowToUpdate,
        orderNum,
        workerName,
        processName: process,
        qcData: answers,
        signatureUrl: signature,
        filesData: []
      });
      if (made && made.url) {
        writeUrlOnLog(rowToUpdate, made.url);
        saved += 1;
      }
    } catch (e) {
      console.error("[qc-pdf] backfill", orderNum, e.message || e);
    }
  }
  return { saved };
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
  parseAnswersFromNotes
};
