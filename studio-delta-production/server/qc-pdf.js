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

function requiredPhotoCount(processName) {
  return photoLabels(processName).filter((label) => String(label).indexOf("Optional") === -1).length;
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

const catalog = require("./product-catalog");
const LAYOUT = 2;
const PAGE_W = 595.28;
const PAGE_H = 841.89;
const MARGIN = 36;
const INK = "#1c1917";
const BRASS = "#b08948";
const MUTED = "#6b645b";
const CREAM = "#fcfbf8";
const ALT = "#f3efe6";
const RULE = "#d7d1c6";

let logoMemo;

async function logoPathForPdf() {
  if (logoMemo !== undefined) return logoMemo;
  const url = catalog.COMPANY_LOGO_URL;
  const dest = path.join(dataDir(), "pdf-images", "studio-delta-logo.jpg");
  try { fs.mkdirSync(path.dirname(dest), { recursive: true }); } catch (e) {}
  if (fs.existsSync(dest) && fs.statSync(dest).size > 400) {
    logoMemo = dest;
    return dest;
  }
  if (!url) {
    logoMemo = null;
    return null;
  }
  try {
    const res = await fetch(url, { redirect: "follow", signal: AbortSignal.timeout(1500) });
    if (!res.ok) {
      logoMemo = fs.existsSync(dest) ? dest : null;
      return logoMemo;
    }
    const buf = Buffer.from(await res.arrayBuffer());
    if (buf.length) fs.writeFileSync(dest, buf);
    logoMemo = dest;
    return dest;
  } catch (e) {
    logoMemo = fs.existsSync(dest) ? dest : null;
    return logoMemo;
  }
}

function drawPdfBox(doc, x, y, w, h) {
  doc.save().lineWidth(0.7).strokeColor(INK).rect(x, y, w, h).stroke().restore();
}

function reportTitle(record) {
  const kind = String((record && record.kind) || (record && record.process) || "QC").toUpperCase();
  if (kind.indexOf("FINAL") !== -1) return "FINAL QC";
  if (kind.indexOf("PRE") !== -1) return "PRE-POWDER QC";
  return "QC REPORT";
}

function answerLabel(raw) {
  const a = String(raw || "").trim();
  if (a === "Y" || /^yes$/i.test(a)) return "Yes";
  if (a === "N" || /^no$/i.test(a)) return "No";
  return a || "—";
}

function drawReportHeader(doc, record, logoPath) {
  const inner = PAGE_W - MARGIN * 2;
  drawPdfBox(doc, MARGIN, MARGIN, 72, 72);
  if (logoPath) {
    try {
      doc.image(logoPath, MARGIN + 4, MARGIN + 4, { fit: [64, 64] });
    } catch (e) {
      doc.fillColor(INK).font("Helvetica-Bold").fontSize(9).text("STUDIO\nDELTA", MARGIN, MARGIN + 26, { width: 72, align: "center" });
    }
  } else {
    doc.fillColor(INK).font("Helvetica-Bold").fontSize(9).text("STUDIO\nDELTA", MARGIN, MARGIN + 26, { width: 72, align: "center" });
  }
  doc.fillColor(INK).font("Helvetica-Bold").fontSize(16).text("STUDIO DELTA", MARGIN + 88, MARGIN + 8);
  doc.fillColor(MUTED).font("Helvetica").fontSize(9).text("Furniture  ·  Steel  ·  Glass", MARGIN + 88, MARGIN + 28);
  doc.fillColor(MUTED).font("Helvetica").fontSize(8).text("studiodelta.co.za", MARGIN + 88, MARGIN + 42);
  doc.fillColor(INK).font("Helvetica-Bold").fontSize(18).text(reportTitle(record), MARGIN, MARGIN + 8, { width: inner, align: "right" });
  doc.fillColor(BRASS).font("Helvetica-Bold").fontSize(12).text(record.order_number || "", MARGIN, MARGIN + 34, { width: inner, align: "right" });
  doc.save().strokeColor(BRASS).lineWidth(2).moveTo(MARGIN, MARGIN + 84).lineTo(MARGIN + inner, MARGIN + 84).stroke().restore();
}

function drawPhotoPageHeader(doc, record, photoName, index, total) {
  const inner = PAGE_W - MARGIN * 2;
  doc.fillColor(INK).font("Helvetica-Bold").fontSize(11).text("STUDIO DELTA", MARGIN, MARGIN);
  doc.fillColor(MUTED).font("Helvetica").fontSize(8).text(
    (record.order_number || "") + "  ·  " + reportTitle(record),
    MARGIN, MARGIN + 16
  );
  doc.fillColor(INK).font("Helvetica-Bold").fontSize(12).text(photoName || "Photo", MARGIN, MARGIN, { width: inner, align: "right" });
  doc.fillColor(MUTED).font("Helvetica").fontSize(8).text(
    "Photo " + String(index) + " of " + String(total),
    MARGIN, MARGIN + 16, { width: inner, align: "right" }
  );
  doc.save().strokeColor(BRASS).lineWidth(2).moveTo(MARGIN, MARGIN + 36).lineTo(MARGIN + inner, MARGIN + 36).stroke().restore();
}

function loadStoredPhotos(id) {
  const dir = photosDir(id);
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir)
    .filter((name) => /\.jpe?g$/i.test(name) && !/^signature\./i.test(name))
    .sort()
    .map((name) => {
      const label = name.replace(/^\d+-/, "").replace(/\.jpe?g$/i, "").replace(/_/g, " ");
      return { name: label || "Photo", buf: fs.readFileSync(path.join(dir, name)) };
    });
}

function loadStoredSignature(id) {
  const file = path.join(photosDir(id), "signature.jpg");
  if (!fs.existsSync(file)) return null;
  const buf = fs.readFileSync(file);
  return buf && buf.length ? buf : null;
}

async function restyleRecord(rec) {
  if (!rec || rec.imported || !rec.pdf_path) return false;
  const photos = loadStoredPhotos(rec.id);
  const signatureBuf = loadStoredSignature(rec.id);
  await renderPdf({
    ...rec,
    photos,
    signatureBuf,
    photoAttempted: photos.length > 0
  }, rec.pdf_path);
  rec.layout = LAYOUT;
  return true;
}

async function restyleStoredReports() {
  const store = loadStore();
  let n = 0;
  for (const rec of store.records) {
    if (!rec || rec.imported || rec.layout === LAYOUT) continue;
    if (!rec.pdf_path || !fs.existsSync(rec.pdf_path)) continue;
    try {
      await restyleRecord(rec);
      n += 1;
    } catch (e) {
      console.error("[qc-pdf] restyle", rec.order_number, e && e.message ? e.message : e);
    }
  }
  if (n) saveStore(store);
  return n;
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
  const logoPath = await logoPathForPdf();
  const inner = PAGE_W - MARGIN * 2;
  const photos = (record.photos || []).filter((photo) => photo && photo.buf);
  const photoSizes = [];
  for (let i = 0; i < photos.length; i++) {
    try {
      const meta = await require("sharp")(photos[i].buf).metadata();
      photoSizes.push({ w: meta.width || 1, h: meta.height || 1 });
    } catch (e) {
      photoSizes.push({ w: 1, h: 1 });
    }
  }
  await new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: "A4",
      margin: MARGIN,
      compress: false,
      info: {
        Title: (record.order_number || "Order") + " " + (record.kind || "QC"),
        Author: "Studio Delta"
      }
    });
    const stream = fs.createWriteStream(dest);
    doc.pipe(stream);
    stream.on("finish", resolve);
    stream.on("error", reject);
    doc.on("error", reject);

    drawReportHeader(doc, record, logoPath);

    let y = MARGIN + 96;
    const boxW = inner * 0.48;
    drawPdfBox(doc, MARGIN, y, boxW, 54);
    drawPdfBox(doc, MARGIN + inner * 0.52, y, boxW, 54);
    doc.fillColor(MUTED).font("Helvetica-Bold").fontSize(8).text("ORDER", MARGIN + 8, y + 8);
    doc.fillColor(INK).font("Helvetica-Bold").fontSize(11).text(record.order_number || "—", MARGIN + 8, y + 22);
    doc.fillColor(MUTED).font("Helvetica").fontSize(8).text(record.product || "Studio Delta", MARGIN + 8, y + 38, { width: boxW - 16, lineBreak: false });
    doc.fillColor(MUTED).font("Helvetica-Bold").fontSize(8).text("COMPLETED", MARGIN + inner * 0.52 + 8, y + 8);
    doc.fillColor(INK).font("Helvetica").fontSize(10)
      .text(record.worker || "—", MARGIN + inner * 0.52 + 8, y + 22)
      .text(formatWhen(record.created_at), MARGIN + inner * 0.52 + 8, y + 36);

    y += 70;
    const answers = Array.isArray(record.answers) ? record.answers : [];
    const numW = 28;
    const resultW = 70;
    const qW = inner - numW - resultW;
    const headerH = 22;
    function drawChecklistHeader() {
      doc.save().fillColor(INK).rect(MARGIN, y, inner, headerH).fill().restore();
      doc.fillColor(CREAM).font("Helvetica-Bold").fontSize(7)
        .text("#", MARGIN + 4, y + 7, { width: numW - 6, lineBreak: false })
        .text("Checklist", MARGIN + numW + 4, y + 7, { width: qW - 8, lineBreak: false })
        .text("Result", MARGIN + numW + qW + 4, y + 7, { width: resultW - 8, lineBreak: false });
      y += headerH;
    }
    drawChecklistHeader();
    answers.forEach((row, i) => {
      const q = String(row.q || "Q" + (i + 1));
      const result = answerLabel(row.a);
      const qHeight = q.length > 72 ? 32 : 20;
      if (y + qHeight > PAGE_H - 120) {
        doc.addPage();
        drawReportHeader(doc, record, logoPath);
        y = MARGIN + 96;
        drawChecklistHeader();
      }
      if (i % 2 === 1) {
        doc.save().fillColor(ALT).rect(MARGIN, y, inner, qHeight).fill().restore();
      }
      doc.save().strokeColor(RULE).lineWidth(0.4).rect(MARGIN, y, inner, qHeight).stroke().restore();
      doc.fillColor(INK).font("Helvetica").fontSize(8)
        .text(String(i + 1), MARGIN + 4, y + 6, { width: numW - 6, lineBreak: false })
        .text(q, MARGIN + numW + 4, y + 6, { width: qW - 8 });
      doc.fillColor(INK).font("Helvetica-Bold").fontSize(8)
        .text(result, MARGIN + numW + qW + 4, y + 6, { width: resultW - 8, align: "center" });
      y += qHeight;
    });
    if (!answers.length) {
      doc.save().strokeColor(RULE).lineWidth(0.4).rect(MARGIN, y, inner, 20).stroke().restore();
      doc.fillColor(MUTED).font("Helvetica").fontSize(8).text("No checklist answers were saved.", MARGIN + 8, y + 6);
      y += 20;
    }

    y += 18;
    const sigH = 88;
    if (y + sigH > PAGE_H - MARGIN) {
      doc.addPage();
      drawReportHeader(doc, record, logoPath);
      y = MARGIN + 96;
    }
    drawPdfBox(doc, MARGIN, y, inner, sigH);
    doc.fillColor(MUTED).font("Helvetica-Bold").fontSize(8).text("SIGNATURE", MARGIN + 8, y + 8);
    if (record.signatureBuf) {
      try {
        doc.image(record.signatureBuf, MARGIN + 8, y + 22, { fit: [220, 56] });
      } catch (e) {
        doc.fillColor(MUTED).font("Helvetica").fontSize(8).text("Signature could not be placed.", MARGIN + 8, y + 40);
      }
    } else {
      doc.save().strokeColor(RULE).lineWidth(0.5).moveTo(MARGIN + 8, y + 70).lineTo(MARGIN + 240, y + 70).stroke().restore();
      doc.fillColor(MUTED).font("Helvetica").fontSize(8).text(record.worker || "Signed", MARGIN + 8, y + 74);
    }
    doc.fillColor(MUTED).font("Helvetica").fontSize(8).text(
      "Quality Control  ·  " + (record.order_number || ""),
      MARGIN, y + 36, { width: inner - 12, align: "right" }
    );

    const photosOnPage = photos;
    photosOnPage.forEach((photo, i) => {
      doc.addPage();
      drawPhotoPageHeader(doc, record, photo.name || ("Photo " + (i + 1)), i + 1, photos.length);
      const frameY = MARGIN + 48;
      const frameH = PAGE_H - frameY - MARGIN;
      drawPdfBox(doc, MARGIN, frameY, inner, frameH);
      const maxW = inner - 24;
      const maxH = frameH - 24;
      const size = photoSizes[i] || { w: maxW, h: maxH };
      const scale = Math.min(maxW / size.w, maxH / size.h);
      const dw = Math.max(1, size.w * scale);
      const dh = Math.max(1, size.h * scale);
      const imgX = MARGIN + (inner - dw) / 2;
      const imgY = frameY + (frameH - dh) / 2;
      try {
        doc.image(photo.buf, imgX, imgY, { width: dw, height: dh });
      } catch (e) {
        doc.fillColor(MUTED).font("Helvetica").fontSize(10).text("This photo could not be placed.", MARGIN + 16, frameY + 24);
      }
    });

    if (record.photoAttempted && !photos.length) {
      doc.addPage();
      drawPhotoPageHeader(doc, record, "Photos", 0, 0);
      doc.fillColor(MUTED).font("Helvetica").fontSize(10).text(
        "Photos were sent with this QC but could not be placed on the PDF.",
        MARGIN, MARGIN + 52, { width: inner }
      );
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
    imported: false,
    layout: LAYOUT
  };
  const files = (job && job.filesData) || [];
  const photoAttempted = files.some((file) => !!file);
  const photos = await packPhotos(rec.process, files);
  const need = requiredPhotoCount(rec.process);
  if (!(job && job.allowMissingPhotos) && photos.length < need) {
    try { fs.rmSync(photosDir(id), { recursive: true, force: true }); } catch (e) {}
    throw new Error(
      "QC PDF needs " + need + " photos (Top and Level 2 can be skipped). Only " +
      photos.length + " could be placed. The job is still running."
    );
  }
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
  rec.layout = LAYOUT;
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

async function attachPhotos(job) {
  const order = String((job && job.orderNum) || "").trim();
  const logId = String((job && job.logId) || "");
  const rowToUpdate = Number(job && job.rowToUpdate) || 0;
  const files = (job && job.filesData) || [];
  const incoming = files.filter((file) => !!file);
  if (!incoming.length) throw new Error("Add the QC photos from the tablet gallery.");
  let rec = existingForLog(logId, rowToUpdate);
  if (!rec && order) {
    const wantKind = qcKind(job && job.processName);
    rec = loadStore().records.find((row) => {
      if (!row || String(row.order_number || "").trim().toLowerCase() !== order.toLowerCase()) return false;
      return qcKind(row.process || row.kind) === wantKind;
    }) || null;
  }
  const answers = (rec && rec.answers && rec.answers.length)
    ? rec.answers
    : parseAnswersFromNotes(job && job.notes);
  const processName = (rec && rec.process) || String((job && job.processName) || "Final QC");
  if (rec) {
    rec.answers = answers;
    rec.process = processName;
    rec.kind = qcKind(processName);
    const ok = await rebuildRecordFromPhotos(rec, files, job && job.signatureUrl);
    if (!ok) throw new Error("Could not place the photos on the PDF.");
    return {
      id: rec.id,
      url: pdfUrlFor(rec.id),
      photo_count: rec.photo_count,
      name: displayName(rec),
      label: qcLabel(rec.process)
    };
  }
  return saveFromFinish({
    logId,
    rowToUpdate,
    orderNum: order,
    workerName: (job && job.workerName) || "",
    processName,
    productName: (job && job.productName) || "",
    qcData: answers,
    signatureUrl: (job && job.signatureUrl) || "",
    filesData: files,
    created_at: job && job.created_at,
    allowMissingPhotos: true
  });
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
      if (!queued.length) continue;
      const made = await saveFromFinish({
        logId,
        rowToUpdate,
        orderNum,
        workerName,
        processName: process,
        qcData: answers,
        signatureUrl: signature,
        filesData: queued,
        created_at: ended,
        allowMissingPhotos: true
      });
      if (made && made.photo_count === 0) {
        const rec = loadStore().records.find((row) => row && row.id === made.id);
        if (rec) await recoverPhotosForRecord(rec, signature);
      }
      if (made && made.photo_count > 0 && made.url) {
        writeUrlOnLog(rowToUpdate, made.url);
        saved += 1;
      }
    } catch (e) {
      console.error("[qc-pdf] backfill", orderNum, e.message || e);
    }
  }
  const recovered = await recoverMissingPhotos();
  const linked = attachDriveUrlsFromLogs();
  const restyled = await restyleStoredReports();
  return { saved: saved + (recovered.saved || 0) + (linked.saved || 0) + restyled };
}

module.exports = {
  saveFromFinish,
  requiredPhotoCount,
  writeUrlOnLog,
  listReports,
  readPdf,
  pdfsForOrder,
  pdfUrlFor,
  qcLabel,
  backfillFromLogs,
  parseAnswersFromNotes,
  recoverMissingPhotos,
  attachPhotos,
  rebuildRecordFromPhotos,
  parseDriveFileId,
  driveIdsFromNotes,
  attachDriveUrlsFromLogs,
  toJpeg,
  packPhotos
};
