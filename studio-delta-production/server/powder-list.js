"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir } = require("./workbook-store");
const catalog = require("./product-catalog");

function storePath() {
  return path.join(dataDir(), "powder-lists.json");
}

function pdfDir() {
  const dir = path.join(dataDir(), "powder-lists");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function loadStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(storePath(), "utf8"));
    if (parsed && Array.isArray(parsed.records)) return parsed;
  } catch (e) {
    if (e && e.code !== "ENOENT") console.error("[powder-list] read", e.message || e);
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
  return "/api/powder-lists/" + encodeURIComponent(id) + "/pdf";
}

function formatPoDate(iso) {
  const d = iso ? new Date(iso) : new Date();
  if (isNaN(d.getTime())) return "";
  return d.toLocaleDateString("en-ZA", {
    timeZone: "Africa/Johannesburg",
    day: "2-digit",
    month: "short",
    year: "numeric"
  });
}

function nextNumber(records, now) {
  const d = now ? new Date(now) : new Date();
  const ymd = isNaN(d.getTime())
    ? "00000000"
    : d.toLocaleDateString("en-CA", { timeZone: "Africa/Johannesburg" }).replace(/-/g, "");
  const prefix = "PCL-" + ymd + "-";
  let max = 0;
  (records || []).forEach((row) => {
    const n = String((row && row.number) || "");
    if (n.indexOf(prefix) !== 0) return;
    const seq = Number(n.slice(prefix.length));
    if (seq > max) max = seq;
  });
  return prefix + String(max + 1);
}

function uniqueOrderNumbers(lines) {
  const out = [];
  (lines || []).forEach((line) => {
    const n = String((line && line.order) || "").trim();
    if (!n || out.indexOf(n) !== -1) return;
    out.push(n);
  });
  return out;
}

function normalizeLines(listData) {
  return (Array.isArray(listData) ? listData : []).map((item) => ({
    order: String((item && item.order) || "").trim(),
    desc: String((item && item.desc) || "").trim(),
    dimensions: String((item && item.dimensions) || "").trim(),
    qty: String((item && item.qty) || "").trim(),
    color: String((item && (item.color || item.colour)) || "").trim()
  })).filter((row) => row.desc || row.order || row.dimensions || row.qty || row.color);
}

async function logoPathForPdf() {
  const url = catalog.COMPANY_LOGO_URL;
  if (!url) return null;
  const dir = path.join(dataDir(), "pdf-images");
  fs.mkdirSync(dir, { recursive: true });
  const dest = path.join(dir, "studio-delta-logo.jpg");
  if (fs.existsSync(dest) && fs.statSync(dest).size > 400) return dest;
  try {
    const res = await fetch(url, { redirect: "follow", signal: AbortSignal.timeout(1500) });
    if (!res.ok) return fs.existsSync(dest) ? dest : null;
    const buf = Buffer.from(await res.arrayBuffer());
    if (buf.length) fs.writeFileSync(dest, buf);
    return dest;
  } catch (e) {
    return fs.existsSync(dest) ? dest : null;
  }
}

function drawPdfBox(doc, x, y, w, h) {
  doc.save().lineWidth(0.7).strokeColor("#1c1917").rect(x, y, w, h).stroke().restore();
}

async function buildPdf(rec) {
  const PDFDocument = require("pdfkit");
  const logoPath = await logoPathForPdf();
  const chunks = [];
  const doc = new PDFDocument({
    size: "A4",
    margin: 36,
    compress: false,
    info: { Title: "Powder Coating List " + rec.number, Author: "Studio Delta" }
  });
  doc.on("data", (c) => chunks.push(c));
  const done = new Promise((resolve, reject) => {
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
  });

  const pageW = 595.28;
  const margin = 36;
  const inner = pageW - margin * 2;
  const ink = "#1c1917";
  const brass = "#b08948";
  const muted = "#6b645b";

  drawPdfBox(doc, margin, margin, 72, 72);
  if (logoPath) {
    try { doc.image(logoPath, margin + 4, margin + 4, { fit: [64, 64] }); } catch (e) {
      doc.fillColor(ink).font("Helvetica-Bold").fontSize(9).text("STUDIO\nDELTA", margin, margin + 26, { width: 72, align: "center" });
    }
  } else {
    doc.fillColor(ink).font("Helvetica-Bold").fontSize(9).text("STUDIO\nDELTA", margin, margin + 26, { width: 72, align: "center" });
  }

  doc.fillColor(ink).font("Helvetica-Bold").fontSize(16).text("STUDIO DELTA", margin + 88, margin + 8);
  doc.fillColor(muted).font("Helvetica").fontSize(9).text("Furniture  ·  Steel  ·  Glass", margin + 88, margin + 28);
  doc.fillColor(muted).font("Helvetica").fontSize(8).text("studiodelta.co.za", margin + 88, margin + 42);
  doc.fillColor(ink).font("Helvetica-Bold").fontSize(16).text("PURCHASE ORDER", margin, margin + 8, { width: inner, align: "right" });
  doc.fillColor(brass).font("Helvetica-Bold").fontSize(11).text(rec.number, margin, margin + 30, { width: inner, align: "right" });
  doc.fillColor(muted).font("Helvetica").fontSize(9).text("Powder coating list", margin, margin + 46, { width: inner, align: "right" });

  doc.save().strokeColor(brass).lineWidth(2).moveTo(margin, margin + 84).lineTo(margin + inner, margin + 84).stroke().restore();

  let y = margin + 96;
  drawPdfBox(doc, margin, y, inner * 0.48, 54);
  drawPdfBox(doc, margin + inner * 0.52, y, inner * 0.48, 54);
  doc.fillColor(muted).font("Helvetica-Bold").fontSize(8).text("SUPPLIER", margin + 8, y + 8);
  doc.fillColor(ink).font("Helvetica").fontSize(10).text("Powder coaters", margin + 8, y + 22);
  doc.fillColor(muted).font("Helvetica").fontSize(8).text("Name / company to be completed when sending", margin + 8, y + 36, { width: inner * 0.48 - 16 });
  doc.fillColor(muted).font("Helvetica-Bold").fontSize(8).text("ORDER DETAILS", margin + inner * 0.52 + 8, y + 8);
  doc.fillColor(ink).font("Helvetica").fontSize(10)
    .text("Date  " + formatPoDate(rec.createdAt), margin + inner * 0.52 + 8, y + 22)
    .text("Prepared by  " + (rec.createdBy || "Studio Delta"), margin + inner * 0.52 + 8, y + 36);

  y += 70;
  doc.fillColor(ink).font("Helvetica").fontSize(9).text(
    "Please powder coat the items listed below. Sizes are millimetres (H x W x D).",
    margin, y, { width: inner }
  );
  y += 24;

  const headers = ["Order Number", "Item Description", "Dimensions (H x W x D)", "QTY", "Colour"];
  const widths = [90, 170, 130, 40, inner - 90 - 170 - 130 - 40];
  const headerH = 22;
  doc.save().fillColor("#1c1917").rect(margin, y, inner, headerH).fill().restore();
  let x = margin;
  headers.forEach((h, i) => {
    doc.fillColor("#fcfbf8").font("Helvetica-Bold").fontSize(7).text(h, x + 3, y + 7, { width: widths[i] - 6, lineBreak: false });
    x += widths[i];
  });
  y += headerH;

  (rec.lines || []).forEach((line, idx) => {
    const desc = String(line.desc || "");
    const descH = Math.max(18, doc.heightOfString(desc, { width: widths[1] - 6, fontSize: 8 }) + 8);
    const rowH = Math.max(20, descH);
    if (y + rowH > 760) {
      doc.addPage();
      y = margin;
    }
    if (idx % 2 === 1) {
      doc.save().fillColor("#f6f3ee").rect(margin, y, inner, rowH).fill().restore();
    }
    const cells = [
      line.order || "",
      desc,
      String(line.dimensions || ""),
      String(line.qty || ""),
      String(line.color || "")
    ];
    x = margin;
    cells.forEach((cell, i) => {
      doc.fillColor(ink).font(i === 0 ? "Helvetica-Bold" : "Helvetica").fontSize(8)
        .text(cell, x + 3, y + 5, { width: widths[i] - 6 });
      x += widths[i];
    });
    y += rowH;
  });

  doc.end();
  return done;
}

async function createList(listData, workerName) {
  const lines = normalizeLines(listData);
  if (!lines.length) {
    return { success: false, error: "Select at least one order to send." };
  }
  for (const line of lines) {
    if (!line.desc || !line.dimensions || !line.qty || !line.color) {
      return { success: false, error: "Fill in item, dimensions, quantity, and colour for every line." };
    }
  }
  const orderNumbers = uniqueOrderNumbers(lines);
  if (!orderNumbers.length) {
    return { success: false, error: "Select at least one order to send." };
  }
  const paint = require("./powder-shop");
  try {
    paint.assertReadyForPowderList(orderNumbers);
  } catch (e) {
    return { success: false, error: (e && e.message) || String(e) };
  }
  const store = loadStore();
  const id = newId();
  const createdAt = new Date().toISOString();
  const rec = {
    id,
    number: nextNumber(store.records, createdAt),
    createdAt,
    createdBy: String(workerName || "").trim() || "Studio Delta",
    lines
  };
  const buffer = await buildPdf(rec);
  fs.writeFileSync(path.join(pdfDir(), id + ".pdf"), buffer);
  store.records.unshift({
    id: rec.id,
    number: rec.number,
    createdAt: rec.createdAt,
    createdBy: rec.createdBy,
    lineCount: rec.lines.length
  });
  saveStore(store);
  let moved = [];
  try {
    const sent = paint.sendToPaintShop(orderNumbers, rec.createdBy, { status: paint.AT_SHOP_STATUS });
    moved = sent.orderNumbers || orderNumbers;
  } catch (e) {
    return { success: false, error: (e && e.message) || String(e), id, url: pdfUrlFor(id), number: rec.number };
  }
  return {
    success: true,
    id,
    number: rec.number,
    url: pdfUrlFor(id),
    filename: rec.number + ".pdf",
    orderNumbers: moved,
    status: paint.AT_SHOP_STATUS
  };
}

function readPdf(id) {
  const safe = String(id || "").replace(/[^a-zA-Z0-9_-]/g, "");
  if (!safe) return null;
  const file = path.join(pdfDir(), safe + ".pdf");
  if (!fs.existsSync(file)) return null;
  const rec = loadStore().records.find((row) => row && row.id === safe) || {};
  return {
    buffer: fs.readFileSync(file),
    filename: String(rec.number || safe) + ".pdf"
  };
}

module.exports = {
  createList,
  readPdf,
  pdfUrlFor,
  normalizeLines
};
