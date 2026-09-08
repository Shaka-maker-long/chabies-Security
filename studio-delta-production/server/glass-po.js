"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir, getBook, persistWorkbook } = require("./workbook-store");
const { parseMoney, money, formatRand } = require("./db");
const catalog = require("./product-catalog");
const glassRates = require("./glass-rates");

const TO_ORDER = "To order";
const ON_PO = "On PO";
const RECEIVED = "Received";
const MAX_INVOICE_BYTES = 15 * 1024 * 1024;
const INVOICE_TYPES = {
  "application/pdf": ".pdf",
  "image/jpeg": ".jpg",
  "image/jpg": ".jpg",
  "image/png": ".png",
  "image/webp": ".webp"
};
const HEADERS = ["ID", "Timestamp", "Order #", "Worker", "Component", "Glass type", "Thickness", "Height", "Width", "Quantity", "Status"];

function nowIso() {
  return new Date().toISOString();
}

function storePath() {
  return path.join(dataDir(), "glass-pos.json");
}

function invoicesDir() {
  const dir = path.join(dataDir(), "glass-po-invoices");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function emptyStore() {
  return { nextPo: 1, pos: [], receives: [], invoices: [], lines: {} };
}

function loadStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(storePath(), "utf8"));
    const store = emptyStore();
    store.nextPo = Number(parsed.nextPo) > 0 ? Number(parsed.nextPo) : 1;
    store.pos = Array.isArray(parsed.pos) ? parsed.pos : [];
    store.receives = Array.isArray(parsed.receives) ? parsed.receives : [];
    store.invoices = Array.isArray(parsed.invoices) ? parsed.invoices : [];
    store.lines = parsed.lines && typeof parsed.lines === "object" ? parsed.lines : {};
    return store;
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[glass-po] could not read", storePath(), e.message || e);
    }
    return emptyStore();
  }
}

function saveStore(store) {
  const file = storePath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(store));
  fs.renameSync(tmp, file);
  return store;
}

function newId(prefix) {
  return prefix + "_" + crypto.randomBytes(8).toString("hex");
}

function poNumber(n) {
  return "GPO-" + String(n).padStart(4, "0");
}

function statusKey(status) {
  return String(status || "").trim().toLowerCase();
}

function isToOrder(status) {
  const s = statusKey(status);
  return !s || s === "to order";
}

function isOutstanding(status) {
  const s = statusKey(status);
  return s === "on po" || s === "ordered";
}

function isReceived(status) {
  return statusKey(status) === "received";
}

function glassSheet() {
  const book = getBook();
  let sheet = book.getSheetByName("Glass_To_Order");
  if (!sheet) {
    sheet = book.insertSheet("Glass_To_Order");
    sheet.appendRow(HEADERS.slice());
    persistWorkbook();
  }
  return sheet;
}

function readGlassLines() {
  const sheet = glassSheet();
  const items = [];
  if (sheet.getLastRow() < 2) return items;
  const grid = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
  for (let i = 0; i < grid.length; i++) {
    const id = String(grid[i][0] || "").trim();
    const typeName = String(grid[i][5] || "").trim();
    if (!id && !typeName) continue;
    items.push({
      id,
      timestamp: grid[i][1] ? new Date(grid[i][1]).toISOString() : "",
      order: String(grid[i][2] || ""),
      worker: String(grid[i][3] || ""),
      component: String(grid[i][4] || ""),
      type: typeName,
      thickness: String(grid[i][6] || ""),
      height: Number(grid[i][7]) || 0,
      width: Number(grid[i][8]) || 0,
      quantity: Number(grid[i][9]) || 0,
      status: String(grid[i][10] || TO_ORDER).trim() || TO_ORDER,
      kind: "glass",
      row: i + 2
    });
  }
  return items;
}

function setGlassStatus(id, status) {
  const want = String(id || "").trim();
  const sheet = glassSheet();
  if (sheet.getLastRow() < 2) throw new Error("Glass line not found.");
  const grid = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < grid.length; i++) {
    if (String(grid[i][0] || "").trim() === want) {
      sheet.getRange(i + 2, 11).setValue(status);
      persistWorkbook();
      try { require("./gas").clearShopCache(); } catch (e) {}
      return true;
    }
  }
  throw new Error("Glass line " + want + " was not found.");
}

function decorateLine(line, extra) {
  const costs = glassRates.costLine(line);
  return Object.assign({
    id: line.id,
    order: line.order,
    worker: line.worker,
    component: line.component,
    type: line.type,
    thickness: line.thickness,
    height: line.height,
    width: line.width,
    quantity: line.quantity,
    status: line.status,
    timestamp: line.timestamp,
    kind: "glass"
  }, costs, extra || {});
}

function safeFilename(name, fallback) {
  const cleaned = String(name || "").replace(/[^A-Za-z0-9._-]+/g, "_");
  return cleaned || fallback || "invoice.pdf";
}

function decodeInvoice(invoice) {
  if (!invoice || typeof invoice !== "object") {
    throw new Error("Attach the glass invoice for this received batch.");
  }
  const raw = String(invoice.data || invoice.dataUrl || invoice.base64 || "");
  if (!raw.trim()) throw new Error("Attach the glass invoice for this received batch.");
  const match = raw.match(/^data:([^;]+);base64,(.+)$/i);
  const mime = String(invoice.mime || invoice.type || (match && match[1]) || "application/pdf")
    .split(";")[0]
    .trim()
    .toLowerCase();
  const b64 = match ? match[2] : raw.replace(/\s+/g, "");
  if (!INVOICE_TYPES[mime]) {
    throw new Error("Invoice must be a PDF or image (JPG, PNG, or WebP).");
  }
  const buffer = Buffer.from(b64, "base64");
  if (!buffer.length) throw new Error("Attach the glass invoice for this received batch.");
  if (buffer.length > MAX_INVOICE_BYTES) throw new Error("Invoice must be 15 MB or smaller.");
  let filename = safeFilename(invoice.filename || invoice.name, "invoice" + INVOICE_TYPES[mime]);
  if (!path.extname(filename)) filename += INVOICE_TYPES[mime];
  return { buffer, mime, filename };
}

function writeInvoiceFile(invoiceId, invoice) {
  const decoded = decodeInvoice(invoice);
  const dir = path.join(invoicesDir(), invoiceId);
  fs.mkdirSync(dir, { recursive: true });
  const storedAs = "invoice" + (path.extname(decoded.filename) || INVOICE_TYPES[decoded.mime] || ".pdf");
  fs.writeFileSync(path.join(dir, storedAs), decoded.buffer);
  return {
    id: invoiceId,
    filename: decoded.filename,
    storedAs,
    mime: decoded.mime,
    size: decoded.buffer.length
  };
}

function readInvoiceFile(invoiceId) {
  const store = loadStore();
  const rec = (store.invoices || []).find((row) => row && row.id === invoiceId);
  if (!rec) return null;
  const file = path.join(invoicesDir(), invoiceId, rec.storedAs || "invoice.pdf");
  if (!fs.existsSync(file)) return null;
  return {
    buffer: fs.readFileSync(file),
    filename: rec.filename || rec.storedAs || "invoice.pdf",
    mime: rec.mime || "application/octet-stream"
  };
}

function snapshot() {
  const store = loadStore();
  const lines = readGlassLines();
  const byId = {};
  lines.forEach((line) => { byId[line.id] = line; });
  const toOrder = lines.filter((line) => isToOrder(line.status)).map((line) => decorateLine(line));
  const outstanding = lines.filter((line) => isOutstanding(line.status)).map((line) => {
    const meta = store.lines[line.id] || {};
    return decorateLine(line, {
      poId: meta.poId || "",
      poNumber: meta.poNumber || "",
      poAt: meta.poAt || "",
      poBy: meta.poBy || ""
    });
  });
  const received = (store.receives || []).slice().reverse().map((batch) => {
    const invoice = (store.invoices || []).find((row) => row.id === batch.invoiceId) || {};
    const po = (store.pos || []).find((row) => row.id === batch.poId) || {};
    return {
      id: batch.id,
      receivedAt: batch.receivedAt,
      receivedBy: batch.receivedBy,
      poId: batch.poId || "",
      poNumber: po.number || batch.poNumber || "",
      invoiceId: batch.invoiceId,
      invoiceFilename: invoice.filename || "",
      lines: (batch.lines || []).map((entry) => {
        const live = byId[entry.id] || { id: entry.id };
        return decorateLine(live, {
          cost: money(entry.cost),
          costLabel: formatRand(entry.cost),
          receiveId: batch.id,
          invoiceId: batch.invoiceId,
          poNumber: po.number || batch.poNumber || ""
        });
      }),
      total: money((batch.lines || []).reduce((sum, entry) => sum + parseMoney(entry.cost), 0)),
      totalLabel: formatRand((batch.lines || []).reduce((sum, entry) => sum + parseMoney(entry.cost), 0))
    };
  });
  const pos = (store.pos || []).slice().reverse().map((po) => {
    const poLines = (po.lineIds || []).map((id) => {
      const live = byId[id];
      const meta = store.lines[id] || {};
      return live ? decorateLine(live, { poId: po.id, poNumber: po.number }) : decorateLine({
        id,
        order: "",
        type: "",
        thickness: "",
        height: 0,
        width: 0,
        quantity: 0,
        status: ""
      }, { poId: po.id, poNumber: po.number, ...meta });
    });
    const estimatedTotal = poLines.reduce((sum, line) => sum + parseMoney(line.estimatedCost), 0);
    return {
      id: po.id,
      number: po.number,
      createdAt: po.createdAt,
      createdBy: po.createdBy,
      lineIds: po.lineIds || [],
      lines: poLines,
      estimatedTotal: money(estimatedTotal),
      estimatedTotalLabel: formatRand(estimatedTotal),
      pdfUrl: "/api/office/glass-po/" + encodeURIComponent(po.id) + "/pdf"
    };
  });
  return {
    toOrder,
    outstanding,
    received,
    pos,
    toOrderCount: toOrder.length,
    outstandingCount: outstanding.length,
    glass: lines.map((line) => {
      const meta = store.lines[line.id] || {};
      return decorateLine(line, {
        poId: meta.poId || "",
        poNumber: meta.poNumber || "",
        cost: meta.cost || "",
        invoiceId: meta.invoiceId || ""
      });
    })
  };
}

function createPurchaseOrder(lineIds, actor) {
  const ids = (Array.isArray(lineIds) ? lineIds : [])
    .map((id) => String(id || "").trim())
    .filter(Boolean);
  const unique = [];
  ids.forEach((id) => { if (unique.indexOf(id) === -1) unique.push(id); });
  if (!unique.length) throw new Error("Select at least one glass line to put on a purchase order.");
  const lines = readGlassLines();
  const byId = {};
  lines.forEach((line) => { byId[line.id] = line; });
  unique.forEach((id) => {
    const line = byId[id];
    if (!line) throw new Error("Glass line was not found.");
    if (!isToOrder(line.status)) {
      throw new Error((line.order || id) + " is already " + line.status + ". Only To order glass can go on a new PO.");
    }
  });
  const store = loadStore();
  const createdAt = nowIso();
  const createdBy = String(actor || "Admin").trim() || "Admin";
  const number = poNumber(store.nextPo++);
  const poId = newId("gpo");
  unique.forEach((id) => {
    setGlassStatus(id, ON_PO);
    store.lines[id] = Object.assign({}, store.lines[id] || {}, {
      poId,
      poNumber: number,
      poAt: createdAt,
      poBy: createdBy,
      receiveId: "",
      receivedAt: "",
      cost: "",
      invoiceId: ""
    });
  });
  store.pos.push({ id: poId, number, createdAt, createdBy, lineIds: unique });
  saveStore(store);
  return { poId, poNumber: number, createdAt, createdBy, lineIds: unique, ...snapshot() };
}

function parseReceiveLines(lines) {
  const rows = Array.isArray(lines) ? lines : [];
  if (!rows.length) throw new Error("Select the glass that arrived.");
  const seen = {};
  return rows.map((row) => {
    const id = String((row && (row.id || row.lineId)) || "").trim();
    if (!id) throw new Error("Each received glass line needs an id.");
    if (seen[id]) throw new Error("A glass line is listed twice on this receive.");
    seen[id] = true;
    if (row == null || row.cost === "" || row.cost == null) {
      throw new Error("Put the glass cost on each received line.");
    }
    const cost = parseMoney(row.cost);
    if (!Number.isFinite(cost)) throw new Error("Put the glass cost on each received line.");
    if (cost < 0) throw new Error("Glass cost cannot be negative.");
    return { id, cost: Number(money(cost)) };
  });
}

function receiveGlass(payload, actor) {
  const body = payload || {};
  const entries = parseReceiveLines(body.lines || body.orders);
  const live = readGlassLines();
  const byId = {};
  live.forEach((line) => { byId[line.id] = line; });
  const found = entries.map((entry) => {
    const line = byId[entry.id];
    if (!line) throw new Error("Glass line was not found.");
    if (!isOutstanding(line.status)) {
      throw new Error((line.order || entry.id) + " is not outstanding on a PO. Generate a PO first, then receive it.");
    }
    return { entry, line };
  });
  decodeInvoice(body.invoice);
  const invoiceId = newId("ginv");
  const file = writeInvoiceFile(invoiceId, body.invoice);
  const store = loadStore();
  const receivedAt = nowIso();
  const receivedBy = String(actor || "Admin").trim() || "Admin";
  const receiveId = newId("grecv");
  const poIds = {};
  found.forEach(({ entry, line }) => {
    setGlassStatus(entry.id, RECEIVED);
    const prev = store.lines[entry.id] || {};
    if (prev.poId) poIds[prev.poId] = true;
    store.lines[entry.id] = Object.assign({}, prev, {
      receiveId,
      receivedAt,
      receivedBy,
      cost: money(entry.cost),
      invoiceId
    });
  });
  const poId = Object.keys(poIds)[0] || "";
  const po = (store.pos || []).find((row) => row.id === poId) || {};
  store.invoices.push({
    id: invoiceId,
    filename: file.filename,
    storedAs: file.storedAs,
    mime: file.mime,
    size: file.size,
    receiveId,
    receivedAt,
    receivedBy,
    poId,
    lineIds: entries.map((entry) => entry.id)
  });
  store.receives.push({
    id: receiveId,
    receivedAt,
    receivedBy,
    invoiceId,
    poId,
    poNumber: po.number || "",
    lines: entries
  });
  saveStore(store);
  return {
    receiveId,
    invoiceId,
    invoiceFilename: file.filename,
    receivedAt,
    receivedBy,
    lineIds: entries.map((entry) => entry.id),
    ...snapshot()
  };
}

function findPo(poId) {
  const store = loadStore();
  const want = String(poId || "").trim();
  return (store.pos || []).find((row) => row && (row.id === want || row.number === want)) || null;
}

function poLines(po) {
  const live = readGlassLines();
  const byId = {};
  live.forEach((line) => { byId[line.id] = line; });
  return (po.lineIds || []).map((id) => {
    const line = byId[id];
    if (!line) {
      return decorateLine({ id, order: "", type: "", thickness: "", height: 0, width: 0, quantity: 0, status: "" });
    }
    return decorateLine(line, { poId: po.id, poNumber: po.number });
  });
}

function formatPoDate(iso) {
  const d = iso ? new Date(iso) : new Date();
  if (isNaN(d.getTime())) return "";
  return d.toLocaleDateString("en-ZA", { timeZone: "Africa/Johannesburg", day: "2-digit", month: "short", year: "numeric" });
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

async function buildPurchaseOrderPdf(poId) {
  const po = findPo(poId);
  if (!po) throw new Error("Purchase order not found.");
  const lines = poLines(po);
  const PDFDocument = require("pdfkit");
  const logoPath = await logoPathForPdf();
  const chunks = [];
  const doc = new PDFDocument({
    size: "A4",
    margin: 36,
    compress: false,
    info: { Title: "Purchase Order " + po.number, Author: "Studio Delta" }
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
  doc.fillColor(ink).font("Helvetica-Bold").fontSize(20).text("PURCHASE ORDER", margin, margin + 8, { width: inner, align: "right" });
  doc.fillColor(brass).font("Helvetica-Bold").fontSize(12).text(po.number, margin, margin + 34, { width: inner, align: "right" });

  doc.save().strokeColor(brass).lineWidth(2).moveTo(margin, margin + 84).lineTo(margin + inner, margin + 84).stroke().restore();

  let y = margin + 96;
  drawPdfBox(doc, margin, y, inner * 0.48, 54);
  drawPdfBox(doc, margin + inner * 0.52, y, inner * 0.48, 54);
  doc.fillColor(muted).font("Helvetica-Bold").fontSize(8).text("SUPPLIER", margin + 8, y + 8);
  doc.fillColor(ink).font("Helvetica").fontSize(10).text("Glass supplier", margin + 8, y + 22);
  doc.fillColor(muted).font("Helvetica").fontSize(8).text("Name / company to be completed when sending", margin + 8, y + 36, { width: inner * 0.48 - 16 });
  doc.fillColor(muted).font("Helvetica-Bold").fontSize(8).text("ORDER DETAILS", margin + inner * 0.52 + 8, y + 8);
  doc.fillColor(ink).font("Helvetica").fontSize(10)
    .text("Date  " + formatPoDate(po.createdAt), margin + inner * 0.52 + 8, y + 22)
    .text("Prepared by  " + (po.createdBy || "Studio Delta"), margin + inner * 0.52 + 8, y + 36);

  y += 70;
  doc.fillColor(ink).font("Helvetica").fontSize(9).text(
    "Please supply the glass listed below. Sizes are millimetres. Estimated cost uses Studio Delta rates per square metre for that glass type and thickness.",
    margin, y, { width: inner }
  );
  y += 28;

  const headers = ["Order Number", "Glass type", "Thickness", "Height", "Width", "Quantity", "Area m²", "Est. cost"];
  const widths = [74, 74, 54, 46, 46, 50, 52, inner - 74 - 74 - 54 - 46 - 46 - 50 - 52];
  const headerH = 22;
  doc.save().fillColor("#1c1917").rect(margin, y, inner, headerH).fill().restore();
  let x = margin;
  headers.forEach((h, i) => {
    doc.fillColor("#fcfbf8").font("Helvetica-Bold").fontSize(7).text(h, x + 3, y + 7, { width: widths[i] - 6, lineBreak: false });
    x += widths[i];
  });
  y += headerH;

  let totalArea = 0;
  let totalCost = 0;
  let missing = 0;
  lines.forEach((line, idx) => {
    const rowH = 18;
    if (y + rowH > 760) {
      doc.addPage();
      y = margin;
    }
    if (idx % 2 === 1) {
      doc.save().fillColor("#f3efe6").rect(margin, y, inner, rowH).fill().restore();
    }
    doc.save().strokeColor("#d7d1c6").lineWidth(0.4).rect(margin, y, inner, rowH).stroke().restore();
    const area = Number(line.areaM2) || 0;
    totalArea += area;
    const cost = parseMoney(line.estimatedCost);
    if (line.rateMissing) missing += 1;
    else totalCost += cost;
    const cells = [
      line.order || "",
      line.type || "",
      line.thickness || "",
      line.height == null ? "" : String(line.height),
      line.width == null ? "" : String(line.width),
      line.quantity == null ? "" : String(line.quantity),
      area > 0 ? formatArea(area) : "—",
      line.estimatedCostLabel || "—"
    ];
    x = margin;
    cells.forEach((cell, i) => {
      doc.fillColor(ink).font("Helvetica").fontSize(8).text(String(cell), x + 3, y + 5, { width: widths[i] - 6, lineBreak: false });
      x += widths[i];
    });
    y += rowH;
  });

  y += 8;
  drawPdfBox(doc, margin + inner - 220, y, 220, 48);
  doc.fillColor(muted).font("Helvetica-Bold").fontSize(8).text("ESTIMATED TOTAL", margin + inner - 212, y + 8);
  doc.fillColor(ink).font("Helvetica-Bold").fontSize(14).text(formatRand(totalCost), margin + inner - 212, y + 22, { width: 204 });
  doc.fillColor(muted).font("Helvetica").fontSize(8).text(
    "Area " + formatArea(totalArea) + " m²" + (missing ? "  ·  " + missing + " line" + (missing === 1 ? "" : "s") + " missing a rate" : ""),
    margin, y + 8, { width: inner - 236 }
  );

  y += 64;
  doc.fillColor(muted).font("Helvetica").fontSize(8).text(
    "This estimated cost is for Studio Delta planning. The supplier invoice is the amount payable. Rates are set on Glass rates (R per m² by glass type and thickness).",
    margin, y, { width: inner }
  );
  y += 28;
  doc.fillColor(ink).font("Helvetica").fontSize(9)
    .text("Authorised ________________________________", margin, y)
    .text("Date ________________________________", margin + inner / 2, y);

  doc.end();
  const buffer = await done;
  return {
    buffer,
    filename: po.number + ".pdf",
    mime: "application/pdf",
    poNumber: po.number
  };
}

function formatArea(area) {
  if (!(Number(area) > 0)) return "0";
  return (Math.round(Number(area) * 10000) / 10000).toFixed(4).replace(/0+$/, "").replace(/\.$/, "");
}

module.exports = {
  TO_ORDER,
  ON_PO,
  RECEIVED,
  snapshot,
  createPurchaseOrder,
  receiveGlass,
  readInvoiceFile,
  readGlassLines,
  findPo,
  buildPurchaseOrderPdf
};
