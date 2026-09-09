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
const HEADERS = ["ID", "Timestamp", "Order #", "Worker", "Component", "Glass type", "Thickness", "Height", "Width", "Quantity", "Status", "Template", "Template spec"];

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

function isTemplateFlag(value) {
  if (value === true || value === 1) return true;
  const s = String(value || "").trim().toLowerCase();
  return s === "yes" || s === "true" || s === "1" || s === "template";
}

function dimensionsLabel(line) {
  if (line && line.isTemplate) {
    const spec = String(line.templateSpec || "").trim();
    return spec ? "Template · " + spec : "Template glass";
  }
  const h = line && line.height;
  const w = line && line.width;
  if (h && w) return h + " × " + w + " mm";
  return "—";
}

function glassSheet() {
  const book = getBook();
  let sheet = book.getSheetByName("Glass_To_Order");
  if (!sheet) {
    sheet = book.insertSheet("Glass_To_Order");
    sheet.appendRow(HEADERS.slice());
    persistWorkbook();
  } else if (sheet.getLastColumn() < HEADERS.length) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS.slice()]);
    persistWorkbook();
  }
  return sheet;
}

function readGlassLines() {
  const sheet = glassSheet();
  const items = [];
  if (sheet.getLastRow() < 2) return items;
  const lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);
  const grid = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
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
      isTemplate: isTemplateFlag(grid[i][11]),
      templateSpec: String(grid[i][12] || "").trim(),
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

function applyLinePatch(id, patch) {
  const want = String(id || "").trim();
  const sheet = glassSheet();
  if (sheet.getLastRow() < 2) throw new Error("Glass line not found.");
  const lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);
  const grid = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  for (let i = 0; i < grid.length; i++) {
    if (String(grid[i][0] || "").trim() !== want) continue;
    const row = i + 2;
    const body = patch && typeof patch === "object" ? patch : {};
    if (body.type != null) sheet.getRange(row, 6).setValue(String(body.type || "").trim());
    if (body.thickness != null) sheet.getRange(row, 7).setValue(String(body.thickness || "").trim());
    if (body.height != null) sheet.getRange(row, 8).setValue(Number(body.height) || 0);
    if (body.width != null) sheet.getRange(row, 9).setValue(Number(body.width) || 0);
    if (body.quantity != null) {
      const qty = Math.round(Number(body.quantity));
      if (!(qty > 0)) throw new Error("Quantity must be more than 0.");
      sheet.getRange(row, 10).setValue(qty);
    }
    if (body.isTemplate != null || body.template != null) {
      sheet.getRange(row, 12).setValue(isTemplateFlag(body.isTemplate != null ? body.isTemplate : body.template) ? "Yes" : "No");
    }
    if (body.templateSpec != null || body.template_spec != null) {
      sheet.getRange(row, 13).setValue(String(body.templateSpec || body.template_spec || "").trim());
    }
    persistWorkbook();
    try { require("./gas").clearShopCache(); } catch (e) {}
    return true;
  }
  throw new Error("Glass line " + want + " was not found.");
}

function purchaseHistory(store, lines) {
  const byId = {};
  lines.forEach((line) => { byId[line.id] = line; });
  const rows = [];
  (store.pos || []).slice().reverse().forEach((po) => {
    (po.lineIds || []).forEach((id) => {
      const live = byId[id] || {};
      const meta = (store.lines && store.lines[id]) || {};
      const merged = Object.assign({}, live, {
        id,
        order: live.order || meta.order || "",
        type: live.type || meta.type || "",
        thickness: live.thickness || meta.thickness || "",
        height: live.height || meta.height || 0,
        width: live.width || meta.width || 0,
        quantity: live.quantity || meta.quantity || 0,
        isTemplate: live.isTemplate || isTemplateFlag(meta.isTemplate),
        templateSpec: live.templateSpec || meta.templateSpec || ""
      });
      const decorated = decorateLine(merged, {
        poId: po.id,
        poNumber: po.number,
        actualCost: meta.cost || "",
        actualCostLabel: meta.cost ? formatRand(meta.cost) : "—"
      });
      rows.push({
        id,
        order: decorated.order,
        type: decorated.type,
        dimensions: decorated.dimensions,
        estimatedCost: decorated.estimatedCost || "",
        estimatedCostLabel: decorated.estimatedCostLabel || "—",
        actualCost: decorated.actualCost,
        actualCostLabel: decorated.actualCostLabel,
        poNumber: po.number,
        isTemplate: decorated.isTemplate,
        templateSpec: decorated.templateSpec,
        status: decorated.status
      });
    });
  });
  return rows;
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
    isTemplate: !!line.isTemplate,
    templateSpec: line.templateSpec || "",
    dimensions: dimensionsLabel(line),
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
    }),
    purchaseHistory: purchaseHistory(store, lines)
  };
}

function createPurchaseOrder(lineIds, actor, edits) {
  const ids = (Array.isArray(lineIds) ? lineIds : [])
    .map((id) => String(id || "").trim())
    .filter(Boolean);
  const unique = [];
  ids.forEach((id) => { if (unique.indexOf(id) === -1) unique.push(id); });
  if (!unique.length) throw new Error("Select at least one glass line to put on a purchase order.");
  const patchById = {};
  (Array.isArray(edits) ? edits : []).forEach((row) => {
    const id = String((row && (row.id || row.lineId)) || "").trim();
    if (!id) return;
    patchById[id] = row;
  });
  unique.forEach((id) => {
    if (patchById[id]) applyLinePatch(id, patchById[id]);
  });
  const lines = readGlassLines();
  const byId = {};
  lines.forEach((line) => { byId[line.id] = line; });
  unique.forEach((id) => {
    const line = byId[id];
    if (!line) throw new Error("Glass line was not found.");
    if (!isToOrder(line.status)) {
      throw new Error((line.order || id) + " is already " + line.status + ". Only To order glass can go on a new PO.");
    }
    if (!String(line.type || "").trim()) throw new Error("Glass type is required on every line.");
    if (!String(line.thickness || "").trim()) throw new Error("Thickness is required on every line.");
    const qty = Number(line.quantity);
    if (!(qty > 0)) throw new Error("Quantity must be more than 0.");
    if (line.isTemplate && !String(line.templateSpec || "").trim()) {
      throw new Error("Specify the template glass for " + (line.order || id) + " before generating the purchase order.");
    }
    if (!line.isTemplate && (!(Number(line.height) > 0) || !(Number(line.width) > 0))) {
      throw new Error("Height and width are required on " + (line.order || id) + " unless it is template glass.");
    }
  });
  const store = loadStore();
  const createdAt = nowIso();
  const createdBy = String(actor || "Admin").trim() || "Admin";
  const number = poNumber(store.nextPo++);
  const poId = newId("gpo");
  unique.forEach((id) => {
    const line = byId[id];
    setGlassStatus(id, ON_PO);
    store.lines[id] = Object.assign({}, store.lines[id] || {}, {
      poId,
      poNumber: number,
      poAt: createdAt,
      poBy: createdBy,
      receiveId: "",
      receivedAt: "",
      cost: "",
      invoiceId: "",
      order: line.order,
      type: line.type,
      thickness: line.thickness,
      height: line.height,
      width: line.width,
      quantity: line.quantity,
      isTemplate: !!line.isTemplate,
      templateSpec: line.templateSpec || ""
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
    "Please supply the glass listed below. Sizes are millimetres. Template glass is not rectangular — see the specification on that line.",
    margin, y, { width: inner }
  );
  y += 28;

  const headers = ["Order Number", "Glass type", "Thickness", "Height", "Width", "Quantity"];
  const widths = [100, 130, 70, 70, 70, inner - 100 - 130 - 70 - 70 - 70];
  const headerH = 22;
  doc.save().fillColor("#1c1917").rect(margin, y, inner, headerH).fill().restore();
  let x = margin;
  headers.forEach((h, i) => {
    doc.fillColor("#fcfbf8").font("Helvetica-Bold").fontSize(7).text(h, x + 3, y + 7, { width: widths[i] - 6, lineBreak: false });
    x += widths[i];
  });
  y += headerH;

  lines.forEach((line, idx) => {
    const spec = String(line.templateSpec || "").trim();
    const template = !!line.isTemplate;
    const rowH = template && spec ? 32 : 18;
    if (y + rowH > 760) {
      doc.addPage();
      y = margin;
    }
    if (idx % 2 === 1) {
      doc.save().fillColor("#f3efe6").rect(margin, y, inner, rowH).fill().restore();
    }
    doc.save().strokeColor("#d7d1c6").lineWidth(0.4).rect(margin, y, inner, rowH).stroke().restore();
    const typeText = template
      ? String(line.type || "") + (spec ? "\nTemplate: " + spec : "\nTemplate glass")
      : String(line.type || "");
    const cells = [
      line.order || "",
      typeText,
      line.thickness || "",
      template ? "Template" : (line.height == null ? "" : String(line.height)),
      template ? "—" : (line.width == null ? "" : String(line.width)),
      line.quantity == null ? "" : String(line.quantity)
    ];
    x = margin;
    cells.forEach((cell, i) => {
      doc.fillColor(ink).font("Helvetica").fontSize(8)
        .text(String(cell), x + 3, y + 5, { width: widths[i] - 6, lineBreak: i === 1 && template });
      x += widths[i];
    });
    y += rowH;
  });

  y += 24;
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
  buildPurchaseOrderPdf,
  applyLinePatch
};
