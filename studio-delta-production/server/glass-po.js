"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir, getBook, persistWorkbook } = require("./workbook-store");
const { parseMoney, money, formatRand } = require("./db");

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
  }, extra || {});
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
  return {
    toOrder,
    outstanding,
    received,
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

module.exports = {
  TO_ORDER,
  ON_PO,
  RECEIVED,
  snapshot,
  createPurchaseOrder,
  receiveGlass,
  readInvoiceFile,
  readGlassLines
};
