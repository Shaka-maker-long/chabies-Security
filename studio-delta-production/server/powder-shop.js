"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir } = require("./workbook-store");
const { listOrders, upsertOrder, formatOrderId, parseMoney, money, formatRand } = require("./db");
const { isAtPaintShop, normalizeShopStatus } = require("./shop-status");
const AT_SHOP_STATUS = "Paint Shop";

const READY_STATUS = "Ready for Powder Coating";
const SENT_STATUS = "Sent to Paint Shop";
const RECEIVED_STATUS = "Ready for Assembly";
const MAX_INVOICE_BYTES = 15 * 1024 * 1024;
const INVOICE_TYPES = {
  "application/pdf": ".pdf",
  "image/jpeg": ".jpg",
  "image/jpg": ".jpg",
  "image/png": ".png",
  "image/webp": ".webp"
};

function nowIso() {
  return new Date().toISOString();
}

function shopPath() {
  return path.join(dataDir(), "paint-shop.json");
}

function invoicesDir() {
  const dir = path.join(dataDir(), "paint-shop-invoices");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function emptyShop() {
  return { sends: [], receives: [], invoices: [], orders: {} };
}

function loadShop() {
  try {
    const parsed = JSON.parse(fs.readFileSync(shopPath(), "utf8"));
    const shop = emptyShop();
    shop.sends = Array.isArray(parsed.sends) ? parsed.sends : [];
    shop.receives = Array.isArray(parsed.receives) ? parsed.receives : [];
    shop.invoices = Array.isArray(parsed.invoices) ? parsed.invoices : [];
    shop.orders = parsed.orders && typeof parsed.orders === "object" ? parsed.orders : {};
    return shop;
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[paint-shop] could not read", shopPath(), e.message || e);
    }
    return emptyShop();
  }
}

function saveShop(shop) {
  const file = shopPath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(shop));
  fs.renameSync(tmp, file);
  return shop;
}

function statusKey(status) {
  return String(status || "").trim().toLowerCase();
}

function isReadyForPowder(status) {
  return statusKey(status) === statusKey(READY_STATUS);
}

function isSentToPaintShop(status) {
  return isAtPaintShop(status) || statusKey(status) === statusKey(SENT_STATUS);
}

function renameOrder(fromRaw, toRaw) {
  const fromId = formatOrderId(fromRaw);
  const toId = formatOrderId(toRaw);
  if (!fromId || !toId || fromId === toId) return false;
  const shop = loadShop();
  let changed = false;
  if (shop.orders[fromId]) {
    shop.orders[toId] = shop.orders[fromId];
    delete shop.orders[fromId];
    changed = true;
  }
  (shop.sends || []).forEach((send) => {
    if (!send || !Array.isArray(send.orderNumbers)) return;
    let hit = false;
    send.orderNumbers = send.orderNumbers.map((n) => {
      if (formatOrderId(n) !== fromId) return n;
      hit = true;
      return toId;
    });
    if (hit) changed = true;
  });
  (shop.receives || []).forEach((recv) => {
    if (!recv) return;
    if (Array.isArray(recv.orderNumbers)) {
      let hit = false;
      recv.orderNumbers = recv.orderNumbers.map((n) => {
        if (formatOrderId(n) !== fromId) return n;
        hit = true;
        return toId;
      });
      if (hit) changed = true;
    }
    if (Array.isArray(recv.lines)) {
      recv.lines.forEach((line) => {
        if (!line) return;
        if (formatOrderId(line.orderNumber || line.order_number) === fromId) {
          line.orderNumber = toId;
          if (line.order_number != null) line.order_number = toId;
          changed = true;
        }
      });
    }
  });
  if (changed) saveShop(shop);
  return changed;
}

function canMarkAtPaintShop(status) {
  if (isSentToPaintShop(status) || isReadyForPowder(status)) return false;
  const s = normalizeShopStatus(status);
  return !s || s === "Not Yet Started";
}

function findOrder(orderNumber) {
  const want = formatOrderId(orderNumber);
  return listOrders().find((o) => formatOrderId(o.order_number) === want) || null;
}

function newId(prefix) {
  return prefix + "_" + crypto.randomBytes(8).toString("hex");
}

function safeFilename(name, fallback) {
  const cleaned = String(name || "").replace(/[^A-Za-z0-9._-]+/g, "_");
  return cleaned || fallback || "invoice.pdf";
}

function decodeInvoice(invoice) {
  if (!invoice || typeof invoice !== "object") {
    throw new Error("Attach the paint shop invoice for this received batch.");
  }
  const raw = String(invoice.data || invoice.dataUrl || invoice.base64 || "");
  if (!raw.trim()) throw new Error("Attach the paint shop invoice for this received batch.");
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
  if (!buffer.length) throw new Error("Attach the paint shop invoice for this received batch.");
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
  const shop = loadShop();
  const rec = (shop.invoices || []).find((row) => row && row.id === invoiceId);
  if (!rec) return null;
  const file = path.join(invoicesDir(), invoiceId, rec.storedAs || "invoice.pdf");
  if (!fs.existsSync(file)) return null;
  return {
    buffer: fs.readFileSync(file),
    filename: rec.filename || rec.storedAs || "invoice.pdf",
    mime: rec.mime || "application/octet-stream"
  };
}

function publicOrder(order, extra) {
  const row = {
    order_number: order.order_number,
    product: order.product || "",
    powder_coating: order.powder_coating || "",
    client_name: order.client_name || "",
    status: order.status || "",
    dimensions: order.dimensions || ""
  };
  return Object.assign(row, extra || {});
}

function snapshot() {
  const shop = loadShop();
  const orders = listOrders();
  const byNumber = {};
  orders.forEach((o) => { byNumber[formatOrderId(o.order_number)] = o; });
  const ready = orders.filter((o) => isReadyForPowder(o.status)).map((o) => publicOrder(o));
  const misplaced = orders.filter((o) => canMarkAtPaintShop(o.status)).map((o) => publicOrder(o));
  const sent = orders.filter((o) => isSentToPaintShop(o.status)).map((o) => {
    const meta = shop.orders[formatOrderId(o.order_number)] || {};
    return publicOrder(o, {
      sendId: meta.sendId || "",
      sentAt: meta.sentAt || "",
      sentBy: meta.sentBy || ""
    });
  });
  const received = (shop.receives || []).slice().reverse().map((batch) => {
    const invoice = (shop.invoices || []).find((row) => row.id === batch.invoiceId) || {};
    return {
      id: batch.id,
      receivedAt: batch.receivedAt,
      receivedBy: batch.receivedBy,
      invoiceId: batch.invoiceId,
      invoiceFilename: invoice.filename || "",
      orders: (batch.orders || []).map((line) => {
        const live = byNumber[formatOrderId(line.orderNumber)] || { order_number: line.orderNumber };
        return publicOrder(live, {
          cost: money(line.cost),
          costLabel: formatRand(line.cost),
          receiveId: batch.id,
          invoiceId: batch.invoiceId
        });
      }),
      total: money((batch.orders || []).reduce((sum, line) => sum + parseMoney(line.cost), 0)),
      totalLabel: formatRand((batch.orders || []).reduce((sum, line) => sum + parseMoney(line.cost), 0))
    };
  });
  const readyForDelivery = orders
    .filter((o) => statusKey(o.status) === statusKey("Ready for Delivery"))
    .map((o) => {
      const meta = shop.orders[formatOrderId(o.order_number)] || {};
      return publicOrder(o, {
        cost: meta.cost || "",
        costLabel: meta.cost ? formatRand(meta.cost) : "",
        receivedAt: meta.receivedAt || "",
        receivedBy: meta.receivedBy || ""
      });
    });
  return {
    ready,
    misplaced,
    sent,
    received,
    readyForDelivery,
    readyCount: ready.length,
    misplacedCount: misplaced.length,
    sentCount: sent.length,
    readyForDeliveryCount: readyForDelivery.length,
    receivedCount: received.length
  };
}

function assertReadyForPowderList(orderNumbers) {
  const nums = (Array.isArray(orderNumbers) ? orderNumbers : [])
    .map((n) => formatOrderId(n))
    .filter(Boolean);
  if (!nums.length) throw new Error("Select at least one order that is Ready for Powder Coating.");
  nums.forEach((num) => {
    const order = findOrder(num);
    if (!order) throw new Error("Order " + num + " was not found.");
    if (!isReadyForPowder(order.status)) {
      throw new Error(num + " is " + (order.status || "not ready") + ". Only Ready for Powder Coating can go on the list.");
    }
  });
}

function sendToPaintShop(orderNumbers, actor, opts) {
  const nums = (Array.isArray(orderNumbers) ? orderNumbers : [])
    .map((n) => formatOrderId(n))
    .filter(Boolean);
  const unique = [];
  nums.forEach((n) => { if (unique.indexOf(n) === -1) unique.push(n); });
  if (!unique.length) throw new Error("Select at least one order that is Ready for Powder Coating.");
  const nextStatus = String((opts && opts.status) || SENT_STATUS).trim() || SENT_STATUS;
  const shop = loadShop();
  const sentAt = nowIso();
  const sentBy = String(actor || "Admin").trim() || "Admin";
  const sendId = newId("send");
  const updated = [];
  unique.forEach((num) => {
    const order = findOrder(num);
    if (!order) throw new Error("Order " + num + " was not found.");
    if (!isReadyForPowder(order.status)) {
      throw new Error(num + " is " + (order.status || "not ready") + ". Only Ready for Powder Coating can be sent.");
    }
    upsertOrder(Object.assign({}, order, { status: nextStatus }));
    shop.orders[num] = Object.assign({}, shop.orders[num] || {}, {
      sendId,
      sentAt,
      sentBy,
      receiveId: "",
      receivedAt: "",
      receivedBy: "",
      cost: "",
      invoiceId: ""
    });
    updated.push(num);
  });
  shop.sends.push({ id: sendId, sentAt, sentBy, orderNumbers: updated });
  saveShop(shop);
  try { require("./gas").clearShopCache(); } catch (e) {}
  return { sendId, sentAt, sentBy, orderNumbers: updated, ...snapshot() };
}

function markAlreadyAtPaintShop(orderNumbers, actor) {
  const nums = (Array.isArray(orderNumbers) ? orderNumbers : [])
    .map((n) => formatOrderId(n))
    .filter(Boolean);
  const unique = [];
  nums.forEach((n) => { if (unique.indexOf(n) === -1) unique.push(n); });
  if (!unique.length) throw new Error("Tick the orders that are already at the powder coaters.");
  const shop = loadShop();
  const sentAt = nowIso();
  const sentBy = String(actor || "Admin").trim() || "Admin";
  const sendId = newId("had");
  const updated = [];
  unique.forEach((num) => {
    const order = findOrder(num);
    if (!order) throw new Error("Order " + num + " was not found.");
    if (isSentToPaintShop(order.status)) {
      updated.push(num);
      return;
    }
    if (!canMarkAtPaintShop(order.status)) {
      throw new Error(num + " is " + (order.status || "not on the sheet") + " and cannot be moved to Paint Shop.");
    }
    upsertOrder(Object.assign({}, order, { status: AT_SHOP_STATUS }));
    shop.orders[num] = Object.assign({}, shop.orders[num] || {}, {
      sendId,
      sentAt,
      sentBy,
      adopted: true,
      receiveId: "",
      receivedAt: "",
      receivedBy: "",
      cost: "",
      invoiceId: ""
    });
    updated.push(num);
  });
  shop.sends.push({ id: sendId, sentAt, sentBy, adopted: true, orderNumbers: updated });
  saveShop(shop);
  try { require("./gas").clearShopCache(); } catch (e) {}
  return { sendId, sentAt, sentBy, orderNumbers: updated, ...snapshot() };
}

function parseReceiveLines(orders) {
  const lines = Array.isArray(orders) ? orders : [];
  if (!lines.length) throw new Error("Select the orders that came back from the paint shop.");
  const seen = {};
  return lines.map((line) => {
    const orderNumber = formatOrderId(line && (line.orderNumber || line.order_number || line.order));
    if (!orderNumber) throw new Error("Each received order needs an order number.");
    if (seen[orderNumber]) throw new Error(orderNumber + " is listed twice on this receive.");
    seen[orderNumber] = true;
    if (line == null || line.cost === "" || line.cost == null) {
      throw new Error("Put the paint shop cost on " + orderNumber + ".");
    }
    const cost = parseMoney(line.cost);
    if (!Number.isFinite(cost)) throw new Error("Put the paint shop cost on " + orderNumber + ".");
    if (cost < 0) throw new Error("The cost on " + orderNumber + " cannot be negative.");
    return { orderNumber, cost: Number(money(cost)) };
  });
}

function receiveFromPaintShop(payload, actor) {
  const body = payload || {};
  const lines = parseReceiveLines(body.orders);
  const found = lines.map((line) => {
    const order = findOrder(line.orderNumber);
    if (!order) throw new Error("Order " + line.orderNumber + " was not found.");
    if (!isSentToPaintShop(order.status)) {
      throw new Error(line.orderNumber + " is not at the paint shop. Send it first, then receive it.");
    }
    return { line, order };
  });
  decodeInvoice(body.invoice);
  const invoiceId = newId("inv");
  const file = writeInvoiceFile(invoiceId, body.invoice);
  const shop = loadShop();
  const receivedAt = nowIso();
  const receivedBy = String(actor || "Admin").trim() || "Admin";
  const receiveId = newId("recv");
  found.forEach(({ line, order }) => {
    upsertOrder(Object.assign({}, order, { status: RECEIVED_STATUS }));
    shop.orders[line.orderNumber] = Object.assign({}, shop.orders[line.orderNumber] || {}, {
      receiveId,
      receivedAt,
      receivedBy,
      cost: money(line.cost),
      invoiceId
    });
  });
  shop.invoices.push({
    id: invoiceId,
    filename: file.filename,
    storedAs: file.storedAs,
    mime: file.mime,
    size: file.size,
    receiveId,
    receivedAt,
    receivedBy,
    orderNumbers: lines.map((line) => line.orderNumber)
  });
  shop.receives.push({
    id: receiveId,
    receivedAt,
    receivedBy,
    invoiceId,
    orders: lines
  });
  saveShop(shop);
  try { require("./gas").clearShopCache(); } catch (e) {}
  return {
    receiveId,
    invoiceId,
    invoiceFilename: file.filename,
    receivedAt,
    receivedBy,
    orderNumbers: lines.map((line) => line.orderNumber),
    ...snapshot()
  };
}

async function parseInvoiceForCosts(payload) {
  const body = payload || {};
  const decoded = decodeInvoice(body.invoice);
  if (decoded.mime !== "application/pdf") {
    throw new Error("Upload a PDF invoice to auto-read costs (photos need typing for now).");
  }
  const prefer = Array.isArray(body.orderNumbers)
    ? body.orderNumbers
    : (Array.isArray(body.orders) ? body.orders.map((o) => o && (o.orderNumber || o.order_number || o)) : []);
  const parser = require("./paint-invoice-parse");
  const parsed = await parser.parsePaintInvoiceBuffer(decoded.buffer, {
    preferOrderNumbers: prefer
  });
  if (!parsed.ok) throw new Error(parsed.error || "Could not read that invoice.");
  const snap = snapshot();
  const sentSet = {};
  const comparedWith = (snap.sent || []).map((o) => {
    const orderNumber = formatOrderId(o.order_number);
    sentSet[orderNumber] = true;
    return {
      orderNumber,
      product: o.product || "",
      powder_coating: o.powder_coating || "",
      client_name: o.client_name || "",
      status: o.status || ""
    };
  });
  const preferIds = (prefer || []).map(formatOrderId).filter(Boolean);
  const allocations = (parsed.allocations || []).map((row) => {
    const orderNumber = formatOrderId(row.orderNumber);
    const atPaintShop = !!sentSet[orderNumber];
    const preferred = preferIds.indexOf(orderNumber) !== -1;
    return Object.assign({}, row, {
      orderNumber,
      atPaintShop,
      preferred,
      matchedToShop: atPaintShop,
      costLabel: formatRand(row.cost)
    });
  });
  const matchedCount = allocations.filter((row) => row.atPaintShop).length;
  const matchedTotal = allocations
    .filter((row) => row.atPaintShop)
    .reduce((sum, row) => sum + Number(row.cost || 0), 0);
  return {
    filename: decoded.filename,
    mime: decoded.mime,
    orderTokens: parsed.orderTokens || [],
    allocations,
    unmatchedLines: parsed.unmatchedLines || [],
    textPreview: parsed.textPreview || "",
    comparedWith,
    comparedCount: comparedWith.length,
    matchedCount,
    grandTotal: parsed.grandTotal != null ? parsed.grandTotal : allocations.reduce((sum, row) => sum + Number(row.cost || 0), 0),
    grandTotalLabel: parsed.grandTotalLabel || formatRand(allocations.reduce((sum, row) => sum + Number(row.cost || 0), 0)),
    matchedTotal: Number(require("./db").money(matchedTotal)),
    matchedTotalLabel: formatRand(matchedTotal),
    filledCount: matchedCount || allocations.filter((row) => preferIds.indexOf(formatOrderId(row.orderNumber)) !== -1).length
  };
}

function updateReceivedCosts(payload, actor) {
  const body = payload || {};
  const receiveId = String(body.receiveId || body.id || "").trim();
  if (!receiveId) throw new Error("Which received batch are you updating?");
  const lines = parseReceiveLines(body.orders);
  const shop = loadShop();
  const batch = (shop.receives || []).find((row) => row && row.id === receiveId);
  if (!batch) throw new Error("That received batch was not found.");
  const byOrder = {};
  (batch.orders || []).forEach((line) => {
    byOrder[formatOrderId(line.orderNumber)] = line;
  });
  lines.forEach((line) => {
    const existing = byOrder[line.orderNumber];
    if (!existing) {
      throw new Error(line.orderNumber + " is not on that received batch.");
    }
    existing.cost = Number(money(line.cost));
    shop.orders[line.orderNumber] = Object.assign({}, shop.orders[line.orderNumber] || {}, {
      cost: money(line.cost),
      receiveId,
      costUpdatedAt: nowIso(),
      costUpdatedBy: String(actor || "Admin").trim() || "Admin"
    });
  });
  saveShop(shop);
  try { require("./gas").clearShopCache(); } catch (e) {}
  return {
    receiveId,
    orderNumbers: lines.map((line) => line.orderNumber),
    ...snapshot()
  };
}

module.exports = {
  READY_STATUS,
  SENT_STATUS,
  AT_SHOP_STATUS,
  RECEIVED_STATUS,
  isReadyForPowder,
  isSentToPaintShop,
  snapshot,
  assertReadyForPowderList,
  sendToPaintShop,
  markAlreadyAtPaintShop,
  receiveFromPaintShop,
  updateReceivedCosts,
  parseInvoiceForCosts,
  readInvoiceFile,
  loadShop,
  renameOrder,
  canMarkAtPaintShop
};
