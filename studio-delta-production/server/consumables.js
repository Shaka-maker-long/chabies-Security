"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir } = require("./workbook-store");
const { parseMoney, money, formatRand } = require("./db");
const CATALOG = require("./consumables-catalog");

const UNITS = ["pcs", "box", "pack", "roll", "pair", "set", "ℓ"];

function storePath() {
  return path.join(dataDir(), "consumables.json");
}

function emptyStore() {
  return { items: [], movements: [], purchases: [] };
}

function loadStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(storePath(), "utf8"));
    const items = Array.isArray(parsed.items) ? parsed.items.filter((row) => row && row.id && row.name) : [];
    const movements = Array.isArray(parsed.movements) ? parsed.movements.filter((row) => row && row.id) : [];
    const purchases = Array.isArray(parsed.purchases) ? parsed.purchases.filter((row) => row && row.id) : [];
    return { items, movements, purchases };
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[consumables] could not read", storePath(), e.message || e);
    }
    return emptyStore();
  }
}

function saveStore(store) {
  const file = storePath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify({
    items: store.items || [],
    movements: store.movements || [],
    purchases: store.purchases || []
  }));
  fs.renameSync(tmp, file);
  return store;
}

function nowIso() {
  return new Date().toISOString();
}

function formatWhen(iso) {
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  return d.toLocaleString("en-ZA", {
    timeZone: "Africa/Johannesburg",
    year: "numeric",
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });
}

function formatQty(n) {
  const v = Number(n);
  if (!Number.isFinite(v)) return "0";
  const s = (Math.round(v * 1000) / 1000).toFixed(3);
  return s.replace(/\.?0+$/, "");
}

function parseQty(raw, label) {
  const n = Number(String(raw == null ? "" : raw).replace(/,/g, "").trim());
  if (!Number.isFinite(n) || n <= 0) throw new Error((label || "Quantity") + " must be more than 0.");
  return Math.round(n * 1000) / 1000;
}

function parseStock(raw, label) {
  const s = String(raw == null ? "" : raw).replace(/,/g, "").trim();
  if (s === "") return 0;
  const n = Number(s);
  if (!Number.isFinite(n) || n < 0) throw new Error((label || "Stock") + " must be 0 or more.");
  return Math.round(n * 1000) / 1000;
}

function parseThreshold(raw) {
  const s = String(raw == null ? "" : raw).replace(/,/g, "").trim();
  if (s === "") return 0;
  const n = Number(s);
  if (!Number.isFinite(n) || n < 0) throw new Error("ROP must be 0 or more.");
  return Math.round(n * 1000) / 1000;
}

function normalizeName(name) {
  return String(name || "").replace(/\s+/g, " ").trim();
}

function nameKey(name) {
  return normalizeName(name).toLowerCase();
}

function normalizeUnit(unit) {
  const s = String(unit || "").trim();
  return s || "pcs";
}

function actorName(actor) {
  const s = String(actor || "").trim();
  return s || "Office";
}

function newId(prefix) {
  return prefix + "_" + crypto.randomBytes(6).toString("hex");
}

function findItem(store, id) {
  const want = String(id || "").trim();
  if (!want) return null;
  return store.items.find((row) => row.id === want) || null;
}

function findItemByName(store, name, exceptId) {
  const want = nameKey(name);
  if (!want) return null;
  return store.items.find((row) => nameKey(row.name) === want && row.id !== exceptId) || null;
}

function shopOrderNumbers() {
  try {
    return require("./db").listOrders()
      .map((row) => String(row.order_number || "").trim())
      .filter(Boolean);
  } catch (e) {
    return [];
  }
}

function isActiveShopStatus(status) {
  const s = String(status || "").trim().toLowerCase();
  if (!s) return true;
  if (s === "delivered") return false;
  return true;
}

function activeOrderNumbers() {
  try {
    return require("./db").listOrders()
      .filter((row) => isActiveShopStatus(row.status))
      .map((row) => String(row.order_number || "").trim())
      .filter(Boolean)
      .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base", numeric: true }));
  } catch (e) {
    return [];
  }
}

function productionWorkers() {
  try {
    const staff = require("./staff");
    return staff.listUsers()
      .filter((u) => u && u.name && String(u.access || "") === "Production")
      .map((u) => u.name)
      .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
  } catch (e) {
    return [];
  }
}

function orderedQtyFor(store, itemId) {
  let n = 0;
  (store && store.purchases ? store.purchases : []).forEach((po) => {
    if (po.status !== "Ordered") return;
    (po.lines || []).forEach((line) => {
      if (line.itemId === itemId) n += Number(line.qty) || 0;
    });
  });
  return Math.round(n * 1000) / 1000;
}

function parsePrice(raw) {
  const s = String(raw == null ? "" : raw).trim();
  if (s === "") return null;
  const n = parseMoney(s);
  if (!Number.isFinite(n) || n < 0) throw new Error("Unit price must be 0 or more.");
  return Number(money(n));
}

function decorateItem(row, store) {
  const stock = Number(row.stock) || 0;
  const minThreshold = Number(row.minThreshold) || 0;
  const orderedQty = orderedQtyFor(store, row.id);
  const unitPrice = row.unitPrice == null || row.unitPrice === "" ? null : Number(row.unitPrice);
  const totalValue = unitPrice == null ? null : Math.round(stock * unitPrice * 100) / 100;
  const low = minThreshold > 0 && stock <= minThreshold;
  return {
    id: row.id,
    name: row.name,
    unit: row.unit || "pcs",
    stock,
    stockLabel: formatQty(stock),
    orderedQty,
    orderedLabel: formatQty(orderedQty),
    minThreshold,
    minLabel: formatQty(minThreshold),
    ropLabel: formatQty(minThreshold),
    unitPrice,
    priceLabel: unitPrice == null ? "—" : formatRand(unitPrice),
    totalValue,
    totalValueLabel: totalValue == null ? "—" : formatRand(totalValue),
    low,
    status: low ? "Low" : "OK",
    createdAt: row.createdAt || "",
    updatedAt: row.updatedAt || ""
  };
}

function decorateMovement(row) {
  const qty = Number(row.qty) || 0;
  return {
    id: row.id,
    at: row.at,
    whenLabel: formatWhen(row.at),
    type: row.type,
    typeLabel: typeLabel(row.type),
    itemId: row.itemId || "",
    itemName: row.itemName || "",
    qty,
    qtyLabel: (qty > 0 && row.type !== "usage" ? "+" : "") + formatQty(qty),
    orderNumber: row.orderNumber || "",
    employee: row.employee || "",
    note: row.note || "",
    supplier: row.supplier || "",
    poId: row.poId || ""
  };
}

function typeLabel(type) {
  if (type === "usage") return "Used";
  if (type === "receive") return "Received";
  if (type === "purchase_order") return "Ordered";
  if (type === "purchase_receive") return "PO received";
  if (type === "count") return "Count";
  if (type === "opening") return "Opening stock";
  return String(type || "");
}

function decoratePurchase(row) {
  const lines = (row.lines || []).map((line) => {
    const qty = Number(line.qty) || 0;
    const unitPrice = line.unitPrice == null || line.unitPrice === "" ? null : Number(line.unitPrice);
    const lineTotal = unitPrice == null ? null : Math.round(qty * unitPrice * 100) / 100;
    return {
      itemId: line.itemId,
      itemName: line.itemName,
      unit: line.unit || "pcs",
      qty,
      qtyLabel: formatQty(line.qty),
      unitPrice,
      priceLabel: unitPrice == null ? "—" : formatRand(unitPrice),
      lineTotal,
      lineTotalLabel: lineTotal == null ? "—" : formatRand(lineTotal),
      receivedQty: Number(line.receivedQty) || 0
    };
  });
  const total = lines.reduce((sum, line) => sum + (line.lineTotal == null ? 0 : line.lineTotal), 0);
  const hasPrices = lines.some((line) => line.unitPrice != null);
  return {
    id: row.id,
    number: row.number || row.id,
    createdAt: row.createdAt,
    whenLabel: formatWhen(row.createdAt),
    receivedAt: row.receivedAt || "",
    receivedLabel: row.receivedAt ? formatWhen(row.receivedAt) : "",
    status: row.status,
    supplier: row.supplier || "",
    note: row.note || "",
    employee: row.employee || "",
    receivedBy: row.receivedBy || "",
    lines,
    totalValue: hasPrices ? Math.round(total * 100) / 100 : null,
    totalValueLabel: hasPrices ? formatRand(total) : "—",
    pdfUrl: row.hasPdf ? ("/api/office/consumables/purchases/" + encodeURIComponent(row.id) + "/pdf") : "",
    hasPdf: !!row.hasPdf
  };
}

function addMovement(store, row) {
  store.movements.unshift(row);
  if (store.movements.length > 2000) store.movements = store.movements.slice(0, 2000);
}

function seedCatalog() {
  const store = loadStore();
  let added = 0;
  const at = nowIso();
  CATALOG.forEach((row) => {
    const name = normalizeName(row && row.name);
    if (!name || findItemByName(store, name)) return;
    store.items.push({
      id: newId("citem"),
      name,
      unit: "pcs",
      stock: Number(row.stock) || 0,
      minThreshold: Number(row.min) || 0,
      unitPrice: row.price == null || row.price === "" ? null : Number(row.price),
      createdAt: at,
      updatedAt: at
    });
    added += 1;
  });
  if (added) saveStore(store);
  return added;
}

function pdfDir() {
  const dir = path.join(dataDir(), "consumables-pos");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function pdfPathFor(id) {
  return path.join(pdfDir(), String(id) + ".pdf");
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

function nextPoNumber(purchases, now) {
  const d = now ? new Date(now) : new Date();
  const ymd = isNaN(d.getTime())
    ? "00000000"
    : d.toLocaleDateString("en-CA", { timeZone: "Africa/Johannesburg" }).replace(/-/g, "");
  const prefix = "CPO-" + ymd + "-";
  let max = 0;
  (purchases || []).forEach((row) => {
    const n = String((row && row.number) || "");
    if (n.indexOf(prefix) !== 0) return;
    const seq = Number(n.slice(prefix.length));
    if (seq > max) max = seq;
  });
  return prefix + String(max + 1);
}

async function logoPathForPdf() {
  try {
    const catalog = require("./product-catalog");
    const url = catalog.COMPANY_LOGO_URL;
    if (!url) return null;
    const dir = path.join(dataDir(), "pdf-images");
    fs.mkdirSync(dir, { recursive: true });
    const dest = path.join(dir, "studio-delta-logo.jpg");
    if (fs.existsSync(dest) && fs.statSync(dest).size > 400) return dest;
    const res = await fetch(url, { redirect: "follow", signal: AbortSignal.timeout(1500) });
    if (!res.ok) return fs.existsSync(dest) ? dest : null;
    const buf = Buffer.from(await res.arrayBuffer());
    if (buf.length) fs.writeFileSync(dest, buf);
    return dest;
  } catch (e) {
    return null;
  }
}

function drawPdfBox(doc, x, y, w, h) {
  doc.save().lineWidth(0.7).strokeColor("#1c1917").rect(x, y, w, h).stroke().restore();
}

async function buildPurchasePdf(po) {
  const PDFDocument = require("pdfkit");
  const logoPath = await logoPathForPdf();
  const chunks = [];
  const doc = new PDFDocument({
    size: "A4",
    margin: 36,
    compress: false,
    info: { Title: "Consumables PO " + (po.number || po.id), Author: "Studio Delta" }
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
  doc.fillColor(brass).font("Helvetica-Bold").fontSize(11).text(po.number || po.id, margin, margin + 30, { width: inner, align: "right" });
  doc.fillColor(muted).font("Helvetica").fontSize(9).text("Consumables", margin, margin + 46, { width: inner, align: "right" });

  doc.save().strokeColor(brass).lineWidth(2).moveTo(margin, margin + 84).lineTo(margin + inner, margin + 84).stroke().restore();

  let y = margin + 96;
  drawPdfBox(doc, margin, y, inner * 0.48, 54);
  drawPdfBox(doc, margin + inner * 0.52, y, inner * 0.48, 54);
  doc.fillColor(muted).font("Helvetica-Bold").fontSize(8).text("SUPPLIER", margin + 8, y + 8);
  doc.fillColor(ink).font("Helvetica").fontSize(10).text(po.supplier || "—", margin + 8, y + 22, { width: inner * 0.48 - 16 });
  doc.fillColor(muted).font("Helvetica").fontSize(8).text(po.note || "", margin + 8, y + 36, { width: inner * 0.48 - 16 });
  doc.fillColor(muted).font("Helvetica-Bold").fontSize(8).text("ORDER DETAILS", margin + inner * 0.52 + 8, y + 8);
  doc.fillColor(ink).font("Helvetica").fontSize(10)
    .text("Date  " + formatPoDate(po.createdAt), margin + inner * 0.52 + 8, y + 22)
    .text("Prepared by  " + (po.employee || "Studio Delta"), margin + inner * 0.52 + 8, y + 36);

  y += 70;
  doc.fillColor(ink).font("Helvetica").fontSize(9).text(
    "Please supply the consumables listed below.",
    margin, y, { width: inner }
  );
  y += 24;

  const headers = ["Item", "Unit", "QTY", "Unit price", "Line total"];
  const widths = [220, 50, 50, 90, inner - 220 - 50 - 50 - 90];
  const headerH = 22;
  doc.save().fillColor("#1c1917").rect(margin, y, inner, headerH).fill().restore();
  let x = margin;
  headers.forEach((h, i) => {
    doc.fillColor("#fcfbf8").font("Helvetica-Bold").fontSize(7).text(h, x + 3, y + 7, { width: widths[i] - 6, lineBreak: false });
    x += widths[i];
  });
  y += headerH;

  let total = 0;
  let hasPrices = false;
  (po.lines || []).forEach((line, idx) => {
    const name = String(line.itemName || "");
    const rowH = Math.max(20, doc.heightOfString(name, { width: widths[0] - 6, fontSize: 8 }) + 8);
    if (y + rowH > 760) {
      doc.addPage();
      y = margin;
    }
    if (idx % 2 === 1) {
      doc.save().fillColor("#f6f3ee").rect(margin, y, inner, rowH).fill().restore();
    }
    const unitPrice = line.unitPrice == null ? null : Number(line.unitPrice);
    const lineTotal = unitPrice == null ? null : Math.round((Number(line.qty) || 0) * unitPrice * 100) / 100;
    if (lineTotal != null) {
      hasPrices = true;
      total += lineTotal;
    }
    const cells = [
      name,
      String(line.unit || "pcs"),
      formatQty(line.qty),
      unitPrice == null ? "—" : formatRand(unitPrice),
      lineTotal == null ? "—" : formatRand(lineTotal)
    ];
    x = margin;
    cells.forEach((cell, i) => {
      doc.fillColor(ink).font(i === 0 ? "Helvetica-Bold" : "Helvetica").fontSize(8)
        .text(cell, x + 3, y + 5, { width: widths[i] - 6 });
      x += widths[i];
    });
    y += rowH;
  });

  if (hasPrices) {
    y += 12;
    doc.fillColor(ink).font("Helvetica-Bold").fontSize(10)
      .text("Total  " + formatRand(total), margin, y, { width: inner, align: "right" });
  }

  doc.end();
  return done;
}

function snapshot() {
  seedCatalog();
  const store = loadStore();
  const items = store.items.slice().sort((a, b) => String(a.name).localeCompare(String(b.name), undefined, { sensitivity: "base" })).map((row) => decorateItem(row, store));
  const low = items.filter((row) => row.low);
  const purchases = store.purchases.slice().sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt))).map(decoratePurchase);
  const openPurchases = purchases.filter((row) => row.status === "Ordered");
  const totalValue = items.reduce((sum, row) => sum + (row.totalValue == null ? 0 : row.totalValue), 0);
  return {
    items,
    low,
    lowCount: low.length,
    itemCount: items.length,
    totalValue: Math.round(totalValue * 100) / 100,
    totalValueLabel: formatRand(totalValue),
    units: UNITS.slice(),
    movements: store.movements.slice(0, 120).map(decorateMovement),
    purchases,
    openPurchases,
    openPurchaseCount: openPurchases.length,
    orderNumbers: shopOrderNumbers(),
    activeOrderNumbers: activeOrderNumbers(),
    productionWorkers: productionWorkers()
  };
}

function upsertItem(body, actor) {
  const name = normalizeName(body && body.name);
  if (!name) throw new Error("Item name is required.");
  const unit = normalizeUnit(body && body.unit);
  const minThreshold = parseThreshold(body && body.minThreshold);
  const opening = body && body.openingStock != null && String(body.openingStock).trim() !== ""
    ? parseStock(body.openingStock, "Opening stock")
    : 0;
  const store = loadStore();
  const existingId = String((body && body.id) || "").trim();
  let row = existingId ? findItem(store, existingId) : null;
  if (findItemByName(store, name, row && row.id)) {
    throw new Error(name + " is already on the list.");
  }
  const at = nowIso();
  const employee = actorName(actor);
  if (!row) {
    row = {
      id: newId("citem"),
      name,
      unit,
      stock: 0,
      minThreshold,
      unitPrice: body && Object.prototype.hasOwnProperty.call(body, "unitPrice") ? parsePrice(body.unitPrice) : null,
      createdAt: at,
      updatedAt: at
    };
    store.items.push(row);
    if (opening > 0) {
      row.stock = opening;
      addMovement(store, {
        id: newId("cmove"),
        at,
        type: "opening",
        itemId: row.id,
        itemName: row.name,
        qty: opening,
        orderNumber: "",
        employee,
        note: "Opening stock"
      });
    }
  } else {
    row.name = name;
    row.unit = unit;
    row.minThreshold = minThreshold;
    row.updatedAt = at;
  }
  if (body && Object.prototype.hasOwnProperty.call(body, "unitPrice")) {
    row.unitPrice = parsePrice(body.unitPrice);
  }
  saveStore(store);
  return decorateItem(row, store);
}

function deleteItem(id) {
  const want = String(id || "").trim();
  if (!want) throw new Error("Item not found.");
  const store = loadStore();
  const row = findItem(store, want);
  if (!row) throw new Error("Item not found.");
  if ((Number(row.stock) || 0) > 0) {
    throw new Error("Use or count " + row.name + " down to 0 before deleting it.");
  }
  const open = store.purchases.some((po) => po.status === "Ordered" && (po.lines || []).some((line) => line.itemId === want));
  if (open) throw new Error("Receive or cancel the open purchase that still has " + row.name + ".");
  store.items = store.items.filter((item) => item.id !== want);
  saveStore(store);
  return true;
}

function logUsage(body, actor) {
  const qty = parseQty(body && body.qty, "Quantity used");
  let orderNumber = String((body && (body.orderNumber || body.order_number)) || "").trim();
  if (/^general$/i.test(orderNumber)) orderNumber = "General";
  const note = String((body && body.note) || "").trim();
  const worker = String((body && (body.worker || body.employee)) || "").trim() || actorName(actor);
  const store = loadStore();
  const row = findItem(store, body && (body.itemId || body.item_id));
  if (!row) throw new Error("Choose a consumable.");
  const stock = Number(row.stock) || 0;
  if (qty > stock) {
    throw new Error(row.name + " has " + formatQty(stock) + " " + (row.unit || "pcs") + " on hand. You cannot use " + formatQty(qty) + ".");
  }
  const at = nowIso();
  row.stock = Math.round((stock - qty) * 1000) / 1000;
  row.updatedAt = at;
  addMovement(store, {
    id: newId("cmove"),
    at,
    type: "usage",
    itemId: row.id,
    itemName: row.name,
    qty: -qty,
    orderNumber,
    employee: worker,
    note
  });
  saveStore(store);
  return decorateItem(row, store);
}

function receiveStock(body, actor) {
  const qty = parseQty(body && body.qty, "Quantity received");
  const note = String((body && body.note) || "").trim();
  const supplier = String((body && body.supplier) || "").trim();
  const store = loadStore();
  const row = findItem(store, body && body.itemId);
  if (!row) throw new Error("Choose a consumable.");
  const at = nowIso();
  row.stock = Math.round(((Number(row.stock) || 0) + qty) * 1000) / 1000;
  row.updatedAt = at;
  addMovement(store, {
    id: newId("cmove"),
    at,
    type: "receive",
    itemId: row.id,
    itemName: row.name,
    qty,
    orderNumber: "",
    employee: actorName(actor),
    note,
    supplier
  });
  saveStore(store);
  return decorateItem(row, store);
}

function countStock(body, actor) {
  const counted = parseStock(body && body.stock, "Counted stock");
  const note = String((body && body.note) || "").trim();
  if (!note) throw new Error("Say why the count changed.");
  const store = loadStore();
  const row = findItem(store, body && body.itemId);
  if (!row) throw new Error("Choose a consumable.");
  const before = Number(row.stock) || 0;
  const delta = Math.round((counted - before) * 1000) / 1000;
  if (delta === 0) throw new Error(row.name + " is already " + formatQty(before) + ".");
  const at = nowIso();
  row.stock = counted;
  row.updatedAt = at;
  addMovement(store, {
    id: newId("cmove"),
    at,
    type: "count",
    itemId: row.id,
    itemName: row.name,
    qty: delta,
    orderNumber: "",
    employee: actorName(actor),
    note
  });
  saveStore(store);
  return decorateItem(row, store);
}

async function createPurchase(body, actor) {
  const supplier = String((body && body.supplier) || "").trim();
  const note = String((body && body.note) || "").trim();
  const rawLines = Array.isArray(body && body.lines) ? body.lines : [];
  if (!rawLines.length) throw new Error("Add at least one item to the purchase.");
  const store = loadStore();
  const lines = [];
  rawLines.forEach((line, i) => {
    const item = findItem(store, line && line.itemId);
    if (!item) throw new Error("Line " + (i + 1) + " needs a consumable.");
    const qty = parseQty(line.qty, item.name + " quantity");
    if (lines.some((existing) => existing.itemId === item.id)) {
      throw new Error(item.name + " is already on this purchase. Combine the quantities.");
    }
    const unitPrice = item.unitPrice == null || item.unitPrice === "" ? null : Number(item.unitPrice);
    lines.push({
      itemId: item.id,
      itemName: item.name,
      unit: item.unit || "pcs",
      qty,
      unitPrice,
      receivedQty: 0
    });
  });
  const at = nowIso();
  const po = {
    id: newId("cpo"),
    number: nextPoNumber(store.purchases, at),
    createdAt: at,
    status: "Ordered",
    supplier,
    note,
    employee: actorName(actor),
    lines,
    hasPdf: false
  };
  const buffer = await buildPurchasePdf(po);
  fs.writeFileSync(pdfPathFor(po.id), buffer);
  po.hasPdf = true;
  store.purchases.unshift(po);
  lines.forEach((line) => {
    addMovement(store, {
      id: newId("cmove"),
      at,
      type: "purchase_order",
      itemId: line.itemId,
      itemName: line.itemName,
      qty: line.qty,
      orderNumber: "",
      employee: po.employee,
      note: supplier ? ("Ordered from " + supplier) : "Ordered",
      supplier,
      poId: po.id
    });
  });
  saveStore(store);
  return decoratePurchase(po);
}

function getPurchase(id) {
  const want = String(id || "").trim();
  if (!want) return null;
  const store = loadStore();
  const po = store.purchases.find((row) => row.id === want);
  return po ? decoratePurchase(po) : null;
}

function readPurchasePdf(id) {
  const want = String(id || "").trim();
  if (!want) return null;
  const store = loadStore();
  const po = store.purchases.find((row) => row.id === want);
  if (!po) return null;
  const file = pdfPathFor(want);
  if (!fs.existsSync(file)) return null;
  return {
    buffer: fs.readFileSync(file),
    filename: (po.number || want) + ".pdf",
    mime: "application/pdf"
  };
}

function receivePurchase(id, actor) {
  const want = String(id || "").trim();
  if (!want) throw new Error("Purchase not found.");
  const store = loadStore();
  const po = store.purchases.find((row) => row.id === want);
  if (!po) throw new Error("Purchase not found.");
  if (po.status === "Received") throw new Error("That purchase is already received.");
  if (po.status === "Cancelled") throw new Error("That purchase was cancelled.");
  const at = nowIso();
  const employee = actorName(actor);
  (po.lines || []).forEach((line) => {
    const item = findItem(store, line.itemId);
    if (!item) throw new Error(line.itemName + " is no longer on the list.");
    const qty = Number(line.qty) || 0;
    item.stock = Math.round(((Number(item.stock) || 0) + qty) * 1000) / 1000;
    item.updatedAt = at;
    line.receivedQty = qty;
    addMovement(store, {
      id: newId("cmove"),
      at,
      type: "purchase_receive",
      itemId: item.id,
      itemName: item.name,
      qty,
      orderNumber: "",
      employee,
      note: po.supplier ? ("Received from " + po.supplier) : "Received purchase",
      supplier: po.supplier || "",
      poId: po.id
    });
  });
  po.status = "Received";
  po.receivedAt = at;
  po.receivedBy = employee;
  saveStore(store);
  return decoratePurchase(po);
}

function cancelPurchase(id, actor) {
  const want = String(id || "").trim();
  if (!want) throw new Error("Purchase not found.");
  const store = loadStore();
  const po = store.purchases.find((row) => row.id === want);
  if (!po) throw new Error("Purchase not found.");
  if (po.status === "Received") throw new Error("A received purchase cannot be cancelled.");
  if (po.status === "Cancelled") throw new Error("That purchase is already cancelled.");
  po.status = "Cancelled";
  po.cancelledAt = nowIso();
  po.cancelledBy = actorName(actor);
  saveStore(store);
  return decoratePurchase(po);
}

module.exports = {
  UNITS,
  snapshot,
  seedCatalog,
  upsertItem,
  deleteItem,
  logUsage,
  receiveStock,
  countStock,
  createPurchase,
  receivePurchase,
  cancelPurchase,
  getPurchase,
  readPurchasePdf,
  activeOrderNumbers,
  productionWorkers,
  loadStore,
  formatQty
};
