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
    priceLabel: unitPrice == null ? "" : formatRand(unitPrice),
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
  const lines = (row.lines || []).map((line) => ({
    itemId: line.itemId,
    itemName: line.itemName,
    qty: Number(line.qty) || 0,
    qtyLabel: formatQty(line.qty),
    receivedQty: Number(line.receivedQty) || 0
  }));
  return {
    id: row.id,
    createdAt: row.createdAt,
    whenLabel: formatWhen(row.createdAt),
    receivedAt: row.receivedAt || "",
    receivedLabel: row.receivedAt ? formatWhen(row.receivedAt) : "",
    status: row.status,
    supplier: row.supplier || "",
    note: row.note || "",
    employee: row.employee || "",
    receivedBy: row.receivedBy || "",
    lines
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

function snapshot() {
  seedCatalog();
  const store = loadStore();
  const items = store.items.slice().sort((a, b) => String(a.name).localeCompare(String(b.name), undefined, { sensitivity: "base" })).map((row) => decorateItem(row, store));
  const low = items.filter((row) => row.low);
  const purchases = store.purchases.slice().sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt))).map(decoratePurchase);
  const openPurchases = purchases.filter((row) => row.status === "Ordered");
  return {
    items,
    low,
    lowCount: low.length,
    itemCount: items.length,
    units: UNITS.slice(),
    movements: store.movements.slice(0, 120).map(decorateMovement),
    purchases,
    openPurchases,
    openPurchaseCount: openPurchases.length,
    orderNumbers: shopOrderNumbers()
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
  const orderNumber = String((body && body.orderNumber) || "").trim();
  const note = String((body && body.note) || "").trim();
  const store = loadStore();
  const row = findItem(store, body && body.itemId);
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
    employee: actorName(actor),
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

function createPurchase(body, actor) {
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
    lines.push({
      itemId: item.id,
      itemName: item.name,
      qty,
      receivedQty: 0
    });
  });
  const at = nowIso();
  const po = {
    id: newId("cpo"),
    createdAt: at,
    status: "Ordered",
    supplier,
    note,
    employee: actorName(actor),
    lines
  };
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
  loadStore,
  formatQty
};
