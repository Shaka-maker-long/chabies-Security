"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir, getBook } = require("./workbook-store");
const { parseMoney, money, formatRand, listOrders, formatOrderId } = require("./db");
const { normalizeShopStatus } = require("./shop-status");
const steelRates = require("./steel-rates");
const glassRates = require("./glass-rates");

/** Tubes / bars / angle iron are bought as fixed lengths. */
const STEEL_LENGTH_M = 6;

/**
 * Auto-deduct (Current → WIP on allocate; WIP → used on deliver) stays off until
 * on-hand counts are loaded. Helpers are ready; callers must check this flag.
 */
const STEEL_AUTO_DEDUCT_ENABLED = false;

function steelPath() {
  return path.join(dataDir(), "inventory-steel.json");
}

function glassPath() {
  return path.join(dataDir(), "inventory-glass.json");
}

function loadKind(file) {
  try {
    const parsed = JSON.parse(fs.readFileSync(file, "utf8"));
    const items = Array.isArray(parsed.items) ? parsed.items.filter((row) => row && row.name) : [];
    return { items };
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[inventory] could not read", file, e.message || e);
    }
    return { items: [] };
  }
}

function saveKind(file, store) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify({ items: store.items || [] }));
  fs.renameSync(tmp, file);
  return store;
}

function nameKey(name) {
  return String(name || "").replace(/\s+/g, " ").trim().toLowerCase();
}

function formatQty(n) {
  const v = Number(n);
  if (!Number.isFinite(v)) return "0";
  const s = (Math.round(v * 1000) / 1000).toFixed(3);
  return s.replace(/\.?0+$/, "");
}

function parseStock(raw) {
  const s = String(raw == null ? "" : raw).replace(/,/g, "").trim();
  if (s === "") return 0;
  const n = Number(s);
  if (!Number.isFinite(n) || n < 0) throw new Error("Stock must be 0 or more.");
  return Math.round(n * 1000) / 1000;
}

function parseThreshold(raw) {
  const s = String(raw == null ? "" : raw).replace(/,/g, "").trim();
  if (s === "") return 0;
  const n = Number(s);
  if (!Number.isFinite(n) || n < 0) throw new Error("ROP must be 0 or more.");
  return Math.round(n * 1000) / 1000;
}

function parsePrice(raw) {
  const s = String(raw == null ? "" : raw).trim();
  if (s === "") return null;
  const n = parseMoney(s);
  if (!Number.isFinite(n) || n < 0) throw new Error("Unit price must be 0 or more.");
  return Number(money(n));
}

function extraByName(store) {
  const map = {};
  (store.items || []).forEach((row) => {
    map[nameKey(row.name)] = row;
  });
  return map;
}

function isPlateName(name) {
  const s = String(name || "");
  return /\bplate\b/i.test(s) || /^\s*LATE\b/i.test(s);
}

function parsePlateSheetMm(name) {
  const m = String(name || "").match(/(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*[xX×]\s*(\d+(?:\.\d+)?)/);
  if (!m) return null;
  return {
    thicknessMm: Number(m[1]),
    widthMm: Number(m[2]),
    lengthMm: Number(m[3])
  };
}

function plateSheetAreaM2(name) {
  const d = parsePlateSheetMm(name);
  if (!d || !(d.widthMm > 0) || !(d.lengthMm > 0)) return null;
  return Math.round((d.widthMm / 1000) * (d.lengthMm / 1000) * 10000) / 10000;
}

function steelBuyUnit(name) {
  return isPlateName(name) ? "sheet" : "length";
}

function resolveBuyUnit(name, extra) {
  const raw = extra && extra.buyUnit != null ? String(extra.buyUnit).trim().toLowerCase() : "";
  if (raw === "sheet" || raw === "length") return raw;
  return steelBuyUnit(name);
}

function buyUnitLabelFor(buyUnit) {
  return buyUnit === "sheet" ? "sheet" : "length (6 m)";
}

function steelBuyUnitLabel(name, extra) {
  return buyUnitLabelFor(resolveBuyUnit(name, extra));
}

function parseBuyUnit(raw) {
  const s = String(raw == null ? "" : raw).trim().toLowerCase();
  if (!s) return null;
  if (s === "sheet" || s.indexOf("sheet") !== -1) return "sheet";
  if (s === "length" || s.indexOf("length") !== -1 || s.indexOf("6 m") !== -1 || s.indexOf("6m") !== -1) return "length";
  throw new Error("Buy unit must be length or sheet.");
}

function steelTypeKeys(type) {
  const raw = String(type || "").replace(/\s+/g, " ").trim();
  if (!raw) return [];
  const keys = [nameKey(raw)];
  if (raw.indexOf(" - ") !== -1) keys.push(nameKey(raw.split(" - ").slice(-1)[0]));
  return keys.filter(Boolean);
}

function steelTypesMatch(a, b) {
  const ak = steelTypeKeys(a);
  const bk = steelTypeKeys(b);
  return ak.some((k) => bk.indexOf(k) !== -1);
}

function allocationQty(size) {
  return parseFloat(String(size == null ? "" : size).replace(/,/g, "").replace(/[^\d.-]/g, "")) || 0;
}

function readSteelUsageRows() {
  try {
    const sheet = getBook().getSheetByName("Steel_Usage");
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues().map((row) => ({
      orderNum: String(row[1] || "").trim(),
      type: String(row[4] || "").trim(),
      size: row[5]
    })).filter((row) => row.orderNum && row.type);
  } catch (e) {
    return [];
  }
}

function orderDeliveredById() {
  const map = {};
  try {
    listOrders().forEach((order) => {
      const id = formatOrderId(order.order_number) || String(order.order_number || "").trim();
      if (!id) return;
      const s = normalizeShopStatus(order.status);
      map[id] = s === "Delivered" || s === "At couriers";
    });
  } catch (e) {}
  return map;
}

/**
 * Allocated metres (or m² for plates) from Steel_Usage, split by order status.
 * Not Delivered → WIP; Delivered → used.
 */
function steelUsageAllocByName(names) {
  const buckets = {};
  (names || []).forEach((name) => {
    buckets[nameKey(name)] = { name, wipAlloc: 0, usedAlloc: 0 };
  });
  const delivered = orderDeliveredById();
  readSteelUsageRows().forEach((row) => {
    const qty = allocationQty(row.size);
    if (!(qty > 0)) return;
    const orderId = formatOrderId(row.orderNum) || row.orderNum;
    const isDelivered = !!delivered[orderId];
    let hitKey = null;
    for (let i = 0; i < (names || []).length; i++) {
      if (steelTypesMatch(row.type, names[i])) {
        hitKey = nameKey(names[i]);
        break;
      }
    }
    if (!hitKey) {
      hitKey = nameKey(row.type);
      if (!buckets[hitKey]) buckets[hitKey] = { name: row.type, wipAlloc: 0, usedAlloc: 0 };
    }
    if (isDelivered) buckets[hitKey].usedAlloc = Math.round((buckets[hitKey].usedAlloc + qty) * 1000) / 1000;
    else buckets[hitKey].wipAlloc = Math.round((buckets[hitKey].wipAlloc + qty) * 1000) / 1000;
  });
  return buckets;
}

function allocToBuyUnits(allocQty, buyUnit, sheetAreaM2) {
  const qty = Number(allocQty) || 0;
  if (!(qty > 0)) return 0;
  if (buyUnit === "sheet") return m2ToSheets(qty, sheetAreaM2);
  return metresToLengths(qty);
}

/** Metres allocated → lengths to deduct (6 m = 1 length). */
function metresToLengths(metres) {
  const m = Number(metres) || 0;
  if (!(m > 0)) return 0;
  return Math.round((m / STEEL_LENGTH_M) * 1000) / 1000;
}

/** m² allocated → sheets (by sheet area from the name). */
function m2ToSheets(m2, sheetAreaM2) {
  const area = Number(sheetAreaM2) || 0;
  const used = Number(m2) || 0;
  if (!(area > 0) || !(used > 0)) return 0;
  return Math.round((used / area) * 1000) / 1000;
}

function defaultLengthUnitPrice(ratePerM) {
  if (ratePerM == null || ratePerM === "") return null;
  const n = Number(ratePerM);
  if (!Number.isFinite(n) || n < 0) return null;
  return Math.round(n * STEEL_LENGTH_M * 100) / 100;
}

function decorateGlass(name, extra, priceFromRate, orderedFromOrders) {
  const stock = extra && extra.stock != null ? Number(extra.stock) || 0 : 0;
  const minThreshold = extra && extra.minThreshold != null ? Number(extra.minThreshold) || 0 : 0;
  const orderedQty = extra && extra.orderedQty != null && String(extra.orderedQty).trim() !== ""
    ? Number(extra.orderedQty) || 0
    : (orderedFromOrders || 0);
  const unitPrice = extra && extra.unitPrice != null && extra.unitPrice !== ""
    ? Number(extra.unitPrice)
    : (priceFromRate == null ? null : Number(priceFromRate));
  const totalValue = unitPrice == null ? null : Math.round(stock * unitPrice * 100) / 100;
  const low = minThreshold > 0 && stock <= minThreshold;
  return {
    id: extra && extra.id ? extra.id : "",
    name,
    stock,
    stockLabel: formatQty(stock),
    orderedQty,
    orderedLabel: formatQty(orderedQty),
    minThreshold,
    ropLabel: formatQty(minThreshold),
    unitPrice,
    priceLabel: unitPrice == null ? "—" : formatRand(unitPrice),
    totalValue,
    totalValueLabel: totalValue == null ? "—" : formatRand(totalValue),
    low,
    status: low ? "Low" : "OK"
  };
}

function decorateSteel(name, extra, ratePerM, usageAlloc) {
  const buyUnit = resolveBuyUnit(name, extra);
  const sheetArea = buyUnit === "sheet" ? plateSheetAreaM2(name) : null;
  const stock = extra && extra.stock != null ? Number(extra.stock) || 0 : 0;
  const orderedQty = extra && extra.orderedQty != null ? Number(extra.orderedQty) || 0 : 0;
  const totalPurchased = extra && extra.totalPurchased != null ? Number(extra.totalPurchased) || 0 : 0;
  const minThreshold = extra && extra.minThreshold != null ? Number(extra.minThreshold) || 0 : 0;
  const wipAlloc = usageAlloc && usageAlloc.wipAlloc != null ? Number(usageAlloc.wipAlloc) || 0 : 0;
  const usedAlloc = usageAlloc && usageAlloc.usedAlloc != null ? Number(usageAlloc.usedAlloc) || 0 : 0;
  const wipStock = allocToBuyUnits(wipAlloc, buyUnit, sheetArea);
  const usedFromUsage = allocToBuyUnits(usedAlloc, buyUnit, sheetArea);
  const hasStoredUsed = !!(extra && Object.prototype.hasOwnProperty.call(extra, "totalUsed") && extra.totalUsed != null && String(extra.totalUsed).trim() !== "");
  const totalUsed = hasStoredUsed ? Number(extra.totalUsed) || 0 : usedFromUsage;
  const defaultPrice = buyUnit === "length" ? defaultLengthUnitPrice(ratePerM) : null;
  const unitPrice = extra && extra.unitPrice != null && extra.unitPrice !== ""
    ? Number(extra.unitPrice)
    : defaultPrice;
  const valueInStock = unitPrice == null ? null : Math.round(stock * unitPrice * 100) / 100;
  const valueInWip = unitPrice == null ? null : Math.round(wipStock * unitPrice * 100) / 100;
  const low = minThreshold > 0 && stock <= minThreshold;
  const rateHint = buyUnit === "length" && ratePerM != null
    ? formatRand(ratePerM) + "/m → " + (unitPrice == null ? "—" : formatRand(unitPrice)) + "/length"
    : (buyUnit === "sheet" && sheetArea != null ? sheetArea + " m² / sheet" : "");
  return {
    id: extra && extra.id ? extra.id : "",
    name,
    buyUnit,
    buyUnitLabel: buyUnitLabelFor(buyUnit),
    lengthM: buyUnit === "length" ? STEEL_LENGTH_M : null,
    sheetAreaM2: sheetArea,
    stock,
    stockLabel: formatQty(stock),
    wipStock,
    wipQty: wipStock,
    wipLabel: formatQty(wipStock),
    wipAlloc,
    wipAllocLabel: formatQty(wipAlloc),
    orderedQty,
    orderedLabel: formatQty(orderedQty),
    totalPurchased,
    purchasedLabel: formatQty(totalPurchased),
    totalUsed,
    usedFromUsage,
    usedLabel: formatQty(totalUsed),
    usedFromUsageLabel: formatQty(usedFromUsage),
    minThreshold,
    ropLabel: formatQty(minThreshold),
    unitPrice,
    priceLabel: unitPrice == null ? "—" : formatRand(unitPrice),
    ratePerM: ratePerM == null ? null : Number(ratePerM),
    rateHint,
    valueInStock,
    valueInStockLabel: valueInStock == null ? "—" : formatRand(valueInStock),
    valueInWip,
    valueInWipLabel: valueInWip == null ? "—" : formatRand(valueInWip),
    totalValue: valueInStock,
    totalValueLabel: valueInStock == null ? "—" : formatRand(valueInStock),
    low,
    status: low ? "Low" : "OK",
    autoDeductEnabled: STEEL_AUTO_DEDUCT_ENABLED
  };
}

function steelProfileNames() {
  const names = [];
  const seen = {};
  function add(label) {
    const n = String(label || "").replace(/\s+/g, " ").trim();
    const k = nameKey(n);
    if (!k || seen[k]) return;
    seen[k] = true;
    names.push(n);
  }
  try {
    const sheet = getBook().getSheetByName("Steel_Profiles");
    if (sheet && sheet.getLastRow() >= 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach((row) => {
        const cat = String(row[0] || "").trim();
        const name = String(row[1] || "").trim();
        if (!name) return;
        add(cat && cat !== "Uncategorized" ? cat + " - " + name : name);
      });
    }
  } catch (e) {}
  (steelRates.snapshotRates().rates || []).forEach((row) => add(row.type));
  (loadKind(steelPath()).items || []).forEach((row) => add(row.name));
  readSteelUsageRows().forEach((row) => add(row.type));
  return names.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}

function snapshotSteel() {
  const store = loadKind(steelPath());
  const extras = extraByName(store);
  const names = steelProfileNames();
  const usageByName = steelUsageAllocByName(names);
  const items = names.map((name) => {
    const rate = steelRates.findRate(name);
    return decorateSteel(
      name,
      extras[nameKey(name)],
      rate ? rate.ratePerM : null,
      usageByName[nameKey(name)]
    );
  });
  return {
    items,
    itemCount: items.length,
    lowCount: items.filter((row) => row.low).length,
    lengthM: STEEL_LENGTH_M,
    autoDeductEnabled: STEEL_AUTO_DEDUCT_ENABLED
  };
}

function upsertSteel(body) {
  const name = String((body && body.name) || "").replace(/\s+/g, " ").trim();
  if (!name) throw new Error("Steel name is required.");
  const store = loadKind(steelPath());
  const prevName = String((body && body.previousName) || "").replace(/\s+/g, " ").trim();
  const wantPrev = nameKey(prevName || name);
  let row = store.items.find((item) => nameKey(item.name) === wantPrev)
    || store.items.find((item) => nameKey(item.name) === nameKey(name));
  if (!row) {
    row = { id: "sinv_" + crypto.randomBytes(6).toString("hex"), name };
    store.items.push(row);
  }
  row.name = name;
  if (body && Object.prototype.hasOwnProperty.call(body, "stock")) row.stock = parseStock(body.stock);
  if (body && Object.prototype.hasOwnProperty.call(body, "minThreshold")) row.minThreshold = parseThreshold(body.minThreshold);
  if (body && Object.prototype.hasOwnProperty.call(body, "orderedQty")) row.orderedQty = parseStock(body.orderedQty);
  if (body && Object.prototype.hasOwnProperty.call(body, "totalPurchased")) row.totalPurchased = parseStock(body.totalPurchased);
  if (body && Object.prototype.hasOwnProperty.call(body, "totalUsed")) row.totalUsed = parseStock(body.totalUsed);
  if (body && Object.prototype.hasOwnProperty.call(body, "unitPrice")) row.unitPrice = parsePrice(body.unitPrice);
  if (body && Object.prototype.hasOwnProperty.call(body, "buyUnit")) {
    const unit = parseBuyUnit(body.buyUnit);
    if (unit) row.buyUnit = unit;
  }
  saveKind(steelPath(), store);
  return snapshotSteel();
}

function findOrCreateSteelRow(store, name) {
  const want = nameKey(name);
  let row = store.items.find((item) => nameKey(item.name) === want);
  if (!row) {
    row = { id: "sinv_" + crypto.randomBytes(6).toString("hex"), name: String(name || "").replace(/\s+/g, " ").trim() };
    store.items.push(row);
  }
  return row;
}

/**
 * Move buy-units from Current → WIP when steel is allocated on a job.
 * Call with lengths/sheets (not metres). Use metresToLengths / m2ToSheets at the call site.
 * Disabled until STEEL_AUTO_DEDUCT_ENABLED is true.
 */
function allocateSteelToWip(name, qtyBuyUnits) {
  if (!STEEL_AUTO_DEDUCT_ENABLED) {
    return { ok: false, skipped: true, reason: "Steel auto-deduct is off until on-hand stock is loaded." };
  }
  const qty = parseStock(qtyBuyUnits);
  if (!(qty > 0)) return { ok: true, stock: 0, wipStock: 0, moved: 0 };
  const store = loadKind(steelPath());
  const row = findOrCreateSteelRow(store, name);
  const stock = Number(row.stock) || 0;
  const wip = Number(row.wipStock) || 0;
  if (stock < qty) {
    throw new Error("Not enough current stock to allocate " + formatQty(qty) + " " + steelBuyUnitLabel(name) + "(s) of " + name + ".");
  }
  row.stock = Math.round((stock - qty) * 1000) / 1000;
  row.wipStock = Math.round((wip + qty) * 1000) / 1000;
  saveKind(steelPath(), store);
  return { ok: true, stock: row.stock, wipStock: row.wipStock, moved: qty };
}

/**
 * When an order is Delivered, move that job's WIP qty into totalUsed.
 * Call with lengths/sheets (not metres). Disabled until STEEL_AUTO_DEDUCT_ENABLED is true.
 */
function consumeSteelFromWipOnDeliver(name, qtyBuyUnits) {
  if (!STEEL_AUTO_DEDUCT_ENABLED) {
    return { ok: false, skipped: true, reason: "Steel auto-deduct is off until on-hand stock is loaded." };
  }
  const qty = parseStock(qtyBuyUnits);
  if (!(qty > 0)) return { ok: true, wipStock: 0, totalUsed: 0, moved: 0 };
  const store = loadKind(steelPath());
  const row = findOrCreateSteelRow(store, name);
  const wip = Number(row.wipStock) || 0;
  const used = Number(row.totalUsed) || 0;
  const move = Math.min(wip, qty);
  row.wipStock = Math.round((wip - move) * 1000) / 1000;
  row.totalUsed = Math.round((used + move) * 1000) / 1000;
  saveKind(steelPath(), store);
  return { ok: true, wipStock: row.wipStock, totalUsed: row.totalUsed, moved: move };
}

function glassLineName(type, thickness) {
  const t = glassRates.normalizeType(type);
  const th = glassRates.normalizeThickness(thickness);
  return [t, th].filter(Boolean).join(" ");
}

function glassOrderedMap() {
  const map = {};
  try {
    const snap = require("./glass-po").snapshot();
    (snap.toOrder || []).concat(snap.outstanding || []).forEach((line) => {
      const name = glassLineName(line.type, line.thickness);
      if (!name) return;
      const k = nameKey(name);
      const qty = Number(line.quantity) || 0;
      if (!map[k]) map[k] = { name, qty: 0 };
      map[k].qty += qty;
    });
  } catch (e) {}
  return map;
}

function snapshotGlass() {
  const store = loadKind(glassPath());
  const extras = extraByName(store);
  const orderedMap = glassOrderedMap();
  const names = [];
  const seen = {};
  function add(label) {
    const n = String(label || "").replace(/\s+/g, " ").trim();
    const k = nameKey(n);
    if (!k || seen[k]) return;
    seen[k] = true;
    names.push(n);
  }
  (glassRates.snapshotRates().rates || []).forEach((row) => add(glassLineName(row.type, row.thickness)));
  Object.keys(orderedMap).forEach((k) => add(orderedMap[k].name));
  (store.items || []).forEach((row) => add(row.name));
  names.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
  const items = names.map((name) => {
    const parts = String(name).trim().split(/\s+/);
    const thickness = parts.length ? parts[parts.length - 1] : "";
    const type = parts.slice(0, -1).join(" ") || name;
    const rate = glassRates.findRate(type, thickness);
    const extra = extras[nameKey(name)];
    const computed = orderedMap[nameKey(name)] ? orderedMap[nameKey(name)].qty : 0;
    const forDecorate = extra ? Object.assign({}, extra) : extra;
    if (forDecorate && (forDecorate.orderedQty == null || String(forDecorate.orderedQty).trim() === "")) {
      delete forDecorate.orderedQty;
    }
    return decorateGlass(name, forDecorate, rate ? rate.ratePerM2 : null, computed);
  });
  return { items, itemCount: items.length, lowCount: items.filter((row) => row.low).length };
}

function upsertGlass(body) {
  const name = String((body && body.name) || "").replace(/\s+/g, " ").trim();
  if (!name) throw new Error("Glass name is required.");
  const store = loadKind(glassPath());
  const want = nameKey(name);
  let row = store.items.find((item) => nameKey(item.name) === want);
  if (!row) {
    row = { id: "ginv_" + crypto.randomBytes(6).toString("hex"), name };
    store.items.push(row);
  } else {
    row.name = name;
  }
  if (body && Object.prototype.hasOwnProperty.call(body, "stock")) row.stock = parseStock(body.stock);
  if (body && Object.prototype.hasOwnProperty.call(body, "minThreshold")) row.minThreshold = parseThreshold(body.minThreshold);
  if (body && Object.prototype.hasOwnProperty.call(body, "orderedQty")) row.orderedQty = parseStock(body.orderedQty);
  if (body && Object.prototype.hasOwnProperty.call(body, "unitPrice")) row.unitPrice = parsePrice(body.unitPrice);
  saveKind(glassPath(), store);
  return snapshotGlass();
}

module.exports = {
  STEEL_LENGTH_M,
  STEEL_AUTO_DEDUCT_ENABLED,
  isPlateName,
  parsePlateSheetMm,
  plateSheetAreaM2,
  steelBuyUnit,
  resolveBuyUnit,
  steelTypesMatch,
  metresToLengths,
  m2ToSheets,
  defaultLengthUnitPrice,
  snapshotSteel,
  upsertSteel,
  allocateSteelToWip,
  consumeSteelFromWipOnDeliver,
  snapshotGlass,
  upsertGlass
};
