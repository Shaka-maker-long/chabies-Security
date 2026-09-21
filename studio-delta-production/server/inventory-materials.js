"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir, getBook } = require("./workbook-store");
const { parseMoney, money, formatRand } = require("./db");
const steelRates = require("./steel-rates");
const glassRates = require("./glass-rates");

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

function decorate(name, extra, priceFromRate, orderedFromOrders) {
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
  return names.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}

function snapshotSteel() {
  const store = loadKind(steelPath());
  const extras = extraByName(store);
  const items = steelProfileNames().map((name) => {
    const rate = steelRates.findRate(name);
    return decorate(name, extras[nameKey(name)], rate ? rate.ratePerM : null, 0);
  });
  return { items, itemCount: items.length, lowCount: items.filter((row) => row.low).length };
}

function upsertSteel(body) {
  const name = String((body && body.name) || "").replace(/\s+/g, " ").trim();
  if (!name) throw new Error("Steel name is required.");
  const store = loadKind(steelPath());
  const want = nameKey(name);
  let row = store.items.find((item) => nameKey(item.name) === want);
  if (!row) {
    row = { id: "sinv_" + crypto.randomBytes(6).toString("hex"), name };
    store.items.push(row);
  } else {
    row.name = name;
  }
  if (body && Object.prototype.hasOwnProperty.call(body, "stock")) row.stock = parseStock(body.stock);
  if (body && Object.prototype.hasOwnProperty.call(body, "minThreshold")) row.minThreshold = parseThreshold(body.minThreshold);
  if (body && Object.prototype.hasOwnProperty.call(body, "orderedQty")) row.orderedQty = parseStock(body.orderedQty);
  if (body && Object.prototype.hasOwnProperty.call(body, "unitPrice")) row.unitPrice = parsePrice(body.unitPrice);
  saveKind(steelPath(), store);
  return snapshotSteel();
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
    return decorate(name, forDecorate, rate ? rate.ratePerM2 : null, computed);
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
  snapshotSteel,
  upsertSteel,
  snapshotGlass,
  upsertGlass
};
