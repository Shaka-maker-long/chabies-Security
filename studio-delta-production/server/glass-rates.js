"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir, getBook } = require("./workbook-store");
const { parseMoney, money, formatRand } = require("./db");

const DEFAULT_TYPES = ["Reeded", "Clear", "Ocean view"];
const DEFAULT_THICKNESS = ["4mm", "6mm", "8mm", "10mm"];

function ratesPath() {
  return path.join(dataDir(), "glass-rates.json");
}

function loadRates() {
  try {
    const parsed = JSON.parse(fs.readFileSync(ratesPath(), "utf8"));
    const rates = Array.isArray(parsed.rates) ? parsed.rates : (Array.isArray(parsed) ? parsed : []);
    return { rates: rates.filter((row) => row && row.type && row.thickness) };
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[glass-rates] could not read", ratesPath(), e.message || e);
    }
    return { rates: [] };
  }
}

function saveRates(store) {
  const file = ratesPath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify({ rates: store.rates || [] }));
  fs.renameSync(tmp, file);
  return store;
}

function normalizeType(type) {
  return String(type || "").trim();
}

function typeKey(type) {
  return normalizeType(type).toLowerCase();
}

function normalizeThickness(raw) {
  let s = String(raw || "").trim().toLowerCase().replace(/\s+/g, "");
  if (!s) return "";
  s = s.replace(/millimetres?$/, "mm").replace(/millimeters?$/, "mm");
  if (/^\d+(\.\d+)?$/.test(s)) s += "mm";
  if (/^\d+(\.\d+)?mm$/.test(s)) {
    const n = s.replace("mm", "");
    return n + "mm";
  }
  return String(raw || "").trim();
}

function thicknessKey(raw) {
  return normalizeThickness(raw).toLowerCase();
}

function rateKey(type, thickness) {
  return typeKey(type) + "|" + thicknessKey(thickness);
}

function listGlassTypes() {
  const names = DEFAULT_TYPES.slice();
  try {
    const sheet = getBook().getSheetByName("Glass_Types");
    if (sheet && sheet.getLastRow() >= 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().forEach((row) => {
        const n = String(row[0] || "").trim();
        if (n && names.indexOf(n) === -1) names.push(n);
      });
    }
  } catch (e) {}
  loadRates().rates.forEach((row) => {
    const n = normalizeType(row.type);
    if (n && names.indexOf(n) === -1) names.push(n);
  });
  return names;
}

function listThickness() {
  const names = DEFAULT_THICKNESS.slice();
  loadRates().rates.forEach((row) => {
    const n = normalizeThickness(row.thickness);
    if (n && names.indexOf(n) === -1) names.push(n);
  });
  return names;
}

function findRate(type, thickness) {
  const want = rateKey(type, thickness);
  if (!want || want === "|") return null;
  return loadRates().rates.find((row) => rateKey(row.type, row.thickness) === want) || null;
}

function upsertRate(body) {
  const type = normalizeType(body && body.type);
  const thickness = normalizeThickness(body && body.thickness);
  if (!type) throw new Error("Glass type is required.");
  if (!thickness) throw new Error("Thickness is required.");
  if (body == null || body.ratePerM2 === "" || body.ratePerM2 == null) {
    throw new Error("Rate per square metre is required.");
  }
  const rate = parseMoney(body.ratePerM2);
  if (!Number.isFinite(rate) || rate < 0) throw new Error("Rate per square metre must be 0 or more.");
  const store = loadRates();
  const want = rateKey(type, thickness);
  let row = store.rates.find((item) => rateKey(item.type, item.thickness) === want);
  if (!row) {
    row = { id: "grate_" + crypto.randomBytes(6).toString("hex"), type, thickness };
    store.rates.push(row);
  } else {
    row.type = type;
    row.thickness = thickness;
  }
  row.ratePerM2 = Number(money(rate));
  saveRates(store);
  return row;
}

function deleteRate(id) {
  const want = String(id || "").trim();
  if (!want) throw new Error("Rate not found.");
  const store = loadRates();
  const next = store.rates.filter((row) => row.id !== want);
  if (next.length === store.rates.length) throw new Error("Rate not found.");
  store.rates = next;
  saveRates(store);
  return true;
}

function snapshotRates() {
  const store = loadRates();
  const rates = store.rates.slice().sort((a, b) => {
    const t = String(a.type).localeCompare(String(b.type));
    if (t) return t;
    return String(a.thickness).localeCompare(String(b.thickness), undefined, { numeric: true });
  }).map((row) => ({
    id: row.id,
    type: row.type,
    thickness: row.thickness,
    ratePerM2: money(row.ratePerM2),
    rateLabel: formatRand(row.ratePerM2)
  }));
  return {
    rates,
    types: listGlassTypes(),
    thickness: listThickness()
  };
}

function dimToMetres(value) {
  const n = Number(String(value == null ? "" : value).replace(/,/g, "").replace(/[^\d.-]/g, ""));
  if (!(n > 0)) return 0;
  return n > 20 ? n / 1000 : n;
}

function lineAreaM2(line) {
  const h = dimToMetres(line && line.height);
  const w = dimToMetres(line && line.width);
  const qty = Number(line && line.quantity) || 0;
  if (!(h > 0) || !(w > 0) || !(qty > 0)) return 0;
  return Math.round(h * w * qty * 10000) / 10000;
}

function formatArea(area) {
  if (!(area > 0)) return "";
  return (Math.round(area * 10000) / 10000).toFixed(4).replace(/0+$/, "").replace(/\.$/, "");
}

function costLine(line) {
  const areaM2 = lineAreaM2(line);
  const rate = findRate(line && line.type, line && line.thickness);
  const ratePerM2 = rate ? Number(rate.ratePerM2) : null;
  const estimatedCost = rate && areaM2 > 0 ? Math.round(areaM2 * ratePerM2 * 100) / 100 : null;
  return {
    areaM2,
    areaLabel: areaM2 > 0 ? formatArea(areaM2) + " m²" : "",
    ratePerM2: ratePerM2 == null ? "" : money(ratePerM2),
    rateLabel: rate ? formatRand(ratePerM2) + " / m²" : "",
    estimatedCost: estimatedCost == null ? "" : money(estimatedCost),
    estimatedCostLabel: estimatedCost == null ? "" : formatRand(estimatedCost),
    rateMissing: !rate
  };
}

module.exports = {
  DEFAULT_TYPES,
  DEFAULT_THICKNESS,
  normalizeType,
  normalizeThickness,
  rateKey,
  findRate,
  upsertRate,
  deleteRate,
  snapshotRates,
  listGlassTypes,
  listThickness,
  dimToMetres,
  lineAreaM2,
  costLine
};
