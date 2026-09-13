"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir, getBook } = require("./workbook-store");
const { parseMoney, money, formatRand } = require("./db");

function ratesPath() {
  return path.join(dataDir(), "backboard-rates.json");
}

function loadRates() {
  try {
    const parsed = JSON.parse(fs.readFileSync(ratesPath(), "utf8"));
    const rates = Array.isArray(parsed.rates) ? parsed.rates : (Array.isArray(parsed) ? parsed : []);
    return { rates: rates.filter((row) => row && row.type) };
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[backboard-rates] could not read", ratesPath(), e.message || e);
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

function typeAliases(type) {
  const raw = normalizeType(type);
  const keys = [typeKey(raw)];
  if (raw.indexOf(" - ") !== -1) {
    keys.push(typeKey(raw.split(" - ").slice(-1)[0]));
  }
  return keys.filter(Boolean);
}

function listBackboardTypes() {
  const names = [];
  try {
    const sheet = getBook().getSheetByName("Backboards");
    if (sheet && sheet.getLastRow() >= 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach((row) => {
        const cat = String(row[0] || "").trim();
        const name = String(row[1] || "").trim();
        const label = name ? (cat && cat !== "Uncategorized" ? cat + " - " + name : name) : "";
        if (label && names.indexOf(label) === -1) names.push(label);
        if (name && names.indexOf(name) === -1) names.push(name);
      });
    }
  } catch (e) {}
  try {
    const usage = getBook().getSheetByName("Backboard_Usage");
    if (usage && usage.getLastRow() >= 2) {
      usage.getRange(2, 5, usage.getLastRow() - 1, 1).getValues().forEach((row) => {
        const n = normalizeType(row[0]);
        if (n && names.indexOf(n) === -1) names.push(n);
      });
    }
  } catch (e) {}
  loadRates().rates.forEach((row) => {
    const n = normalizeType(row.type);
    if (n && names.indexOf(n) === -1) names.push(n);
  });
  return names.sort((a, b) => a.localeCompare(b));
}

function findRate(type) {
  const aliases = typeAliases(type);
  if (!aliases.length) return null;
  const rates = loadRates().rates;
  for (let i = 0; i < aliases.length; i++) {
    const hit = rates.find((row) => typeKey(row.type) === aliases[i] || typeAliases(row.type).indexOf(aliases[i]) !== -1);
    if (hit) return hit;
  }
  return null;
}

function upsertRate(body) {
  const type = normalizeType(body && body.type);
  if (!type) throw new Error("Backboard type is required.");
  if (body == null || body.ratePerM2 === "" || body.ratePerM2 == null) {
    throw new Error("Rate per m² is required.");
  }
  const rate = parseMoney(body.ratePerM2);
  if (!Number.isFinite(rate) || rate < 0) throw new Error("Rate per m² must be 0 or more.");
  const store = loadRates();
  const want = typeKey(type);
  let row = store.rates.find((item) => typeKey(item.type) === want);
  if (!row) {
    row = { id: "brate_" + crypto.randomBytes(6).toString("hex"), type };
    store.rates.push(row);
  } else {
    row.type = type;
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
  const rates = store.rates.slice().sort((a, b) => String(a.type).localeCompare(String(b.type))).map((row) => ({
    id: row.id,
    type: row.type,
    ratePerM2: money(row.ratePerM2),
    rateLabel: formatRand(row.ratePerM2) + " / m²"
  }));
  return { rates, types: listBackboardTypes() };
}

function costUsage(type, size) {
  const qty = parseFloat(String(size == null ? "" : size).replace(/,/g, "").replace(/[^\d.-]/g, "")) || 0;
  const rate = findRate(type);
  const ratePerM2 = rate ? Number(rate.ratePerM2) : null;
  const cost = rate && qty > 0 ? Math.round(qty * ratePerM2 * 100) / 100 : null;
  return {
    qty,
    ratePerM2: ratePerM2 == null ? "" : money(ratePerM2),
    cost: cost == null ? 0 : cost,
    rateMissing: !rate || !(qty > 0)
  };
}

module.exports = {
  normalizeType,
  findRate,
  upsertRate,
  deleteRate,
  snapshotRates,
  listBackboardTypes,
  costUsage
};
