"use strict";

const fs = require("fs");
const path = require("path");
const { dataDir } = require("./workbook-store");
const { formatOrderId } = require("./db");

function storePath() {
  return path.join(dataDir(), "no-plates.json");
}

function emptyStore() {
  return { orders: {} };
}

function loadStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(storePath(), "utf8"));
    const store = emptyStore();
    store.orders = parsed && parsed.orders && typeof parsed.orders === "object" ? parsed.orders : {};
    return store;
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[no-plates] could not read", storePath(), e.message || e);
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

function isNoPlate(orderNumber) {
  const id = formatOrderId(orderNumber);
  if (!id) return false;
  const rec = loadStore().orders[id];
  return !!(rec && rec.noPlate);
}

function markNoPlate(orderNumber, actor) {
  const id = formatOrderId(orderNumber);
  if (!id) throw new Error("Order number is required.");
  const store = loadStore();
  store.orders[id] = {
    noPlate: true,
    markedBy: String(actor || "").trim(),
    markedAt: new Date().toISOString()
  };
  saveStore(store);
  try { require("./gas").clearShopCache(); } catch (e) {}
  return store.orders[id];
}

function clearNoPlate(orderNumber) {
  const id = formatOrderId(orderNumber);
  if (!id) throw new Error("Order number is required.");
  const store = loadStore();
  delete store.orders[id];
  saveStore(store);
  try { require("./gas").clearShopCache(); } catch (e) {}
  return { noPlate: false };
}

function listNoPlates() {
  const orders = loadStore().orders;
  return Object.keys(orders).filter((id) => orders[id] && orders[id].noPlate);
}

function canMarkNoPlate(profile) {
  if (!profile) return false;
  const title = String(profile.jobTitle || profile.role || "").trim().toLowerCase();
  if (title === "manager" || title === "production manager") return true;
  if (profile.canManageUsers) return true;
  const tasks = Array.isArray(profile.tasks) ? profile.tasks : [];
  return tasks.indexOf("Plate Cutting") !== -1;
}

module.exports = {
  isNoPlate,
  markNoPlate,
  clearNoPlate,
  listNoPlates,
  canMarkNoPlate
};
