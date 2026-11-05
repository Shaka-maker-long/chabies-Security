"use strict";

const fs = require("fs");
const path = require("path");
const { dataDir } = require("./workbook-store");

function storePath() {
  return path.join(dataDir(), "order-created.json");
}

function emptyStore() {
  return { orders: {} };
}

function formatId(raw) {
  try {
    return require("./db").formatOrderId(raw);
  } catch (e) {
    return String(raw || "").trim();
  }
}

function loadStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(storePath(), "utf8"));
    const store = emptyStore();
    store.orders = parsed && parsed.orders && typeof parsed.orders === "object" ? parsed.orders : {};
    return store;
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[order-created] could not read", storePath(), e.message || e);
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

function record(rows, opts) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) return [];
  const store = loadStore();
  const at = String((opts && opts.at) || "").trim() || new Date().toISOString();
  const by = String((opts && opts.actor) || (opts && opts.by) || "").trim();
  const readyAt = String((opts && opts.readyAt) || "").trim();
  const enquiryNo = String((opts && opts.enquiryNo) || "").trim();
  const out = [];
  list.forEach((row) => {
    const id = formatId(row && (row.order_number || row.orderNumber));
    if (!id) return;
    if (!store.orders[id]) {
      store.orders[id] = {
        at,
        by,
        enquiryNo: String((row && row.enquiry_no) || enquiryNo || "").trim(),
        readyAt: String((row && row.readyAt) || readyAt || "").trim(),
        product: String((row && row.product) || "").trim(),
        client_name: String((row && (row.client_name || row.clientName)) || "").trim()
      };
    }
    out.push(Object.assign({ order_number: id }, store.orders[id]));
  });
  saveStore(store);
  return out;
}

function get(orderNumber) {
  const id = formatId(orderNumber);
  if (!id) return null;
  return loadStore().orders[id] || null;
}

function list() {
  const store = loadStore();
  return Object.keys(store.orders).map((id) => Object.assign({ order_number: id }, store.orders[id]));
}

function renameOrder(fromRaw, toRaw) {
  const fromId = formatId(fromRaw);
  const toId = formatId(toRaw);
  if (!fromId || !toId || fromId === toId) return false;
  const store = loadStore();
  if (!store.orders[fromId]) return false;
  if (!store.orders[toId]) store.orders[toId] = store.orders[fromId];
  delete store.orders[fromId];
  saveStore(store);
  return true;
}

module.exports = {
  record,
  get,
  list,
  renameOrder
};
