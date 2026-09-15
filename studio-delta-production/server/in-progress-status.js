"use strict";

const { getBook } = require("./workbook-store");
const { listOrders, upsertOrder, formatOrderId } = require("./db");
const { waitingStatusIfProfileCuttingIdle, normalizeShopStatus } = require("./shop-status");

function openProfileCuttingOrders() {
  const book = getBook();
  const sheet = book.getSheetByName("Production_Log");
  const open = {};
  if (!sheet || sheet.getLastRow() < 2) return open;
  const grid = sheet.getRange(1, 1, sheet.getLastRow(), 13).getValues();
  for (let i = 1; i < grid.length; i++) {
    if (grid[i][6]) continue;
    const role = String(grid[i][3] || "").trim();
    const process = String(grid[i][4] || "").trim();
    if (role === "Plate Cutting" || role === "Indirect" || role === "Out for Delivery") continue;
    const isCut = /^profile cutting$/i.test(role) || /^profile cutting$/i.test(process);
    if (!isCut) continue;
    const id = formatOrderId(grid[i][1]);
    if (id) open[id] = true;
  }
  return open;
}

function reconcile() {
  const open = openProfileCuttingOrders();
  let updated = 0;
  listOrders().forEach((o) => {
    const id = formatOrderId(o.order_number);
    if (normalizeShopStatus(o.status) !== "Profile Cutting") return;
    const next = waitingStatusIfProfileCuttingIdle(o.status, !!open[id]);
    if (next === "Profile Cutting") return;
    upsertOrder(Object.assign({}, o, { status: next }));
    updated += 1;
  });
  return { updated };
}

module.exports = {
  openProfileCuttingOrders,
  reconcile
};
