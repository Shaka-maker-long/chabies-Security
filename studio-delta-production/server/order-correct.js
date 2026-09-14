"use strict";

const { getBook, persistWorkbook } = require("./workbook-store");
const {
  listOrders,
  upsertOrder,
  formatOrderId
} = require("./db");
const {
  SHOP_STATUSES,
  isShopStatus,
  normalizeShopStatus,
  waitingStatusIfProfileCuttingIdle
} = require("./shop-status");

function closeOpenLogsForOrder(orderNumber) {
  const id = formatOrderId(orderNumber);
  if (!id) return { closed: 0 };
  const now = new Date();
  let closed = 0;
  [
    { title: "Production_Log", endCol: 7 },
    { title: "Overview", endCol: 6 }
  ].forEach((spec) => {
    const sheet = getBook().getSheetByName(spec.title);
    if (!sheet || sheet.getLastRow() < 2) return;
    const lastCol = Math.max(sheet.getLastColumn(), spec.endCol);
    const grid = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    const endIdx = spec.endCol - 1;
    for (let i = 0; i < grid.length; i++) {
      if (formatOrderId(grid[i][1]) !== id) continue;
      if (grid[i][endIdx]) continue;
      sheet.getRange(i + 2, spec.endCol).setValue(now);
      closed += 1;
    }
  });
  if (closed) persistWorkbook();
  return { closed };
}

function correctOrderShop(body) {
  const row = body || {};
  if (!row.confirmed) {
    throw new Error("Tick that this is where the order is now.");
  }
  const orderNumber = formatOrderId(row.order_number || row.orderNumber);
  if (!orderNumber) throw new Error("Order number is required.");
  const existing = listOrders().find((o) => formatOrderId(o.order_number) === orderNumber);
  if (!existing) throw new Error(orderNumber + " is not on Orders.");
  let status = normalizeShopStatus(row.status);
  if (!isShopStatus(status)) throw new Error("Pick where this order is on the floor now.");
  const assigned = String(row.assigned_operator != null ? row.assigned_operator : row.assignedOperator || "").trim();
  const logs = closeOpenLogsForOrder(orderNumber);
  let openProfile = {};
  try { openProfile = require("./in-progress-status").openProfileCuttingOrders(); } catch (e) {}
  status = waitingStatusIfProfileCuttingIdle(status, !!openProfile[orderNumber]);
  const saved = upsertOrder(Object.assign({}, existing, {
    status,
    assigned_operator: assigned
  }));
  let planningRemoved = 0;
  try {
    planningRemoved = require("./floor-planning").unscheduleOrder(orderNumber).removed || 0;
  } catch (e) {}
  try { require("./gas").clearShopCache(); } catch (e) {}
  return {
    row: saved,
    closedLogs: logs.closed,
    planningRemoved,
    status
  };
}

module.exports = {
  closeOpenLogsForOrder,
  correctOrderShop,
  SHOP_STATUSES
};
