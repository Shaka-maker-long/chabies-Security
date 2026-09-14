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

const SKIP_ROLES = { "plate cutting": true, "indirect": true, "out for delivery": true };

function liveTaskForStatus(status) {
  const s = normalizeShopStatus(status);
  if (s === "Profile Cutting") return "Profile Cutting";
  if (s === "Tagging") return "Tagging";
  if (s === "Welding") return "Welding";
  if (s === "Grinding") return "Grinding";
  if (s === "Pre-Powder Coating") return "Quality Control";
  if (s === "Assembly") return "Assembly";
  if (s === "Paint Preparation") return "Paint Preparation";
  if (s === "Painting") return "Painting";
  if (s === "Final QC") return "Quality Control";
  return "";
}

function parseMeta(raw) {
  try {
    const parsed = JSON.parse(String(raw || ""));
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (e) {
    return {};
  }
}

function closeOpenPauses(meta) {
  const now = Date.now();
  const pauses = Array.isArray(meta.pauses) ? meta.pauses : [];
  pauses.forEach((p) => {
    if (p && !p.end) p.end = now;
  });
  meta.pauses = pauses;
  return meta;
}

function isSkippedRole(role) {
  return !!SKIP_ROLES[String(role || "").trim().toLowerCase()];
}

function runningByOrder() {
  const sheet = getBook().getSheetByName("Production_Log");
  const out = {};
  if (!sheet || sheet.getLastRow() < 2) return out;
  const lastCol = Math.max(sheet.getLastColumn(), 13);
  const grid = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  for (let i = 0; i < grid.length; i++) {
    if (grid[i][6]) continue;
    const id = formatOrderId(grid[i][1]);
    if (!id) continue;
    const role = String(grid[i][3] || "").trim();
    if (isSkippedRole(role)) continue;
    const meta = parseMeta(grid[i][12]);
    const paused = Array.isArray(meta.pauses) && meta.pauses.some((p) => p && !p.end);
    out[id] = {
      worker: String(grid[i][2] || "").trim(),
      task: role || String(grid[i][4] || "").trim(),
      paused
    };
  }
  return out;
}

function annotateOrders(rows) {
  const running = runningByOrder();
  return (rows || []).map((row) => {
    const id = formatOrderId(row && row.order_number);
    const live = running[id];
    return Object.assign({}, row, {
      running: !!live,
      running_worker: live ? live.worker : "",
      running_task: live ? live.task : "",
      running_paused: !!(live && live.paused)
    });
  });
}

function closeOpenLogsForOrder(orderNumber) {
  const id = formatOrderId(orderNumber);
  if (!id) return { closed: 0 };
  const now = new Date();
  let closed = 0;
  [
    { title: "Production_Log", endCol: 7, roleIdx: 3 },
    { title: "Overview", endCol: 6, roleIdx: -1 }
  ].forEach((spec) => {
    const sheet = getBook().getSheetByName(spec.title);
    if (!sheet || sheet.getLastRow() < 2) return;
    const lastCol = Math.max(sheet.getLastColumn(), spec.endCol);
    const grid = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    const endIdx = spec.endCol - 1;
    for (let i = 0; i < grid.length; i++) {
      if (formatOrderId(grid[i][1]) !== id) continue;
      if (grid[i][endIdx]) continue;
      if (spec.roleIdx >= 0 && isSkippedRole(grid[i][spec.roleIdx])) continue;
      sheet.getRange(i + 2, spec.endCol).setValue(now);
      closed += 1;
    }
  });
  if (closed) persistWorkbook();
  return { closed };
}

function transferOpenLogsForOrder(orderNumber, worker, process, status) {
  const id = formatOrderId(orderNumber);
  if (!id) return { transferred: 0 };
  let transferred = 0;
  const logSheet = getBook().getSheetByName("Production_Log");
  if (logSheet && logSheet.getLastRow() >= 2) {
    const lastCol = Math.max(logSheet.getLastColumn(), 13);
    const grid = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, lastCol).getValues();
    for (let i = 0; i < grid.length; i++) {
      if (formatOrderId(grid[i][1]) !== id) continue;
      if (grid[i][6]) continue;
      if (isSkippedRole(grid[i][3])) continue;
      const nextWorker = worker || String(grid[i][2] || "").trim();
      const meta = closeOpenPauses(parseMeta(grid[i][12]));
      logSheet.getRange(i + 2, 3, 1, 3).setValues([[nextWorker, process, status]]);
      logSheet.getRange(i + 2, 10, 1, 4).setValues([["", "", "", JSON.stringify(meta)]]);
      transferred += 1;
    }
  }
  const overview = getBook().getSheetByName("Overview");
  if (overview && overview.getLastRow() >= 2) {
    const lastCol = Math.max(overview.getLastColumn(), 6);
    const grid = overview.getRange(2, 1, overview.getLastRow() - 1, lastCol).getValues();
    for (let i = 0; i < grid.length; i++) {
      if (formatOrderId(grid[i][1]) !== id) continue;
      if (grid[i][5]) continue;
      const nextWorker = worker || String(grid[i][2] || "").trim();
      overview.getRange(i + 2, 3, 1, 2).setValues([[nextWorker, status]]);
    }
  }
  if (transferred) persistWorkbook();
  return { transferred };
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
  const task = liveTaskForStatus(status);
  let closedLogs = 0;
  let transferred = 0;
  if (task) {
    transferred = transferOpenLogsForOrder(orderNumber, assigned, task, status).transferred;
  } else {
    closedLogs = closeOpenLogsForOrder(orderNumber).closed;
  }
  let openProfile = {};
  try { openProfile = require("./in-progress-status").openProfileCuttingOrders(); } catch (e) {}
  status = waitingStatusIfProfileCuttingIdle(status, !!openProfile[orderNumber]);
  const saved = upsertOrder(Object.assign({}, existing, {
    status,
    assigned_operator: assigned
  }));
  let planningRemoved = 0;
  const statusChanged = normalizeShopStatus(existing.status) !== status;
  if (statusChanged) {
    try {
      planningRemoved = require("./floor-planning").unscheduleOrder(orderNumber).removed || 0;
    } catch (e) {}
  }
  try { require("./gas").clearShopCache(); } catch (e) {}
  return {
    row: saved,
    closedLogs,
    transferred,
    planningRemoved,
    status
  };
}

function correctManyOrders(body) {
  const pack = body || {};
  if (!pack.confirmed) throw new Error("Tick that these are the right status and person.");
  const rows = Array.isArray(pack.rows) ? pack.rows : [];
  if (!rows.length) throw new Error("Pick at least one order to correct.");
  const updated = [];
  let closedLogs = 0;
  let transferred = 0;
  let planningRemoved = 0;
  rows.forEach((row) => {
    const result = correctOrderShop(Object.assign({}, row, { confirmed: true }));
    updated.push(result.row);
    closedLogs += result.closedLogs || 0;
    transferred += result.transferred || 0;
    planningRemoved += result.planningRemoved || 0;
  });
  return { rows: updated, closedLogs, transferred, planningRemoved, updated: updated.length };
}

module.exports = {
  liveTaskForStatus,
  runningByOrder,
  annotateOrders,
  closeOpenLogsForOrder,
  transferOpenLogsForOrder,
  correctOrderShop,
  correctManyOrders,
  SHOP_STATUSES
};
