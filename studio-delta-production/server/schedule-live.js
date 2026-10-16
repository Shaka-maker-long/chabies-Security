"use strict";

/**
 * Keep production-schedule letters in sync with the Planning page plan.
 * Paints the same process days as Planning Journey Plan rows (buildJourney).
 * Never overwrites or clears LD / LC / LD* / LC* / PD / C (Courier).
 * Never auto-plans PD or C.
 * In Stock / Ready for Delivery / Out for Delivery / At couriers / Delivered: no auto production letters.
 */

const sched = require("./office-schedule");
const { isProductionFinished } = require("./shop-status");

function todayIso() {
  try {
    return require("./floor-planning").isoDateFromMs(Date.now());
  } catch (e) {
    return new Date().toISOString().slice(0, 10);
  }
}

function letterForProcess(process) {
  return sched.shopProcessToScheduleCode(process) || "";
}

function orderStatusById(db) {
  const out = new Map();
  try {
    db.listOrders().forEach((o) => {
      const id = db.formatOrderId(o.order_number);
      if (id) out.set(id, o.status || "");
    });
  } catch (e) {}
  return out;
}

function finishedOrderIds(db) {
  const out = new Set();
  orderStatusById(db).forEach((status, id) => {
    if (isProductionFinished(status)) out.add(id);
  });
  return out;
}

function readOpenShopLetters(db, plan) {
  const out = [];
  try {
    const { getBook } = require("./workbook-store");
    const sheet = getBook().getSheetByName("Production_Log");
    if (!sheet || sheet.getLastRow() < 2) return out;
    const lastCol = Math.max(sheet.getLastColumn(), 7);
    const grid = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
    const today = todayIso();
    for (let i = 1; i < grid.length; i++) {
      const row = grid[i] || [];
      if (row[6]) continue;
      const orderId = db.formatOrderId(row[1]);
      if (!orderId) continue;
      const role = String(row[3] || "").trim();
      const process = String(row[4] || "").trim() || role;
      let letter = letterForProcess(process) || letterForProcess(role);
      if (!letter) {
        const matched = plan.matchJourneyProcess(role) || plan.matchJourneyProcess(process);
        letter = letterForProcess(matched);
      }
      if (!letter) continue;
      out.push({ orderId, day: today, code: letter });
    }
  } catch (e) {}
  return out;
}

/**
 * Drop letters that would break CS→T→P→W→G→PQC→PC→U→A→FQC day order
 * for a single order (later station must not sit on an earlier day than an earlier station).
 */
function enforceChainDayOrder(marks) {
  const byOrder = new Map();
  marks.forEach((code, key) => {
    const split = key.indexOf("|");
    const orderId = key.slice(0, split);
    const day = key.slice(split + 1);
    if (!byOrder.has(orderId)) byOrder.set(orderId, []);
    byOrder.get(orderId).push({ day, code, key });
  });
  const order = sched.AUTO_SCHEDULE_ORDER;
  byOrder.forEach((rows) => {
    const earliest = {};
    rows.forEach((row) => {
      sched.parseScheduleCodes(row.code).forEach((letter) => {
        const ix = order.indexOf(letter);
        if (ix === -1) return;
        if (!earliest[letter] || row.day < earliest[letter]) earliest[letter] = row.day;
      });
    });
    rows.forEach((row) => {
      const kept = sched.parseScheduleCodes(row.code).filter((letter) => {
        const ix = order.indexOf(letter);
        if (ix === -1) return true;
        for (let i = 0; i < ix; i++) {
          const earlier = order[i];
          if (earliest[earlier] && row.day < earliest[earlier]) return false;
        }
        return true;
      });
      if (!kept.length) {
        marks.delete(row.key);
        return;
      }
      const next = sched.formatScheduleCodes(kept);
      if (next !== row.code) marks.set(row.key, next);
    });
  });
}

function collectAutoMarks(db, plan) {
  const marks = new Map();
  const done = finishedOrderIds(db);

  function put(orderNumber, day, code) {
    const orderId = db.formatOrderId(orderNumber);
    const dayIso = String(day || "").slice(0, 10);
    const letter = sched.normalizeScheduleToken(code);
    if (!orderId || !dayIso || !letter || !/^\d{4}-\d{2}-\d{2}$/.test(dayIso)) return;
    if (done.has(orderId)) return;
    const upper = letter.toUpperCase();
    if (upper === "LD" || upper === "LC" || upper === "LD*" || upper === "LC*") return;
    if (sched.isNeverAutoScheduleCode(letter)) return;
    const key = orderId + "|" + dayIso;
    marks.set(key, sched.mergeScheduleCodes(marks.get(key) || "", letter));
  }

  // Mirror Planning Journey Plan days exactly (not Actual / remaining-only).
  const store = plan.load();
  const journey = plan.buildJourney(store.blocks || []);
  ((journey && journey.orders) || []).forEach((order) => {
    const orderId = db.formatOrderId(order.orderId);
    if (done.has(orderId)) return;
    (order.rows || []).forEach((row) => {
      const letter = letterForProcess(row.process) || sched.normalizeScheduleToken(row.code);
      if (!letter) return;
      (row.days || []).forEach((iso) => put(order.orderId, iso, letter));
    });
  });

  // Live open clocks still mark today so the grid stays current mid-shift.
  readOpenShopLetters(db, plan).forEach((hit) => put(hit.orderId, hit.day, hit.code));

  enforceChainDayOrder(marks);
  return marks;
}

function dropFinishedPlanBlocks(db, plan) {
  const done = finishedOrderIds(db);
  if (!done.size) return 0;
  const store = plan.load();
  const before = (store.blocks || []).length;
  store.blocks = (store.blocks || []).filter((b) => {
    if (!b || b.kind === "other") return true;
    return !done.has(db.formatOrderId(b.orderId));
  });
  const removed = before - store.blocks.length;
  if (removed) plan.save(store);
  return removed;
}

function syncLiveScheduleCodes() {
  const db = require("./db");
  const plan = require("./floor-planning");
  db.syncScheduleFromOrders();
  dropFinishedPlanBlocks(db, plan);
  const rows = db.listSchedule("2000-01-01", "2099-12-31");
  const byOrder = new Map(rows.map((r) => [db.formatOrderId(r.order_number), r]));
  const done = finishedOrderIds(db);
  const marks = collectAutoMarks(db, plan);

  let cleared = 0;
  let written = 0;

  rows.forEach((row) => {
    const sources = row.cell_sources || {};
    const orderId = db.formatOrderId(row.order_number);
    Object.keys(row.cells || {}).forEach((day) => {
      const value = row.cells[day];
      if (sched.isDeliveryProtectedToken(value)) return;
      if (sources[day] !== "auto") return;
      if (!done.has(orderId) && marks.get(orderId + "|" + day)) return;
      db.setScheduleCell(row.id, day, "", false, { skipPlan: true, source: "auto" });
      cleared += 1;
      delete row.cells[day];
    });
  });

  marks.forEach((code, key) => {
    const split = key.indexOf("|");
    const orderId = key.slice(0, split);
    const day = key.slice(split + 1);
    if (done.has(orderId)) return;
    const row = byOrder.get(orderId);
    if (!row) return;
    const current = (row.cells && row.cells[day]) || "";
    if (sched.isProtectedScheduleCode(current)) return;
    if (sched.isDeliveryProtectedToken(current)) return;
    const source = (row.cell_sources && row.cell_sources[day]) || "";
    if (current && source !== "auto") return;
    if (String(current) === String(code) && source === "auto") return;
    db.setScheduleCell(row.id, day, code, false, { skipPlan: true, source: "auto" });
    if (!row.cells) row.cells = {};
    row.cells[day] = code;
    if (!row.cell_sources) row.cell_sources = {};
    row.cell_sources[day] = "auto";
    written += 1;
  });

  try { db.persistOffice(); } catch (e) {}
  return { cleared, written, marks: marks.size };
}

module.exports = {
  syncLiveScheduleCodes,
  collectAutoMarks,
  todayIso,
  readOpenShopLetters,
  enforceChainDayOrder
};
