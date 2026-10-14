"use strict";

/**
 * Keep production-schedule letters in sync with Planning + live shop clocks.
 * Never overwrites or clears LD / LC / LD* / LC* / PD / C (Courier).
 * Never auto-plans PD or C. Production chain paints as CS → T → P → W → G → PQC → PC → U? → A → FQC.
 * Ready for Delivery / Out for Delivery / At couriers / Delivered: no auto production letters.
 * Only paints stations still remaining for the order's shop status (so PQC never appears after A).
 */

const sched = require("./office-schedule");
const { isProductionFinished, remainingPlanForStatus } = require("./shop-status");

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

/** Processes still allowed to paint for this status (letter → true). */
function allowedLettersForStatus(status) {
  const remaining = remainingPlanForStatus(status);
  const bag = {};
  (remaining.processes || []).forEach((p) => {
    const letter = letterForProcess(p);
    if (letter) bag[letter] = true;
  });
  // Paint-shop Monday block maps to PC while paintWait is active.
  if (remaining.paintWait) bag.PC = true;
  return bag;
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
  const statuses = orderStatusById(db);
  const done = finishedOrderIds(db);
  const allowedByOrder = new Map();

  function allowed(orderId) {
    if (allowedByOrder.has(orderId)) return allowedByOrder.get(orderId);
    const bag = allowedLettersForStatus(statuses.get(orderId) || "");
    allowedByOrder.set(orderId, bag);
    return bag;
  }

  function put(orderNumber, day, code) {
    const orderId = db.formatOrderId(orderNumber);
    const dayIso = String(day || "").slice(0, 10);
    const letter = sched.normalizeScheduleToken(code);
    if (!orderId || !dayIso || !letter || !/^\d{4}-\d{2}-\d{2}$/.test(dayIso)) return;
    if (done.has(orderId)) return;
    const upper = letter.toUpperCase();
    if (upper === "LD" || upper === "LC" || upper === "LD*" || upper === "LC*") return;
    if (sched.isNeverAutoScheduleCode(letter)) return;
    // Only paint stations still remaining for this shop status.
    if (!allowed(orderId)[letter]) return;
    const key = orderId + "|" + dayIso;
    marks.set(key, sched.mergeScheduleCodes(marks.get(key) || "", letter));
  }

  const store = plan.load();
  const journey = plan.attachActuals(plan.buildJourney(store.blocks || []));
  ((journey && journey.orders) || []).forEach((order) => {
    const orderId = db.formatOrderId(order.orderId);
    if (done.has(orderId)) return;
    (order.rows || []).forEach((row) => {
      const letter = letterForProcess(row.process);
      if (!letter) return;
      (row.days || []).forEach((iso) => put(order.orderId, iso, letter));
      ((row.actual && row.actual.days) || []).forEach((iso) => put(order.orderId, iso, letter));
    });
  });

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
      // Never clear LD / LC / moved / Courier / PD cells.
      if (sched.isDeliveryProtectedToken(value)) return;
      if (sources[day] !== "auto") return;
      // Drop auto production letters once production is finished, or when the mark is gone.
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
