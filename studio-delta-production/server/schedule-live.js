"use strict";

/**
 * Keep production-schedule letters in sync with Planning + live shop clocks.
 * Never overwrites or clears LD / LC / LD* / LC* / PD / C (Courier).
 * Never auto-plans PD or C. Production chain paints as CS → T → P → W → G → PQC → PC → U? → A → FQC.
 * Delivered / Out for Delivery / At couriers orders get no auto production letters.
 */

const sched = require("./office-schedule");
const { isHandedOff } = require("./shop-status");

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

function handedOffOrderIds(db) {
  const out = new Set();
  try {
    db.listOrders().forEach((o) => {
      if (!isHandedOff(o.status)) return;
      const id = db.formatOrderId(o.order_number);
      if (id) out.add(id);
    });
  } catch (e) {}
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

function collectAutoMarks(db, plan) {
  const marks = new Map();
  const done = handedOffOrderIds(db);
  function put(orderNumber, day, code) {
    const orderId = db.formatOrderId(orderNumber);
    const dayIso = String(day || "").slice(0, 10);
    const letter = sched.normalizeScheduleToken(code);
    if (!orderId || !dayIso || !letter || !/^\d{4}-\d{2}-\d{2}$/.test(dayIso)) return;
    if (done.has(orderId)) return;
    // Never emit LD / LC / moved / PD / Courier as auto marks.
    const upper = letter.toUpperCase();
    if (upper === "LD" || upper === "LC" || upper === "LD*" || upper === "LC*") return;
    if (sched.isNeverAutoScheduleCode(letter)) return;
    const key = orderId + "|" + dayIso;
    marks.set(key, sched.mergeScheduleCodes(marks.get(key) || "", letter));
  }

  const store = plan.load();
  const journey = plan.attachActuals(plan.buildJourney(store.blocks || []));
  ((journey && journey.orders) || []).forEach((order) => {
    if (done.has(db.formatOrderId(order.orderId))) return;
    (order.rows || []).forEach((row) => {
      const letter = letterForProcess(row.process);
      if (!letter) return;
      (row.days || []).forEach((iso) => put(order.orderId, iso, letter));
      ((row.actual && row.actual.days) || []).forEach((iso) => put(order.orderId, iso, letter));
    });
  });

  readOpenShopLetters(db, plan).forEach((hit) => put(hit.orderId, hit.day, hit.code));

  return marks;
}

function syncLiveScheduleCodes() {
  const db = require("./db");
  const plan = require("./floor-planning");
  db.syncScheduleFromOrders();
  const rows = db.listSchedule("2000-01-01", "2099-12-31");
  const byOrder = new Map(rows.map((r) => [db.formatOrderId(r.order_number), r]));
  const done = handedOffOrderIds(db);
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
      // Drop auto production letters once the job is handed off / delivered.
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
  readOpenShopLetters
};
