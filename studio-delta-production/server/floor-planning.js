"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir, getBook } = require("./workbook-store");
const { listOrders, formatOrderId } = require("./db");
const staff = require("./staff");

const ZONE = "+02:00";
const PAINT_WORKER_ID = "__paint_shop__";
const PAINT_WORKER_NAME = "Paint shop";
const PAINT_WAIT_DAYS = 5;
const QUEUE_STATUSES = ["Not Yet Started", "Ready for Steelwork"];

const PLANNED_PROCESSES = [
  "Profile Cutting",
  "Tagging",
  "Plate Cutting",
  "Welding",
  "Grinding",
  "Assembly"
];

const PROCESS_COLORS = {
  "Profile Cutting": { bg: "#fef3c7", fg: "#92400e", border: "#d97706" },
  "Tagging": { bg: "#e0f2fe", fg: "#075985", border: "#0284c7" },
  "Plate Cutting": { bg: "#ede9fe", fg: "#5b21b6", border: "#7c3aed" },
  "Welding": { bg: "#fee2e2", fg: "#991b1b", border: "#dc2626" },
  "Grinding": { bg: "#dcfce7", fg: "#166534", border: "#16a34a" },
  "Powder coating": { bg: "#e5e7eb", fg: "#374151", border: "#6b7280" },
  "Assembly": { bg: "#fae8ff", fg: "#86198f", border: "#c026d3" },
  Other: { bg: "#ffedd5", fg: "#9a3412", border: "#f97316" }
};

const OTHER_TASKS = ["Cleaning", "Production Meeting", "Maintenance", "Material", "Setup"];

const DAY_BANDS = [
  { id: "meeting", label: "Production Meeting", startMin: 7 * 60 + 45, endMin: 8 * 60, bg: "#fde68a", fg: "#92400e" },
  { id: "lunch", label: "Lunch", startMin: 12 * 60, endMin: 12 * 60 + 30, bg: "#e5e7eb", fg: "#374151" },
  { id: "cleaning", label: "Cleaning", startMin: 15 * 60 + 15, endMin: 15 * 60 + 45, bg: "#a7f3d0", fg: "#065f46" }
];

const WINDOWS = [
  { startMin: 7 * 60 + 45, endMin: 12 * 60 },
  { startMin: 12 * 60 + 30, endMin: 15 * 60 + 45 }
];

const WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
const WEEKDAY_LONG = {
  Mon: "Monday",
  Tue: "Tuesday",
  Wed: "Wednesday",
  Thu: "Thursday",
  Fri: "Friday",
  Sat: "Saturday",
  Sun: "Sunday"
};

function pad(n) {
  return String(n).padStart(2, "0");
}

function planningPath() {
  return path.join(dataDir(), "floor-planning.json");
}

function emptyStore() {
  return { blocks: [] };
}

function load() {
  try {
    const parsed = JSON.parse(fs.readFileSync(planningPath(), "utf8"));
    const blocks = Array.isArray(parsed.blocks) ? parsed.blocks.filter(Boolean) : [];
    return { blocks };
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[planning] could not read", planningPath(), e.message || e);
    }
    return emptyStore();
  }
}

function save(store) {
  const file = planningPath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify({ blocks: (store && store.blocks) || [] }, null, 2));
  fs.renameSync(tmp, file);
  return store;
}

function newId() {
  return crypto.randomBytes(8).toString("hex");
}

function sastMs(year, month, day, hour, minute) {
  return new Date(
    year + "-" + pad(month) + "-" + pad(day) + "T" + pad(hour) + ":" + pad(minute) + ":00" + ZONE
  ).getTime();
}

function partsFromMs(ms) {
  const dtf = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Africa/Johannesburg",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    weekday: "short",
    hourCycle: "h23"
  });
  const map = {};
  for (const part of dtf.formatToParts(new Date(ms))) {
    if (part.type !== "literal") map[part.type] = part.value;
  }
  const weekday = String(map.weekday || "").slice(0, 3);
  return {
    year: Number(map.year),
    month: Number(map.month),
    day: Number(map.day),
    hour: Number(map.hour),
    minute: Number(map.minute),
    weekday,
    weekdayLong: WEEKDAY_LONG[weekday] || weekday
  };
}

function isoFromMs(ms) {
  const p = partsFromMs(ms);
  return p.year + "-" + pad(p.month) + "-" + pad(p.day) + "T" + pad(p.hour) + ":" + pad(p.minute) + ":00" + ZONE;
}

function isoDateFromMs(ms) {
  const p = partsFromMs(ms);
  return p.year + "-" + pad(p.month) + "-" + pad(p.day);
}

function toMs(value) {
  if (value instanceof Date) return value.getTime();
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const s = String(value || "").trim();
  if (!s) return NaN;
  const n = Date.parse(s);
  return Number.isFinite(n) ? n : NaN;
}

function addCalendarDays(year, month, day, n) {
  const ms = sastMs(year, month, day, 12, 0) + n * 86400000;
  const p = partsFromMs(ms);
  return { year: p.year, month: p.month, day: p.day };
}

function isWeekend(parts) {
  return parts.weekday === "Sat" || parts.weekday === "Sun";
}

function workWindowsOnDay(year, month, day) {
  const noon = partsFromMs(sastMs(year, month, day, 12, 0));
  if (isWeekend(noon)) return [];
  return WINDOWS.map((w) => ({
    start: sastMs(year, month, day, Math.floor(w.startMin / 60), w.startMin % 60),
    end: sastMs(year, month, day, Math.floor(w.endMin / 60), w.endMin % 60)
  }));
}

function currentOrNextWindow(ms) {
  const p = partsFromMs(ms);
  for (let i = 0; i < 21; i++) {
    const day = addCalendarDays(p.year, p.month, p.day, i);
    const wins = workWindowsOnDay(day.year, day.month, day.day);
    for (let w = 0; w < wins.length; w++) {
      if (wins[w].end > ms) return wins[w];
    }
  }
  return null;
}

function nextWorkInstant(ms) {
  const t = toMs(ms);
  const win = currentOrNextWindow(Number.isFinite(t) ? t : Date.now());
  if (!win) return Number.isFinite(t) ? t : Date.now();
  return Math.max(Number.isFinite(t) ? t : win.start, win.start);
}

function ceilToMinute(ms) {
  const p = partsFromMs(ms);
  const exact = sastMs(p.year, p.month, p.day, p.hour, p.minute);
  if (exact < ms) return nextWorkInstant(exact + 60000);
  return ms;
}

function minutesBetween(startMs, endMs) {
  return Math.max(0, Math.round((endMs - startMs) / 60000));
}

function normalizeBusy(busy) {
  return (busy || [])
    .map((b) => ({ start: toMs(b.start), end: toMs(b.end) }))
    .filter((b) => Number.isFinite(b.start) && Number.isFinite(b.end) && b.end > b.start)
    .sort((a, b) => a.start - b.start || a.end - b.end);
}

function placeTask(fromMs, durationMinutes, busy) {
  const minutes = Math.max(0, Math.round(Number(durationMinutes) || 0));
  if (!(minutes > 0)) return { segments: [], start: null, end: null };
  const busyList = normalizeBusy(busy);
  let cursor = ceilToMinute(nextWorkInstant(fromMs));
  let left = minutes;
  const segments = [];
  let guard = 0;
  while (left > 0 && guard++ < 20000) {
    const win = currentOrNextWindow(cursor);
    if (!win) break;
    if (cursor < win.start) cursor = win.start;
    if (cursor >= win.end) {
      cursor = nextWorkInstant(win.end);
      continue;
    }
    let availEnd = win.end;
    let skipTo = null;
    for (let i = 0; i < busyList.length; i++) {
      const b = busyList[i];
      if (b.end <= cursor) continue;
      if (b.start >= win.end) break;
      if (b.start <= cursor && b.end > cursor) {
        skipTo = b.end;
        break;
      }
      if (b.start > cursor && b.start < availEnd) {
        availEnd = b.start;
        break;
      }
    }
    if (skipTo != null) {
      cursor = ceilToMinute(nextWorkInstant(skipTo));
      continue;
    }
    const availMin = minutesBetween(cursor, availEnd);
    if (availMin <= 0) {
      cursor = nextWorkInstant(availEnd);
      continue;
    }
    const take = Math.min(left, availMin);
    const end = cursor + take * 60000;
    segments.push({ start: isoFromMs(cursor), end: isoFromMs(end) });
    left -= take;
    cursor = end;
  }
  return {
    segments,
    start: segments.length ? segments[0].start : null,
    end: segments.length ? segments[segments.length - 1].end : null
  };
}

function addWorkMinutes(fromMs, minutes) {
  const placed = placeTask(fromMs, minutes, []);
  return placed.end ? toMs(placed.end) : null;
}

function earliestPaintMonday(grindEndMs) {
  const p = partsFromMs(toMs(grindEndMs));
  let y = p.year;
  let m = p.month;
  let d = p.day;
  for (let i = 0; i < 8; i++) {
    const morning = sastMs(y, m, d, 7, 45);
    const q = partsFromMs(morning);
    if (q.weekday === "Mon") return morning;
    const next = addCalendarDays(y, m, d, 1);
    y = next.year;
    m = next.month;
    d = next.day;
  }
  return sastMs(p.year, p.month, p.day, 7, 45);
}

function weekMondayIso(week) {
  const raw = String(week || "").trim();
  let iso = /^\d{4}-\d{2}-\d{2}$/.test(raw) ? raw : isoDateFromMs(Date.now());
  const parts = iso.split("-").map(Number);
  let ms = sastMs(parts[0], parts[1], parts[2], 12, 0);
  for (let i = 0; i < 7; i++) {
    const p = partsFromMs(ms);
    if (p.weekday === "Mon") return isoDateFromMs(ms);
    ms -= 86400000;
  }
  return iso;
}

function shiftWeek(mondayIso, weeks) {
  const iso = weekMondayIso(mondayIso);
  const parts = iso.split("-").map(Number);
  const next = addCalendarDays(parts[0], parts[1], parts[2], weeks * 7);
  return next.year + "-" + pad(next.month) + "-" + pad(next.day);
}

function weekDays(mondayIso) {
  const iso = weekMondayIso(mondayIso);
  const parts = iso.split("-").map(Number);
  const out = [];
  for (let i = 0; i < 5; i++) {
    const day = addCalendarDays(parts[0], parts[1], parts[2], i);
    const ms = sastMs(day.year, day.month, day.day, 12, 0);
    const p = partsFromMs(ms);
    const dateIso = day.year + "-" + pad(day.month) + "-" + pad(day.day);
    out.push({
      iso: dateIso,
      weekday: p.weekdayLong,
      weekdayShort: p.weekday,
      label: p.weekdayLong.slice(0, 3) + " " + p.day + " " + monthShort(p.month)
    });
  }
  return out;
}

function monthShort(month) {
  return ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"][month - 1] || "";
}

const JOURNEY_PROCESS_ORDER = {
  "Profile Cutting": 1,
  "Tagging": 2,
  "Plate Cutting": 3,
  "Welding": 4,
  "Grinding": 5,
  "Powder coating": 6,
  "Assembly": 7
};

const JOURNEY_WEEK_COUNT = 20;

const PROCESS_CODES = {
  "Profile Cutting": "C",
  "Tagging": "T",
  "Plate Cutting": "P",
  "Welding": "W",
  "Grinding": "G",
  "Powder coating": "PC",
  "Assembly": "A"
};

function processCode(process) {
  return PROCESS_CODES[process] || String(process || "").slice(0, 2).toUpperCase();
}

function formatWeekRange(mondayIso) {
  const days = weekDays(mondayIso);
  if (!days.length) return "";
  const a = days[0].iso.split("-").map(Number);
  const b = days[4].iso.split("-").map(Number);
  const am = monthShort(a[1]);
  const bm = monthShort(b[1]);
  if (a[1] === b[1]) return a[2] + "–" + b[2] + " " + am;
  return a[2] + " " + am + " – " + b[2] + " " + bm;
}

function journeyWeeks(mondayIso, count) {
  const n = count > 0 ? count : 5;
  const out = [];
  let iso = weekMondayIso(mondayIso);
  const todayMonday = weekMondayIso();
  for (let i = 0; i < n; i++) {
    out.push({
      index: i + 1,
      label: formatWeekRange(iso),
      start: iso,
      days: weekDays(iso),
      current: iso === todayMonday
    });
    iso = shiftWeek(iso, 1);
  }
  return out;
}

function addJourneyIso(iso, bounds) {
  if (!iso) return;
  if (!bounds.min || iso < bounds.min) bounds.min = iso;
  if (!bounds.max || iso > bounds.max) bounds.max = iso;
}

function journeyDayBounds(journey) {
  const bounds = { min: "", max: "" };
  ((journey && journey.orders) || []).forEach((o) => {
    (o.rows || []).forEach((row) => {
      (row.days || []).forEach((iso) => addJourneyIso(iso, bounds));
      ((row.actual && row.actual.days) || []).forEach((iso) => addJourneyIso(iso, bounds));
    });
  });
  return bounds;
}

function journeyWeeksFromPlan(journey, minCount) {
  const bounds = journeyDayBounds(journey);
  const min = bounds.min;
  const max = bounds.max;
  if (!min) return journeyWeeks(undefined, minCount || 5);
  const start = weekMondayIso(min);
  const last = weekMondayIso(max);
  let n = 1;
  let iso = start;
  while (iso < last && n < 20) {
    iso = shiftWeek(iso, 1);
    n += 1;
  }
  return journeyWeeks(start, Math.max(minCount || 5, n));
}

function isWorkIso(iso) {
  const parts = String(iso || "").split("-").map(Number);
  if (parts.length < 3 || !parts[0]) return false;
  const p = partsFromMs(sastMs(parts[0], parts[1], parts[2], 12, 0));
  return !isWeekend(p);
}

function formatDayHeader(iso) {
  const parts = String(iso || "").split("-").map(Number);
  if (parts.length < 3 || !parts[0]) return "";
  const p = partsFromMs(sastMs(parts[0], parts[1], parts[2], 12, 0));
  return pad(p.day) + "-" + monthShort(p.month);
}

function workdaysFromTo(fromIso, toIso) {
  const out = [];
  let iso = String(fromIso || "").slice(0, 10);
  const last = String(toIso || "").slice(0, 10);
  if (!iso || !last || iso > last) return out;
  let guard = 0;
  while (iso <= last && guard++ < 400) {
    const parts = iso.split("-").map(Number);
    const p = partsFromMs(sastMs(parts[0], parts[1], parts[2], 12, 0));
    if (!isWeekend(p)) {
      out.push({
        iso,
        weekday: p.weekdayLong,
        weekdayShort: p.weekday,
        label: pad(p.day) + "-" + monthShort(p.month)
      });
    }
    const next = addCalendarDays(parts[0], parts[1], parts[2], 1);
    iso = next.year + "-" + pad(next.month) + "-" + pad(next.day);
  }
  return out;
}

function occupiedWorkdays(block) {
  const startMs = toMs(block && block.start);
  const endMs = toMs(block && block.end);
  if (!Number.isFinite(startMs) || !Number.isFinite(endMs) || endMs <= startMs) return [];
  const startIso = isoDateFromMs(startMs);
  const endParts = partsFromMs(endMs);
  let endIso = isoDateFromMs(endMs);
  if (endParts.hour === 7 && endParts.minute === 45) {
    const prev = addCalendarDays(endParts.year, endParts.month, endParts.day, -1);
    endIso = prev.year + "-" + pad(prev.month) + "-" + pad(prev.day);
  }
  if (endIso < startIso) endIso = startIso;
  return workdaysFromTo(startIso, endIso).map((d) => d.iso);
}

function emptyActual() {
  return { start: "", end: "", days: [], workerId: "", workerName: "", bouts: [] };
}

function matchJourneyProcess(task) {
  const s = String(task || "").toLowerCase();
  if (!s) return "";
  if (/pre-powder|final qc|quality control|paint prep|painting/.test(s) && !/powder coating/.test(s)) return "";
  if (/powder coating|paint shop/.test(s)) return "Powder coating";
  const names = PLANNED_PROCESSES.concat(["Powder coating"]);
  return names.find((name) => s.indexOf(name.toLowerCase().split(" ")[0]) !== -1) || "";
}

function readProductionActuals() {
  try {
    const sheet = getBook().getSheetByName("Production_Log");
    if (!sheet || sheet.getLastRow() < 2) return [];
    const lastCol = Math.max(sheet.getLastColumn(), 13);
    const grid = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
    const nowMs = Date.now();
    const out = [];
    for (let i = 1; i < grid.length; i++) {
      const row = grid[i] || [];
      const orderId = formatOrderId(row[1]);
      const process = matchJourneyProcess(row[3]);
      if (!orderId || !process) continue;
      const startMs = toMs(row[5]);
      if (!Number.isFinite(startMs)) continue;
      let endMs = toMs(row[6]);
      if (!Number.isFinite(endMs) || endMs <= startMs) endMs = nowMs;
      const bout = { start: isoFromMs(startMs), end: isoFromMs(endMs) };
      out.push({
        orderId,
        process,
        workerId: String(row[2] || "").trim(),
        workerName: String(row[2] || "").trim(),
        start: bout.start,
        end: bout.end,
        bouts: [bout]
      });
    }
    return out;
  } catch (e) {
    return [];
  }
}

function groupProductionActuals(logs) {
  const groups = {};
  (logs || []).forEach((log) => {
    const key = log.orderId + "||" + log.process;
    if (!groups[key]) {
      groups[key] = {
        orderId: log.orderId,
        process: log.process,
        workerId: log.workerId,
        workerName: log.workerName,
        start: log.start,
        end: log.end,
        bouts: [],
        daySet: {}
      };
    }
    const g = groups[key];
    if (String(log.start || "") < String(g.start || "")) g.start = log.start;
    if (String(log.end || "") > String(g.end || "")) g.end = log.end;
    if (log.workerName) {
      g.workerId = log.workerId;
      g.workerName = log.workerName;
    }
    (log.bouts || [{ start: log.start, end: log.end }]).forEach((bout) => {
      g.bouts.push({ start: bout.start, end: bout.end });
      occupiedWorkdays(bout).forEach((iso) => { g.daySet[iso] = true; });
    });
  });
  Object.keys(groups).forEach((k) => {
    const g = groups[k];
    g.days = Object.keys(g.daySet).sort();
    g.bouts.sort((a, b) => String(a.start).localeCompare(String(b.start)));
    delete g.daySet;
  });
  return groups;
}

function attachActuals(journey) {
  const groups = groupProductionActuals(readProductionActuals());
  ((journey && journey.orders) || []).forEach((o) => {
    (o.rows || []).forEach((row) => {
      const key = formatOrderId(row.orderId || o.orderId) + "||" + String(row.process || "");
      const act = groups[key];
      row.actual = act
        ? {
          start: act.start,
          end: act.end,
          days: act.days.slice(),
          workerId: act.workerId,
          workerName: act.workerName,
          bouts: act.bouts.map((b) => ({ start: b.start, end: b.end }))
        }
        : emptyActual();
    });
  });
  const bounds = journeyDayBounds(journey);
  if (journey) {
    journey.days = bounds.min && bounds.max ? workdaysFromTo(bounds.min, bounds.max) : [];
  }
  return journey;
}

function buildJourney(blocks) {
  const groups = {};
  (blocks || []).forEach((b) => {
    if (!b || !b.orderId || b.kind === "other") return;
    const orderId = formatOrderId(b.orderId);
    const key = orderId + "||" + String(b.process || "");
    if (!groups[key]) {
      groups[key] = {
        orderId,
        product: String(b.product || ""),
        process: String(b.process || ""),
        code: processCode(b.process),
        workerId: b.workerId,
        workerName: b.workerName || b.workerId,
        start: b.start,
        end: b.end,
        blockId: b.id,
        daySet: {}
      };
    }
    const g = groups[key];
    if (String(b.start || "") < String(g.start || "")) g.start = b.start;
    if (String(b.end || "") > String(g.end || "")) g.end = b.end;
    occupiedWorkdays(b).forEach((iso) => { g.daySet[iso] = true; });
  });
  const byOrder = {};
  Object.keys(groups).forEach((k) => {
    const g = groups[k];
    if (!byOrder[g.orderId]) byOrder[g.orderId] = [];
      byOrder[g.orderId].push({
        orderId: g.orderId,
        product: g.product,
        process: g.process,
        code: g.code || processCode(g.process),
        workerId: g.workerId,
        workerName: g.workerName,
        start: g.start,
        end: g.end,
        blockId: g.blockId,
        days: Object.keys(g.daySet).sort()
      });
  });
  const orderIds = Object.keys(byOrder).sort();
  let minIso = "";
  let maxIso = "";
  orderIds.forEach((id) => {
    byOrder[id].sort((a, b) => (JOURNEY_PROCESS_ORDER[a.process] || 99) - (JOURNEY_PROCESS_ORDER[b.process] || 99));
    byOrder[id].forEach((row) => {
      row.days.forEach((iso) => {
        if (!minIso || iso < minIso) minIso = iso;
        if (!maxIso || iso > maxIso) maxIso = iso;
      });
    });
  });
  const built = {
    days: minIso && maxIso ? workdaysFromTo(minIso, maxIso) : [],
    orders: orderIds.map((id) => {
      const rows = byOrder[id];
      let start = "";
      let end = "";
      (rows || []).forEach((row) => {
        if (row.start && (!start || row.start < start)) start = row.start;
        if (row.end && (!end || row.end > end)) end = row.end;
      });
      return {
        orderId: id,
        product: rows[0] ? rows[0].product : "",
        start,
        end,
        rows
      };
    }).sort((a, b) => String(a.start || "").localeCompare(String(b.start || "")) || String(a.orderId).localeCompare(String(b.orderId)))
  };
  return attachActuals(built);
}

const USER_ASSIGNED_PROCESSES = [
  "Profile Cutting",
  "Tagging",
  "Plate Cutting",
  "Welding",
  "Assembly"
];
const GRIND_POOL_TASKS = ["Profile Cutting", "Tagging", "Welding"];

function namesEqual(a, b) {
  return String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase();
}

function busyForWorker(blocks, workerId) {
  return (blocks || [])
    .filter((b) => b && namesEqual(b.workerId, workerId) && b.kind !== "paint")
    .map((b) => ({ start: b.start, end: b.end }));
}

function minutesByProcessFor(product) {
  const out = {};
  PLANNED_PROCESSES.forEach((process) => {
    out[process] = staff.durationMinutes(product, process) || 0;
  });
  return out;
}

function workersForProcess(process, users) {
  return (users || staff.listUsers())
    .filter((u) => (u.tasks || []).indexOf(process) !== -1)
    .map((u) => ({ id: u.name, name: u.name }));
}

function findUser(name, users) {
  const list = users || staff.listUsers();
  return list.find((u) => namesEqual(u.name, name)) || null;
}

function placeProcess(order, process, fromMs, workerName, busyBlocks) {
  const minutes = staff.durationMinutes(order.product, process);
  if (!(minutes > 0)) return { endMs: fromMs, blocks: [] };
  const who = String(workerName || "").trim();
  if (!who) throw new Error("Assign someone for " + process + " on " + order.order_number + ".");
  const user = findUser(who);
  if (!user) throw new Error(who + " is not on Users.");
  if ((user.tasks || []).indexOf(process) === -1) {
    throw new Error(who + " is not ticked for " + process + " on Users.");
  }
  const placed = placeTask(fromMs, minutes, busyForWorker(busyBlocks, who));
  if (!placed.segments.length) {
    throw new Error("Could not place " + process + " for " + order.order_number + ".");
  }
  const blocks = placed.segments.map((seg) => ({
    id: newId(),
    orderId: formatOrderId(order.order_number),
    studioNo: formatOrderId(order.order_number),
    product: String(order.product || ""),
    process,
    workerId: user.name,
    workerName: user.name,
    start: seg.start,
    end: seg.end,
    kind: "work"
  }));
  return { endMs: toMs(placed.end), blocks };
}

function workerBusyMinutes(blocks, workerId) {
  return busyForWorker(blocks, workerId).reduce((n, b) => {
    return n + minutesBetween(toMs(b.start), toMs(b.end));
  }, 0);
}

function fillContiguous(startMs, minutes, busyList) {
  let cursor = startMs;
  let left = minutes;
  const segments = [];
  let guard = 0;
  while (left > 0 && guard++ < 5000) {
    const win = currentOrNextWindow(cursor);
    if (!win) return null;
    if (cursor < win.start) cursor = win.start;
    if (cursor >= win.end) {
      cursor = nextWorkInstant(win.end);
      continue;
    }
    let availEnd = win.end;
    for (let i = 0; i < busyList.length; i++) {
      const b = busyList[i];
      if (b.end <= cursor) continue;
      if (b.start >= win.end) break;
      if (b.start <= cursor && b.end > cursor) return null;
      if (b.start > cursor && b.start < availEnd) {
        availEnd = b.start;
        break;
      }
    }
    const availMin = minutesBetween(cursor, availEnd);
    if (availMin <= 0) return null;
    if (availEnd < win.end && availMin < left) return null;
    const take = Math.min(left, availMin);
    const end = cursor + take * 60000;
    segments.push({ start: isoFromMs(cursor), end: isoFromMs(end) });
    left -= take;
    cursor = end;
  }
  if (left > 0 || !segments.length) return null;
  return {
    segments,
    start: segments[0].start,
    end: segments[segments.length - 1].end
  };
}

function placeContiguousTask(fromMs, durationMinutes, busy) {
  const minutes = Math.max(0, Math.round(Number(durationMinutes) || 0));
  if (!(minutes > 0)) return { segments: [], start: null, end: null };
  const busyList = normalizeBusy(busy);
  let cursor = ceilToMinute(nextWorkInstant(fromMs));
  let guard = 0;
  while (guard++ < 5000) {
    const win = currentOrNextWindow(cursor);
    if (!win) break;
    if (cursor < win.start) cursor = win.start;
    if (cursor >= win.end) {
      cursor = nextWorkInstant(win.end);
      continue;
    }
    let skipTo = null;
    for (let i = 0; i < busyList.length; i++) {
      const b = busyList[i];
      if (b.end <= cursor) continue;
      if (b.start >= win.end) break;
      if (b.start <= cursor && b.end > cursor) {
        skipTo = b.end;
        break;
      }
    }
    if (skipTo != null) {
      cursor = ceilToMinute(nextWorkInstant(skipTo));
      continue;
    }
    const attempt = fillContiguous(cursor, minutes, busyList);
    if (attempt) return attempt;
    let blocker = null;
    for (let i = 0; i < busyList.length; i++) {
      const b = busyList[i];
      if (b.end <= cursor) continue;
      blocker = b;
      break;
    }
    if (!blocker) break;
    cursor = ceilToMinute(nextWorkInstant(blocker.end));
  }
  return { segments: [], start: null, end: null };
}

function grindingPool(users) {
  return (users || staff.listUsers())
    .filter((u) => {
      const tasks = u.tasks || [];
      if (tasks.indexOf("Grinding") === -1) return false;
      return GRIND_POOL_TASKS.some((t) => tasks.indexOf(t) !== -1);
    })
    .sort((a, b) => String(a.name).localeCompare(String(b.name)));
}

function placeGrindingOnOpenSlot(order, minutes, fromMs, busy, users) {
  const pool = grindingPool(users);
  if (!pool.length) {
    throw new Error(
      "Grinding is auto-assigned. On Users, tick Grinding plus Profile cutting, Tagging, or Welding for at least one person."
    );
  }
  let best = null;
  pool.forEach((user) => {
    const theirs = busyForWorker(busy, user.name);
    const contiguous = placeContiguousTask(fromMs, minutes, theirs);
    const placed = contiguous.segments.length ? contiguous : placeTask(fromMs, minutes, theirs);
    if (!placed.segments.length || !placed.start) return;
    const load = workerBusyMinutes(busy, user.name);
    const split = !contiguous.segments.length;
    const candidate = { user, placed, load, split };
    if (!best) {
      best = candidate;
      return;
    }
    if (candidate.split !== best.split) {
      if (!candidate.split) best = candidate;
      return;
    }
    if (candidate.load !== best.load) {
      if (candidate.load < best.load) best = candidate;
      return;
    }
    if (candidate.placed.start !== best.placed.start) {
      if (candidate.placed.start < best.placed.start) best = candidate;
      return;
    }
    if (String(user.name) < String(best.user.name)) best = candidate;
  });
  if (!best) throw new Error("Could not place Grinding for " + order.order_number + ".");
  const blocks = best.placed.segments.map((seg) => ({
    id: newId(),
    orderId: formatOrderId(order.order_number),
    studioNo: formatOrderId(order.order_number),
    product: String(order.product || ""),
    process: "Grinding",
    workerId: best.user.name,
    workerName: best.user.name,
    start: seg.start,
    end: seg.end,
    kind: "work"
  }));
  return { endMs: toMs(best.placed.end), blocks };
}

function scheduleOrder({ order, assignments, existingBlocks, fromMs }) {
  return scheduleBatch({
    jobs: [{ order, assignments }],
    existingBlocks,
    fromMs
  }).blocks;
}

function scheduleBatch({ jobs, existingBlocks, fromMs }) {
  const users = staff.listUsers();
  let busy = (existingBlocks || []).slice();
  const startMs = nextWorkInstant(fromMs != null ? fromMs : Date.now());
  const drafts = [];

  (jobs || []).forEach((job) => {
    const order = job.order;
    const minutes = minutesByProcessFor(order.product);
    const hasWork = PLANNED_PROCESSES.some((p) => minutes[p] > 0);
    if (!hasWork) {
      throw new Error("No Task times for " + (order.product || order.order_number) + ". Add hours on Task times first.");
    }
    const assign = job.assignments || {};
    const metal = [];
    function run(process, from) {
      if (!(minutes[process] > 0)) return from;
      const placed = placeProcess(order, process, from, assign[process], busy);
      metal.push.apply(metal, placed.blocks);
      busy = busy.concat(placed.blocks);
      return placed.endMs;
    }
    let afterProfile = startMs;
    if (minutes["Profile Cutting"] > 0) afterProfile = run("Profile Cutting", startMs);
    let afterTag = afterProfile;
    if (minutes.Tagging > 0) afterTag = run("Tagging", afterProfile);
    let afterWeld = afterTag;
    if (minutes.Welding > 0) afterWeld = run("Welding", afterTag);
    if (minutes["Plate Cutting"] > 0) run("Plate Cutting", afterTag);
    drafts.push({ order, assign, minutes, afterWeld, metal, grind: [], rest: [] });
  });

  drafts.forEach((draft) => {
    draft.afterGrind = draft.afterWeld;
    if (!(draft.minutes.Grinding > 0)) return;
    const placed = placeGrindingOnOpenSlot(
      draft.order,
      draft.minutes.Grinding,
      draft.afterWeld,
      busy,
      users
    );
    draft.grind = placed.blocks;
    draft.afterGrind = placed.endMs;
    busy = busy.concat(placed.blocks);
  });

  drafts.forEach((draft) => {
    const hadMetal = ["Profile Cutting", "Tagging", "Plate Cutting", "Welding", "Grinding"]
      .some((p) => draft.minutes[p] > 0);
    let after = draft.afterGrind;
    if (hadMetal) {
      const drop = earliestPaintMonday(after);
      const paintEnd = drop + PAINT_WAIT_DAYS * 86400000;
      const paint = {
        id: newId(),
        orderId: formatOrderId(draft.order.order_number),
        studioNo: formatOrderId(draft.order.order_number),
        product: String(draft.order.product || ""),
        process: "Powder coating",
        workerId: PAINT_WORKER_ID,
        workerName: PAINT_WORKER_NAME,
        start: isoFromMs(drop),
        end: isoFromMs(paintEnd),
        kind: "paint"
      };
      draft.rest.push(paint);
      busy = busy.concat([paint]);
      after = paintEnd;
    }
    if (draft.minutes.Assembly > 0) {
      const placed = placeProcess(
        draft.order,
        "Assembly",
        nextWorkInstant(after),
        draft.assign.Assembly,
        busy
      );
      draft.rest.push.apply(draft.rest, placed.blocks);
      busy = busy.concat(placed.blocks);
    }
  });

  const blocks = [];
  drafts.forEach((draft) => {
    blocks.push.apply(blocks, draft.metal);
    blocks.push.apply(blocks, draft.grind);
    blocks.push.apply(blocks, draft.rest);
  });
  return { blocks };
}

function findOrder(orderNumber) {
  const want = formatOrderId(orderNumber);
  return listOrders().find((o) => formatOrderId(o.order_number) === want) || null;
}

function scheduleSelected(body) {
  const ids = Array.isArray(body && body.orderIds) ? body.orderIds.map(formatOrderId).filter(Boolean) : [];
  if (!ids.length) throw new Error("Tick at least one order.");
  const assignments = (body && body.assignments) || {};
  const fromMs = body && body.from ? toMs(body.from) : Date.now();
  if (body && body.from && !Number.isFinite(fromMs)) throw new Error("From date is not a valid time.");
  const store = load();
  const remaining = store.blocks.filter((b) => ids.indexOf(formatOrderId(b.orderId)) === -1);
  const jobs = ids.map((id) => {
    const order = findOrder(id);
    if (!order) throw new Error("Order " + id + " was not found.");
    return {
      order,
      assignments: assignments[id] || assignments[order.order_number] || {}
    };
  });
  const created = scheduleBatch({ jobs, existingBlocks: remaining, fromMs }).blocks;
  store.blocks = remaining.concat(created);
  save(store);
  return { blocks: created, count: created.length };
}

function unscheduleOrder(orderNumber) {
  const id = formatOrderId(orderNumber);
  if (!id) throw new Error("Order is required.");
  const store = load();
  const before = store.blocks.length;
  store.blocks = store.blocks.filter((b) => formatOrderId(b.orderId) !== id);
  save(store);
  return { removed: before - store.blocks.length };
}

function jobKeyFor(block) {
  if (!block) return "";
  if (block.kind === "other") return "other:" + (block.jobId || block.id);
  if (block.kind === "paint") return "paint:" + formatOrderId(block.orderId);
  return "work:" + formatOrderId(block.orderId) + "||" + String(block.process || "");
}

function collectJobs(blocks) {
  const map = {};
  (blocks || []).forEach((b) => {
    if (!b) return;
    const key = jobKeyFor(b);
    if (!key) return;
    if (!map[key]) {
      map[key] = {
        key,
        kind: b.kind || "work",
        orderId: b.orderId ? formatOrderId(b.orderId) : "",
        process: String(b.process || b.title || ""),
        title: String(b.title || b.process || ""),
        product: String(b.product || ""),
        studioNo: String(b.studioNo || b.orderId || b.title || ""),
        workerId: b.workerId,
        workerName: b.workerName || b.workerId,
        jobId: b.jobId || b.id,
        minutes: Number(b.durationMinutes) > 0 ? Number(b.durationMinutes) : 0,
        start: b.start,
        end: b.end,
        counted: Number(b.durationMinutes) > 0
      };
    }
    const job = map[key];
    if (!job.counted) job.minutes += minutesBetween(toMs(b.start), toMs(b.end));
    if (String(b.start || "") < String(job.start || "")) job.start = b.start;
    if (String(b.end || "") > String(job.end || "")) job.end = b.end;
  });
  return Object.keys(map).map((k) => map[k]).filter((j) => j.kind === "paint" || j.minutes > 0);
}

function chainPreds(process, present) {
  const has = (p) => (present || []).indexOf(p) !== -1;
  if (process === "Profile Cutting") return [];
  if (process === "Tagging") return has("Profile Cutting") ? ["Profile Cutting"] : [];
  if (process === "Plate Cutting" || process === "Welding") {
    if (has("Tagging")) return ["Tagging"];
    if (has("Profile Cutting")) return ["Profile Cutting"];
    return [];
  }
  if (process === "Grinding") {
    if (has("Welding")) return ["Welding"];
    if (has("Tagging")) return ["Tagging"];
    return has("Profile Cutting") ? ["Profile Cutting"] : [];
  }
  if (process === "Powder coating") {
    if (has("Grinding")) return ["Grinding"];
    if (has("Welding")) return ["Welding"];
    if (has("Tagging")) return ["Tagging"];
    return has("Profile Cutting") ? ["Profile Cutting"] : [];
  }
  if (process === "Assembly") {
    if (has("Powder coating")) return ["Powder coating"];
    if (has("Grinding")) return ["Grinding"];
    if (has("Welding")) return ["Welding"];
    return has("Tagging") ? ["Tagging"] : (has("Profile Cutting") ? ["Profile Cutting"] : []);
  }
  return [];
}

function attachPreds(jobs) {
  const byOrder = {};
  (jobs || []).forEach((j) => {
    if (!j.orderId || j.kind === "other") return;
    byOrder[j.orderId] = byOrder[j.orderId] || [];
    if (byOrder[j.orderId].indexOf(j.process) === -1) byOrder[j.orderId].push(j.process);
  });
  (jobs || []).forEach((j) => {
    j.predKeys = [];
    if (j.kind === "other" || !j.orderId) return;
    chainPreds(j.process, byOrder[j.orderId]).forEach((p) => {
      const key = p === "Powder coating" ? "paint:" + j.orderId : "work:" + j.orderId + "||" + p;
      if ((jobs || []).some((x) => x.key === key)) j.predKeys.push(key);
    });
  });
  return jobs;
}

function defaultSequences(jobs) {
  const seq = {};
  (jobs || []).forEach((j) => {
    if (j.kind === "paint") return;
    const w = String(j.workerId || "");
    seq[w] = seq[w] || [];
    seq[w].push(j);
  });
  Object.keys(seq).forEach((w) => {
    seq[w].sort((a, b) => String(a.start).localeCompare(String(b.start)) || String(a.key).localeCompare(String(b.key)));
    seq[w] = seq[w].map((j) => j.key);
  });
  return seq;
}

function packPlan(blocks, sequences, notBefore) {
  const jobs = attachPreds(collectJobs(blocks));
  const byKey = {};
  jobs.forEach((j) => { byKey[j.key] = j; });
  const seq = sequences || defaultSequences(jobs);
  const held = notBefore || {};
  const placed = {};
  const busy = {};

  function predEnd(job) {
    let ms = Number(held[job.key]) || 0;
    (job.predKeys || []).forEach((k) => {
      if (placed[k]) ms = Math.max(ms, toMs(placed[k].end));
    });
    return ms;
  }
  function workerPrevEnd(job) {
    const list = seq[job.workerId] || [];
    const i = list.indexOf(job.key);
    if (i <= 0) return 0;
    const prev = list[i - 1];
    if (placed[prev]) return toMs(placed[prev].end);
    return null;
  }

  let guard = 0;
  while (Object.keys(placed).length < jobs.length && guard++ < 8000) {
    const left = jobs.filter((j) => !placed[j.key]);
    const cand = left.filter((j) => (j.predKeys || []).every((k) => placed[k] || !byKey[k]));
    if (!cand.length) throw new Error("Could not reflow the plan.");
    let pick = cand.filter((j) => {
      if (j.kind === "paint") return true;
      if ((j.predKeys || []).length) return true;
      return workerPrevEnd(j) !== null;
    });
    if (!pick.length) pick = cand;
    pick.sort((a, b) => {
      const ae = predEnd(a);
      const be = predEnd(b);
      if (ae !== be) return ae - be;
      return String(a.start).localeCompare(String(b.start)) || String(a.key).localeCompare(String(b.key));
    });
    const job = pick[0];
    let from = predEnd(job);
    const prevEnd = workerPrevEnd(job);
    if (prevEnd != null) from = Math.max(from, prevEnd);
    if (job.kind === "paint") {
      const drop = earliestPaintMonday(from);
      const paintEnd = drop + PAINT_WAIT_DAYS * 86400000;
      placed[job.key] = {
        start: isoFromMs(drop),
        end: isoFromMs(paintEnd),
        segments: [{ start: isoFromMs(drop), end: isoFromMs(paintEnd) }]
      };
      continue;
    }
    const wBusy = busy[job.workerId] || [];
    const result = placeTask(from, job.minutes, wBusy);
    if (!result.segments.length) {
      throw new Error("Could not place " + (job.studioNo || job.process) + ".");
    }
    placed[job.key] = result;
    busy[job.workerId] = wBusy.concat(result.segments);
  }

  const out = [];
  jobs.forEach((job) => {
    const hit = placed[job.key];
    (hit.segments || []).forEach((seg, i) => {
      out.push({
        id: i === 0 ? (job.jobId || newId()) : newId(),
        jobId: job.jobId || job.key,
        orderId: job.orderId,
        studioNo: job.kind === "other" ? job.title : job.studioNo,
        product: job.product,
        process: job.process,
        title: job.title,
        workerId: job.workerId,
        workerName: job.workerName,
        start: seg.start,
        end: seg.end,
        kind: job.kind === "paint" ? "paint" : (job.kind === "other" ? "other" : "work"),
        durationMinutes: job.kind === "other" ? job.minutes : undefined
      });
    });
  });
  return out;
}

function jobContainingBlock(blocks, blockId) {
  const raw = (blocks || []).find((b) => b && String(b.id) === String(blockId));
  if (!raw) return { raw: null, jobs: [], job: null };
  const jobs = collectJobs(blocks);
  const job = jobs.find((j) => j.key === jobKeyFor(raw)) || null;
  return { raw, jobs, job };
}

function processRank(process) {
  return JOURNEY_PROCESS_ORDER[process] || 99;
}

function laterJobsOnOrder(jobs, job) {
  if (!job || !job.orderId || job.kind === "other") return [];
  const rank = processRank(job.process);
  return (jobs || [])
    .filter((j) => j && j.orderId === job.orderId && j.key !== job.key && j.kind !== "other" && processRank(j.process) > rank)
    .sort((a, b) => processRank(a.process) - processRank(b.process) || String(a.key).localeCompare(String(b.key)));
}

function busyForWorkerExcept(blocks, workerId, excludeKeys) {
  const skip = {};
  (excludeKeys || []).forEach((k) => { skip[k] = true; });
  return (blocks || [])
    .filter((b) => b && namesEqual(b.workerId, workerId) && b.kind !== "paint" && !skip[jobKeyFor(b)])
    .map((b) => ({ start: b.start, end: b.end }));
}

function blocksFromJob(job, placed) {
  return (placed.segments || []).map((seg, i) => ({
    id: i === 0 ? (job.jobId || newId()) : newId(),
    jobId: job.jobId || job.key,
    orderId: job.orderId || "",
    studioNo: job.kind === "other" ? job.title : job.studioNo,
    product: job.product || "",
    process: job.process,
    title: job.title,
    workerId: job.workerId,
    workerName: job.workerName,
    start: seg.start,
    end: seg.end,
    kind: job.kind === "paint" ? "paint" : (job.kind === "other" ? "other" : "work"),
    durationMinutes: job.kind === "other" ? job.minutes : undefined
  }));
}

function replaceJobBlocks(blocks, job, placed) {
  const kept = (blocks || []).filter((b) => jobKeyFor(b) !== job.key);
  return kept.concat(blocksFromJob(job, placed));
}

function placeJobAt(job, fromMs, busy) {
  if (job.kind === "paint") {
    const drop = earliestPaintMonday(fromMs);
    const paintEnd = drop + PAINT_WAIT_DAYS * 86400000;
    return {
      segments: [{ start: isoFromMs(drop), end: isoFromMs(paintEnd) }],
      start: isoFromMs(drop),
      end: isoFromMs(paintEnd)
    };
  }
  const placed = placeTask(fromMs, job.minutes, busy);
  if (!placed.segments.length) {
    throw new Error("Could not place " + (job.studioNo || job.process) + ".");
  }
  return placed;
}

function moveBlock(blockId, toStart) {
  const store = load();
  const { raw, jobs, job } = jobContainingBlock(store.blocks, blockId);
  if (!raw || !job) throw new Error("That calendar block was not found.");
  if (job.kind === "paint") throw new Error("Paint shop wait cannot be dragged. Move grinding instead.");
  const dropMs = nextWorkInstant(toMs(toStart));
  if (!Number.isFinite(dropMs)) throw new Error("Drop time is not valid.");
  const followers = laterJobsOnOrder(jobs, job);
  const excludeSelf = [job.key];
  const busy = busyForWorkerExcept(store.blocks, job.workerId, excludeSelf.concat(followers.map((j) => j.key)));
  const placed = placeJobAt(job, dropMs, busy);
  if (toMs(placed.start) > dropMs + 60000) {
    throw new Error((job.workerName || "That person") + " already has work there. Pick a free slot.");
  }
  store.blocks = replaceJobBlocks(store.blocks, job, placed);
  attachPreds(jobs);
  followers.forEach((follower) => {
    const live = collectJobs(store.blocks);
    const predKeys = (attachPreds(live).find((j) => j.key === follower.key) || {}).predKeys || [];
    let from = 0;
    predKeys.forEach((k) => {
      const pred = live.find((j) => j.key === k);
      if (pred) from = Math.max(from, toMs(pred.end));
    });
    const nextFrom = nextWorkInstant(from || dropMs);
    const followerBusy = follower.kind === "paint"
      ? []
      : busyForWorkerExcept(store.blocks, follower.workerId, [follower.key]);
    const next = placeJobAt(follower, nextFrom, followerBusy);
    store.blocks = replaceJobBlocks(store.blocks, follower, next);
  });
  save(store);
  return { blocks: store.blocks, job: job.key };
}

function insertOtherTask(body) {
  const title = String((body && (body.title || body.process)) || "").trim();
  const workerId = String((body && (body.workerId || body.worker)) || "").trim();
  const minutes = Math.max(15, Math.round(Number(body && body.minutes) || 0));
  if (!title) throw new Error("Name the other task.");
  if (!workerId) throw new Error("Pick a person for the other task.");
  if (namesEqual(workerId, PAINT_WORKER_ID) || namesEqual(workerId, PAINT_WORKER_NAME)) {
    throw new Error("Other tasks go on a person, not the paint shop.");
  }
  const startMs = nextWorkInstant(body && body.start ? toMs(body.start) : Date.now());
  if (!Number.isFinite(startMs)) throw new Error("Start time is not valid.");
  const user = findUser(workerId);
  const who = user ? user.name : workerId;
  const jobId = newId();
  const store = load();
  const placed = placeTask(startMs, minutes, busyForWorker(store.blocks, who));
  if (!placed.segments.length) throw new Error("Could not place that other task.");
  if (toMs(placed.start) > startMs + 60000) {
    throw new Error(who + " already has work there. Pick a free slot.");
  }
  placed.segments.forEach((seg, i) => {
    store.blocks.push({
      id: i === 0 ? jobId : newId(),
      jobId,
      kind: "other",
      title,
      process: title,
      studioNo: title,
      orderId: "",
      product: "",
      workerId: who,
      workerName: who,
      start: seg.start,
      end: seg.end,
      durationMinutes: minutes
    });
  });
  save(store);
  return { blocks: store.blocks, jobId };
}

function removeBlock(blockId) {
  const store = load();
  const { raw, job } = jobContainingBlock(store.blocks, blockId);
  if (!raw || !job) throw new Error("That calendar block was not found.");
  if (job.kind !== "other") throw new Error("Unschedule the order to remove shop work.");
  store.blocks = store.blocks.filter((b) => jobKeyFor(b) !== job.key);
  save(store);
  return { removed: true };
}

function plannedWorkers() {
  const users = staff.listUsers();
  const fromTasks = users
    .filter((u) => (u.tasks || []).length)
    .map((u) => ({
      id: u.name,
      name: u.name,
      tasks: (u.tasks || []).slice()
    }));
  const seen = {};
  fromTasks.forEach((w) => { seen[w.id.toLowerCase()] = w; });
  load().blocks.forEach((b) => {
    if (!b || b.kind === "paint" || !b.workerId || namesEqual(b.workerId, PAINT_WORKER_ID)) return;
    const key = String(b.workerId).toLowerCase();
    if (!seen[key]) {
      seen[key] = { id: b.workerName || b.workerId, name: b.workerName || b.workerId, tasks: [] };
    }
  });
  return Object.keys(seen)
    .map((k) => seen[k])
    .sort((a, b) => String(a.name).localeCompare(String(b.name)));
}

function queueOrders() {
  const users = staff.listUsers();
  const scheduled = {};
  load().blocks.forEach((b) => {
    if (b && b.orderId) scheduled[formatOrderId(b.orderId)] = true;
  });
  return listOrders()
    .filter((o) => QUEUE_STATUSES.indexOf(String(o.status || "").trim()) !== -1)
    .map((o) => {
      const processes = PLANNED_PROCESSES.map((process) => {
        const minutes = staff.durationMinutes(o.product, process) || 0;
        const hours = minutes > 0 ? Math.round((minutes / 60) * 100) / 100 : 0;
        return {
          process,
          hours,
          minutes,
          auto: process === "Grinding",
          workers: process === "Grinding" ? [] : workersForProcess(process, users)
        };
      }).filter((p) => p.minutes > 0);
      return {
        order_number: formatOrderId(o.order_number),
        product: String(o.product || ""),
        status: String(o.status || ""),
        type: String(o.type || ""),
        category: String(o.category || ""),
        scheduled: !!scheduled[formatOrderId(o.order_number)],
        processes
      };
    });
}

function blockOverlapsWeek(block, weekStartIso) {
  const days = weekDays(weekStartIso);
  const weekStart = sastMs.apply(null, days[0].iso.split("-").map(Number).concat([0, 0]));
  const last = days[4].iso.split("-").map(Number);
  const weekEnd = sastMs(last[0], last[1], last[2], 23, 59) + 60000;
  const start = toMs(block.start);
  const end = toMs(block.end);
  return Number.isFinite(start) && Number.isFinite(end) && start < weekEnd && end > weekStart;
}

function weekStartingOrders(journey, days) {
  const isos = (days || []).map((d) => d.iso);
  return ((journey && journey.orders) || [])
    .filter((o) => isos.indexOf(String(o.start || "").slice(0, 10)) !== -1)
    .sort((a, b) => String(a.start).localeCompare(String(b.start)) || String(a.orderId).localeCompare(String(b.orderId)));
}

function getBoard(week) {
  const weekStart = weekMondayIso(week);
  const days = weekDays(weekStart);
  const store = load();
  const workers = plannedWorkers();
  const journey = buildJourney(store.blocks);
  return {
    weekStart,
    prevWeek: shiftWeek(weekStart, -1),
    nextWeek: shiftWeek(weekStart, 1),
    weekDays: days,
    workers,
    paintShop: { id: PAINT_WORKER_ID, name: PAINT_WORKER_NAME },
    blocks: store.blocks,
    queue: queueOrders(),
    colors: PROCESS_COLORS,
    processCodes: Object.assign({}, PROCESS_CODES),
    processes: PLANNED_PROCESSES.slice(),
    journeyWeeks: journeyWeeks(weekStart, JOURNEY_WEEK_COUNT),
    firstPlannedWeek: journey.days[0] ? weekMondayIso(journey.days[0].iso) : "",
    autoProcesses: ["Grinding"],
    windows: {
      morning: "07:45–12:00",
      afternoon: "12:30–15:45",
      lunch: "12:00–12:30"
    },
    bands: DAY_BANDS.slice(),
    otherTasks: OTHER_TASKS.slice(),
    now: isoFromMs(Date.now()),
    journey,
    weekStarting: weekStartingOrders(journey, days)
  };
}

module.exports = {
  ZONE,
  PAINT_WORKER_ID,
  PAINT_WORKER_NAME,
  PAINT_WAIT_DAYS,
  PLANNED_PROCESSES,
  PROCESS_COLORS,
  OTHER_TASKS,
  WINDOWS,
  DAY_BANDS,
  WEEKDAYS,
  sastMs,
  partsFromMs,
  isoFromMs,
  isoDateFromMs,
  toMs,
  nextWorkInstant,
  ceilToMinute,
  placeTask,
  placeContiguousTask,
  addWorkMinutes,
  earliestPaintMonday,
  weekMondayIso,
  shiftWeek,
  weekDays,
  scheduleOrder,
  scheduleSelected,
  unscheduleOrder,
  moveBlock,
  insertOtherTask,
  removeBlock,
  packPlan,
  getBoard,
  load,
  save,
  queueOrders,
  buildJourney,
  workdaysFromTo,
  occupiedWorkdays,
  formatDayHeader,
  JOURNEY_PROCESS_ORDER,
  PROCESS_CODES,
  JOURNEY_WEEK_COUNT,
  processCode,
  journeyWeeks,
  journeyWeeksFromPlan,
  formatWeekRange,
  grindingPool,
  USER_ASSIGNED_PROCESSES,
  weekStartingOrders,
  attachActuals,
  matchJourneyProcess
};
