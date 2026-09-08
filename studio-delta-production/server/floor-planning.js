"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir } = require("./workbook-store");
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
  "Assembly": { bg: "#fae8ff", fg: "#86198f", border: "#c026d3" }
};

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

function scheduleOrder({ order, assignments, existingBlocks, fromMs }) {
  const minutes = minutesByProcessFor(order.product);
  const hasWork = PLANNED_PROCESSES.some((p) => minutes[p] > 0);
  if (!hasWork) {
    throw new Error("No Task times for " + (order.product || order.order_number) + ". Add hours on Task times first.");
  }
  const assign = assignments || {};
  let busy = (existingBlocks || []).slice();
  const out = [];

  function run(process, from) {
    if (!(minutes[process] > 0)) return from;
    const placed = placeProcess(order, process, from, assign[process], busy);
    out.push.apply(out, placed.blocks);
    busy = busy.concat(placed.blocks);
    return placed.endMs;
  }

  const startMs = nextWorkInstant(fromMs != null ? fromMs : Date.now());
  let afterProfile = startMs;
  if (minutes["Profile Cutting"] > 0) afterProfile = run("Profile Cutting", startMs);

  let afterTag = afterProfile;
  if (minutes.Tagging > 0) afterTag = run("Tagging", afterProfile);

  let afterWeld = afterTag;
  if (minutes.Welding > 0) afterWeld = run("Welding", afterTag);
  if (minutes["Plate Cutting"] > 0) run("Plate Cutting", afterTag);

  let afterGrind = afterWeld;
  if (minutes.Grinding > 0) afterGrind = run("Grinding", afterWeld);

  const hadMetal = ["Profile Cutting", "Tagging", "Plate Cutting", "Welding", "Grinding"]
    .some((p) => minutes[p] > 0);
  if (hadMetal) {
    const drop = earliestPaintMonday(afterGrind);
    const paintEnd = drop + PAINT_WAIT_DAYS * 86400000;
    out.push({
      id: newId(),
      orderId: formatOrderId(order.order_number),
      studioNo: formatOrderId(order.order_number),
      product: String(order.product || ""),
      process: "Powder coating",
      workerId: PAINT_WORKER_ID,
      workerName: PAINT_WORKER_NAME,
      start: isoFromMs(drop),
      end: isoFromMs(paintEnd),
      kind: "paint"
    });
    afterGrind = paintEnd;
  }

  if (minutes.Assembly > 0) run("Assembly", nextWorkInstant(afterGrind));
  return out;
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
  let busy = store.blocks.filter((b) => ids.indexOf(formatOrderId(b.orderId)) === -1);
  const created = [];
  ids.forEach((id) => {
    const order = findOrder(id);
    if (!order) throw new Error("Order " + id + " was not found.");
    const assign = assignments[id] || assignments[order.order_number] || {};
    const blocks = scheduleOrder({
      order,
      assignments: assign,
      existingBlocks: busy,
      fromMs
    });
    created.push.apply(created, blocks);
    busy = busy.concat(blocks);
  });
  store.blocks = busy;
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

function plannedWorkers() {
  const users = staff.listUsers();
  const fromTasks = users
    .filter((u) => (u.tasks || []).some((t) => PLANNED_PROCESSES.indexOf(t) !== -1))
    .map((u) => ({
      id: u.name,
      name: u.name,
      tasks: (u.tasks || []).filter((t) => PLANNED_PROCESSES.indexOf(t) !== -1)
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
          workers: workersForProcess(process, users)
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

function getBoard(week) {
  const weekStart = weekMondayIso(week);
  const days = weekDays(weekStart);
  const store = load();
  const workers = plannedWorkers();
  return {
    weekStart,
    prevWeek: shiftWeek(weekStart, -1),
    nextWeek: shiftWeek(weekStart, 1),
    weekDays: days,
    workers,
    paintShop: { id: PAINT_WORKER_ID, name: PAINT_WORKER_NAME },
    blocks: store.blocks.filter((b) => blockOverlapsWeek(b, weekStart)),
    queue: queueOrders(),
    colors: PROCESS_COLORS,
    processes: PLANNED_PROCESSES.slice(),
    windows: {
      morning: "07:45–12:00",
      afternoon: "12:30–15:45",
      lunch: "12:00–12:30"
    },
    now: isoFromMs(Date.now())
  };
}

module.exports = {
  ZONE,
  PAINT_WORKER_ID,
  PAINT_WORKER_NAME,
  PAINT_WAIT_DAYS,
  PLANNED_PROCESSES,
  PROCESS_COLORS,
  WINDOWS,
  WEEKDAYS,
  sastMs,
  partsFromMs,
  isoFromMs,
  isoDateFromMs,
  toMs,
  nextWorkInstant,
  ceilToMinute,
  placeTask,
  addWorkMinutes,
  earliestPaintMonday,
  weekMondayIso,
  shiftWeek,
  weekDays,
  scheduleOrder,
  scheduleSelected,
  unscheduleOrder,
  getBoard,
  load,
  save,
  queueOrders
};
