"use strict";

const staff = require("./staff");

const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function sastParts(d) {
  const x = d instanceof Date ? d : new Date();
  const sast = new Date(x.getTime() + 2 * 60 * 60 * 1000);
  return { y: sast.getUTCFullYear(), m: sast.getUTCMonth() + 1 };
}

function currentMonth() {
  const p = sastParts(new Date());
  return p.y + "-" + String(p.m).padStart(2, "0");
}

function parseMonth(raw) {
  const s = String(raw || "").trim();
  if (/^\d{4}-(0[1-9]|1[0-2])$/.test(s)) return s;
  return currentMonth();
}

function monthTitle(ym) {
  const parts = String(ym).split("-");
  const m = Number(parts[1]) - 1;
  return (MONTH_NAMES[m] || ym) + " " + parts[0];
}

function refMs(ym) {
  const parts = String(ym).split("-");
  const y = Number(parts[0]);
  const m = Number(parts[1]);
  return Date.UTC(y, m - 1, 15, 10, 0, 0);
}

function nameKey(s) {
  return String(s || "").replace(/\s+/g, " ").trim().toLowerCase();
}

function baseOrderNo(raw) {
  const s = String(raw || "").trim().toUpperCase().replace(/\s+/g, " ");
  if (!s) return "";
  const tagged = s.match(/^(S\d+)\s*[A-Z]$/);
  if (tagged) return tagged[1];
  const glued = s.match(/^(S\d+)[A-Z]$/);
  if (glued) return glued[1];
  return s;
}

function orderLineId(raw) {
  const s = String(raw || "").trim().toUpperCase().replace(/\s+/g, " ");
  if (!s) return "";
  const tagged = s.match(/^(S\d+)\s*([A-Z])$/);
  if (tagged) return tagged[1] + " " + tagged[2];
  return s;
}

function isIndirectTask(t) {
  if (!t) return true;
  if (String(t.entryType || "") === "indirect") return true;
  if (!String(t.order || "").trim()) return true;
  return String(t.role || t.task || "").toLowerCase() === "indirect";
}

function isShopUser(u) {
  if (!u || !u.name) return false;
  const access = String(u.access || "");
  if (access === "Marketing") return false;
  const tasks = Array.isArray(u.tasks) ? u.tasks : [];
  if (tasks.length) return true;
  return access === "Production";
}

function salesSeatOf(name) {
  if (staff.isDrawingOwnerName(name)) return "erin";
  if (staff.isDrawingApprovalOwnerName(name)) return "shabana";
  return "";
}

function isSalesUser(u) {
  if (!u || !u.name) return false;
  return !!salesSeatOf(u.name);
}

function salesRoster() {
  const users = staff.listUsers() || [];
  const erin = users.find((u) => staff.isDrawingOwnerName(u.name));
  const shabana = users.find((u) => staff.isDrawingApprovalOwnerName(u.name));
  return [
    { seat: "erin", user: erin || { name: "Erin", role: "Interior Designer" } },
    { seat: "shabana", user: shabana || { name: "Shabana", role: "Clients Relations Manager" } }
  ];
}

function isoMonth(raw) {
  const n = Date.parse(raw);
  if (!Number.isFinite(n)) return "";
  const p = sastParts(new Date(n));
  return p.y + "-" + String(p.m).padStart(2, "0");
}

function durationLabel(ms) {
  if (!Number.isFinite(ms) || ms < 0) return "";
  const mins = Math.round(ms / 60000);
  if (mins < 1) return "same time";
  if (mins < 60) return mins + " min";
  const hours = Math.floor(mins / 60);
  const rem = mins % 60;
  if (hours < 48) return hours + "h" + (rem ? " " + rem + "m" : "");
  const days = Math.floor(hours / 24);
  const h = hours % 24;
  return days + "d" + (h ? " " + h + "h" : "");
}

function median(nums) {
  const s = (nums || []).slice().filter((n) => Number.isFinite(n)).sort((a, b) => a - b);
  if (!s.length) return 0;
  const mid = Math.floor(s.length / 2);
  return s.length % 2 ? s[mid] : Math.round((s[mid - 1] + s[mid]) / 2);
}

const SALES_KINDS = [
  { kind: "created", key: "captured", label: "Captured" },
  { kind: "complete_quote", key: "quoted", label: "Quoted" },
  { kind: "complete_followup", key: "followed", label: "Followed up" },
  { kind: "complete_order", key: "ordered", label: "Client ordered" },
  { kind: "create_orders", key: "added", label: "Added to Orders" },
  { kind: "complete_drawing", key: "drawing", label: "Drawings" },
  { kind: "complete_drawing_approval", key: "approval", label: "Drawing approval" }
];

function salesKindOf(kind) {
  return SALES_KINDS.find((k) => k.kind === kind) || null;
}

function emptySalesPerson(name, profile) {
  return {
    worker: name,
    role: (profile && (profile.role || profile.jobTitle)) || "Sales",
    worked: false,
    captured: 0,
    quoted: 0,
    followed: 0,
    ordered: 0,
    added: 0,
    drawing: 0,
    approval: 0,
    kinds: SALES_KINDS.map((k) => ({ name: k.label, key: k.key, count: 0 })),
    rows: []
  };
}

function buildSales(month) {
  const db = require("./db");
  const created = require("./order-created");
  const byName = {};
  salesRoster().forEach((seat) => {
    byName[seat.seat] = emptySalesPerson(seat.user.name, seat.user);
  });

  const orders = db.listOrders() || [];
  (db.listEnquiries() || []).forEach((row) => {
    const list = Array.isArray(row.events) ? row.events : [];
    list.forEach((ev) => {
      const spec = salesKindOf(ev && ev.kind);
      if (!spec) return;
      if (isoMonth(ev.at) !== month) return;
      const actor = String((ev && ev.actor) || "").trim();
      if (!actor) return;
      const key = salesSeatOf(actor);
      if (!key || !byName[key]) return;
      const person = byName[key];
      person.worked = true;
      person[spec.key] = (person[spec.key] || 0) + 1;
      const kindRow = person.kinds.find((k) => k.key === spec.key);
      if (kindRow) kindRow.count += 1;
      person.rows.push({
        kind: spec.key,
        kindLabel: spec.label,
        enquiry_no: row.enquiry_no || "",
        client: row.client_name || "",
        product: row.product || "",
        order: row.order_number || "",
        at: ev.at || "",
        atLabel: ev.at_label || db.formatSastDateTime(ev.at) || ev.at || ""
      });
    });
  });

  const added = [];
  const lags = [];
  (created.list() || []).forEach((rec) => {
    if (!rec || isoMonth(rec.at) !== month) return;
    const start = Date.parse(rec.readyAt);
    const end = Date.parse(rec.at);
    const lagMs = Number.isFinite(start) && Number.isFinite(end) ? Math.max(0, end - start) : null;
    if (lagMs != null) lags.push(lagMs);
    const orderMeta = orders.find((o) => orderLineId(o.order_number) === orderLineId(rec.order_number));
    added.push({
      order: rec.order_number,
      product: rec.product || (orderMeta && orderMeta.product) || "",
      client: rec.client_name || (orderMeta && orderMeta.client_name) || "",
      by: rec.by || "",
      enquiry_no: rec.enquiryNo || "",
      readyAt: rec.readyAt || "",
      readyLabel: rec.readyAt ? (db.formatSastDateTime(rec.readyAt) || rec.readyAt) : "",
      at: rec.at || "",
      atLabel: rec.at ? (db.formatSastDateTime(rec.at) || rec.at) : "",
      lagMs,
      lagLabel: lagMs == null ? "Not timed" : durationLabel(lagMs)
    });
  });
  added.sort((a, b) => String(a.at).localeCompare(String(b.at)));

  const people = Object.keys(byName).map((k) => byName[k]);
  people.forEach((p) => {
    p.kinds = p.kinds.filter((k) => k.count > 0);
    p.rows.sort((a, b) => String(a.at).localeCompare(String(b.at)));
  });
  people.sort((a, b) => {
    const aN = a.captured + a.quoted + a.followed + a.ordered + a.added + a.drawing + a.approval;
    const bN = b.captured + b.quoted + b.followed + b.ordered + b.added + b.drawing + b.approval;
    if (bN !== aN) return bN - aN;
    return String(a.worker).localeCompare(String(b.worker));
  });

  const worked = people.filter((p) => p.worked);
  return {
    people,
    workedCount: worked.length,
    idleCount: people.filter((p) => !p.worked).length,
    headcount: people.length,
    captured: people.reduce((s, p) => s + p.captured, 0),
    quoted: people.reduce((s, p) => s + p.quoted, 0),
    followed: people.reduce((s, p) => s + p.followed, 0),
    ordered: people.reduce((s, p) => s + p.ordered, 0),
    drawing: people.reduce((s, p) => s + p.drawing, 0),
    approval: people.reduce((s, p) => s + p.approval, 0),
    addedCount: added.length,
    timedCount: lags.length,
    untimedCount: added.length - lags.length,
    medianLagMs: median(lags),
    medianLagLabel: lags.length ? durationLabel(median(lags)) : "",
    averageLagMs: lags.length ? Math.round(lags.reduce((s, n) => s + n, 0) / lags.length) : 0,
    averageLagLabel: lags.length ? durationLabel(Math.round(lags.reduce((s, n) => s + n, 0) / lags.length)) : "",
    added
  };
}

function orderMetaMap() {
  const db = require("./db");
  const catalog = require("./product-catalog");
  const map = {};
  (db.listOrders() || []).forEach((o) => {
    const id = db.formatOrderId(o.order_number);
    if (!id) return;
    const product = String(o.product || "").trim();
    const found = product ? catalog.lookupProduct(product) : null;
    const rec = {
      product,
      imageUrl: (found && found.imageUrl) || ""
    };
    map[id] = rec;
    const base = baseOrderNo(id);
    if (base && !map[base]) map[base] = rec;
  });
  return map;
}

function decorateTasks(tasks, orderMap) {
  const db = require("./db");
  return (tasks || []).map((t) => {
    const order = String(t.order || "").trim();
    const formatted = order ? db.formatOrderId(order) : "";
    const meta = formatted ? (orderMap[formatted] || orderMap[baseOrderNo(formatted)] || orderMap[order] || null) : null;
    return {
      order: formatted || order,
      orderLine: formatted ? orderLineId(formatted) : "",
      baseOrder: formatted ? baseOrderNo(formatted) : "",
      role: t.role || "",
      task: t.task || t.role || "",
      minutes: Number(t.minutes || 0),
      date: t.date || "",
      start: t.start || null,
      end: t.end || null,
      entryType: t.entryType || "production",
      product: (meta && meta.product) || "",
      imageUrl: (meta && meta.imageUrl) || ""
    };
  });
}

function summarize(w, profile, orderMap) {
  const processes = {};
  const orders = {};
  let finished = 0;
  let running = 0;
  let indirectMins = 0;
  let productionMins = 0;
  const tasks = decorateTasks(w.tasks || [], orderMap || {});
  tasks.forEach((t) => {
    const mins = Number(t.minutes || 0);
    const role = String(t.role || t.task || "Other").trim() || "Other";
    const indirect = isIndirectTask(t);
    processes[role] = processes[role] || { name: role, minutes: 0, clocks: 0, orders: {}, tasks: {} };
    processes[role].minutes += mins;
    processes[role].clocks += 1;
    if (indirect) {
      indirectMins += mins;
      const taskName = String(t.task || t.role || "Indirect").trim() || "Indirect";
      processes[role].tasks[taskName] = true;
    } else {
      productionMins += mins;
      const order = t.orderLine || orderLineId(t.order);
      if (order) {
        orders[order] = (orders[order] || 0) + mins;
        processes[role].orders[order] = true;
      }
    }
    if (t.end) finished += 1;
    else running += 1;
  });
  const processList = Object.keys(processes).map((k) => {
    const row = processes[k];
    const indirect = String(row.name || "").toLowerCase() === "indirect";
    return {
      name: row.name,
      minutes: row.minutes,
      clocks: row.clocks,
      orderCount: Object.keys(row.orders).length,
      taskCount: Object.keys(row.tasks).length,
      count: indirect ? Object.keys(row.tasks).length : Object.keys(row.orders).length,
      indirect
    };
  }).sort((a, b) => b.minutes - a.minutes);
  return {
    worker: w.worker,
    role: (profile && (profile.role || profile.jobTitle)) || "",
    clocked: (w.totalMinutes || 0) > 0 || tasks.length > 0,
    daysWorked: (w.days || []).length,
    totalMinutes: Number(w.totalMinutes || 0),
    totalRegular: Number(w.totalRegular || 0),
    totalOvertime: Number(w.totalOvertime || 0),
    overLimitDays: (w.days || []).filter((d) => d.overLimit).length,
    processes: processList,
    orderCount: Object.keys(orders).length,
    finished,
    running,
    indirectMins,
    productionMins,
    tasks
  };
}

function emptyWorker(name, profile, orderMap) {
  return summarize({
    worker: name,
    days: [],
    tasks: [],
    totalMinutes: 0,
    totalRegular: 0,
    totalOvertime: 0
  }, profile, orderMap);
}

async function build(query) {
  const month = parseMonth(query && query.month);
  const { callShopFunction } = require("./gas");
  const report = await callShopFunction("getActivityReport", ["month", refMs(month), ""]);
  const orderMap = orderMetaMap();
  const byName = {};
  (report.workers || []).forEach((w) => {
    byName[nameKey(w.worker)] = w;
  });
  const users = staff.listUsers();
  const profiles = {};
  users.forEach((u) => { profiles[nameKey(u.name)] = u; });

  const seen = {};
  const people = [];
  users.filter(isShopUser).forEach((u) => {
    const key = nameKey(u.name);
    seen[key] = true;
    const rec = byName[key];
    people.push(rec ? summarize(rec, u, orderMap) : emptyWorker(u.name, u, orderMap));
  });
  (report.workers || []).forEach((w) => {
    const key = nameKey(w.worker);
    if (seen[key]) return;
    seen[key] = true;
    people.push(summarize(w, profiles[key], orderMap));
  });
  people.sort((a, b) => {
    if (b.totalMinutes !== a.totalMinutes) return b.totalMinutes - a.totalMinutes;
    return String(a.worker).localeCompare(String(b.worker));
  });

  const clocked = people.filter((p) => p.clocked);
  const absent = people.filter((p) => !p.clocked);
  const processTotals = {};
  clocked.forEach((p) => {
    p.processes.forEach((pr) => {
      processTotals[pr.name] = processTotals[pr.name] || { name: pr.name, minutes: 0, n: 0, people: 0 };
      processTotals[pr.name].minutes += pr.minutes;
      processTotals[pr.name].n += pr.n;
      processTotals[pr.name].people += 1;
    });
  });
  const stations = Object.keys(processTotals).map((k) => processTotals[k]).sort((a, b) => b.minutes - a.minutes);
  const busiest = clocked[0] || null;
  const topStation = stations[0] || null;
  const uniqueOrders = {};
  clocked.forEach((p) => {
    (p.tasks || []).forEach((t) => {
      if (isIndirectTask(t)) return;
      const id = t.orderLine || orderLineId(t.order);
      if (id) uniqueOrders[id] = true;
    });
  });

  return {
    month,
    title: monthTitle(month),
    standardDayMins: report.standardDayMins || 450,
    people,
    clockedCount: clocked.length,
    absentCount: absent.length,
    headcount: people.length,
    totalMinutes: clocked.reduce((s, p) => s + p.totalMinutes, 0),
    totalRegular: clocked.reduce((s, p) => s + p.totalRegular, 0),
    totalOvertime: clocked.reduce((s, p) => s + p.totalOvertime, 0),
    orderCount: Object.keys(uniqueOrders).length,
    stations,
    busiest: busiest ? { worker: busiest.worker, minutes: busiest.totalMinutes } : null,
    topStation: topStation ? { name: topStation.name, minutes: topStation.minutes } : null,
    sales: buildSales(month)
  };
}

module.exports = {
  parseMonth,
  monthTitle,
  currentMonth,
  isShopUser,
  isSalesUser,
  baseOrderNo,
  orderLineId,
  isIndirectTask,
  summarize,
  orderMetaMap,
  durationLabel,
  buildSales,
  build
};
