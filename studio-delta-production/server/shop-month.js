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
    topStation: topStation ? { name: topStation.name, minutes: topStation.minutes } : null
  };
}

module.exports = {
  parseMonth,
  monthTitle,
  currentMonth,
  isShopUser,
  baseOrderNo,
  orderLineId,
  isIndirectTask,
  summarize,
  orderMetaMap,
  build
};
