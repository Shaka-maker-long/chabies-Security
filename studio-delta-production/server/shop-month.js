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

function isShopUser(u) {
  if (!u || !u.name) return false;
  const access = String(u.access || "");
  if (access === "Marketing") return false;
  const tasks = Array.isArray(u.tasks) ? u.tasks : [];
  if (tasks.length) return true;
  return access === "Production";
}

function summarize(w, profile) {
  const processes = {};
  const orders = {};
  let finished = 0;
  let running = 0;
  let indirectMins = 0;
  let productionMins = 0;
  (w.tasks || []).forEach((t) => {
    const mins = Number(t.minutes || 0);
    const role = String(t.role || t.task || "Other").trim() || "Other";
    processes[role] = processes[role] || { name: role, minutes: 0, n: 0 };
    processes[role].minutes += mins;
    processes[role].n += 1;
    const indirect = t.entryType === "indirect" || !t.order;
    if (indirect) indirectMins += mins;
    else {
      productionMins += mins;
      const order = String(t.order || "").trim();
      if (order) orders[order] = (orders[order] || 0) + mins;
    }
    if (t.end) finished += 1;
    else running += 1;
  });
  const processList = Object.keys(processes).map((k) => processes[k]).sort((a, b) => b.minutes - a.minutes);
  const orderList = Object.keys(orders).map((order) => ({ order, minutes: orders[order] }))
    .sort((a, b) => b.minutes - a.minutes);
  return {
    worker: w.worker,
    role: (profile && (profile.role || profile.jobTitle)) || "",
    clocked: (w.totalMinutes || 0) > 0 || (w.tasks || []).length > 0,
    daysWorked: (w.days || []).length,
    totalMinutes: Number(w.totalMinutes || 0),
    totalRegular: Number(w.totalRegular || 0),
    totalOvertime: Number(w.totalOvertime || 0),
    overLimitDays: (w.days || []).filter((d) => d.overLimit).length,
    processes: processList,
    orders: orderList,
    orderCount: orderList.length,
    finished,
    running,
    indirectMins,
    productionMins,
    days: w.days || [],
    tasks: w.tasks || []
  };
}

function emptyWorker(name, profile) {
  return summarize({
    worker: name,
    days: [],
    tasks: [],
    totalMinutes: 0,
    totalRegular: 0,
    totalOvertime: 0
  }, profile);
}

async function build(query) {
  const month = parseMonth(query && query.month);
  const { callShopFunction } = require("./gas");
  const report = await callShopFunction("getActivityReport", ["month", refMs(month), ""]);
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
    people.push(rec ? summarize(rec, u) : emptyWorker(u.name, u));
  });
  (report.workers || []).forEach((w) => {
    const key = nameKey(w.worker);
    if (seen[key]) return;
    seen[key] = true;
    people.push(summarize(w, profiles[key]));
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
  clocked.forEach((p) => p.orders.forEach((o) => { uniqueOrders[o.order] = true; }));

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
  summarize,
  build
};
