"use strict";

const { getBook, persistWorkbook } = require("./workbook-store");
const { listOrders, parseMoney, money, formatRand } = require("./db");
const catalog = require("./product-catalog");
const steelRates = require("./steel-rates");
const { FLOOR_TASKS } = require("./staff");

const TASKS = ["Profile Cutting", "Plate Cutting", "Tagging", "Welding", "Grinding", "Assembly"];
const OVERHEAD_PCT = 0.4;
const RATES_HEADER = ["Employee Name", "Hourly Rate"];
const DEFAULT_EMPLOYEE_RATES = [
  { employee: "Willard", ratePerHour: 110 },
  { employee: "John", ratePerHour: 65.16 },
  { employee: "Admire", ratePerHour: 49.41 },
  { employee: "Uriah", ratePerHour: 32.97 },
  { employee: "Sam", ratePerHour: 46.16 },
  { employee: "Thabile", ratePerHour: 39.56 }
];
const STATION_NAMES = new Set(TASKS.concat(FLOOR_TASKS).map((name) => name.toLowerCase()));

function sastYmd(date) {
  if (!date || isNaN(date.getTime())) return "";
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Africa/Johannesburg",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).format(date);
}

function monthKey(date) {
  const ymd = sastYmd(date);
  return ymd ? ymd.slice(0, 7) : "";
}

function sastDate(y, m, d, h, min, sec) {
  return new Date(Date.UTC(Number(y), Number(m) - 1, Number(d), Number(h) - 2, Number(min) || 0, Number(sec) || 0));
}

function parseFilterRange(mode, from, to) {
  const kind = String(mode || "month").trim().toLowerCase();
  if (kind === "all" || !from) return { kind: "all", start: null, end: null };
  if (kind === "month") {
    const [y, m] = String(from).split("-").map(Number);
    if (!y || !m) return { kind: "all", start: null, end: null };
    const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
    return { kind, start: sastDate(y, m, 1, 0, 0, 0), end: sastDate(y, m, last, 23, 59, 59) };
  }
  if (kind === "custom") {
    const a = String(from).slice(0, 10).split("-").map(Number);
    const b = String(to || from).slice(0, 10).split("-").map(Number);
    if (a.length < 3 || b.length < 3) return { kind: "all", start: null, end: null };
    return { kind, start: sastDate(a[0], a[1], a[2], 0, 0, 0), end: sastDate(b[0], b[1], b[2], 23, 59, 59) };
  }
  if (kind === "week") {
    const parts = String(from).slice(0, 10).split("-").map(Number);
    if (parts.length < 3) return { kind: "all", start: null, end: null };
    const seed = sastDate(parts[0], parts[1], parts[2], 12, 0, 0);
    const dow = Number(new Intl.DateTimeFormat("en-US", { timeZone: "Africa/Johannesburg", weekday: "short" }).format(seed)
      .replace("Sun", "0").replace("Mon", "1").replace("Tue", "2").replace("Wed", "3").replace("Thu", "4").replace("Fri", "5").replace("Sat", "6"));
    const start = new Date(seed.getTime() - dow * 86400000);
    start.setUTCHours(22, 0, 0, 0);
    const ymd = sastYmd(start).split("-");
    const startDay = sastDate(ymd[0], ymd[1], ymd[2], 0, 0, 0);
    const endDay = new Date(startDay.getTime() + 7 * 86400000 - 1);
    return { kind, start: startDay, end: endDay };
  }
  const parts = String(from).slice(0, 10).split("-").map(Number);
  if (parts.length < 3) return { kind: "all", start: null, end: null };
  return { kind: "day", start: sastDate(parts[0], parts[1], parts[2], 0, 0, 0), end: sastDate(parts[0], parts[1], parts[2], 23, 59, 59) };
}

function inRange(date, range) {
  if (!range || !range.start) return true;
  if (!date || isNaN(date.getTime())) return false;
  return date.getTime() >= range.start.getTime() && date.getTime() <= range.end.getTime();
}

function matchTask(task) {
  const s = String(task || "").toLowerCase();
  if (!s || /pre-powder|final qc|quality control|paint prep|painting|powder coating/.test(s)) return "";
  return TASKS.find((name) => s.indexOf(name.toLowerCase().split(" ")[0]) !== -1) || "";
}

function employeeNameFromBody(body) {
  return String((body && (body.employee || body.name || body.process)) || "").trim();
}

function looksLikeStationRateSheet(rows) {
  if (!rows.length) return true;
  return rows.every((row) => STATION_NAMES.has(String(row.employee || "").toLowerCase()));
}

function writeRatesHeader(sheet, force) {
  if (!force) {
    const header = sheet.getRange(1, 1, 1, 2).getValues()[0] || [];
    if (String(header[0] || "").trim() === RATES_HEADER[0] && String(header[1] || "").trim() === RATES_HEADER[1]) {
      return false;
    }
  }
  sheet.getRange(1, 1, 1, 2).setValues([RATES_HEADER]);
  return true;
}

function getOrCreateRatesSheet() {
  const book = getBook();
  let sheet = book.getSheetByName("Rates");
  if (!sheet) {
    sheet = book.insertSheet("Rates");
    writeRatesHeader(sheet);
    persistWorkbook();
  }
  return sheet;
}

function readLabourRates(opts) {
  if (!opts || !opts.skipEnsure) ensureDefaultEmployeeRates();
  const sheet = getOrCreateRatesSheet();
  const byEmployee = {};
  const rows = [];
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach((row, i) => {
      const employee = String(row[0] || "").trim();
      if (!employee) return;
      const rate = parseMoney(row[1]);
      byEmployee[employee.toLowerCase()] = Number.isFinite(rate) ? rate : 0;
      rows.push({
        id: "lrate_" + (i + 2),
        employee,
        process: employee,
        ratePerHour: money(rate || 0),
        rateLabel: formatRand(rate || 0) + " / h",
        row: i + 2
      });
    });
  }
  return { byEmployee, byProcess: byEmployee, rows };
}

function knownEmployees(rateRows) {
  const seen = {};
  const out = [];
  function add(name) {
    const n = String(name || "").trim();
    if (!n) return;
    const key = n.toLowerCase();
    if (seen[key] || STATION_NAMES.has(key)) return;
    seen[key] = true;
    out.push(n);
  }
  DEFAULT_EMPLOYEE_RATES.forEach((row) => add(row.employee));
  (rateRows || []).forEach((row) => add(row.employee));
  try {
    require("./staff").listUsers().forEach((user) => add(user.name));
  } catch (e) {}
  return out;
}

function ensureDefaultEmployeeRates() {
  const { rows } = readLabourRates({ skipEnsure: true });
  const sheet = getOrCreateRatesSheet();
  if (rows.length && !looksLikeStationRateSheet(rows)) {
    if (writeRatesHeader(sheet)) persistWorkbook();
    return;
  }
  writeRatesHeader(sheet);
  for (let r = sheet.getLastRow(); r >= 2; r--) sheet.deleteRow(r);
  DEFAULT_EMPLOYEE_RATES.forEach((row) => {
    sheet.appendRow([row.employee, Number(money(row.ratePerHour))]);
  });
  persistWorkbook();
  try { require("./gas").clearShopCache(); } catch (e) {}
}

function snapshotLabourRates() {
  const { rows } = readLabourRates();
  const employees = knownEmployees(rows);
  return { rates: rows, employees, processes: employees };
}

function upsertLabourRate(body) {
  const employee = employeeNameFromBody(body);
  if (!employee) throw new Error("Employee name is required.");
  if (body == null || body.ratePerHour === "" || body.ratePerHour == null) {
    throw new Error("Hourly rate is required.");
  }
  const rate = parseMoney(body.ratePerHour);
  if (!Number.isFinite(rate) || rate < 0) throw new Error("Hourly rate must be 0 or more.");
  ensureDefaultEmployeeRates();
  const sheet = getOrCreateRatesSheet();
  writeRatesHeader(sheet);
  if (sheet.getLastRow() >= 2) {
    const grid = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < grid.length; i++) {
      if (String(grid[i][0] || "").trim().toLowerCase() === employee.toLowerCase()) {
        sheet.getRange(i + 2, 1).setValue(employee);
        sheet.getRange(i + 2, 2).setValue(Number(money(rate)));
        persistWorkbook();
        try { require("./gas").clearShopCache(); } catch (e) {}
        return { employee, process: employee, ratePerHour: money(rate) };
      }
    }
  }
  sheet.appendRow([employee, Number(money(rate))]);
  persistWorkbook();
  try { require("./gas").clearShopCache(); } catch (e) {}
  return { employee, process: employee, ratePerHour: money(rate) };
}

function deleteLabourRate(employee) {
  const want = String(employee || "").trim().toLowerCase();
  if (!want) throw new Error("Rate not found.");
  const sheet = getBook().getSheetByName("Rates");
  if (!sheet || sheet.getLastRow() < 2) throw new Error("Rate not found.");
  const grid = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < grid.length; i++) {
    if (String(grid[i][0] || "").trim().toLowerCase() === want) {
      sheet.deleteRow(i + 2);
      persistWorkbook();
      try { require("./gas").clearShopCache(); } catch (e) {}
      return true;
    }
  }
  throw new Error("Rate not found.");
}

function deleteSheetRowsFromBottom(sheet, shouldDelete) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  let removed = 0;
  for (let r = sheet.getLastRow(); r >= 2; r--) {
    if (!shouldDelete || shouldDelete(r)) {
      sheet.deleteRow(r);
      removed++;
    }
  }
  return removed;
}

function clearSteelUsage() {
  const removed = deleteSheetRowsFromBottom(getBook().getSheetByName("Steel_Usage"));
  persistWorkbook();
  try { require("./gas").clearShopCache(); } catch (e) {}
  return { removed };
}

function clearProductionLogs() {
  const book = getBook();
  const removed = deleteSheetRowsFromBottom(book.getSheetByName("Production_Log"));
  const overviewRemoved = deleteSheetRowsFromBottom(book.getSheetByName("Overview"));
  persistWorkbook();
  try { require("./gas").clearShopCache(); } catch (e) {}
  return { removed, overviewRemoved };
}

function clearFinishedProductionLogs() {
  return clearProductionLogs();
}

function hourlyRate(byEmployee, worker) {
  const key = String(worker || "").trim().toLowerCase();
  if (!key) return 0;
  return byEmployee[key] != null ? byEmployee[key] : 0;
}

function readSteelUsage() {
  const sheet = getBook().getSheetByName("Steel_Usage");
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues().map((row) => {
    const ts = row[0] ? new Date(row[0]) : null;
    return {
      timestamp: ts && !isNaN(ts.getTime()) ? ts : null,
      orderNum: String(row[1] || "").trim(),
      worker: String(row[2] || "").trim(),
      process: String(row[3] || "").trim(),
      type: String(row[4] || "").trim(),
      size: row[5]
    };
  }).filter((row) => row.orderNum);
}

function readBackboardUsage() {
  const sheet = getBook().getSheetByName("Backboard_Usage");
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues().map((row) => {
    const ts = row[0] ? new Date(row[0]) : null;
    return {
      timestamp: ts && !isNaN(ts.getTime()) ? ts : null,
      orderNum: String(row[1] || "").trim(),
      worker: String(row[2] || "").trim(),
      process: String(row[3] || "").trim(),
      type: String(row[4] || "").trim(),
      size: row[5]
    };
  }).filter((row) => row.orderNum);
}

function readWoodUsage() {
  const sheet = getBook().getSheetByName("Wood_To_Order");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lastCol = Math.max(sheet.getLastColumn(), 11);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues().map((row) => {
    const ts = row[1] ? new Date(row[1]) : null;
    const type = String(row[5] || "").trim();
    const orderNum = String(row[2] || "").trim();
    if (!orderNum || !type) return null;
    return {
      timestamp: ts && !isNaN(ts.getTime()) ? ts : null,
      orderNum,
      worker: String(row[3] || "").trim(),
      component: String(row[4] || "").trim(),
      type,
      thickness: String(row[6] || "").trim(),
      height: row[7],
      width: row[8],
      quantity: row[9],
      status: String(row[10] || "").trim()
    };
  }).filter(Boolean);
}

function parseSoldAt(order) {
  const month = String((order && order.month_of_sale) || "").trim();
  const m = month.match(/^(\d{4})-(\d{2})/);
  if (m) return sastDate(m[1], m[2], 15, 12, 0, 0);
  const pay = String((order && order.payment_date) || "").slice(0, 10);
  const p = pay.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (p) return sastDate(p[1], p[2], p[3], 12, 0, 0);
  return null;
}

function orderMeta() {
  const byId = {};
  listOrders().forEach((order) => {
    const id = String(order.order_number || "").trim();
    if (!id) return;
    const found = catalog.lookupProduct(order.product);
    byId[id] = {
      product: order.product || "Unknown product",
      sellingPrice: parseMoney(order.price_excl_vat) || 0,
      img: (found && found.imageUrl) || "",
      soldAt: parseSoldAt(order)
    };
  });
  return byId;
}

function emptyOrder(id, meta) {
  const info = meta[id] || { product: "Unknown product", sellingPrice: 0, img: "", soldAt: null };
  const tasks = {};
  TASKS.forEach((t) => { tasks[t] = { h: 0, c: 0 }; });
  return {
    orderNum: id,
    product: info.product,
    sellingPrice: info.sellingPrice,
    img: info.img,
    totalHours: 0,
    laborCost: 0,
    materialCost: 0,
    glassCost: 0,
    powderCost: 0,
    woodCost: 0,
    backboardCost: 0,
    materialEntries: [],
    tasks,
    staff: {},
    warnings: [],
    monthHours: {},
    monthCosts: {},
    monthMaterials: {},
    soldAt: info.soldAt || null
  };
}

function addHours(order, month, hours, cost) {
  order.totalHours += hours;
  order.laborCost += cost;
  order.monthHours[month] = (order.monthHours[month] || 0) + hours;
  order.monthCosts[month] = (order.monthCosts[month] || 0) + cost;
}

async function loadLogs() {
  const { callShopFunction } = require("./gas");
  const rows = await callShopFunction("listCostingLogs", []);
  return Array.isArray(rows) ? rows : [];
}

function addMaterial(rec, entry) {
  rec.materialEntries.push(entry);
  rec.materialCost += entry.cost;
  if (entry.type === "glass") rec.glassCost += entry.cost;
  if (entry.type === "powder") rec.powderCost += entry.cost;
  if (entry.type === "wood") rec.woodCost += entry.cost;
  if (entry.type === "backboard") rec.backboardCost += entry.cost;
  if (entry.month && entry.month !== "unknown") {
    rec.monthMaterials[entry.month] = (rec.monthMaterials[entry.month] || 0) + entry.cost;
  }
}

function addGlassActuals(order) {
  let snap;
  try { snap = require("./glass-po").snapshot(); } catch (e) { return; }
  (snap.glass || []).forEach((line) => {
    const cost = parseMoney(line.cost);
    if (!(cost > 0)) return;
    const orderNum = String(line.order || "").trim();
    if (!orderNum) return;
    const rec = order(orderNum);
    const month = line.receivedAt ? monthKey(new Date(line.receivedAt)) : "unknown";
    addMaterial(rec, {
      type: "glass",
      item: [line.type, line.thickness].filter(Boolean).join(" ") || "Glass",
      qty: [line.height, line.width].filter(Boolean).join(" × ") + (line.quantity ? " × " + line.quantity : ""),
      cost,
      month
    });
  });
}

function addPowderActuals(order) {
  let shop;
  try { shop = require("./powder-shop").loadShop(); } catch (e) { return; }
  Object.keys(shop.orders || {}).forEach((orderNum) => {
    const meta = shop.orders[orderNum] || {};
    const cost = parseMoney(meta.cost);
    if (!(cost > 0)) return;
    const rec = order(orderNum);
    const month = meta.receivedAt ? monthKey(new Date(meta.receivedAt)) : "unknown";
    addMaterial(rec, {
      type: "powder",
      item: "Powder coating",
      qty: "",
      cost,
      month
    });
  });
}

async function getAppData(query) {
  const q = query || {};
  const range = parseFilterRange(q.mode, q.from || q.value, q.to || q.end);
  const { byEmployee } = readLabourRates();
  const meta = orderMeta();
  const byId = {};
  const missingEmployeeRates = {};

  function order(id) {
    if (!byId[id]) byId[id] = emptyOrder(id, meta);
    return byId[id];
  }

  const logs = await loadLogs();
  logs.forEach((log) => {
    const start = log.start ? new Date(log.start) : null;
    const end = log.end ? new Date(log.end) : null;
    if (!start || isNaN(start.getTime()) || !end || isNaN(end.getTime())) return;
    const hours = (Number(log.minutes) || 0) / 60;
    if (!(hours > 0)) return;
    const month = monthKey(start);
    const rate = hourlyRate(byEmployee, log.worker);
    const cost = hours * rate;
    const rec = order(log.orderNum);
    addHours(rec, month, hours, cost);
    const name = log.worker || "Unknown";
    if (!(rate > 0) && name && name !== "Unknown") missingEmployeeRates[name] = true;
    if (range.kind === "all" || inRange(start, range)) {
      if (!rec.staff[name]) rec.staff[name] = { h: 0, c: 0 };
      rec.staff[name].h += hours;
      rec.staff[name].c += cost;
    }
    const cat = matchTask(log.task);
    if (cat && inRange(start, range)) {
      rec.tasks[cat].h += hours;
      rec.tasks[cat].c += cost;
      if (hours > 8) {
        rec.warnings.push({
          task: cat,
          hours,
          employee: name,
          message: cat + " took " + hours.toFixed(1) + "h (" + name + ")"
        });
      }
    }
  });

  readSteelUsage().forEach((row) => {
    const ts = row.timestamp;
    const month = ts ? monthKey(ts) : "unknown";
    const priced = steelRates.costUsage(row.type, row.size);
    const rec = order(row.orderNum);
    rec.materialEntries.push({
      type: "steel",
      item: row.type || "Steel",
      qty: String(row.size == null ? "" : row.size),
      cost: priced.cost,
      month,
      rateMissing: priced.rateMissing,
      worker: row.worker,
      process: row.process
    });
    rec.materialCost += priced.cost;
    if (month && month !== "unknown") {
      rec.monthMaterials[month] = (rec.monthMaterials[month] || 0) + priced.cost;
    }
  });

  addGlassActuals(order);
  addPowderActuals(order);

  readBackboardUsage().forEach((row) => {
    const ts = row.timestamp;
    const month = ts ? monthKey(ts) : "unknown";
    addMaterial(order(row.orderNum), {
      type: "backboard",
      item: row.type || "Backboard",
      qty: String(row.size == null ? "" : row.size),
      cost: 0,
      month,
      rateMissing: true,
      worker: row.worker,
      process: row.process
    });
  });

  readWoodUsage().forEach((row) => {
    const ts = row.timestamp;
    const month = ts ? monthKey(ts) : "unknown";
    const qty = [row.height, row.width].filter((v) => v !== "" && v != null).join(" × ")
      + (row.quantity ? " × " + row.quantity : "")
      + (row.thickness ? " · " + row.thickness : "");
    addMaterial(order(row.orderNum), {
      type: "wood",
      item: [row.type, row.component].filter(Boolean).join(" ") || "Wood",
      qty,
      cost: 0,
      month,
      rateMissing: true,
      worker: row.worker,
      status: row.status
    });
  });

  Object.keys(meta).forEach((id) => { order(id); });

  const startMonth = range.start ? monthKey(range.start) : "";

  const orders = Object.keys(byId).map((id) => {
    const o = byId[id];
    const totalOverhead = (o.sellingPrice || 0) * OVERHEAD_PCT;
    const absHours = o.totalHours > 0 ? o.totalHours : 1;
    const monthlyBreakdown = {};
    Object.keys(o.monthHours).forEach((m) => {
      monthlyBreakdown[m] = (o.monthHours[m] / absHours) * totalOverhead;
    });

    let fHours = 0;
    let fLabor = 0;
    let fMat = 0;
    let displayOH = 0;
    const prunedMonthHours = {};
    const prunedMonthCosts = {};
    const prunedMonthMaterials = {};
    const prunedMats = [];

    if (range.kind === "all") {
      fHours = o.totalHours;
      fLabor = o.laborCost;
      fMat = o.materialCost;
      displayOH = totalOverhead;
      Object.assign(prunedMonthHours, o.monthHours);
      Object.assign(prunedMonthCosts, o.monthCosts);
      Object.assign(prunedMonthMaterials, o.monthMaterials);
      prunedMats.push.apply(prunedMats, o.materialEntries);
    } else {
      Object.keys(o.monthHours).forEach((m) => {
        const [y, mo] = m.split("-").map(Number);
        const mStart = sastDate(y, mo, 1, 0, 0, 0);
        const last = new Date(Date.UTC(y, mo, 0)).getUTCDate();
        const mEnd = sastDate(y, mo, last, 23, 59, 59);
        if (mStart <= range.end && mEnd >= range.start) {
          prunedMonthHours[m] = o.monthHours[m];
          prunedMonthCosts[m] = o.monthCosts[m];
          fHours += o.monthHours[m];
          fLabor += o.monthCosts[m];
          displayOH += monthlyBreakdown[m] || 0;
        }
      });
      Object.keys(o.monthMaterials).forEach((m) => {
        const [y, mo] = m.split("-").map(Number);
        const mStart = sastDate(y, mo, 1, 0, 0, 0);
        const last = new Date(Date.UTC(y, mo, 0)).getUTCDate();
        const mEnd = sastDate(y, mo, last, 23, 59, 59);
        if (mStart <= range.end && mEnd >= range.start) {
          prunedMonthMaterials[m] = o.monthMaterials[m];
        }
      });
      o.materialEntries.forEach((entry) => {
        if (entry.month === "unknown") return;
        const [y, mo] = String(entry.month).split("-").map(Number);
        if (!y || !mo) return;
        const mStart = sastDate(y, mo, 1, 0, 0, 0);
        const last = new Date(Date.UTC(y, mo, 0)).getUTCDate();
        const mEnd = sastDate(y, mo, last, 23, 59, 59);
        if (mStart <= range.end && mEnd >= range.start) {
          prunedMats.push(entry);
          fMat += entry.cost;
        }
      });
    }

    let pastCosts = 0;
    if (range.kind !== "all" && startMonth) {
      Object.keys(o.monthCosts).forEach((m) => { if (m < startMonth) pastCosts += o.monthCosts[m]; });
      Object.keys(o.monthMaterials).forEach((m) => { if (m < startMonth) pastCosts += o.monthMaterials[m]; });
      Object.keys(monthlyBreakdown).forEach((m) => { if (m < startMonth) pastCosts += monthlyBreakdown[m]; });
    }

    const hasWork = Math.round(fLabor || 0) !== 0 || Math.round(fMat || 0) !== 0 || (fHours || 0) >= 0.05;
    const hasMaterials = prunedMats.length > 0;
    const priced = (o.sellingPrice || 0) > 0;
    const soldInRange = !range.start || inRange(o.soldAt, range);
    if (!hasWork && !hasMaterials && !(priced && (range.kind === "all" || soldInRange))) return null;

    return {
      orderNum: o.orderNum,
      product: o.product,
      img: o.img,
      sellingPrice: o.sellingPrice,
      displaySale: Math.max(0, (o.sellingPrice || 0) - pastCosts),
      pastCosts,
      totalHours: fHours,
      laborCost: fLabor,
      materialCost: fMat,
      glassCost: prunedMats.filter((e) => e.type === "glass").reduce((s, e) => s + (e.cost || 0), 0),
      powderCost: prunedMats.filter((e) => e.type === "powder").reduce((s, e) => s + (e.cost || 0), 0),
      steelCost: prunedMats.filter((e) => e.type === "steel").reduce((s, e) => s + (e.cost || 0), 0),
      woodCost: prunedMats.filter((e) => e.type === "wood").reduce((s, e) => s + (e.cost || 0), 0),
      backboardCost: prunedMats.filter((e) => e.type === "backboard").reduce((s, e) => s + (e.cost || 0), 0),
      overheadCost: displayOH,
      materialEntries: prunedMats,
      tasks: o.tasks,
      staff: o.staff,
      warnings: o.warnings,
      monthHours: prunedMonthHours,
      monthCosts: prunedMonthCosts,
      monthMaterials: prunedMonthMaterials,
      monthlyBreakdown
    };
  }).filter(Boolean);

  const staffHours = {};
  orders.forEach((o) => {
    Object.keys(o.staff || {}).forEach((name) => {
      staffHours[name] = (staffHours[name] || 0) + (o.staff[name].h || 0);
    });
  });
  const topName = Object.keys(staffHours).sort((a, b) => staffHours[b] - staffHours[a])[0] || "";

  return {
    orders,
    tasks: TASKS,
    topPerformer: topName,
    warningCount: orders.filter((o) => o.warnings && o.warnings.length).length,
    labourRates: snapshotLabourRates(),
    steelRates: steelRates.snapshotRates(),
    missingEmployeeRates: Object.keys(missingEmployeeRates)
  };
}

module.exports = {
  TASKS,
  OVERHEAD_PCT,
  DEFAULT_EMPLOYEE_RATES,
  parseFilterRange,
  matchTask,
  snapshotLabourRates,
  upsertLabourRate,
  deleteLabourRate,
  clearSteelUsage,
  clearProductionLogs,
  clearFinishedProductionLogs,
  getAppData
};
