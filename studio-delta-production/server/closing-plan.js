"use strict";

const fs = require("fs");
const path = require("path");
const { getBook, dataDir } = require("./workbook-store");
const { listOrders, formatOrderId, listLiveDeliveries } = require("./db");
const staff = require("./staff");
const dates = require("./sast-date");
const { remainingPlanForStatus, isProductionFinished, isParked, PLANNED_PROCESSES } = require("./shop-status");

const SHIFT_MIN = 7.5 * 60;
const PRODUCTIVE_MIN = 6 * 60; // fallback when a person has no daily clocks yet (80% of 7.5)
const PAINT_WAIT_DAYS = 5;
const DEFAULT_PROCESS_MIN = 90;
const MIN_SAMPLE = 5;
const MAX_SAMPLE = 12 * 60;
const MIN_DAY_MIN = 30;

function settingsPath() {
  return path.join(dataDir(), "closing-plan.json");
}

function loadSettings() {
  try {
    const parsed = JSON.parse(fs.readFileSync(settingsPath(), "utf8"));
    return normalizeSettings(parsed);
  } catch (e) {
    return normalizeSettings({});
  }
}

function normalizeSettings(raw) {
  const closingDate = String((raw && raw.closingDate) || "").slice(0, 10);
  let weekdayOtHours = Number(raw && raw.weekdayOtHours);
  if (!(weekdayOtHours >= 0) || weekdayOtHours > 6) weekdayOtHours = 2;
  const saturday = raw && raw.saturday === false ? false : true;
  return { closingDate, weekdayOtHours, saturday };
}

function saveSettings(body) {
  const next = normalizeSettings(Object.assign({}, loadSettings(), body || {}));
  fs.mkdirSync(path.dirname(settingsPath()), { recursive: true });
  const tmp = settingsPath() + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(next, null, 2));
  fs.renameSync(tmp, settingsPath());
  return next;
}

function isoDay(d) {
  if (typeof d === "string" && /^\d{4}-\d{2}-\d{2}$/.test(d)) return d;
  if (d instanceof Date && !isNaN(d.getTime())) return dates.isoFromDate(d);
  return dates.isoFromDate(new Date());
}

function addDays(iso, n) {
  const [y, m, d] = iso.split("-").map(Number);
  const utc = Date.UTC(y, m - 1, d + n);
  return dates.isoFromDate(new Date(utc - dates.SAST_OFFSET_MS));
}

function weekday(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(Date.UTC(y, m - 1, d)).getUTCDay(); // 0 Sun
}

function formatHours(mins) {
  const n = Math.max(0, Number(mins) || 0);
  const h = Math.floor(n / 60);
  const m = Math.round(n % 60);
  if (!h) return m + " min";
  if (!m) return h + " h";
  return h + " h " + m + " min";
}

function median(nums) {
  const a = (nums || []).filter((n) => n > 0).slice().sort((x, y) => x - y);
  if (!a.length) return 0;
  const mid = Math.floor(a.length / 2);
  return a.length % 2 ? a[mid] : Math.round((a[mid - 1] + a[mid]) / 2);
}

/**
 * Typical output per calendar day from clocks, not a month total.
 * A month average hides slow days and one long Saturday.
 */
function buildDailyRates(clocks) {
  const byWorkerDay = {};
  const names = {};
  (clocks || []).forEach((row) => {
    const worker = String(row.worker || "").trim();
    const day = String(row.day || "").slice(0, 10);
    if (!worker || !day) return;
    names[worker.toLowerCase()] = worker;
    const key = worker.toLowerCase() + "||" + day;
    byWorkerDay[key] = (byWorkerDay[key] || 0) + (Number(row.minutes) || 0);
  });
  const buckets = {};
  Object.keys(byWorkerDay).forEach((key) => {
    const parts = key.split("||");
    const wk = parts[0];
    const day = parts[1];
    const mins = Math.min(MAX_SAMPLE, byWorkerDay[key]);
    if (mins < MIN_DAY_MIN) return;
    buckets[wk] = buckets[wk] || { name: names[wk] || wk, days: [], weekday: { 1: [], 2: [], 3: [], 4: [], 5: [], 6: [] } };
    buckets[wk].days.push(mins);
    const wd = weekday(day);
    if (buckets[wk].weekday[wd]) buckets[wk].weekday[wd].push(mins);
  });
  const out = {};
  Object.keys(buckets).forEach((wk) => {
    const row = buckets[wk];
    const typicalMin = Math.min(SHIFT_MIN, median(row.days));
    const byWeekday = {};
    Object.keys(row.weekday).forEach((wd) => {
      const list = row.weekday[wd];
      if (list.length >= 2) byWeekday[wd] = Math.min(SHIFT_MIN, median(list));
    });
    out[wk] = {
      worker: row.name,
      typicalMin,
      typicalHours: Math.round((typicalMin / 60) * 10) / 10,
      dayCount: row.days.length,
      pct: Math.round((typicalMin / SHIFT_MIN) * 1000) / 10,
      byWeekday
    };
  });
  return out;
}

function dayCapacity(rates, worker, day, extraMin) {
  const row = rates && worker ? rates[String(worker).trim().toLowerCase()] : null;
  const wd = weekday(day);
  let base = PRODUCTIVE_MIN;
  if (row) {
    if (row.byWeekday && row.byWeekday[wd] != null) base = row.byWeekday[wd];
    else base = row.typicalMin;
  }
  const paidRate = Math.min(1.15, Math.max(0.35, base / SHIFT_MIN));
  return Math.max(60, Math.round(base + (Number(extraMin) || 0) * paidRate));
}

function parseMeta(raw) {
  try {
    const j = JSON.parse(String(raw || "{}"));
    return j && typeof j === "object" ? j : {};
  } catch (e) {
    return {};
  }
}

function netMinutes(start, end, row) {
  const s = start ? new Date(start).getTime() : 0;
  const e = end ? new Date(end).getTime() : 0;
  if (!(s > 0) || !(e > s)) return 0;
  let pause = 0;
  const meta = parseMeta(row[12]);
  const pauses = Array.isArray(meta.pauses) ? meta.pauses : [];
  if (pauses.length) {
    pauses.forEach((p) => {
      const a = p && p.start ? new Date(p.start).getTime() : 0;
      const b = p && p.end ? new Date(p.end).getTime() : e;
      if (a && b > a) pause += (Math.min(b, e) - Math.max(a, s)) / 60000;
    });
  } else {
    pause = Number(row[10]) || 0;
  }
  return Math.round(Math.max(0, (e - s) / 60000 - pause));
}

function normProcess(name) {
  const s = String(name || "").trim();
  if (/^powder/i.test(s)) return "Powder coating";
  if (/^pre-?powder/i.test(s)) return "Pre-Powder Coating";
  if (/^final qc/i.test(s) || /^quality control$/i.test(s)) return "Final QC";
  return s;
}

function skipLogProcess(process) {
  const p = String(process || "").trim();
  if (!p) return true;
  if (p === "Indirect") return true;
  if (/^rework/i.test(p)) return true;
  if (/meeting|lunch|cleaning/i.test(p)) return true;
  return false;
}

function loadFinishedClocks() {
  const sheet = getBook().getSheetByName("Production_Log");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lastCol = Math.max(sheet.getLastColumn(), 13);
  const grid = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
  const orders = {};
  listOrders().forEach((o) => {
    orders[formatOrderId(o.order_number)] = String(o.product || "").trim();
  });
  const out = [];
  for (let i = 1; i < grid.length; i++) {
    const row = grid[i];
    const orderNumber = formatOrderId(row[1]);
    const worker = String(row[2] || "").trim();
    const process = normProcess(row[4] || row[3]);
    if (!orderNumber || !worker || skipLogProcess(process)) continue;
    if (String(row[3] || "").trim() === "Indirect") continue;
    const meta = parseMeta(row[12]);
    if (String(meta.entryType || "") === "indirect") continue;
    if (!row[6]) continue;
    const minutes = netMinutes(row[5], row[6], row);
    if (minutes < MIN_SAMPLE || minutes > MAX_SAMPLE) continue;
    out.push({
      orderNumber,
      worker,
      process,
      product: orders[orderNumber] || "",
      minutes,
      day: isoDay(new Date(row[6]))
    });
  }
  return out;
}

function productMinutesMap(product, durationRows) {
  const out = {};
  (durationRows || []).forEach((row) => {
    if (String(row.product || "").trim() !== product) return;
    out[normProcess(row.process)] = Number(row.minutes) || 0;
  });
  return out;
}

function userCanDoProcess(user, process) {
  if (!user) return false;
  const tasks = user.tasks || [];
  if (process === "Pre-Powder Coating" || process === "Final QC") {
    return tasks.indexOf("Quality Control") !== -1 || tasks.indexOf(process) !== -1;
  }
  return tasks.indexOf(process) !== -1;
}

function peopleForProcess(users, process) {
  return (users || []).filter((u) => {
    const access = String(u.access || "").toLowerCase();
    if (access === "marketing") return false;
    return userCanDoProcess(u, process);
  });
}

function buildLearner(clocks, durationRows) {
  const byPP = {};
  const byWPP = {};
  const byProc = {};
  clocks.forEach((row) => {
    const pp = row.product + "||" + row.process;
    const wpp = row.worker.toLowerCase() + "||" + pp;
    (byPP[pp] = byPP[pp] || []).push(row.minutes);
    (byWPP[wpp] = byWPP[wpp] || []).push(row.minutes);
    (byProc[row.process] = byProc[row.process] || []).push(row.minutes);
  });
  const sheet = {};
  (durationRows || []).forEach((row) => {
    sheet[String(row.product || "").trim() + "||" + normProcess(row.process)] = Number(row.minutes) || 0;
  });

  function estimate(product, process, worker) {
    const pp = product + "||" + process;
    const person = worker
      ? (byWPP[String(worker).toLowerCase() + "||" + pp] || [])
      : [];
    const group = byPP[pp] || [];
    const sheetMins = Number(sheet[pp]) || 0;
    const proc = byProc[process] || [];
    let minutes = 0;
    let source = "default";
    let samples = 0;
    if (person.length >= 2) {
      minutes = median(person);
      source = "this person on this product";
      samples = person.length;
    } else if (group.length >= 2) {
      minutes = person.length === 1
        ? Math.round(person[0] * 0.6 + median(group) * 0.4)
        : median(group);
      source = person.length === 1 ? "this person (1 job) mixed with the product" : "this product (any person)";
      samples = group.length;
    } else if (person.length === 1) {
      minutes = person[0];
      source = "this person (1 job)";
      samples = 1;
    } else if (group.length === 1) {
      minutes = group[0];
      source = "this product (1 job)";
      samples = 1;
    } else if (sheetMins > 0) {
      minutes = sheetMins;
      source = "Task times";
      samples = 0;
    } else if (proc.length) {
      minutes = median(proc);
      source = "this station, any product";
      samples = proc.length;
    } else {
      minutes = DEFAULT_PROCESS_MIN;
      source = "default 90 min";
      samples = 0;
    }
    return { minutes: Math.max(15, minutes), source, samples };
  }

  const learned = Object.keys(byPP).map((key) => {
    const parts = key.split("||");
    return {
      product: parts[0],
      process: parts[1],
      medianMinutes: median(byPP[key]),
      jobs: byPP[key].length
    };
  }).sort((a, b) => a.product.localeCompare(b.product) || a.process.localeCompare(b.process));

  const peopleSkill = {};
  clocks.forEach((row) => {
    if (!row.product) return;
    const key = row.worker + "||" + row.process + "||" + row.product;
    peopleSkill[key] = peopleSkill[key] || { worker: row.worker, process: row.process, product: row.product, minutes: [], jobs: 0 };
    peopleSkill[key].minutes.push(row.minutes);
    peopleSkill[key].jobs += 1;
  });
  const skills = Object.keys(peopleSkill).map((key) => {
    const row = peopleSkill[key];
    return {
      worker: row.worker,
      process: row.process,
      product: row.product,
      medianMinutes: median(row.minutes),
      jobs: row.jobs
    };
  }).sort((a, b) => a.worker.localeCompare(b.worker) || a.process.localeCompare(b.process));

  return { estimate, learned, skills, clockCount: clocks.length };
}

function workdays(fromIso, toIso, saturday) {
  const out = [];
  if (!fromIso || !toIso || toIso < fromIso) return out;
  let day = fromIso;
  for (let i = 0; i < 800; i++) {
    const wd = weekday(day);
    if (wd !== 0 && (wd !== 6 || saturday)) out.push(day);
    if (day === toIso) break;
    day = addDays(day, 1);
  }
  return out;
}

function pickPerson(process, product, people, learner, load) {
  if (!people.length) return { name: "", estimate: learner.estimate(product, process, "") };
  let best = null;
  people.forEach((user) => {
    const est = learner.estimate(product, process, user.name);
    const used = load[user.name] || 0;
    const score = est.minutes + used * 0.05;
    if (!best || score < best.score) {
      best = { name: user.name, estimate: est, score };
    }
  });
  return best;
}

function placeMinutes(left, days, used, extra, capForDay) {
  extra = extra || {};
  let startDay = "";
  let endDay = "";
  const remain = Object.assign({}, left);
  for (let i = 0; i < days.length && remain.minutes > 0; i++) {
    const day = days[i];
    const extraMin = Number(extra[day]) || 0;
    const cap = capForDay ? capForDay(day, extraMin) : (PRODUCTIVE_MIN + extraMin);
    const already = used[day] || 0;
    const free = cap - already;
    if (free <= 0) continue;
    const take = Math.min(free, remain.minutes);
    used[day] = already + take;
    remain.minutes -= take;
    if (!startDay) startDay = day;
    endDay = day;
  }
  return { startDay, endDay, leftover: remain.minutes, used };
}

function simulate(opts) {
  const today = isoDay(opts.now);
  const closingDate = opts.closingDate;
  const days = workdays(today, closingDate, !!opts.saturday);
  const extra = opts.extra || {}; // { worker: { day: extraMin } }
  const used = {};
  (opts.peopleNames || []).forEach((name) => { used[name] = {}; });
  const assignments = [];
  let lastFinish = "";
  let leftoverMin = 0;

  (opts.jobs || []).forEach((job) => {
    let ready = today;
    job.steps.forEach((step) => {
      if (step.waitDays) {
        ready = addDays(ready, step.waitDays);
        return;
      }
      const name = step.worker;
      if (!used[name]) used[name] = {};
      const from = days.filter((d) => d >= ready);
      const rates = opts.rates || {};
      const placed = placeMinutes(
        { minutes: step.minutes },
        from,
        used[name],
        extra[name] || {},
        (day, extraMin) => dayCapacity(rates, name, day, extraMin)
      );
      leftoverMin += placed.leftover;
      if (placed.endDay) {
        ready = addDays(placed.endDay, 1);
        if (!lastFinish || placed.endDay > lastFinish) lastFinish = placed.endDay;
      }
      assignments.push({
        orderNumber: job.orderNumber,
        product: job.product,
        process: step.process,
        worker: name,
        minutes: step.minutes,
        hoursLabel: formatHours(step.minutes),
        source: step.source,
        samples: step.samples,
        startDay: placed.startDay,
        endDay: placed.endDay,
        overflow: placed.leftover > 0
      });
    });
  });

  return { assignments, lastFinish, leftoverMin, used };
}

function suggestOvertime(sim, settings, peopleNames, days, rates) {
  if (!sim.leftoverMin && sim.lastFinish && sim.lastFinish <= settings.closingDate) {
    return { rows: [], extra: {} };
  }
  const extra = {};
  const rows = [];
  const weekdayExtra = Math.round((Number(settings.weekdayOtHours) || 0) * 60);
  const load = peopleNames.map((name) => {
    let total = 0;
    Object.keys(sim.used[name] || {}).forEach((d) => { total += sim.used[name][d]; });
    return { name, total };
  }).sort((a, b) => b.total - a.total);

  let need = sim.leftoverMin || 0;
  if (!need && sim.lastFinish > settings.closingDate) {
    need = PRODUCTIVE_MIN * 2;
  }

  load.forEach((row) => {
    if (need <= 0) return;
    days.forEach((day) => {
      if (need <= 0) return;
      const wd = weekday(day);
      if (wd === 6 && settings.saturday) {
        const satMin = dayCapacity(rates, row.name, day, 0);
        need -= satMin;
        rows.push({
          worker: row.name,
          day,
          extraHours: Math.round((satMin / 60) * 10) / 10,
          reason: "Saturday at this person's usual daily output"
        });
        return;
      }
      if (wd < 1 || wd > 5 || !weekdayExtra) return;
      extra[row.name] = extra[row.name] || {};
      extra[row.name][day] = (extra[row.name][day] || 0) + weekdayExtra;
      const gained = dayCapacity(rates, row.name, day, weekdayExtra) - dayCapacity(rates, row.name, day, 0);
      need -= gained;
      rows.push({
        worker: row.name,
        day,
        extraHours: Math.round((weekdayExtra / 60) * 10) / 10,
        reason: "After 15:45 (" + settings.weekdayOtHours + " h paid; ~" +
          (Math.round((gained / 60) * 10) / 10) + " h of their usual output)"
      });
    });
  });
  return { rows, extra };
}

function prettyDay(iso) {
  const s = String(iso || "").slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return "";
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const parts = s.split("-");
  return Number(parts[2]) + " " + months[Number(parts[1]) - 1] + " " + parts[0];
}

function shopDaysBetween(fromIso, toIso, saturday) {
  if (!fromIso || !toIso || toIso <= fromIso) return 0;
  return workdays(addDays(fromIso, 1), toIso, !!saturday).length;
}

function finishesFromAssignments(assignments) {
  const out = {};
  (assignments || []).forEach((row) => {
    const id = row.orderNumber;
    if (!id) return;
    const cur = out[id] || { finish: "", leftoverMin: 0, lateProcess: "", worker: "" };
    if (row.endDay && (!cur.finish || row.endDay > cur.finish)) cur.finish = row.endDay;
    if (row.overflow) {
      cur.leftoverMin += Number(row.minutes) || 0;
      cur.lateProcess = row.process;
      cur.worker = row.worker;
    }
    out[id] = cur;
  });
  return out;
}

function peopleLoad(sim, rates, today, closingDate, saturday) {
  const days = workdays(today, closingDate, !!saturday);
  return Object.keys(sim.used || {}).map((name) => {
    let booked = 0;
    Object.keys(sim.used[name] || {}).forEach((d) => { booked += sim.used[name][d]; });
    let cap = 0;
    days.forEach((day) => { cap += dayCapacity(rates, name, day, 0); });
    return {
      worker: name,
      bookedHours: Math.round((booked / 60) * 10) / 10,
      capacityHours: Math.round((cap / 60) * 10) / 10,
      full: cap > 0 && booked >= cap * 0.95
    };
  }).sort((a, b) => b.bookedHours - a.bookedHours);
}

function explainOverrun(opts) {
  const settings = opts.settings;
  const closing = settings.closingDate;
  const leftoverH = Math.round((opts.leftoverMin / 60) * 10) / 10;
  const finish = opts.lastFinish || "";
  const lateDays = finish ? shopDaysBetween(closing, finish, settings.saturday) : 0;
  const missed = opts.missed || [];
  const bottlenecks = (opts.load || []).filter((p) => p.full).slice(0, 4);
  const drop = opts.drop || [];
  const facts = [];
  let headline = "";

  if (finish && finish > closing) {
    headline = "Last order finishes " + prettyDay(finish) + " — " + lateDays + " shop day" +
      (lateDays === 1 ? "" : "s") + " after " + prettyDay(closing) + ".";
  } else if (leftoverH > 0) {
    headline = "About " + leftoverH + " h of shop work still has no seat before " + prettyDay(closing) + ".";
  } else {
    headline = "The open orders do not fit before " + prettyDay(closing) + ".";
  }

  if (leftoverH > 0) {
    facts.push("After weekday overtime" + (settings.saturday ? " and Saturdays" : "") +
      ", about " + leftoverH + " h still cannot be booked.");
  }
  if (finish && finish > closing) {
    facts.push("If the same crew keeps going past closing, the last clock would fall on " + prettyDay(finish) + ".");
  }
  if (missed.length) {
    const names = missed.slice(0, 8).map((j) => j.orderNumber + (j.finish ? " (" + prettyDay(j.finish) + ")" : "")).join(", ");
    facts.push(missed.length + " order" + (missed.length === 1 ? "" : "s") + " miss the date: " + names +
      (missed.length > 8 ? "…" : "."));
  }
  if (bottlenecks.length) {
    facts.push("Full before closing: " + bottlenecks.map((p) => p.worker + " (" + p.bookedHours + " h booked / " +
      p.capacityHours + " h typical)").join("; ") + ".");
  }
  const missingCrew = (opts.unassigned || []).slice(0, 5);
  if (missingCrew.length) {
    facts.push("No one on Users is ticked for: " + missingCrew.join(", ") + ".");
  }
  if (drop.length) {
    facts.push("Dropping " + drop.map((j) => j.orderNumber).join(", ") + " would free about " +
      drop.reduce((s, j) => s + j.hours, 0) + " h — enough to bring the rest inside the date.");
  } else if (missed.length > 1) {
    facts.push("Cutting the latest orders, adding a person on the full stations, or moving closing to " +
      prettyDay(finish || closing) + " are the three ways to close the gap.");
  } else if (finish) {
    facts.push("Move closing to " + prettyDay(finish) + ", or add capacity on the full stations.");
  }
  return { headline, facts };
}

function openShopOrders() {
  const due = {};
  try {
    listLiveDeliveries().forEach((d) => {
      due[formatOrderId(d.order_number)] = d;
    });
  } catch (e) {}
  return listOrders()
    .filter((o) => {
      const status = o.status || "";
      if (isProductionFinished(status) || isParked(status)) return false;
      return true;
    })
    .map((o) => {
      const id = formatOrderId(o.order_number);
      const delivery = due[id] || null;
      return {
        orderNumber: id,
        product: String(o.product || "").trim(),
        status: o.status || "",
        assigned: String(o.assigned_operator || "").trim(),
        deliveryDay: delivery && delivery.day || "",
        deliveryCode: delivery && delivery.code || ""
      };
    });
}

function build(opts) {
  opts = opts || {};
  const settings = normalizeSettings(Object.assign({}, loadSettings(), opts.settings || {}, {
    closingDate: opts.closingDate || (opts.settings && opts.settings.closingDate) || loadSettings().closingDate
  }));
  const today = isoDay(opts.now);
  const clocks = loadFinishedClocks();
  const durationRows = staff.listDurations();
  const learner = buildLearner(clocks, durationRows);
  const dailyRates = buildDailyRates(clocks);
  const users = staff.listUsers();
  const orders = openShopOrders();

  const jobs = orders.map((order) => {
    const remaining = remainingPlanForStatus(order.status, productMinutesMap(order.product, durationRows));
    const steps = [];
    remaining.processes.forEach((process) => {
      const people = peopleForProcess(users, process);
      const pick = pickPerson(process, order.product, people, learner, {});
      const worker = pick.name || (people[0] && people[0].name) || "Unassigned";
      const est = pick.estimate || learner.estimate(order.product, process, worker);
      steps.push({
        process,
        worker,
        minutes: est.minutes,
        source: est.source,
        samples: est.samples
      });
      if (process === "Pre-Powder Coating" && remaining.paintWait) {
        steps.push({ waitDays: PAINT_WAIT_DAYS, process: "Paint shop wait", worker: "", minutes: 0, source: "5 calendar days", samples: 0 });
      }
    });
    const remainingMinutes = steps.reduce((sum, s) => sum + (s.minutes || 0), 0);
    return {
      orderNumber: order.orderNumber,
      product: order.product,
      status: order.status,
      deliveryDay: order.deliveryDay,
      deliveryCode: order.deliveryCode,
      remainingMinutes,
      remainingLabel: formatHours(remainingMinutes),
      steps
    };
  }).sort((a, b) => {
    if (a.deliveryDay && b.deliveryDay && a.deliveryDay !== b.deliveryDay) return a.deliveryDay.localeCompare(b.deliveryDay);
    if (a.deliveryDay && !b.deliveryDay) return -1;
    if (!a.deliveryDay && b.deliveryDay) return 1;
    return a.orderNumber.localeCompare(b.orderNumber);
  });

  const peopleNames = [];
  jobs.forEach((job) => {
    job.steps.forEach((s) => {
      if (s.worker && peopleNames.indexOf(s.worker) === -1) peopleNames.push(s.worker);
    });
  });

  const paidDays = workdays(today, settings.closingDate, false);
  const allDays = workdays(today, settings.closingDate, true);
  const base = simulate({
    now: today,
    closingDate: settings.closingDate,
    saturday: false,
    jobs,
    peopleNames,
    extra: {},
    rates: dailyRates
  });
  const ot = suggestOvertime(base, settings, peopleNames, allDays, dailyRates);
  const withOt = simulate({
    now: today,
    closingDate: settings.closingDate,
    saturday: !!settings.saturday,
    jobs,
    peopleNames,
    extra: ot.extra,
    rates: dailyRates
  });

  const canFinishPaid = !!(settings.closingDate && !base.leftoverMin && base.lastFinish && base.lastFinish <= settings.closingDate);
  const canFinishWithOt = !!(settings.closingDate && !withOt.leftoverMin && withOt.lastFinish && withOt.lastFinish <= settings.closingDate);

  const farDate = addDays(today, 540);
  const onward = simulate({
    now: today,
    closingDate: farDate,
    saturday: !!settings.saturday,
    jobs,
    peopleNames,
    extra: ot.extra,
    rates: dailyRates
  });
  const finishMap = finishesFromAssignments(onward.assignments);
  const load = peopleLoad(withOt, dailyRates, today, settings.closingDate, false);
  const missed = jobs.map((job) => {
    const fin = finishMap[job.orderNumber] || {};
    const finish = fin.finish || "";
    const miss = !!(fin.leftoverMin || (settings.closingDate && finish && finish > settings.closingDate) || (settings.closingDate && !finish && job.remainingMinutes));
    return {
      orderNumber: job.orderNumber,
      product: job.product,
      remainingLabel: job.remainingLabel,
      remainingMinutes: job.remainingMinutes,
      hours: Math.round((job.remainingMinutes / 60) * 10) / 10,
      finish,
      miss
    };
  }).filter((j) => j.miss).sort((a, b) => (b.finish || "").localeCompare(a.finish || "") || b.remainingMinutes - a.remainingMinutes);

  let dropHours = 0;
  const drop = [];
  const rankedMiss = missed.slice().sort((a, b) => b.remainingMinutes - a.remainingMinutes);
  for (let i = 0; i < rankedMiss.length; i++) {
    if (withOt.leftoverMin > 0 && dropHours >= withOt.leftoverMin) break;
    if (withOt.leftoverMin <= 0 && drop.length >= Math.min(3, missed.length)) break;
    drop.push(rankedMiss[i]);
    dropHours += rankedMiss[i].remainingMinutes;
  }

  const unassigned = [];
  jobs.forEach((job) => {
    job.steps.forEach((s) => {
      if (s.worker === "Unassigned" && unassigned.indexOf(s.process) === -1) unassigned.push(s.process);
    });
  });

  let verdict = "Set the last shop day. The forecast uses finished clocks first, then Task times.";
  let verdictKind = "wait";
  let why = [];
  if (!settings.closingDate) {
    verdict = "Set the last shop day. The forecast uses finished clocks first, then Task times.";
    verdictKind = "wait";
  } else if (!orders.length) {
    verdict = "No open shop orders. Nothing left to manufacture before closing.";
    verdictKind = "ok";
  } else if (canFinishPaid) {
    verdict = "Paid weekdays look enough. Last predicted finish is " + prettyDay(base.lastFinish) + ".";
    why = ["No overtime needed if clocks keep matching history."];
    verdictKind = "ok";
  } else if (canFinishWithOt) {
    verdict = "Not on paid weekdays alone. With the overtime below, last predicted finish is " + prettyDay(withOt.lastFinish) + ".";
    why = [
      "Weekdays only would run to " + prettyDay(base.lastFinish || settings.closingDate) +
        (base.leftoverMin ? " with " + (Math.round((base.leftoverMin / 60) * 10) / 10) + " h still unbooked" : "") + ".",
      "The overtime list is what closes that gap."
    ];
    verdictKind = "ot";
  } else {
    const explained = explainOverrun({
      settings,
      leftoverMin: withOt.leftoverMin,
      lastFinish: onward.lastFinish || withOt.lastFinish,
      missed,
      load,
      drop: dropHours >= withOt.leftoverMin && withOt.leftoverMin > 0 ? drop : [],
      unassigned
    });
    verdict = explained.headline;
    why = explained.facts;
    verdictKind = "bad";
  }

  return {
    today,
    settings,
    clockCount: learner.clockCount,
    openOrderCount: orders.length,
    paidDayCount: paidDays.length,
    canFinishPaid,
    canFinishWithOt,
    lastFinishPaid: base.lastFinish || "",
    lastFinishWithOt: onward.lastFinish || withOt.lastFinish || "",
    leftoverPaidHours: Math.round((base.leftoverMin / 60) * 10) / 10,
    leftoverOtHours: Math.round((withOt.leftoverMin / 60) * 10) / 10,
    finishIfWeContinue: onward.lastFinish || "",
    shopDaysLate: onward.lastFinish && settings.closingDate
      ? shopDaysBetween(settings.closingDate, onward.lastFinish, settings.saturday)
      : 0,
    missedOrders: missed,
    bottlenecks: load.filter((p) => p.full),
    why,
    verdict,
    verdictKind,
    overtime: ot.rows,
    jobs: jobs.map((job) => {
      const fin = finishMap[job.orderNumber] || {};
      return {
        orderNumber: job.orderNumber,
        product: job.product,
        status: job.status,
        deliveryDay: job.deliveryDay,
        deliveryCode: job.deliveryCode,
        remainingLabel: job.remainingLabel,
        finishDay: fin.finish || "",
        miss: missed.some((m) => m.orderNumber === job.orderNumber),
        steps: job.steps.filter((s) => s.process && !s.waitDays).map((s) => ({
          process: s.process,
          worker: s.worker,
          hoursLabel: formatHours(s.minutes),
          source: s.source,
          samples: s.samples
        }))
      };
    }),
    learned: learner.learned,
    skills: learner.skills.slice(0, 80),
    daily: Object.keys(dailyRates).map((k) => dailyRates[k]).sort((a, b) => a.worker.localeCompare(b.worker)),
    note: "Each person's day is their median finished-clock minutes on days they actually worked — not a month-pack average. Weekdays with enough history use that weekday's median. No daily clocks yet → 6 h (80% of 7.5). Task times fill in when a product has too few finished jobs. Paint shop is a 5-calendar-day wait."
  };
}

function snapshot(opts) {
  return build(opts || {});
}

function update(body) {
  const settings = saveSettings(body || {});
  return build({ settings });
}

module.exports = {
  PRODUCTIVE_MIN,
  loadSettings,
  saveSettings,
  snapshot,
  update,
  build,
  median,
  netMinutes,
  formatHours,
  buildDailyRates,
  dayCapacity,
  SHIFT_MIN
};
