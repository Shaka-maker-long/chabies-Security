"use strict";

const db = require("./db");
const shop = require("./shop-status");
const sched = require("./office-schedule");
const noPlates = require("./no-plates");

const SAST_OFFSET_MS = 2 * 60 * 60 * 1000;

const AGE_BUCKETS = [
  { id: "0-3d", maxDays: 3 },
  { id: "4-7d", maxDays: 7 },
  { id: "8-14d", maxDays: 14 },
  { id: "15-30d", maxDays: 30 },
  { id: "31-60d", maxDays: 60 },
  { id: "60d+", maxDays: Infinity }
];

const PIPELINE = [
  { id: "drawing", label: "Waiting for drawing" },
  { id: "not_started", label: "Not yet started" },
  { id: "steelwork", label: "Steelwork" },
  { id: "paint", label: "Paint shop" },
  { id: "assembly", label: "Assembly" },
  { id: "qc", label: "QC" },
  { id: "delivery", label: "Delivery" },
  { id: "delivered", label: "Delivered" },
  { id: "other", label: "Other" }
];

const STEELWORK = [
  "Profile Cutting", "Ready for Tagging", "Tagging",
  "Ready for Welding", "Welding", "Ready for Grinding", "Grinding"
];
const PAINT = [
  "Ready for Pre-Powder Coating", "Pre-Powder Coating",
  "Ready for Powder Coating", "Sent to Paint Shop", "Paint Shop", "Powder Coating"
];
const ASSEMBLY = [
  "Ready for Assembly", "Assembly", "Paint Preparation", "Ready for Painting", "Painting"
];

function sastParts(d) {
  const sast = new Date(d.getTime() + SAST_OFFSET_MS);
  return { y: sast.getUTCFullYear(), m: sast.getUTCMonth(), day: sast.getUTCDate() };
}

function pad(n) {
  return String(n).padStart(2, "0");
}

function monthKey(d) {
  const p = sastParts(d);
  return p.y + "-" + pad(p.m + 1);
}

function monthLabel(key) {
  const m = String(key || "").match(/^(\d{4})-(\d{2})$/);
  if (!m) return String(key || "").trim();
  const names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return names[Number(m[2]) - 1] + " " + m[1];
}

function sastUtcDate(d) {
  const p = sastParts(d);
  return new Date(Date.UTC(p.y, p.m, p.day));
}

function weekKey(d) {
  const date = sastUtcDate(d);
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const year = date.getUTCFullYear();
  const yearStart = new Date(Date.UTC(year, 0, 1));
  const week = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return year + "-W" + pad(week);
}

function weekLabel(key) {
  return String(key || "").replace("-W", " W");
}

function addUtcMonths(y, m, delta) {
  const d = new Date(Date.UTC(y, m + delta, 1));
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() };
}

function bucketKeys(grain, fromMs, toMs) {
  const keys = [];
  if (grain === "week") {
    let t = fromMs;
    const seen = new Set();
    while (t <= toMs + 86400000) {
      const key = weekKey(new Date(t));
      if (!seen.has(key)) {
        seen.add(key);
        keys.push(key);
      }
      t += 86400000;
    }
    return keys;
  }
  const start = sastParts(new Date(fromMs));
  const end = sastParts(new Date(toMs));
  let y = start.y;
  let m = start.m;
  while (y < end.y || (y === end.y && m <= end.m)) {
    keys.push(y + "-" + pad(m + 1));
    const next = addUtcMonths(y, m, 1);
    y = next.y;
    m = next.m;
  }
  return keys;
}

function monthBounds(ym) {
  const m = String(ym || "").match(/^(\d{4})-(\d{2})$/);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]) - 1;
  if (mo < 0 || mo > 11) return null;
  const from = Date.UTC(y, mo, 1) - SAST_OFFSET_MS;
  const to = Date.UTC(y, mo + 1, 1) - SAST_OFFSET_MS - 1;
  return { from, to };
}

function periodWindow(range, grain) {
  const now = Date.now();
  const today = sastParts(new Date(now));
  let from;
  if (range === "all") {
    from = now - 5 * 365 * 86400000;
  } else if (range === "year") {
    from = Date.UTC(today.y, 0, 1) - SAST_OFFSET_MS;
  } else if (range === "6m") {
    const start = addUtcMonths(today.y, today.m, -5);
    from = Date.UTC(start.y, start.m, 1) - SAST_OFFSET_MS;
  } else if (grain === "week") {
    from = now - 16 * 7 * 86400000;
  } else {
    const start = addUtcMonths(today.y, today.m, -11);
    from = Date.UTC(start.y, start.m, 1) - SAST_OFFSET_MS;
  }
  return { from, to: now, grain: grain === "week" ? "week" : "month" };
}

function resolveWindow(query) {
  const grain = String((query && query.grain) || "month").toLowerCase() === "week" ? "week" : "month";
  const month = String((query && query.month) || "").trim();
  const rangeRaw = String((query && query.range) || "year").toLowerCase();
  const range = rangeRaw === "all" || rangeRaw === "year" || rangeRaw === "6m" ? rangeRaw : "year";
  const single = monthBounds(month);
  if (single) {
    return {
      from: single.from,
      to: single.to,
      grain,
      range,
      month,
      windowLabel: monthLabel(month)
    };
  }
  const win = periodWindow(range, grain);
  return {
    from: win.from,
    to: win.to,
    grain,
    range,
    month: "",
    windowLabel: range === "all" ? "All time" : range === "year" ? "This year" : "Last 6 months"
  };
}

function saleDate(row) {
  return db.asDate(row && row.payment_date);
}

function saleMs(row) {
  const d = saleDate(row);
  return d ? d.getTime() : 0;
}

function inWindow(ms, from, to) {
  return ms >= from && ms <= to;
}

function bucketOf(ms, grain) {
  const d = new Date(ms);
  return grain === "week" ? weekKey(d) : monthKey(d);
}

function blankLabel(v, empty) {
  const s = String(v || "").trim();
  return s || empty;
}

function moneyOf(row) {
  const billedIncl = db.parseMoney(row && row.price_incl_vat);
  const storedExcl = db.parseMoney(row && row.price_excl_vat);
  const billed = storedExcl || db.parseMoney(db.exclFromIncl(billedIncl || ""));
  const paidIncl = db.parseMoney(row && row.amount_paid);
  const paid = billedIncl > 0
    ? roundMoney(paidIncl * billed / billedIncl)
    : db.parseMoney(db.exclFromIncl(paidIncl || ""));
  return {
    billed: roundMoney(billed),
    paid: roundMoney(paid),
    owing: Math.max(0, roundMoney(billed - paid))
  };
}

function statusOf(row) {
  return shop.normalizeShopStatus(row && row.status) || String((row && row.status) || "").trim();
}

function isDelivered(status) {
  return shop.normalizeShopStatus(status) === "Delivered";
}

function isOpen(status) {
  return !isDelivered(status);
}

function isReadyOrOut(status) {
  const s = shop.normalizeShopStatus(status);
  return s === "Ready for Delivery" || s === "Out for Delivery";
}

function pipelineId(status) {
  if (shop.isWaitingForDrawing(status)) return "drawing";
  const s = shop.normalizeShopStatus(status);
  if (s === "Not Yet Started" || s === "Ready for Steelwork") return "not_started";
  if (STEELWORK.indexOf(s) >= 0) return "steelwork";
  if (PAINT.indexOf(s) >= 0 || shop.isAtPaintShop(status)) return "paint";
  if (ASSEMBLY.indexOf(s) >= 0) return "assembly";
  if (s === "Ready for Final QC" || s === "Final QC") return "qc";
  if (s === "Ready for Delivery" || s === "Out for Delivery") return "delivery";
  if (s === "Delivered") return "delivered";
  return "other";
}

function pipelineLabel(id) {
  const hit = PIPELINE.find((p) => p.id === id);
  return hit ? hit.label : "Other";
}

function ageBucket(days) {
  for (const b of AGE_BUCKETS) {
    if (days <= b.maxDays) return b.id;
  }
  return "60d+";
}

function median(nums) {
  const list = nums.filter((n) => Number.isFinite(n) && n >= 0).sort((a, b) => a - b);
  if (!list.length) return null;
  const mid = Math.floor(list.length / 2);
  return list.length % 2 ? list[mid] : (list[mid - 1] + list[mid]) / 2;
}

function round1(n) {
  return n == null ? null : Math.round(n * 10) / 10;
}

function roundMoney(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

function todayIso() {
  const p = sastParts(new Date());
  return p.y + "-" + pad(p.m + 1) + "-" + pad(p.day);
}

function addIsoDays(iso, days) {
  const d = new Date(String(iso).slice(0, 10) + "T12:00:00");
  d.setDate(d.getDate() + days);
  return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
}

function daysOpen(row, nowMs) {
  const ms = saleMs(row);
  if (!ms) return null;
  return Math.max(0, Math.floor((nowMs - ms) / 86400000));
}

function cardOf(row) {
  const money = moneyOf(row);
  const d = saleDate(row);
  return {
    order_number: row.order_number,
    client_name: row.client_name || "",
    status: row.status || "",
    product: row.product || "",
    category: row.category || "",
    source: row.source || "",
    payment_date: row.payment_date || "",
    month_of_sale: row.month_of_sale || "",
    sale_date: d ? db.formatPaymentDate(d) : "",
    price_excl_vat: money.billed,
    price_incl_vat: money.billed,
    amount_paid: money.paid,
    owing: money.owing
  };
}

function pickerMonths(rows) {
  const keys = new Set();
  const today = sastParts(new Date());
  for (let i = 0; i < 24; i++) {
    const p = addUtcMonths(today.y, today.m, -i);
    keys.add(p.y + "-" + pad(p.m + 1));
  }
  (rows || []).forEach((row) => {
    const d = saleDate(row);
    if (d) keys.add(monthKey(d));
  });
  return Array.from(keys).sort().reverse().map((key) => ({ key, label: monthLabel(key) }));
}

function sortedMoney(map, limit) {
  return Object.keys(map).map((label) => ({
    label,
    income: roundMoney(map[label].income),
    count: map[label].count || 0
  })).sort((a, b) => b.income - a.income || a.label.localeCompare(b.label))
    .slice(0, limit || 10);
}

function bumpMoney(map, label, income) {
  const key = label;
  if (!map[key]) map[key] = { income: 0, count: 0 };
  map[key].income += income;
  map[key].count += 1;
}

function deliveryWeekSummary(items, weekKeyValue, today) {
  const days = { Monday: 0, Tuesday: 0, Wednesday: 0, Thursday: 0, Friday: 0 };
  const rows = [];
  let late = 0;
  (items || []).forEach((it) => {
    if (it.weekKey !== weekKeyValue) return;
    const weekday = it.weekday || sched.weekdayLong(it.day);
    if (days[weekday] != null) days[weekday] += 1;
    const done = isDelivered(it.status) || shop.normalizeShopStatus(it.status) === "Out for Delivery";
    const isLate = it.day < today && !done;
    if (isLate) late += 1;
    rows.push({
      order_number: it.order_number,
      product: it.product || "",
      status: it.status || "",
      day: it.day,
      weekday,
      code: it.code,
      codeLabel: it.codeLabel,
      late: isLate
    });
  });
  return {
    weekKey: weekKeyValue,
    weekLabel: "Week " + String(weekKeyValue).replace(/^\d+-W/, ""),
    days,
    count: rows.length,
    late,
    items: rows
  };
}

function buildDashboard(query) {
  const win = resolveWindow(query || {});
  const rows = db.listOrders();
  const nowMs = Date.now();
  const today = todayIso();
  const thisMonday = sched.mondayOf(today);
  const nextMonday = addIsoDays(thisMonday, 7);
  const thisWeekKey = sched.weekKey(sched.isoWeekInfo(thisMonday));
  const nextWeekKey = sched.weekKey(sched.isoWeekInfo(nextMonday));

  const keys = bucketKeys(win.grain, win.from, win.to);
  const seriesMap = {};
  keys.forEach((key) => { seriesMap[key] = { key, income: 0, billed: 0, count: 0, delivered: 0 }; });

  const catMap = {};
  const sourceMap = {};
  const productIncome = {};
  const productQty = {};
  const pipeCounts = {};
  const pipeAges = {};
  PIPELINE.forEach((p) => {
    pipeCounts[p.id] = 0;
    pipeAges[p.id] = [];
  });
  const ageCounts = {};
  const ageStatuses = {};
  AGE_BUCKETS.forEach((b) => {
    ageCounts[b.id] = 0;
    ageStatuses[b.id] = {};
  });

  let income = 0;
  let billed = 0;
  let windowCount = 0;
  let openJobs = 0;
  let waitingForDrawing = 0;
  let atPaintShop = 0;
  let readyOrOut = 0;
  let unpaidOpen = 0;
  const stuck = [];

  rows.forEach((row) => {
    const money = moneyOf(row);
    const status = row.status;
    const pipe = pipelineId(status);
    pipeCounts[pipe] = (pipeCounts[pipe] || 0) + 1;
    const open = isOpen(status);
    if (open) {
      openJobs += 1;
      unpaidOpen += money.owing;
      if (shop.isWaitingForDrawing(status)) waitingForDrawing += 1;
      if (shop.isAtPaintShop(status)) atPaintShop += 1;
      if (isReadyOrOut(status)) readyOrOut += 1;
      const days = daysOpen(row, nowMs);
      if (days != null) {
        const bucket = ageBucket(days);
        ageCounts[bucket] += 1;
        const statusLabel = blankLabel(row.status, "(Blank)");
        ageStatuses[bucket][statusLabel] = (ageStatuses[bucket][statusLabel] || 0) + 1;
        pipeAges[pipe].push(days);
        stuck.push({
          order_number: row.order_number,
          product: row.product || "",
          status: row.status || "",
          pipeline: pipelineLabel(pipe),
          days,
          client_name: row.client_name || ""
        });
      }
    }

    const ms = saleMs(row);
    if (!ms || !inWindow(ms, win.from, win.to)) return;
    windowCount += 1;
    income += money.paid;
    billed += money.billed;
    const key = bucketOf(ms, win.grain);
    if (!seriesMap[key]) seriesMap[key] = { key, income: 0, billed: 0, count: 0, delivered: 0 };
    seriesMap[key].income += money.paid;
    seriesMap[key].billed += money.billed;
    seriesMap[key].count += 1;
    if (isDelivered(status)) seriesMap[key].delivered += 1;
    bumpMoney(catMap, blankLabel(row.category, "(Blank)"), money.paid);
    bumpMoney(sourceMap, blankLabel(row.source, "(Blank)"), money.paid);
    bumpMoney(productIncome, blankLabel(row.product, "(No product)"), money.paid);
    productQty[blankLabel(row.product, "(No product)")] = (productQty[blankLabel(row.product, "(No product)")] || 0) + 1;
  });

  const series = keys.map((key) => ({
    key,
    label: win.grain === "week" ? weekLabel(key) : monthLabel(key),
    income: roundMoney(seriesMap[key] ? seriesMap[key].income : 0),
    billed: roundMoney(seriesMap[key] ? seriesMap[key].billed : 0),
    count: seriesMap[key] ? seriesMap[key].count : 0,
    delivered: seriesMap[key] ? seriesMap[key].delivered : 0
  }));

  stuck.sort((a, b) => b.days - a.days || String(a.order_number).localeCompare(String(b.order_number)));

  const noPlateIds = new Set(noPlates.listNoPlates());
  let noPlateOpen = 0;
  rows.forEach((row) => {
    if (isOpen(row.status) && noPlateIds.has(db.formatOrderId(row.order_number))) noPlateOpen += 1;
  });

  let deliveryItems = [];
  try {
    deliveryItems = db.listDeliveryItems().items || [];
  } catch (e) {
    deliveryItems = [];
  }
  const thisWeek = deliveryWeekSummary(deliveryItems, thisWeekKey, today);
  const nextWeek = deliveryWeekSummary(deliveryItems, nextWeekKey, today);
  const lateItems = deliveryItems.filter((it) => {
    const done = isDelivered(it.status) || shop.normalizeShopStatus(it.status) === "Out for Delivery";
    return it.day < today && !done;
  });

  const blockers = [
    { id: "drawing", label: "Waiting for drawing", count: waitingForDrawing },
    { id: "no_plates", label: "No plates", count: noPlateOpen },
    { id: "paint", label: "At paint shop", count: atPaintShop },
    { id: "ready_delivery", label: "Ready for Delivery", count: rows.filter((r) => shop.normalizeShopStatus(r.status) === "Ready for Delivery").length }
  ];

  return {
    orderCount: rows.length,
    windowCount,
    windowLabel: win.windowLabel,
    grain: win.grain,
    range: win.range,
    month: win.month,
    months: pickerMonths(rows),
    kpis: {
      openJobs,
      waitingForDrawing,
      atPaintShop,
      readyOrOut,
      unpaidOpen: roundMoney(unpaidOpen),
      income: roundMoney(income),
      billed: roundMoney(billed),
      windowCount
    },
    series,
    categories: sortedMoney(catMap, 16),
    sources: sortedMoney(sourceMap, 16),
    topIncome: sortedMoney(productIncome, 10),
    topQuantity: Object.keys(productQty).map((label) => ({
      label,
      count: productQty[label]
    })).sort((a, b) => b.count - a.count || a.label.localeCompare(b.label)).slice(0, 10),
    pipeline: PIPELINE.map((p) => ({
      id: p.id,
      label: p.label,
      count: pipeCounts[p.id] || 0,
      medianDays: round1(median(pipeAges[p.id] || []))
    })),
    ageing: AGE_BUCKETS.map((b) => {
      const statuses = Object.keys(ageStatuses[b.id] || {}).map((label) => ({
        label,
        count: ageStatuses[b.id][label]
      })).sort((a, c) => c.count - a.count || a.label.localeCompare(c.label));
      return { id: b.id, label: b.id, count: ageCounts[b.id] || 0, statuses };
    }),
    stuck: stuck.slice(0, 10),
    delivery: {
      today,
      thisWeek,
      nextWeek,
      late: lateItems.length
    },
    blockers
  };
}

function matchesDrill(row, query, win) {
  const kind = String((query && query.kind) || "");
  const value = String((query && query.value) || "");
  const key = String((query && query.key) || "");
  const group = String((query && query.group) || value);
  const ms = saleMs(row);
  const inWin = ms && inWindow(ms, win.from, win.to);
  const status = row.status;

  if (kind === "income" || kind === "billed" || kind === "period" || kind === "throughput") {
    if (!inWin) return false;
    if (key && bucketOf(ms, win.grain) !== key) return false;
    if (kind === "throughput" && String((query && query.slice) || "") === "delivered") return isDelivered(status);
    return true;
  }
  if (kind === "category") return inWin && blankLabel(row.category, "(Blank)") === value;
  if (kind === "source") return inWin && blankLabel(row.source, "(Blank)") === value;
  if (kind === "product" || kind === "productIncome") {
    return inWin && blankLabel(row.product, "(No product)") === value;
  }
  if (kind === "productQty") return inWin && blankLabel(row.product, "(No product)") === value;
  if (kind === "pipeline") return pipelineId(status) === group;
  if (kind === "age") {
    if (!isOpen(status)) return false;
    const days = daysOpen(row, Date.now());
    if (days == null || ageBucket(days) !== value) return false;
    const slice = String((query && query.slice) || "").trim();
    if (slice) return blankLabel(row.status, "(Blank)") === slice;
    return true;
  }
  if (kind === "stuck") {
    return isOpen(status) && (!value || String(row.order_number) === value);
  }
  if (kind === "open") return isOpen(status);
  if (kind === "drawing") return shop.isWaitingForDrawing(status);
  if (kind === "paint") return shop.isAtPaintShop(status);
  if (kind === "ready_delivery") return shop.normalizeShopStatus(status) === "Ready for Delivery";
  if (kind === "readyOrOut") return isReadyOrOut(status);
  if (kind === "unpaid") return isOpen(status) && moneyOf(row).owing > 0;
  if (kind === "no_plates") {
    return isOpen(status) && noPlates.isNoPlate(row.order_number);
  }
  if (kind === "delivery" || kind === "late") {
    return false;
  }
  return false;
}

function deliveryDrill(query) {
  const today = todayIso();
  let items = [];
  try {
    items = db.listDeliveryItems().items || [];
  } catch (e) {
    items = [];
  }
  const kind = String((query && query.kind) || "");
  const week = String((query && query.week) || "");
  const weekday = String((query && query.weekday) || "");
  const filtered = items.filter((it) => {
    if (kind === "late") {
      const done = isDelivered(it.status) || shop.normalizeShopStatus(it.status) === "Out for Delivery";
      return it.day < today && !done;
    }
    if (week && it.weekKey !== week) return false;
    if (weekday && (it.weekday || sched.weekdayLong(it.day)) !== weekday) return false;
    return true;
  });
  const orders = db.listOrders();
  const byId = new Map(orders.map((o) => [db.formatOrderId(o.order_number), o]));
  const rows = filtered.map((it) => {
    const order = byId.get(db.formatOrderId(it.order_number)) || { order_number: it.order_number, product: it.product, status: it.status };
    const card = cardOf(order);
    card.delivery_day = it.day;
    card.delivery_code = it.code;
    card.late = it.day < today && !(isDelivered(it.status) || shop.normalizeShopStatus(it.status) === "Out for Delivery");
    return card;
  });
  const title = kind === "late"
    ? "Late deliveries"
    : (weekday ? weekday + " · " : "") + (week ? weekLabel(week) : "Deliveries");
  return {
    title,
    rows,
    totals: {
      count: rows.length,
      income: roundMoney(rows.reduce((s, r) => s + (r.amount_paid || 0), 0)),
      billed: roundMoney(rows.reduce((s, r) => s + (r.price_excl_vat || r.price_incl_vat || 0), 0)),
      owing: roundMoney(rows.reduce((s, r) => s + (r.owing || 0), 0))
    }
  };
}

function drillTitle(query, win) {
  const kind = String((query && query.kind) || "");
  const value = String((query && query.value) || "");
  const key = String((query && query.key) || "");
  const period = key ? (win.grain === "week" ? weekLabel(key) : monthLabel(key)) : win.windowLabel;
  if (kind === "income") return "Income · " + period;
  if (kind === "billed") return "Sale value · " + period;
  if (kind === "period") return period;
  if (kind === "throughput") {
    return (String((query && query.slice) || "") === "delivered" ? "Delivered · " : "Booked · ") + period;
  }
  if (kind === "category") return "CATERGORY · " + value;
  if (kind === "source") return "Source · " + value;
  if (kind === "product" || kind === "productIncome") return "Income · " + value;
  if (kind === "productQty") return "Orders · " + value;
  if (kind === "pipeline") return pipelineLabel(String((query && query.group) || value));
  if (kind === "age") {
    const slice = String((query && query.slice) || "").trim();
    return slice ? "Open " + value + " · " + slice : "Open " + value;
  }
  if (kind === "stuck") return value ? "Order " + value : "Oldest open jobs";
  if (kind === "open") return "Open jobs";
  if (kind === "drawing") return "Waiting for drawing";
  if (kind === "paint") return "At paint shop";
  if (kind === "ready_delivery") return "Ready for Delivery";
  if (kind === "readyOrOut") return "Ready / out for delivery";
  if (kind === "unpaid") return "Unpaid on open jobs";
  if (kind === "no_plates") return "No plates";
  return "Orders";
}

function buildDrill(query) {
  const win = resolveWindow(query || {});
  const kind = String((query && query.kind) || "");
  if (kind === "delivery" || kind === "late") return deliveryDrill(query);
  const rows = db.listOrders().filter((row) => matchesDrill(row, query, win)).map(cardOf);
  rows.sort((a, b) => String(b.sale_date).localeCompare(String(a.sale_date)) || String(a.order_number).localeCompare(String(b.order_number)));
  return {
    title: drillTitle(query, win),
    rows,
    totals: {
      count: rows.length,
      income: roundMoney(rows.reduce((s, r) => s + (r.amount_paid || 0), 0)),
      billed: roundMoney(rows.reduce((s, r) => s + (r.price_excl_vat || r.price_incl_vat || 0), 0)),
      owing: roundMoney(rows.reduce((s, r) => s + (r.owing || 0), 0))
    }
  };
}

module.exports = {
  buildDashboard,
  buildDrill,
  resolveWindow,
  saleDate,
  pipelineId,
  PIPELINE
};
