const db = require("./db");
const { CLOSED_STATUSES } = require("./enquiry-pipeline");

const SAST_OFFSET_MS = 2 * 60 * 60 * 1000;
const FOLLOW_UP_DAYS = 7;

const FUNNEL = [
  { id: "captured", label: "Captured" },
  { id: "costing", label: "Reached costing" },
  { id: "quoted", label: "Quoted" },
  { id: "followed", label: "Followed up" },
  { id: "ordered", label: "Ordered" }
];

const PIPELINE_ORDER = [
  "New",
  "Waiting on clients personal details",
  "Waiting on clients specifictions",
  "Waiting on productions confirmation",
  "Costing",
  "Re-Cost",
  "Waiting on Supplier",
  "Costed",
  "Quoted",
  "Followed Up",
  "Ordered",
  "Rejected",
  "Not Interested",
  "Not within scope"
];

const AGE_BUCKETS = [
  { id: "0-3d", maxDays: 3 },
  { id: "4-7d", maxDays: 7 },
  { id: "8-14d", maxDays: 14 },
  { id: "15-30d", maxDays: 30 },
  { id: "31-60d", maxDays: 60 },
  { id: "60d+", maxDays: Infinity }
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

function weekKey(d) {
  const p = sastParts(d);
  const date = new Date(Date.UTC(p.y, p.m, p.day));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const year = date.getUTCFullYear();
  const yearStart = new Date(Date.UTC(year, 0, 1));
  const week = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return year + "-W" + pad(week);
}

function weekLabel(key) {
  return key.replace("-W", " W");
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

function pickerMonths(rows) {
  const keys = new Set();
  const today = sastParts(new Date());
  for (let i = 0; i < 24; i++) {
    const p = addUtcMonths(today.y, today.m, -i);
    keys.add(p.y + "-" + pad(p.m + 1));
  }
  (rows || []).forEach((row) => {
    ["opened", "quoted", "ordered"].forEach((field) => {
      const t = parseWhen(row, field);
      if (t) keys.add(monthKey(new Date(t)));
    });
  });
  return Array.from(keys).sort().reverse().map((key) => ({ key, label: monthLabel(key) }));
}

function namedProducts(row) {
  const lines = Array.isArray(row.products) ? row.products : [];
  const named = lines.filter((p) => String((p && p.product) || "").trim());
  if (named.length) {
    return named.map((p) => ({
      product: String(p.product).trim(),
      category: String((p && p.category) || "").trim(),
      value_excl_vat: db.parseMoney(p.value_excl_vat)
    }));
  }
  return String(row.product || "").split(",").map((p) => p.trim()).filter(Boolean).map((product) => ({
    product,
    category: String(row.category || "").trim(),
    value_excl_vat: 0
  }));
}

function deliveryExclOf(row) {
  return db.parseMoney(row.delivery_excl_vat);
}

function productsTotalExclOf(row) {
  const fromLines = namedProducts(row).reduce((sum, p) => sum + (p.value_excl_vat || 0), 0);
  if (fromLines) return Math.round(fromLines * 100) / 100;
  return db.parseMoney(row.products_total_excl_vat);
}

function moneyOf(row) {
  const products = namedProducts(row);
  const productsTotal = productsTotalExclOf(row);
  const delivery = deliveryExclOf(row);
  const quote = db.parseMoney(row.quote_total_excl_vat) || Math.round((productsTotal + delivery) * 100) / 100;
  return {
    products,
    products_total_excl_vat: productsTotal,
    delivery_excl_vat: delivery,
    quote_total_excl_vat: quote
  };
}

function cardOf(row) {
  const money = moneyOf(row);
  return {
    enquiry_no: row.enquiry_no,
    client_name: row.client_name || "",
    status: statusOf(row),
    product: row.product || "",
    category: row.category || "",
    enquiry_type: row.enquiry_type || "",
    enquiry_source: row.enquiry_source || row.source || "",
    province: row.province || "",
    opened_at_label: row.opened_at_label || "",
    date_quoted: row.date_quoted || "",
    products: money.products,
    products_total_excl_vat: money.products_total_excl_vat,
    delivery_excl_vat: money.delivery_excl_vat,
    quote_total_excl_vat: money.quote_total_excl_vat || "",
    quote_no: row.quote_no || ""
  };
}

function parseWhen(row, field) {
  if (field === "opened") {
    if (row.created_at) {
      const t = Date.parse(row.created_at);
      if (Number.isFinite(t)) return t;
    }
    const d = db.asDate(row.date_enquired);
    return d ? d.getTime() : 0;
  }
  if (field === "quoted") {
    const d = db.asDate(row.date_quoted);
    if (d) return d.getTime();
    const ev = (row.events || []).filter((e) => e.kind === "complete_quote").slice(-1)[0];
    return ev && ev.at ? Date.parse(ev.at) : 0;
  }
  if (field === "ordered") {
    if (row.ordered_at) {
      const t = Date.parse(row.ordered_at);
      if (Number.isFinite(t)) return t;
    }
    const ev = (row.events || []).filter((e) => e.kind === "complete_order" || e.status === "Ordered").slice(-1)[0];
    return ev && ev.at ? Date.parse(ev.at) : 0;
  }
  return 0;
}

function inWindow(ms, from, to) {
  return ms >= from && ms <= to;
}

function statusOf(row) {
  return String(row.status || "New").trim() || "New";
}

function isClosed(status) {
  return CLOSED_STATUSES.indexOf(status) >= 0;
}

function reachedCosting(status) {
  return [
    "Costing", "Re-Cost", "Waiting on Supplier", "Costed",
    "Quoted", "Followed Up", "Ordered"
  ].indexOf(status) >= 0;
}

function reachedQuoted(status) {
  return ["Quoted", "Followed Up", "Ordered"].indexOf(status) >= 0;
}

function revenueOf(row) {
  return moneyOf(row).quote_total_excl_vat;
}

function bump(map, key, amount) {
  const k = key || "(blank)";
  map[k] = (map[k] || 0) + (amount == null ? 1 : amount);
}

function customKindLabels(row) {
  const specs = Array.isArray(row.custom_specs) ? row.custom_specs : [];
  const kinds = [];
  specs.forEach((spec) => {
    let kind = String((spec && spec.kind) || "").trim();
    const other = String((spec && spec.other) || "").trim();
    if (/^other$/i.test(kind) && other) kind = other;
    if (/^color$/i.test(kind)) kind = "Colour";
    if (!kind) kind = "Other";
    if (kinds.indexOf(kind) < 0) kinds.push(kind);
  });
  return kinds.length ? kinds : ["Other"];
}

function designKey(row) {
  return String(row.design_description || row.request || "").trim() || "(blank)";
}

function shorten(s, n) {
  const t = String(s || "").replace(/\s+/g, " ").trim();
  const max = n || 42;
  return t.length <= max ? t : t.slice(0, max - 1) + "…";
}

function weekOfMonth(d) {
  const p = sastParts(d);
  return Math.min(5, Math.floor((p.day - 1) / 7) + 1);
}

function emptyWom() {
  return {
    enquiries: [0, 0, 0, 0, 0],
    quotes: [0, 0, 0, 0, 0],
    ordered: [0, 0, 0, 0, 0],
    quoteRevenue: [0, 0, 0, 0, 0],
    orderedRevenue: [0, 0, 0, 0, 0]
  };
}

function ensureWom(map, key) {
  if (!map[key]) map[key] = emptyWom();
  return map[key];
}

function splitPairs(countMap, moneyMap) {
  return Object.keys(countMap).map((label) => ({
    label: shorten(label),
    value: label,
    count: countMap[label] || 0,
    quoteValue: Math.round((moneyMap[label] || 0) * 100) / 100
  })).sort((a, b) => b.count - a.count || a.label.localeCompare(b.label));
}

function sortedPairs(map, limit) {
  return Object.keys(map).map((label) => ({ label, value: map[label] }))
    .sort((a, b) => b.value - a.value || a.label.localeCompare(b.label))
    .slice(0, limit || 12);
}

function median(nums) {
  const list = nums.filter((n) => Number.isFinite(n) && n >= 0).sort((a, b) => a - b);
  if (!list.length) return null;
  const mid = Math.floor(list.length / 2);
  return list.length % 2 ? list[mid] : (list[mid - 1] + list[mid]) / 2;
}

function percentile(nums, p) {
  const list = nums.filter((n) => Number.isFinite(n) && n >= 0).sort((a, b) => a - b);
  if (!list.length) return null;
  const i = Math.min(list.length - 1, Math.max(0, Math.ceil(list.length * p) - 1));
  return list[i];
}

function round1(n) {
  return n == null ? null : Math.round(n * 10) / 10;
}

function countPairs(map, limit) {
  return sortedPairs(map, limit).map((row) => ({ label: row.label, count: row.value }));
}

function ageBucket(days) {
  for (const b of AGE_BUCKETS) {
    if (days <= b.maxDays) return b.id;
  }
  return "60d+";
}

function eventTime(row, kind) {
  const hit = (row.events || []).filter((e) => e.kind === kind).slice(-1)[0];
  return hit && hit.at ? Date.parse(hit.at) : 0;
}

function firstEvent(row, kind) {
  const hit = (row.events || []).find((e) => e.kind === kind);
  return hit && hit.at ? Date.parse(hit.at) : 0;
}

function followUpOverdue(row) {
  const status = statusOf(row);
  if (status !== "Quoted" && status !== "Followed Up") return false;
  const list = Array.isArray(row.follow_ups) ? row.follow_ups : [];
  const currentNo = String(row.quote_no || "").trim();
  const tagged = list.filter((f) => String((f && f.quote_no) || "").trim());
  const use = currentNo && tagged.length
    ? list.filter((f) => String((f && f.quote_no) || "").trim() === currentNo)
    : list;
  if (use.length >= 3) return false;
  let from = 0;
  if (use.length) from = Date.parse(use[use.length - 1].uploaded_at || "");
  if (!from) {
    const d = db.asDate(row.date_quoted);
    from = d ? d.getTime() : 0;
  }
  if (!from) return false;
  return Date.now() >= from + FOLLOW_UP_DAYS * 86400000;
}

function stageDays(row, fromKind, toKind) {
  const a = firstEvent(row, fromKind) || (fromKind === "created" ? parseWhen(row, "opened") : 0);
  const b = firstEvent(row, toKind) || (toKind === "complete_quote" ? parseWhen(row, "quoted") : 0)
    || (toKind === "complete_order" ? parseWhen(row, "ordered") : 0);
  if (!a || !b || b < a) return null;
  return (b - a) / 86400000;
}

function emptySeries(keys, grain) {
  const label = grain === "week" ? weekLabel : monthLabel;
  return keys.map((key) => ({
    key,
    label: label(key),
    enquiries: 0,
    quotes: 0,
    ordered: 0,
    enquiryRevenue: 0,
    quoteRevenue: 0,
    orderedRevenue: 0,
    rejected: 0,
    notInterested: 0,
    notInScope: 0
  }));
}

function buildDashboard(query) {
  const win = resolveWindow(query);
  const grain = win.grain;
  const range = win.range;
  const keys = bucketKeys(grain, win.from, win.to);
  const keyOf = grain === "week" ? weekKey : monthKey;
  const seriesMap = {};
  emptySeries(keys, grain).forEach((row) => { seriesMap[row.key] = row; });

  const rows = db.listEnquiries();
  const pipeline = {};
  PIPELINE_ORDER.forEach((s) => { pipeline[s] = 0; });
  const source = {};
  const type = {};
  const typeQuote = {};
  const customCount = {};
  const customQuote = {};
  const designCount = {};
  const designQuote = {};
  const province = {};
  const provinceOrdered = {};
  const provinceQuote = {};
  const provinceOrderedQuote = {};
  const category = {};
  const product = {};
  const productValueMap = {};
  const womMap = {};
  bucketKeys("month", win.from, win.to).forEach((key) => ensureWom(womMap, key));
  const assignee = {};
  const stuck = { Costing: [], Quoted: [], "Followed Up": [], Waiting: [] };
  const timeToOrderDays = [];
  const stage = {
    toCosting: [],
    costing: [],
    toQuote: [],
    quoteToOrder: []
  };
  let outlookWith = 0;
  let outlookWithout = 0;
  let funnelCaptured = 0;
  let funnelCosting = 0;
  let funnelQuoted = 0;
  let funnelFollowed = 0;
  let funnelOrdered = 0;
  let orderedInRange = 0;
  let overdueFollowUps = 0;
  let openNow = 0;
  let costingOpen = 0;
  let quotedWaiting = 0;
  let followUpDue = 0;
  let waitingOnSupplier = 0;
  let oldestOpenDays = null;
  let quotedExclVat = 0;
  let orderedExclVat = 0;
  let winRejected = 0;
  let winNotInterested = 0;
  let winNotInScope = 0;

  rows.forEach((row) => {
    const status = statusOf(row);
    const opened = parseWhen(row, "opened");
    const quoted = parseWhen(row, "quoted");
    const ordered = parseWhen(row, "ordered");
    const openedIn = opened && inWindow(opened, win.from, win.to);
    const quotedIn = quoted && inWindow(quoted, win.from, win.to);
    const orderedIn = ordered && inWindow(ordered, win.from, win.to);
    const rev = revenueOf(row);

    if (openedIn) {
      pipeline[status] = (pipeline[status] || 0) + 1;
      if (!isClosed(status) && status !== "Ordered") openNow += 1;
      if (followUpOverdue(row)) overdueFollowUps += 1;
      (row.tasks || []).forEach((t) => {
        if (t.status !== "open" || !t.assignee) return;
        bump(assignee, t.assignee);
      });
      const ageDays = opened ? (Date.now() - opened) / 86400000 : 0;
      if (!isClosed(status) && status !== "Ordered" && opened) {
        if (oldestOpenDays == null || ageDays > oldestOpenDays) oldestOpenDays = ageDays;
      }
      if (status === "Costing" || status === "Re-Cost") costingOpen += 1;
      if (status === "Quoted") quotedWaiting += 1;
      if (status === "Followed Up") followUpDue += 1;
      if (status === "Waiting on Supplier") waitingOnSupplier += 1;
      if (status === "Costing" || status === "Re-Cost" || status === "Waiting on Supplier") stuck.Costing.push(ageDays);
      if (status === "Quoted") stuck.Quoted.push(ageDays);
      if (status === "Followed Up") stuck["Followed Up"].push(ageDays);
      if (/Waiting on clients/.test(status)) stuck.Waiting.push(ageDays);
    }

    if (openedIn) {
      funnelCaptured += 1;
      if (reachedCosting(status) || reachedQuoted(status)) funnelCosting += 1;
      if (reachedQuoted(status)) funnelQuoted += 1;
      if (status === "Followed Up" || status === "Ordered") funnelFollowed += 1;
      if (status === "Ordered") funnelOrdered += 1;
      bump(source, row.enquiry_source || row.source);
      bump(type, row.enquiry_type);
      bump(typeQuote, row.enquiry_type, rev);
      bump(province, row.province);
      bump(provinceQuote, row.province, rev);
      if (status === "Ordered") {
        bump(provinceOrdered, row.province);
        bump(provinceOrderedQuote, row.province, rev);
      }
      bump(category, row.category);
      namedProducts(row).forEach((p) => {
        bump(product, p.product);
        bump(productValueMap, p.product, p.value_excl_vat);
      });
      const etype = String(row.enquiry_type || "").trim();
      if (etype === "Custom") {
        customKindLabels(row).forEach((kind) => {
          bump(customCount, kind);
          bump(customQuote, kind, rev);
        });
      }
      if (etype === "New Design") {
        const desc = designKey(row);
        bump(designCount, desc);
        bump(designQuote, desc, rev);
      }
      if (opened) {
        const bucket = ensureWom(womMap, monthKey(new Date(opened)));
        const w = weekOfMonth(new Date(opened)) - 1;
        bucket.enquiries[w] += 1;
      }
      const mails = (((row.correspondence || {}).mails) || []).length;
      if (mails) outlookWith += 1;
      else outlookWithout += 1;
      const bucket = seriesMap[keyOf(new Date(opened))];
      if (bucket) {
        bucket.enquiries += 1;
        bucket.enquiryRevenue += rev;
      }
    }
    if (quotedIn) {
      quotedExclVat += rev;
      const bucket = seriesMap[keyOf(new Date(quoted))];
      if (bucket) {
        bucket.quotes += 1;
        bucket.quoteRevenue += rev;
      }
      const wom = ensureWom(womMap, monthKey(new Date(quoted)));
      const w = weekOfMonth(new Date(quoted)) - 1;
      wom.quotes[w] += 1;
      wom.quoteRevenue[w] += rev;
    }
    if (orderedIn) {
      orderedInRange += 1;
      orderedExclVat += rev;
      const bucket = seriesMap[keyOf(new Date(ordered))];
      if (bucket) {
        bucket.ordered += 1;
        bucket.orderedRevenue += rev;
      }
      const wom = ensureWom(womMap, monthKey(new Date(ordered)));
      const w = weekOfMonth(new Date(ordered)) - 1;
      wom.ordered[w] += 1;
      wom.orderedRevenue[w] += rev;
      const days = (row.lifespan_ms ? row.lifespan_ms / 86400000 : 0)
        || (opened && ordered ? (ordered - opened) / 86400000 : 0);
      if (days > 0) timeToOrderDays.push(days);
    }
    if (openedIn) {
      const bucket = seriesMap[keyOf(new Date(opened))];
      if (bucket) {
        if (status === "Rejected") bucket.rejected += 1;
        if (status === "Not Interested") bucket.notInterested += 1;
        if (status === "Not within scope") bucket.notInScope += 1;
      }
      if (status === "Rejected") winRejected += 1;
      if (status === "Not Interested") winNotInterested += 1;
      if (status === "Not within scope") winNotInScope += 1;
      const toCost = stageDays(row, "created", "assign_costing");
      const costingDays = stageDays(row, "assign_costing", "complete_cost_sheet");
      const toQuote = stageDays(row, "complete_cost_sheet", "complete_quote");
      const qToO = stageDays(row, "complete_quote", "complete_order");
      if (toCost != null) stage.toCosting.push(toCost);
      if (costingDays != null) stage.costing.push(costingDays);
      if (toQuote != null) stage.toQuote.push(toQuote);
      if (qToO != null) stage.quoteToOrder.push(qToO);
    }
  });

  const series = keys.map((key) => seriesMap[key] || emptySeries([key], grain)[0]);
  const timeToOrderHist = AGE_BUCKETS.map((b) => ({
    label: b.id,
    count: timeToOrderDays.filter((d) => ageBucket(d) === b.id).length
  }));

  const stuckBars = Object.keys(stuck).map((label) => {
    const days = stuck[label];
    return {
      label,
      count: days.length,
      medianDays: round1(median(days))
    };
  });

  const provinceRows = sortedPairs(province, 12).map((row) => ({
    label: row.label,
    opened: row.value,
    ordered: provinceOrdered[row.label] || 0,
    conversion: row.value ? Math.round(1000 * (provinceOrdered[row.label] || 0) / row.value) / 10 : 0,
    quoteValue: Math.round((provinceQuote[row.label] || 0) * 100) / 100,
    orderedValue: Math.round((provinceOrderedQuote[row.label] || 0) * 100) / 100
  }));

  const productRows = Object.keys(product).map((label) => ({
    label,
    count: product[label],
    productValue: Math.round((productValueMap[label] || 0) * 100) / 100
  })).sort((a, b) => b.productValue - a.productValue || b.count - a.count || a.label.localeCompare(b.label));

  const typeRows = Object.keys(type).map((label) => ({
    label,
    count: type[label],
    quoteValue: Math.round((typeQuote[label] || 0) * 100) / 100
  })).sort((a, b) => b.count - a.count || a.label.localeCompare(b.label));

  const weekCompare = {
    weeks: [1, 2, 3, 4, 5],
    months: Object.keys(womMap).sort().map((key) => Object.assign({
      key,
      label: monthLabel(key)
    }, womMap[key]))
  };

  return {
    tz: "Africa/Johannesburg",
    grain,
    range,
    month: win.month || "",
    windowLabel: win.windowLabel,
    months: pickerMonths(rows),
    enquiryCount: rows.length,
    kpis: {
      openNow,
      quotedWaiting: (pipeline.Quoted || 0) + (pipeline["Followed Up"] || 0),
      overdueFollowUps,
      orderedInPeriod: orderedInRange,
      medianDaysToOrder: round1(median(timeToOrderDays)),
      p80DaysToOrder: round1(percentile(timeToOrderDays, 0.8))
    },
    funnel: [
      { label: "Captured", count: funnelCaptured },
      { label: "Reached costing", count: funnelCosting },
      { label: "Quoted", count: funnelQuoted },
      { label: "Followed up", count: funnelFollowed },
      { label: "Ordered", count: funnelOrdered }
    ],
    pipeline: PIPELINE_ORDER.map((status) => ({ status, count: pipeline[status] || 0 })).filter((r) => r.count),
    series,
    timeToOrder: { buckets: timeToOrderHist },
    stuck: {
      costingOpen,
      quotedWaiting,
      followUpDue,
      overdueFollowUps,
      waitingOnSupplier,
      oldestOpenDays: oldestOpenDays == null ? null : Math.round(oldestOpenDays),
      ageing: stuckBars
    },
    winLoss: [
      { label: "Ordered", count: orderedInRange },
      { label: "Rejected", count: winRejected },
      { label: "Not Interested", count: winNotInterested },
      { label: "Not within scope", count: winNotInScope }
    ],
    winLossByPeriod: series.map((s) => ({
      label: s.label,
      ordered: s.ordered,
      rejected: s.rejected,
      notInterested: s.notInterested,
      notInScope: s.notInScope
    })),
    sources: countPairs(source, 12),
    types: typeRows,
    typeSplits: {
      Custom: splitPairs(customCount, customQuote),
      "New Design": splitPairs(designCount, designQuote)
    },
    stageTime: [
      { label: "Capture → costing", medianDays: round1(median(stage.toCosting)) || 0, n: stage.toCosting.length },
      { label: "In costing", medianDays: round1(median(stage.costing)) || 0, n: stage.costing.length },
      { label: "Costed → quote", medianDays: round1(median(stage.toQuote)) || 0, n: stage.toQuote.length },
      { label: "Quote → order", medianDays: round1(median(stage.quoteToOrder)) || 0, n: stage.quoteToOrder.length }
    ],
    workload: countPairs(assignee, 20).map((row) => ({ name: row.label, count: row.count })),
    money: {
      quotedExclVat: Math.round(quotedExclVat * 100) / 100,
      orderedExclVat: Math.round(orderedExclVat * 100) / 100
    },
    categories: countPairs(category, 12),
    products: productRows,
    provinces: provinceRows,
    weekCompare
  };
}

function fieldOrBlank(v) {
  const s = String(v || "").trim();
  return s || "(blank)";
}

function inBucket(ms, grain, key) {
  if (!key) return true;
  if (!ms) return false;
  const keyOf = grain === "week" ? weekKey : monthKey;
  return keyOf(new Date(ms)) === key;
}

function matchesDrill(row, query, win) {
  const kind = String((query && query.kind) || "enquiries");
  const key = String((query && query.key) || "");
  const value = String((query && query.value) || "");
  const status = statusOf(row);
  const opened = parseWhen(row, "opened");
  const quoted = parseWhen(row, "quoted");
  const ordered = parseWhen(row, "ordered");
  const openedIn = !!(opened && inWindow(opened, win.from, win.to));
  const quotedIn = !!(quoted && inWindow(quoted, win.from, win.to));
  const orderedIn = !!(ordered && inWindow(ordered, win.from, win.to));

  if (kind === "enquiries") return openedIn && inBucket(opened, win.grain, key);
  if (kind === "quotes") return quotedIn && inBucket(quoted, win.grain, key);
  if (kind === "ordered") return orderedIn && inBucket(ordered, win.grain, key);
  if (kind === "funnel") {
    if (!openedIn) return false;
    const stage = String((query && query.stage) || "captured");
    if (stage === "costing") return reachedCosting(status) || reachedQuoted(status);
    if (stage === "quoted") return reachedQuoted(status);
    if (stage === "followed") return status === "Followed Up" || status === "Ordered";
    if (stage === "ordered") return status === "Ordered";
    return true;
  }
  if (kind === "pipeline") return openedIn && status === value;
  if (kind === "source") return openedIn && fieldOrBlank(row.enquiry_source || row.source) === value;
  if (kind === "type") return openedIn && fieldOrBlank(row.enquiry_type) === value;
  if (kind === "typeSubtype") {
    if (!openedIn) return false;
    const typeName = String((query && query.type) || "");
    if (fieldOrBlank(row.enquiry_type) !== typeName) return false;
    if (typeName === "Custom") return customKindLabels(row).indexOf(value) >= 0;
    if (typeName === "New Design") return designKey(row) === value;
    return true;
  }
  if (kind === "wom") {
    const ym = String((query && query.month) || "");
    const week = Number((query && query.week) || 0);
    const series = String((query && query.series) || "enquiries");
    if (series === "quotes" || series === "quoteRevenue") {
      return !!(quoted && monthKey(new Date(quoted)) === ym && weekOfMonth(new Date(quoted)) === week);
    }
    if (series === "ordered" || series === "orderedRevenue") {
      return !!(ordered && monthKey(new Date(ordered)) === ym && weekOfMonth(new Date(ordered)) === week);
    }
    return !!(opened && monthKey(new Date(opened)) === ym && weekOfMonth(new Date(opened)) === week);
  }
  if (kind === "category") return openedIn && fieldOrBlank(row.category) === value;
  if (kind === "product") {
    if (!openedIn) return false;
    const names = namedProducts(row).map((p) => p.product);
    return names.indexOf(value) >= 0 || (value === "(blank)" && !names.length);
  }
  if (kind === "province") {
    if (String((query && query.slice) || "") === "ordered") {
      return openedIn && status === "Ordered" && fieldOrBlank(row.province) === value;
    }
    return openedIn && fieldOrBlank(row.province) === value;
  }
  if (kind === "outlook") {
    if (!openedIn) return false;
    const mails = ((((row.correspondence || {}).mails) || []).length);
    return value === "with" ? mails > 0 : mails === 0;
  }
  if (kind === "workload") {
    if (!openedIn) return false;
    return (row.tasks || []).some((t) => t.status === "open" && t.assignee === value);
  }
  if (kind === "stuck") {
    if (!openedIn) return false;
    if (value === "Costing") return status === "Costing" || status === "Re-Cost" || status === "Waiting on Supplier";
    if (value === "Quoted") return status === "Quoted";
    if (value === "Followed Up") return status === "Followed Up";
    if (value === "Waiting on Supplier" || value === "supplier") return status === "Waiting on Supplier";
    if (value === "overdue") return followUpOverdue(row);
    return false;
  }
  if (kind === "winloss") {
    if (value === "Ordered") return orderedIn && inBucket(ordered, win.grain, key);
    if (!openedIn || !inBucket(opened, win.grain, key)) return false;
    if (value === "Rejected") return status === "Rejected";
    if (value === "Not Interested") return status === "Not Interested";
    if (value === "Not within scope") return status === "Not within scope";
    return false;
  }
  return openedIn;
}

function drillTitle(query, win) {
  const kind = String((query && query.kind) || "enquiries");
  const key = String((query && query.key) || "");
  const value = String((query && query.value) || "");
  const bucket = key ? (win.grain === "week" ? weekLabel(key) : monthLabel(key)) : win.windowLabel;
  const funnelNames = { captured: "Captured", costing: "Reached costing", quoted: "Quoted", followed: "Followed up", ordered: "Ordered" };
  if (kind === "enquiries") return "Enquiries · " + bucket;
  if (kind === "quotes") return "Quotes · " + bucket;
  if (kind === "ordered") return "Ordered · " + bucket;
  if (kind === "funnel") return (funnelNames[query.stage] || "Funnel") + " · " + win.windowLabel;
  if (kind === "typeSubtype") return (String((query && query.type) || "") + " · " + value + " · " + win.windowLabel).trim();
  if (kind === "wom") {
    const series = String((query && query.series) || "enquiries");
    const name = series === "quotes" || series === "quoteRevenue" ? "Quotes"
      : series === "ordered" || series === "orderedRevenue" ? "Ordered"
      : "Enquiries";
    const monthBit = monthLabel(String((query && query.month) || ""));
    const weekBit = String((query && query.week) || "");
    return name + " · " + (monthBit ? monthBit + " " : "") + "week " + weekBit;
  }
  if (kind === "pipeline") return (value || "Pipeline") + " · " + win.windowLabel;
  if (kind === "stuck") return (value === "overdue" ? "Overdue follow-ups" : value) + " · " + win.windowLabel;
  if (value) return value + " · " + win.windowLabel;
  return "Enquiries · " + win.windowLabel;
}

function buildDrill(query) {
  const win = resolveWindow(query);
  const rows = db.listEnquiries()
    .filter((row) => matchesDrill(row, query, win))
    .map(cardOf)
    .sort((a, b) => String(b.enquiry_no).localeCompare(String(a.enquiry_no), undefined, { numeric: true }));
  const totals = rows.reduce((acc, row) => {
    acc.products += Number(row.products_total_excl_vat) || 0;
    acc.delivery += Number(row.delivery_excl_vat) || 0;
    return acc;
  }, { products: 0, delivery: 0 });
  return {
    tz: "Africa/Johannesburg",
    grain: win.grain,
    range: win.range,
    month: win.month || "",
    windowLabel: win.windowLabel,
    kind: String((query && query.kind) || "enquiries"),
    key: String((query && query.key) || ""),
    title: drillTitle(query, win),
    totals: {
      products_excl_vat: Math.round(totals.products * 100) / 100,
      delivery_excl_vat: Math.round(totals.delivery * 100) / 100
    },
    rows
  };
}

module.exports = {
  buildDashboard,
  buildDrill,
  resolveWindow,
  weekKey,
  monthKey,
  FUNNEL
};
