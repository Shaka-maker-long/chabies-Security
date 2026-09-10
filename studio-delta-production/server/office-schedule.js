const SCHEDULE_CODES = [
  { code: "LD", label: "Latest Delivery", group: "Delivery", bg: "#ff0000", fg: "#ffffff" },
  { code: "LC", label: "Latest Courier", group: "Delivery", bg: "#ff9900", fg: "#1d2939" },
  { code: "LD*", label: "Moved Latest Delivery", group: "Moved", bg: "#d0d5dd", fg: "#667085", moved: true },
  { code: "LC*", label: "Moved Latest Courier", group: "Moved", bg: "#d0d5dd", fg: "#667085", moved: true },
  { code: "PD", label: "Planned Delivery", group: "Delivery", bg: "#00ff00", fg: "#1d2939" },
  { code: "C", label: "Planned Courier", group: "Delivery", bg: "#00ffff", fg: "#1d2939" },
  { code: "A", label: "Assemble", group: "Production", bg: "#f4cccc", fg: "#1d2939" },
  { code: "M", label: "Manufacturing", group: "Production", bg: "#fff2cc", fg: "#1d2939" },
  { code: "U", label: "Upholstery", group: "Production", bg: "#4a86e8", fg: "#ffffff" },
  { code: "PC", label: "Powder Coating", group: "Production", bg: "#d9d9d9", fg: "#1d2939" },
  { code: "P", label: "Photograpy", group: "Production", bg: "#cfe2f3", fg: "#1d2939" },
  { code: "QC", label: "QC", group: "Production", bg: "#fce5cd", fg: "#1d2939" },
  { code: "CS", label: "Cutting Steel", group: "Production", bg: "#b45f06", fg: "#ffffff" },
  { code: "Gr", label: "Grinding", group: "Production", bg: "#b6d7a8", fg: "#1d2939" },
  { code: "Wr", label: "Wrapping", group: "Production", bg: "#d9ead3", fg: "#1d2939" }
];

const DELIVERY_CODES = ["LD", "LC"];
const MOVED_DELIVERY_CODES = ["LD*", "LC*"];

function liveDeliveryCode(value) {
  const c = String(value || "").trim().toUpperCase();
  if (c === "LD" || c === "LC") return c;
  return "";
}

function starredDeliveryCode(value) {
  const live = liveDeliveryCode(value);
  return live ? live + "*" : "";
}
const SCHEDULE_WORKDAYS = 180;

function parseDay(iso) {
  const s = String(iso || "").slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
  const d = new Date(s + "T12:00:00");
  return Number.isNaN(d.getTime()) ? null : d;
}

function mondayOf(dateIso) {
  const fallback = new Date().toISOString().slice(0, 10);
  const d = parseDay(dateIso) || parseDay(fallback);
  const day = d.getDay() || 7;
  d.setDate(d.getDate() - day + 1);
  return d.toISOString().slice(0, 10);
}

function workdays(fromIso, days) {
  const out = [];
  const d = parseDay(fromIso) || parseDay(mondayOf());
  while (out.length < days) {
    const dow = d.getDay();
    if (dow !== 0 && dow !== 6) out.push(d.toISOString().slice(0, 10));
    d.setDate(d.getDate() + 1);
  }
  return out;
}

function isoWeekInfo(iso) {
  const d = parseDay(iso) || parseDay(mondayOf());
  const utc = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = utc.getUTCDay() || 7;
  utc.setUTCDate(utc.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(utc.getUTCFullYear(), 0, 1));
  const week = Math.ceil((((utc - yearStart) / 86400000) + 1) / 7);
  return { year: utc.getUTCFullYear(), week };
}

function weekKey(info) {
  return info.year + "-W" + String(info.week).padStart(2, "0");
}

function weekdayLong(iso) {
  const d = parseDay(iso);
  return d ? d.toLocaleDateString("en-GB", { weekday: "long" }) : "";
}

function formatDayLabel(iso) {
  const d = parseDay(iso);
  if (!d) return "";
  return d.toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "short" });
}

const MONTH_IX = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };

function formatOrderDate(v) {
  const s = String(v || "").trim();
  if (!s) return "";
  let d = parseDay(s);
  if (!d) {
    const dmy = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (dmy) d = new Date(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1]), 12, 0, 0);
  }
  if (!d) {
    const named = s.match(/^(\d{1,2})[-\s]+([A-Za-z]{3,})(?:[-\s,]+(\d{4}))?$/);
    if (named) {
      const m = MONTH_IX[named[2].slice(0, 3).toLowerCase()];
      if (m != null) {
        const y = named[3] ? Number(named[3]) : new Date().getFullYear();
        d = new Date(y, m, Number(named[1]), 12, 0, 0);
      }
    }
  }
  if (!d || Number.isNaN(d.getTime())) return s;
  const mon = d.toLocaleString("en-GB", { month: "short" }).replace(/\./g, "").slice(0, 3);
  return String(d.getDate()).padStart(2, "0") + "-" + mon;
}

function gridWeekLabel(iso) {
  return "Week" + String(isoWeekInfo(iso).week).padStart(2, "0");
}

function collectDeliveryItems(rows, cells) {
  const byId = new Map((rows || []).map((r) => [r.id, r]));
  const items = [];
  for (const c of cells || []) {
    const code = String(c.value || "").trim().toUpperCase();
    if (DELIVERY_CODES.indexOf(code) === -1) continue;
    const row = byId.get(c.row_id);
    if (!row) continue;
    const day = String(c.day || "").slice(0, 10);
    if (!parseDay(day)) continue;
    const info = isoWeekInfo(day);
    items.push({
      day,
      year: info.year,
      week: info.week,
      weekKey: weekKey(info),
      weekLabel: "Week " + info.week,
      weekday: weekdayLong(day),
      dateLabel: formatDayLabel(day),
      order_number: row.order_number,
      product: row.product || "",
      category: row.category || "",
      status: row.status || "",
      code,
      codeLabel: code === "LD" ? "Latest Delivery" : "Latest Courier"
    });
  }
  items.sort((a, b) => a.day.localeCompare(b.day) || String(a.order_number).localeCompare(String(b.order_number)));
  return items;
}

function weekOptions(items, aroundIso) {
  const monday = mondayOf(aroundIso);
  const map = new Map();
  function add(iso) {
    const info = isoWeekInfo(iso);
    const key = weekKey(info);
    if (!map.has(key)) {
      map.set(key, {
        key,
        year: info.year,
        week: info.week,
        start: mondayOf(iso),
        label: "Week " + info.week
      });
    }
  }
  for (let i = -12; i <= 16; i++) {
    const d = parseDay(monday);
    d.setDate(d.getDate() + i * 7);
    add(d.toISOString().slice(0, 10));
  }
  (items || []).forEach((it) => add(it.day));
  return [...map.values()].sort((a, b) => a.start.localeCompare(b.start));
}

module.exports = {
  SCHEDULE_CODES,
  DELIVERY_CODES,
  MOVED_DELIVERY_CODES,
  liveDeliveryCode,
  starredDeliveryCode,
  SCHEDULE_WORKDAYS,
  mondayOf,
  workdays,
  isoWeekInfo,
  weekKey,
  weekdayLong,
  formatDayLabel,
  formatOrderDate,
  gridWeekLabel,
  collectDeliveryItems,
  weekOptions
};
