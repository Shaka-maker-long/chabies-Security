"use strict";

const db = require("./db");
const orderDash = require("./order-dashboard");

const SAST_OFFSET_MS = 2 * 60 * 60 * 1000;
const TOP_LINE_CAMPAIGNS = 8;
const PALETTE = [
  "#027a48", "#3538cd", "#b54708", "#b42318", "#026aa2",
  "#6941c6", "#1d2939", "#667085", "#dd2590", "#3b7c0f"
];

function pad(n) {
  return String(n).padStart(2, "0");
}

function sastParts(d) {
  const sast = new Date(d.getTime() + SAST_OFFSET_MS);
  return { y: sast.getUTCFullYear(), m: sast.getUTCMonth(), day: sast.getUTCDate() };
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

function bucketOf(ms, grain) {
  const d = new Date(ms);
  return grain === "week" ? weekKey(d) : monthKey(d);
}

function roundMoney(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
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
    paid: roundMoney(paid)
  };
}

function campaignLabel(row) {
  const s = String((row && row.campaign) || "").trim();
  return s || "(No campaign)";
}

function pickerMonths(rows) {
  const keys = new Set();
  const today = sastParts(new Date());
  for (let i = 0; i < 24; i++) {
    const p = addUtcMonths(today.y, today.m, -i);
    keys.add(p.y + "-" + pad(p.m + 1));
  }
  (rows || []).forEach((row) => {
    const d = orderDash.saleDate(row);
    if (d) keys.add(monthKey(d));
  });
  return Array.from(keys).sort().reverse().map((key) => ({ key, label: monthLabel(key) }));
}

function buildDashboard(query) {
  const win = orderDash.resolveWindow(query || {});
  const keys = bucketKeys(win.grain, win.from, win.to);
  const labelOf = win.grain === "week" ? weekLabel : monthLabel;
  const seriesShell = keys.map((key) => ({
    key,
    label: labelOf(key),
    income: 0,
    orders: 0
  }));

  const byCampaign = {};
  let income = 0;
  let billed = 0;
  let orderCount = 0;
  let missingIncome = 0;
  let missingOrders = 0;

  db.listOrders().forEach((row) => {
    const d = orderDash.saleDate(row);
    if (!d) return;
    const ms = d.getTime();
    if (ms < win.from || ms > win.to) return;
    const money = moneyOf(row);
    const campaign = campaignLabel(row);
    const key = bucketOf(ms, win.grain);
    if (!byCampaign[campaign]) {
      byCampaign[campaign] = {
        label: campaign,
        income: 0,
        billed: 0,
        orders: 0,
        byBucket: {}
      };
    }
    const c = byCampaign[campaign];
    c.income = roundMoney(c.income + money.paid);
    c.billed = roundMoney(c.billed + money.billed);
    c.orders += 1;
    c.byBucket[key] = roundMoney((c.byBucket[key] || 0) + money.paid);

    income = roundMoney(income + money.paid);
    billed = roundMoney(billed + money.billed);
    orderCount += 1;
    if (campaign === "(No campaign)") {
      missingIncome = roundMoney(missingIncome + money.paid);
      missingOrders += 1;
    }

    const point = seriesShell.find((s) => s.key === key);
    if (point) {
      point.income = roundMoney(point.income + money.paid);
      point.orders += 1;
    }
  });

  const campaigns = Object.keys(byCampaign).map((label) => {
    const c = byCampaign[label];
    return {
      label,
      income: c.income,
      billed: c.billed,
      orders: c.orders,
      share: income > 0 ? Math.round((c.income / income) * 1000) / 10 : 0,
      series: keys.map((key) => roundMoney(c.byBucket[key] || 0))
    };
  }).sort((a, b) => b.income - a.income || a.label.localeCompare(b.label));

  const lineCampaigns = campaigns
    .filter((c) => c.label !== "(No campaign)")
    .slice(0, TOP_LINE_CAMPAIGNS);
  // Always include missing campaign on the chart when it has income, so gaps stay visible.
  const missing = campaigns.find((c) => c.label === "(No campaign)");
  const chartCampaigns = missing && missing.income > 0
    ? lineCampaigns.concat([missing]).slice(0, TOP_LINE_CAMPAIGNS + 1)
    : lineCampaigns;

  return {
    windowLabel: win.windowLabel,
    grain: win.grain,
    range: win.range,
    month: win.month,
    months: pickerMonths(db.listOrders()),
    orderCount,
    kpis: {
      income,
      billed,
      orders: orderCount,
      campaigns: campaigns.filter((c) => c.label !== "(No campaign)").length,
      missingOrders,
      missingIncome
    },
    series: seriesShell,
    campaigns,
    chart: {
      labels: seriesShell.map((s) => s.label),
      keys,
      datasets: chartCampaigns.map((c, i) => ({
        label: c.label,
        data: c.series,
        color: PALETTE[i % PALETTE.length]
      }))
    }
  };
}

module.exports = {
  buildDashboard,
  campaignLabel,
  TOP_LINE_CAMPAIGNS
};
