"use strict";

const db = require("./db");
const shop = require("./shop-status");
const orderDash = require("./order-dashboard");

const MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
const FORECAST_FROM = "AUG 2026";

function sastYear(d) {
  const x = d instanceof Date ? d : new Date();
  return new Date(x.getTime() + 2 * 60 * 60 * 1000).getUTCFullYear();
}

function monthTagFromDate(d) {
  if (!d || isNaN(d.getTime())) return "";
  const sast = new Date(d.getTime() + 2 * 60 * 60 * 1000);
  return MONTHS[sast.getUTCMonth()] + " " + sast.getUTCFullYear();
}

function monthTag(raw, fallbackDate) {
  const s = String(raw || "").trim().toUpperCase().replace(/,/g, " ").replace(/\s+/g, " ");
  const named = s.match(/\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)[A-Z]*\s+(\d{4})\b/);
  if (named) return named[1] + " " + named[2];
  const iso = s.match(/^(\d{4})-(\d{2})/);
  if (iso) {
    const m = Number(iso[2]) - 1;
    if (m >= 0 && m < 12) return MONTHS[m] + " " + iso[1];
  }
  const d = db.asDate(raw) || db.asDate(fallbackDate);
  return monthTagFromDate(d);
}

function quoteToken(raw) {
  const s = String(raw || "").trim().toUpperCase().replace(/\s+/g, "");
  if (!s) return "";
  if (/^\d+$/.test(s)) return s;
  return db.normalizeQuoteNo(raw, { allowEmpty: true }) || s;
}

function quoteTokens(raw) {
  const out = [];
  const seen = {};
  String(raw || "").split(/[\s&,\/;]+/).forEach((part) => {
    const t = quoteToken(part);
    if (!t || seen[t]) return;
    seen[t] = true;
    out.push(t);
  });
  return out;
}

function enquiryQuotes(row) {
  const tokens = quoteTokens(row && row.quote_no);
  const quotes = Array.isArray(row && row.quotes) ? row.quotes : [];
  quotes.forEach((q) => {
    quoteTokens(q && q.quote_no).forEach((t) => {
      if (tokens.indexOf(t) === -1) tokens.push(t);
    });
  });
  return tokens;
}

function isQuoted(row) {
  return enquiryQuotes(row).length > 0;
}

function isSoqToken(token) {
  return String(token || "").indexOf("SOQ") === 0;
}

function enquiryRev(row) {
  const quotes = Array.isArray(row && row.quotes) ? row.quotes : [];
  let best = 0;
  quotes.forEach((q) => {
    const n = db.parseMoney(q && (q.price_excl_vat || q.excl || q.total_excl));
    if (n > best) best = n;
  });
  return best;
}

function orderRev(row) {
  const stored = db.parseMoney(row && row.price_excl_vat);
  if (stored) return stored;
  return db.parseMoney(db.exclFromIncl(row && row.price_incl_vat));
}

function build(query) {
  const year = Number((query && query.year) || sastYear(new Date())) || sastYear(new Date());
  const yearBit = String(year);
  const months = MONTHS.map((m) => m + " " + year);

  const enquiries = [];
  (db.listEnquiries() || []).forEach((row) => {
    const month = monthTag(row.month_enquired, row.date_enquired);
    if (!month || month.indexOf(yearBit) === -1) return;
    const quotes = enquiryQuotes(row);
    const date = db.asDate(row.date_enquired);
    enquiries.push({
      no: row.enquiry_no || "",
      month,
      client: String(row.client_name || "").trim() || "Unknown",
      source: String(row.source || row.enquiry_source || "").trim() || "Unknown",
      quoteNo: quotes[0] || "",
      quotes,
      status: row.status || "",
      rev: enquiryRev(row),
      date: date ? db.formatPaymentDate(date) : String(row.date_enquired || ""),
      is_q: isQuoted(row),
      is_o: String(row.status || "").toUpperCase().indexOf("ORDER") !== -1
    });
  });

  const quoteSet = {};
  enquiries.forEach((e) => {
    (e.quotes || []).forEach((q) => { quoteSet[q] = true; });
  });

  const orders = [];
  (db.listOrders() || []).forEach((row) => {
    if (shop.isCancelled(row.status) || shop.isOnHold(row.status)) return;
    const sale = orderDash.saleDate(row);
    const month = monthTag(row.month_of_sale, sale || row.payment_date);
    if (!month || month.indexOf(yearBit) === -1) return;
    const quoteNo = quoteToken(row.quote_number);
    const soq = isSoqToken(quoteNo);
    const linked = !!(quoteNo && quoteSet[quoteNo]);
    orders.push({
      no: row.order_number || "",
      month,
      client: String(row.client_name || "").trim() || "Unknown",
      source: String(row.source || "").trim() || (soq ? "Offline Direct" : (quoteNo ? "Website" : "Unknown")),
      quoteNo,
      status: row.status || "",
      rev: orderRev(row),
      date: sale ? db.formatPaymentDate(sale) : String(row.payment_date || ""),
      isLinked: linked,
      isSoq: soq
    });
  });

  const orderByQuote = {};
  orders.forEach((o) => {
    if (o.quoteNo) orderByQuote[o.quoteNo] = o.month;
  });
  enquiries.forEach((e) => {
    e.isLinked = (e.quotes || []).some((q) => !!orderByQuote[q]);
  });

  return {
    year,
    months,
    forecastFrom: FORECAST_FROM,
    enquiries,
    orders,
    enquiryCount: enquiries.length,
    orderCount: orders.length
  };
}

module.exports = {
  MONTHS,
  FORECAST_FROM,
  monthTag,
  quoteToken,
  quoteTokens,
  enquiryQuotes,
  isQuoted,
  isSoqToken,
  build
};
