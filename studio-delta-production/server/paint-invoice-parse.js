"use strict";

/**
 * Parse powder-shop / paint invoices so costs can be auto-filled onto orders.
 * Works on selectable-text PDFs (not photo scans without OCR).
 */

const { formatOrderId, parseMoney, money } = require("./db");

const ORDER_TOKEN_RE = /\bS\d{5,}(?:\([A-Za-z0-9,-]+\))?/gi;
// R2,219.80 or R2219.80 (must not stop at R2,21)
const RAND_NUM = "(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:[.,]\\d{2})";
const QTY_PRICE_TAIL_RE = new RegExp(
  "(?:^|\\s)(\\d+(?:[.,]\\d+)?)\\s+R\\s*(" + RAND_NUM + ")(?:\\s+[\\d]+[.,][\\d]+\\s*%){0,2}\\s+R\\s*(" + RAND_NUM + ")(?:\\s+R\\s*(" + RAND_NUM + "))?",
  "i"
);

function normalizeOrderToken(raw) {
  const s = String(raw || "").trim().toUpperCase();
  if (!s) return "";
  const base = s.replace(/\([^)]*\)/g, "");
  return formatOrderId(base) || formatOrderId(s);
}

function parseRandToken(raw) {
  const cleaned = String(raw || "")
    .replace(/\s/g, "")
    .replace(/R/gi, "")
    .replace(/,(?=\d{3}(?:\D|$))/g, "")
    .replace(",", ".");
  const n = parseMoney(cleaned);
  return Number.isFinite(n) ? Number(money(n)) : null;
}

function extractOrderTokens(text) {
  const out = [];
  const seen = {};
  String(text || "").replace(ORDER_TOKEN_RE, (m) => {
    const id = normalizeOrderToken(m);
    if (id && !seen[id]) {
      seen[id] = true;
      out.push(id);
    }
    return m;
  });
  return out;
}

function collapsePdfText(text) {
  return String(text || "")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/**
 * Walk invoice text and build per-order cost totals (excl VAT when available).
 */
function parsePaintInvoiceText(text, opts) {
  const options = opts || {};
  const preferOrders = (options.preferOrderNumbers || [])
    .map((n) => formatOrderId(n))
    .filter(Boolean);
  const preferSet = {};
  preferOrders.forEach((n) => { preferSet[n] = true; });

  const raw = collapsePdfText(text);
  if (!raw) {
    return {
      ok: false,
      error: "No text found in that PDF. Export a text PDF (not a photo scan).",
      allocations: [],
      unmatchedLines: [],
      orderTokens: []
    };
  }

  const lines = raw.split(/\n+/).map((l) => l.trim()).filter(Boolean);
  const allocations = {};
  const unmatchedLines = [];
  let activeOrder = preferOrders[0] || "";
  const headerOrders = extractOrderTokens(lines.slice(0, 12).join("\n"));
  if (!activeOrder && headerOrders.length === 1) activeOrder = headerOrders[0];

  function ensure(orderNumber) {
    if (!allocations[orderNumber]) {
      allocations[orderNumber] = {
        orderNumber,
        cost: 0,
        matched: !!preferSet[orderNumber],
        lines: []
      };
    }
    return allocations[orderNumber];
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const next = lines[i + 1] || "";
    const combo = line + " " + next;

    // Colour / job tags like S260244-MATT BLACK SMOOTH set the active order.
    const colourHit = line.match(/\b(S\d{5,})(?:\([A-Za-z0-9,-]+\))?\s*[-–—]/i);
    if (colourHit) {
      activeOrder = normalizeOrderToken(colourHit[1]);
    } else {
      const alone = line.match(/^\s*(S\d{5,}(?:\([A-Za-z0-9,-]+\))?)\s*$/i);
      if (alone) activeOrder = normalizeOrderToken(alone[1]);
    }

    const candidates = [line, combo];
    let parsed = null;
    for (const candidate of candidates) {
      const m = candidate.match(QTY_PRICE_TAIL_RE);
      if (!m) continue;
      const qty = parseMoney(String(m[1]).replace(",", "."));
      const unit = parseRandToken(m[2]);
      const excl = parseRandToken(m[3]);
      const incl = m[4] != null ? parseRandToken(m[4]) : null;
      if (!Number.isFinite(qty) || qty <= 0 || unit == null) continue;
      // Prefer explicit excl subtotal; else qty × unit.
      let lineExcl = excl != null ? excl : Number(money(qty * unit));
      // If only two R amounts look like unit + incl, derive excl.
      if (excl != null && incl != null && Math.abs(excl - unit) < 0.005 && qty > 1) {
        lineExcl = Number(money(qty * unit));
      }
      const desc = candidate
        .replace(QTY_PRICE_TAIL_RE, " ")
        .replace(/\s+/g, " ")
        .trim()
        .slice(0, 160);
      parsed = { qty, unitPrice: unit, lineExcl, lineIncl: incl, description: desc };
      if (candidate === combo) i += 1; // consumed next line
      break;
    }

    if (!parsed) continue;

    let orderNumber = activeOrder;
    const inDesc = extractOrderTokens(parsed.description);
    if (inDesc.length) orderNumber = inDesc[0];
    if (!orderNumber) {
      unmatchedLines.push(parsed);
      continue;
    }
    const bucket = ensure(orderNumber);
    bucket.cost = Number(money(bucket.cost + parsed.lineExcl));
    bucket.lines.push(parsed);
    bucket.matched = bucket.matched || !!preferSet[orderNumber];
  }

  let list = Object.keys(allocations).map((k) => allocations[k]).sort((a, b) => {
    if (a.matched !== b.matched) return a.matched ? -1 : 1;
    return String(a.orderNumber).localeCompare(String(b.orderNumber));
  });

  // If nothing keyed by order but we have preferred orders and one preferred + unmatched lines,
  // dump unmatched onto the single preferred order when only one is selected.
  if (!list.length && unmatchedLines.length && preferOrders.length === 1) {
    const orderNumber = preferOrders[0];
    const cost = unmatchedLines.reduce((sum, row) => sum + Number(row.lineExcl || 0), 0);
    list = [{
      orderNumber,
      cost: Number(money(cost)),
      matched: true,
      lines: unmatchedLines.slice(),
      assumedSingleOrder: true
    }];
    unmatchedLines.length = 0;
  }

  return {
    ok: true,
    orderTokens: extractOrderTokens(raw),
    allocations: list,
    unmatchedLines,
    textPreview: raw.slice(0, 1200)
  };
}

async function extractPdfText(buffer) {
  const pdfParse = require("pdf-parse");
  const data = await pdfParse(Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer || ""));
  return String((data && data.text) || "");
}

async function parsePaintInvoiceBuffer(buffer, opts) {
  const text = await extractPdfText(buffer);
  return parsePaintInvoiceText(text, opts);
}

module.exports = {
  normalizeOrderToken,
  parseRandToken,
  extractOrderTokens,
  parsePaintInvoiceText,
  extractPdfText,
  parsePaintInvoiceBuffer
};
