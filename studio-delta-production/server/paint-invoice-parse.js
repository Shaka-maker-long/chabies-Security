"use strict";

/**
 * Parse powder-shop / paint invoices so costs can be auto-filled onto orders.
 * Works on selectable-text PDFs (not photo scans without OCR).
 */

const { formatOrderId, parseMoney, money, formatRand } = require("./db");

const ORDER_TOKEN_RE = /\bS\d{5,}(?:\([A-Za-z0-9,-]+\))?/gi;
// R2,219.80 or R2219.80 (must not stop at R2,21)
const RAND_NUM = "(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:[.,]\\d{2})";
// Spaces are often missing in PDF extraction: 1R877.500,00%15,00%R877.50R1,009.13
const QTY_PRICE_TAIL_RE = new RegExp(
  "(?:^|\\s)(\\d+(?:[.,]\\d+)?)\\s*R\\s*(" + RAND_NUM + ")(?:\\s*[\\d]+[.,][\\d]+\\s*%){0,2}\\s*R\\s*(" + RAND_NUM + ")(?:\\s*R\\s*(" + RAND_NUM + "))?",
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

/** Insert spaces where PDF extract glued qty / Rand / % columns together. */
function relaxInvoiceText(text) {
  let s = String(text || "");
  // (2000X1000X450)1R877.50 → ) 1 R877.50
  s = s.replace(/\)(\d)/g, ") $1");
  s = s.replace(/(\d)R(\d)/gi, "$1 R$2");
  // R877.500,00% → R877.50 0,00%
  s = s.replace(new RegExp("(R\\s*" + RAND_NUM + ")(\\d+[.,]\\d+\\s*%)", "gi"), "$1 $2");
  s = s.replace(/(%\s*)(\d+[.,]\d+\s*%)/gi, "$1 $2");
  s = s.replace(/(%)(R)/gi, "$1 $2");
  s = s.replace(new RegExp("(R\\s*" + RAND_NUM + ")(R)", "gi"), "$1 $2");
  return s;
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

  const raw = relaxInvoiceText(collapsePdfText(text));
  if (!raw) {
    return {
      ok: false,
      error: "No text found in that PDF. Export a text PDF (not a photo scan).",
      allocations: [],
      unmatchedLines: [],
      orderTokens: [],
      grandTotal: 0,
      grandTotalLabel: formatRand(0)
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

  const COLOUR_TAG_RE = /\b(S\d{5,})(?:\([A-Za-z0-9,-]+\))?\s*[-–—/]/gi;
  const ALONE_ORDER_RE = /^\s*(S\d{5,}(?:\([A-Za-z0-9,-]+\))?)\s*$/i;
  const PRICE_GLOBAL_RE = new RegExp(QTY_PRICE_TAIL_RE.source, "gi");

  function colourTagsIn(text) {
    const out = [];
    const re = new RegExp(COLOUR_TAG_RE.source, "gi");
    let m;
    while ((m = re.exec(String(text || "")))) {
      out.push({ index: m.index, orderNumber: normalizeOrderToken(m[1]) });
    }
    return out;
  }

  function orderBefore(tags, atIndex, fallback) {
    let found = fallback || "";
    for (let t = 0; t < tags.length; t++) {
      if (tags[t].index <= atIndex) found = tags[t].orderNumber;
      else break;
    }
    return found;
  }

  function consumePriceMatch(fullText, match, orderNumber) {
    const qty = parseMoney(String(match[1]).replace(",", "."));
    const unit = parseRandToken(match[2]);
    const excl = parseRandToken(match[3]);
    const incl = match[4] != null ? parseRandToken(match[4]) : null;
    if (!Number.isFinite(qty) || qty <= 0 || unit == null) return null;
    let lineExcl = excl != null ? excl : Number(money(qty * unit));
    if (excl != null && incl != null && Math.abs(excl - unit) < 0.005 && qty > 1) {
      lineExcl = Number(money(qty * unit));
    }
    const start = match.index;
    const end = start + match[0].length;
    const desc = (fullText.slice(Math.max(0, start - 120), start) + " " + fullText.slice(end, end + 40))
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 160);
    return { qty, unitPrice: unit, lineExcl, lineIncl: incl, description: desc, orderNumber };
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const alone = line.match(ALONE_ORDER_RE);
    if (alone) activeOrder = normalizeOrderToken(alone[1]);

    const lineTags = colourTagsIn(line);
    if (lineTags.length) activeOrder = lineTags[lineTags.length - 1].orderNumber;

    const priceRe = new RegExp(PRICE_GLOBAL_RE.source, "gi");
    const matches = [];
    let m;
    while ((m = priceRe.exec(line))) {
      matches.push({
        1: m[1], 2: m[2], 3: m[3], 4: m[4],
        index: m.index,
        0: m[0]
      });
    }

    // If this line has no prices but the next line is only a price tail, join once.
    if (!matches.length && lines[i + 1]) {
      const combo = line + " " + lines[i + 1];
      const comboRe = new RegExp(PRICE_GLOBAL_RE.source, "gi");
      const comboMatches = [];
      while ((m = comboRe.exec(combo))) {
        comboMatches.push({
          1: m[1], 2: m[2], 3: m[3], 4: m[4],
          index: m.index,
          0: m[0]
        });
      }
      if (comboMatches.length) {
        const tags = colourTagsIn(combo);
        comboMatches.forEach((hit) => {
          let orderNumber = orderBefore(tags, hit.index, activeOrder);
          const parsed = consumePriceMatch(combo, hit, orderNumber);
          if (!parsed) return;
          if (!orderNumber) {
            unmatchedLines.push(parsed);
            return;
          }
          activeOrder = orderNumber;
          const bucket = ensure(orderNumber);
          bucket.cost = Number(money(bucket.cost + parsed.lineExcl));
          bucket.lines.push(parsed);
          bucket.matched = bucket.matched || !!preferSet[orderNumber];
        });
        i += 1;
        continue;
      }
    }

    matches.forEach((hit) => {
      let orderNumber = orderBefore(lineTags, hit.index, activeOrder);
      const parsed = consumePriceMatch(line, hit, orderNumber);
      if (!parsed) return;
      if (!orderNumber) {
        unmatchedLines.push(parsed);
        return;
      }
      activeOrder = orderNumber;
      const bucket = ensure(orderNumber);
      bucket.cost = Number(money(bucket.cost + parsed.lineExcl));
      bucket.lines.push(parsed);
      bucket.matched = bucket.matched || !!preferSet[orderNumber];
    });
  }

  let list = Object.keys(allocations).map((k) => {
    const row = allocations[k];
    return Object.assign({}, row, {
      costLabel: formatRand(row.cost)
    });
  }).sort((a, b) => {
    if (a.matched !== b.matched) return a.matched ? -1 : 1;
    return String(a.orderNumber).localeCompare(String(b.orderNumber));
  });

  if (!list.length && unmatchedLines.length && preferOrders.length === 1) {
    const orderNumber = preferOrders[0];
    const cost = unmatchedLines.reduce((sum, row) => sum + Number(row.lineExcl || 0), 0);
    list = [{
      orderNumber,
      cost: Number(money(cost)),
      costLabel: formatRand(cost),
      matched: true,
      lines: unmatchedLines.slice(),
      assumedSingleOrder: true
    }];
    unmatchedLines.length = 0;
  }

  const grandTotal = list.reduce((sum, row) => sum + Number(row.cost || 0), 0);

  return {
    ok: true,
    orderTokens: extractOrderTokens(raw),
    allocations: list,
    unmatchedLines,
    grandTotal: Number(money(grandTotal)),
    grandTotalLabel: formatRand(grandTotal),
    textPreview: raw.slice(0, 1600)
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
  relaxInvoiceText,
  parsePaintInvoiceText,
  extractPdfText,
  parsePaintInvoiceBuffer
};
