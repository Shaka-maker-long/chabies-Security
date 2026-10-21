"use strict";

/**
 * Parse powder-shop / paint invoices so costs can be auto-filled onto orders.
 * Works on selectable-text PDFs (not photo scans without OCR).
 *
 * Letter ranges on the invoice like S260249(A-L) expand to Studio Delta
 * split orders: S260249 A … S260249 L, with the line excl total split across them.
 */

const { formatOrderId, parseMoney, money, formatRand } = require("./db");

// Letter split orders are "S260249 A" (single letter). Do not eat the "S" from "SD-CABINET…".
const ORDER_TOKEN_RE = /\bS\d{5,}(?:\([A-Za-z0-9,-]+\)|\s+[A-Z](?![A-Z]))?/gi;
// R2,219.80 or R2219.80 (must not stop at R2,21)
const RAND_NUM = "(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:[.,]\\d{2})";
// Spaces are often missing in PDF extraction: 1R877.500,00%15,00%R877.50R1,009.13
const QTY_PRICE_TAIL_RE = new RegExp(
  "(?:^|\\s)(\\d+(?:[.,]\\d+)?)\\s*R\\s*(" + RAND_NUM + ")(?:\\s*[\\d]+[.,][\\d]+\\s*%){0,2}\\s*R\\s*(" + RAND_NUM + ")(?:\\s*R\\s*(" + RAND_NUM + "))?",
  "i"
);

function lettersFromSpec(spec) {
  const s = String(spec || "").trim().toUpperCase().replace(/\s+/g, "");
  if (!s) return [];
  if (/^[A-Z]-[A-Z]$/.test(s)) {
    const a = s.charCodeAt(0);
    const b = s.charCodeAt(2);
    const out = [];
    for (let c = Math.min(a, b); c <= Math.max(a, b); c++) out.push(String.fromCharCode(c));
    return out;
  }
  if (/^[A-Z]$/.test(s)) return [s];
  if (s.indexOf(",") !== -1) {
    const out = [];
    s.split(",").forEach((part) => {
      lettersFromSpec(part).forEach((L) => {
        if (out.indexOf(L) === -1) out.push(L);
      });
    });
    return out;
  }
  return [];
}

/** Expand invoice refs like S260249(A-L) into shop order numbers S260249 A … S260249 L. */
function expandOrderRef(raw) {
  const s = String(raw || "").trim().toUpperCase();
  if (!s) return { base: "", letters: [], orderNumbers: [] };
  // No trailing \b — it would stop at S260249 before "(A-L)" because "(" is a boundary.
  // Single trailing letter only (S260249 A), never the "S" from "SD-CABINET…".
  const m = s.match(/\b(S\d{5,})(?:\(([A-Za-z0-9,-]+)\))?(?:\s+([A-Z])(?![A-Z]))?/);
  if (!m) {
    const id = formatOrderId(s);
    return { base: id, letters: [], orderNumbers: id ? [id] : [] };
  }
  const base = formatOrderId(m[1]);
  let letters = lettersFromSpec(m[2] || "");
  if (!letters.length && m[3]) letters = [m[3]];
  const orderNumbers = letters.length
    ? letters.map((L) => formatOrderId(base + " " + L))
    : [base];
  return { base, letters, orderNumbers: orderNumbers.filter(Boolean) };
}

function normalizeOrderToken(raw) {
  const expanded = expandOrderRef(raw);
  if (expanded.orderNumbers.length === 1) return expanded.orderNumbers[0];
  return expanded.base || "";
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
    expandOrderRef(m).orderNumbers.forEach((id) => {
      if (id && !seen[id]) {
        seen[id] = true;
        out.push(id);
      }
    });
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

/** Split cents across n orders so the sum still equals the line total. */
function splitCostShares(total, n) {
  const count = Math.max(1, n | 0);
  const cents = Math.round(Number(money(total)) * 100);
  const base = Math.floor(cents / count);
  let rem = cents - base * count;
  const shares = [];
  for (let i = 0; i < count; i++) {
    const c = base + (rem > 0 ? 1 : 0);
    if (rem > 0) rem -= 1;
    shares.push(Number(money(c / 100)));
  }
  return shares;
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
  let activeOrderNumbers = [];
  const headerOrders = extractOrderTokens(lines.slice(0, 12).join("\n"));
  if (!activeOrderNumbers.length && headerOrders.length === 1) {
    activeOrderNumbers = [headerOrders[0]];
  }

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

  // Capture full token including (A-L) before the colour dash.
  const COLOUR_TAG_RE = /\b(S\d{5,}(?:\([A-Za-z0-9,-]+\))?)\s*[-–—/]/gi;
  const ALONE_ORDER_RE = /^\s*(S\d{5,}(?:\([A-Za-z0-9,-]+\)|\s+[A-Z](?![A-Z]))?)\s*$/i;
  const PRICE_GLOBAL_RE = new RegExp(QTY_PRICE_TAIL_RE.source, "gi");

  function colourTagsIn(text) {
    const out = [];
    const re = new RegExp(COLOUR_TAG_RE.source, "gi");
    let m;
    while ((m = re.exec(String(text || "")))) {
      const expanded = expandOrderRef(m[1]);
      out.push({
        index: m.index,
        raw: m[1],
        orderNumbers: expanded.orderNumbers.slice()
      });
    }
    return out;
  }

  function orderNumbersBefore(tags, atIndex, fallback) {
    let found = fallback && fallback.length ? fallback.slice() : [];
    for (let t = 0; t < tags.length; t++) {
      if (tags[t].index <= atIndex) found = tags[t].orderNumbers.slice();
      else break;
    }
    return found;
  }

  function colourAhead(fromLine, afterIndexOnLine) {
    for (let j = fromLine; j < Math.min(lines.length, fromLine + 5); j++) {
      const tags = colourTagsIn(lines[j]);
      for (let t = 0; t < tags.length; t++) {
        if (j === fromLine && tags[t].index <= afterIndexOnLine) continue;
        if (tags[t].orderNumbers.length) return tags[t].orderNumbers.slice();
      }
      const alone = lines[j].match(ALONE_ORDER_RE);
      if (alone && (j > fromLine || alone.index > afterIndexOnLine)) {
        const expanded = expandOrderRef(alone[1]);
        if (expanded.orderNumbers.length) return expanded.orderNumbers;
      }
    }
    return [];
  }

  function consumePriceMatch(fullText, match) {
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
    return { qty, unitPrice: unit, lineExcl, lineIncl: incl, description: desc };
  }

  function addParsedToOrders(parsed, orderNumbers) {
    const targets = (orderNumbers || []).filter(Boolean);
    if (!targets.length) {
      unmatchedLines.push(parsed);
      return;
    }
    if (targets.length === 1) {
      const bucket = ensure(targets[0]);
      bucket.cost = Number(money(bucket.cost + parsed.lineExcl));
      bucket.lines.push(Object.assign({}, parsed, { orderNumber: targets[0] }));
      bucket.matched = bucket.matched || !!preferSet[targets[0]];
      return;
    }
    const shares = splitCostShares(parsed.lineExcl, targets.length);
    targets.forEach((orderNumber, idx) => {
      const share = shares[idx];
      const bucket = ensure(orderNumber);
      bucket.cost = Number(money(bucket.cost + share));
      bucket.lines.push(Object.assign({}, parsed, {
        orderNumber,
        lineExcl: share,
        splitFrom: parsed.lineExcl,
        splitCount: targets.length
      }));
      bucket.matched = bucket.matched || !!preferSet[orderNumber];
      bucket.fromRange = true;
    });
  }

  function resolveOrders(lineTags, hitIndex, lineIndex) {
    let orders = orderNumbersBefore(lineTags, hitIndex, activeOrderNumbers);
    if (!orders.length) orders = colourAhead(lineIndex, hitIndex);
    return orders;
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const alone = line.match(ALONE_ORDER_RE);
    if (alone) activeOrderNumbers = expandOrderRef(alone[1]).orderNumbers;

    const lineTags = colourTagsIn(line);
    if (lineTags.length) {
      activeOrderNumbers = lineTags[lineTags.length - 1].orderNumbers.slice();
    }

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
          const orders = resolveOrders(tags, hit.index, i);
          const parsed = consumePriceMatch(combo, hit);
          if (!parsed) return;
          if (orders.length) activeOrderNumbers = orders.slice();
          addParsedToOrders(parsed, orders);
        });
        i += 1;
        continue;
      }
    }

    matches.forEach((hit) => {
      const orders = resolveOrders(lineTags, hit.index, i);
      const parsed = consumePriceMatch(line, hit);
      if (!parsed) return;
      if (orders.length) activeOrderNumbers = orders.slice();
      addParsedToOrders(parsed, orders);
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
  const tokenSeen = {};
  const orderTokens = [];
  extractOrderTokens(raw).forEach((id) => {
    if (!tokenSeen[id]) {
      tokenSeen[id] = true;
      orderTokens.push(id);
    }
  });
  list.forEach((row) => {
    if (row.orderNumber && !tokenSeen[row.orderNumber]) {
      tokenSeen[row.orderNumber] = true;
      orderTokens.push(row.orderNumber);
    }
  });

  return {
    ok: true,
    orderTokens,
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
  expandOrderRef,
  lettersFromSpec,
  parseRandToken,
  extractOrderTokens,
  relaxInvoiceText,
  splitCostShares,
  parsePaintInvoiceText,
  extractPdfText,
  parsePaintInvoiceBuffer
};
