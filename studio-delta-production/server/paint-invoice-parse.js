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

/**
 * Paint-shop SKUs like SD-CABINET-006A / 006B / 006C are components of one
 * furniture item (not Studio Delta letter-split orders).
 */
function skuFamilyFromText(text) {
  const s = String(text || "").toUpperCase();
  const m = s.match(/\b(SD-[A-Z]+-\d+)([A-Z])\b/);
  if (m) return { family: m[1], part: m[2], sku: m[1] + m[2] };
  const m2 = s.match(/\b(SD-[A-Z]+-\d+)\b/);
  if (m2) return { family: m2[1], part: "", sku: m2[1] };
  return null;
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

  const PRICE_GLOBAL_RE = new RegExp(QTY_PRICE_TAIL_RE.source, "gi");

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

  function lastSkuIn(text) {
    let last = null;
    const re = /\bSD-[A-Z]+-\d+[A-Z]?\b/g;
    let m;
    while ((m = re.exec(String(text || "")))) {
      last = Object.assign({ index: m.index }, skuFamilyFromText(m[0]) || {});
    }
    return last && last.family ? last : null;
  }

  function skuNear(lineIndex, sourceText, atIndex) {
    const prefix = String(sourceText || "").slice(0, atIndex == null ? undefined : atIndex);
    const fromPrefix = lastSkuIn(prefix);
    if (fromPrefix) return Object.assign({ lineIndex: lineIndex }, fromPrefix);
    const onLine = lastSkuIn(lines[lineIndex]);
    if (onLine) return Object.assign({ lineIndex: lineIndex }, onLine);
    const prev = lastSkuIn(lines[lineIndex - 1]);
    if (prev) return Object.assign({ lineIndex: lineIndex - 1 }, prev);
    const prev2 = lastSkuIn(lines[lineIndex - 2]);
    if (prev2) return Object.assign({ lineIndex: lineIndex - 2 }, prev2);
    return null;
  }

  function posOf(lineIndex, charIndex) {
    return (lineIndex * 100000) + (charIndex || 0);
  }

  const orderMarks = [];
  const pricedRecords = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    const tokenRe = new RegExp(ORDER_TOKEN_RE.source, "gi");
    let tokenHit;
    while ((tokenHit = tokenRe.exec(line))) {
      const expanded = expandOrderRef(tokenHit[0]);
      if (expanded.orderNumbers.length) {
        activeOrderNumbers = expanded.orderNumbers;
        orderMarks.push({
          lineIndex: i,
          charIndex: tokenHit.index,
          pos: posOf(i, tokenHit.index),
          orderNumbers: expanded.orderNumbers.slice()
        });
      }
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

    function takeHits(sourceText, hits, lineIndex) {
      hits.forEach((hit) => {
        const parsed = consumePriceMatch(sourceText, hit);
        if (!parsed) return;
        const sku = skuNear(lineIndex, sourceText, hit.index);
        const pricePos = posOf(lineIndex, hit.index);
        const skuPos = sku && sku.index != null
          ? posOf(sku.lineIndex != null ? sku.lineIndex : lineIndex, sku.index)
          : pricePos;
        pricedRecords.push({
          parsed,
          lineIndex,
          charIndex: hit.index,
          pos: pricePos,
          skuPos,
          skuFamily: sku && sku.family,
          sku: sku && sku.sku
        });
      });
    }

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
        takeHits(combo, comboMatches, i);
        i += 1;
        continue;
      }
    }

    takeHits(line, matches, i);
  }

  const firstPricePos = pricedRecords.length ? pricedRecords[0].pos : -1;
  const firstMarkPos = orderMarks.length ? orderMarks[0].pos : -1;
  // Real Ultimate invoices put colour AFTER the A/B/C parts. Other PDFs put
  // the colour line first. Decide from whether the first price or first
  // order tag appears earlier — not from SKU position (SKU often sits above
  // both the colour and the price).
  const colourAfterStyle = firstPricePos >= 0 && (firstMarkPos < 0 || firstPricePos < firstMarkPos);

  pricedRecords.forEach((row) => {
    let next = null;
    let prev = null;
    orderMarks.forEach((mark) => {
      if (mark.pos > row.pos && (!next || mark.pos < next.pos)) next = mark;
      if (mark.pos < row.pos && (!prev || mark.pos > prev.pos)) prev = mark;
    });
    let chosen = [];
    if (colourAfterStyle) chosen = (next && next.orderNumbers) || (prev && prev.orderNumbers) || [];
    else chosen = (prev && prev.orderNumbers) || (next && next.orderNumbers) || [];
    addParsedToOrders(row.parsed, chosen);
  });

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
  skuFamilyFromText,
  relaxInvoiceText,
  splitCostShares,
  parsePaintInvoiceText,
  extractPdfText,
  parsePaintInvoiceBuffer
};
