"use strict";

/**
 * Read a glass supplier invoice (National Glass layout) and match each piece of glass to the
 * open glass lines on Studio Delta purchase orders.
 *
 * The invoice often has no order number on a line, so matching uses what the glass IS:
 * thickness + glass type + height × width (either way round, within ±2 mm) + quantity.
 * An order reference on the line (e.g. "S26270A+B, 283A+B") is only a tie-breaker.
 *
 * One piece = the m² glass line + the polish (m) lines under it. Logo and toughening
 * lines carry no money. Toughened pieces have their polish discounted 100%.
 */

const { formatOrderId, money, formatRand } = require("./db");

const SIZE_TOLERANCE_MM = 2;
// Area (3 decimals) followed by the unit: "0.282 M2", "7.000 M POL…", "2.000 EA".
const QTY_UNIT_RE = /(\d+[.,]\d{3})\s*(M2|EA|M)(?![a-z0-9])/g;
const PIECE_RE = /(\d{1,3})\s*of\s*(\d{1,2}\s\d{3}|\d{2,4})\s*[xX×]\s*(\d{1,2}\s\d{3}|\d{2,4})/;
// Supplier stock codes such as XTOBS4.35, XTCFG04.01, TGPR.04, POL3.15 — never money.
const STOCK_CODE_RE = /\b[A-Z]{3,}\d*\.\d{2}/g;

function num(raw) {
  const n = Number(String(raw == null ? "" : raw).replace(/\s/g, "").replace(",", "."));
  return Number.isFinite(n) ? n : NaN;
}

function round2(n) {
  return Number(money(n));
}

function flatText(text) {
  return String(text || "").replace(/\r/g, "\n").replace(/[ \t\f\v]+/g, " ").replace(/\s*\n\s*/g, " ").trim();
}

// Money in a line: try spaced amounts (with "1 234.56" thousands) and glued ones ("2.82479.00").
function moneyLists(text) {
  const clean = String(text || "")
    .replace(STOCK_CODE_RE, " ")
    .replace(/\d+[.,]\d{2}\s*%/g, " ");
  const spaced = [];
  clean.replace(/(?<![\d.])\d{1,3}(?:\s\d{3})*\.\d{2}(?![\d])/g, (m) => {
    spaced.push(num(m));
    return m;
  });
  const glued = [];
  clean.replace(/\d+\.\d{2}/g, (m) => {
    glued.push(num(m));
    return m;
  });
  return [spaced, glued];
}

// Find unit price and gross where measure × unit ≈ gross; net is the amount right after gross.
// Only a row's own money passes this check, so a neighbouring row's price is never taken.
function pricedFromAmounts(measure, text) {
  const lists = moneyLists(text);
  for (const list of lists) {
    for (let i = 0; i + 1 < list.length; i++) {
      const unit = list[i];
      const gross = list[i + 1];
      if (!(unit > 0) || !(gross >= 0)) continue;
      const want = measure * unit;
      if (Math.abs(want - gross) <= Math.max(0.06, gross * 0.002)) {
        const next = list[i + 2];
        const net = next != null && next <= gross + 0.01 ? next : gross;
        return { unit, gross, net: round2(net) };
      }
    }
  }
  return null;
}

// PDF text layers put a row's money either after its "0.282 M2" (as printed) or before it
// (money columns first). The invoice's layout is decided once, then every row is read that way.
function rowPrice(measure, afterText, beforeText, layout) {
  if (layout === "before") return pricedFromAmounts(measure, beforeText);
  if (layout === "after") return pricedFromAmounts(measure, afterText);
  return pricedFromAmounts(measure, afterText) || pricedFromAmounts(measure, beforeText);
}

function glassTypeKey(raw) {
  const s = String(raw || "").toUpperCase();
  if (!s.trim()) return "";
  if (/REED|SLENDER|FLUT/.test(s)) return "reeded";
  if (/OCEAN/.test(s)) return "ocean view";
  if (/MIRROR/.test(s)) return "mirror";
  if (/BRONZE/.test(s)) return "bronze";
  if (/GREY|GRAY/.test(s)) return "grey";
  if (/SATIN|ETCH|FROST|OBSCUR/.test(s)) return "satin";
  if (/CLEAR|FLOAT/.test(s)) return "clear";
  return "";
}

function thicknessOf(raw) {
  const m = String(raw == null ? "" : raw).match(/(\d{1,2})(?:[.,]\d+)?\s*(?:mm)?/i);
  return m ? Number(m[1]) : 0;
}

function thicknessFromInvoice(glassText) {
  const desc = String(glassText || "").match(/(?<![\d.])(\d{1,2})\s?mm\s*[A-Za-z(]/);
  if (desc) return Number(desc[1]);
  const code = String(glassText || "").match(/\bXT[A-Z]{2,4}0?(\d{1,2})\.\d{2}/);
  return code ? Number(code[1]) : 0;
}

/** "S26270A+B, 283A+B" → S260270 A, S260270 B, S260283 A, S260283 B. */
function parseOrderRefs(text) {
  const out = [];
  let prefix = "";
  String(text || "").split(/[,;()]+/).forEach((chunk) => {
    const c = chunk.replace(/\s+/g, "").toUpperCase();
    const m = c.match(/^(S)?(\d{3,6})([A-Z](?:\+[A-Z])*)?$/);
    if (!m) return;
    let digits = m[2];
    if (m[1]) {
      if (digits.length === 5) digits = digits.slice(0, 2) + "0" + digits.slice(2);
      if (digits.length !== 6) return;
      prefix = digits.slice(0, 3);
    } else if (digits.length === 3 && prefix) {
      digits = prefix + digits;
    } else if (digits.length !== 6) {
      return;
    }
    const letters = m[3] ? m[3].split("+") : [""];
    letters.forEach((L) => {
      const id = formatOrderId("S" + digits + (L ? " " + L : ""));
      if (id && out.indexOf(id) === -1) out.push(id);
    });
  });
  return out;
}

function parseInvoiceTotals(flat) {
  const amounts = [];
  flat.replace(/R\s?(\d{1,3}(?:[ ,]\d{3})*\.\d{2})/g, (_m, a) => {
    const n = num(a.replace(/,/g, ""));
    if (Number.isFinite(n)) amounts.push(n);
    return _m;
  });
  if (!amounts.length) return { subtotal: null, vat: null, total: null };
  const total = Math.max.apply(null, amounts);
  const vat = amounts.find((a) => a > 0 && Math.abs(a - total * 15 / 115) <= 0.05);
  if (vat != null) return { subtotal: round2(total - vat), vat: round2(vat), total: round2(total) };
  return { subtotal: null, vat: null, total: round2(total) };
}

/** Turn invoice text into pieces of glass with their cost (excl VAT). */
function parseGlassInvoiceText(text) {
  const flat = flatText(text);
  if (!flat) {
    return { ok: false, error: "No text found in that PDF. Upload the supplier's PDF invoice, not a photo or scan.", pieces: [] };
  }
  const events = [];
  let m;
  QTY_UNIT_RE.lastIndex = 0;
  while ((m = QTY_UNIT_RE.exec(flat))) {
    events.push({ index: m.index, end: m.index + m[0].length, measure: num(m[1]), unit: m[2] });
  }
  const glassAt = events.filter((e) => e.unit === "M2");
  const pieces = [];
  const warnings = [];
  // Text between the previous qty/unit token and this one (where money sits when a PDF
  // writes each row's numbers before its quantity).
  function textBefore(ev) {
    const prev = events.filter((e) => e.end <= ev.index).pop();
    const from = Math.max(prev ? prev.end : 0, ev.index - 260);
    return flat.slice(from, ev.index).replace(/\d\s*-\s*\d+\s*mm/gi, " ");
  }
  const DISCOUNT_RE = /100(?:[.,]0+)?\s*%/;
  function textAfter(ev) {
    const next = events.find((e) => e.index >= ev.end);
    return flat.slice(ev.end, Math.min(next ? next.index : flat.length, ev.end + 400)).replace(/\d\s*-\s*\d+\s*mm/gi, " ");
  }
  let afterHits = 0;
  let beforeHits = 0;
  glassAt.forEach((ev) => {
    if (pricedFromAmounts(ev.measure, textAfter(ev))) afterHits += 1;
    if (pricedFromAmounts(ev.measure, textBefore(ev))) beforeHits += 1;
  });
  const layout = beforeHits > afterHits ? "before" : (afterHits > 0 ? "after" : "");
  glassAt.forEach((ev, gi) => {
    const segEnd = gi + 1 < glassAt.length ? glassAt[gi + 1].index : flat.length;
    const seg = flat.slice(ev.end, segEnd);
    const pieceMatch = seg.match(PIECE_RE);
    const firstOther = events.find((e) => e.index >= ev.end && e.index < segEnd);
    let glassEnd = pieceMatch ? pieceMatch.index : seg.length;
    if (firstOther && firstOther.index - ev.end < glassEnd) glassEnd = firstOther.index - ev.end;
    const glassText = seg.slice(0, Math.min(glassEnd, 260));
    const priced = rowPrice(ev.measure, glassText, textBefore(ev), layout);
    let polish = 0;
    let costCheck = !priced;
    events
      .filter((e) => e.unit === "M" && e.index > ev.index && e.index < segEnd)
      .forEach((pe) => {
        const next = events.find((e) => e.index > pe.index);
        const stop = Math.min(next ? next.index : segEnd, pe.end + 160, segEnd);
        const after = flat.slice(pe.end, stop).replace(/\d\s*-\s*\d+\s*mm/gi, " ");
        const before = textBefore(pe);
        const p = rowPrice(pe.measure, after, before, layout);
        if (p) {
          polish += p.net;
          return;
        }
        // No priced polish row: fine when it is discounted 100% (toughened), otherwise check it.
        if (!DISCOUNT_RE.test(after) && !DISCOUNT_RE.test(before)) costCheck = true;
      });
    const refText = glassText.split(STOCK_CODE_RE)[0] + " " +
      (pieceMatch ? ((seg.slice(pieceMatch.index, pieceMatch.index + 120).match(/\(([^)]*)\)/) || [])[1] || "") : "");
    if (!pieceMatch) {
      warnings.push("A glass line of " + ev.measure + " m² has no 'N of H X W' size and was skipped.");
      return;
    }
    const qty = Number(pieceMatch[1]);
    const height = num(pieceMatch[2]);
    const width = num(pieceMatch[3]);
    const glassCost = priced ? priced.net : 0;
    pieces.push({
      index: pieces.length + 1,
      quantity: qty,
      height,
      width,
      thickness: thicknessFromInvoice(glassText),
      type: glassTypeKey(glassText),
      typeText: (glassText.match(/\d{1,2}\s?mm\s*([A-Za-z][A-Za-z ()]+?)(?=\s*\d|$)/) || [])[1] || "",
      toughened: /TSG|TGPR|TOUGHEN/i.test(seg.slice(0, (pieceMatch.index || 0) + 120)),
      areaM2: ev.measure,
      glassCost: round2(glassCost),
      polishCost: round2(polish),
      cost: round2(glassCost + polish),
      costCheck,
      refs: parseOrderRefs(refText)
    });
  });
  if (!pieces.length) {
    return {
      ok: false,
      error: "Could not find any glass lines (like '4 of 775 X 91') on that invoice. Check it is the supplier's PDF.",
      pieces: [],
      textPreview: flat.slice(0, 1200)
    };
  }
  const totals = parseInvoiceTotals(flat);
  const invoiceNo = (flat.match(/Tax Invoice:?\s*([A-Z]{0,3}\d{3,})/i) || [])[1] || "";
  const piecesTotal = round2(pieces.reduce((s, p) => s + p.cost, 0));
  return {
    ok: true,
    invoiceNo,
    supplier: /national\s*glass|natglass/i.test(flat) ? "National Glass" : "",
    pieces,
    piecesTotal,
    invoiceSubtotal: totals.subtotal,
    invoiceVat: totals.vat,
    invoiceTotal: totals.total,
    warnings,
    textPreview: flat.slice(0, 1200)
  };
}

function sizeMatch(piece, line) {
  const h = Number(line.height) || 0;
  const w = Number(line.width) || 0;
  if (!(h > 0) || !(w > 0)) return -1;
  const ph = piece.height;
  const pw = piece.width;
  if ((h === ph && w === pw) || (h === pw && w === ph)) return 0;
  const near = (a, b) => Math.abs(a - b) <= SIZE_TOLERANCE_MM;
  if ((near(h, ph) && near(w, pw)) || (near(h, pw) && near(w, ph))) return 1;
  return -1;
}

function typeCompatible(piece, line) {
  const a = piece.type;
  const b = glassTypeKey(line.type);
  return !a || !b || a === b;
}

function refHit(piece, line) {
  if (!piece.refs || !piece.refs.length) return false;
  const id = formatOrderId(line.order);
  const base = id.replace(/\s+[A-Z]$/i, "");
  return piece.refs.some((r) => r === id || r === base);
}

// Smallest-first subset of candidate lines whose quantities add up to the invoice quantity.
function pickLines(cands, qty) {
  const list = cands.slice(0, 14);
  let best = null;
  function walk(i, sum, chosen) {
    if (best) return;
    if (sum === qty) { best = chosen.slice(); return; }
    if (sum > qty || i >= list.length) return;
    chosen.push(list[i]);
    walk(i + 1, sum + (Number(list[i].quantity) || 0), chosen);
    chosen.pop();
    walk(i + 1, sum, chosen);
  }
  walk(0, 0, []);
  return best;
}

function splitByQuantity(total, lines) {
  const qty = lines.reduce((s, l) => s + (Number(l.quantity) || 0), 0) || lines.length;
  let left = round2(total);
  return lines.map((l, i) => {
    if (i === lines.length - 1) return left;
    const share = round2(total * (Number(l.quantity) || 1) / qty);
    left = round2(left - share);
    return share;
  });
}

/**
 * Match invoice pieces to open glass lines.
 * openLines: [{ id, order, type, thickness, height, width, quantity, isTemplate, poNumber }]
 */
function matchPiecesToLines(pieces, openLines) {
  const remaining = (openLines || []).filter((l) => l && !l.isTemplate && Number(l.quantity) > 0);
  const results = (pieces || []).map((p) => ({ piece: p, lines: [], status: "unmatched", reason: "", sizeOff: false }));

  function candidatesFor(piece, allowNear) {
    return remaining
      .filter((l) => {
        const t = thicknessOf(l.thickness);
        if (piece.thickness && t && piece.thickness !== t) return false;
        if (!typeCompatible(piece, l)) return false;
        const s = sizeMatch(piece, l);
        return s === 0 || (allowNear && s === 1);
      })
      .map((l) => ({ line: l, ref: refHit(piece, l), exact: sizeMatch(piece, l) === 0 }))
      .sort((a, b) => (b.ref - a.ref) || (b.exact - a.exact) || String(a.line.order).localeCompare(String(b.line.order)));
  }

  function tryMatch(res, allowNear) {
    if (res.status === "matched") return;
    const cands = candidatesFor(res.piece, allowNear);
    if (!cands.length) return;
    const withRef = cands.filter((c) => c.ref);
    const picked = (withRef.length && pickLines(withRef.map((c) => c.line), res.piece.quantity))
      || pickLines(cands.map((c) => c.line), res.piece.quantity);
    if (!picked) {
      res.reason = "Same glass on " + cands.map((c) => c.line.order + " (" + c.line.quantity + ")").join(", ") +
        " but the quantities do not add up to " + res.piece.quantity + ".";
      res.near = cands.map((c) => c.line.id);
      return;
    }
    const shares = splitByQuantity(res.piece.cost, picked);
    res.lines = picked.map((l, i) => ({
      id: l.id,
      order: l.order,
      poNumber: l.poNumber || "",
      quantity: Number(l.quantity) || 0,
      height: l.height,
      width: l.width,
      cost: shares[i],
      costLabel: formatRand(shares[i]),
      exact: sizeMatch(res.piece, l) === 0
    }));
    res.status = "matched";
    res.reason = "";
    res.sizeOff = res.lines.some((l) => !l.exact);
    res.byRef = picked.some((l) => refHit(res.piece, l));
    picked.forEach((l) => {
      const at = remaining.indexOf(l);
      if (at !== -1) remaining.splice(at, 1);
    });
  }

  // Exact sizes first so a ±2 mm match never steals a line another piece fits exactly.
  results.forEach((r) => tryMatch(r, false));
  results.forEach((r) => tryMatch(r, true));
  results.forEach((r) => {
    if (r.status !== "matched" && !r.reason) {
      r.reason = "No open glass line with " + (r.piece.thickness ? r.piece.thickness + "mm " : "") +
        (r.piece.type || "glass") + " at " + r.piece.height + " × " + r.piece.width + " mm.";
    }
  });

  const perLine = {};
  results.forEach((r) => r.lines.forEach((l) => { perLine[l.id] = l.cost; }));
  const matchedTotal = round2(results.filter((r) => r.status === "matched").reduce((s, r) => s + r.piece.cost, 0));
  const unmatchedTotal = round2(results.filter((r) => r.status !== "matched").reduce((s, r) => s + r.piece.cost, 0));
  return {
    results,
    perLine,
    matchedTotal,
    unmatchedTotal,
    notOnInvoice: remaining.map((l) => ({
      id: l.id,
      order: l.order,
      type: l.type,
      thickness: l.thickness,
      height: l.height,
      width: l.width,
      quantity: l.quantity,
      poNumber: l.poNumber || ""
    })),
    templateLines: (openLines || []).filter((l) => l && l.isTemplate).map((l) => ({
      id: l.id,
      order: l.order,
      templateSpec: l.templateSpec || "",
      poNumber: l.poNumber || ""
    }))
  };
}

module.exports = {
  SIZE_TOLERANCE_MM,
  glassTypeKey,
  thicknessOf,
  parseOrderRefs,
  parseGlassInvoiceText,
  matchPiecesToLines,
  splitByQuantity
};
