const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-glass-inv-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const reader = require("./glass-invoice-parse");

// Text of National Glass invoice F42678 as a PDF text layer gives it.
const F42678 = `Customer Tax Invoice: F42678 Date: 20 Nov 2026
Invoice
Sold To: Address:
VAT No:
585 Krige Street
Tannery Industrial Park
Silverton
Studio Delta(Pty)Ltd 320 The Hillside
Lynnwood
Pretoria
Shipped To:
Internal Sales:
Contact No:
Winona Nell
Acc No: 010 593 8221
SO No:
INV No:
INV Terms:
PO Ref:
S1055
692019
F42678
S26270A+B, 283A+B,
4570300618
30 Days
Delivery Truck: Pt West
External Sales: Chris Van Der Walt
Contact No: 083 654 5335
Qty Unit Reference Stock Code Description Mass(kg) Unit Price (R) GrossValue (R) Disc% Net Value (R)
 0.282 M2 S26270A+B, XTOBS4.35 4mm SLENDER (NREED) CTS 2.82 479.00 135.08 135.08
283A+B,
4 of 775 X 91 +Polish All Round (S26270A+B, 283A+B, 1)
 7.000 M POL3.15 S/Edge Mach Matt Polish 2-6mm 19.00 133.00 133.00
 0.663 M2 XTOBS4.35 + TGPR.04 4mm SLENDER (NREED) CTS TSG 6.63 588.00 389.85 389.85
2 of 780 X 425 +Polish All Round +Toughening (S26270A+B,
283A+B, 1)
 4.900 M POL3.15 S/Edge Mach Matt Polish 2-6mm 100.00%
 2.000 EA LOGOPOS.STD LOGO POS NG STD
 2.000 EA TGPRL.NLS TGP Standard Logo w. Size
 0.150 M2 XTOBS4.35 4mm SLENDER (NREED) CTS 1.50 479.00 71.85 71.85
4 of 407 X 92 +Polish All Round (S26270A+B, 283A+B, 1)
 4.200 M POL3.15 S/Edge Mach Matt Polish 2-6mm 19.00 79.80 79.80
 0.860 M2 XTOBS4.35 + TGPR.04 4mm SLENDER (NREED) CTS TSG 8.60 588.00 505.68 505.68
2 of 735 X 585 +Polish All Round +Toughening (S26270A+B,
283A+B, 1)
 5.400 M POL3.15 S/Edge Mach Matt Polish 2-6mm 100.00%
 2.000 EA LOGOPOS.STD LOGO POS NG STD
 2.000 EA TGPRL.NLS TGP Standard Logo w. Size
 2.015 M2 4mm XTCFG04.01 + TGPR.04 4mm CLEAR FLOAT CTS TSG 20.15 323.00 650.85 650.85
4 of 1 550 X 325 +Polish All Round +Toughening
(4mm)
 15.000 M POL3.15 S/Edge Mach Matt Polish 2-6mm 100.00%
 4.000 EA LOGOPOS.STD LOGO POS NG STD
 4.000 EA TGPRL.NLS TGP Standard Logo w. Size
 0.459 M2 XTCFG04.01 + TGPR.04 4mm CLEAR FLOAT CTS TSG 4.59 323.00 148.26 148.26
1 of 1 610 X 285 +Polish All Round +Toughening (.)
 3.850 M POL3.15 S/Edge Mach Matt Polish 2-6mm 100.00%
 1.000 EA LOGOPOS.STD LOGO POS NG STD
 1.000 EA TGPRL.NLS TGP Standard Logo w. Size
 0.293 M2 6mm XTCFG06.01 6mm CLEAR FLOAT CTS 4.40 305.00 89.37 89.37
8 of 398 X 92 +Polish All Round (6mm)
 8.000 M POL3.15 S/Edge Mach Matt Polish 2-6mm 19.00 152.00 152.00
 0.580 M2 XTCFG06.01 + TGPR.06 6mm CLEAR FLOAT CTS TSG 8.70 436.00 252.88 252.88
3 of 730 X 265 +Polish All Round +Toughening (6mm)
 6.150 M POL3.15 S/Edge Mach Matt Polish 2-6mm 100.00%
Page 1 of 2
4030148680 +27(0)10 593 8400 orders@natglass.co.za portal.natglass.co.za
11 Nov 2026 11:21 Reprinted: 20 Nov 2026 10:15
164 Russell Road, Gqeberha
Customer Tax Invoice: F42678 Date: 20 Nov 2026
Invoice
Qty Unit Reference Stock Code Description Mass(kg) Unit Price (R) GrossValue (R) Disc% Net Value (R)
 3.000 EA LOGOPOS.STD LOGO POS NG STD
 3.000 EA TGPRL.NLS TGP Standard Logo w. Size
 0.578 M2 XTCFG06.01 + TGPR.06 6mm CLEAR FLOAT CTS TSG 8.67 436.00 252.01 252.01
3 of 727 X 265 +Polish All Round +Toughening (6mm)
 6.150 M POL3.15 S/Edge Mach Matt Polish 2-6mm 100.00%
 3.000 EA LOGOPOS.STD LOGO POS NG STD
 3.000 EA TGPRL.NLS TGP Standard Logo w. Size
.
Many Thanks,
Winona Nell
Note: Sub Total
Plus Transport Levy
Net Amount Before VAT
VAT
Total ZAR
R2 860.63
R0.00
R2 860.63
R429.10
R3 289.73
Mass (kg): 66,06
Page 2 of 2`;

// Small helpers
assert.deepStrictEqual(reader.parseOrderRefs("S26270A+B, 283A+B, 1"), ["S260270 A", "S260270 B", "S260283 A", "S260283 B"]);
assert.deepStrictEqual(reader.parseOrderRefs("(4mm)"), []);
assert.strictEqual(reader.glassTypeKey("4mm SLENDER (NREED) CTS"), "reeded");
assert.strictEqual(reader.glassTypeKey("Reeded"), "reeded");
assert.strictEqual(reader.glassTypeKey("6mm CLEAR FLOAT CTS"), "clear");
assert.strictEqual(reader.glassTypeKey("Ocean view"), "ocean view");
assert.strictEqual(reader.thicknessOf("4mm"), 4);
assert.strictEqual(reader.thicknessOf("6"), 6);

// Reading the invoice: 9 pieces, each glass line plus its polish.
const parsed = reader.parseGlassInvoiceText(F42678);
assert.ok(parsed.ok, parsed.error);
assert.strictEqual(parsed.invoiceNo, "F42678");
assert.strictEqual(parsed.supplier, "National Glass");
assert.strictEqual(parsed.pieces.length, 9);
const brief = parsed.pieces.map((p) => [p.quantity, p.height, p.width, p.thickness, p.type, p.toughened, p.cost]);
assert.deepStrictEqual(brief, [
  [4, 775, 91, 4, "reeded", false, 268.08],
  [2, 780, 425, 4, "reeded", true, 389.85],
  [4, 407, 92, 4, "reeded", false, 151.65],
  [2, 735, 585, 4, "reeded", true, 505.68],
  [4, 1550, 325, 4, "clear", true, 650.85],
  [1, 1610, 285, 4, "clear", true, 148.26],
  [8, 398, 92, 6, "clear", false, 241.37],
  [3, 730, 265, 6, "clear", true, 252.88],
  [3, 727, 265, 6, "clear", true, 252.01]
]);
assert.strictEqual(parsed.pieces[0].glassCost, 135.08);
assert.strictEqual(parsed.pieces[0].polishCost, 133);
assert.strictEqual(parsed.pieces[1].polishCost, 0, "toughened pieces include polish (100% discount)");
assert.deepStrictEqual(parsed.pieces[0].refs, ["S260270 A", "S260270 B", "S260283 A", "S260283 B"]);
assert.deepStrictEqual(parsed.pieces[4].refs, [], "clear lines carry no order reference");
assert.strictEqual(parsed.piecesTotal, 2860.63);
assert.strictEqual(parsed.invoiceSubtotal, 2860.63);
assert.strictEqual(parsed.invoiceVat, 429.1);
assert.strictEqual(parsed.invoiceTotal, 3289.73);

// Glued text (no spaces between PDF columns) still reads the same money.
const glued = reader.parseGlassInvoiceText(
  "0.282M2S26270A+B,XTOBS4.354mm SLENDER (NREED) CTS2.82479.00135.08135.08 4 of 775 X 91 +Polish (S26270A+B, 1) 7.000MPOL3.15 S/Edge Mach Matt Polish 2-6mm19.00133.00133.00"
);
assert.ok(glued.ok, glued.error);
assert.deepStrictEqual(
  [glued.pieces[0].thickness, glued.pieces[0].type, glued.pieces[0].glassCost, glued.pieces[0].polishCost],
  [4, "reeded", 135.08, 133]
);

// Money-first text (each row's mass/price/gross/net comes BEFORE its "0.282 M2"), as the live
// server's PDF reader produced. Each piece must keep its own price, not the next row's.
const moneyFirst = reader.parseGlassInvoiceText([
  "Qty Unit Reference Stock Code Description Mass(kg) Unit Price (R) GrossValue (R) Disc% Net Value (R)",
  "2.82 479.00 135.08 135.08 0.282 M2 S26270A+B, XTOBS4.35 4mm SLENDER (NREED) CTS 283A+B,",
  "4 of 775 X 91 +Polish All Round (S26270A+B, 283A+B, 1)",
  "19.00 133.00 133.00 7.000 M POL3.15 S/Edge Mach Matt Polish 2-6mm",
  "6.63 588.00 389.85 389.85 0.663 M2 XTOBS4.35 + TGPR.04 4mm SLENDER (NREED) CTS TSG",
  "2 of 780 X 425 +Polish All Round +Toughening (S26270A+B, 283A+B, 1)",
  "100.00% 4.900 M POL3.15 S/Edge Mach Matt Polish 2-6mm",
  "2.000 EA LOGOPOS.STD LOGO POS NG STD",
  "4.40 305.00 89.37 89.37 0.293 M2 6mm XTCFG06.01 6mm CLEAR FLOAT CTS",
  "8 of 398 X 92 +Polish All Round (6mm)",
  "19.00 152.00 152.00 8.000 M POL3.15 S/Edge Mach Matt Polish 2-6mm",
  "8.70 436.00 252.88 252.88 0.580 M2 XTCFG06.01 + TGPR.06 6mm CLEAR FLOAT CTS TSG",
  "3 of 730 X 265 +Polish All Round +Toughening (6mm)",
  "100.00% 6.150 M POL3.15 S/Edge Mach Matt Polish 2-6mm"
].join("\n"));
assert.ok(moneyFirst.ok, moneyFirst.error);
assert.deepStrictEqual(
  moneyFirst.pieces.map((p) => [p.quantity, p.height, p.width, p.thickness, p.type, p.toughened, p.glassCost, p.polishCost, p.cost, p.costCheck]),
  [
    [4, 775, 91, 4, "reeded", false, 135.08, 133, 268.08, false],
    [2, 780, 425, 4, "reeded", true, 389.85, 0, 389.85, false],
    [8, 398, 92, 6, "clear", false, 89.37, 152, 241.37, false],
    [3, 730, 265, 6, "clear", true, 252.88, 0, 252.88, false]
  ]
);
// The printed layout reads the same, and no piece needs a cost check.
assert.ok(parsed.pieces.every((p) => !p.costCheck));

// A polish row with no money and no 100% discount is flagged, never filled from a neighbour.
const missingPolish = reader.parseGlassInvoiceText(
  "0.282 M2 XTOBS4.35 4mm SLENDER (NREED) CTS 2.82 479.00 135.08 135.08 4 of 775 X 91 +Polish 7.000 M POL3.15 S/Edge Mach Matt Polish 2-6mm 0.663 M2 XTOBS4.35 + TGPR.04 4mm SLENDER (NREED) CTS TSG 6.63 588.00 389.85 389.85 2 of 780 X 425 +Toughening"
);
assert.strictEqual(missingPolish.pieces[0].cost, 135.08);
assert.strictEqual(missingPolish.pieces[0].costCheck, true);

// Matching against open PO lines: order numbers mostly NOT on the invoice.
const L = (id, order, type, thickness, height, width, quantity, extra) =>
  Object.assign({ id, order, type, thickness, height, width, quantity, isTemplate: false, poNumber: "GPO-0007" }, extra || {});
const open = [
  L("g1", "S260270 A", "Reeded", "4mm", 775, 91, 2),
  L("g2", "S260270 A", "Reeded", "4mm", 780, 425, 1),
  L("g3", "S260270 B", "Reeded", "4mm", 775, 91, 2),
  L("g4", "S260270 B", "Reeded", "4mm", 780, 425, 1),
  L("g5", "S260283 A", "Reeded", "4mm", 92, 407, 2),
  L("g6", "S260283 A", "Reeded", "4mm", 735, 585, 1),
  L("g7", "S260283 B", "Reeded", "4mm", 407, 92, 2),
  L("g8", "S260283 B", "Reeded", "4mm", 735, 585, 1),
  L("g9", "S260291", "Clear", "4mm", 1550, 325, 4),
  L("g10", "S260295", "Clear", "4mm", 1612, 285, 1),
  L("g11", "S260301", "Clear", "6mm", 398, 92, 8),
  L("g12", "S260305 A", "Clear", "6mm", 730, 265, 3),
  L("g13", "S260305 B", "Clear", "6mm", 727, 265, 3),
  L("g14", "S260310", "Ocean view", "4mm", 600, 400, 2),
  L("g15", "S260320", "Clear", "4mm", 0, 0, 1, { isTemplate: true, templateSpec: "Arch" })
];
const matched = reader.matchPiecesToLines(parsed.pieces, open);
assert.ok(matched.results.every((r) => r.status === "matched"), JSON.stringify(matched.results.filter((r) => r.status !== "matched")));
const cost = matched.perLine;
const near = (a, b) => Math.abs(a - b) < 0.011;
assert.ok(near(cost.g1, 134.04) && near(cost.g3, 134.04), "4 × 775×91 split 2 + 2 across S260270 A and B");
assert.ok(near(cost.g2 + cost.g4, 389.85) && near(cost.g2, 194.925) && near(cost.g4, 194.925));
assert.ok(near(cost.g5 + cost.g7, 151.65), "92×407 matches 407×92 the other way round");
assert.ok(near(cost.g6, 252.84) && near(cost.g8, 252.84));
assert.strictEqual(cost.g9, 650.85);
assert.strictEqual(cost.g10, 148.26, "1612 ordered vs 1610 invoiced is within 2 mm");
assert.ok(matched.results[5].sizeOff, "the 2 mm difference is flagged");
assert.strictEqual(cost.g11, 241.37);
assert.strictEqual(cost.g12, 252.88, "730 goes to S260305 A");
assert.strictEqual(cost.g13, 252.01, "727 goes to S260305 B, 3 mm is too far for A");
assert.strictEqual(matched.matchedTotal, 2860.63);
assert.strictEqual(matched.unmatchedTotal, 0);
assert.deepStrictEqual(matched.notOnInvoice.map((l) => l.order), ["S260310"], "Ocean view is still outstanding");
assert.deepStrictEqual(matched.templateLines.map((l) => l.order), ["S260320"]);
assert.ok(!("g14" in cost) && !("g15" in cost));

// Quantities that do not add up are not forced onto an order.
const short = reader.matchPiecesToLines(parsed.pieces.slice(0, 1), [L("x1", "S260270 A", "Reeded", "4mm", 775, 91, 3)]);
assert.strictEqual(short.results[0].status, "unmatched");
assert.ok(/do not add up/.test(short.results[0].reason));

// Wrong thickness or type never matches.
const wrong = reader.matchPiecesToLines(parsed.pieces.slice(6, 7), [
  L("w1", "S260301", "Clear", "4mm", 398, 92, 8),
  L("w2", "S260302", "Reeded", "6mm", 398, 92, 8)
]);
assert.strictEqual(wrong.results[0].status, "unmatched");

// The order reference breaks a tie between two orders with the same glass.
const tie = reader.matchPiecesToLines(parsed.pieces.slice(1, 2), [
  L("t1", "S260999 A", "Reeded", "4mm", 780, 425, 2),
  L("t2", "S260270 A", "Reeded", "4mm", 780, 425, 1),
  L("t3", "S260270 B", "Reeded", "4mm", 780, 425, 1)
]);
assert.deepStrictEqual(tie.results[0].lines.map((l) => l.id).sort(), ["t2", "t3"]);

// No text at all.
assert.strictEqual(reader.parseGlassInvoiceText("").ok, false);

console.log("glass-invoice-parse tests ok");
