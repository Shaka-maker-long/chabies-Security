"use strict";

const assert = require("assert");
const {
  parsePaintInvoiceText,
  normalizeOrderToken,
  parseRandToken,
  expandOrderRef
} = require("./paint-invoice-parse");

assert.strictEqual(normalizeOrderToken("S260244"), "S260244");
assert.strictEqual(normalizeOrderToken("S260249(A-L)"), "S260249");
assert.strictEqual(normalizeOrderToken("S260249 A"), "S260249 A");
assert.strictEqual(normalizeOrderToken("S260214A"), "S260214 A");
assert.deepStrictEqual(expandOrderRef("S260214A-F/G BLACK").orderNumbers, ["S260214 A"]);
assert.deepStrictEqual(expandOrderRef("S260214B-F/G BLACK").orderNumbers, ["S260214 B"]);
assert.deepStrictEqual(expandOrderRef("S260131-F/G BLACK").orderNumbers, ["S260131"]);
assert.strictEqual(parseRandToken("R2,219.80"), 2219.8);
assert.strictEqual(parseRandToken("R74.40"), 74.4);

const range = expandOrderRef("S260249(A-L)");
assert.strictEqual(range.base, "S260249");
assert.strictEqual(range.letters.length, 12);
assert.strictEqual(range.orderNumbers[0], "S260249 A");
assert.strictEqual(range.orderNumbers[11], "S260249 L");
assert.deepStrictEqual(expandOrderRef("S260249(A)").orderNumbers, ["S260249 A"]);

const cabinetInvoice = `
Invoice
Order S260301
SD-CABINET-112A - NAOMI ARCHED CABINET (2000X1400X680) 1 R2,219.80 0,00% 15,00% R2,219.80 R2,552.77
SD-CABINET-112B - NAOMI ARCHED CABINET DOOR (840X330) 4 R74.40 0,00% 15,00% R297.60 R342.24
SD-CABINET-112C - NAOMI ARCHED CABINET DOOR (730X660) 2 R130.20 0,00% 15,00% R260.40 R299.46
`;

const cabinet = parsePaintInvoiceText(cabinetInvoice, { preferOrderNumbers: ["S260301"] });
assert.ok(cabinet.ok);
assert.strictEqual(cabinet.allocations.length, 1);
assert.strictEqual(cabinet.allocations[0].orderNumber, "S260301");
assert.strictEqual(cabinet.allocations[0].cost, 2777.8); // 2219.80 + 297.60 + 260.40
assert.strictEqual(cabinet.allocations[0].lines.length, 3);

const colourInvoice = `
Paint shop invoice
SD-SIDEBOARD-001B - THREE DOOR SIDEBOARD DOOR (680X400)
S260244-MATT BLACK SMOOTH
3 R30.90 0,00% 15,00% R92.70 R106.61
SD-CABINET-115B - BATHROOM CABINET DOOR (1000X400)
S260235-RED BEIGE
1 R68.77 0,00% 15,00% R68.77 R79.09
`;

const coloured = parsePaintInvoiceText(colourInvoice, {
  preferOrderNumbers: ["S260244", "S260235"]
});
assert.ok(coloured.ok);
const byOrder = {};
coloured.allocations.forEach((row) => { byOrder[row.orderNumber] = row; });
assert.ok(byOrder.S260244);
assert.ok(byOrder.S260235);
assert.strictEqual(byOrder.S260244.cost, 92.7);
assert.strictEqual(byOrder.S260235.cost, 68.77);

// Real paint invoices often extract without spaces between columns.
const mashed = `
S260131-F/G BLACK
SD-CABINET-110A - CHARLENE ARCHED CABINET (2000X1000X450)1R877.500,00%15,00%R877.50R1,009.13
SD-CABINET-110B - CHARLENE ARCHED CABINET DOOR (450X350)2R93.900,00%15,00%R187.80R215.97
S260193-F/G BLACK
SD-SHELF-001 - TALITHA BOOKSHELF (1800X800X300)1R268.900,00%15,00%R268.90R309.24
`;
const mashedParsed = parsePaintInvoiceText(mashed, {
  preferOrderNumbers: ["S260131", "S260193"]
});
assert.ok(mashedParsed.ok);
const mashedBy = {};
mashedParsed.allocations.forEach((row) => { mashedBy[row.orderNumber] = row; });
assert.ok(mashedBy.S260131, "mashed PDF must total S260131");
assert.ok(mashedBy.S260193, "mashed PDF must total S260193");
assert.strictEqual(mashedBy.S260131.cost, 1065.3); // 877.50 + 187.80
assert.strictEqual(mashedBy.S260193.cost, 268.9);
assert.ok(mashedBy.S260131.costLabel);
assert.strictEqual(mashedParsed.grandTotal, 1334.2);

// One long glued extract (common for some PDF writers) must still split per order.
const glued = "S260131-F/G BLACK SD-CABINET-110A - CHARLENE ARCHED CABINET (2000X1000X450)1R877.500,00%15,00%R877.50R1,009.13 SD-CABINET-110B - DOOR (450X350)1R93.900,00%15,00%R93.90R107.99 S260193-F/G BLACK SD-SHELF-001 - TALITHA BOOKSHELF (1800X800X300)1R268.900,00%15,00%R268.90R309.24";
const gluedParsed = parsePaintInvoiceText(glued, {
  preferOrderNumbers: ["S260131", "S260193"]
});
assert.ok(gluedParsed.ok);
const gluedBy = {};
gluedParsed.allocations.forEach((row) => { gluedBy[row.orderNumber] = row; });
assert.ok(gluedBy.S260131, "glued PDF must total S260131");
assert.ok(gluedBy.S260193, "glued PDF must total S260193");
assert.strictEqual(gluedBy.S260131.cost, 971.4); // 877.50 + 93.90
assert.strictEqual(gluedBy.S260193.cost, 268.9);
assert.strictEqual(gluedParsed.grandTotal, 1240.3);

// Invoice S260249(A-L) must split excl total across shop orders S260249 A … L.
const letterPrefer = [];
for (let c = 65; c <= 76; c++) letterPrefer.push("S260249 " + String.fromCharCode(c));
const chairInvoice = `
SD-CHAIR-002 - CUSTOM STEEL CHAIR
(1020X470X430)

12 R232.50 0.00% 15.00% R2,790.00 R3,208.50

S260249(A-L)-ULTRA MARINE BLUE
`;
const chairs = parsePaintInvoiceText(chairInvoice, { preferOrderNumbers: letterPrefer });
assert.ok(chairs.ok);
assert.strictEqual(chairs.allocations.length, 12, "A–L must become 12 letter orders");
assert.strictEqual(chairs.grandTotal, 2790);
const chairBy = {};
chairs.allocations.forEach((row) => { chairBy[row.orderNumber] = row; });
assert.ok(chairBy["S260249 A"]);
assert.ok(chairBy["S260249 L"]);
assert.strictEqual(chairBy["S260249 A"].cost, 232.5);
assert.strictEqual(chairBy["S260249 L"].cost, 232.5);
assert.ok(chairBy["S260249 A"].matched);
assert.ok(chairs.orderTokens.indexOf("S260249 A") !== -1);
assert.ok(chairs.orderTokens.indexOf("S260249 L") !== -1);

// Colour tag before the price line must also expand.
const chairBefore = parsePaintInvoiceText(`
S260249(A-L)-ULTRA MARINE BLUE
SD-CHAIR-002 - CUSTOM STEEL CHAIR (1020X470X430)
12 R232.50 0.00% 15.00% R2,790.00 R3,208.50
`, { preferOrderNumbers: letterPrefer });
assert.strictEqual(chairBefore.allocations.length, 12);
assert.strictEqual(chairBefore.grandTotal, 2790);

// 006A / 006B / 006C are one product. Order number may sit only after the last part.
const components = `
SD-CABINET-006A - CHARLENE DISPLAY
CABINET (2220X980X350)
1 R877.50 0,00% 15,00% R877.50 R1,009.13
SD-CABINET-006B - CHARLENE DISPLAY
CABINET DOOR (1820X400)
1 R93.90 0,00% 15,00% R93.90 R107.99
SD-CABINET-006C - ARCHED DISPLAY
CABINET PLATE (980X350)
1 R133.70 0,00% 15,00% R133.70 R153.76
S260131-F/G BLACK
SD-BOOKSHELF-002 - TALITHA BOOKSHELF (1470X260X300)
1 R268.90 0,00% 15,00% R268.90 R309.24
S260193-F/G BLACK
`;
const parts = parsePaintInvoiceText(components, {
  preferOrderNumbers: ["S260131", "S260193"]
});
assert.ok(parts.ok);
assert.strictEqual(parts.unmatchedLines.length, 0, "006A/B/C must not be unmatched");
const partsBy = {};
parts.allocations.forEach((row) => { partsBy[row.orderNumber] = row; });
assert.ok(partsBy.S260131);
assert.ok(partsBy.S260193);
assert.strictEqual(partsBy.S260131.cost, 1105.1); // 877.50 + 93.90 + 133.70
assert.strictEqual(partsBy.S260131.lines.length, 3);
assert.strictEqual(partsBy.S260193.cost, 268.9);
assert.strictEqual(parts.grandTotal, 1374);

// Order on the first component still covers B and C.
const orderOnA = parsePaintInvoiceText(`
SD-CABINET-006A - CHARLENE DISPLAY CABINET (2220X980X350)
S260131-F/G BLACK
1 R877.50 0,00% 15,00% R877.50 R1,009.13
SD-CABINET-006B - CHARLENE DISPLAY CABINET DOOR (1820X400)
1 R93.90 0,00% 15,00% R93.90 R107.99
SD-CABINET-006C - ARCHED DISPLAY CABINET PLATE (980X350)
1 R133.70 0,00% 15,00% R133.70 R153.76
`, { preferOrderNumbers: ["S260131"] });
assert.strictEqual(orderOnA.unmatchedLines.length, 0);
assert.strictEqual(orderOnA.allocations.length, 1);
assert.strictEqual(orderOnA.allocations[0].cost, 1105.1);

// Ultimate invoice: colour tag sits under the last component, next product has its own tag.
const ultimate = `
Copy Tax Invoice
ULTIMATE POWDER COATERS (PTY) LTD
SD-CABINET-006A - CHARLENE DISPLAY
CABINET (2220X980X350)
1 R877.50 0,00% 15,00% R877.50 R1,009.13
SD-CABINET-006B - CHARLENE DISPLAY
CABINET DOOR (1820X400)
1 R93.90 0,00% 15,00% R93.90 R107.99
SD-CABINET-006C - ARCHED DISPLAY
CABINET PLATE (980X350)
1 R133.70 0,00% 15,00% R133.70 R153.76
S260131-F/G BLACK
SD-BOOKSHELF-002 - TALITHA
BOOKSHELF (1470X260X300)
1 R268.90 0,00% 15,00% R268.90 R309.24
S260193-F/G BLACK
`;
const ult = parsePaintInvoiceText(ultimate, {
  preferOrderNumbers: ["S260131", "S260193"]
});
assert.strictEqual(ult.unmatchedLines.length, 0, "006A/B must follow 006C onto S260131");
const ultBy = {};
ult.allocations.forEach((row) => { ultBy[row.orderNumber] = row; });
assert.strictEqual(ultBy.S260131.cost, 1105.1);
assert.strictEqual(ultBy.S260131.lines.length, 3);
assert.strictEqual(ultBy.S260193.cost, 268.9);
assert.strictEqual(ult.grandTotal, 1374);

// Two colour tags S260214A / S260214B after shared 001A+001B components.
const dualLetters = `
SD-CABINET-001A - ARCHED DISPLAY
CABINET (2100X450X1000)	2	R1,406.50	0,00%	15,00%	R2,813.00	R3,234.95
SD-CABINET-001B - ARCHED DISPLAY
CABINET DOOR (1820X475X20)	4	R93.90	0,00%	15,00%	R375.60	R431.94
S260214A-F/G BLACK
S260214B-F/G BLACK
`;
const dual = parsePaintInvoiceText(dualLetters, {
  preferOrderNumbers: ["S260214 A", "S260214 B"]
});
assert.strictEqual(dual.unmatchedLines.length, 0);
assert.strictEqual(dual.allocations.length, 2);
const dualBy = {};
dual.allocations.forEach((row) => { dualBy[row.orderNumber] = row; });
assert.ok(dualBy["S260214 A"]);
assert.ok(dualBy["S260214 B"]);
assert.strictEqual(dualBy["S260214 A"].cost, 1594.3); // 1406.50 + 187.80
assert.strictEqual(dualBy["S260214 B"].cost, 1594.3);
assert.strictEqual(dual.grandTotal, 3188.6);

console.log("paint-invoice-parse.test.js ok");
