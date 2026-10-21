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

console.log("paint-invoice-parse.test.js ok");
