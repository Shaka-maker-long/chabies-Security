"use strict";

const assert = require("assert");
const { parsePaintInvoiceText, normalizeOrderToken, parseRandToken } = require("./paint-invoice-parse");

assert.strictEqual(normalizeOrderToken("S260244"), "S260244");
assert.strictEqual(normalizeOrderToken("S260249(A-L)"), "S260249");
assert.strictEqual(parseRandToken("R2,219.80"), 2219.8);
assert.strictEqual(parseRandToken("R74.40"), 74.4);

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

(async function pdfRoundTrip() {
  const PDFDocument = require("pdfkit");
  const { parsePaintInvoiceBuffer } = require("./paint-invoice-parse");
  const doc = new PDFDocument({ size: "A4", margin: 40 });
  const chunks = [];
  doc.on("data", (c) => chunks.push(c));
  const done = new Promise((resolve) => doc.on("end", () => resolve(Buffer.concat(chunks))));
  doc.fontSize(11).text("Paint shop invoice");
  doc.text("Order S260301");
  doc.text("SD-CABINET-112A - NAOMI ARCHED CABINET (2000X1400X680) 1 R2,219.80 0,00% 15,00% R2,219.80 R2,552.77");
  doc.text("SD-CABINET-112B - NAOMI ARCHED CABINET DOOR (840X330) 4 R74.40 0,00% 15,00% R297.60 R342.24");
  doc.end();
  const buf = await done;
  const parsed = await parsePaintInvoiceBuffer(buf, { preferOrderNumbers: ["S260301"] });
  assert.ok(parsed.ok);
  assert.strictEqual(parsed.allocations[0].orderNumber, "S260301");
  assert.ok(parsed.allocations[0].cost >= 2500);
  console.log("paint-invoice-parse.test.js ok");
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
