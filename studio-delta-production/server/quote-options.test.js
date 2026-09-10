const assert = require("assert");
const quoteOptions = require("./quote-options");

assert.strictEqual(quoteOptions.optionLetter(0), "A");
assert.strictEqual(quoteOptions.optionLetter(1), "B");
assert.strictEqual(quoteOptions.nextOptionLetter([]), "A");

const first = quoteOptions.normalizeQuotes([{ quote_no: "SOQ1" }]);
assert.strictEqual(first[0].option, "A");
assert.strictEqual(first[0].kind, "option");

const revised = quoteOptions.normalizeQuotes([
  { quote_no: "SOQ1" },
  { quote_no: "SOQ2" }
]);
assert.strictEqual(revised[1].option, "A");
assert.strictEqual(revised[1].kind, "revision");
const liveRev = quoteOptions.liveQuoteOptions(revised);
assert.strictEqual(liveRev.length, 1);
assert.strictEqual(liveRev[0].quote_no, "SOQ2");
assert.strictEqual(quoteOptions.nextOptionLetter(revised), "B");
assert.strictEqual(quoteOptions.quoteOptionsLabel(revised), "SOQ2 A");

const two = [
  { quote_no: "SOQ1", option: "A", kind: "option", products: [{ product: "Air Chair" }] },
  { quote_no: "SOQ2", option: "B", kind: "option", products: [{ product: "Air Chair", variation: "Extra shelf" }] }
];
assert.strictEqual(quoteOptions.liveQuoteOptions(two).length, 2);
assert.strictEqual(quoteOptions.nextOptionLetter(two), "C");
assert.strictEqual(quoteOptions.quoteOptionsLabel(two), "SOQ1 A / SOQ2 B");
assert.strictEqual(quoteOptions.findLiveOption(two, "B").quote_no, "SOQ2");
assert.strictEqual(quoteOptions.findLiveOption(two, "SOQ1").option, "A");
assert.strictEqual(quoteOptions.chosenQuote({ quotes: two, chosen_option: "B" }).quote_no, "SOQ2");
assert.ok(/Extra shelf/.test(quoteOptions.optionSummary(two[1])));

const row = { products: [{ product: "Old" }], quote_no: "SOQ1" };
quoteOptions.applyQuoteSnapshot(row, two[1]);
assert.strictEqual(row.quote_no, "SOQ2");
assert.strictEqual(row.products[0].variation, "Extra shelf");

console.log("quote-options.test.js ok");
