const assert = require("assert");
const dates = require("./sast-date");

assert.strictEqual(dates.formatPaymentDate("21/07/2026"), "21/07/2026");
assert.strictEqual(dates.formatMonthOfSale("21/07/2026"), "July 2026");
assert.strictEqual(dates.formatPaymentDate("28/08/2026"), "28/08/2026");
assert.strictEqual(dates.formatPaymentDate("07/21/2026"), "21/07/2026", "US leftover 07/21 is 21 July");
assert.strictEqual(dates.formatPaymentDate("07/01/2026"), "07/01/2026", "07/01/2026 is 7 January in South Africa");
assert.strictEqual(dates.formatMonthOfSale("07/01/2026"), "January 2026");
assert.strictEqual(dates.formatPaymentDate("03-Sep"), "03/09/" + new Date().getFullYear());
assert.strictEqual(dates.formatPaymentDate("09/03/2026", "September 2026"), "03/09/2026");
assert.strictEqual(dates.formatMonthOfSale("09/03/2026", "September 2026"), "September 2026");
assert.strictEqual(dates.formatPaymentDate("09/03/2026"), "09/03/2026", "without a month hint, slash dates stay day-first");
assert.strictEqual(dates.formatMonthOfSale("09/03/2026"), "March 2026");

const mixed = [
  { order_number: "S260178", payment_date: "21/07/2026" },
  { order_number: "S260221 A", payment_date: "28/08/2026" },
  { order_number: "S260221 B", payment_date: "28/08/2026" },
  { order_number: "S260222", payment_date: "09/01/2026" },
  { order_number: "S260223", payment_date: "09/04/2026" },
  { order_number: "S260224", payment_date: "09/01/2026" },
  { order_number: "S260225", payment_date: "09/03/2026" },
  { order_number: "S260226 D", payment_date: "09/03/2026" },
  { order_number: "S260227", payment_date: "09/04/2026" }
];
const anchors = dates.paymentAnchors(mixed);
assert.strictEqual(dates.resolvePaymentDate("21/07/2026", "S260178", "July 2026", anchors), "21/07/2026");
assert.strictEqual(dates.resolvePaymentDate("28/08/2026", "S260221 A", "August 2026", anchors), "28/08/2026");
assert.strictEqual(dates.resolvePaymentDate("09/01/2026", "S260222", "January 2026", anchors), "01/09/2026");
assert.strictEqual(dates.resolvePaymentDate("09/04/2026", "S260223", "April 2026", anchors), "04/09/2026");
assert.strictEqual(dates.resolvePaymentDate("09/01/2026", "S260224", "January 2026", anchors), "01/09/2026");
assert.strictEqual(dates.resolvePaymentDate("09/03/2026", "S260226 D", "March 2026", anchors), "03/09/2026");
assert.strictEqual(dates.resolvePaymentDate("09/04/2026", "S260227", "April 2026", anchors), "04/09/2026");
assert.strictEqual(dates.formatMonthOfSale("03/09/2026"), "September 2026");
assert.strictEqual(dates.formatMonthOfSale("01/09/2026"), "September 2026");
assert.strictEqual(dates.resolvePaymentDate("01/09/2026", "S260186", "", anchors), "01/09/2026", "true SA 1 September stays 1 September");
assert.strictEqual(dates.resolvePaymentDate("07/01/2026", "S-JAN-PAY", "September 2026", anchors), "07/01/2026");
assert.strictEqual(
  dates.resolvePaymentDate("10/07/2026", "S260404", "", [
    { seq: 260401, t: dates.asDate("14/09/2026").getTime() }
  ]),
  "10/07/2026",
  "10/07/2026 next to a September order stays 10 July"
);

console.log("sast-date.test.js ok");
