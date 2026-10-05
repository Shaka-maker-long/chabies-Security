const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-debt-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, persistWorkbook } = require("./workbook-store");
const db = require("./db");

const TINY_PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const proof = { filename: "pop.png", mime: "image/png", data: TINY_PNG };

initWorkbook();
persistWorkbook();

const order = db.upsertOrder({
  order_number: "S-D1",
  status: "Not Yet Started",
  product: "Air Chair",
  client_name: "Winelands Design Studio",
  price_excl_vat: "1000.00",
  amount_paid: "0"
});
assert.strictEqual(order.product, "Air Chair");
const decorated = db.decorateMoney(order);
assert.ok(decorated.total_excl.indexOf("R ") === 0);
assert.ok(db.listDebtors().some((o) => o.order_number === "S-D1" && o.product === "Air Chair"));

assert.throws(() => db.recordPayment("S-D1", "150", "deposit"), /proof of payment/i);
assert.throws(() => db.recordPayment("S-D1", "150", "deposit", {}), /proof of payment/i);

const paid = db.recordPayment("S-D1", "150", "deposit", proof);
assert.strictEqual(paid.paid, "R 150.00");
assert.ok(db.parseMoney(paid.owing) > 0);
assert.ok(Array.isArray(paid.payments) && paid.payments[0].filename === "pop.png");
assert.ok(paid.payments[0].id);

const hist = db.listDebtorHistory();
assert.ok(hist.some((h) => h.order_number === "S-D1" && h.has_file && h.product === "Air Chair"));
const rec = hist.find((h) => h.order_number === "S-D1");
const file = db.readPaymentProof(rec.id);
assert.ok(file);
assert.ok(file.buffer.length > 0);
assert.ok(/png/i.test(file.mime) || /png/i.test(file.filename));
assert.ok(fs.existsSync(path.join(dir, "debtor-payments", rec.id, rec.has_file ? "proof.png" : "proof.png")));

// Office order edits omit payments — history must survive the upsert.
const edited = db.upsertOrder({
  order_number: "S-D1",
  status: "In Progress",
  product: "Air Chair",
  client_name: "Winelands Design Studio",
  price_excl_vat: "1000.00",
  amount_paid: "150.00"
});
assert.ok(Array.isArray(edited.payments) && edited.payments.length === 1);
assert.strictEqual(edited.payments[0].id, rec.id);
assert.ok(db.listDebtorHistory().some((h) => h.order_number === "S-D1" && h.id === rec.id && h.has_file));
assert.ok(db.readPaymentProof(rec.id));

// Paid on the order with no ledger entry still shows on History (recovery after wipe).
db.upsertOrder({
  order_number: "S-D2",
  status: "Not Yet Started",
  product: "Serena Sideboard",
  client_name: "Paid Client",
  price_excl_vat: "2000.00",
  amount_paid: "500.00",
  payment_date: "24/09/2026",
  payments: []
});
const recovered = db.listDebtorHistory().filter((h) => h.order_number === "S-D2");
assert.strictEqual(recovered.length, 1, "paid order without POP still appears in history");
assert.strictEqual(recovered[0].synthetic, true);
assert.ok(db.parseMoney(recovered[0].amount) === 500);
assert.ok(!recovered[0].has_file);
assert.ok(/no proof/i.test(recovered[0].note));

// Ledger payment must not also synthesize a duplicate for the same order.
const d1Rows = db.listDebtorHistory().filter((h) => h.order_number === "S-D1");
assert.strictEqual(d1Rows.length, 1);
assert.strictEqual(d1Rows[0].synthetic, false);

// Empty payment placeholders must not create synthetic history for unpaid orders.
db.upsertOrder({
  order_number: "S-D3",
  status: "Not Yet Started",
  product: "Air Chair",
  client_name: "Zero Paid",
  price_excl_vat: "100.00",
  amount_paid: "0",
  payments: []
});
assert.ok(!db.listDebtorHistory().some((h) => h.order_number === "S-D3"));

db.deleteOrder("S-D1");
assert.ok(!db.readPaymentProof(rec.id));
assert.ok(!db.listDebtorHistory().some((h) => h.order_number === "S-D1"));
assert.ok(!db.listDebtors().some((o) => o.order_number === "S-D1"), "removing the order also takes it off Debtors");

db.deleteOrder("S-D2");
db.deleteOrder("S-D3");

console.log("debtors.test.js ok");
