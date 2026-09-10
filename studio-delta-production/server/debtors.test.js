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

db.deleteOrder("S-D1");
assert.ok(!db.readPaymentProof(rec.id));
assert.ok(!db.listDebtorHistory().some((h) => h.order_number === "S-D1"));

console.log("debtors.test.js ok");
