const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-cat-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const catalog = require("./product-catalog");

initWorkbook();
persistWorkbook();

const snap = catalog.snapshotCatalog();
assert.ok(snap.total >= 100);
assert.ok(snap.missingCount >= 1);
const air = snap.products.find((p) => p.name === "Air Chair");
assert.ok(air);
assert.strictEqual(air.hasImage, false);

assert.throws(() => catalog.upsertProduct({ name: "Air Chair", imageUrl: "not-a-url" }), /http/i);

const saved = catalog.upsertProduct({
  name: "Air Chair",
  imageUrl: "https://www.studiodelta.co.za/air-chair.jpg"
});
assert.strictEqual(saved.name, "Air Chair");
assert.strictEqual(saved.imageUrl, "https://www.studiodelta.co.za/air-chair.jpg");
assert.ok(catalog.lookupProduct("Air Chair").imageUrl.indexOf("air-chair") !== -1);
assert.ok(catalog.snapshotCatalog().products.find((p) => p.name === "Air Chair").hasImage);

const file = path.join(dir, "product-catalog-overrides.json");
assert.ok(fs.existsSync(file));

const neu = catalog.upsertProduct({
  name: "Nova Test Bench",
  imageUrl: "https://www.studiodelta.co.za/nova-test.jpg",
  height: 900,
  width: 1200,
  depth: 400
});
assert.ok(neu.custom);
assert.strictEqual(neu.height, 900);
assert.ok(catalog.lookupProduct("Nova Test Bench"));
assert.ok(db.listDropdowns().product.some((p) => p === "Nova Test Bench"));

catalog.deleteProductOverride("Air Chair");
assert.strictEqual(catalog.lookupProduct("Air Chair").imageUrl, "");
assert.ok(catalog.lookupProduct("Nova Test Bench").imageUrl);

const html = fs.readFileSync(path.join(__dirname, "../public/orders-products.html"), "utf8");
assert.ok(html.indexOf("/api/office/product-catalog") !== -1);
assert.ok(html.indexOf("Missing photo") !== -1);
assert.ok(html.indexOf("Save product photo") !== -1);

console.log("product-catalog.test.js ok");
