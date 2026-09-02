const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-sql-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const sqlite = require("./sqlite-store");

assert.ok(sqlite.sqliteAvailable(), "Node sqlite must be available");
initWorkbook();
const book = getBook();
book.getSheetByName("Users").appendRow(["Shaka", "Admin", "x", "", "Admin", "Yes"]);
persistWorkbook();

const enquiry = db.upsertEnquiry({
  date_enquired: "01/09/2026",
  client_name: "SQL Client",
  product: "Air Chair",
  status: "New"
});
db.upsertOrder({
  order_number: "S-SQL1",
  status: "Not Yet Started",
  client_name: "SQL Client",
  price_excl_vat: "10.00"
});

const n = sqlite.counts();
assert.ok(n.users >= 1, "users copied into SQLite");
assert.ok(n.orders >= 1, "orders copied into SQLite");
assert.ok(n.enquiries >= 1, "enquiries copied into SQLite");
assert.ok(fs.existsSync(sqlite.sqlitePath()));

const user = sqlite.open().prepare("SELECT name, access FROM users WHERE name = ?").get("Shaka");
assert.ok(user);
assert.strictEqual(user.access, "Admin");
const order = sqlite.open().prepare("SELECT client_name, status FROM orders WHERE order_number = ?").get("S-SQL1");
assert.ok(order);
assert.strictEqual(order.client_name, "SQL Client");
const row = sqlite.open().prepare("SELECT client_name, status FROM enquiries WHERE enquiry_no = ?").get(enquiry.enquiry_no);
assert.ok(row);
assert.strictEqual(row.status, "New");

const loaded = sqlite.loadOffice();
assert.ok(loaded.enquiries.some((e) => e.client_name === "SQL Client"));
assert.ok(sqlite.loadWorkbookJson());

const info = db.persistenceInfo();
assert.ok(/SQLite/i.test(info.database));
assert.ok(info.sqliteExists);
assert.ok(info.sqliteUsers >= 1);

console.log("sqlite-store.test.js ok");
