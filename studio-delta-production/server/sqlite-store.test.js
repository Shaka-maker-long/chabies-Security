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

db.recordPayment("S-SQL1", "2.50", "deposit");
db.upsertScheduleRow({
  order_number: "S-SQL1",
  product: "Air Chair",
  status: "Not Yet Started",
  cells: { "2026-09-02": "cut" }
});

const after = sqlite.counts();
assert.ok(after.sheets >= 8, "every workbook tab is a SQLite sheet");
assert.ok(after.sheetRows >= 8, "every tab’s rows are in sheet_rows");
assert.ok(after.dropdowns >= 1, "dropdowns copied into SQLite");
assert.ok(after.payments >= 1, "payments copied into SQLite");
assert.ok(after.officeScheduleRows >= 1, "office schedule copied into SQLite");

const titles = sqlite.open().prepare("SELECT title FROM sheets").all().map((r) => r.title);
["Users", "ORDERS", "Production_Log", "Overview", "Rates", "Steel_Profiles", "Backboards", "Idle_Alerts", "Schedule", "Task_Durations"].forEach((title) => {
  assert.ok(titles.indexOf(title) !== -1, "missing sheet " + title);
});
const userRow = sqlite.open().prepare("SELECT json FROM sheet_rows WHERE title = 'Users' AND row_idx = 1").get();
assert.ok(userRow);
assert.ok(JSON.parse(userRow.json).indexOf("Shaka") !== -1);
const drop = sqlite.open().prepare("SELECT value FROM dropdowns WHERE group_name = 'product' LIMIT 1").get();
assert.ok(drop && drop.value);
const pay = sqlite.open().prepare("SELECT amount, note FROM payments WHERE order_number = ?").get("S-SQL1");
assert.ok(pay);
assert.strictEqual(pay.note, "deposit");

sqlite.saveSessions({ tok1: { name: "Shaka", access: "Admin", savedAt: Date.now() } });
const sess = sqlite.loadSessions();
assert.ok(sess && sess.tok1 && sess.tok1.name === "Shaka");

const loaded = sqlite.loadOffice();
assert.ok(loaded.enquiries.some((e) => e.client_name === "SQL Client"));
assert.ok(loaded.paymentsByOrder["S-SQL1"]);
assert.ok(loaded.schedule_rows.some((r) => r.order_number === "S-SQL1"));
assert.ok(Object.keys(loaded.dropdowns).length >= 1);

sqlite.open().prepare("DELETE FROM blobs WHERE kind = 'workbook'").run();
const fromSheets = sqlite.loadWorkbookFromSheets();
assert.ok(fromSheets && fromSheets.sheets.Users);
assert.ok(fromSheets.sheets.Production_Log);
assert.ok(sqlite.loadWorkbookJson().sheets.Users);

const info = db.persistenceInfo();
assert.ok(/SQLite/i.test(info.database));
assert.ok(info.sqliteExists);
assert.ok(info.sqliteUsers >= 1);
assert.ok(info.sqliteSheets >= 8);
assert.ok(info.sqliteDropdowns >= 1);

console.log("sqlite-store.test.js ok");
