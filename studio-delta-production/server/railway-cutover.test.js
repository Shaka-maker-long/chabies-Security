const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-cutover-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_MIGRATE;
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.SHEET_ID;

const { Spreadsheet } = require("./sheets");
const store = require("./workbook-store");
const db = require("./db");

store.initWorkbook();

assert.strictEqual(store.googleMigrateEnabled(), false);
assert.strictEqual(store.usersNeedGoogleCopy(), true);

(async function main() {
  const book = await store.maybeImportGoogleOnce();
  assert.ok(book);
  assert.strictEqual(store.usersEmpty(), true);

  const users = store.getBook().getSheetByName("Users");
  users.getRange(2, 1, 1, 6).setValues([["Admin", "Admin", "admin", "", "Admin", "Yes"]]);
  assert.strictEqual(store.usersNeedGoogleCopy(), true);
  users.getRange(2, 1, 1, 6).setValues([["Siya", "Admin", "1234", "", "Admin", "Yes"]]);
  assert.strictEqual(store.usersNeedGoogleCopy(), false);

  process.env.GOOGLE_SERVICE_ACCOUNT_JSON = "{\"type\":\"service_account\"}";
  process.env.SHEET_ID = "not-used";
  assert.strictEqual(store.googleMigrateEnabled(), false, "credentials alone must not enable Google as the database");
  process.env.GOOGLE_MIGRATE = "1";
  assert.strictEqual(store.googleMigrateEnabled(), true);
  delete process.env.GOOGLE_MIGRATE;
  delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  delete process.env.SHEET_ID;

  const remote = new Spreadsheet(null, "old-sheets");
  const sh = remote.insertSheet("Enquiries");
  sh.getRange(1, 1, 1, 5).setValues([["ENQUIRY NO", "CLIENT NAME", "PRODUCT", "STATUS", "PROVINCE"]]);
  sh.getRange(2, 1, 1, 5).setValues([["#2101", "Pat Client", "Slider", "New", "Gauteng"]]);
  const copied = db.copyEnquiriesFromWorkbook(remote);
  assert.strictEqual(copied.imported, 1);
  const row = db.getEnquiry("#2101");
  assert.ok(row);
  assert.strictEqual(row.client_name, "Pat Client");
  assert.strictEqual(row.product, "Slider");

  const pack = db.railwayBackup();
  assert.strictEqual(pack.database, "railway");
  assert.ok(pack.workbook.sheets.Users);
  assert.ok(Array.isArray(pack.office.enquiries));
  assert.ok(pack.office.enquiries.some((e) => e.enquiry_no === "#2101"));

  const healthJs = fs.readFileSync(path.join(__dirname, "index.js"), "utf8");
  assert.ok(healthJs.indexOf("sheetsLive: false") !== -1);
  assert.ok(healthJs.indexOf("SHEET_ID ||") === -1);

  console.log("railway-cutover.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
