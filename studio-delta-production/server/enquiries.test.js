const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-enq-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";

const db = require("./db");

assert.strictEqual(db.nextEnquiryNo(), "#1996");
assert.strictEqual(db.monthFromEnquiryDate("30/11/2025"), "Nov");
assert.strictEqual(db.monthFromEnquiryDate("05/01/2026"), "Jan");

const first = db.upsertEnquiry({
  date_enquired: "30/11/2025",
  enquiry_source: "Website",
  enquiry_type: "Custom",
  client_name: "Michael Cost",
  product: "Daphne Rectangular Mirror",
  status: "New"
});
assert.strictEqual(first.enquiry_no, "#1996");
assert.strictEqual(first.month_enquired, "Nov");

const second = db.upsertEnquiry({
  date_enquired: "05/01/2026",
  client_name: "Sas Promotions",
  status: "Quoted"
});
assert.strictEqual(second.enquiry_no, "#1997");
assert.strictEqual(second.month_enquired, "Jan");
assert.strictEqual(db.nextEnquiryNo(), "#1998");

const listed = db.listEnquiries();
assert.strictEqual(listed[0].enquiry_no, "#1997");
assert.strictEqual(listed[1].enquiry_no, "#1996");

const drops = db.listEnquiryDropdowns();
assert.ok(drops.enquiry_source.indexOf("Whatsapp") >= 0);
assert.ok(drops.enquiry_type.indexOf("Catologue") >= 0);
assert.ok(drops.product.indexOf("Violet Sideboard 3-Door") >= 0);
assert.ok(drops.category.indexOf("Gate") >= 0);
assert.ok(drops.status.indexOf("Waiting on clients specifictions") >= 0);

const saved = JSON.parse(fs.readFileSync(db.dbPath, "utf8"));
assert.strictEqual(saved.enquiries.length, 2);

console.log("enquiries.test.js ok");
