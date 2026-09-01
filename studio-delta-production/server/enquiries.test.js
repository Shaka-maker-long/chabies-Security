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
assert.strictEqual(first.products[0].product, "Daphne Rectangular Mirror");

const second = db.upsertEnquiry({
  date_enquired: "05/01/2026",
  client_name: "Sas Promotions",
  status: "Followed Up"
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

assert.throws(
  () => db.upsertEnquiry({ enquiry_no: "#1996", status: "Quoted", products: [{ product: "Daphne Rectangular Mirror", value_excl_vat: "1000" }] }),
  /quote PDF/
);

const pdf = Buffer.from("%PDF-1.1\n1 0 obj\n<<>>\nendobj\ntrailer\n<<>>\n%%EOF\n");
const quoted = db.upsertEnquiry({
  enquiry_no: "#1996",
  date_enquired: "30/11/2025",
  client_name: "Michael Cost",
  status: "Quoted",
  products: [
    { product: "Daphne Rectangular Mirror", category: "Mirror", value_excl_vat: "2500" },
    { product: "Eve Patio Table", category: "Table", value_excl_vat: "1800.5" }
  ],
  delivery_excl_vat: "350",
  quote_pdf_base64: "data:application/pdf;base64," + pdf.toString("base64"),
  quote_pdf_name: "Michael Cost quote.pdf",
  quote_pdf_confirmed: true
});
assert.strictEqual(quoted.status, "Quoted");
assert.strictEqual(quoted.products.length, 2);
assert.strictEqual(quoted.delivery_excl_vat, "350.00");
assert.strictEqual(quoted.quote_total_excl_vat, "4650.50");
assert.strictEqual(quoted.has_quote_pdf, true);
assert.ok(quoted.date_quoted);
assert.strictEqual(quoted.date_quoted, db.todayEnquiryDate());
assert.ok(db.readEnquiryQuotePdf("#1996"));

const removed = db.upsertEnquiry({
  enquiry_no: "#1996",
  date_enquired: "30/11/2025",
  client_name: "Michael Cost",
  status: "Quoted",
  products: [
    { product: "Eve Patio Table", category: "Table", value_excl_vat: "1800.5" }
  ],
  delivery_excl_vat: "350"
});
assert.strictEqual(removed.products.length, 1);
assert.strictEqual(removed.product, "Eve Patio Table");

assert.throws(
  () => db.upsertEnquiry({
    enquiry_no: "#1997",
    status: "Quoted",
    products: [{ product: "Air Chair", value_excl_vat: "100" }],
    quote_pdf_base64: "data:application/pdf;base64," + pdf.toString("base64")
  }),
  /confirm/
);

const saved = JSON.parse(fs.readFileSync(db.dbPath, "utf8"));
assert.strictEqual(saved.enquiries.length, 2);

console.log("enquiries.test.js ok");
