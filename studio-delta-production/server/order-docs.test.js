const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const assert = require("assert");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-order-docs-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const staff = require("./staff");
const db = require("./db");
const { mountOffice } = require("./office");

initWorkbook();
staff.upsertUser({
  name: "Office Boss",
  access: "Admin",
  role: "Manager",
  password: "admin",
  seeDebtors: "Yes",
  canManageUsers: true
});

const enquiry = db.upsertEnquiry({
  date_enquired: "14/09/2026",
  client_name: "Docs Client",
  product: "Gate",
  status: "Ordered",
  quote_no: "SOQ4401"
});
const file = db.saveEnquiryAttachment(
  enquiry.enquiry_no,
  "drawing",
  "data:application/pdf;base64,JVBERi0x",
  "plan.pdf"
);
const raw = db.getEnquiryRaw(enquiry.enquiry_no);
raw.drawing = { required: true, file: file, assignee: "Erin" };
db.saveEnquiryRecord(raw);

db.upsertOrder({
  order_number: "S260441",
  enquiry_no: enquiry.enquiry_no,
  quote_number: "SOQ4401",
  status: "Ready for Steelwork",
  type: "Standard",
  category: "Driveway",
  product: "Gate",
  client_name: "Docs Client"
});
db.upsertOrder({
  order_number: "S260442",
  status: "Not Yet Started",
  type: "Standard",
  product: "Shelf"
});

const waiting = db.upsertEnquiry({
  date_enquired: "14/09/2026",
  client_name: "Wait Client",
  product: "Desk",
  status: "Ordered",
  quote_no: "SOQ4402"
});
const waitRaw = db.getEnquiryRaw(waiting.enquiry_no);
waitRaw.drawing = { required: true, file: null, assignee: "Erin" };
db.saveEnquiryRecord(waitRaw);
db.upsertOrder({
  order_number: "S260443",
  enquiry_no: waiting.enquiry_no,
  quote_number: "SOQ4402",
  status: "Waiting for drawing",
  type: "Standard",
  product: "Desk"
});

const log = getBook().getSheetByName("Production_Log");
log.appendRow([
  "log-1",
  "S260441",
  "Willard",
  "Pre-Powder Coating QC",
  "Complete",
  new Date("2026-09-14T08:00:00+02:00"),
  new Date("2026-09-14T09:00:00+02:00"),
  "Pass\n\nQC PDF: https://drive.google.com/file/d/qc441",
  "",
  "",
  0,
  "",
  "{}"
]);
persistWorkbook();

(async function main() {
  const app = express();
  app.use(express.json());
  mountOffice(app);
  const server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const base = "http://127.0.0.1:" + server.address().port;
  const login = await fetch(base + "/api/office/login", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: "Office Boss", password: "admin" })
  }).then((r) => r.json());
  assert.ok(login.ok, JSON.stringify(login));
  const listed = await fetch(base + "/api/office/orders", {
    headers: { "x-sd-token": login.token }
  }).then((r) => r.json());
  assert.ok(listed.ok, JSON.stringify(listed));
  const withDraw = (listed.rows || []).find((r) => r.order_number === "S260441");
  const plain = (listed.rows || []).find((r) => r.order_number === "S260442");
  const waitingRow = (listed.rows || []).find((r) => r.order_number === "S260443");
  assert.ok(withDraw, "order with drawing");
  assert.ok(withDraw.drawing_url.indexOf("/files/drawing") !== -1, JSON.stringify(withDraw));
  assert.ok((withDraw.qc_pdfs || []).some((p) => String(p.url).indexOf("qc441") !== -1), JSON.stringify(withDraw.qc_pdfs));
  assert.ok(!plain.drawing_url);
  assert.ok(!(plain.qc_pdfs || []).length);
  assert.strictEqual(waitingRow.drawing_required, true);
  assert.ok(!waitingRow.drawing_url, "waiting drawing has no file yet");
  server.close();
  console.log("order-docs.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
