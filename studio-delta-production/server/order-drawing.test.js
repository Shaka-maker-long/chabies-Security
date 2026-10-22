const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-order-draw-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const staff = require("./staff");
const db = require("./db");
const orderDrawing = require("./order-drawing");
const alerts = require("./drawing-print-alerts");

initWorkbook();
staff.upsertUser({ name: "Office Boss", access: "Admin", role: "Manager", password: "x", seeDebtors: "Yes" });
staff.upsertUser({ name: "Siya", access: "Admin", role: "Production Manager", password: "x", seeDebtors: "Yes" });
staff.upsertUser({ name: "Coster", access: "Admin", role: "Admin", password: "x", seeDebtors: "Yes" });

const pdf = "data:application/pdf;base64," + Buffer.from("%PDF-1.1\n%%EOF\n").toString("base64");

const waitEnq = db.upsertEnquiry({
  date_enquired: "22/10/2026",
  client_name: "Wait Draw",
  product: "Gate",
  status: "Ordered",
  quote_no: "SOQ8801"
});
const waitRaw = db.getEnquiryRaw(waitEnq.enquiry_no);
waitRaw.drawing = { required: true, file: null, assignee: "Erin" };
db.saveEnquiryRecord(waitRaw);
db.upsertOrder({
  order_number: "S260880",
  enquiry_no: waitEnq.enquiry_no,
  quote_number: "SOQ8801",
  status: "Waiting for drawing",
  product: "Gate",
  client_name: "Wait Draw"
});

assert.throws(
  () => orderDrawing.updateFromOrder("S260880", "Office Boss", { file_base64: pdf, file_name: "a.pdf" }),
  /Tick|correct drawing/i
);

const first = orderDrawing.updateFromOrder("S260880", "Office Boss", {
  file_base64: pdf,
  file_name: "gate.pdf",
  file_confirmed: true
});
assert.ok(first.drawing_url.indexOf("/files/drawing") !== -1);
assert.strictEqual(first.notified_siya, false);
assert.strictEqual(db.listOrders().find((o) => o.order_number === "S260880").status, "Not Yet Started");
assert.strictEqual(alerts.pollFor("Siya").alerts.length, 0);
assert.strictEqual(db.getEnquiryRaw(waitEnq.enquiry_no).drawing.approval.status, "approved");

const liveEnq = db.upsertEnquiry({
  date_enquired: "22/10/2026",
  client_name: "Live Draw",
  product: "Table",
  status: "Ordered",
  quote_no: "SOQ8802"
});
const liveRaw = db.getEnquiryRaw(liveEnq.enquiry_no);
liveRaw.drawing = {
  required: true,
  file: db.saveEnquiryAttachment(liveEnq.enquiry_no, "drawing", pdf, "old.pdf"),
  approval: { status: "approved" }
};
db.saveEnquiryRecord(liveRaw);
db.upsertOrder({
  order_number: "S260881",
  enquiry_no: liveEnq.enquiry_no,
  quote_number: "SOQ8802",
  status: "Welding",
  product: "Table",
  client_name: "Live Draw"
});
fs.writeFileSync(path.join(dir, "job-cards.json"), JSON.stringify({
  S260881: { order_number: "S260881", created_at: "2026-10-22T08:00:00.000Z" }
}));

const updated = orderDrawing.updateFromOrder("S260881", "Office Boss", {
  file_base64: pdf,
  file_name: "table-new.pdf",
  file_confirmed: true
});
assert.strictEqual(updated.notified_siya, true);
assert.strictEqual(db.listOrders().find((o) => o.order_number === "S260881").status, "Welding");
const siya = alerts.pollFor("Siya");
assert.strictEqual(siya.alerts.length, 1);
assert.strictEqual(siya.alerts[0].order_number, "S260881");
assert.ok(siya.alerts[0].drawing_url);
assert.strictEqual(alerts.pollFor("Coster").alerts.length, 0);
assert.throws(() => alerts.acknowledge(siya.alerts[0].id, "Coster"), /Siya/);
const seen = alerts.acknowledge(siya.alerts[0].id, "Siya");
assert.strictEqual(seen.status, "seen");
assert.strictEqual(alerts.pollFor("Siya").alerts.length, 0);

console.log("order-drawing.test.js ok");
