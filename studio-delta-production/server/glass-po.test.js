const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const assert = require("assert");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-gpo-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const staff = require("./staff");
const { mountOffice } = require("./office");
const glassPo = require("./glass-po");

initWorkbook();
staff.upsertUser({
  name: "Office Boss",
  access: "Admin",
  role: "Admin",
  password: "admin",
  seeDebtors: "Yes"
});

function addGlass(id, order) {
  const sheet = getBook().getSheetByName("Glass_To_Order");
  sheet.appendRow([id, new Date(), order, "Nomsa", "Door", "Reeded", "6mm", 1800, 500, 2, "To order"]);
  persistWorkbook();
}

const invoice = {
  filename: "glass-batch.pdf",
  mime: "application/pdf",
  data: "data:application/pdf;base64," + Buffer.from("%PDF-1.4 glass invoice").toString("base64")
};

addGlass("g1", "SD-G1");
addGlass("g2", "SD-G2");
addGlass("g3", "SD-G3");
addGlass("g4", "SD-G4");
addGlass("g5", "SD-G5");

const rates = require("./glass-rates");
rates.upsertRate({ type: "Reeded", thickness: "6mm", ratePerM2: "450" });
assert.strictEqual(rates.lineAreaM2({ height: 1800, width: 500, quantity: 2 }), 1.8);
assert.strictEqual(rates.costLine({ type: "Reeded", thickness: "6 mm", height: 1800, width: 500, quantity: 2 }).estimatedCost, "810.00");

const priced = glassPo.snapshot().toOrder.find((l) => l.id === "g1");
assert.ok(priced);
assert.strictEqual(priced.areaM2, 1.8);
assert.strictEqual(priced.estimatedCost, "810.00");

const po = glassPo.createPurchaseOrder(["g1", "g2", "g3", "g4", "g5"], "Office Boss");
assert.strictEqual(po.poNumber, "GPO-0001");
assert.strictEqual(po.toOrderCount, 0);
assert.strictEqual(po.outstandingCount, 5);

let failed = false;
try {
  glassPo.createPurchaseOrder(["g1"], "Office Boss");
} catch (e) {
  failed = /already/i.test(e.message);
}
assert.ok(failed, "cannot put the same line on a second PO");

failed = false;
try {
  glassPo.receiveGlass({
    lines: [
      { id: "g1", cost: "90" },
      { id: "g2", cost: "70" },
      { id: "g3" }
    ],
    invoice
  }, "Office Boss");
} catch (e) {
  failed = /cost/i.test(e.message);
}
assert.ok(failed, "receive must require a cost on every selected line");

failed = false;
try {
  glassPo.receiveGlass({
    lines: [
      { id: "g1", cost: "90" },
      { id: "g2", cost: "70" },
      { id: "g3", cost: "40" }
    ]
  }, "Office Boss");
} catch (e) {
  failed = /invoice/i.test(String(e.message || e));
}
assert.ok(failed, "receive must require one invoice for the batch");

const recv = glassPo.receiveGlass({
  lines: [
    { id: "g1", cost: "90.50" },
    { id: "g2", cost: "70" },
    { id: "g3", cost: "40" }
  ],
  invoice
}, "Office Boss");
assert.strictEqual(recv.lineIds.length, 3);
assert.strictEqual(recv.outstandingCount, 2);
assert.ok(recv.outstanding.some((l) => l.id === "g4"));
assert.ok(recv.outstanding.some((l) => l.id === "g5"));
assert.ok(!recv.outstanding.some((l) => l.id === "g1"));
assert.strictEqual(recv.received[0].lines.length, 3);
assert.strictEqual(recv.received[0].total, "200.50");
assert.ok(glassPo.readInvoiceFile(recv.invoiceId));

const later = glassPo.receiveGlass({
  lines: [
    { id: "g4", cost: "12" },
    { id: "g5", cost: "18" }
  ],
  invoice: Object.assign({}, invoice, { filename: "glass-rest.pdf" })
}, "Office Boss");
assert.strictEqual(later.outstandingCount, 0);
assert.notStrictEqual(later.invoiceId, recv.invoiceId);

(async function main() {
  const pdf = await glassPo.buildPurchaseOrderPdf(po.poId);
  assert.ok(pdf.buffer.slice(0, 4).toString() === "%PDF");
  const latin = pdf.buffer.toString("latin1");
  const decoded = [];
  latin.replace(/<([0-9A-Fa-f]+)>/g, (_, hex) => {
    try { decoded.push(Buffer.from(hex, "hex").toString("latin1")); } catch (e) {}
    return "";
  });
  const text = latin + "\n" + decoded.join("");
  assert.ok(text.indexOf("PURCHASE ORDER") !== -1);
  assert.ok(text.indexOf("GPO-0001") !== -1);
  assert.ok(text.indexOf("Order Number") !== -1);
  assert.ok(text.indexOf("Quantity") !== -1);
  assert.ok(text.indexOf("SD-G1") !== -1);
  assert.ok(text.indexOf("Reeded") !== -1);
  addGlass("g-api-1", "SD-API-G1");
  addGlass("g-api-2", "SD-API-G2");
  const app = express();
  app.use(express.json({ limit: "8mb" }));
  mountOffice(app);
  const server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const base = "http://127.0.0.1:" + server.address().port;
  const login = await fetch(base + "/api/office/login", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: "Office Boss", password: "admin" })
  });
  const session = await login.json();
  assert.ok(session.ok, JSON.stringify(session));
  const made = await fetch(base + "/api/office/glass-po", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({ lineIds: ["g-api-1", "g-api-2"] })
  });
  const madeJson = await made.json();
  assert.ok(madeJson.ok, JSON.stringify(madeJson));
  const rec = await fetch(base + "/api/office/glass-po/receive", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({
      lines: [{ id: "g-api-1", cost: "25" }],
      invoice
    })
  });
  const recJson = await rec.json();
  assert.ok(recJson.ok, JSON.stringify(recJson));
  assert.ok((recJson.outstanding || []).some((l) => l.id === "g-api-2"));
  assert.ok(!(recJson.outstanding || []).some((l) => l.id === "g-api-1"));
  const file = await fetch(base + "/api/office/glass-po/invoices/" + recJson.invoiceId, {
    headers: { "x-sd-token": session.token }
  });
  assert.strictEqual(file.status, 200);
  const pdfRes = await fetch(base + "/api/office/glass-po/" + encodeURIComponent(madeJson.poId) + "/pdf", {
    headers: { "x-sd-token": session.token }
  });
  assert.strictEqual(pdfRes.status, 200);
  assert.ok(String(pdfRes.headers.get("content-type") || "").indexOf("pdf") !== -1);
  const savedRate = await fetch(base + "/api/office/glass-rates", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({ type: "Clear", thickness: "8mm", ratePerM2: "320" })
  });
  const savedJson = await savedRate.json();
  assert.ok(savedJson.ok, JSON.stringify(savedJson));
  assert.ok((savedJson.rates || []).some((r) => r.type === "Clear" && r.thickness === "8mm"));
  server.close();
  console.log("glass-po.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
