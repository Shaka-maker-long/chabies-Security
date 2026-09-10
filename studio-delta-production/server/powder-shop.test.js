const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const assert = require("assert");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-paint-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const db = require("./db");
const staff = require("./staff");
const { mountOffice } = require("./office");
const { callShopFunction } = require("./gas");
const paint = require("./powder-shop");

initWorkbook();
staff.upsertUser({
  name: "Office Boss",
  access: "Admin",
  role: "Admin",
  password: "admin",
  seeDebtors: "Yes"
});
staff.upsertUser({
  name: "Welder Sam",
  access: "Production",
  role: "Welding",
  password: "1234",
  tasks: ["Welding"]
});

function seed(num, status) {
  return db.upsertOrder({
    order_number: num,
    status,
    type: "Cabinet",
    category: "Bedroom",
    product: "Blaire Dresser",
    powder_coating: "Ferrograin Black",
    price_excl_vat: "100.00"
  });
}

const invoice = {
  filename: "paint-batch.pdf",
  mime: "application/pdf",
  data: "data:application/pdf;base64," + Buffer.from("%PDF-1.4 paint shop invoice").toString("base64")
};

(async function main() {
  ["SD-P1", "SD-P2", "SD-P3", "SD-P4", "SD-P5"].forEach((num) => seed(num, "Ready for Powder Coating"));
  seed("SD-WELD", "Ready for Welding");

  const sendFive = paint.sendToPaintShop(["SD-P1", "SD-P2", "SD-P3", "SD-P4", "SD-P5"], "Office Boss");
  assert.strictEqual(sendFive.orderNumbers.length, 5);
  assert.strictEqual(sendFive.sentCount, 5);
  assert.strictEqual(sendFive.readyCount, 0);
  ["SD-P1", "SD-P2", "SD-P3", "SD-P4", "SD-P5"].forEach((num) => {
    assert.strictEqual(db.listOrders().find((o) => o.order_number === num).status, "Sent to Paint Shop");
  });

  let failed = false;
  try {
    paint.sendToPaintShop(["SD-WELD"], "Office Boss");
  } catch (e) {
    failed = /Ready for Powder Coating/.test(e.message);
  }
  assert.ok(failed, "welding orders must not be sent to the paint shop");

  failed = false;
  try {
    paint.receiveFromPaintShop({
      orders: [
        { orderNumber: "SD-P1", cost: "120" },
        { orderNumber: "SD-P2", cost: "80" },
        { orderNumber: "SD-P3" }
      ],
      invoice
    }, "Office Boss");
  } catch (e) {
    failed = /cost on SD-P3/.test(e.message);
  }
  assert.ok(failed, "receive must require a cost on every selected order");

  failed = false;
  try {
    paint.receiveFromPaintShop({
      orders: [
        { orderNumber: "SD-P1", cost: "120" },
        { orderNumber: "SD-P2", cost: "80" },
        { orderNumber: "SD-P3", cost: "40" }
      ]
    }, "Office Boss");
  } catch (e) {
    failed = /invoice/.test(String(e.message || e).toLowerCase());
  }
  assert.ok(failed, "receive must require one invoice for the batch");

  const recvThree = paint.receiveFromPaintShop({
    orders: [
      { orderNumber: "SD-P1", cost: "120.50" },
      { orderNumber: "SD-P2", cost: "80" },
      { orderNumber: "SD-P3", cost: "40" }
    ],
    invoice
  }, "Office Boss");
  assert.strictEqual(recvThree.orderNumbers.length, 3);
  assert.strictEqual(recvThree.sentCount, 2);
  assert.ok(recvThree.invoiceId);
  ["SD-P1", "SD-P2", "SD-P3"].forEach((num) => {
    assert.strictEqual(db.listOrders().find((o) => o.order_number === num).status, "Ready for Assembly");
  });
  ["SD-P4", "SD-P5"].forEach((num) => {
    assert.strictEqual(db.listOrders().find((o) => o.order_number === num).status, "Sent to Paint Shop");
  });
  const firstBatch = recvThree.received[0];
  assert.strictEqual(firstBatch.orders.length, 3);
  assert.strictEqual(firstBatch.invoiceId, recvThree.invoiceId);
  assert.strictEqual(firstBatch.total, "240.50");

  const recvTwo = paint.receiveFromPaintShop({
    orders: [
      { orderNumber: "SD-P4", cost: "15" },
      { orderNumber: "SD-P5", cost: "25" }
    ],
    invoice: Object.assign({}, invoice, { filename: "second-load.pdf" })
  }, "Office Boss");
  assert.strictEqual(recvTwo.sentCount, 0);
  assert.notStrictEqual(recvTwo.invoiceId, recvThree.invoiceId);
  assert.strictEqual(db.listOrders().find((o) => o.order_number === "SD-P4").status, "Ready for Assembly");
  assert.strictEqual(db.listOrders().find((o) => o.order_number === "SD-P5").status, "Ready for Assembly");
  assert.ok(paint.readInvoiceFile(recvThree.invoiceId));
  assert.ok(paint.readInvoiceFile(recvTwo.invoiceId));

  const fromRfaPrep = await callShopFunction("getStartStatusForRole", ["Ready for Assembly", "Paint Preparation"]);
  assert.strictEqual(fromRfaPrep, "Paint Preparation");
  const fromRfaAssembly = await callShopFunction("getStartStatusForRole", ["Ready for Assembly", "Assembly"]);
  assert.strictEqual(fromRfaAssembly, "Assembly");
  const fromRfaPaint = await callShopFunction("getStartStatusForRole", ["Ready for Assembly", "Painting"]);
  assert.strictEqual(fromRfaPaint, "Assembly", "painting is not a choice on Ready for Assembly");
  const fromReadyPaint = await callShopFunction("getStartStatusForRole", ["Ready for Painting", "Painting"]);
  assert.strictEqual(fromReadyPaint, "Painting");

  const sentRow = seed("SD-SENT", "Sent to Paint Shop");
  const blocked = await callShopFunction("startOrder", [sentRow.id, "Welder Sam", "Welding", [], "", false, [], { understood: true, highlights: [] }]);
  assert.strictEqual(blocked.success, false, JSON.stringify(blocked));
  assert.ok(/paint shop/i.test(blocked.message || ""), blocked.message);

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
  seed("SD-API1", "Ready for Powder Coating");
  seed("SD-API2", "Ready for Powder Coating");
  const sentApi = await fetch(base + "/api/office/paint-shop/send", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({ orderNumbers: ["SD-API1", "SD-API2"] })
  });
  const sentJson = await sentApi.json();
  assert.ok(sentJson.ok, JSON.stringify(sentJson));
  const recvApi = await fetch(base + "/api/office/paint-shop/receive", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({
      orders: [{ orderNumber: "SD-API1", cost: "33" }],
      invoice
    })
  });
  const recvJson = await recvApi.json();
  assert.ok(recvJson.ok, JSON.stringify(recvJson));
  assert.strictEqual(db.listOrders().find((o) => o.order_number === "SD-API1").status, "Ready for Assembly");
  assert.strictEqual(db.listOrders().find((o) => o.order_number === "SD-API2").status, "Sent to Paint Shop");
  assert.ok((recvJson.sent || []).some((o) => o.order_number === "SD-API2"));
  assert.ok(!(recvJson.sent || []).some((o) => o.order_number === "SD-API1"));
  const file = await fetch(base + "/api/office/paint-shop/invoices/" + recvJson.invoiceId, {
    headers: { "x-sd-token": session.token }
  });
  assert.strictEqual(file.status, 200);
  server.close();
  console.log("powder-shop.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
