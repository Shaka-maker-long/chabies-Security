const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const assert = require("assert");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-pclist-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const db = require("./db");
const paint = require("./powder-shop");
const powderList = require("./powder-list");
const { callShopFunction } = require("./gas");

initWorkbook();

function seedReady(num, extra) {
  return db.upsertOrder(Object.assign({
    order_number: num,
    status: "Ready for Powder Coating",
    type: "Custom",
    category: "Gate",
    product: "Steel Gate",
    powder_coating: "Ferrograin Black",
    dimensions: "1800mm H x 900mm W x 40mm D"
  }, extra || {}));
}

function pdfText(buf) {
  const latin = buf.toString("latin1");
  const decoded = [];
  latin.replace(/<([0-9A-Fa-f]+)>/g, (_, hex) => {
    try { decoded.push(Buffer.from(hex, "hex").toString("latin1")); } catch (e) {}
    return "";
  });
  return latin + "\n" + decoded.join("");
}

(async function main() {
  const empty = await powderList.createList([], "Siya");
  assert.strictEqual(empty.success, false);

  const missing = await powderList.createList([
    { order: "S260247", desc: "Steel Gate", dimensions: "1800x900x40", qty: "1" }
  ], "Siya");
  assert.strictEqual(missing.success, false);
  assert.ok(/colour/i.test(missing.error));

  seedReady("S260247");
  seedReady("S260248", { product: "Door" });

  const saved = await powderList.createList([
    {
      order: "S260247",
      desc: "Steel Gate",
      dimensions: "1800x900x40",
      qty: "1",
      color: "Ferrograin Black",
      profiles: "19mm tube"
    },
    {
      order: "",
      desc: "↳ Door",
      dimensions: "1700x400x20",
      qty: "2",
      colour: "Ferrograin Black"
    }
  ], "Siya");
  assert.ok(saved.success, JSON.stringify(saved));
  assert.ok(saved.url.indexOf("/api/powder-lists/") === 0);
  assert.ok(/^PCL-/.test(saved.number), saved.number);
  const file = powderList.readPdf(saved.id);
  assert.ok(file && file.buffer.slice(0, 5).toString() === "%PDF-");
  const text = pdfText(file.buffer);
  assert.ok(text.indexOf("PURCHASE ORDER") !== -1);
  assert.ok(text.indexOf("S260247") !== -1);
  assert.ok(text.indexOf("Steel Gate") !== -1);
  assert.ok(text.indexOf("Ferrograin Black") !== -1);
  assert.ok(text.indexOf("Profiles Used") === -1);
  assert.ok(text.indexOf("19mm tube") === -1, "profile used must not print on the PO");
  assert.strictEqual(saved.status, "Paint Shop");
  assert.deepStrictEqual(saved.orderNumbers, ["S260247"]);
  assert.strictEqual(db.listOrders().find((o) => o.order_number === "S260247").status, "Paint Shop");
  assert.ok(paint.snapshot().sent.some((o) => o.order_number === "S260247"));
  assert.strictEqual(db.listOrders().find((o) => o.order_number === "S260248").status, "Ready for Powder Coating");

  const again = await powderList.createList([
    { order: "S260247", desc: "Steel Gate", dimensions: "1800x900x40", qty: "1", color: "Ferrograin Black" }
  ], "Siya");
  assert.strictEqual(again.success, false);
  assert.ok(/Ready for Powder Coating/.test(again.error));

  seedReady("S260502", { product: "Zahara Arched Mirror", powder_coating: "Smooth Matt Black" });
  const viaGas = await callShopFunction("generatePowderCoatingList", [[
    { order: "S260502", desc: "Zahara Arched Mirror", dimensions: "1800x600x25", qty: "1", color: "Smooth Matt Black" }
  ], "Admin"]);
  assert.ok(viaGas.success, JSON.stringify(viaGas));
  assert.ok(viaGas.url.indexOf("/api/powder-lists/") === 0);
  assert.strictEqual(db.listOrders().find((o) => o.order_number === "S260502").status, "Paint Shop");

  const app = express();
  app.get("/api/powder-lists/:id/pdf", (req, res) => {
    const got = powderList.readPdf(req.params.id);
    if (!got) {
      res.status(404).end();
      return;
    }
    res.setHeader("Content-Type", "application/pdf");
    res.send(got.buffer);
  });
  const server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  const res = await fetch("http://127.0.0.1:" + port + viaGas.url);
  assert.strictEqual(res.status, 200);
  assert.ok((res.headers.get("content-type") || "").indexOf("pdf") !== -1);
  const buf = Buffer.from(await res.arrayBuffer());
  assert.ok(buf.slice(0, 5).toString() === "%PDF-");
  server.close();

  const floor = fs.readFileSync(path.join(__dirname, "../index.html"), "utf8");
  assert.ok(floor.indexOf("Profiles Used") === -1);
  assert.ok(floor.indexOf("function powderDimsFromOrder") !== -1);
  assert.ok(floor.indexOf("window.open(res.url") !== -1);
  assert.ok(floor.indexOf(" Paint Shop.") !== -1);
  assert.ok(floor.indexOf("loadDashboard()") !== -1);

  console.log("powder-list.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
