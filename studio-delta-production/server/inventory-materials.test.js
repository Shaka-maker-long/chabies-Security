const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");
const http = require("http");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-invmat-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const staff = require("./staff");
const { mountOffice } = require("./office");
const steelRates = require("./steel-rates");
const glassRates = require("./glass-rates");
const inventory = require("./inventory-materials");

initWorkbook();
staff.upsertUser({
  name: "Office Boss",
  access: "Admin",
  role: "Admin",
  password: "admin",
  seeDebtors: "Yes"
});

steelRates.upsertRate({ type: "25x25 SHS", ratePerM: "45.50" });
let steel = inventory.snapshotSteel();
const shs = steel.items.find((row) => row.name === "25x25 SHS");
assert.ok(shs, "steel rates appear on the Steel inventory list");
assert.strictEqual(shs.stock, 0);
assert.strictEqual(shs.orderedQty, 0);
assert.strictEqual(shs.minThreshold, 0);
assert.strictEqual(shs.low, false, "ROP 0 does not flag steel as low");
assert.strictEqual(Number(shs.unitPrice), 45.5);

steel = inventory.upsertSteel({
  name: "25x25 SHS",
  stock: "12",
  orderedQty: "8",
  minThreshold: "5",
  unitPrice: "50"
});
const saved = steel.items.find((row) => row.name === "25x25 SHS");
assert.strictEqual(saved.stock, 12);
assert.strictEqual(saved.orderedQty, 8);
assert.strictEqual(saved.minThreshold, 5);
assert.strictEqual(Number(saved.unitPrice), 50);
assert.strictEqual(saved.low, false);
assert.ok(saved.priceLabel.indexOf("50.00") !== -1);

steel = inventory.upsertSteel({ name: "25x25 SHS", stock: "5", minThreshold: "5" });
const lowShs = steel.items.find((row) => row.name === "25x25 SHS");
assert.strictEqual(lowShs.stock, 5);
assert.strictEqual(lowShs.low, true);
assert.strictEqual(steel.lowCount, 1);

inventory.upsertSteel({ name: "50x50 SHS", stock: "0", minThreshold: "0" });
const emptyShs = inventory.snapshotSteel().items.find((row) => row.name === "50x50 SHS");
assert.ok(emptyShs);
assert.strictEqual(emptyShs.low, false);

glassRates.upsertRate({ type: "Reeded", thickness: "6mm", ratePerM2: "450" });
const sheet = getBook().getSheetByName("Glass_To_Order");
sheet.appendRow(["g-inv", new Date(), "S260199", "Nomsa", "Door", "Reeded", "6mm", 1800, 500, 4, "To order"]);
persistWorkbook();

let glass = inventory.snapshotGlass();
const reeded = glass.items.find((row) => row.name === "Reeded 6mm");
assert.ok(reeded, "glass rates and To order lines appear on Glass inventory");
assert.strictEqual(reeded.orderedQty, 4, "ordered stock comes from outstanding To order quantity");
assert.strictEqual(Number(reeded.unitPrice), 450);
assert.strictEqual(reeded.stock, 0);
assert.strictEqual(reeded.low, false);

glass = inventory.upsertGlass({
  name: "Reeded 6mm",
  stock: "2",
  orderedQty: "1",
  minThreshold: "3",
  unitPrice: "410"
});
const glassSaved = glass.items.find((row) => row.name === "Reeded 6mm");
assert.strictEqual(glassSaved.stock, 2);
assert.strictEqual(glassSaved.orderedQty, 1, "typed ordered stock overrides To order");
assert.strictEqual(glassSaved.minThreshold, 3);
assert.strictEqual(Number(glassSaved.unitPrice), 410);
assert.strictEqual(glassSaved.low, true);

(async function main() {
  const app = express();
  app.use(express.json({ limit: "2mb" }));
  mountOffice(app);
  const server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;

  async function api(pathname, opts) {
    const r = await fetch("http://127.0.0.1:" + port + pathname, opts);
    const j = await r.json();
    return { status: r.status, json: j };
  }

  const login = await api("/api/office/login", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ name: "Office Boss", password: "admin" })
  });
  assert.strictEqual(login.json.ok, true, JSON.stringify(login.json));
  const token = login.json.token;
  const headers = { "content-type": "application/json", "x-sd-token": token };

  const listedSteel = await api("/api/office/inventory/steel", { headers });
  assert.strictEqual(listedSteel.json.ok, true);
  assert.ok((listedSteel.json.items || []).some((row) => row.name === "25x25 SHS"));

  const listedGlass = await api("/api/office/inventory/glass", { headers });
  assert.strictEqual(listedGlass.json.ok, true);
  assert.ok((listedGlass.json.items || []).some((row) => row.name === "Reeded 6mm"));

  const posted = await api("/api/office/inventory/steel", {
    method: "POST",
    headers,
    body: JSON.stringify({ name: "25x25 SHS", stock: "9", orderedQty: "2", minThreshold: "4" })
  });
  assert.strictEqual(posted.json.ok, true, JSON.stringify(posted.json));
  const after = (posted.json.items || []).find((row) => row.name === "25x25 SHS");
  assert.strictEqual(after.stock, 9);
  assert.strictEqual(after.orderedQty, 2);

  server.close();
  console.log("inventory-materials.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
