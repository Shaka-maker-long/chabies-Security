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
const db = require("./db");
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
assert.strictEqual(shs.wipStock, 0);
assert.strictEqual(shs.orderedQty, 0);
assert.strictEqual(shs.totalPurchased, 0);
assert.strictEqual(shs.totalUsed, 0);
assert.strictEqual(shs.minThreshold, 0);
assert.strictEqual(shs.buyUnit, "length");
assert.strictEqual(shs.lengthM, 6);
assert.strictEqual(shs.low, false, "ROP 0 does not flag steel as low");
assert.strictEqual(Number(shs.unitPrice), 273, "default unit price is rate/m × 6 m length");
assert.strictEqual(steel.autoDeductEnabled, false, "auto-deduct stays off until on-hand stock is loaded");
assert.strictEqual(inventory.metresToLengths(6), 1);
assert.strictEqual(inventory.metresToLengths(12), 2);
assert.ok(inventory.isPlateName("PLATE - 1.6X1500X3000 PLATE"));
assert.ok(inventory.isPlateName("LATE - 1.2X1220X2450 PLATE"));
assert.ok(!inventory.isPlateName("25x25 SHS"));
assert.strictEqual(inventory.plateSheetAreaM2("LATE - 1.2X1220X2450 PLATE"), 2.989);
assert.strictEqual(inventory.steelBuyUnit("LATE - 1.2X1220X2450 PLATE"), "sheet");
assert.strictEqual(inventory.allocateSteelToWip("25x25 SHS", 1).skipped, true);

steel = inventory.upsertSteel({
  name: "25x25 SHS",
  stock: "12",
  orderedQty: "8",
  minThreshold: "5",
  unitPrice: "50",
  totalPurchased: "20",
  totalUsed: "3"
});
const saved = steel.items.find((row) => row.name === "25x25 SHS");
assert.strictEqual(saved.stock, 12);
assert.strictEqual(saved.orderedQty, 8);
assert.strictEqual(saved.minThreshold, 5);
assert.strictEqual(Number(saved.unitPrice), 50);
assert.strictEqual(saved.totalPurchased, 20);
assert.strictEqual(saved.totalUsed, 3);
assert.strictEqual(saved.valueInStock, 600);
assert.strictEqual(saved.valueInWip, 0);
assert.strictEqual(saved.low, false);
assert.ok(saved.priceLabel.indexOf("50.00") !== -1);

steel = inventory.upsertSteel({
  name: "25x25 SHS",
  previousName: "25x25 SHS",
  buyUnit: "length",
  unitPrice: "50"
});
const withUnit = steel.items.find((row) => row.name === "25x25 SHS");
assert.strictEqual(withUnit.buyUnit, "length");

db.upsertOrder({
  order_number: "S-STEEL-WIP",
  product: "Gate",
  status: "Welding",
  price_excl_vat: "1000.00"
});
db.upsertOrder({
  order_number: "S-STEEL-DONE",
  product: "Gate",
  status: "Delivered",
  price_excl_vat: "1000.00"
});
const start = new Date("2026-09-09T06:00:00.000Z");
getBook().getSheetByName("Steel_Usage").appendRow([
  start, "S-STEEL-WIP", "Willard", "Welding", "25x25 SHS", 12
]);
getBook().getSheetByName("Steel_Usage").appendRow([
  start, "S-STEEL-DONE", "Willard", "Welding", "25x25 SHS", 6
]);
persistWorkbook();

steel = inventory.snapshotSteel();
const fromUsage = steel.items.find((row) => row.name === "25x25 SHS");
assert.strictEqual(fromUsage.wipStock, 2, "12 m on undelivered order → 2 lengths in WIP Qty");
assert.strictEqual(fromUsage.valueInWip, 100, "WIP value = WIP Qty × unit price");
assert.strictEqual(fromUsage.usedFromUsage, 1, "6 m on Delivered order → 1 length used from usage");
assert.strictEqual(fromUsage.totalUsed, 3, "stored totalUsed override still wins over usage");

steel = inventory.upsertSteel({
  name: "25x25 SHS",
  previousName: "25x25 SHS",
  stock: "12",
  unitPrice: "50"
});
// clear stored totalUsed by writing usage-default path: delete field via fresh row logic —
// re-save without totalUsed keeps prior stored 3. Explicitly set to usage by saving usedFromUsage.
steel = inventory.upsertSteel({
  name: "25x25 SHS",
  previousName: "25x25 SHS",
  totalUsed: "1"
});
const usedSynced = steel.items.find((row) => row.name === "25x25 SHS");
assert.strictEqual(usedSynced.totalUsed, 1);
assert.strictEqual(usedSynced.wipQty, 2);

steel = inventory.upsertSteel({
  name: "LATE - 1.2X1220X2450 PLATE",
  stock: "4",
  orderedQty: "2",
  unitPrice: "160.35",
  totalPurchased: "4",
  buyUnit: "sheet"
});
const plate = steel.items.find((row) => row.name === "LATE - 1.2X1220X2450 PLATE");
assert.ok(plate);
assert.strictEqual(plate.buyUnit, "sheet");
assert.strictEqual(plate.sheetAreaM2, 2.989);
assert.strictEqual(Number(plate.unitPrice), 160.35);
assert.strictEqual(plate.valueInStock, 641.4);

steel = inventory.upsertSteel({
  name: "25x25 SHS renamed",
  previousName: "25x25 SHS",
  stock: "5",
  minThreshold: "5"
});
const renamed = steel.items.find((row) => row.name === "25x25 SHS renamed");
assert.ok(renamed, "steel name can be renamed via previousName");
assert.strictEqual(renamed.stock, 5);
assert.strictEqual(renamed.low, true);
assert.ok(steel.lowCount >= 1);

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
