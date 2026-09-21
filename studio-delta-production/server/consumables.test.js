const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");
const http = require("http");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-cons-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const staff = require("./staff");
const { mountOffice } = require("./office");
const consumables = require("./consumables");

initWorkbook();
staff.upsertUser({
  name: "Office Boss",
  access: "Admin",
  role: "Admin",
  password: "admin",
  seeDebtors: "Yes"
});

const catalog = require("./consumables-catalog");
assert.ok(catalog.length >= 150);

const hinge = consumables.upsertItem({
  name: "  Bullet hinge  ",
  unit: "pcs",
  minThreshold: "20",
  openingStock: "40"
}, "Office Boss");
assert.strictEqual(hinge.name, "Bullet hinge");
assert.strictEqual(hinge.stock, 40);
assert.strictEqual(hinge.minThreshold, 20);
assert.strictEqual(hinge.low, false);

const windowCleaner = consumables.upsertItem({
  name: "WINDOW CLEANER",
  unit: "pcs",
  openingStock: "99"
}, "Office Boss");
assert.strictEqual(windowCleaner.stock, 99);

let dupFailed = false;
try {
  consumables.upsertItem({ name: "bullet hinge", minThreshold: 1 }, "Office Boss");
} catch (e) {
  dupFailed = /already/i.test(e.message);
}
assert.ok(dupFailed, "duplicate names are rejected");

const used = consumables.logUsage({
  itemId: hinge.id,
  qty: "6",
  orderNumber: "S260100",
  note: "Assembly"
}, "Admire");
assert.strictEqual(used.stock, 34);
assert.strictEqual(used.low, false);

let overFailed = false;
try {
  consumables.logUsage({ itemId: hinge.id, qty: "100" }, "Admire");
} catch (e) {
  overFailed = /on hand/i.test(e.message);
}
assert.ok(overFailed, "cannot use more than on-hand");

consumables.logUsage({ itemId: hinge.id, qty: "20", orderNumber: "S260101" }, "Admire");
const afterUse = consumables.snapshot().items.find((row) => row.id === hinge.id);
assert.strictEqual(afterUse.stock, 14);
assert.strictEqual(afterUse.low, true);
assert.strictEqual(afterUse.status, "Low");

const screw = consumables.upsertItem({
  name: "4mm chipboard screw",
  unit: "box",
  minThreshold: "2",
  openingStock: "1"
}, "Office Boss");
assert.strictEqual(screw.low, true);

const multiUsed = consumables.logUsage({
  orderNumber: "General",
  worker: "Admire",
  note: "Floor batch",
  lines: [
    { itemId: hinge.id, qty: "2" },
    { itemId: screw.id, qty: "1" }
  ]
}, "Admire");
assert.ok(Array.isArray(multiUsed));
assert.strictEqual(multiUsed.length, 2);
assert.strictEqual(multiUsed.find((row) => row.id === hinge.id).stock, 12);
assert.strictEqual(multiUsed.find((row) => row.id === screw.id).stock, 0);
let multiDupFailed = false;
try {
  consumables.logUsage({
    orderNumber: "General",
    worker: "Admire",
    lines: [
      { itemId: hinge.id, qty: "1" },
      { itemId: hinge.id, qty: "1" }
    ]
  }, "Admire");
} catch (e) {
  multiDupFailed = /twice/i.test(e.message);
}
assert.ok(multiDupFailed, "duplicate items in one usage batch are rejected");
consumables.receiveStock({ itemId: hinge.id, qty: "2", note: "Restore after batch test" }, "Siya");
consumables.receiveStock({ itemId: screw.id, qty: "1", note: "Restore after batch test" }, "Siya");
assert.strictEqual(consumables.snapshot().items.find((row) => row.id === hinge.id).stock, 14);
assert.strictEqual(consumables.snapshot().items.find((row) => row.id === screw.id).stock, 1);

(async function purchaseFlow() {
const po = await consumables.createPurchase({
  supplier: "Wellington Hardware",
  note: "Weekly top-up",
  lines: [
    { itemId: hinge.id, qty: "50" },
    { itemId: screw.id, qty: "4" }
  ]
}, "Office Boss");
assert.strictEqual(po.status, "Ordered");
assert.strictEqual(po.lines.length, 2);
assert.ok(po.number);
assert.ok(po.hasPdf);
assert.ok(po.pdfUrl);
const pdf = consumables.readPurchasePdf(po.id);
assert.ok(pdf && pdf.buffer && pdf.buffer.length > 100);
const orderedSnap = consumables.snapshot();
const orderedHinge = orderedSnap.items.find((row) => row.id === hinge.id);
assert.strictEqual(orderedHinge.stock, 14, "ordering does not add stock yet");
assert.strictEqual(orderedHinge.orderedQty, 50);
assert.strictEqual(orderedHinge.orderedLabel, "50");
assert.strictEqual(orderedSnap.items.find((row) => row.id === screw.id).orderedQty, 4);
assert.ok(orderedSnap.activeOrderNumbers);
assert.ok(Array.isArray(orderedSnap.productionWorkers));

const received = consumables.receivePurchase(po.id, "Shaka");
assert.strictEqual(received.status, "Received");
assert.strictEqual(consumables.snapshot().items.find((row) => row.id === hinge.id).stock, 64);
assert.strictEqual(consumables.snapshot().items.find((row) => row.id === hinge.id).orderedQty, 0, "received PO is no longer ordered");
assert.strictEqual(consumables.snapshot().items.find((row) => row.id === screw.id).stock, 5);
assert.strictEqual(consumables.snapshot().items.find((row) => row.id === screw.id).low, false);

let twiceFailed = false;
try {
  consumables.receivePurchase(po.id, "Shaka");
} catch (e) {
  twiceFailed = /already received/i.test(e.message);
}
assert.ok(twiceFailed);

consumables.receiveStock({ itemId: screw.id, qty: "1", supplier: "Walk-in", note: "Emergency box" }, "Siya");
assert.strictEqual(consumables.snapshot().items.find((row) => row.id === screw.id).stock, 6);

const counted = consumables.countStock({
  itemId: screw.id,
  stock: "5",
  note: "Box was short"
}, "Siya");
assert.strictEqual(counted.stock, 5);

const openPo = await consumables.createPurchase({
  lines: [{ itemId: hinge.id, qty: "10" }]
}, "Office Boss");
consumables.cancelPurchase(openPo.id, "Office Boss");
assert.strictEqual(consumables.loadStore().purchases.find((row) => row.id === openPo.id).status, "Cancelled");
assert.strictEqual(consumables.snapshot().items.find((row) => row.id === hinge.id).stock, 64, "cancel does not change stock");

const snap = consumables.snapshot();
assert.ok(snap.units.indexOf("pcs") !== -1);
assert.ok(snap.movements.some((row) => row.type === "usage" && row.orderNumber === "S260100"));
assert.ok(snap.movements.some((row) => row.type === "purchase_receive"));
assert.ok(snap.itemCount >= catalog.length);
const catalogHinges = snap.items.find((row) => row.name === "BULLET HINGES - 12mm x 70mm");
assert.ok(catalogHinges);
assert.strictEqual(catalogHinges.stock, 75);
assert.strictEqual(catalogHinges.minThreshold, 0);
assert.strictEqual(catalogHinges.low, false, "ROP 0 does not flag low stock");
const bumper = snap.items.find((row) => row.name === "BUMPER SPRAY");
assert.ok(bumper);
assert.strictEqual(bumper.stock, 13);
assert.strictEqual(bumper.unitPrice, 272.27);
assert.ok(bumper.priceLabel.indexOf("272.27") !== -1);
const ears = snap.items.find((row) => row.name === "RE-USABLE EAR PLUGS");
assert.ok(ears);
assert.strictEqual(ears.stock, 63);
assert.strictEqual(ears.minThreshold, 10);
assert.strictEqual(ears.low, false);
const drywall = snap.items.find((row) => row.name === "35mm DRYWALL SCREWS");
assert.ok(drywall);
assert.strictEqual(drywall.stock, 152);
assert.strictEqual(drywall.minThreshold, 50);
assert.strictEqual(drywall.low, false);
const goggles = snap.items.find((row) => row.name === "CLEAR SAFETY GOGGLES");
assert.ok(goggles);
assert.strictEqual(goggles.stock, 0);
assert.strictEqual(goggles.minThreshold, 0);
assert.strictEqual(goggles.low, false);
const seededWindow = snap.items.find((row) => row.name === "WINDOW CLEANER");
assert.ok(seededWindow);
assert.strictEqual(seededWindow.stock, 99, "catalog seed does not overwrite existing stock");
assert.strictEqual(snap.items.filter((row) => row.name === "WINDOW CLEANER").length, 1);

const sticker = consumables.upsertItem({ name: "Powder colour sticker", unit: "pcs", minThreshold: 0 }, "Office Boss");
assert.strictEqual(sticker.stock, 0);
assert.strictEqual(sticker.low, false);
consumables.deleteItem(sticker.id);
assert.ok(!consumables.snapshot().items.some((row) => row.id === sticker.id));

let deleteFailed = false;
try {
  consumables.deleteItem(hinge.id);
} catch (e) {
  deleteFailed = /down to 0/i.test(e.message);
}
assert.ok(deleteFailed, "cannot delete an item that still has stock");
})().then(function () {
return (async function main() {
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

  const listed = await api("/api/office/consumables", { headers });
  assert.strictEqual(listed.json.ok, true);
  assert.ok(listed.json.items.some((row) => row.name === "Bullet hinge"));
  assert.ok(listed.json.itemCount >= 2);
  assert.ok((listed.json.movements || []).some((row) => row.type === "usage"));

  const usedApi = await api("/api/office/consumables/use", {
    method: "POST",
    headers,
    body: JSON.stringify({ itemId: hinge.id, qty: 2, orderNumber: "S260200" })
  });
  assert.strictEqual(usedApi.json.ok, true, JSON.stringify(usedApi.json));
  assert.strictEqual(usedApi.json.item.stock, 62);

  const batchApi = await api("/api/office/consumables/use", {
    method: "POST",
    headers,
    body: JSON.stringify({
      orderNumber: "General",
      worker: "Admire",
      note: "API batch",
      lines: [
        { itemId: hinge.id, qty: 1 },
        { itemId: screw.id, qty: 1 }
      ]
    })
  });
  assert.strictEqual(batchApi.json.ok, true, JSON.stringify(batchApi.json));
  assert.ok(Array.isArray(batchApi.json.usedItems));
  assert.strictEqual(batchApi.json.usedItems.length, 2);
  assert.strictEqual(batchApi.json.usedItems.find((row) => row.id === hinge.id).stock, 61);
  assert.ok(Array.isArray(batchApi.json.items) && batchApi.json.items.length > 2, "snapshot items stay intact");

  server.close();
  console.log("consumables.test.js ok");
})();
}).catch((err) => {
  console.error(err);
  process.exit(1);
});
