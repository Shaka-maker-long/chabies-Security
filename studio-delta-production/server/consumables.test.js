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

const po = consumables.createPurchase({
  supplier: "Wellington Hardware",
  note: "Weekly top-up",
  lines: [
    { itemId: hinge.id, qty: "50" },
    { itemId: screw.id, qty: "4" }
  ]
}, "Office Boss");
assert.strictEqual(po.status, "Ordered");
assert.strictEqual(po.lines.length, 2);
assert.strictEqual(consumables.snapshot().items.find((row) => row.id === hinge.id).stock, 14, "ordering does not add stock yet");

const received = consumables.receivePurchase(po.id, "Shaka");
assert.strictEqual(received.status, "Received");
assert.strictEqual(consumables.snapshot().items.find((row) => row.id === hinge.id).stock, 64);
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

const openPo = consumables.createPurchase({
  lines: [{ itemId: hinge.id, qty: "10" }]
}, "Office Boss");
consumables.cancelPurchase(openPo.id, "Office Boss");
assert.strictEqual(consumables.loadStore().purchases.find((row) => row.id === openPo.id).status, "Cancelled");
assert.strictEqual(consumables.snapshot().items.find((row) => row.id === hinge.id).stock, 64, "cancel does not change stock");

const snap = consumables.snapshot();
assert.ok(snap.units.indexOf("pcs") !== -1);
assert.ok(snap.movements.some((row) => row.type === "usage" && row.orderNumber === "S260100"));
assert.ok(snap.movements.some((row) => row.type === "purchase_receive"));
assert.ok(snap.items.every((row) => !/steel|glass|wood/i.test(row.name)));

const sticker = consumables.upsertItem({ name: "Powder colour sticker", unit: "pcs", minThreshold: 0 }, "Office Boss");
assert.strictEqual(sticker.stock, 0);
assert.strictEqual(sticker.low, true);
consumables.deleteItem(sticker.id);
assert.ok(!consumables.snapshot().items.some((row) => row.id === sticker.id));

let deleteFailed = false;
try {
  consumables.deleteItem(hinge.id);
} catch (e) {
  deleteFailed = /down to 0/i.test(e.message);
}
assert.ok(deleteFailed, "cannot delete an item that still has stock");

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

  server.close();
  console.log("consumables.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
