"use strict";

const assert = require("assert");
const fs = require("fs");
const os = require("os");
const path = require("path");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-inv-proto-"));
process.env.DATA_DIR = dir;

const proto = require("./inventory-prototype");

proto.resetPrototype();
let snap = proto.snapshot();
assert.strictEqual(snap.materials.length, 2);
assert.strictEqual(snap.materials[0].onHand, 0);

const purchase = proto.createPurchase({
  supplier: "Test",
  lines: [
    { code: "A", qty: 20 },
    { code: "B", qty: 15 }
  ]
});
assert.ok(/^PROTO-PO-/.test(purchase.number));
assert.strictEqual(purchase.status, "ordered");

const received = proto.receivePurchase(purchase.id);
assert.strictEqual(received.purchase.status, "received");
assert.strictEqual(received.labels.length, 2);
assert.strictEqual(received.labels[0].qr, "SDINV|MAT|A");
assert.strictEqual(received.labels[1].qr, "SDINV|MAT|B");

snap = proto.snapshot();
assert.strictEqual(snap.materials.find((m) => m.code === "A").onHand, 20);
assert.strictEqual(snap.materials.find((m) => m.code === "B").onHand, 15);

proto.setActiveJob({ workerId: "w1", orderNumber: "S260301" });
const alloc = proto.allocate({ workerId: "w1", qr: "SDINV|MAT|A", qty: 3 });
assert.strictEqual(alloc.allocation.orderNumber, "S260301");
assert.strictEqual(alloc.material.onHand, 17);
assert.strictEqual(alloc.material.wip, 3);

snap = proto.snapshot();
assert.strictEqual(snap.materials.find((m) => m.code === "A").onHand, 17);
assert.strictEqual(snap.materials.find((m) => m.code === "A").wip, 3);
assert.ok(snap.wipByOrder.some((w) => w.orderNumber === "S260301"));

assert.throws(() => proto.allocate({ workerId: "w1", qr: "A", qty: 100 }), /available/i);
assert.strictEqual(proto.parseQrPayload("SDINV|MAT|B"), "B");
assert.strictEqual(proto.parseQrPayload("a"), "A");

console.log("inventory-prototype.test.js ok");
