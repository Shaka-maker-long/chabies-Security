"use strict";

const assert = require("assert");
const fs = require("fs");
const os = require("os");
const path = require("path");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-inv-proto-"));
process.env.DATA_DIR = dir;

const proto = require("./inventory-prototype");

async function main() {
  proto.resetPrototype();
  let snap = proto.snapshot();
  assert.strictEqual(snap.materials.length, 2);
  assert.strictEqual(snap.materials[0].onHand, 0);

  const purchase = await proto.createPurchase({
    supplier: "Test",
    lines: [
      { code: "A", qty: 20 },
      { code: "B", qty: 15 }
    ]
  });
  assert.ok(/^PROTO-PO-/.test(purchase.number));
  assert.strictEqual(purchase.status, "ordered");
  assert.ok(purchase.hasPoPdf, "PO PDF created on purchase");
  assert.ok(purchase.hasLabelsPdf, "QR labels PDF created on purchase (both docs together)");
  assert.ok(purchase.poPdfUrl.indexOf("/po.pdf") !== -1);
  assert.ok(purchase.lines[0].qrPdfUrl.indexOf("/qr/") !== -1);
  assert.ok(purchase.lines[1].qrPdfUrl.indexOf("/qr/") !== -1);
  assert.ok(fs.existsSync(path.join(dir, "inventory-prototype-pdfs", purchase.id + "-po.pdf")));
  assert.ok(fs.statSync(path.join(dir, "inventory-prototype-pdfs", purchase.id + "-po.pdf")).size > 500);
  assert.ok(fs.existsSync(path.join(dir, "inventory-prototype-pdfs", purchase.id + "-qr-mat-a.pdf")));
  assert.ok(fs.existsSync(path.join(dir, "inventory-prototype-pdfs", purchase.id + "-qr-mat-b.pdf")));

  const received = await proto.receivePurchase(purchase.id);
  assert.strictEqual(received.purchase.status, "received");
  assert.strictEqual(received.labels.length, 2);
  assert.strictEqual(received.labels[0].qr, "SDINV|MAT|A");
  assert.strictEqual(received.labels[1].qr, "SDINV|MAT|B");
  assert.ok(String(received.labels[0].qrImage || "").indexOf("data:image/png") === 0, "server must render QR PNG");
  assert.ok(received.purchase.hasLabelsPdf, "QR labels PDF still available after receive");
  assert.ok(Math.abs(proto.cmToPt(5) - ((5 * 72) / 2.54)) < 0.01);

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
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
