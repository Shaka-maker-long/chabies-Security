"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-delivery-pod-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const db = require("./db");
const staff = require("./staff");
const delivery = require("./delivery-pod");
const { callShopFunction, ALLOWED, clearShopCache } = require("./gas");
const sharp = require("sharp");

initWorkbook();

staff.upsertUser({
  name: "Siya",
  access: "Admin",
  role: "Production Manager",
  password: "s",
  tasks: ["Quality Control"]
});
staff.upsertUser({
  name: "Lebo",
  access: "Production",
  role: "Driver",
  password: "van",
  tasks: ["Delivery"]
});
staff.upsertUser({
  name: "Willard",
  access: "Production",
  role: "Welding",
  password: "1234",
  tasks: ["Welding"]
});

const lebo = staff.listUsers().find((u) => u.name === "Lebo");
assert.ok(delivery.isDriverProfile(lebo), "Driver job title is a driver");
assert.ok(lebo.tasks.indexOf("Delivery") !== -1, "Driver title ticks Delivery");
assert.ok(delivery.canSubmitPod(lebo));
assert.ok(delivery.canLoadTruck(staff.listUsers().find((u) => u.name === "Siya")));
assert.ok(!delivery.canLoadTruck(staff.listUsers().find((u) => u.name === "Willard")));

db.upsertOrder({
  order_number: "S260401 A",
  status: "Ready for Delivery",
  product: "Air Chair",
  client_name: "Naledi Botha",
  address: "12 Loop Street",
  city: "Cape Town",
  province: "Western Cape",
  assigned_operator: ""
});
db.upsertOrder({
  order_number: "S260401 B",
  status: "Ready for Delivery",
  product: "Air Chair",
  client_name: "Naledi Botha",
  address: "12 Loop Street",
  city: "Cape Town",
  province: "Western Cape",
  assigned_operator: ""
});
db.upsertOrder({
  order_number: "S260402",
  status: "Ready for Delivery",
  product: "Air Chair",
  client_name: "Johan",
  address: "88 Beach Road",
  city: "Hermanus",
  province: "Western Cape",
  assigned_operator: ""
});
db.upsertOrder({
  order_number: "S260403",
  status: "Welding",
  product: "Air Chair",
  client_name: "Skip",
  address: "Sandton",
  assigned_operator: "Willard"
});

assert.ok(ALLOWED.has("loadOrdersOnTruck"));
assert.ok(ALLOWED.has("submitDeliveryPod"));

(async function main() {
  const jpeg = await sharp({
    create: { width: 80, height: 60, channels: 3, background: { r: 40, g: 80, b: 40 } }
  }).jpeg({ quality: 70 }).toBuffer();
  const jpegB64 = jpeg.toString("base64");
  const pngDataUrl = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

  const loaded = await callShopFunction("loadOrdersOnTruck", [["S260401 A"], "Siya"]);
  assert.ok(loaded.count >= 2, "split units load together: " + loaded.count);
  const afterA = db.listOrders().find((o) => db.formatOrderId(o.order_number) === "S260401 A");
  const afterB = db.listOrders().find((o) => db.formatOrderId(o.order_number) === "S260401 B");
  assert.strictEqual(String(afterA.status), "Out for Delivery");
  assert.strictEqual(String(afterB.status), "Out for Delivery");
  assert.strictEqual(String(afterA.assigned_operator), "Siya");

  delivery.loadOnTruck(["S260402"], "Lebo");
  const hermanus = db.listOrders().find((o) => o.order_number === "S260402");
  assert.strictEqual(String(hermanus.status), "Out for Delivery");

  let blocked = null;
  try {
    delivery.loadOnTruck(["S260403"], "Willard");
  } catch (e) {
    blocked = e;
  }
  assert.ok(blocked, "welding orders cannot load");

  const route = await delivery.buildRoute({
    now: new Date("2026-10-06T08:00:00+02:00"),
    geocodeFn: async (q) => delivery.defaultGeocode(q)
  });
  assert.ok(route.stops.length >= 2, "route has loaded stops");
  assert.ok(route.stops.every((s) => s.eta_label), "each stop has an ETA");
  assert.ok(route.stops[0].minutes >= 4);

  clearShopCache();
  const listed = await callShopFunction("listDeliveryRun", []);
  assert.ok((listed.orders || []).length >= 3);

  const pod = await delivery.submitPod({
    order_number: "S260401 A",
    client_is_receiver: true,
    photos: [{ name: "unit.jpg", mime: "image/jpeg", data: jpegB64 }],
    signature: pngDataUrl,
    lat: -33.9249,
    lng: 18.4241,
    delivered_at: "2026-10-06T10:15:00.000Z",
    rating_delivery: 5,
    rating_sales: 4,
    rating_craft: 5,
    comments: "Left with the client."
  }, "Lebo");
  assert.ok(pod.id);
  assert.ok(pod.url.indexOf("/api/delivery-forms/") === 0);
  const deliveredA = db.listOrders().find((o) => db.formatOrderId(o.order_number) === "S260401 A");
  const deliveredB = db.listOrders().find((o) => db.formatOrderId(o.order_number) === "S260401 B");
  assert.strictEqual(String(deliveredA.status), "Delivered");
  assert.strictEqual(String(deliveredB.status), "Delivered");

  const forms = delivery.listForms();
  assert.strictEqual(forms.length, 1);
  assert.strictEqual(forms[0].receiver_name, "Naledi Botha");
  const pdf = delivery.readPdf(pod.id);
  assert.ok(pdf && pdf.buffer && pdf.buffer.length > 400);
  const pdfText = pdf.buffer.toString("latin1");
  assert.ok(/DELIVERY/.test(pdfText) || /STUDIO/.test(pdfText), "PDF letterhead");

  let missingPhoto = null;
  try {
    await delivery.submitPod({
      order_number: "S260402",
      client_is_receiver: true,
      photos: [],
      signature: pngDataUrl
    }, "Lebo");
  } catch (e) {
    missingPhoto = e;
  }
  assert.ok(missingPhoto, "POD needs a photo");

  const other = await delivery.submitPod({
    order_number: "S260402",
    client_is_receiver: false,
    receiver_name: "Gate guard",
    photos: [{ name: "door.jpg", mime: "image/jpeg", data: jpegB64 }],
    signature: pngDataUrl,
    lat: -34.4187,
    lng: 19.2345
  }, "Lebo");
  assert.ok(other.id);
  const hermanusDone = db.listOrders().find((o) => o.order_number === "S260402");
  assert.strictEqual(String(hermanusDone.status), "Delivered");
  const listedForms = delivery.listForms();
  assert.strictEqual(listedForms[0].receiver_name, "Gate guard");
  assert.strictEqual(listedForms.length, 2);

  console.log("delivery-pod.test.js ok", { splits: loaded.count, forms: listedForms.length });
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
