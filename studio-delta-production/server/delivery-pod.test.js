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

db.upsertOrder({
  order_number: "S260404",
  status: "Ready for Delivery",
  product: "Air Chair",
  client_name: "Thandi",
  address: "9 Rivonia Road",
  city: "Sandton",
  province: "Gauteng",
  assigned_operator: ""
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
  delivery.loadOnTruck(["S260404"], "Lebo");

  const capeDrop = delivery.decorateOrder(db.listOrders().find((o) => db.formatOrderId(o.order_number) === "S260401 A"));
  assert.strictEqual(capeDrop.drop_kind, "third_party");
  assert.ok(capeDrop.drop_address.indexOf("Milkyway") !== -1, "non-Gauteng drops at Frankenwald");
  assert.ok(capeDrop.client_address.indexOf("Cape Town") !== -1);
  const gautengDrop = delivery.decorateOrder(db.listOrders().find((o) => o.order_number === "S260404"));
  assert.strictEqual(gautengDrop.drop_kind, "client");
  assert.ok(gautengDrop.drop_address.indexOf("Rivonia") !== -1);

  const depotPin = delivery.defaultGeocode(delivery.THIRD_PARTY_DEPOT.full_address);
  assert.ok(Math.abs(depotPin.lat - delivery.THIRD_PARTY_DEPOT.lat) < 0.01);
  assert.ok(String(delivery.FACTORY.address || "").indexOf("Derdepoort") !== -1);
  assert.ok(delivery.FACTORY.lat < -25.7 && delivery.FACTORY.lat > -25.8, "factory is Silverton / Pretoria");

  const route = await delivery.buildRoute({
    now: new Date("2026-10-06T08:00:00+02:00"),
    geocodeFn: async (q) => delivery.defaultGeocode(q),
    driveFn: null
  });
  assert.strictEqual(route.stops.length, 2, "Gauteng client stop plus one 3rd party depot: " + route.stops.length);
  const depotStop = route.stops.find((s) => s.drop_kind === "third_party");
  const clientStop = route.stops.find((s) => s.drop_kind === "client");
  assert.ok(depotStop, "out-of-Gauteng units share the Frankenwald stop");
  assert.ok(clientStop, "Gauteng still goes to the client");
  assert.ok(depotStop.orders.length >= 3, "Cape Town splits and Hermanus share the depot");
  assert.ok(depotStop.address.indexOf("Milkyway") !== -1);
  assert.ok(Math.abs(route.origin.lat - delivery.FACTORY.lat) < 0.001, "route starts at Studio Delta");
  assert.ok(String(route.origin.address || "").indexOf("Derdepoort") !== -1);
  assert.ok(route.stops.every((s) => s.eta_label), "each stop has an ETA");
  assert.ok(route.stops[0].minutes >= 4);
  assert.ok(depotStop.waze_url.indexOf("waze.com") !== -1);
  assert.strictEqual(route.on_roads, false);

  const roaded = await delivery.buildRoute({
    now: new Date("2026-10-06T08:00:00+02:00"),
    geocodeFn: async (q) => delivery.defaultGeocode(q),
    driveFn: async (pts) => ({
      path: [[-25.725, 28.295], [-25.8, 28.22], [-26.067, 28.111]],
      legs: pts.slice(1).map(() => ({ km: 18.4, minutes: 22 }))
    })
  });
  assert.strictEqual(roaded.on_roads, true);
  assert.ok(roaded.path.length > 2);
  assert.strictEqual(roaded.stops[0].km, 18.4);
  assert.ok(roaded.stops[0].minutes >= 22);

  const parsed = delivery.parseOsrmRoute({
    code: "Ok",
    routes: [{
      geometry: { coordinates: [[28.29, -25.72], [28.11, -26.06]] },
      legs: [{ distance: 54200, duration: 2736 }]
    }]
  }, 2);
  assert.strictEqual(parsed.path[0][0], -25.72);
  assert.strictEqual(parsed.legs[0].km, 54.2);

  let blocked = null;
  try {
    delivery.loadOnTruck(["S260403"], "Willard");
  } catch (e) {
    blocked = e;
  }
  assert.ok(blocked, "welding orders cannot load");

  clearShopCache();
  const listed = await callShopFunction("listDeliveryRun", []);
  assert.ok((listed.orders || []).length >= 3);

  let missingName = null;
  try {
    await delivery.submitPod({
      order_numbers: ["S260401 A", "S260402"],
      batch_third_party: true,
      photos: [{ name: "unit.jpg", mime: "image/jpeg", data: jpegB64 }],
      signature: pngDataUrl
    }, "Lebo");
  } catch (e) {
    missingName = e;
  }
  assert.ok(missingName, "courier handover needs the depot receiver name");

  const pod = await delivery.submitPod({
    order_numbers: ["S260401 A", "S260401 B", "S260402"],
    batch_third_party: true,
    receiver_name: "Depot clerk",
    photos: [
      { name: "unit.jpg", mime: "image/jpeg", data: jpegB64 },
      { name: "extra.jpg", mime: "image/jpeg", data: jpegB64 }
    ],
    signature: pngDataUrl,
    lat: -26.0674,
    lng: 28.1112,
    delivered_at: "2026-10-06T10:15:00.000Z",
    rating_delivery: 5,
    rating_sales: 4,
    rating_craft: 5,
    comments: "Handed to Frankenwald together."
  }, "Lebo");
  assert.ok(pod.id);
  assert.ok(pod.url.indexOf("/api/delivery-forms/") === 0);
  assert.ok((pod.order_numbers || []).length >= 3, "all third-party units share one handover");
  const deliveredA = db.listOrders().find((o) => db.formatOrderId(o.order_number) === "S260401 A");
  const deliveredB = db.listOrders().find((o) => db.formatOrderId(o.order_number) === "S260401 B");
  const hermanusDone = db.listOrders().find((o) => o.order_number === "S260402");
  assert.strictEqual(String(deliveredA.status), "At couriers");
  assert.strictEqual(String(deliveredB.status), "At couriers");
  assert.strictEqual(String(hermanusDone.status), "At couriers");

  const forms = delivery.listForms();
  assert.strictEqual(forms.length, 1, "one form for the whole third-party batch");
  assert.strictEqual(forms[0].receiver_name, "Depot clerk");
  assert.strictEqual(forms[0].rating_delivery, 0, "courier handover has no ratings");
  assert.ok(String(forms[0].address || "").indexOf("Milkyway") !== -1, "PDF stop is the 3rd party depot");
  assert.ok(String(forms[0].client_address || "").indexOf("Cape Town") !== -1);
  assert.ok(String(forms[0].client_address || "").indexOf("Hermanus") !== -1);
  const pdf = delivery.readPdf(pod.id);
  assert.ok(pdf && pdf.buffer && pdf.buffer.slice(0, 4).toString() === "%PDF");
  const latin = pdf.buffer.toString("latin1");
  assert.ok(/DCTDecode/.test(latin), "JPEG photo must be embedded");
  const decoded = [];
  latin.replace(/<([0-9A-Fa-f]+)>/g, (_, hex) => {
    try { decoded.push(Buffer.from(hex, "hex").toString("latin1")); } catch (e) {}
    return "";
  });
  const pdfText = latin + "\n" + decoded.join("");
  assert.ok(pdfText.indexOf("STUDIO DELTA") !== -1 || /STUDIO/.test(pdfText), "delivery PDF letterhead");
  assert.ok(/COURIER HANDOVER/.test(pdfText), "third-party PDF is a courier handover");
  assert.ok(/At couriers/.test(pdfText), "third-party PDF shows At couriers");
  assert.ok(/COURIER DEPOT/.test(pdfText), "depot field is labelled");
  assert.ok(/CLIENT DESTINATION/.test(pdfText), "client destination is labelled");
  assert.ok(/Milkyway|Frankenwald/i.test(pdfText), "non-Gauteng PDF names the Frankenwald depot");
  assert.ok(!/DROPPED AT/i.test(pdfText), "PDF must not use informal dropped-at copy");
  assert.ok(/GPS PIN/.test(pdfText), "GPS is labelled as a pin, not raw LOCATION");
  assert.ok(!/\bRATINGS\b/.test(pdfText), "courier handover PDF has no ratings block");
  assert.ok(/S260401 A/.test(pdfText) && /S260402/.test(pdfText), "batched orders appear on the PDF");

  let missingPhoto = null;
  try {
    await delivery.submitPod({
      order_number: "S260404",
      client_is_receiver: true,
      photos: [],
      signature: pngDataUrl
    }, "Lebo");
  } catch (e) {
    missingPhoto = e;
  }
  assert.ok(missingPhoto, "POD needs a photo");

  const listedForms = delivery.listForms();
  assert.strictEqual(listedForms.length, 1);
  assert.strictEqual(listedForms[0].status, "At couriers");

  const gautengPod = await delivery.submitPod({
    order_number: "S260404",
    client_is_receiver: true,
    photos: [{ name: "door.jpg", mime: "image/jpeg", data: jpegB64 }],
    signature: pngDataUrl,
    lat: -26.1076,
    lng: 28.0567,
    delivered_at: "2026-10-06T12:00:00.000Z",
    rating_delivery: 4,
    rating_sales: 5,
    rating_craft: 3
  }, "Lebo");
  assert.ok(gautengPod.id);
  const gautengDone = db.listOrders().find((o) => o.order_number === "S260404");
  assert.strictEqual(String(gautengDone.status), "Delivered");
  const gautengPdf = delivery.readPdf(gautengPod.id);
  const gautengLatin = gautengPdf.buffer.toString("latin1");
  const gautengDecoded = [];
  gautengLatin.replace(/<([0-9A-Fa-f]+)>/g, (_, hex) => {
    try { gautengDecoded.push(Buffer.from(hex, "hex").toString("latin1")); } catch (e) {}
    return "";
  });
  const gautengText = gautengLatin + "\n" + gautengDecoded.join("");
  assert.ok(/DELIVERY NOTE/.test(gautengText), "Gauteng PDF is a delivery note");
  assert.ok(/Delivered/.test(gautengText));
  assert.ok(/DELIVERED TO/.test(gautengText));
  assert.ok(!/COURIER HANDOVER/.test(gautengText));
  assert.strictEqual(delivery.podStatusForDrop("third_party"), "At couriers");
  assert.strictEqual(delivery.podStatusForDrop("client"), "Delivered");
  assert.ok(/S/.test(delivery.formatGps(-25.76822, 28.26716)));

  db.upsertOrder({
    order_number: "S260405",
    status: "Out for Delivery",
    product: "Air Chair",
    client_name: "Ada",
    address: "1 Test Street",
    city: "Pretoria",
    province: "Gauteng",
    assigned_operator: "Lebo"
  });
  const offlineId = "offline-pod-1";
  const firstOffline = await delivery.submitPod({
    order_number: "S260405",
    client_is_receiver: true,
    client_submit_id: offlineId,
    photos: [{ name: "door.jpg", mime: "image/jpeg", data: jpegB64 }],
    signature: pngDataUrl,
    lat: -25.74,
    lng: 28.21
  }, "Lebo");
  assert.ok(firstOffline.id);
  const againOffline = await delivery.submitPod({
    order_number: "S260405",
    client_is_receiver: true,
    client_submit_id: offlineId,
    photos: [{ name: "door.jpg", mime: "image/jpeg", data: jpegB64 }],
    signature: pngDataUrl
  }, "Lebo");
  assert.ok(againOffline.already);
  assert.strictEqual(againOffline.id, firstOffline.id);
  const replay = await delivery.submitPod({
    order_number: "S260404",
    client_is_receiver: true,
    photos: [{ name: "door.jpg", mime: "image/jpeg", data: jpegB64 }],
    signature: pngDataUrl
  }, "Lebo");
  assert.ok(replay.already, "already-delivered Gauteng POD can replay after reconnect");
  assert.strictEqual(replay.id, gautengPod.id);

  const costDay = delivery.deliveryDayKey("2026-10-06T10:15:00.000Z");
  const cost = await delivery.deliveryCostForDay(costDay, { rate: 6, road: false });
  assert.strictEqual(cost.day, costDay);
  assert.ok(cost.forms_with_gps >= 2, "cost uses GPS pins from the day’s forms");
  assert.ok(cost.total_km > 0);
  assert.ok(cost.total_zar > 0);
  assert.ok((cost.batches || []).length >= 1);
  const batch = cost.batches[0];
  assert.ok(batch.return_km > 0, "return to factory is included");
  assert.ok((batch.stops_detail || []).length >= 2);
  const shareSum = batch.stops_detail.reduce((s, row) => s + Number(row.return_share_km || 0), 0);
  assert.ok(Math.abs(shareSum - batch.return_km) < 0.2, "return is fully allocated by weighted shares");

  console.log("delivery-pod.test.js ok", { splits: loaded.count, forms: listedForms.length, costZar: cost.total_zar });
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
