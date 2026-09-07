const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const fromEnquiry = require("./create-order-from-enquiry");

assert.strictEqual(fromEnquiry.orderTypeFromEnquiryType("Catologue"), "Standard");
assert.strictEqual(fromEnquiry.orderTypeFromEnquiryType("Standard"), "Standard");
assert.strictEqual(fromEnquiry.orderTypeFromEnquiryType("Custom"), "Custom");
assert.strictEqual(fromEnquiry.orderTypeFromEnquiryType("Variation on Standard"), "Custom");
assert.strictEqual(fromEnquiry.orderTypeFromEnquiryType("New Design"), "New Design");
assert.strictEqual(fromEnquiry.strictestOrderType(["Standard", "Custom", "New Design"]), "New Design");
assert.strictEqual(fromEnquiry.strictestOrderType(["Standard", "Custom"]), "Custom");
assert.strictEqual(fromEnquiry.scheduleCodeForProvince("Gauteng"), "LD");
assert.strictEqual(fromEnquiry.scheduleCodeForProvince("Western Cape"), "LC");
assert.strictEqual(fromEnquiry.leadWeeksFor("Standard", "Gauteng"), 5);
assert.strictEqual(fromEnquiry.leadWeeksFor("Standard", "KwaZulu-Natal"), 7);
assert.strictEqual(fromEnquiry.leadWeeksFor("Custom", "Gauteng"), 7);
assert.strictEqual(fromEnquiry.leadWeeksFor("Custom", "Western Cape"), 8);
assert.strictEqual(fromEnquiry.leadWeeksFor("New Design", "Gauteng"), 8);
assert.strictEqual(fromEnquiry.leadWeeksFor("New Design", "Limpopo"), 10);

assert.strictEqual(fromEnquiry.normalizeBaseOrderNumber("s260100 a"), "S260100");
assert.strictEqual(fromEnquiry.normalizeBaseOrderNumber("S260100"), "S260100");
assert.strictEqual(fromEnquiry.studioOrderSeq("S260100 B"), 260100);
assert.strictEqual(
  fromEnquiry.nextStudioOrderNumberFrom([{ order_number: "S260100 A" }, { order_number: "S260100 C" }]),
  "S260101"
);
assert.ok(fromEnquiry.orderBaseTaken("S260100", [{ order_number: "S260100 A" }]));
assert.ok(!fromEnquiry.orderBaseTaken("S260101", [{ order_number: "S260100 A" }]));

assert.strictEqual(fromEnquiry.splitCents(20000, 2, 0), "10000.00");
assert.strictEqual(fromEnquiry.splitCents(20000, 2, 1), "10000.00");
assert.strictEqual(fromEnquiry.splitCents("20,000.00", 3, 0), "6666.66");
assert.strictEqual(fromEnquiry.splitCents("20,000.00", 3, 2), "6666.68");

const now = "2026-09-07T10:00:00+02:00";
const gautengStandard = fromEnquiry.estimateDelivery({
  types: ["Standard"],
  province: "Gauteng",
  now
});
assert.strictEqual(gautengStandard.weeks, 5);
assert.strictEqual(gautengStandard.date, "2026-10-13");
assert.strictEqual(gautengStandard.scheduleCode, "LD");
assert.ok(/Tuesday 13 October 2026/.test(gautengStandard.label));
assert.ok(/week 42/.test(gautengStandard.label));
assert.ok(fromEnquiry.isTueOrThu(gautengStandard.date));

const capeCustom = fromEnquiry.estimateDelivery({
  types: ["Custom"],
  province: "Western Cape",
  now
});
assert.strictEqual(capeCustom.weeks, 8);
assert.strictEqual(capeCustom.scheduleCode, "LC");
assert.ok(fromEnquiry.isTueOrThu(capeCustom.date));

const mixed = fromEnquiry.estimateDelivery({
  types: ["Standard", "New Design"],
  province: "Gauteng",
  now
});
assert.strictEqual(mixed.typeUsed, "New Design");
assert.strictEqual(mixed.weeks, 8);

assert.throws(() => fromEnquiry.assertTueThu("2026-10-14"), /Tuesday and Thursday/);
assert.strictEqual(fromEnquiry.assertTueThu("2026-10-15"), "2026-10-15");

const planned = fromEnquiry.planCreate(
  { enquiry_no: "#5001", quote_no: "SOQ88" },
  {
    order_number: "S260100",
    delivery_date: "2026-10-15",
    shared: {
      client_name: "Split Client",
      province: "Gauteng",
      city: "Sandton",
      email: "split@example.com"
    },
    products: [
      { product: "Air Chair", category: "Chair", type: "Standard", quantity: 2, price_incl_vat: "20000", amount_paid: "2000" },
      { product: "Air Bar Stool", category: "Chair", type: "Custom", quantity: 1, price_incl_vat: "5000", amount_paid: "0" }
    ]
  },
  []
);
assert.strictEqual(planned.units.length, 3);
assert.deepStrictEqual(planned.units.map((u) => u.order_number), ["S260100 A", "S260100 B", "S260100 C"]);
assert.strictEqual(planned.units[0].product, "Air Chair");
assert.strictEqual(planned.units[1].product, "Air Chair");
assert.strictEqual(planned.units[2].product, "Air Bar Stool");
assert.strictEqual(planned.units[0].price_incl_vat, "10000.00");
assert.strictEqual(planned.units[1].price_incl_vat, "10000.00");
assert.strictEqual(planned.units[2].price_incl_vat, "5000.00");
assert.strictEqual(planned.units[0].amount_paid, "1000.00");
assert.strictEqual(planned.delivery.scheduleCode, "LD");
assert.strictEqual(planned.delivery.date, "2026-10-15");
assert.strictEqual(planned.quote_number, "SOQ88");

const single = fromEnquiry.planCreate(
  { enquiry_no: "#5002", quote_no: "SOQ89" },
  {
    order_number: "S260200",
    delivery_date: "2026-10-13",
    shared: { client_name: "One Chair", province: "Western Cape" },
    products: [{ product: "Air Chair", type: "Standard", quantity: 1, price_incl_vat: "28750" }]
  },
  []
);
assert.strictEqual(single.units.length, 1);
assert.strictEqual(single.units[0].order_number, "S260200");
assert.strictEqual(single.delivery.scheduleCode, "LC");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-from-enq-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const db = require("./db");
initWorkbook();

db.upsertEnquiry({
  enquiry_no: "#5001",
  status: "Ordered",
  client_name: "Split Client",
  client_number: "0820000001",
  client_email: "split@example.com",
  province: "Gauteng",
  source: "Website",
  enquiry_type: "Catologue",
  quote_no: "SOQ88",
  request: "Two chairs and a stool",
  products: [
    { product: "Air Chair", category: "Chair", value_incl_vat: "23000.00" },
    { product: "Air Bar Stool", category: "Chair", value_incl_vat: "5750.00" }
  ],
  ready_for_orders: true
}, { fromPipeline: true, fromMigrate: true });

const draft = db.createOrderDraftFromEnquiry("#5001", now);
assert.strictEqual(draft.quote_number, "SOQ88");
assert.ok(/^S\d+$/.test(draft.order_number));
assert.strictEqual(draft.products.length, 2);
assert.strictEqual(draft.products[0].type, "Standard");
assert.strictEqual(draft.products[0].quantity, 1);
assert.strictEqual(draft.shared.city, "");
assert.strictEqual(draft.delivery.scheduleCode, "LD");
assert.ok(draft.dropdowns.type.indexOf("Standard") !== -1);

const created = db.createOrdersFromEnquiryForm("#5001", {
  order_number: "S260100",
  delivery_date: "2026-10-15",
  shared: {
    client_name: "Split Client",
    client_number: "0820000001",
    email: "split@example.com",
    province: "Gauteng",
    city: "Sandton",
    source: "Website"
  },
  products: [
    { product: "Air Chair", category: "Chair", type: "Standard", quantity: 2, price_incl_vat: "20000", amount_paid: "2000", variation: "Black", doors: "", detailed_description: "Pair", dimensions: "800x400x400", powder_coating: "Matt Black" },
    { product: "Air Bar Stool", category: "Chair", type: "Custom", quantity: 1, price_incl_vat: "5000", amount_paid: "0", detailed_description: "Stool" }
  ]
});
assert.strictEqual(created.existing, false);
assert.strictEqual(created.rows.length, 3);
assert.deepStrictEqual(created.rows.map((r) => r.order_number), ["S260100 A", "S260100 B", "S260100 C"]);
assert.strictEqual(created.rows[0].status, "Not Yet Started");
assert.strictEqual(created.rows[0].quote_number, "SOQ88");
assert.strictEqual(created.rows[0].enquiry_no, "#5001");
assert.strictEqual(created.rows[0].city, "Sandton");
const savedA = db.listOrders().find((o) => o.order_number === "S260100 A");
assert.strictEqual(savedA.price_incl_vat, "10000.00");
assert.ok(Number(savedA.price_excl_vat) > 0);
assert.strictEqual(created.rows[2].product, "Air Bar Stool");
assert.strictEqual(created.rows[2].type, "Custom");
assert.strictEqual(db.getEnquiry("#5001").order_number, "S260100 A");
assert.strictEqual(db.nextStudioOrderNumber(), "S260101");

const items = db.listDeliveryItems().items.filter((it) => String(it.order_number).indexOf("S260100") === 0);
assert.strictEqual(items.length, 3);
items.forEach((it) => {
  assert.strictEqual(it.day, "2026-10-15");
  assert.strictEqual(it.code, "LD");
});

const again = db.createOrdersFromEnquiryForm("#5001", {
  order_number: "S260199",
  delivery_date: "2026-10-15",
  shared: { client_name: "Split Client", province: "Gauteng" },
  products: [{ product: "Air Chair", quantity: 1, price_incl_vat: "1000" }]
});
assert.strictEqual(again.existing, true);
assert.strictEqual(again.rows.length, 3);

db.upsertEnquiry({
  enquiry_no: "#5002",
  status: "Ordered",
  client_name: "Cape Client",
  province: "Western Cape",
  enquiry_type: "New Design",
  quote_no: "SOQ90",
  products: [{ product: "Custom Table", category: "Table", value_incl_vat: "11500.00" }],
  ready_for_orders: true
}, { fromPipeline: true, fromMigrate: true });

const one = db.createOrdersFromEnquiryForm("#5002", {
  order_number: "S260200",
  delivery_date: "2026-11-05",
  shared: { client_name: "Cape Client", province: "Western Cape", city: "Stellenbosch" },
  products: [{ product: "Custom Table", category: "Table", type: "New Design", quantity: 1, price_incl_vat: "11500" }]
});
assert.strictEqual(one.rows.length, 1);
assert.strictEqual(one.rows[0].order_number, "S260200");
assert.strictEqual(one.delivery.scheduleCode, "LC");
const capeItems = db.listDeliveryItems().items.filter((it) => it.order_number === "S260200");
assert.strictEqual(capeItems.length, 1);
assert.strictEqual(capeItems[0].code, "LC");
assert.strictEqual(capeItems[0].day, "2026-11-05");

assert.ok(fromEnquiry.orderBaseTaken("S260200", db.listOrders()));

const page = fs.readFileSync(path.join(__dirname, "../public/orders-from-enquiry.html"), "utf8");
assert.ok(page.indexOf("Create orders from enquiry") !== -1);
assert.ok(page.indexOf("f_order_number") !== -1);
assert.ok(page.indexOf("Quote number") !== -1);
assert.ok(page.indexOf("f_status") === -1, "do not put a status field on the enquiry-to-order form");
assert.ok(page.indexOf("Tuesday and Thursday") !== -1);
assert.ok(page.indexOf("Latest Delivery") !== -1);
assert.ok(page.indexOf("Latest Courier") !== -1);
assert.ok(page.indexOf("/api/office/enquiries/") !== -1);
assert.ok(page.indexOf("create-orders") !== -1);
assert.ok(page.indexOf("data-f=\\\"quantity\\\"") !== -1);
assert.ok(page.indexOf("Office schedule") === -1);

const officeJs = fs.readFileSync(path.join(__dirname, "office.js"), "utf8");
assert.ok(officeJs.indexOf("/create-order-draft") !== -1);
assert.ok(officeJs.indexOf("createOrdersFromEnquiryForm") !== -1);

console.log("create-order-from-enquiry.test.js ok");
