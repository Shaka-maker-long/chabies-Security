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
assert.ok(/lead time is 5 weeks/.test(gautengStandard.label));
assert.ok(/strictest item is Standard/.test(gautengStandard.label));
assert.ok(/province is Gauteng/.test(gautengStandard.label));
assert.ok(/Latest Delivery/.test(gautengStandard.label));
assert.ok(fromEnquiry.isTueOrThu(gautengStandard.date));

assert.strictEqual(fromEnquiry.snapToScheduleDay("2026-09-11", "Western Cape", "2026-09-11"), "2026-09-14");
assert.strictEqual(fromEnquiry.snapToScheduleDay("2026-09-11", "Gauteng", "2026-09-11"), "2026-09-15");
assert.ok(fromEnquiry.isMonday(fromEnquiry.snapToScheduleDay("2026-09-16", "KwaZulu-Natal", "2026-09-07")));
assert.strictEqual(fromEnquiry.snapToScheduleDay("2026-09-16", "KwaZulu-Natal", "2026-09-07"), "2026-09-14");
assert.ok(fromEnquiry.isNextIsoWeek("2026-09-07", "2026-09-16"));
assert.ok(!fromEnquiry.isNextIsoWeek("2026-09-07", "2026-09-10"));

const nextWeekLc = fromEnquiry.estimateDelivery({
  types: ["Standard"],
  province: "Western Cape",
  now: "2026-09-07T10:00:00+02:00",
  date: "2026-09-14"
});
assert.strictEqual(nextWeekLc.date, "2026-09-14");
assert.strictEqual(nextWeekLc.scheduleCode, "LC");
assert.ok(nextWeekLc.nextWeekLc);
assert.ok(/Monday 14 September 2026/.test(nextWeekLc.label));
assert.ok(/next week/.test(nextWeekLc.label));
assert.ok(/not Gauteng/.test(nextWeekLc.label));
assert.ok(/Latest Courier/.test(nextWeekLc.label));

assert.throws(
  () => fromEnquiry.assertDeliveryDay("2026-09-15", "Western Cape", "2026-09-07"),
  /Monday/
);
assert.strictEqual(fromEnquiry.assertDeliveryDay("2026-09-14", "Western Cape", "2026-09-07"), "2026-09-14");
assert.strictEqual(fromEnquiry.assertDeliveryDay("2026-09-15", "Gauteng", "2026-09-07"), "2026-09-15");
assert.throws(
  () => fromEnquiry.assertDeliveryDay("2026-09-14", "Gauteng", "2026-09-07"),
  /Tuesday and Thursday/
);

const capeCustom = fromEnquiry.estimateDelivery({
  types: ["Custom"],
  province: "Western Cape",
  now
});
assert.strictEqual(capeCustom.weeks, 8);
assert.strictEqual(capeCustom.scheduleCode, "LC");
assert.ok(fromEnquiry.isTueOrThu(capeCustom.date));
assert.ok(/strictest item is Custom/.test(capeCustom.label));
assert.ok(/not Gauteng/.test(capeCustom.label));
assert.ok(/lead time is 8 weeks/.test(capeCustom.label));

const mixed = fromEnquiry.estimateDelivery({
  types: ["Standard", "New Design"],
  province: "Gauteng",
  now
});
assert.strictEqual(mixed.typeUsed, "New Design");
assert.strictEqual(mixed.weeks, 8);
assert.ok(/lead time is 8 weeks/.test(mixed.label));
assert.ok(/strictest item is New Design/.test(mixed.label));

assert.throws(() => fromEnquiry.assertTueThu("2026-10-14"), /Tuesday and Thursday/);
assert.strictEqual(fromEnquiry.assertTueThu("2026-10-15"), "2026-10-15");

function shopFields(extra) {
  return Object.assign({
    variation: "Black",
    doors: "None",
    powder_coating: "Matt Black",
    dimensions: "800x400x400",
    detailed_description: "Shop notes",
    amount_paid: "0",
    qty_price_confirmed: true
  }, extra || {});
}

const planned = fromEnquiry.planCreate(
  { enquiry_no: "#5001", quote_no: "SOQ88" },
  {
    order_number: "S260100",
    delivery_date: "2026-10-15",
    shared: {
      client_name: "Split Client",
      province: "Gauteng",
      address: "12 Main Road",
      city: "Sandton",
      email: "split@example.com"
    },
    products: [
      shopFields({ product: "Air Chair", category: "Chair", type: "Standard", quantity: 2, price_incl_vat: "20000", amount_paid: "2000" }),
      shopFields({ product: "Air Bar Stool", category: "Chair", type: "Custom", quantity: 1, price_incl_vat: "5000", amount_paid: "0" })
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
    shared: { client_name: "One Chair", province: "Western Cape", address: "1 Beach Road", city: "Stellenbosch" },
    products: [shopFields({ product: "Air Chair", type: "Standard", quantity: 1, price_incl_vat: "28750" })]
  },
  []
);
assert.strictEqual(single.units.length, 1);
assert.strictEqual(single.units[0].order_number, "S260200");
assert.strictEqual(single.delivery.scheduleCode, "LC");

assert.throws(
  () => fromEnquiry.planCreate(
    { enquiry_no: "#5003" },
    {
      order_number: "S260300",
      delivery_date: "2026-10-15",
      shared: { client_name: "No Address", province: "Gauteng", city: "Sandton" },
      products: [shopFields({ product: "Air Chair", quantity: 1, price_incl_vat: "1000" })]
    },
    []
  ),
  /Address/
);
assert.throws(
  () => fromEnquiry.planCreate(
    { enquiry_no: "#5003" },
    {
      order_number: "S260300",
      delivery_date: "2026-10-15",
      shared: { client_name: "No City", province: "Gauteng", address: "12 Main Road" },
      products: [shopFields({ product: "Air Chair", quantity: 1, price_incl_vat: "1000" })]
    },
    []
  ),
  /City/
);
assert.throws(
  () => fromEnquiry.planCreate(
    { enquiry_no: "#5003" },
    {
      order_number: "S260300",
      delivery_date: "2026-10-15",
      shared: { client_name: "No Confirm", province: "Gauteng", address: "12 Main Road", city: "Sandton" },
      products: [shopFields({ product: "Air Chair", quantity: 1, price_incl_vat: "1000", qty_price_confirmed: false })]
    },
    []
  ),
  /Confirm quantity and price excl VAT/
);
assert.throws(
  () => fromEnquiry.planCreate(
    { enquiry_no: "#5003" },
    {
      order_number: "S260300",
      delivery_date: "2026-10-15",
      shared: { client_name: "No Paid", province: "Gauteng", address: "12 Main Road", city: "Sandton" },
      products: [shopFields({ product: "Air Chair", quantity: 1, price_incl_vat: "1000", amount_paid: "" })]
    },
    []
  ),
  /Amount paid is required/
);

const paidFullPlan = fromEnquiry.planCreate(
  { enquiry_no: "#5006", quote_no: "SOQ92" },
  {
    order_number: "S260500",
    delivery_date: "2026-10-15",
    shared: { client_name: "Paid Full", province: "Gauteng", address: "12 Main Road", city: "Sandton" },
    products: [shopFields({
      product: "Air Chair",
      category: "Chair",
      type: "Standard",
      quantity: 2,
      price_incl_vat: "29322.00",
      amount_paid: "",
      paid_in_full: true
    })]
  },
  []
);
assert.strictEqual(paidFullPlan.units.length, 2);
assert.strictEqual(paidFullPlan.units[0].price_incl_vat, "14661.00");
assert.strictEqual(paidFullPlan.units[1].price_incl_vat, "14661.00");
assert.strictEqual(paidFullPlan.units[0].amount_paid, "14661.00");
assert.strictEqual(paidFullPlan.units[1].amount_paid, "14661.00");

assert.throws(
  () => fromEnquiry.planCreate(
    { enquiry_no: "#5003" },
    {
      order_number: "S260300",
      delivery_date: "2026-10-15",
      shared: { client_name: "No Doors", province: "Gauteng", address: "12 Main Road", city: "Sandton" },
      products: [shopFields({ product: "Air Chair", quantity: 1, price_incl_vat: "1000", doors: "" })]
    },
    []
  ),
  /Doors/
);

const mondayLcPlan = fromEnquiry.planCreate(
  { enquiry_no: "#5004", quote_no: "SOQ91" },
  {
    order_number: "S260400",
    delivery_date: "2026-09-14",
    now: "2026-09-07T10:00:00+02:00",
    shared: { client_name: "Cape Rush", province: "Western Cape", address: "1 Beach Road", city: "Stellenbosch" },
    products: [shopFields({ product: "Air Chair", type: "Standard", quantity: 1, price_incl_vat: "1000" })]
  },
  []
);
assert.strictEqual(mondayLcPlan.delivery.date, "2026-09-14");
assert.strictEqual(mondayLcPlan.delivery.scheduleCode, "LC");
assert.ok(mondayLcPlan.delivery.nextWeekLc);

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
assert.strictEqual(draft.products[0].detailed_description, "");
assert.strictEqual(draft.shared.city, "");
assert.strictEqual(draft.delivery.scheduleCode, "LD");
assert.ok(draft.dropdowns.type.indexOf("Standard") !== -1);

assert.ok(db.listEnquiriesWaitingForOrders().some((row) => row.enquiry_no === "#5001"), "ordered enquiries with no Orders row are ready to add");

const created = db.createOrdersFromEnquiryForm("#5001", {
  order_number: "S260100",
  delivery_date: "2026-10-15",
  shared: {
    client_name: "Split Client",
    client_number: "0820000001",
    email: "split@example.com",
    province: "Gauteng",
    city: "Sandton",
    address: "12 Main Road",
    source: "Website"
  },
  products: [
    shopFields({ product: "Air Chair", category: "Chair", type: "Standard", quantity: 2, price_incl_vat: "20000", amount_paid: "2000", detailed_description: "Pair" }),
    shopFields({ product: "Air Bar Stool", category: "Chair", type: "Custom", quantity: 1, price_incl_vat: "5000", amount_paid: "0", detailed_description: "Stool" })
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
assert.ok(!db.listEnquiriesWaitingForOrders().some((row) => row.enquiry_no === "#5001"), "enquiries already on Orders leave the ready-to-add list");

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
  shared: { client_name: "Cape Client", province: "Western Cape", address: "1 Beach Road", city: "Stellenbosch" },
  products: [shopFields({ product: "Custom Table", category: "Table", type: "New Design", quantity: 1, price_incl_vat: "11500" })]
});
assert.strictEqual(one.rows.length, 1);
assert.strictEqual(one.rows[0].order_number, "S260200");
assert.strictEqual(one.delivery.scheduleCode, "LC");
const capeItems = db.listDeliveryItems().items.filter((it) => it.order_number === "S260200");
assert.strictEqual(capeItems.length, 1);
assert.strictEqual(capeItems[0].code, "LC");
assert.strictEqual(capeItems[0].day, "2026-11-05");

assert.ok(fromEnquiry.orderBaseTaken("S260200", db.listOrders()));

db.upsertEnquiry({
  enquiry_no: "#5003",
  status: "Ordered",
  client_name: "Winelands Design Studio",
  client_number: "0764405425",
  client_email: "chante@winelandsdesignstudio.com",
  province: "Western Cape",
  enquiry_type: "Catologue",
  quote_no: "SOQ91",
  products: [{ product: "Air Chair", category: "Chair", value_incl_vat: "29322.00" }],
  ready_for_orders: true
}, { fromPipeline: true, fromMigrate: true });

const winelands = db.createOrdersFromEnquiryForm("#5003", {
  order_number: "S260001",
  delivery_date: "2026-11-05",
  shared: {
    client_name: "Winelands Design Studio",
    client_number: "0764405425",
    email: "chante@winelandsdesignstudio.com",
    province: "Western Cape",
    city: "Stellenbosch",
    address: "1 Beach Road"
  },
  products: [shopFields({
    product: "Air Chair",
    category: "Chair",
    type: "Standard",
    quantity: 2,
    price_incl_vat: "29322.00",
    amount_paid: "29322.00",
    detailed_description: "Pair"
  })]
});
assert.strictEqual(winelands.rows.length, 2);
assert.strictEqual(winelands.rows[0].price_incl_vat, "R 14,661.00");
assert.strictEqual(winelands.rows[1].price_incl_vat, "R 14,661.00");
assert.strictEqual(winelands.rows[0].amount_paid, "R 14,661.00");
assert.strictEqual(winelands.rows[1].amount_paid, "R 14,661.00");
assert.strictEqual(winelands.rows[0].owing, "R 0.00");
assert.strictEqual(winelands.rows[1].owing, "R 0.00");
assert.ok(!db.listDebtors().some((o) => String(o.order_number).indexOf("S260001") === 0));

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
assert.ok(page.indexOf("Address *") !== -1);
assert.ok(page.indexOf("City *") !== -1);
assert.ok(page.indexOf("qty_price_confirmed") !== -1);
assert.ok(page.indexOf("Paid in full") !== -1);
assert.ok(page.indexOf("data-f=\\\"paid_in_full\\\"") !== -1);
assert.ok(page.indexOf("strictest item") !== -1);
assert.ok(page.indexOf("put LC on Monday") !== -1);
assert.ok(page.indexOf("Mark selected as important") !== -1);
assert.ok(page.indexOf("does not copy from the enquiry") !== -1);
assert.ok(page.indexOf("Office schedule") === -1);

const officeJs = fs.readFileSync(path.join(__dirname, "office.js"), "utf8");
assert.ok(officeJs.indexOf("/create-order-draft") !== -1);
assert.ok(officeJs.indexOf("createOrdersFromEnquiryForm") !== -1);

console.log("create-order-from-enquiry.test.js ok");
