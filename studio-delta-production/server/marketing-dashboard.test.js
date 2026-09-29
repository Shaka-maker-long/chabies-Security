const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-mkt-dash-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const db = require("./db");
const dash = require("./marketing-dashboard");

initWorkbook();

const empty = dash.buildDashboard({ range: "year", grain: "month" });
assert.strictEqual(empty.orderCount, 0);
assert.strictEqual(empty.kpis.income, 0);
assert.ok(Array.isArray(empty.chart.labels));
assert.deepStrictEqual(empty.chart.datasets, []);

db.upsertOrder({
  order_number: "S261001",
  status: "Delivered",
  product: "Air Chair",
  client_name: "Ann",
  campaign: "Spring Push",
  source: "Instagram",
  province: "Gauteng",
  city: "Johannesburg",
  price_excl_vat: "1000.00",
  amount_paid: "1150.00",
  price_incl_vat: "1150.00",
  payment_date: "10/09/2026"
});
db.upsertOrder({
  order_number: "S261002",
  status: "Assembly",
  product: "Maria Desk",
  client_name: "Ben",
  campaign: "Spring Push",
  source: "Website",
  province: "Gauteng",
  city: "Pretoria",
  price_excl_vat: "2000.00",
  amount_paid: "2300.00",
  price_incl_vat: "2300.00",
  payment_date: "18/09/2026"
});
db.upsertOrder({
  order_number: "S261003",
  status: "Not Yet Started",
  product: "Ella Arched Cabinet",
  client_name: "Cara",
  campaign: "Decor Fair",
  source: "Showroom",
  province: "Western Cape",
  city: "Cape Town",
  price_excl_vat: "5000.00",
  amount_paid: "5750.00",
  price_incl_vat: "5750.00",
  payment_date: "05/08/2026"
});
db.upsertOrder({
  order_number: "S261004",
  status: "Welding",
  product: "Jasmine Planter",
  client_name: "Dan",
  campaign: "",
  source: "Google",
  province: "KwaZulu-Natal",
  city: "Durban",
  price_excl_vat: "800.00",
  amount_paid: "920.00",
  price_incl_vat: "920.00",
  payment_date: "12/09/2026"
});

const year = dash.buildDashboard({ range: "year", grain: "month" });
assert.ok(year.orderCount >= 4);
assert.strictEqual(year.kpis.campaigns, 2);
assert.ok(year.kpis.missingOrders >= 1);
assert.ok(year.kpis.missingIncome > 0);

const spring = year.campaigns.find((c) => c.label === "Spring Push");
assert.ok(spring);
assert.strictEqual(spring.orders, 2);
assert.ok(spring.income > 0);

const decor = year.campaigns.find((c) => c.label === "Decor Fair");
assert.ok(decor);

const none = year.campaigns.find((c) => c.label === "(No campaign)");
assert.ok(none);

assert.ok(year.chart.datasets.some((d) => d.label === "Spring Push"));
assert.ok(year.chart.datasets.some((d) => d.label === "(No campaign)"), "missing campaign stays on the line chart");
assert.ok(year.chart.labels.length >= 1);

const sep = dash.buildDashboard({ month: "2026-09", grain: "month" });
assert.ok(/sep/i.test(sep.windowLabel), sep.windowLabel);
assert.ok(sep.orderCount >= 3);
const sepSpring = sep.campaigns.find((c) => c.label === "Spring Push");
assert.ok(sepSpring);
assert.strictEqual(sepSpring.orders, 2);
assert.ok(!sep.campaigns.some((c) => c.label === "Decor Fair"), "August Decor Fair stays out of September");

const week = dash.buildDashboard({ range: "6m", grain: "week" });
assert.strictEqual(week.grain, "week");
assert.ok(week.chart.labels.length >= 1);

const sepKey = "2026-09";
const provinces = dash.buildDrill({
  kind: "provinces",
  campaign: "Spring Push",
  key: sepKey,
  month: sepKey,
  grain: "month"
});
assert.strictEqual(provinces.kind, "provinces");
assert.ok(provinces.slices.some((s) => s.label === "Gauteng"));
assert.ok(provinces.totals.income > 0);

const cities = dash.buildDrill({
  kind: "cities",
  campaign: "Spring Push",
  key: sepKey,
  province: "Gauteng",
  month: sepKey,
  grain: "month"
});
assert.strictEqual(cities.kind, "cities");
assert.ok(cities.slices.some((s) => s.label === "Johannesburg" || s.label === "(No city)" || s.orders >= 1));

const orders = dash.buildDrill({
  kind: "orders",
  campaign: "Spring Push",
  key: sepKey,
  province: "Gauteng",
  city: cities.slices[0].label,
  month: sepKey,
  grain: "month"
});
assert.strictEqual(orders.kind, "orders");
assert.ok(orders.rows.length >= 1);
assert.ok(orders.rows.every((r) => r.campaign === "Spring Push"));

const html = fs.readFileSync(path.join(__dirname, "../public/marketing-dashboard.html"), "utf8");
assert.ok(html.indexOf("Campaign income per month") !== -1);
assert.ok(html.indexOf("Campaign income per week") !== -1);
assert.ok(html.indexOf("/api/office/marketing/dashboard") !== -1);
assert.ok(html.indexOf("/api/office/marketing/dashboard/drill") !== -1);
assert.ok(html.indexOf("openDrill") !== -1);
assert.ok(html.indexOf("kind: \"provinces\"") !== -1 || html.indexOf('kind: "provinces"') !== -1);
assert.ok(html.indexOf("chart.js@4.4.7") !== -1);
assert.ok(html.indexOf("sdRequireOffice(\"marketing\")") !== -1);

console.log("marketing-dashboard.test.js ok");
