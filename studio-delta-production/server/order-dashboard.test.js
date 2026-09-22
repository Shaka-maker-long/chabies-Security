const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-order-dash-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const db = require("./db");
const noPlates = require("./no-plates");
const dash = require("./order-dashboard");

initWorkbook();

const empty = dash.buildDashboard({ range: "year", grain: "month" });
assert.strictEqual(empty.orderCount, 0);
assert.strictEqual(empty.kpis.openJobs, 0);
assert.strictEqual(empty.kpis.income, 0);
assert.ok(Array.isArray(empty.series));
assert.ok(empty.pipeline.some((p) => p.id === "drawing"));
assert.deepStrictEqual(dash.buildDrill({ kind: "open" }).rows, []);

db.upsertOrder({
  order_number: "S260401",
  status: "Welding",
  category: "Mirror",
  product: "Daphne Rectangular Mirror",
  source: "Website",
  quote_number: "22300",
  client_name: "Ann",
  price_incl_vat: "11500.00",
  amount_paid: "11500.00",
  payment_date: "14/09/2026"
});
db.upsertOrder({
  order_number: "S260402",
  status: "Waiting for drawing",
  category: "Table",
  product: "Air Chair",
  source: "Instagram",
  quote_number: "SOQ2954",
  client_name: "Ben",
  price_incl_vat: "4650.50",
  amount_paid: "2000.00",
  payment_date: "01/08/2026"
});
db.upsertOrder({
  order_number: "S260403",
  status: "Paint Shop",
  category: "Mirror",
  product: "Daphne Rectangular Mirror",
  source: "Website",
  quote_number: "22301",
  client_name: "Cara",
  price_incl_vat: "8000.00",
  amount_paid: "8000.00",
  month_of_sale: "September 2026"
});
db.upsertOrder({
  order_number: "S260404",
  status: "Delivered",
  category: "Gate",
  product: "Driveway Gate",
  source: "Walk-in",
  quote_number: "SOQ2900",
  client_name: "Dan",
  price_incl_vat: "22000.00",
  amount_paid: "22000.00",
  payment_date: "10/07/2026"
});
db.upsertOrder({
  order_number: "S260405",
  status: "Ready for Delivery",
  category: "Cabinet",
  product: "Vivienne Arched Cabinet",
  source: "Website",
  quote_number: "22450",
  client_name: "Eve",
  price_incl_vat: "18000.00",
  amount_paid: "9000.00",
  payment_date: "02/09/2026"
});
db.upsertOrder({
  order_number: "S260406",
  status: "Not Yet Started",
  category: "",
  product: "",
  source: "",
  quote_number: "",
  client_name: "Fay",
  price_incl_vat: "1000.00",
  amount_paid: "0",
  payment_date: "20/06/2026"
});

noPlates.markNoPlate("S260401", "Thabile");

const synced = db.syncScheduleFromOrders();
assert.ok(synced.synced >= 6);
const schedRows = db.listSchedule();
const weld = schedRows.find((r) => r.order_number === "S260401");
const ready = schedRows.find((r) => r.order_number === "S260405");
const paint = schedRows.find((r) => r.order_number === "S260403");
assert.ok(weld && ready && paint);
const sched = require("./office-schedule");
function isoAddDays(iso, days) {
  const d = new Date(String(iso).slice(0, 10) + "T12:00:00");
  d.setDate(d.getDate() + days);
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
}
const todayParts = (() => {
  const sast = new Date(Date.now() + 2 * 60 * 60 * 1000);
  return sast.getUTCFullYear() + "-" + String(sast.getUTCMonth() + 1).padStart(2, "0") + "-" + String(sast.getUTCDate()).padStart(2, "0");
})();
const thisMonday = sched.mondayOf(todayParts);
const thisTuesday = isoAddDays(thisMonday, 1);
const nextMonday = isoAddDays(thisMonday, 7);
const lastWednesday = isoAddDays(thisMonday, -5);
db.setScheduleCell(weld.id, thisTuesday, "LD");
db.setScheduleCell(ready.id, nextMonday, "LC");
db.setScheduleCell(paint.id, lastWednesday, "LD");

const year = dash.buildDashboard({ range: "year", grain: "month" });
assert.strictEqual(year.orderCount, 6);
assert.strictEqual(year.kpis.openJobs, 5);
assert.strictEqual(year.kpis.waitingForDrawing, 1);
assert.strictEqual(year.kpis.atPaintShop, 1);
assert.strictEqual(year.kpis.readyOrOut, 1);
assert.strictEqual(year.kpis.income, 38695.65);
assert.ok(year.kpis.unpaidOpen > 0);

assert.ok(dash.saleDate({ payment_date: "14/09/2026" }));
assert.strictEqual(dash.saleDate({ month_of_sale: "September 2026" }), null);
assert.strictEqual(dash.saleDate({ payment_date: "", month_of_sale: "September 2026" }), null);

const sep = year.series.find((s) => s.key === "2026-09");
assert.ok(sep, "September bucket exists");
assert.strictEqual(sep.count, 2);
assert.strictEqual(sep.income, 17826.09);
assert.strictEqual(sep.billed, 25652.17);
assert.strictEqual(sep.delivered, 0);
assert.strictEqual(sep.website, 2, "Sep website quotes 22300 and 22450");
assert.strictEqual(sep.offline, 0);

const jul = year.series.find((s) => s.key === "2026-07");
assert.ok(jul);
assert.strictEqual(jul.delivered, 1);
assert.strictEqual(jul.website, 0);
assert.strictEqual(jul.offline, 1, "Jul offline SOQ2900");

const aug = year.series.find((s) => s.key === "2026-08");
assert.ok(aug);
assert.strictEqual(aug.offline, 1, "Aug offline SOQ2954");
assert.strictEqual(aug.website, 0);

assert.strictEqual(dash.quoteChannel("22300"), "website");
assert.strictEqual(dash.quoteChannel("SOQ2954"), "offline");
assert.strictEqual(dash.quoteChannel("soq 2954"), "offline");
assert.strictEqual(dash.quoteChannel(""), "");

const channelWeb = dash.buildDrill({ kind: "channel", slice: "website", key: "2026-09", range: "year", grain: "month" });
assert.strictEqual(channelWeb.rows.length, 2);
assert.ok(channelWeb.rows.every((r) => r.quote_channel === "website"));
assert.ok(channelWeb.title.indexOf("Website") !== -1);
const channelOff = dash.buildDrill({ kind: "channel", slice: "offline", key: "2026-07", range: "year", grain: "month" });
assert.strictEqual(channelOff.rows.length, 1);
assert.strictEqual(channelOff.rows[0].order_number, "S260404");
assert.strictEqual(channelOff.rows[0].quote_number, "SOQ2900");
assert.strictEqual(jul.income, 19130.43);
assert.strictEqual(jul.billed, 19130.43);

const mirror = year.categories.find((c) => c.label === "Mirror");
assert.ok(mirror);
assert.strictEqual(mirror.count, 1);
assert.strictEqual(mirror.income, 10000);

const website = year.sources.find((s) => s.label === "Website");
assert.ok(website);
assert.strictEqual(website.income, 17826.09);

const topItem = year.topIncome[0];
assert.strictEqual(topItem.label, "Driveway Gate");
assert.strictEqual(topItem.income, 19130.43);
const daphne = year.topIncome.find((p) => p.label === "Daphne Rectangular Mirror");
assert.ok(daphne);
assert.strictEqual(daphne.income, 10000);
assert.strictEqual(daphne.count, 1);

const topQty = year.topQuantity.find((p) => p.label === "Daphne Rectangular Mirror");
assert.ok(topQty);
assert.strictEqual(topQty.count, 1);

const steel = year.pipeline.find((p) => p.id === "steelwork");
assert.strictEqual(steel.count, 1);
const drawing = year.pipeline.find((p) => p.id === "drawing");
assert.strictEqual(drawing.count, 1);
const delivered = year.pipeline.find((p) => p.id === "delivered");
assert.strictEqual(delivered.count, 1);
const notStarted = year.pipeline.find((p) => p.id === "not_started");
assert.ok(notStarted.count >= 1);

assert.ok(year.stuck.some((r) => r.order_number === "S260406"));
assert.ok(year.ageing.some((b) => b.count > 0));
year.ageing.forEach((b) => {
  assert.ok(Array.isArray(b.statuses));
  assert.strictEqual(b.statuses.reduce((n, s) => n + s.count, 0), b.count);
});
const weldingAge = year.ageing.find((b) => (b.statuses || []).some((s) => s.label === "Welding"));
assert.ok(weldingAge, "Welding sits in an age band");
assert.ok(weldingAge.count >= 1);
const ageWeld = dash.buildDrill({ kind: "age", value: weldingAge.id, slice: "Welding" });
assert.ok(ageWeld.rows.some((r) => r.order_number === "S260401"));
assert.ok(ageWeld.rows.every((r) => r.status === "Welding"));
assert.ok(ageWeld.title.indexOf("Welding") !== -1);
assert.ok(ageWeld.title.indexOf(weldingAge.id) !== -1);
const ageAll = dash.buildDrill({ kind: "age", value: weldingAge.id });
assert.ok(ageAll.rows.length >= ageWeld.rows.length);

assert.strictEqual(year.delivery.thisWeek.count, 1);
assert.strictEqual(year.delivery.thisWeek.days.Tuesday, 1);
assert.strictEqual(year.delivery.nextWeek.count, 1);
assert.strictEqual(year.delivery.nextWeek.days.Monday, 1);
assert.ok(year.delivery.late >= 1);

const noPlate = year.blockers.find((b) => b.id === "no_plates");
assert.strictEqual(noPlate.count, 1);
const readyBlock = year.blockers.find((b) => b.id === "ready_delivery");
assert.strictEqual(readyBlock.count, 1);

const monthOnly = dash.buildDashboard({ month: "2026-09", grain: "month" });
assert.strictEqual(monthOnly.windowCount, 2);
assert.strictEqual(monthOnly.kpis.income, 17826.09);
assert.strictEqual(monthOnly.kpis.billed, 25652.17);
assert.strictEqual(monthOnly.windowLabel, "Sep 2026");
assert.strictEqual(monthOnly.kpis.openJobs, 5, "shop KPIs stay live when a month is picked");
assert.ok(!monthOnly.categories.some((c) => c.label === "Mirror" && c.count > 1));

const week = dash.buildDashboard({ range: "year", grain: "week" });
assert.ok(week.series.some((s) => s.key.indexOf("-W") !== -1));
assert.ok(week.series.some((s) => s.income > 0));

const catDrill = dash.buildDrill({ kind: "category", value: "Mirror", range: "year", grain: "month" });
assert.strictEqual(catDrill.rows.length, 1);
assert.strictEqual(catDrill.rows[0].order_number, "S260401");
assert.ok(catDrill.title.indexOf("CATERGORY") !== -1);
assert.strictEqual(catDrill.totals.income, 10000);
assert.strictEqual(catDrill.rows[0].amount_paid, 10000);
assert.strictEqual(catDrill.rows[0].price_excl_vat, 10000);
assert.ok(!catDrill.rows.some((r) => r.order_number === "S260403"), "month of sale is not the order date");

const pipeDrill = dash.buildDrill({ kind: "pipeline", group: "drawing" });
assert.strictEqual(pipeDrill.rows.length, 1);
assert.strictEqual(pipeDrill.rows[0].order_number, "S260402");

const unpaid = dash.buildDrill({ kind: "unpaid" });
assert.ok(unpaid.rows.some((r) => r.order_number === "S260402"));
assert.ok(unpaid.rows.every((r) => r.owing > 0));
const unpaidBen = unpaid.rows.find((r) => r.order_number === "S260402");
assert.strictEqual(unpaidBen.owing, 2304.78);
assert.strictEqual(unpaidBen.amount_paid, 1739.13);

const plates = dash.buildDrill({ kind: "no_plates" });
assert.strictEqual(plates.rows.length, 1);
assert.strictEqual(plates.rows[0].order_number, "S260401");

const late = dash.buildDrill({ kind: "late" });
assert.ok(late.rows.some((r) => r.order_number === "S260403"));
assert.ok(late.rows.every((r) => r.late));

const tue = dash.buildDrill({ kind: "delivery", week: year.delivery.thisWeek.weekKey, weekday: "Tuesday" });
assert.strictEqual(tue.rows.length, 1);
assert.strictEqual(tue.rows[0].order_number, "S260401");

const blank = dash.buildDashboard({ range: "year" });
const blankCat = blank.categories.find((c) => c.label === "(Blank)");
assert.ok(blankCat);
const noProd = blank.topQuantity.find((p) => p.label === "(No product)");
assert.ok(noProd);

const html = fs.readFileSync(path.join(__dirname, "../public/orders-dashboard.html"), "utf8");
assert.ok(html.indexOf("No orders yet.") !== -1);
assert.ok(html.indexOf("Could not load the dashboard") !== -1);
assert.ok(html.indexOf("Income per month") !== -1);
assert.ok(html.indexOf("Income per week") !== -1);
assert.ok(html.indexOf("Website vs offline orders") !== -1);
assert.ok(html.indexOf("drawChannel") !== -1);
assert.ok(html.indexOf("id=\"channel\"") !== -1);
assert.ok(html.indexOf("SOQ2954") !== -1 || html.indexOf("SOQ") !== -1);
assert.ok(html.indexOf("Income by CATERGORY") !== -1);
assert.ok(html.indexOf("Income by source") !== -1);
assert.ok(html.indexOf("Top 10 items by income") !== -1);
assert.ok(html.indexOf("Top 10 products by quantity") !== -1);
assert.ok(html.indexOf("Shop pipeline") !== -1);
assert.ok(html.indexOf("Stuck / ageing") !== -1);
assert.ok(html.indexOf("openAgeStatus") !== -1);
assert.ok(html.indexOf("ageStatusPie") !== -1);
assert.ok(html.indexOf("Click a pie slice") !== -1);
assert.ok(html.indexOf("Back to statuses") !== -1);
assert.ok(html.indexOf("Delivery this week and next") !== -1);
assert.ok(html.indexOf("Throughput") !== -1);
assert.ok(html.indexOf("Blockers") !== -1);
assert.ok(html.indexOf("chart.js@4.4.7") !== -1);
assert.ok(html.indexOf("/api/office/orders/dashboard") !== -1);
assert.ok(html.indexOf("sdOfficeFetch") !== -1);
assert.ok(html.indexOf("grouped by the payment date") !== -1);
assert.ok(html.indexOf("the day the order was placed") === -1);
assert.ok(html.indexOf("month of sale if there is no payment date") === -1);
assert.ok(html.indexOf("Unpaid on open jobs (Excl VAT)") !== -1);
assert.ok(html.indexOf("Sale value (Excl VAT)") !== -1);
assert.ok(html.indexOf("Amount paid (Excl VAT)") !== -1);
assert.ok(html.indexOf("Income paid Excl VAT") !== -1);
assert.ok(html.indexOf("Paid Excl VAT") !== -1);
assert.ok(html.indexOf("Money excludes VAT") !== -1);
assert.ok(html.indexOf("Incl VAT") === -1);
assert.ok(html.indexOf("including VAT") === -1);

db.upsertOrder({
  order_number: "S260410",
  status: "Not Yet Started",
  category: "Table",
  product: "Date Check",
  source: "Website",
  client_name: "Gia",
  price_incl_vat: "100.00",
  amount_paid: "100.00",
  payment_date: "16-Sep",
  month_of_sale: "January 2020"
});
const dated = db.listOrders().find((o) => o.order_number === "S260410");
assert.strictEqual(dated.payment_date, "16/09/2026");
assert.strictEqual(dated.month_of_sale, "September 2026", "month of sale is the payment month");
assert.ok(dash.saleDate({ payment_date: "16-Sep" }));
const schedDate = db.listSchedule().find((r) => r.order_number === "S260410");
assert.strictEqual(schedDate.order_date_label, "16-Sep");
const sepDash = dash.buildDashboard({ month: "2026-09", grain: "month" });
assert.ok(sepDash.windowCount >= 3, "16-Sep payment lands in September");

db.upsertOrder({
  order_number: "S260411",
  status: "Assembly",
  category: "Mirror",
  product: "Age Pie A",
  source: "Website",
  client_name: "Hal",
  price_incl_vat: "115.00",
  amount_paid: "115.00",
  payment_date: "17/09/2026"
});
db.upsertOrder({
  order_number: "S260412",
  status: "Grinding",
  category: "Mirror",
  product: "Age Pie B",
  source: "Website",
  client_name: "Ivy",
  price_incl_vat: "115.00",
  amount_paid: "115.00",
  payment_date: "17/09/2026"
});
const pieDash = dash.buildDashboard({ range: "year" });
const pieBucket = pieDash.ageing.find((b) =>
  (b.statuses || []).some((s) => s.label === "Assembly")
  && (b.statuses || []).some((s) => s.label === "Grinding")
);
assert.ok(pieBucket, "same-age jobs split by status");
assert.ok(pieBucket.statuses.length >= 2);
const assemblySlice = dash.buildDrill({ kind: "age", value: pieBucket.id, slice: "Assembly" });
assert.ok(assemblySlice.rows.some((r) => r.order_number === "S260411"));
assert.ok(!assemblySlice.rows.some((r) => r.order_number === "S260412"));
const grindSlice = dash.buildDrill({ kind: "age", value: pieBucket.id, slice: "Grinding" });
assert.ok(grindSlice.rows.some((r) => r.order_number === "S260412"));
assert.ok(!grindSlice.rows.some((r) => r.order_number === "S260411"));

console.log("order-dashboard tests ok");
