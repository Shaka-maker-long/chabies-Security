const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-dash-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";

const db = require("./db");
const dash = require("./enquiry-dashboard");

const now = Date.now();
function iso(ms) { return new Date(ms).toISOString(); }
function daysAgo(n) { return iso(now - n * 86400000); }

const rows = [
  {
    enquiry_no: "#1996",
    status: "New",
    created_at: daysAgo(10),
    date_enquired: "23/08/2026",
    enquiry_source: "Website",
    enquiry_type: "Catologue",
    category: "Mirror",
    product: "Daphne Rectangular Mirror",
    province: "Gauteng",
    events: [{ kind: "created", at: daysAgo(10) }],
    tasks: [],
    correspondence: { mails: [] }
  },
  {
    enquiry_no: "#1997",
    status: "Quoted",
    created_at: daysAgo(20),
    date_quoted: "20/08/2026",
    date_enquired: "13/08/2026",
    enquiry_source: "Instagram",
    enquiry_type: "Custom",
    category: "Table",
    product: "Air Chair",
    province: "Western Cape",
    quote_total_excl_vat: "4650.50",
    custom_specs: [
      { kind: "Dimensions", detail: "800 x 600" },
      { kind: "Colour", detail: "black" }
    ],
    events: [
      { kind: "created", at: daysAgo(20) },
      { kind: "complete_quote", at: daysAgo(12) }
    ],
    tasks: [{ status: "open", assignee: "Coster", kind: "follow_up" }],
    follow_ups: [],
    correspondence: { mails: [{ subject: "Quote" }] }
  },
  {
    enquiry_no: "#1998",
    status: "Ordered",
    created_at: daysAgo(40),
    ordered_at: daysAgo(5),
    date_enquired: "24/07/2026",
    date_quoted: "01/08/2026",
    enquiry_source: "Website",
    enquiry_type: "Catologue",
    category: "Mirror",
    product: "Daphne Rectangular Mirror",
    province: "Gauteng",
    quote_total_excl_vat: "12000.00",
    lifespan_ms: 35 * 86400000,
    events: [
      { kind: "created", at: daysAgo(40) },
      { kind: "complete_quote", at: daysAgo(25) },
      { kind: "complete_order", at: daysAgo(5), status: "Ordered" }
    ],
    tasks: [],
    correspondence: { mails: [] }
  },
  {
    enquiry_no: "#1999",
    status: "Rejected",
    created_at: daysAgo(8),
    date_enquired: "25/08/2026",
    enquiry_source: "Walk-in",
    enquiry_type: "Inexss",
    category: "Gate",
    product: "Driveway Gate",
    province: "KwaZulu-Natal",
    events: [
      { kind: "created", at: daysAgo(8) },
      { kind: "complete_reject", at: daysAgo(1), status: "Rejected" }
    ],
    tasks: [],
    correspondence: { mails: [] }
  },
  {
    enquiry_no: "#2000",
    status: "Costing",
    created_at: daysAgo(3),
    date_enquired: "30/08/2026",
    enquiry_source: "Website",
    enquiry_type: "Catologue",
    category: "Mirror",
    product: "Round Mirror",
    province: "Gauteng",
    events: [
      { kind: "created", at: daysAgo(3) },
      { kind: "assign_costing", at: daysAgo(2) }
    ],
    tasks: [{ status: "open", assignee: "Coster", kind: "cost_sheet" }],
    correspondence: { mails: [] }
  },
  {
    enquiry_no: "#2001",
    status: "New",
    created_at: "2026-06-15T08:00:00.000Z",
    date_enquired: "15/06/2026",
    enquiry_source: "Website",
    enquiry_type: "Catologue",
    category: "Mirror",
    product: "June Mirror",
    province: "Gauteng",
    client_name: "June Client",
    events: [{ kind: "created", at: "2026-06-15T08:00:00.000Z" }],
    tasks: [],
    correspondence: { mails: [] }
  },
  {
    enquiry_no: "#2002",
    status: "New",
    created_at: daysAgo(6),
    date_enquired: "27/08/2026",
    enquiry_source: "Website",
    enquiry_type: "New Design",
    design_description: "Steel dining table with a live-edge oak top and arched black base.",
    category: "Table",
    product: "Custom table",
    province: "Gauteng",
    client_name: "Design Client",
    events: [{ kind: "created", at: daysAgo(6) }],
    tasks: [],
    correspondence: { mails: [] }
  }
];

const orig = db.listEnquiries;
db.listEnquiries = () => rows;

try {
  const month = dash.buildDashboard({ grain: "month", range: "6m" });
  assert.strictEqual(month.grain, "month");
  assert.strictEqual(month.range, "6m");
  assert.strictEqual(month.tz, "Africa/Johannesburg");
  assert.strictEqual(month.enquiryCount, 7);
  assert.strictEqual(month.kpis.openNow, 5);
  assert.strictEqual(month.kpis.orderedInPeriod, 1);
  assert.ok(month.kpis.quotedWaiting >= 1);
  assert.ok(month.kpis.overdueFollowUps >= 1, "quoted with no follow-up after 7 days is overdue");
  assert.strictEqual(month.kpis.medianDaysToOrder, 35);
  assert.ok(month.series.length >= 6);
  assert.ok(month.series.every((s) => /^\d{4}-\d{2}$/.test(s.key)));
  const opened = month.series.reduce((n, s) => n + s.enquiries, 0);
  const quoted = month.series.reduce((n, s) => n + s.quotes, 0);
  assert.strictEqual(opened, 7);
  assert.strictEqual(quoted, 2);
  assert.strictEqual(month.money.quotedExclVat, 16650.5);
  assert.strictEqual(month.money.orderedExclVat, 12000);
  const captured = month.funnel.find((x) => x.label === "Captured");
  assert.strictEqual(captured.count, 7);
  const orderedFunnel = month.funnel.find((x) => x.label === "Ordered");
  assert.strictEqual(orderedFunnel.count, 1);
  assert.ok(month.pipeline.some((p) => p.status === "Costing" && p.count === 1));
  assert.strictEqual(month.stuck.costingOpen, 1);
  assert.strictEqual(month.stuck.quotedWaiting, 1);
  assert.ok(month.winLoss.find((x) => x.label === "Rejected").count >= 1);
  assert.ok(month.winLossByPeriod.length === month.series.length);
  assert.ok(month.sources.some((s) => s.label === "Website" && s.count >= 2));
  assert.ok(month.types.some((s) => s.label === "Catologue"));
  assert.ok(month.types.some((s) => s.label === "Inexss"));
  assert.ok(month.categories.some((s) => s.label === "Mirror"));
  assert.ok(month.provinces.some((p) => p.label === "Gauteng" && p.opened >= 2 && p.ordered >= 1));
  assert.ok(month.types.some((s) => s.label === "New Design"));
  assert.ok(month.types.some((s) => s.label === "Custom" && s.quoteValue === 4650.5));
  assert.ok(month.typeSplits.Custom.some((s) => s.label === "Dimensions"));
  assert.ok(month.typeSplits.Custom.some((s) => s.label === "Colour"));
  assert.ok(month.typeSplits["New Design"].some((s) => /Steel dining table/.test(s.value)));
  assert.ok(month.weekCompare.weeks.length === 5);
  assert.ok(month.weekCompare.months.some((m) => m.key === "2026-06" && m.enquiries[2] === 1));
  assert.ok(month.provinces.some((p) => p.label === "Gauteng" && p.quoteValue >= 0));
  assert.ok(month.products.some((p) => p.label === "Daphne Rectangular Mirror"));
  assert.ok(month.workload.some((w) => w.name === "Coster" && w.count >= 2));
  assert.ok(month.timeToOrder.buckets.some((b) => b.count === 1));
  assert.ok(month.stageTime.some((s) => s.n >= 1));

  const week = dash.buildDashboard({ grain: "week", range: "6m" });
  assert.strictEqual(week.grain, "week");
  assert.ok(week.series.length > month.series.length);
  assert.ok(week.series.every((s) => /^\d{4}-W\d{2}$/.test(s.key)));
  assert.ok(week.series.every((s) => / W\d{2}$/.test(s.label)));
  assert.strictEqual(week.series.reduce((n, s) => n + s.enquiries, 0), 7);
  assert.strictEqual(week.series.reduce((n, s) => n + s.quotes, 0), 2);

  const empty = dash.buildDashboard({ grain: "month", range: "year" });
  db.listEnquiries = () => [];
  const none = dash.buildDashboard({ grain: "month", range: "all" });
  assert.strictEqual(none.enquiryCount, 0);
  assert.strictEqual(none.kpis.openNow, 0);
  assert.strictEqual(none.kpis.medianDaysToOrder, null);
  assert.ok(none.series.length > 0);
  db.listEnquiries = () => rows;
  assert.ok(empty.enquiryCount === 7);

  const june = dash.buildDashboard({ grain: "month", month: "2026-06" });
  assert.strictEqual(june.month, "2026-06");
  assert.strictEqual(june.windowLabel, "Jun 2026");
  assert.ok(june.series.length === 1 && june.series[0].key === "2026-06");
  assert.strictEqual(june.series[0].enquiries, 1);
  assert.ok(june.pipeline.every((p) => p.status === "New"));
  assert.strictEqual(june.funnel.find((x) => x.label === "Captured").count, 1);
  assert.ok(!june.pipeline.some((p) => p.status === "Costing"), "June filter must not include later costing work");

  const drillJune = dash.buildDrill({ grain: "month", month: "2026-06", kind: "enquiries", key: "2026-06" });
  assert.strictEqual(drillJune.rows.length, 1);
  assert.strictEqual(drillJune.rows[0].enquiry_no, "#2001");
  assert.ok(drillJune.rows[0].client_name);

  const drillQuotes = dash.buildDrill({ grain: "month", range: "6m", kind: "quotes", key: "2026-08" });
  assert.ok(drillQuotes.rows.some((r) => r.enquiry_no === "#1997"));
  assert.ok(drillQuotes.rows.some((r) => r.enquiry_no === "#1998"));

  const drillEnq = dash.buildDrill({ grain: "month", range: "6m", kind: "enquiries", key: "2026-08" });
  assert.ok(drillEnq.rows.every((r) => r.enquiry_no !== "#2001"));
  const drillCustom = dash.buildDrill({ grain: "month", range: "6m", kind: "typeSubtype", type: "Custom", value: "Dimensions" });
  assert.ok(drillCustom.rows.some((r) => r.enquiry_no === "#1997"));
  const drillDesign = dash.buildDrill({ grain: "month", range: "6m", kind: "typeSubtype", type: "New Design", value: "Steel dining table with a live-edge oak top and arched black base." });
  assert.ok(drillDesign.rows.some((r) => r.enquiry_no === "#2002"));
  const drillWom = dash.buildDrill({ kind: "wom", month: "2026-06", week: 3, series: "enquiries" });
  assert.ok(drillWom.rows.some((r) => r.enquiry_no === "#2001"));
} finally {
  db.listEnquiries = orig;
}

console.log("enquiry-dashboard.test.js ok");
