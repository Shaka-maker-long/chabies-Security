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
  }
];

const orig = db.listEnquiries;
db.listEnquiries = () => rows;

try {
  const month = dash.buildDashboard({ grain: "month", range: "6m" });
  assert.strictEqual(month.grain, "month");
  assert.strictEqual(month.range, "6m");
  assert.strictEqual(month.tz, "Africa/Johannesburg");
  assert.strictEqual(month.enquiryCount, 5);
  assert.strictEqual(month.kpis.openNow, 3);
  assert.strictEqual(month.kpis.orderedInPeriod, 1);
  assert.ok(month.kpis.quotedWaiting >= 1);
  assert.ok(month.kpis.overdueFollowUps >= 1, "quoted with no follow-up after 7 days is overdue");
  assert.strictEqual(month.kpis.medianDaysToOrder, 35);
  assert.ok(month.series.length >= 6);
  assert.ok(month.series.every((s) => /^\d{4}-\d{2}$/.test(s.key)));
  const opened = month.series.reduce((n, s) => n + s.enquiries, 0);
  const quoted = month.series.reduce((n, s) => n + s.quotes, 0);
  assert.strictEqual(opened, 5);
  assert.strictEqual(quoted, 2);
  assert.strictEqual(month.money.quotedExclVat, 16650.5);
  assert.strictEqual(month.money.orderedExclVat, 12000);
  const captured = month.funnel.find((x) => x.label === "Captured");
  assert.strictEqual(captured.count, 5);
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
  assert.ok(month.outlook.some((o) => o.label === "With Outlook email" && o.count === 1));
  assert.ok(month.workload.some((w) => w.name === "Coster" && w.count >= 2));
  assert.ok(month.timeToOrder.buckets.some((b) => b.count === 1));
  assert.ok(month.stageTime.some((s) => s.n >= 1));

  const week = dash.buildDashboard({ grain: "week", range: "6m" });
  assert.strictEqual(week.grain, "week");
  assert.ok(week.series.length > month.series.length);
  assert.ok(week.series.every((s) => /^\d{4}-W\d{2}$/.test(s.key)));
  assert.ok(week.series.every((s) => / W\d{2}$/.test(s.label)));
  assert.strictEqual(week.series.reduce((n, s) => n + s.enquiries, 0), 5);
  assert.strictEqual(week.series.reduce((n, s) => n + s.quotes, 0), 2);

  const empty = dash.buildDashboard({ grain: "month", range: "year" });
  db.listEnquiries = () => [];
  const none = dash.buildDashboard({ grain: "month", range: "all" });
  assert.strictEqual(none.enquiryCount, 0);
  assert.strictEqual(none.kpis.openNow, 0);
  assert.strictEqual(none.kpis.medianDaysToOrder, null);
  assert.ok(none.series.length > 0);
  db.listEnquiries = () => rows;
  assert.ok(empty.enquiryCount === 5);

  const key = dash.weekKey(new Date("2026-09-02T10:00:00+02:00"));
  assert.ok(/^\d{4}-W\d{2}$/.test(key));
  assert.strictEqual(dash.monthKey(new Date("2026-09-02T10:00:00+02:00")), "2026-09");
} finally {
  db.listEnquiries = orig;
}

console.log("enquiry-dashboard.test.js ok");
