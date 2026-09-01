const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-pipe-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const staff = require("./staff");
const db = require("./db");
const pipeline = require("./enquiry-pipeline");

initWorkbook();
staff.upsertUser({ name: "Coster", access: "Admin", role: "Admin", password: "x", seeDebtors: "Yes" });
staff.upsertUser({ name: "Approver", access: "Admin", role: "Admin", password: "x", seeDebtors: "Yes" });
staff.upsertUser({ name: "Quoter", access: "Admin", role: "Admin", password: "x", seeDebtors: "Yes" });
staff.upsertUser({ name: "Welder", access: "Production", role: "Welding", password: "x", tasks: ["Welding"] });

assert.ok(pipeline.officeAssignees().indexOf("Coster") >= 0);
assert.ok(pipeline.officeAssignees().indexOf("Welder") === -1);
assert.throws(() => pipeline.applyAction("#1996", "Coster", { action: "assign_costing", assignee: "Coster" }), /not found/i);

const enquiry = db.upsertEnquiry({
  date_enquired: "30/11/2025",
  client_name: "Michael Cost",
  product: "Daphne Rectangular Mirror",
  category: "Mirror",
  status: "Quoted"
});
assert.strictEqual(enquiry.enquiry_no, "#1996");
assert.strictEqual(enquiry.status, "New");
assert.strictEqual(enquiry.products[0].value_excl_vat, "");

assert.throws(
  () => pipeline.applyAction("#1996", "Coster", { action: "assign_costing", assignee: "Welder" }),
  /office Admin/
);

const costing = pipeline.applyAction("#1996", "Coster", { action: "assign_costing", assignee: "Coster" });
assert.strictEqual(costing.row.status, "Costing");
assert.ok(costing.actions.some((a) => a.id === "assign_costing" && a.label === "Change costing person"));
const myCost = pipeline.listMyTasks("Coster");
assert.strictEqual(myCost.length, 1);
assert.strictEqual(myCost[0].kind, "cost_sheet");
assert.strictEqual(pipeline.listMyTasks("Quoter").length, 0);

const moved = pipeline.applyAction("#1996", "Coster", { action: "assign_costing", assignee: "Approver" });
assert.ok(pipeline.listMyTasks("Approver").some((t) => t.kind === "cost_sheet"));
assert.ok(!pipeline.listMyTasks("Coster").some((t) => t.kind === "cost_sheet"));
const self = pipeline.applyAction("#1996", "Approver", { action: "assign_costing", assignee: "Approver" });
assert.ok(self.row.tasks.some((t) => t.kind === "cost_sheet" && t.status === "open" && t.assignee === "Approver"));
pipeline.applyAction("#1996", "Approver", { action: "assign_costing", assignee: "Coster" });

assert.throws(
  () => pipeline.applyAction("#1996", "Coster", { action: "assign_waiting", waiting_status: "Waiting on clients personal details", assignee: "Coster" }),
  /capture/i
);

const csv = "data:text/csv;base64," + Buffer.from("item,value\nMirror,2500\n").toString("base64");
assert.throws(
  () => pipeline.applyAction("#1996", "Coster", {
    action: "complete_cost_sheet",
    products: [{ product: "Daphne Rectangular Mirror", category: "Mirror", value_excl_vat: "2500" }],
    delivery_excl_vat: "",
    file_base64: csv,
    file_name: "cost.csv",
    file_confirmed: true,
    assignee: "Approver"
  }),
  /Delivery/
);

const costed = pipeline.applyAction("#1996", "Coster", {
  action: "complete_cost_sheet",
  products: [{ product: "Daphne Rectangular Mirror", category: "Mirror", value_excl_vat: "2500" }],
  delivery_excl_vat: "350",
  file_base64: csv,
  file_name: "cost.csv",
  file_confirmed: true,
  assignee: "Approver"
});
assert.strictEqual(costed.row.status, "Costed");
assert.strictEqual(costed.row.delivery_excl_vat, "350.00");
assert.ok(costed.row.cost_sheet && costed.row.cost_sheet.stored_as);
assert.strictEqual(pipeline.listMyTasks("Approver")[0].kind, "approval");

assert.throws(
  () => pipeline.applyAction("#1996", "Approver", { action: "complete_approval", decision: "reject", assignee: "Coster" }),
  /Comments/
);

const recost = pipeline.applyAction("#1996", "Approver", {
  action: "complete_approval",
  decision: "reject",
  comments: "Supplier price is wrong",
  assignee: "Coster"
});
assert.strictEqual(recost.row.status, "Re-Cost");
assert.ok(pipeline.listMyTasks("Coster").some((t) => t.kind === "cost_sheet"));

pipeline.applyAction("#1996", "Coster", {
  action: "complete_cost_sheet",
  products: [{ product: "Daphne Rectangular Mirror", category: "Mirror", value_excl_vat: "2500" }],
  delivery_excl_vat: "350",
  file_base64: csv,
  file_name: "cost.csv",
  file_confirmed: true,
  assignee: "Approver"
});

const approved = pipeline.applyAction("#1996", "Approver", {
  action: "complete_approval",
  decision: "approve",
  assignee: "Quoter"
});
assert.strictEqual(approved.row.approval.status, "approved");
assert.ok(pipeline.listMyTasks("Quoter").some((t) => t.kind === "quote"));

const pdf = Buffer.from("%PDF-1.1\n1 0 obj\n<<>>\nendobj\ntrailer\n<<>>\n%%EOF\n");
const pdfB64 = "data:application/pdf;base64," + pdf.toString("base64");
assert.throws(
  () => pipeline.applyAction("#1996", "Quoter", {
    action: "complete_quote",
    products: [{ product: "Daphne Rectangular Mirror", category: "Mirror", value_excl_vat: "2500" }],
    delivery_excl_vat: "350",
    file_base64: pdfB64,
    file_name: "quote.pdf",
    file_confirmed: false,
    follow_up_assignee: "Quoter"
  }),
  /confirm/
);

const quoted = pipeline.applyAction("#1996", "Quoter", {
  action: "complete_quote",
  products: [
    { product: "Daphne Rectangular Mirror", category: "Mirror", value_excl_vat: "2500" },
    { product: "Eve Patio Table", category: "Table", value_excl_vat: "1800.5" }
  ],
  delivery_excl_vat: "350",
  file_base64: pdfB64,
  file_name: "Michael Cost quote.pdf",
  file_confirmed: true,
  follow_up_assignee: "Quoter"
});
assert.strictEqual(quoted.row.status, "Quoted");
assert.strictEqual(quoted.row.has_quote_pdf, true);
assert.ok(quoted.row.date_quoted);
assert.strictEqual(quoted.row.quote_total_excl_vat, "4650.50");
assert.ok(pipeline.listMyTasks("Quoter").some((t) => t.kind === "follow_up"));

const png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const followed = pipeline.applyAction("#1996", "Quoter", {
  action: "complete_followup",
  file_base64: png,
  file_name: "whatsapp.png",
  file_confirmed: true,
  assignee: "Quoter"
});
assert.strictEqual(followed.row.status, "Followed Up");
assert.strictEqual(followed.row.follow_ups[0].label, "Follow up");

assert.throws(
  () => pipeline.applyAction("#1996", "Quoter", { action: "complete_reject" }),
  /reason/
);

const ordered = pipeline.applyAction("#1996", "Quoter", {
  action: "complete_order",
  file_base64: pdfB64,
  file_name: "pop.pdf",
  file_confirmed: true,
  drawing_required: "yes",
  assignee: "Coster"
});
assert.strictEqual(ordered.row.status, "Ordered");
assert.strictEqual(ordered.row.ready_for_orders, false);
assert.ok(pipeline.listMyTasks("Coster").some((t) => t.kind === "drawing"));

const drawn = pipeline.applyAction("#1996", "Coster", {
  action: "complete_drawing",
  file_base64: pdfB64,
  file_name: "drawing.pdf",
  file_confirmed: true
});
assert.strictEqual(drawn.row.ready_for_orders, true);
assert.ok(db.readEnquiryAttachment("#1996", "drawing"));
assert.ok(db.readEnquiryAttachment("#1996", "pop"));
assert.ok(db.readEnquiryAttachment("#1996", "cost_sheet"));

const waiting = db.upsertEnquiry({
  date_enquired: "05/01/2026",
  client_name: "Sas Promotions",
  product: "Air Chair",
  status: "New"
});
pipeline.applyAction(waiting.enquiry_no, "Coster", {
  action: "assign_waiting",
  waiting_status: "Waiting on clients specifictions",
  assignee: "Approver"
});
assert.strictEqual(db.getEnquiry(waiting.enquiry_no).status, "Waiting on clients specifictions");
assert.ok(pipeline.listMyTasks("Approver").some((t) => t.kind === "chase_info" && t.enquiry_no === waiting.enquiry_no));

assert.throws(
  () => pipeline.applyAction(waiting.enquiry_no, "Approver", { action: "supplier_wait", assignee: "Coster" }),
  /costing/
);

console.log("enquiry-pipeline.test.js ok");
