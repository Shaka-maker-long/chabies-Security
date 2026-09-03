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
assert.ok(enquiry.events.some((ev) => ev.kind === "created"));
assert.ok(enquiry.opened_at_label);
assert.ok(enquiry.lifespan_label);

assert.throws(
  () => pipeline.applyAction("#1996", "Coster", { action: "assign_costing", assignee: "Welder" }),
  /office Admin/
);

const costing = pipeline.applyAction("#1996", "Coster", {
  action: "assign_costing",
  assignee: "Coster"
});
assert.strictEqual(costing.row.status, "Costing");
assert.ok(costing.row.events.some((ev) => ev.kind === "assign_costing" && ev.at));
assert.ok(costing.row.events.some((ev) => ev.kind === "assign_costing" && ev.actor === "Coster"));
assert.ok(costing.actions.some((a) => a.id === "assign_costing" && a.label === "Change costing person"));
assert.ok(costing.actions.some((a) => a.id === "add_correspondence"));
const myCost = pipeline.listMyTasks("Coster");
assert.strictEqual(myCost.length, 1);
assert.strictEqual(myCost[0].kind, "cost_sheet");
assert.strictEqual(myCost[0].correspondence_mails, 0);
assert.strictEqual(pipeline.listMyTasks("Quoter").length, 0);

const moved = pipeline.applyAction("#1996", "Coster", { action: "assign_costing", assignee: "Approver" });
assert.ok(pipeline.listMyTasks("Approver").some((t) => t.kind === "cost_sheet"));
assert.ok(!pipeline.listMyTasks("Coster").some((t) => t.kind === "cost_sheet"));
const self = pipeline.applyAction("#1996", "Approver", { action: "assign_costing", assignee: "Approver" });
assert.ok(self.row.tasks.some((t) => t.kind === "cost_sheet" && t.status === "open" && t.assignee === "Approver"));
pipeline.applyAction("#1996", "Approver", { action: "assign_costing", assignee: "Coster" });
const withCopy = pipeline.applyAction("#1996", "Coster", {
  action: "assign_costing",
  assignee: "Coster",
  correspondence_mails: [{
    subject: "Re: Order #S260025 - LAUGE SORENSEN",
    from: "Studio Delta",
    rest_id: "AAMkADrestid025",
    internet_message_id: "<s260025@studio-delta.test>"
  }]
});
assert.strictEqual(withCopy.row.correspondence.mails.length, 1);
assert.ok(!withCopy.row.correspondence.mails[0].stored_as);
assert.ok(withCopy.row.correspondence.mails[0].outlook_url.indexOf("ms-outlook://") === 0);
assert.strictEqual(withCopy.row.correspondence.mails[0].order_no, "S260025");
assert.strictEqual(withCopy.row.correspondence.mails[0].customer, "LAUGE SORENSEN");
assert.ok(!db.readEnquiryAttachment("#1996", "correspondence_1"));
const extraMail = pipeline.applyAction("#1996", "Coster", {
  action: "add_correspondence",
  correspondence_links: "https://outlook.office.com/owa/?ItemID=AAMkADrestid023&exvsurl=1&viewmodel=ReadMessageItem"
});
assert.strictEqual(extraMail.row.correspondence.mails.length, 2);
assert.ok(extraMail.row.correspondence.mails[1].outlook_url.indexOf("https://outlook.office.com/") === 0);
assert.ok(extraMail.row.correspondence.mails[1].outlook_url.indexOf("ms-outlook://") === -1);
assert.ok(extraMail.row.correspondence.mails[1].outlook_url.indexOf("AAMkADrestid023") >= 0);
const serverLink = pipeline.applyAction("#1996", "Coster", {
  action: "add_correspondence",
  correspondence_links: "https://chabies-security-production.up.railway.app/api/office/enquiries/%231996/files/quote"
});
assert.ok(serverLink.row.correspondence.mails.some((m) => /railway\.app/.test(m.outlook_url || "") && m.title === "Correspondance link"));
assert.ok(!serverLink.row.correspondence.mails.some((m) => /railway\.app/.test(m.outlook_url || "") && /^ms-outlook:/i.test(m.outlook_url || "")));
assert.throws(
  () => pipeline.applyAction("#1996", "Coster", {
    action: "add_correspondence",
    correspondence_files: [{ file_base64: "data:application/octet-stream;base64,QQ==", file_name: "note.msg" }]
  }),
  /Paste the Correspondance link|Paste the email|subject appears/
);
const fromDrop = pipeline.applyAction("#1996", "Coster", {
  action: "add_correspondence",
  correspondence_files: [{
    file_base64: "data:application/octet-stream;base64," + Buffer.from("Message-ID: <s260023@studio-delta.test>\r\nSubject: Re: Order #S260023 - GERNOT CANTO\r\nFrom: Office\r\n").toString("base64"),
    file_name: "Re_ Order #S260023 - GERNOT CANTO.msg"
  }]
});
assert.strictEqual(fromDrop.row.correspondence.mails.length, 4);
assert.strictEqual(fromDrop.row.correspondence.mails[3].order_no, "S260023");
assert.ok(fromDrop.row.correspondence.mails[3].stored_as);
assert.ok(db.readEnquiryAttachment("#1996", fromDrop.row.correspondence.mails[3].kind));
const both = pipeline.applyAction("#1996", "Coster", {
  action: "add_correspondence",
  correspondence_mails: [{ title: "P21136 - District 6 Phase 5.msg" }],
  correspondence_files: [{
    file_base64: "data:application/octet-stream;base64," + Buffer.from("Message-ID: <p21136@studio-delta.test>\r\nSubject: P21136 - District 6 Phase 5\r\nFrom: Office\r\n").toString("base64"),
    file_name: "P21136 - District 6 Phase 5.msg"
  }]
});
const district = both.row.correspondence.mails.filter((m) => String(m.title || m.filename || "").indexOf("P21136") >= 0);
assert.strictEqual(district.length, 1);
assert.ok(district[0].stored_as);
assert.ok(district[0].kind);
assert.ok(db.readEnquiryAttachment("#1996", district[0].kind));
const titleOnlyFirst = pipeline.applyAction("#1996", "Coster", {
  action: "add_correspondence",
  correspondence_mails: [{ title: "Ready to save subject.msg" }]
});
assert.ok(!titleOnlyFirst.row.correspondence.mails.find((m) => /Ready to save subject/i.test(m.title || "")).stored_as);
const upgradeDrop = pipeline.applyAction("#1996", "Coster", {
  action: "add_correspondence",
  correspondence_files: [{
    file_base64: "data:application/octet-stream;base64," + Buffer.alloc(80, 65).toString("base64"),
    file_name: "Ready to save subject.msg"
  }]
});
const upgraded = upgradeDrop.row.correspondence.mails.filter((m) => /Ready to save subject/i.test(m.title || m.filename || ""));
assert.strictEqual(upgraded.length, 1);
assert.ok(upgraded[0].stored_as);
assert.ok(upgraded[0].kind);
assert.ok(db.readEnquiryAttachment("#1996", upgraded[0].kind));

assert.throws(
  () => pipeline.applyAction("#1996", "Coster", { action: "assign_waiting", waiting_status: "Waiting on clients personal details", assignee: "Coster" }),
  /capture/i
);

const csv = "data:text/csv;base64," + Buffer.from("item,value\nMirror,2500\n").toString("base64");
assert.throws(
  () => pipeline.applyAction("#1996", "Coster", {
    action: "complete_cost_sheet",
    file_base64: csv,
    file_name: "cost.csv",
    file_confirmed: false,
    assignee: "Approver"
  }),
  /Tick|confirm|correct file/i
);

const costed = pipeline.applyAction("#1996", "Coster", {
  action: "complete_cost_sheet",
  file_base64: csv,
  file_name: "cost.csv",
  file_confirmed: true,
  assignee: "Approver",
  quote_assignee: "Quoter"
});
assert.strictEqual(costed.row.status, "Costed");
assert.strictEqual(costed.row.delivery_excl_vat, "");
assert.strictEqual(costed.row.products[0].value_excl_vat, "");
assert.strictEqual(costed.row.quote_assignee, "Quoter");
assert.ok(costed.row.cost_sheet && costed.row.cost_sheet.stored_as);
assert.strictEqual(pipeline.listMyTasks("Approver")[0].kind, "approval");
assert.ok(!pipeline.listMyTasks("Coster").some((t) => t.kind === "cost_sheet"));
const doneCost = pipeline.listMyCompletedTasks("Coster");
assert.ok(doneCost.some((t) => t.kind === "cost_sheet" && t.status === "done"));
assert.ok(doneCost[0].completed_at);
assert.ok(doneCost[0].completed_at_label);
assert.ok(!doneCost.some((t) => t.status === "open"));

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
  file_base64: csv,
  file_name: "cost.csv",
  file_confirmed: true,
  assignee: "Approver",
  quote_assignee: "Quoter"
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
    follow_up_assignee: "Quoter",
    quote_no: "SOQ2361"
  }),
  /confirm/
);

assert.throws(
  () => pipeline.applyAction("#1996", "Quoter", {
    action: "complete_quote",
    products: [{ product: "Daphne Rectangular Mirror", category: "Mirror", value_excl_vat: "" }],
    delivery_excl_vat: "",
    file_base64: pdfB64,
    file_name: "quote.pdf",
    file_confirmed: true,
    follow_up_assignee: "Quoter"
  }),
  /value|Delivery/i
);

assert.throws(
  () => pipeline.applyAction("#1996", "Quoter", {
    action: "complete_quote",
    products: [
      { product: "Daphne Rectangular Mirror", category: "Mirror", value_excl_vat: "2500" },
      { product: "Eve Patio Table", category: "Table", value_excl_vat: "1800.5" }
    ],
    delivery_excl_vat: "350",
    file_base64: pdfB64,
    file_name: "quote.pdf",
    file_confirmed: true,
    follow_up_assignee: "Quoter"
  }),
  /quotation number/i
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
  follow_up_assignee: "Quoter",
  quote_no: "SOQ2361"
});
assert.strictEqual(quoted.row.status, "Quoted");
assert.strictEqual(quoted.row.has_quote_pdf, true);
assert.ok(quoted.row.date_quoted);
assert.strictEqual(quoted.row.quote_no, "SOQ2361");
assert.strictEqual(quoted.quoteNo.next, "SOQ2362");
assert.ok(quoted.quoteNo.recent.indexOf("SOQ2361") >= 0);
assert.strictEqual(quoted.row.quote_total_excl_vat, "4650.50");
assert.strictEqual(quoted.row.products[0].value_incl_vat, "2875.00");
assert.strictEqual(quoted.row.delivery_incl_vat, "402.50");
assert.strictEqual(quoted.row.quote_total_incl_vat, "5348.07");
const fromIncl = db.normalizeEnquiryLines({
  products: [{ product: "Eve Patio Table", category: "Table", value_incl_vat: "1150" }]
}, null);
assert.strictEqual(fromIncl[0].value_excl_vat, "1000.00");
assert.strictEqual(fromIncl[0].value_incl_vat, "1150.00");
assert.ok(pipeline.listMyTasks("Quoter").some((t) => t.kind === "follow_up"));
assert.strictEqual(quoted.row.quotes.length, 1);
assert.strictEqual(quoted.row.quotes[0].quote_no, "SOQ2361");
assert.ok(quoted.row.quotes[0].file && quoted.row.quotes[0].file.stored_as);
assert.ok(quoted.row.deliverables.some((d) => d.group === "quote"));
assert.ok(db.readEnquiryAttachment("#1996", "quote_1"));
assert.ok(quoted.actions.some((a) => a.id === "complete_quote" && a.label === "Issue another quote"));
assert.ok(quoted.actions.some((a) => a.id === "assign_costing" && /recost/i.test(a.label)));
assert.ok(db.readEnquiryAttachment("#1996", "quote_1"));
assert.throws(
  () => pipeline.applyAction("#1996", "Quoter", {
    action: "complete_quote",
    products: [
      { product: "Daphne Rectangular Mirror", category: "Mirror", value_excl_vat: "2500" },
      { product: "Eve Patio Table", category: "Table", value_excl_vat: "1800.5" }
    ],
    delivery_excl_vat: "350",
    file_base64: pdfB64,
    file_name: "quote-2.pdf",
    file_confirmed: true,
    follow_up_assignee: "Quoter",
    quote_no: "SOQ2361"
  }),
  /already used/i
);
const requoted = pipeline.applyAction("#1996", "Quoter", {
  action: "complete_quote",
  products: [
    { product: "Daphne Rectangular Mirror", category: "Mirror", value_excl_vat: "2800" },
    { product: "Eve Patio Table", category: "Table", value_excl_vat: "1800.5" }
  ],
  delivery_excl_vat: "350",
  file_base64: pdfB64,
  file_name: "quote-revised.pdf",
  file_confirmed: true,
  follow_up_assignee: "Quoter",
  quote_no: "SOQ2362"
});
assert.strictEqual(requoted.row.status, "Quoted");
assert.strictEqual(requoted.row.quote_no, "SOQ2362");
assert.strictEqual(requoted.row.quotes.length, 2);
assert.strictEqual(requoted.row.quotes[0].quote_no, "SOQ2361");
assert.strictEqual(requoted.row.quotes[1].quote_no, "SOQ2362");
assert.strictEqual(requoted.row.quote_total_excl_vat, "4950.50");
assert.ok(db.readEnquiryAttachment("#1996", "quote_1"));
assert.ok(db.readEnquiryAttachment("#1996", "quote_2"));
assert.ok(requoted.row.deliverables.filter((d) => d.group === "quote").length >= 2);
assert.strictEqual(requoted.quoteNo.next, "SOQ2363");

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
const files = db.listEnquiryDeliverables(drawn.row);
const groups = files.map((f) => f.group);
["correspondence", "cost_sheet", "quote", "follow_up", "pop", "drawing"].forEach((g) => {
  assert.ok(groups.indexOf(g) >= 0, "missing deliverable " + g);
});
assert.ok(files.filter((f) => f.open).every((f) => f.kind || f.href), "open files need a server path or a link");
assert.ok(files.every((f) => f.outlook === false), "do not treat files as Outlook attachments");
assert.ok(files.some((f) => f.group === "correspondence" && f.href && f.title === "Correspondance link"));
assert.ok(drawn.row.deliverable_count >= 6);
assert.ok(db.getEnquiry("#1996").deliverable_count >= 6);
const emptyNew = db.upsertEnquiry({
  date_enquired: "07/01/2026",
  client_name: "No files yet",
  product: "Air Chair",
  status: "New"
});
assert.strictEqual(emptyNew.deliverable_count, 0);
assert.deepStrictEqual(emptyNew.deliverables, []);

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

const skip = db.upsertEnquiry({
  date_enquired: "06/01/2026",
  client_name: "Skip Approval",
  product: "Eve Patio Table",
  category: "Table",
  status: "New"
});
pipeline.applyAction(skip.enquiry_no, "Coster", { action: "assign_costing", assignee: "Coster" });
assert.throws(
  () => pipeline.applyAction(skip.enquiry_no, "Coster", {
    action: "complete_cost_sheet",
    file_base64: csv,
    file_name: "cost.csv",
    file_confirmed: true
  }),
  /quoting person/i
);
const skipped = pipeline.applyAction(skip.enquiry_no, "Coster", {
  action: "complete_cost_sheet",
  file_base64: csv,
  file_name: "cost.csv",
  file_confirmed: true,
  quote_assignee: "Quoter"
});
assert.strictEqual(skipped.row.status, "Costed");
assert.strictEqual(skipped.row.approval.status, "approved");
assert.strictEqual(skipped.row.approval.comments, "Approval skipped");
assert.strictEqual(skipped.row.quote_assignee, "Quoter");
assert.ok(pipeline.listMyTasks("Quoter").some((t) => t.kind === "quote" && t.enquiry_no === skip.enquiry_no));
assert.ok(!pipeline.listMyTasks("Approver").some((t) => t.kind === "approval" && t.enquiry_no === skip.enquiry_no));

assert.throws(
  () => pipeline.applyAction(skip.enquiry_no, "Quoter", {
    action: "complete_quote",
    products: [{ product: "Eve Patio Table", category: "Table", value_excl_vat: "1000" }],
    delivery_excl_vat: "100",
    file_base64: pdfB64,
    file_name: "quote.pdf",
    file_confirmed: true,
    follow_up_assignee: "Quoter",
    quote_no: "SOQ2361"
  }),
  /already used/i
);
assert.throws(
  () => pipeline.applyAction(skip.enquiry_no, "Quoter", {
    action: "complete_quote",
    products: [{ product: "Eve Patio Table", category: "Table", value_excl_vat: "1000" }],
    delivery_excl_vat: "100",
    file_base64: pdfB64,
    file_name: "quote.pdf",
    file_confirmed: true,
    follow_up_assignee: "Quoter",
    quote_no: "soq2361"
  }),
  /already used/i
);
const skipQuoted = pipeline.applyAction(skip.enquiry_no, "Quoter", {
  action: "complete_quote",
  products: [{ product: "Eve Patio Table", category: "Table", value_excl_vat: "1000" }],
  delivery_excl_vat: "100",
  file_base64: pdfB64,
  file_name: "quote.pdf",
  file_confirmed: true,
  follow_up_assignee: "Quoter",
  quote_no: "SOQ2400"
});
assert.strictEqual(skipQuoted.row.quote_no, "SOQ2400");
assert.strictEqual(skipQuoted.quoteNo.next, "SOQ2401");
assert.ok(skipQuoted.quoteNo.recent.indexOf("SOQ2361") >= 0);
assert.ok(skipQuoted.quoteNo.recent.indexOf("SOQ2362") >= 0);
assert.ok(skipQuoted.quoteNo.recent.indexOf("SOQ2400") >= 0);
assert.strictEqual(skipQuoted.row.quotes.length, 1);

const recostSkip = pipeline.applyAction(skip.enquiry_no, "Quoter", {
  action: "assign_costing",
  assignee: "Coster"
});
assert.strictEqual(recostSkip.row.status, "Re-Cost");
assert.ok(pipeline.listMyTasks("Coster").some((t) => t.kind === "cost_sheet" && t.enquiry_no === skip.enquiry_no));
assert.ok(!pipeline.listMyTasks("Quoter").some((t) => t.kind === "follow_up" && t.enquiry_no === skip.enquiry_no));
pipeline.applyAction(skip.enquiry_no, "Coster", {
  action: "complete_cost_sheet",
  file_base64: csv,
  file_name: "cost.csv",
  file_confirmed: true,
  quote_assignee: "Quoter"
});
const skipRequote = pipeline.applyAction(skip.enquiry_no, "Quoter", {
  action: "complete_quote",
  products: [{ product: "Eve Patio Table", category: "Table", value_excl_vat: "1200" }],
  delivery_excl_vat: "100",
  file_base64: pdfB64,
  file_name: "quote-2.pdf",
  file_confirmed: true,
  follow_up_assignee: "Quoter",
  quote_no: "SOQ2401"
});
assert.strictEqual(skipRequote.row.quote_no, "SOQ2401");
assert.strictEqual(skipRequote.row.quotes.length, 2);
assert.strictEqual(skipRequote.row.quotes[0].quote_no, "SOQ2400");
assert.ok(db.readEnquiryAttachment(skip.enquiry_no, "quote_1"));
assert.ok(db.readEnquiryAttachment(skip.enquiry_no, "quote_2"));

staff.upsertUser({ name: "Coster", access: "Admin", role: "Admin", password: "x", seeDebtors: "Yes", enquiryRoles: ["Costing"] });
staff.upsertUser({ name: "Approver", access: "Admin", role: "Admin", password: "x", seeDebtors: "Yes", enquiryRoles: ["Approval"] });
staff.upsertUser({ name: "Quoter", access: "Admin", role: "Admin", password: "x", seeDebtors: "Yes", enquiryRoles: ["Quoting"] });
assert.strictEqual(staff.defaultEnquiryAssignee("Costing"), "Coster");
assert.strictEqual(staff.defaultEnquiryAssignee("Quoting"), "Quoter");
assert.strictEqual(staff.defaultEnquiryAssignee("Approval"), "Approver");
const roleSnap = pipeline.processSnapshot(skip.enquiry_no);
assert.strictEqual(roleSnap.enquiryRoles.costing, "Coster");
assert.strictEqual(roleSnap.enquiryRoles.quoting, "Quoter");
assert.strictEqual(roleSnap.enquiryRoles.approval, "Approver");

const roleEnq = db.upsertEnquiry({
  date_enquired: "02/09/2026",
  client_name: "Role Defaults",
  product: "Gate",
  category: "Gate"
});
const autoCost = pipeline.applyAction(roleEnq.enquiry_no, "Coster", { action: "assign_costing" });
assert.strictEqual(autoCost.row.status, "Costing");
assert.ok(pipeline.listMyTasks("Coster").some((t) => t.kind === "cost_sheet" && t.enquiry_no === roleEnq.enquiry_no));
const autoSheet = pipeline.applyAction(roleEnq.enquiry_no, "Coster", {
  action: "complete_cost_sheet",
  file_base64: csv,
  file_name: "cost.csv",
  file_confirmed: true
});
assert.strictEqual(autoSheet.row.quote_assignee, "Quoter");
assert.strictEqual(autoSheet.row.approval.requested_from, "Approver");
assert.ok(pipeline.listMyTasks("Approver").some((t) => t.kind === "approval" && t.enquiry_no === roleEnq.enquiry_no));
const autoApprove = pipeline.applyAction(roleEnq.enquiry_no, "Approver", {
  action: "complete_approval",
  decision: "approve"
});
assert.ok(pipeline.listMyTasks("Quoter").some((t) => t.kind === "quote" && t.enquiry_no === roleEnq.enquiry_no));
assert.strictEqual(autoApprove.row.quote_assignee, "Quoter");

const slashPaste = pipeline.applyAction(roleEnq.enquiry_no, "Coster", {
  action: "add_correspondence",
  correspondence_links: "https://files.example/jobs//Q1//mirror.xlsx\n\\\\nas\\studio-delta\\cost//sheet.xlsx"
});
assert.ok(slashPaste.row.correspondence.mails.some((m) => m.outlook_url === "https://files.example/jobs//Q1//mirror.xlsx"));
assert.ok(slashPaste.row.correspondence.mails.some((m) => m.outlook_url === "\\\\nas\\studio-delta\\cost//sheet.xlsx"));

const multi = db.upsertEnquiry({
  date_enquired: "08/01/2026",
  client_name: "Two Products",
  products: [
    { product: "Daphne Rectangular Mirror", category: "Mirror" },
    { product: "Eve Patio Table", category: "Table" }
  ],
  status: "New"
});
pipeline.applyAction(multi.enquiry_no, "Coster", { action: "assign_costing", assignee: "Coster" });
assert.throws(
  () => pipeline.applyAction(multi.enquiry_no, "Coster", {
    action: "complete_cost_sheet",
    file_base64: csv,
    file_name: "cost.csv",
    file_confirmed: true,
    quote_assignee: "Quoter"
  }),
  /Eve Patio Table/
);
assert.throws(
  () => pipeline.applyAction(multi.enquiry_no, "Coster", {
    action: "complete_cost_sheet",
    cost_sheets: [{
      product: "Daphne Rectangular Mirror",
      files: [{ file_base64: csv, file_name: "mirror.csv", file_confirmed: true }]
    }],
    quote_assignee: "Quoter"
  }),
  /Eve Patio Table/
);
const multiCost = pipeline.applyAction(multi.enquiry_no, "Coster", {
  action: "complete_cost_sheet",
  cost_sheets: [
    {
      product: "Daphne Rectangular Mirror",
      files: [
        { file_base64: csv, file_name: "mirror.csv", file_confirmed: true },
        { file_base64: csv, file_name: "mirror-extra.csv", file_confirmed: true }
      ]
    },
    {
      product: "Eve Patio Table",
      files: [{ file_base64: csv, file_name: "table.csv", file_confirmed: true }]
    }
  ],
  quote_assignee: "Quoter"
});
assert.strictEqual(multiCost.row.status, "Costed");
assert.strictEqual(multiCost.row.cost_sheets.length, 3);
assert.strictEqual(multiCost.row.cost_sheets.filter((s) => s.product === "Daphne Rectangular Mirror").length, 2);
assert.strictEqual(multiCost.row.cost_sheets.filter((s) => s.product === "Eve Patio Table").length, 1);
assert.ok(db.readEnquiryAttachment(multi.enquiry_no, "cost_sheet"));
assert.ok(db.readEnquiryAttachment(multi.enquiry_no, "cost_sheet_1"));
assert.ok(db.readEnquiryAttachment(multi.enquiry_no, "cost_sheet_2"));
assert.ok(db.readEnquiryAttachment(multi.enquiry_no, "cost_sheet_3"));
const multiFiles = db.listEnquiryDeliverables(multiCost.row);
assert.strictEqual(multiFiles.filter((f) => f.group === "cost_sheet").length, 3);
assert.ok(multiFiles.some((f) => f.label === "Cost sheet · Daphne Rectangular Mirror"));
assert.ok(multiFiles.some((f) => f.label === "Cost sheet · Eve Patio Table"));

console.log("enquiry-pipeline.test.js ok");
