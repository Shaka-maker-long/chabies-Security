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

assert.throws(
  () => pipeline.applyAction("#1996", "Coster", { action: "assign_costing", assignee: "Coster" }),
  /Correspondance/
);
const costing = pipeline.applyAction("#1996", "Coster", {
  action: "assign_costing",
  assignee: "Coster",
  correspondence_mails: [{
    subject: "Re: Order #S260025 - LAUGE SORENSEN",
    from: "Studio Delta",
    rest_id: "AAMkADrestid025",
    internet_message_id: "<s260025@studio-delta.test>"
  }]
});
assert.strictEqual(costing.row.status, "Costing");
assert.ok(costing.row.events.some((ev) => ev.kind === "assign_costing" && ev.at));
assert.ok(costing.row.events.some((ev) => ev.kind === "assign_costing" && ev.actor === "Coster"));
assert.ok(costing.actions.some((a) => a.id === "assign_costing" && a.label === "Change costing person"));
assert.ok(costing.actions.some((a) => a.id === "add_correspondence" && a.label === "Save Correspondance link"));
const myCost = pipeline.listMyTasks("Coster");
assert.strictEqual(myCost.length, 1);
assert.strictEqual(myCost[0].kind, "cost_sheet");
assert.strictEqual(myCost[0].correspondence_mails, 1);
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
  /Enter a Correspondance link|Paste the email|subject appears/
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
pipeline.applyAction(skip.enquiry_no, "Coster", {
  action: "assign_costing",
  assignee: "Coster",
  correspondence_links: "https://files.example/skip"
});
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
const autoCost = pipeline.applyAction(roleEnq.enquiry_no, "Coster", {
  action: "assign_costing",
  correspondence_links: "https://files.example/role"
});
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
pipeline.applyAction(multi.enquiry_no, "Coster", {
  action: "assign_costing",
  assignee: "Coster",
  correspondence_links: "https://files.example/multi"
});
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

staff.upsertUser({ name: "Pat", access: "Admin", role: "Admin", password: "x", seeDebtors: "Yes" });
staff.upsertUser({ name: "Lesedi", access: "Admin", role: "Manager", password: "x", seeDebtors: "Yes" });
assert.ok(staff.canManageUsers({ name: "Lesedi" }));

const locked = db.upsertEnquiry({
  date_enquired: "08/01/2026",
  client_name: "Access Lock Client",
  product: "Air Chair",
  category: "Chair",
  status: "New"
});
pipeline.applyAction(locked.enquiry_no, "Coster", {
  action: "assign_costing",
  assignee: "Coster",
  correspondence_links: "https://files.example/locked"
});
assert.throws(
  () => pipeline.applyAction(locked.enquiry_no, "Pat", {
    action: "complete_cost_sheet",
    file_base64: csv,
    file_name: "cost.csv",
    file_confirmed: true,
    quote_assignee: "Quoter"
  }),
  /Coster is assigned/i
);
const asked = pipeline.applyAction(locked.enquiry_no, "Pat", {
  action: "request_access",
  for_action: "complete_cost_sheet",
  kind: "cost_sheet"
});
assert.ok(asked.access.mine.some((g) => g.kind === "cost_sheet" && g.status === "pending"));
assert.ok(pipeline.listAccessInbox("Coster").some((g) => g.enquiry_no === locked.enquiry_no && g.requester === "Pat"));
assert.ok(pipeline.listAccessInbox("Lesedi").some((g) => g.enquiry_no === locked.enquiry_no));
assert.throws(
  () => pipeline.applyAction(locked.enquiry_no, "Quoter", { action: "grant_access", grant_id: asked.access.mine[0].id }),
  /Manager/i
);
const granted = pipeline.applyAction(locked.enquiry_no, "Coster", {
  action: "grant_access",
  grant_id: asked.access.mine[0].id
});
assert.ok(granted.actions.some((a) => a.id === "complete_cost_sheet" && a.can_act));
const asPat = pipeline.processSnapshot(locked.enquiry_no, "Pat");
assert.ok(asPat.actions.some((a) => a.id === "complete_cost_sheet" && a.can_act));
pipeline.applyAction(locked.enquiry_no, "Pat", {
  action: "complete_cost_sheet",
  file_base64: csv,
  file_name: "cost.csv",
  file_confirmed: true,
  quote_assignee: "Quoter"
});
assert.strictEqual(db.getEnquiry(locked.enquiry_no).status, "Costed");

const managerActs = db.upsertEnquiry({
  date_enquired: "09/01/2026",
  client_name: "Manager Can Act",
  product: "Air Chair",
  category: "Chair",
  status: "New"
});
pipeline.applyAction(managerActs.enquiry_no, "Coster", {
  action: "assign_costing",
  assignee: "Coster",
  correspondence_links: "https://files.example/manager"
});
pipeline.applyAction(managerActs.enquiry_no, "Lesedi", {
  action: "complete_cost_sheet",
  file_base64: csv,
  file_name: "cost.csv",
  file_confirmed: true,
  quote_assignee: "Quoter"
});
assert.strictEqual(db.getEnquiry(managerActs.enquiry_no).status, "Costed");

const fromEnquiry = db.createOrderFromEnquiry("#1996");
assert.ok(fromEnquiry.row.order_number);
assert.strictEqual(fromEnquiry.row.enquiry_no, "#1996");
assert.strictEqual(fromEnquiry.row.client_name, "Michael Cost");
assert.strictEqual(fromEnquiry.existing, false);
assert.strictEqual(db.getEnquiry("#1996").order_number, fromEnquiry.row.order_number);
const again = db.createOrderFromEnquiry("#1996");
assert.strictEqual(again.existing, true);
assert.strictEqual(again.row.order_number, fromEnquiry.row.order_number);

const dups = db.findOpenEnquiryDuplicates({
  client_email: "lock@example.com",
  client_number: "0821112222"
});
assert.deepStrictEqual(dups, []);
db.upsertEnquiry({
  enquiry_no: locked.enquiry_no,
  client_name: "Access Lock Client",
  client_email: "lock@example.com",
  client_number: "082 111 2222",
  product: "Air Chair"
});
const hit = db.findOpenEnquiryDuplicates({
  client_email: "lock@example.com",
  client_number: "0821112222"
}, managerActs.enquiry_no);
assert.ok(hit.some((r) => r.enquiry_no === locked.enquiry_no));

assert.strictEqual(
  pipeline.classifyCapture({ client_name: "Name only" }),
  "Waiting on clients personal details"
);
assert.strictEqual(
  pipeline.classifyCapture({
    client_name: "Pat Client",
    client_email: "pat@example.com",
    province: "Gauteng"
  }),
  "Waiting on clients specifictions"
);
assert.strictEqual(
  pipeline.classifyCapture({
    client_name: "Pat Client",
    client_number: "0821110000",
    province: "Gauteng",
    enquiry_type: "Catologue",
    products: [{ product: "Air Chair" }]
  }),
  "Costing"
);
assert.strictEqual(
  pipeline.classifyCapture({
    client_name: "Pat Client",
    client_email: "pat@example.com",
    province: "Gauteng",
    enquiry_type: "Custom",
    products: [{ product: "Air Chair" }],
    custom_specs: []
  }),
  "Waiting on clients specifictions"
);
assert.strictEqual(
  pipeline.classifyCapture({
    client_name: "Pat Client",
    client_email: "pat@example.com",
    province: "Gauteng",
    enquiry_type: "Custom",
    products: [{ product: "Air Chair" }],
    custom_specs: [{ kind: "Dimensions", detail: "800 x 600" }]
  }),
  "Costing"
);
assert.strictEqual(
  pipeline.classifyCapture({
    client_name: "Pat Client",
    client_email: "pat@example.com",
    province: "Gauteng",
    enquiry_type: "New Design",
    products: [{ product: "Air Chair" }],
    design_description: "a table"
  }),
  "Waiting on clients specifictions"
);
assert.ok(pipeline.isAutoCaptureStatus("New"));
assert.ok(pipeline.isAutoCaptureStatus("Waiting on clients personal details"));
assert.ok(pipeline.isAutoCaptureStatus("Waiting on clients specifictions"));
assert.ok(!pipeline.isAutoCaptureStatus("Waiting on productions confirmation"));
assert.ok(!pipeline.isAutoCaptureStatus("Quoted"));
assert.ok(!pipeline.isAutoCaptureStatus("Costing"));

const thin = db.upsertEnquiry({
  date_enquired: "10/01/2026",
  client_name: "Name Only Route"
});
const thinRouted = pipeline.applyCaptureRoute(thin.enquiry_no, "Coster");
assert.strictEqual(thinRouted.row.status, "Waiting on clients personal details");
assert.ok(pipeline.listMyTasks("Coster").some((t) => t.kind === "chase_info" && t.enquiry_no === thin.enquiry_no));

db.upsertEnquiry({
  enquiry_no: thin.enquiry_no,
  date_enquired: "10/01/2026",
  client_name: "Name Only Route",
  client_email: "nameonly@example.com",
  province: "Gauteng"
});
const specRouted = pipeline.applyCaptureRoute(thin.enquiry_no, "Coster");
assert.strictEqual(specRouted.row.status, "Waiting on clients specifictions");
assert.ok(pipeline.listMyTasks("Coster").some((t) => t.kind === "chase_info" && t.enquiry_no === thin.enquiry_no));

db.upsertEnquiry({
  enquiry_no: thin.enquiry_no,
  date_enquired: "10/01/2026",
  client_name: "Name Only Route",
  client_email: "nameonly@example.com",
  province: "Gauteng",
  enquiry_type: "Catologue",
  products: [{ product: "Air Chair", category: "Chair" }]
});
assert.throws(
  () => pipeline.applyCaptureRoute(thin.enquiry_no, "Coster"),
  /Correspondance/
);
assert.strictEqual(db.getEnquiry(thin.enquiry_no).status, "Waiting on clients specifictions");
pipeline.applyAction(thin.enquiry_no, "Coster", {
  action: "add_correspondence",
  correspondence_links: "https://files.example/thin"
});
const costRouted = pipeline.applyCaptureRoute(thin.enquiry_no, "Coster");
assert.strictEqual(costRouted.row.status, "Costing");
assert.ok(pipeline.listMyTasks("Coster").some((t) => t.kind === "cost_sheet" && t.enquiry_no === thin.enquiry_no));
assert.ok(!pipeline.listMyTasks("Coster").some((t) => t.kind === "chase_info" && t.enquiry_no === thin.enquiry_no));

const prodHold = db.upsertEnquiry({
  date_enquired: "11/01/2026",
  client_name: "Prod Confirm",
  client_email: "prod@example.com",
  province: "Gauteng",
  enquiry_type: "Catologue",
  products: [{ product: "Air Chair", category: "Chair" }]
});
pipeline.applyAction(prodHold.enquiry_no, "Coster", {
  action: "assign_waiting",
  waiting_status: "Waiting on productions confirmation",
  assignee: "Coster"
});
const prodKept = pipeline.applyCaptureRoute(prodHold.enquiry_no, "Coster");
assert.strictEqual(prodKept.row.status, "Waiting on productions confirmation");

const quoteHold = db.upsertEnquiry({
  date_enquired: "12/01/2026",
  client_name: "Quoted Hold",
  client_email: "quoted@example.com",
  province: "Gauteng",
  enquiry_type: "Catologue",
  products: [{ product: "Air Chair", category: "Chair" }]
});
const quoteRaw = db.getEnquiryRaw(quoteHold.enquiry_no);
quoteRaw.status = "Quoted";
db.saveEnquiryRecord(quoteRaw);
const quoteKept = pipeline.applyCaptureRoute(quoteHold.enquiry_no, "Coster");
assert.strictEqual(quoteKept.row.status, "Quoted");

const supplierJob = db.upsertEnquiry({
  date_enquired: "13/01/2026",
  client_name: "Supplier Wait Client",
  enquiry_type: "Catologue",
  products: [{ product: "Air Chair", category: "Chair" }]
});
pipeline.applyAction(supplierJob.enquiry_no, "Coster", {
  action: "assign_costing",
  assignee: "Coster",
  correspondence_links: "https://files.example/supplier"
});
const waitingSupplier = pipeline.applyAction(supplierJob.enquiry_no, "Coster", {
  action: "supplier_wait",
  assignee: "Quoter"
});
assert.strictEqual(waitingSupplier.row.status, "Waiting on Supplier");
assert.ok(pipeline.listMyTasks("Quoter").some((t) => t.kind === "supplier" && t.enquiry_no === supplierJob.enquiry_no));
assert.ok(!pipeline.listMyTasks("Coster").some((t) => t.kind === "cost_sheet" && t.enquiry_no === supplierJob.enquiry_no));
assert.throws(
  () => pipeline.applyAction(supplierJob.enquiry_no, "Quoter", { action: "complete_supplier" }),
  /quotation from the supplier|Upload/
);
assert.throws(
  () => pipeline.applyAction(supplierJob.enquiry_no, "Quoter", {
    action: "complete_supplier",
    file_base64: csv,
    file_name: "supplier.xlsx",
    file_confirmed: false
  }),
  /correct file/
);
const supplierDone = pipeline.applyAction(supplierJob.enquiry_no, "Quoter", {
  action: "complete_supplier",
  file_base64: csv,
  file_name: "supplier.xlsx",
  file_confirmed: true
});
assert.strictEqual(supplierDone.row.status, "Costing");
assert.ok(!pipeline.listMyTasks("Quoter").some((t) => t.kind === "supplier" && t.enquiry_no === supplierJob.enquiry_no));
assert.ok(pipeline.listMyTasks("Coster").some((t) => t.kind === "cost_sheet" && t.enquiry_no === supplierJob.enquiry_no));
assert.ok(supplierDone.row.deliverables.some((d) => d.group === "supplier" && /supplier\.xlsx/i.test(d.filename || d.title || "")));
assert.ok(db.readEnquiryAttachment(supplierJob.enquiry_no, "supplier_1"));

const selfSupplier = db.upsertEnquiry({
  date_enquired: "14/01/2026",
  client_name: "Coster Waits Supplier",
  enquiry_type: "Catologue",
  products: [{ product: "Air Chair", category: "Chair" }]
});
pipeline.applyAction(selfSupplier.enquiry_no, "Coster", {
  action: "assign_costing",
  assignee: "Coster",
  correspondence_links: "https://files.example/self-supplier"
});
pipeline.applyAction(selfSupplier.enquiry_no, "Coster", { action: "supplier_wait", assignee: "Coster" });
assert.throws(
  () => pipeline.applyAction(selfSupplier.enquiry_no, "Coster", { action: "complete_supplier", assignee: "Coster" }),
  /quotation from the supplier|Upload/
);
const selfDone = pipeline.applyAction(selfSupplier.enquiry_no, "Coster", {
  action: "complete_supplier",
  file_base64: png,
  file_name: "supplier-quote.png",
  file_confirmed: true
});
assert.strictEqual(selfDone.row.status, "Costing");
assert.ok(pipeline.listMyTasks("Coster").some((t) => t.kind === "cost_sheet" && t.enquiry_no === selfSupplier.enquiry_no));
assert.ok(selfDone.row.deliverables.some((d) => d.group === "supplier"));

staff.upsertUser({ name: "Quoter", access: "Admin", role: "Admin", password: "x", seeDebtors: "Yes", enquiryRoles: ["Quoting", "Follow-up"] });
staff.upsertUser({ name: "Pat", access: "Admin", role: "Admin", password: "x", seeDebtors: "Yes", enquiryRoles: ["Follow-up"] });
assert.deepStrictEqual(staff.enquiryRoleHolders("Follow-up").sort(), ["Pat", "Quoter"]);

function toQuoted(clientName, quoteNo) {
  const enq = db.upsertEnquiry({
    date_enquired: "16/01/2026",
    client_name: clientName,
    enquiry_type: "Catologue",
    products: [{ product: "Air Chair", category: "Chair" }]
  });
  pipeline.applyAction(enq.enquiry_no, "Coster", {
    action: "assign_costing",
    assignee: "Coster",
    correspondence_links: "https://files.example/follow"
  });
  pipeline.applyAction(enq.enquiry_no, "Coster", {
    action: "complete_cost_sheet",
    file_base64: csv,
    file_name: "cost.csv",
    file_confirmed: true,
    quote_assignee: "Quoter"
  });
  pipeline.applyAction(enq.enquiry_no, "Approver", { action: "complete_approval", decision: "approve" });
  return pipeline.applyAction(enq.enquiry_no, "Quoter", {
    action: "complete_quote",
    products: [{ product: "Air Chair", category: "Chair", value_excl_vat: "1000" }],
    delivery_excl_vat: "100",
    file_base64: pdfB64,
    file_name: "quote.pdf",
    file_confirmed: true,
    quote_no: quoteNo
  });
}

const pooled = toQuoted("Follow Pool", "SOQ2501");
assert.ok(pipeline.listMyTasks("Quoter").some((t) => t.kind === "follow_up" && t.enquiry_no === pooled.row.enquiry_no));
assert.ok(pipeline.listMyTasks("Pat").some((t) => t.kind === "follow_up" && t.enquiry_no === pooled.row.enquiry_no));
const firstFollow = pipeline.applyAction(pooled.row.enquiry_no, "Pat", {
  action: "complete_followup",
  file_base64: png,
  file_name: "fu1.png",
  file_confirmed: true
});
assert.strictEqual(firstFollow.row.status, "Followed Up");
assert.strictEqual(pipeline.currentQuoteFollowUps(firstFollow.row).length, 1);
assert.ok(firstFollow.row.tasks.some((t) => t.kind === "follow_up" && t.status === "done" && t.assignee === "Pat"));
assert.ok(firstFollow.row.tasks.some((t) => t.kind === "follow_up" && t.status === "cancelled" && t.assignee === "Quoter"));
assert.ok(pipeline.listMyTasks("Quoter").some((t) => t.kind === "follow_up" && t.enquiry_no === pooled.row.enquiry_no));
assert.ok(pipeline.listMyTasks("Pat").some((t) => t.kind === "follow_up" && t.enquiry_no === pooled.row.enquiry_no));

const capped = toQuoted("Three Follows", "SOQ2502");
pipeline.applyAction(capped.row.enquiry_no, "Quoter", { action: "complete_followup", file_base64: png, file_name: "a.png", file_confirmed: true });
pipeline.applyAction(capped.row.enquiry_no, "Quoter", { action: "complete_followup", file_base64: png, file_name: "b.png", file_confirmed: true });
const third = pipeline.applyAction(capped.row.enquiry_no, "Pat", { action: "complete_followup", file_base64: png, file_name: "c.png", file_confirmed: true });
assert.strictEqual(third.row.status, "Followed Up");
assert.ok(pipeline.followUpsExhausted(third.row));
assert.ok(!pipeline.listMyTasks("Quoter").some((t) => t.kind === "follow_up" && t.enquiry_no === capped.row.enquiry_no));
assert.ok(!pipeline.listMyTasks("Pat").some((t) => t.kind === "follow_up" && t.enquiry_no === capped.row.enquiry_no));
assert.ok(!third.actions.some((a) => a.id === "complete_followup"));
assert.ok(third.actions.some((a) => a.id === "complete_quote"));
assert.throws(
  () => pipeline.applyAction(capped.row.enquiry_no, "Quoter", { action: "complete_followup", file_base64: png, file_name: "d.png", file_confirmed: true }),
  /3 follow-ups|another quote/
);
const resetQuote = pipeline.applyAction(capped.row.enquiry_no, "Quoter", {
  action: "complete_quote",
  products: [{ product: "Air Chair", category: "Chair", value_excl_vat: "1100" }],
  delivery_excl_vat: "100",
  file_base64: pdfB64,
  file_name: "quote-new.pdf",
  file_confirmed: true,
  quote_no: "SOQ2503"
});
assert.strictEqual(resetQuote.row.status, "Quoted");
assert.strictEqual(pipeline.currentQuoteFollowUps(resetQuote.row).length, 0);
assert.ok(!pipeline.followUpsExhausted(resetQuote.row));
assert.ok(pipeline.listMyTasks("Quoter").some((t) => t.kind === "follow_up" && t.enquiry_no === capped.row.enquiry_no));
assert.ok(pipeline.listMyTasks("Pat").some((t) => t.kind === "follow_up" && t.enquiry_no === capped.row.enquiry_no));
const afterNew = pipeline.applyAction(capped.row.enquiry_no, "Quoter", {
  action: "complete_followup",
  file_base64: png,
  file_name: "new-quote-fu.png",
  file_confirmed: true
});
assert.strictEqual(pipeline.currentQuoteFollowUps(afterNew.row).length, 1);
assert.strictEqual(afterNew.row.follow_ups.filter((f) => f.quote_no === "SOQ2503").length, 1);
assert.strictEqual(afterNew.row.follow_ups.filter((f) => f.quote_no === "SOQ2502").length, 3);

assert.throws(
  () => pipeline.onboardEnquiry("Coster", {
    enquiry_no: "#3101",
    date_enquired: "01/08/2026",
    client_name: "Old Costing",
    client_email: "old@example.com",
    province: "Gauteng",
    enquiry_type: "Catologue",
    product: "Air Chair",
    status: "Costing",
    costing_assignee: "Coster"
  }),
  /Correspondance/
);
assert.ok(!db.getEnquiryRaw("#3101"));

const quotedOn = pipeline.onboardEnquiry("Quoter", {
  enquiry_no: "#3102",
  date_enquired: "01/08/2026",
  date_quoted: "15/08/2026",
  client_name: "Brought Across Quoted",
  client_email: "quoted-onboard@example.com",
  province: "Western Cape",
  enquiry_type: "Catologue",
  products: [{ product: "Air Chair", category: "Chair", value_incl_vat: "2300" }],
  delivery_incl_vat: "115",
  status: "Quoted",
  quote_no: "SOQ2601",
  quote_assignee: "Quoter",
  correspondence_links: "https://files.example/onboard-quoted",
  file_base64: pdfB64,
  file_name: "old-quote.pdf",
  file_confirmed: true
});
assert.strictEqual(quotedOn.row.enquiry_no, "#3102");
assert.strictEqual(quotedOn.row.status, "Quoted");
assert.strictEqual(quotedOn.row.date_quoted, "15/08/2026");
assert.strictEqual(quotedOn.row.quote_no, "SOQ2601");
assert.ok(quotedOn.row.events.some((ev) => ev.kind === "onboard"));
assert.ok(pipeline.listMyTasks("Quoter").some((t) => t.kind === "follow_up" && t.enquiry_no === "#3102"));
assert.ok(pipeline.listMyTasks("Pat").some((t) => t.kind === "follow_up" && t.enquiry_no === "#3102"));
assert.ok(pipeline.isAutoCaptureStatus(quotedOn.row.status) === false);

const holdOn = pipeline.onboardEnquiry("Coster", {
  enquiry_no: "#3103",
  date_enquired: "02/08/2026",
  client_name: "Production Hold",
  client_email: "hold@example.com",
  client_number: "0820000000",
  province: "Gauteng",
  enquiry_type: "Catologue",
  product: "Air Chair",
  status: "Waiting on productions confirmation",
  chase_assignee: "Coster",
  correspondence_links: "https://files.example/hold"
});
assert.strictEqual(holdOn.row.status, "Waiting on productions confirmation");
assert.ok(holdOn.row.tasks.some((t) => t.kind === "chase_info" && t.status === "open" && t.assignee === "Coster"));
assert.ok(!holdOn.row.tasks.some((t) => t.kind === "cost_sheet" && t.status === "open"));

assert.throws(
  () => pipeline.onboardEnquiry("Coster", {
    enquiry_no: "#3104",
    date_enquired: "03/08/2026",
    client_name: "Two Cost Items",
    client_email: "twocost@example.com",
    province: "Gauteng",
    enquiry_type: "Catologue",
    products: [
      { product: "Air Chair", category: "Chair" },
      { product: "Eve Patio Table", category: "Table" }
    ],
    status: "Costed",
    quote_assignee: "Quoter",
    correspondence_links: "https://files.example/two-cost",
    cost_sheets: [{
      product: "Air Chair",
      files: [{ file_base64: csv, file_name: "chair.csv", file_confirmed: true }]
    }]
  }),
  /cost sheet/
);
assert.ok(!db.getEnquiryRaw("#3104"));
const twoCost = pipeline.onboardEnquiry("Coster", {
  enquiry_no: "#3105",
  date_enquired: "03/08/2026",
  client_name: "Two Cost Items",
  client_email: "twocost-ok@example.com",
  province: "Gauteng",
  enquiry_type: "Catologue",
  products: [
    { product: "Air Chair", category: "Chair" },
    { product: "Eve Patio Table", category: "Table" }
  ],
  status: "Costed",
  quote_assignee: "Quoter",
  correspondence_links: "https://files.example/two-cost-ok",
  cost_sheets: [
    { product: "Air Chair", files: [{ file_base64: csv, file_name: "chair.csv", file_confirmed: true }] },
    { product: "Eve Patio Table", files: [{ file_base64: csv, file_name: "table.csv", file_confirmed: true }] }
  ]
});
assert.strictEqual(twoCost.row.status, "Costed");
assert.strictEqual((twoCost.row.cost_sheets || []).length, 2);
assert.ok(twoCost.row.cost_sheets.some((s) => s.product === "Air Chair"));
assert.ok(twoCost.row.cost_sheets.some((s) => s.product === "Eve Patio Table"));

const keptOrder = db.upsertOrder({
  order_number: "9001",
  client_name: "Keep Me",
  status: "Not Yet Started"
});
assert.strictEqual(keptOrder.order_number, "9001");
assert.ok(db.listEnquiries().length >= 1);
const wiped = db.deleteAllEnquiries();
assert.ok(wiped >= 1);
assert.strictEqual(db.listEnquiries().length, 0);
assert.ok(db.listOrders().some((o) => String(o.order_number) === "9001"));
assert.strictEqual(db.nextEnquiryNo(), "#1996");

console.log("enquiry-pipeline.test.js ok");
