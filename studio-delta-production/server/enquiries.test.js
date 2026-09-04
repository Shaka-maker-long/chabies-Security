const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-enq-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";

const db = require("./db");

assert.strictEqual(db.nextEnquiryNo(), "#1996");
assert.strictEqual(db.monthFromEnquiryDate("30/11/2025"), "Nov");
assert.strictEqual(db.monthFromEnquiryDate("05/01/2026"), "Jan");

const first = db.upsertEnquiry({
  date_enquired: "30/11/2025",
  enquiry_source: "Website",
  enquiry_type: "Custom",
  client_name: "Michael Cost",
  product: "Daphne Rectangular Mirror",
  status: "New",
  custom_specs: [{ kind: "Dimensions", detail: "800 x 600" }]
});
assert.strictEqual(first.enquiry_no, "#1996");
assert.strictEqual(first.month_enquired, "Nov");
assert.strictEqual(first.products[0].product, "Daphne Rectangular Mirror");
assert.strictEqual(first.products[0].value_excl_vat, "");
assert.strictEqual(first.custom_specs[0].kind, "Dimensions");
assert.ok(first.request.indexOf("Dimensions") >= 0);
assert.ok(first.events.some((ev) => ev.kind === "created" && ev.at));
assert.ok(first.opened_at_label);
assert.ok(/open|to /.test(first.lifespan_label));
const same = db.upsertEnquiry({
  enquiry_no: first.enquiry_no,
  date_enquired: "30/11/2025",
  enquiry_source: "Website",
  enquiry_type: "Custom",
  client_name: "Michael Cost",
  product: "Daphne Rectangular Mirror",
  custom_specs: [{ kind: "Dimensions", detail: "800 x 600" }]
});
assert.strictEqual(same.events.filter((ev) => ev.kind === "edited").length, 0);
const renamed = db.upsertEnquiry({
  enquiry_no: first.enquiry_no,
  date_enquired: "30/11/2025",
  enquiry_source: "Website",
  enquiry_type: "Custom",
  client_name: "Michael Cost edited",
  product: "Daphne Rectangular Mirror",
  custom_specs: [{ kind: "Dimensions", detail: "800 x 600" }]
});
assert.ok(renamed.events.some((ev) => ev.kind === "edited"));

const noSpec = db.upsertEnquiry({
  date_enquired: "01/09/2026",
  enquiry_type: "Custom",
  client_name: "No Spec"
});
assert.strictEqual(noSpec.status, "New");
assert.deepStrictEqual(noSpec.custom_specs, []);

const leftoverOther = db.upsertEnquiry({
  date_enquired: "01/09/2026",
  enquiry_type: "Custom",
  client_name: "Other missing",
  custom_specs: [{ kind: "Other", other: "", detail: "something" }]
});
assert.strictEqual(leftoverOther.status, "New");
assert.strictEqual(leftoverOther.custom_specs[0].kind, "Other");
assert.strictEqual(leftoverOther.custom_specs[0].other, "");

const othered = db.upsertEnquiry({
  date_enquired: "01/09/2026",
  enquiry_type: "Custom",
  client_name: "Handle job",
  product: "Air Chair",
  custom_specs: [{ kind: "Other", other: "Handle", detail: "brass D-pull" }]
});
assert.strictEqual(othered.custom_specs[0].kind, "Handle");
assert.ok(db.listEnquiryDropdowns().custom_spec.indexOf("Handle") >= 0);
assert.ok(db.listEnquiryDropdowns().custom_spec.indexOf("Dimensions") >= 0);

const shortDesign = db.upsertEnquiry({
  date_enquired: "01/09/2026",
  enquiry_type: "New Design",
  client_name: "Sketch",
  design_description: "a table"
});
assert.strictEqual(shortDesign.status, "New");
assert.strictEqual(shortDesign.design_description, "a table");

const design = db.upsertEnquiry({
  date_enquired: "02/09/2026",
  enquiry_type: "New Design",
  client_name: "New brief",
  design_description: "Steel dining table with a live-edge oak top and arched black base."
});
assert.ok(design.request.indexOf("live-edge") >= 0);
assert.strictEqual(design.custom_specs.length, 0);

const pricedTooSoon = db.upsertEnquiry({
  enquiry_no: "#1996",
  date_enquired: "30/11/2025",
  client_name: "Michael Cost",
  status: "New",
  products: [{ product: "Daphne Rectangular Mirror", category: "Mirror", value_excl_vat: "999" }],
  delivery_excl_vat: "50"
});
assert.strictEqual(pricedTooSoon.products[0].value_excl_vat, "");
assert.strictEqual(pricedTooSoon.delivery_excl_vat, "");

const second = db.upsertEnquiry({
  date_enquired: "05/01/2026",
  client_name: "Sas Promotions",
  status: "Followed Up"
});
assert.strictEqual(second.enquiry_no, "#2002");
assert.strictEqual(second.month_enquired, "Jan");
assert.strictEqual(second.status, "New");
assert.strictEqual(db.nextEnquiryNo(), "#2003");

const listed = db.listEnquiries();
assert.strictEqual(listed[0].enquiry_no, "#2002");
assert.ok(listed.some((r) => r.enquiry_no === "#1996"));

const drops = db.listEnquiryDropdowns();
assert.ok(drops.enquiry_source.indexOf("Whatsapp") >= 0);
assert.ok(drops.enquiry_type.indexOf("Catologue") >= 0);
assert.ok(drops.product.indexOf("Violet Sideboard 3-Door") >= 0);
assert.ok(drops.category.indexOf("Gate") >= 0);
assert.ok(drops.status.indexOf("Waiting on clients specifictions") >= 0);
assert.ok(drops.status.indexOf("Costing") >= 0);
assert.ok(drops.custom_spec.indexOf("Handle") >= 0);
assert.ok(drops.custom_spec.indexOf("Dimensions") >= 0);
assert.ok(drops.custom_spec.indexOf("Colour") >= 0);

const jumped = db.upsertEnquiry({
  enquiry_no: "#1996",
  date_enquired: "30/11/2025",
  client_name: "Michael Cost",
  status: "Quoted",
  products: [{ product: "Daphne Rectangular Mirror", value_excl_vat: "1000" }],
  delivery_excl_vat: "50",
  quote_pdf_base64: "data:application/pdf;base64,eA==",
  quote_pdf_confirmed: true
});
assert.strictEqual(jumped.status, "New");
assert.strictEqual(jumped.products[0].value_excl_vat, "");
assert.strictEqual(jumped.has_quote_pdf, false);

assert.strictEqual(db.nextQuoteNo(), "SOQ1");
assert.deepStrictEqual(db.recentQuoteNos(3), []);
assert.strictEqual(db.normalizeQuoteNo("soq 2361"), "SOQ2361");
assert.throws(() => db.normalizeQuoteNo("ABC"), /look like/i);

const qseq = db.upsertEnquiry({
  date_enquired: "03/09/2026",
  client_name: "Quote sequence",
  enquiry_type: "Catologue"
});
const qraw = db.getEnquiryRaw(qseq.enquiry_no);
qraw.quote_no = "SOQ2360";
db.saveEnquiryRecord(qraw);
assert.strictEqual(db.nextQuoteNo(), "SOQ2361");
assert.deepStrictEqual(db.recentQuoteNos(5), ["SOQ2360"]);
assert.throws(() => db.requireUniqueQuoteNo("SOQ2360"), /already used/i);
assert.strictEqual(db.requireUniqueQuoteNo("SOQ2360", qseq.enquiry_no), "SOQ2360");
assert.strictEqual(db.parseCorrespondenceName("Re_ Order #S260023 - GERNOT CANTO.msg").order_no, "S260023");
assert.strictEqual(db.parseCorrespondenceName("Re: Order #S260025 - LAUGE SORENSEN").customer, "LAUGE SORENSEN");
assert.strictEqual(db.sanitizeOutlookOpenUrl("javascript:alert(1)"), "");
assert.strictEqual(db.sanitizeOutlookOpenUrl("file:///P:/ORDERS/mail.msg"), "");
assert.ok(db.outlookDesktopUrl({ title: "Checklist Manifesto" }).indexOf("ms-outlook://search") === 0);
const extracted = db.extractOutlookFromBuffer(Buffer.from("Message-ID: <a@b.com>\r\nSubject: Re: Order #S260025 - LAUGE SORENSEN\r\nFrom: Office\r\n"), "mail.msg");
assert.strictEqual(extracted.order_no, "S260025");
assert.strictEqual(extracted.internet_message_id, "<a@b.com>");
assert.ok(!db.extractOutlookFromDataUrl("data:application/octet-stream;base64,QQ==", "note.msg"));
assert.deepStrictEqual(
  db.parseOutlookLinks("see https://outlook.office.com/owa/?ItemID=ABC123&exvsurl=1"),
  ["https://outlook.office.com/owa/?ItemID=ABC123&exvsurl=1"]
);
assert.ok(db.sanitizeSavedLink("https://chabies-security-production.up.railway.app/api/office/enquiries/%231996/files/quote").indexOf("railway.app") !== -1);
assert.strictEqual(db.sanitizeSavedLink("javascript:alert(1)"), "");
assert.strictEqual(db.sanitizeSavedLink("https://share.example/a//b//c#frag"), "https://share.example/a//b//c#frag");
assert.strictEqual(db.sanitizeSavedLink("\\\\server\\share\\job//folder"), "\\\\server\\share\\job//folder");
assert.strictEqual(db.sanitizeSavedLink("file://server/share/a//b"), "file://server/share/a//b");
assert.strictEqual(db.sanitizeSavedLink("http://intranet/a//b"), "http://intranet/a//b");
assert.ok(db.parseOutlookLinks("https://chabies-security-production.up.railway.app/api/office/enquiries/%231996/files/quote").length === 1);
const slashMail = db.mailsFromPastedLinks("https://files.example/a//b//c\n\\\\nas\\jobs\\Q1//folder");
assert.strictEqual(slashMail.length, 2);
assert.strictEqual(slashMail[0].outlook_url, "https://files.example/a//b//c");
assert.strictEqual(slashMail[1].outlook_url, "\\\\nas\\jobs\\Q1//folder");
assert.ok(!/^ms-outlook:/i.test(slashMail[0].outlook_url));

const linked = db.getEnquiryRaw("#1996");
linked.correspondence = {
  saved_at: "2026-04-20T13:21:00.000Z",
  saved_by: "Admin",
  mails: [{
    title: "Re: Order #S260025 - LAUGE SORENSEN",
    rest_id: "AAMkKeep",
    from: "Office"
  }]
};
db.saveEnquiryRecord(linked);
const kept = db.upsertEnquiry({
  enquiry_no: "#1996",
  date_enquired: "30/11/2025",
  client_name: "Michael Cost",
  comment: "keep correspondence"
});
assert.strictEqual(kept.correspondence.mails.length, 1);
assert.ok(kept.correspondence.mails[0].outlook_url.indexOf("ms-outlook://") === 0);
assert.strictEqual(kept.correspondence.mails[0].order_no, "S260025");

const rawQuote = db.getEnquiryRaw("#1996");
rawQuote.quote_no = "";
rawQuote.date_quoted = "";
rawQuote.quotes = [{ n: 1, quote_no: "SOQ9", date_quoted: "", uploaded_at: "2026-09-02T10:00:00.000Z" }];
const shownQuote = db.getEnquiry("#1996");
assert.strictEqual(shownQuote.quote_no, "SOQ9");
assert.strictEqual(shownQuote.date_quoted, "02/09/2026");

const saved = JSON.parse(fs.readFileSync(db.dbPath, "utf8"));
assert.strictEqual(saved.enquiries.length, 8);

assert.throws(
  () => db.upsertEnquiry({
    enquiry_no: "#1996",
    date_enquired: "01/09/2026",
    client_name: "Taken number"
  }, { createOnly: true }),
  /already/
);
const customNo = db.upsertEnquiry({
  enquiry_no: "2200",
  date_enquired: "01/09/2026",
  client_name: "Custom No"
}, { createOnly: true });
assert.strictEqual(customNo.enquiry_no, "#2200");
assert.strictEqual(db.nextEnquiryNo(), "#2201");

console.log("enquiries.test.js ok");
