const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-desk-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";

const db = require("./db");
const desk = require("./enquiry-desk");

const origList = db.listEnquiries;
const origGet = db.getEnquiry;
db.listEnquiries = () => [{
  enquiry_no: "#2604",
  client_name: "Blaire Bedroom Client",
  client_number: "0820000000",
  client_email: "blaire@example.com",
  product: "Blaire Dressing Table",
  quote_no: "SOQ2604",
  status: "Quoted",
  province: "KwaZulu-Natal",
  source: "Instagram",
  enquiry_type: "Catologue"
}];
db.getEnquiry = (no) => (no === "#2604" ? db.listEnquiries()[0] : null);

try {
  const replies = desk.loadReplies();
  assert.ok(replies.length >= 25);
  assert.ok(replies.some((r) => r.id === "costing-in-progress"));
  assert.ok(replies.some((r) => r.id === "standard-quotes" && /quotation/.test(r.subject)));
  assert.ok(replies.some((r) => r.id === "inquiry-with-wood" && /rubiomonocoat/.test(r.body)));
  assert.ok(replies.some((r) => r.id === "request-float-link" && /float\.co\.za/.test(r.body)));
  assert.ok(replies.some((r) => r.id === "delivery-gauteng" && /Silverton/.test(r.body)));
  assert.ok(replies.every((r) => r.subject && r.body && r.title));
  const filled = desk.fillReply(replies.find((r) => r.id === "standard-quotes"), "#2604");
  assert.ok(filled.subject.indexOf("SOQ2604") !== -1);
  assert.ok(filled.body.indexOf("Blaire Bedroom Client") !== -1);
  assert.ok(desk.fillText("Good day [Client's Name],", { client_name: "Pat" }).indexOf("Pat") !== -1);
  assert.ok(desk.fillText("Hi {{client_name}}", null).indexOf("there") !== -1);
  assert.ok(desk.fillText("{{client_email}} {{client_number}} {{province}} {{status}}", {
    client_email: "pat@studio",
    client_number: "0821",
    province: "Gauteng",
    status: "Quoted"
  }).indexOf("pat@studio") !== -1);
  const filledMore = desk.fillReply({
    subject: "{{enquiry_no}} {{quote_no}}",
    body: "{{client_email}} / {{client_number}} / {{province}}"
  }, "#2604");
  assert.ok(filledMore.body.indexOf("blaire@example.com") !== -1);
  assert.ok(filledMore.body.indexOf("0820000000") !== -1);
  assert.ok(filledMore.body.indexOf("KwaZulu-Natal") !== -1);

  const added = desk.upsertReply({
    title: "Custom chase",
    topic: "Follow-up",
    subject: "Studio Delta {{enquiry_no}}",
    body: "Hello {{client_name}}"
  });
  assert.ok(/^er-/.test(added.reply.id));
  assert.strictEqual(desk.loadReplies().filter((r) => r.id === added.reply.id).length, 1);
  desk.upsertReply({ id: added.reply.id, title: "Custom chase 2", topic: "Follow-up", subject: "S", body: "B" });
  assert.strictEqual(desk.loadReplies().find((r) => r.id === added.reply.id).title, "Custom chase 2");
  desk.deleteReply(added.reply.id);
  assert.ok(!desk.loadReplies().some((r) => r.id === added.reply.id));

  const restored = desk.restoreReplies();
  assert.ok(restored.some((r) => r.id === "costing-in-progress"));
  assert.ok(restored.some((r) => r.topic === "December" && r.id === "december-mds"));

  assert.throws(() => desk.upsertReply({ title: "x", subject: "", body: "b" }), /Subject/);

  const first = desk.upsertBooking({
    enquiry_no: "#2604",
    date: "2026-09-10",
    time: "10:00",
    duration_min: 45,
    notes: "Bring sizes"
  }, "Admin");
  assert.strictEqual(first.booking.client_name, "Blaire Bedroom Client");
  assert.strictEqual(first.booking.status, "Booked");
  assert.strictEqual(first.booking.time, "10:00");
  assert.ok(first.booking.starts_at.indexOf("2026-09-10T10:00:00+02:00") !== -1);
  assert.strictEqual(first.overlap, null);

  const clash = desk.upsertBooking({
    client_name: "Walk-in",
    date: "2026-09-10",
    time: "10:15",
    duration_min: 45
  }, "Pat");
  assert.ok(clash.overlap && clash.overlap.id === first.booking.id);

  const later = desk.upsertBooking({
    client_name: "Afternoon",
    date: "2026-09-10",
    time: "11:00",
    duration_min: 30
  }, "Pat");
  assert.strictEqual(later.overlap, null);

  assert.throws(() => desk.upsertBooking({ client_name: "X", date: "10/09/2026", time: "09:00" }), /date/);
  assert.throws(() => desk.upsertBooking({ date: "2026-09-11", time: "09:00" }), /Client name/);

  desk.deleteBooking(first.booking.id);
  assert.ok(!desk.loadBookings().some((b) => b.id === first.booking.id));
  const opt = desk.enquiryOptions().find((r) => r.enquiry_no === "#2604");
  assert.ok(opt);
  assert.strictEqual(opt.client_email, "blaire@example.com");
  assert.strictEqual(opt.province, "KwaZulu-Natal");
  assert.strictEqual(opt.product, "Blaire Dressing Table");
} finally {
  db.listEnquiries = origList;
  db.getEnquiry = origGet;
}

console.log("enquiry-desk.test.js ok");
