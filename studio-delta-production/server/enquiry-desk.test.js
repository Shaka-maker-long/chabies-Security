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
  status: "Quoted"
}];
db.getEnquiry = (no) => (no === "#2604" ? db.listEnquiries()[0] : null);

try {
  const replies = desk.loadReplies();
  assert.ok(replies.length >= 8);
  assert.ok(replies.some((r) => r.id === "quote-sent" && /quotation/.test(r.subject)));
  assert.ok(replies.every((r) => r.subject && r.body && r.title));
  const filled = desk.fillReply(replies.find((r) => r.id === "quote-sent"), "#2604");
  assert.ok(filled.subject.indexOf("SOQ2604") !== -1);
  assert.ok(filled.body.indexOf("Blaire Bedroom Client") !== -1);
  assert.ok(filled.body.indexOf("Blaire Dressing Table") !== -1);
  assert.ok(desk.fillText("Hi {{client_name}}", null).indexOf("there") !== -1);

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
  assert.ok(restored.some((r) => r.id === "thank-you"));

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
  assert.ok(desk.enquiryOptions().some((r) => r.enquiry_no === "#2604"));
} finally {
  db.listEnquiries = origList;
  db.getEnquiry = origGet;
}

console.log("enquiry-desk.test.js ok");
