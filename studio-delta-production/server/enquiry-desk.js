const fs = require("fs");
const nodePath = require("path");
const { dataDir } = require("./workbook-store");
const db = require("./db");
const { TOPICS, REPLIES_VERSION, DEFAULT_REPLIES } = require("./enquiry-replies-default");

const BOOKING_STATUSES = ["Booked", "Done", "No-show", "Cancelled"];
const PLACEHOLDERS = [
  { key: "{{client_name}}", label: "Client name" },
  { key: "{{enquiry_no}}", label: "Enquiry number" },
  { key: "{{quote_no}}", label: "Quote number" },
  { key: "{{product}}", label: "Product" }
];

function repliesPath() {
  return nodePath.join(dataDir(), "email-replies.json");
}

function bookingsPath() {
  return nodePath.join(dataDir(), "showroom-bookings.json");
}

function readJson(file, fallback) {
  try {
    return JSON.parse(fs.readFileSync(file, "utf8"));
  } catch (e) {
    if (e && e.code !== "ENOENT") throw e;
    return fallback;
  }
}

function writeJson(file, data) {
  fs.mkdirSync(nodePath.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(data, null, 2));
}

function cloneDefaults() {
  return DEFAULT_REPLIES.map((row) => Object.assign({}, row));
}

function normalizeReply(row, fallbackId) {
  const id = String((row && row.id) || fallbackId || "").trim();
  const title = String((row && row.title) || "").trim();
  const subject = String((row && row.subject) || "").trim();
  const body = String((row && row.body) || "").replace(/\r\n/g, "\n").trim();
  let topic = String((row && row.topic) || "Enquiry").trim();
  if (TOPICS.indexOf(topic) < 0) topic = "Other";
  if (!id) throw new Error("A reply id is required");
  if (!title) throw new Error("Give the reply a name");
  if (!subject) throw new Error("Subject is required");
  if (!body) throw new Error("Body is required");
  return { id, topic, title, subject, body };
}

function loadReplies() {
  const raw = readJson(repliesPath(), null);
  if (!raw || raw.version !== REPLIES_VERSION || !Array.isArray(raw.replies) || !raw.replies.length) {
    const replies = cloneDefaults();
    writeJson(repliesPath(), { version: REPLIES_VERSION, replies });
    return replies;
  }
  return raw.replies.map((row, i) => normalizeReply(row, "er-" + (i + 1)));
}

function saveReplies(list) {
  const replies = (Array.isArray(list) ? list : []).map((row, i) => normalizeReply(row, "er-" + (i + 1)));
  writeJson(repliesPath(), { version: REPLIES_VERSION, replies });
  return replies;
}

function upsertReply(body) {
  const replies = loadReplies();
  const incoming = Object.assign({}, body || {});
  if (!incoming.id) incoming.id = "er-" + Date.now();
  const row = normalizeReply(incoming);
  const idx = replies.findIndex((r) => r.id === row.id);
  if (idx >= 0) replies[idx] = row;
  else replies.push(row);
  return { reply: row, replies: saveReplies(replies) };
}

function deleteReply(id) {
  const key = String(id || "").trim();
  const replies = loadReplies().filter((r) => r.id !== key);
  return saveReplies(replies);
}

function restoreReplies() {
  return saveReplies(cloneDefaults());
}

function productLine(row) {
  if (!row) return "";
  if (Array.isArray(row.products) && row.products.length) {
    return row.products.map((p) => String((p && p.product) || "").trim()).filter(Boolean).join(", ");
  }
  return String(row.product || row.design_description || row.request || "").trim();
}

function fillText(text, enquiry) {
  const row = enquiry || {};
  const map = {
    "{{client_name}}": String(row.client_name || "").trim() || "there",
    "{{enquiry_no}}": String(row.enquiry_no || "").trim() || "the enquiry",
    "{{quote_no}}": String(row.quote_no || "").trim() || "the quotation",
    "{{product}}": productLine(row) || "your request"
  };
  return String(text || "")
    .replace(/\[Client['’]s Name\]/gi, map["{{client_name}}"])
    .replace(/\{\{(client_name|enquiry_no|quote_no|product)\}\}/g, (m) => map[m] || m);
}

function fillReply(reply, enquiryNo) {
  const enquiry = enquiryNo ? db.getEnquiry(enquiryNo) : null;
  return {
    subject: fillText(reply && reply.subject, enquiry),
    body: fillText(reply && reply.body, enquiry)
  };
}

function nextBookingId(list) {
  let max = 0;
  (list || []).forEach((row) => {
    const m = String(row.id || "").match(/^sb-(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return "sb-" + (max + 1);
}

function parseHm(t) {
  const m = String(t || "").trim().match(/^(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const h = Number(m[1]);
  const min = Number(m[2]);
  if (h < 0 || h > 23 || min < 0 || min > 59) return null;
  return h * 60 + min;
}

function normalizeBooking(row, fallbackId) {
  const id = String((row && row.id) || fallbackId || "").trim();
  const date = String((row && row.date) || "").trim();
  const time = String((row && row.time) || "").trim();
  if (!id) throw new Error("A booking id is required");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error("Pick a showroom date");
  if (!parseHm(time)) throw new Error("Pick a start time");
  let status = String((row && row.status) || "Booked").trim();
  if (BOOKING_STATUSES.indexOf(status) < 0) status = "Booked";
  let duration_min = Number(row && row.duration_min);
  if (!Number.isFinite(duration_min) || duration_min < 15) duration_min = 45;
  duration_min = Math.min(180, Math.round(duration_min));
  let enquiry_no = String((row && row.enquiry_no) || "").trim();
  let client_name = String((row && row.client_name) || "").trim();
  let client_number = String((row && row.client_number) || "").trim();
  let client_email = String((row && row.client_email) || "").trim();
  if (enquiry_no) {
    const enq = db.getEnquiry(enquiry_no);
    if (enq) {
      if (!client_name) client_name = String(enq.client_name || "").trim();
      if (!client_number) client_number = String(enq.client_number || "").trim();
      if (!client_email) client_email = String(enq.client_email || "").trim();
    }
  }
  if (!client_name) throw new Error("Client name is required");
  const minutes = parseHm(time);
  const starts_at = date + "T" + String(Math.floor(minutes / 60)).padStart(2, "0") + ":" + String(minutes % 60).padStart(2, "0") + ":00+02:00";
  return {
    id,
    enquiry_no,
    client_name,
    client_number,
    client_email,
    date,
    time: String(Math.floor(minutes / 60)).padStart(2, "0") + ":" + String(minutes % 60).padStart(2, "0"),
    duration_min,
    notes: String((row && row.notes) || "").trim(),
    status,
    booked_by: String((row && row.booked_by) || "").trim(),
    created_at: String((row && row.created_at) || new Date().toISOString()),
    updated_at: new Date().toISOString(),
    starts_at
  };
}

function loadBookings() {
  const raw = readJson(bookingsPath(), { bookings: [] });
  const list = Array.isArray(raw.bookings) ? raw.bookings : [];
  return list.map((row, i) => normalizeBooking(row, "sb-" + (i + 1)));
}

function saveBookings(list) {
  const bookings = (Array.isArray(list) ? list : []).map((row, i) => normalizeBooking(row, "sb-" + (i + 1)));
  bookings.sort((a, b) => String(a.starts_at).localeCompare(String(b.starts_at)) || a.client_name.localeCompare(b.client_name));
  writeJson(bookingsPath(), { bookings });
  return bookings;
}

function bookingOverlap(list, row) {
  if (row.status !== "Booked") return null;
  const start = parseHm(row.time);
  const end = start + row.duration_min;
  return (list || []).find((other) => {
    if (other.id === row.id || other.status !== "Booked" || other.date !== row.date) return false;
    const oStart = parseHm(other.time);
    const oEnd = oStart + other.duration_min;
    return start < oEnd && oStart < end;
  }) || null;
}

function upsertBooking(body, actor) {
  const bookings = loadBookings();
  const incoming = Object.assign({}, body || {});
  if (!incoming.id) incoming.id = nextBookingId(bookings);
  if (!incoming.booked_by) incoming.booked_by = String(actor || "").trim();
  if (!incoming.created_at) {
    const prev = bookings.find((b) => b.id === incoming.id);
    incoming.created_at = prev ? prev.created_at : new Date().toISOString();
  }
  const row = normalizeBooking(incoming);
  const clash = bookingOverlap(bookings, row);
  const idx = bookings.findIndex((b) => b.id === row.id);
  if (idx >= 0) bookings[idx] = row;
  else bookings.push(row);
  return { booking: row, bookings: saveBookings(bookings), overlap: clash ? { id: clash.id, time: clash.time, client_name: clash.client_name } : null };
}

function deleteBooking(id) {
  const key = String(id || "").trim();
  return saveBookings(loadBookings().filter((b) => b.id !== key));
}

function enquiryOptions() {
  return db.listEnquiries().map((row) => ({
    enquiry_no: row.enquiry_no,
    client_name: row.client_name || "",
    client_number: row.client_number || "",
    client_email: row.client_email || "",
    product: productLine(row),
    quote_no: row.quote_no || "",
    status: row.status || ""
  }));
}

module.exports = {
  TOPICS,
  REPLIES_VERSION,
  BOOKING_STATUSES,
  PLACEHOLDERS,
  DEFAULT_REPLIES,
  loadReplies,
  upsertReply,
  deleteReply,
  restoreReplies,
  fillReply,
  fillText,
  loadBookings,
  upsertBooking,
  deleteBooking,
  bookingOverlap,
  enquiryOptions
};
