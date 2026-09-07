const fs = require("fs");
const nodePath = require("path");
const { dataDir } = require("./workbook-store");
const db = require("./db");

const TOPICS = ["Enquiry", "Quote", "Follow-up", "Showroom", "Order", "Close"];
const BOOKING_STATUSES = ["Booked", "Done", "No-show", "Cancelled"];
const PLACEHOLDERS = [
  { key: "{{client_name}}", label: "Client name" },
  { key: "{{enquiry_no}}", label: "Enquiry number" },
  { key: "{{quote_no}}", label: "Quote number" },
  { key: "{{product}}", label: "Product" }
];

const DEFAULT_REPLIES = [
  {
    id: "thank-you",
    topic: "Enquiry",
    title: "Thank you for the enquiry",
    subject: "Studio Delta — we have your enquiry {{enquiry_no}}",
    body:
      "Good day {{client_name}}\n\n" +
      "Thank you for getting in touch with Studio Delta. We have logged enquiry {{enquiry_no}} for {{product}}.\n\n" +
      "Someone from the office will come back to you as soon as we have checked the request. If we still need sizes, finish, or contact details, we will ask in a separate mail.\n\n" +
      "Kind regards\nStudio Delta"
  },
  {
    id: "waiting-details",
    topic: "Enquiry",
    title: "Waiting on personal details",
    subject: "Studio Delta {{enquiry_no}} — details we still need",
    body:
      "Good day {{client_name}}\n\n" +
      "We have enquiry {{enquiry_no}} on the system, but we cannot start costing until we have your name, email or cell number, and province.\n\n" +
      "Please reply with those details so we can continue.\n\n" +
      "Kind regards\nStudio Delta"
  },
  {
    id: "waiting-specs",
    topic: "Enquiry",
    title: "Waiting on specifications",
    subject: "Studio Delta {{enquiry_no}} — specifications for {{product}}",
    body:
      "Good day {{client_name}}\n\n" +
      "Thank you — enquiry {{enquiry_no}} is with us. To cost {{product}} we still need the missing specifications (sizes, finish / colour, or a clear description of the change).\n\n" +
      "Please reply with those details, or say if you would rather visit the showroom to go through them.\n\n" +
      "Kind regards\nStudio Delta"
  },
  {
    id: "quote-sent",
    topic: "Quote",
    title: "Quote attached",
    subject: "Studio Delta quotation {{quote_no}} — {{enquiry_no}}",
    body:
      "Good day {{client_name}}\n\n" +
      "Please find our quotation {{quote_no}} for enquiry {{enquiry_no}} ({{product}}).\n\n" +
      "Figures on the quote exclude VAT unless the line says otherwise. Delivery is shown separately.\n\n" +
      "We will follow up if we have not heard from you. You are welcome to reply with questions, or to book a showroom visit.\n\n" +
      "Kind regards\nStudio Delta"
  },
  {
    id: "follow-up",
    topic: "Follow-up",
    title: "Follow-up on the quote",
    subject: "Studio Delta {{quote_no}} — following up on {{enquiry_no}}",
    body:
      "Good day {{client_name}}\n\n" +
      "I am following up on quotation {{quote_no}} for enquiry {{enquiry_no}} ({{product}}).\n\n" +
      "Have you had a chance to go through it? If you would like a change, another quote, or a showroom visit, reply to this mail and we will arrange it.\n\n" +
      "Kind regards\nStudio Delta"
  },
  {
    id: "showroom-invite",
    topic: "Showroom",
    title: "Invite to the showroom",
    subject: "Studio Delta showroom — {{enquiry_no}}",
    body:
      "Good day {{client_name}}\n\n" +
      "You are welcome to visit the Studio Delta showroom to look at {{product}} for enquiry {{enquiry_no}}.\n\n" +
      "Please reply with a day and time that suits you (weekday mornings work best) and we will book you in. Bring sizes and any pictures if you have them.\n\n" +
      "Kind regards\nStudio Delta"
  },
  {
    id: "showroom-confirm",
    topic: "Showroom",
    title: "Showroom booking confirmed",
    subject: "Studio Delta showroom booking — {{enquiry_no}}",
    body:
      "Good day {{client_name}}\n\n" +
      "Your showroom visit for enquiry {{enquiry_no}} ({{product}}) is booked. We look forward to seeing you.\n\n" +
      "If you need to move the time, reply to this mail and we will change the diary.\n\n" +
      "Kind regards\nStudio Delta"
  },
  {
    id: "pop-thanks",
    topic: "Order",
    title: "Proof of payment received",
    subject: "Studio Delta {{enquiry_no}} — thank you, we have your payment",
    body:
      "Good day {{client_name}}\n\n" +
      "Thank you. We have the proof of payment for enquiry {{enquiry_no}} ({{product}}). The job now moves onto Orders and into production.\n\n" +
      "We will be in touch if we need anything further.\n\n" +
      "Kind regards\nStudio Delta"
  },
  {
    id: "not-in-scope",
    topic: "Close",
    title: "Not within scope",
    subject: "Studio Delta {{enquiry_no}} — not a product we make",
    body:
      "Good day {{client_name}}\n\n" +
      "Thank you for enquiry {{enquiry_no}}. After reviewing {{product}}, this is not a job we can take on — it falls outside what Studio Delta manufactures.\n\n" +
      "We are sorry we cannot help on this one. You are welcome to send a different request.\n\n" +
      "Kind regards\nStudio Delta"
  },
  {
    id: "not-interested",
    topic: "Close",
    title: "Client not going ahead",
    subject: "Studio Delta {{enquiry_no}} — closing the enquiry",
    body:
      "Good day {{client_name}}\n\n" +
      "Thank you for letting us know you will not go ahead with enquiry {{enquiry_no}} ({{product}}).\n\n" +
      "We have closed it on our side. If you come back to it later, reply to this mail and we will reopen it.\n\n" +
      "Kind regards\nStudio Delta"
  }
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
  if (TOPICS.indexOf(topic) < 0) topic = "Enquiry";
  if (!id) throw new Error("A reply id is required");
  if (!title) throw new Error("Give the reply a name");
  if (!subject) throw new Error("Subject is required");
  if (!body) throw new Error("Body is required");
  return { id, topic, title, subject, body };
}

function loadReplies() {
  const raw = readJson(repliesPath(), null);
  if (!raw || !Array.isArray(raw.replies) || !raw.replies.length) {
    const replies = cloneDefaults();
    writeJson(repliesPath(), { replies });
    return replies;
  }
  return raw.replies.map((row, i) => normalizeReply(row, "er-" + (i + 1)));
}

function saveReplies(list) {
  const replies = (Array.isArray(list) ? list : []).map((row, i) => normalizeReply(row, "er-" + (i + 1)));
  writeJson(repliesPath(), { replies });
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
  return String(text || "").replace(/\{\{(client_name|enquiry_no|quote_no|product)\}\}/g, (m) => map[m] || m);
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
