"use strict";

const db = require("./db");
const jobCard = require("./job-card");
const orderDocs = require("./order-docs");
const drawingPrintAlerts = require("./drawing-print-alerts");

function requireFile(body) {
  const raw = (body && (body.file_base64 || body.fileBase64)) || "";
  if (!String(raw).trim()) throw new Error("Upload the drawing, check the preview, then confirm it");
  if (!(body.file_confirmed || body.fileConfirmed)) {
    throw new Error("Tick that this is the correct drawing before saving");
  }
  return raw;
}

function isPdfOrImage(name) {
  const n = String(name || "").toLowerCase();
  return /\.(pdf|png|jpe?g|webp|gif)$/.test(n);
}

// Tasks for this order number only (or old enquiry-wide tasks with no order number).
function taskIsForOrder(t, orderNumber) {
  const own = db.formatOrderId(t && t.order_number);
  return !own || own === orderNumber;
}

function closeOpenKind(row, kind, actor, orderNumber) {
  (row.tasks || []).forEach((t) => {
    if (t.kind === kind && t.status === "open" && taskIsForOrder(t, orderNumber)) {
      t.status = "done";
      t.completed_at = db.nowIso();
      t.completed_by = actor;
      t.note = "Drawing updated on Orders";
    }
  });
}

function cancelOpenKind(row, kind, orderNumber) {
  (row.tasks || []).forEach((t) => {
    if (t.kind === kind && t.status === "open" && taskIsForOrder(t, orderNumber)) t.status = "cancelled";
  });
}

function drawingUrl(enquiryNo, uploadedAt, orderNumber) {
  const kind = orderNumber ? db.drawingFileKindForOrder(orderNumber) : "drawing";
  const base = "/api/office/enquiries/" + encodeURIComponent(enquiryNo) + "/files/" + encodeURIComponent(kind);
  const v = String(uploadedAt || "").trim();
  return v ? (base + "?v=" + encodeURIComponent(v)) : base;
}

// Siya reprints only the order whose drawing changed.
function notifyJobCards(enquiry, originOrder, actor, uploadedAt) {
  const list = originOrder ? [originOrder] : db.ordersLinkedToEnquiry(enquiry);
  const seen = new Set();
  const alerts = [];
  list.forEach((order) => {
    const num = db.formatOrderId(order && order.order_number);
    if (!num || seen.has(num)) return;
    seen.add(num);
    if (!jobCard.getJobCard(num)) return;
    alerts.push(drawingPrintAlerts.upsertPending({
      order_number: num,
      client_name: order.client_name || originOrder.client_name || enquiry.client_name || "",
      product: order.product || originOrder.product || enquiry.product || "",
      drawing_url: drawingUrl(enquiry.enquiry_no, uploadedAt, db.drawingOrderEntry(enquiry.drawing, num) ? num : ""),
      uploaded_at: uploadedAt,
      uploaded_by: actor
    }));
  });
  return alerts.filter(Boolean);
}

function updateFromOrder(orderNumber, actor, body) {
  const who = String(actor || "").trim();
  if (!who) throw new Error("Not signed in");
  const want = db.formatOrderId(orderNumber);
  if (!want) throw new Error("Order number is required");
  const order = db.listOrders().find((o) => db.formatOrderId(o.order_number) === want);
  if (!order) throw new Error("Order not found");
  const enquiryHit = orderDocs.findEnquiry(order, orderDocs.enquiryIndex());
  if (!enquiryHit) throw new Error("This order has no enquiry to attach a drawing to");
  const raw = db.getEnquiryRaw(enquiryHit.enquiry_no);
  if (!raw) throw new Error("This order has no enquiry to attach a drawing to");
  const filename = String((body && (body.file_name || body.fileName)) || "drawing.pdf").trim() || "drawing.pdf";
  if (!isPdfOrImage(filename)) throw new Error("Drawing must be a PDF or image");
  const fileRaw = requireFile(body);
  if (!raw.drawing || typeof raw.drawing !== "object") raw.drawing = { required: true, file: null };
  raw.drawing.required = true;
  if (!raw.drawing.orders || typeof raw.drawing.orders !== "object") raw.drawing.orders = {};
  const uploadedAt = db.nowIso();
  raw.drawing.orders[want] = {
    file: db.saveEnquiryAttachment(raw.enquiry_no, db.drawingFileKindForOrder(want), fileRaw, filename),
    uploaded_at: uploadedAt,
    uploaded_by: who,
    approval: {
      status: "approved",
      decided_at: uploadedAt,
      decided_by: who,
      note: "Updated on the Orders sheet"
    }
  };
  if (!Array.isArray(raw.tasks)) raw.tasks = [];
  closeOpenKind(raw, "drawing", who, want);
  cancelOpenKind(raw, "drawing_approval", want);
  raw.ready_for_orders = true;
  db.applyDrawingShopStatus(raw);
  db.appendEnquiryEvent(raw, {
    kind: "order_drawing_update",
    actor: who,
    status: raw.status || "",
    label: "Drawing updated on Orders — " + want,
    note: filename
  });
  raw.updated_at = db.nowIso();
  db.saveEnquiryRecord(raw);
  const alerts = notifyJobCards(raw, order, who, uploadedAt);
  const live = db.listOrders().find((o) => db.formatOrderId(o.order_number) === want) || order;
  return {
    order: live,
    enquiry_no: raw.enquiry_no,
    drawing_url: drawingUrl(raw.enquiry_no, uploadedAt, want),
    notified_siya: alerts.length > 0,
    print_alerts: alerts
  };
}

module.exports = {
  updateFromOrder,
  notifyJobCards
};
