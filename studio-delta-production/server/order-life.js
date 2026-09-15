"use strict";

const fs = require("fs");
const path = require("path");
const { dataDir, getBook } = require("./workbook-store");
const {
  listOrders,
  getEnquiry,
  ordersLinkedToEnquiry,
  formatOrderId,
  formatSastDateTime,
  formatRand,
  asDate
} = require("./db");
const { normalizeBaseOrderNumber } = require("./create-order-from-enquiry");
const jobCard = require("./job-card");
const paintShop = require("./powder-shop");
const glassPo = require("./glass-po");

const WOOD_HEADERS = [
  "ID", "Timestamp", "Order #", "Worker", "Component", "Wood type",
  "Thickness", "Height", "Width", "Quantity", "Status"
];

const ENQUIRY_STAGES = {
  created: "enquiry",
  capture: "enquiry",
  status_change: "enquiry",
  request_access: "enquiry",
  grant_access: "enquiry",
  deny_access: "enquiry",
  assign_costing: "enquiry",
  complete_cost_sheet: "enquiry",
  complete_supplier: "enquiry",
  complete_chase: "enquiry",
  complete_approval: "enquiry",
  add_correspondence: "enquiry",
  close: "enquiry",
  complete_quote: "quote",
  complete_quote_option: "quote",
  quote: "quote",
  complete_followup: "enquiry",
  follow_up: "enquiry",
  complete_order: "money",
  complete_drawing: "enquiry"
};

function sameOrder(left, right) {
  const a = formatOrderId(left);
  const b = formatOrderId(right);
  if (!a || !b) return false;
  if (a === b) return true;
  const baseA = normalizeBaseOrderNumber(a);
  const baseB = normalizeBaseOrderNumber(b);
  if (baseA && baseA === baseB && a === baseA) return true;
  return false;
}

function toIso(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.toISOString();
  if (typeof value === "number" && Number.isFinite(value)) {
    const d = asDate(value) || new Date(value);
    return d && !isNaN(d.getTime()) ? d.toISOString() : "";
  }
  const s = String(value || "").trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s + "T06:00:00.000Z";
  const parsed = asDate(s);
  if (parsed && !isNaN(parsed.getTime())) return parsed.toISOString();
  const n = Date.parse(s);
  return Number.isFinite(n) ? new Date(n).toISOString() : "";
}

function eventOf(partial) {
  const at = toIso(partial.at);
  const label = formatSastDateTime(at) || String(partial.at_label || "").trim();
  return {
    id: String(partial.id || ""),
    at,
    at_label: label,
    stage: String(partial.stage || "order"),
    title: String(partial.title || "").trim(),
    detail: String(partial.detail || "").trim(),
    actor: String(partial.actor || "").trim(),
    order_number: String(partial.order_number || "").trim(),
    kind: String(partial.kind || "")
  };
}

function findOrder(orderNumber) {
  const want = formatOrderId(orderNumber);
  if (!want) return null;
  return listOrders().find((o) => formatOrderId(o && o.order_number) === want) || null;
}

function enquiryEvents(enquiry, orderNumber) {
  const out = [];
  if (!enquiry) return out;
  const events = Array.isArray(enquiry.events) ? enquiry.events : [];
  events.forEach((ev) => {
    const stage = ENQUIRY_STAGES[String((ev && ev.kind) || "")] || "enquiry";
    const bits = [];
    if (ev && ev.status) bits.push(ev.status);
    if (ev && ev.note) bits.push(ev.note);
    out.push(eventOf({
      id: "enquiry-" + String((ev && ev.id) || out.length + 1),
      at: ev && ev.at,
      at_label: ev && ev.at_label,
      stage,
      title: (ev && (ev.label || ev.kind)) || "Enquiry",
      detail: bits.join(" · "),
      actor: ev && ev.actor,
      order_number: orderNumber || enquiry.order_number || "",
      kind: ev && ev.kind
    }));
  });
  const outcome = enquiry.client_outcome || {};
  if (outcome.decided_at && !events.some((ev) => ev && ev.kind === "complete_order")) {
    out.push(eventOf({
      id: "enquiry-pop",
      at: outcome.decided_at,
      stage: "money",
      title: "Proof of payment",
      detail: String(outcome.reason || outcome.status || "").trim(),
      order_number: orderNumber || enquiry.order_number || "",
      kind: "complete_order"
    }));
  }
  return out;
}

function orderOpenedEvents(order) {
  const out = [];
  if (!order) return out;
  const at = order.payment_date || order.month_of_sale || "";
  if (!toIso(at)) return out;
  out.push(eventOf({
    id: "order-opened-" + formatOrderId(order.order_number),
    at,
    stage: "order",
    title: "Order opened",
    detail: [order.status, order.product].filter(Boolean).join(" · "),
    order_number: order.order_number,
    kind: "order_opened"
  }));
  return out;
}

function jobCardEvents(order) {
  const out = [];
  const want = formatOrderId(order && order.order_number);
  if (!want) return out;
  try {
    (jobCard.listGeneratedJobCards() || []).forEach((card) => {
      if (formatOrderId(card && card.order_number) !== want) return;
      const at = card.created_at || card.created_date;
      if (!toIso(at)) return;
      out.push(eventOf({
        id: "job-card-" + want,
        at,
        stage: "shop",
        title: "Job card generated",
        detail: card.product || order.product || "",
        order_number: order.order_number,
        kind: "job_card"
      }));
    });
  } catch (e) {}
  return out;
}

function shopStage(process) {
  const s = String(process || "").toLowerCase();
  if (/pre-powder|final qc|quality control|\bqc\b/.test(s)) return "qc";
  if (/paint prep|painting|powder coating|paint shop/.test(s)) return "paint";
  return "shop";
}

function parseMeta(raw) {
  if (!raw) return {};
  if (typeof raw === "object") return raw;
  try {
    const parsed = JSON.parse(String(raw));
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch (e) {
    return {};
  }
}

function productionLogEvents(order) {
  const out = [];
  const want = formatOrderId(order && order.order_number);
  if (!want) return out;
  try {
    const sheet = getBook().getSheetByName("Production_Log");
    if (!sheet || sheet.getLastRow() < 2) return out;
    const lastCol = Math.max(sheet.getLastColumn(), 13);
    const grid = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
    for (let i = 1; i < grid.length; i++) {
      const row = grid[i] || [];
      const orderId = formatOrderId(row[1]);
      if (!orderId || orderId !== want) continue;
      const process = String(row[3] || "").trim() || "Shop work";
      const worker = String(row[2] || "").trim();
      const status = String(row[4] || "").trim();
      const start = toIso(row[5]);
      const end = toIso(row[6]);
      const meta = parseMeta(row[12]);
      const logId = String(row[0] || ("r" + i));
      const stage = shopStage(process);
      if (start) {
        out.push(eventOf({
          id: "log-" + logId + "-start",
          at: start,
          stage,
          title: "Started " + process,
          detail: status && !/^done|complete/i.test(status) ? status : "",
          actor: worker,
          order_number: order.order_number,
          kind: "shop_start"
        }));
      }
      (Array.isArray(meta.pauses) ? meta.pauses : []).forEach((pause, p) => {
        if (!pause) return;
        if (pause.start) {
          out.push(eventOf({
            id: "log-" + logId + "-pause-" + p,
            at: pause.start,
            stage,
            title: "Paused " + process,
            detail: String(pause.reason || "").trim(),
            actor: worker,
            order_number: order.order_number,
            kind: "shop_pause"
          }));
        }
        if (pause.end) {
          out.push(eventOf({
            id: "log-" + logId + "-resume-" + p,
            at: pause.end,
            stage,
            title: "Resumed " + process,
            actor: worker,
            order_number: order.order_number,
            kind: "shop_resume"
          }));
        }
      });
      if (end) {
        out.push(eventOf({
          id: "log-" + logId + "-end",
          at: end,
          stage,
          title: "Finished " + process,
          actor: worker,
          order_number: order.order_number,
          kind: "shop_end"
        }));
      }
    }
  } catch (e) {}
  return out;
}

function glassEvents(order) {
  const out = [];
  const want = order && order.order_number;
  if (!want) return out;
  try {
    const lines = glassPo.readGlassLines() || [];
    const store = JSON.parse(fs.readFileSync(path.join(dataDir(), "glass-pos.json"), "utf8"));
    const lineMeta = (store && store.lines) || {};
    const pos = Array.isArray(store && store.pos) ? store.pos : [];
    const receives = Array.isArray(store && store.receives) ? store.receives : [];
    lines.forEach((line) => {
      if (!sameOrder(line.order, want)) return;
      const spec = [line.type, line.thickness, line.isTemplate ? "template" : "", line.component]
        .filter(Boolean).join(" · ");
      if (line.timestamp) {
        out.push(eventOf({
          id: "glass-log-" + (line.id || line.row),
          at: line.timestamp,
          stage: "materials",
          title: "Glass logged",
          detail: spec || line.status || "",
          actor: line.worker,
          order_number: order.order_number,
          kind: "glass_logged"
        }));
      }
      const meta = lineMeta[line.id] || {};
      if (meta.poAt) {
        out.push(eventOf({
          id: "glass-po-" + (meta.poId || line.id),
          at: meta.poAt,
          stage: "materials",
          title: "Glass on purchase order",
          detail: [meta.poNumber, spec].filter(Boolean).join(" · "),
          actor: meta.poBy,
          order_number: order.order_number,
          kind: "glass_po"
        }));
      }
    });
    pos.forEach((po) => {
      const hit = (po.lineIds || []).some((id) => {
        const line = lines.find((l) => l.id === id);
        return line && sameOrder(line.order, want);
      });
      if (!hit || !po.createdAt) return;
      out.push(eventOf({
        id: "glass-po-" + po.id,
        at: po.createdAt,
        stage: "materials",
        title: "Glass purchase order " + (po.number || ""),
        actor: po.createdBy,
        order_number: order.order_number,
        kind: "glass_po"
      }));
    });
    receives.forEach((batch) => {
      const batchLines = batch.lines || [];
      const hit = batchLines.some((entry) => {
        const line = lines.find((l) => l.id === entry.id);
        return line && sameOrder(line.order, want);
      });
      if (!hit || !batch.receivedAt) return;
      out.push(eventOf({
        id: "glass-recv-" + batch.id,
        at: batch.receivedAt,
        stage: "materials",
        title: "Glass received",
        actor: batch.receivedBy,
        order_number: order.order_number,
        kind: "glass_received"
      }));
    });
  } catch (e) {
    try {
      (glassPo.readGlassLines() || []).forEach((line) => {
        if (!sameOrder(line.order, want) || !line.timestamp) return;
        out.push(eventOf({
          id: "glass-log-" + (line.id || line.row),
          at: line.timestamp,
          stage: "materials",
          title: "Glass logged",
          detail: [line.type, line.status].filter(Boolean).join(" · "),
          actor: line.worker,
          order_number: order.order_number,
          kind: "glass_logged"
        }));
      });
    } catch (err) {}
  }
  return out;
}

function readWoodLines() {
  const book = getBook();
  const sheet = book.getSheetByName("Wood_To_Order");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lastCol = Math.max(sheet.getLastColumn(), WOOD_HEADERS.length);
  const grid = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  return grid.map((row, i) => {
    const id = String(row[0] || "").trim();
    const typeName = String(row[5] || "").trim();
    if (!id && !typeName) return null;
    return {
      id,
      timestamp: row[1] ? toIso(row[1]) : "",
      order: String(row[2] || ""),
      worker: String(row[3] || ""),
      component: String(row[4] || ""),
      type: typeName,
      thickness: String(row[6] || ""),
      status: String(row[10] || "").trim(),
      row: i + 2
    };
  }).filter(Boolean);
}

function woodEvents(order) {
  const out = [];
  const want = order && order.order_number;
  if (!want) return out;
  try {
    readWoodLines().forEach((line) => {
      if (!sameOrder(line.order, want) || !line.timestamp) return;
      const received = /^received$/i.test(line.status);
      const spec = [line.type, line.thickness, line.component, line.status].filter(Boolean).join(" · ");
      out.push(eventOf({
        id: "wood-" + (line.id || line.row),
        at: line.timestamp,
        stage: "materials",
        title: received ? "Wood received" : "Wood logged",
        detail: spec,
        actor: line.worker,
        order_number: order.order_number,
        kind: received ? "wood_received" : "wood_logged"
      }));
    });
  } catch (e) {}
  return out;
}

function paintEvents(order) {
  const out = [];
  const want = formatOrderId(order && order.order_number);
  if (!want) return out;
  try {
    const shop = paintShop.loadShop();
    const meta = (shop.orders && shop.orders[want]) || {};
    if (meta.sentAt) {
      out.push(eventOf({
        id: "paint-sent-" + want,
        at: meta.sentAt,
        stage: "paint",
        title: "Sent to paint shop",
        actor: meta.sentBy,
        order_number: order.order_number,
        kind: "paint_sent"
      }));
    }
    if (meta.receivedAt) {
      out.push(eventOf({
        id: "paint-recv-" + want,
        at: meta.receivedAt,
        stage: "paint",
        title: "Received from paint shop",
        actor: meta.receivedBy,
        order_number: order.order_number,
        kind: "paint_received"
      }));
    }
    (shop.receives || []).forEach((batch) => {
      const hit = (batch.orders || []).some((line) => formatOrderId(line.orderNumber || line.order_number) === want);
      if (!hit || !batch.receivedAt || meta.receivedAt) return;
      out.push(eventOf({
        id: "paint-recv-" + batch.id,
        at: batch.receivedAt,
        stage: "paint",
        title: "Received from paint shop",
        actor: batch.receivedBy,
        order_number: order.order_number,
        kind: "paint_received"
      }));
    });
  } catch (e) {}
  return out;
}

function paymentEvents(order) {
  const out = [];
  (Array.isArray(order && order.payments) ? order.payments : []).forEach((p, i) => {
    if (!p || !toIso(p.at)) return;
    out.push(eventOf({
      id: "pay-" + (p.id || i),
      at: p.at,
      stage: "money",
      title: "Payment received",
      detail: [p.amount ? formatRand(p.amount) : "", p.note].filter(Boolean).join(" · "),
      order_number: order.order_number,
      kind: "payment"
    }));
  });
  return out;
}

function deliveryEvents(order) {
  const out = [];
  const want = formatOrderId(order && order.order_number);
  if (!want) return out;
  try {
    const { listSchedule } = require("./db");
    const rows = listSchedule("2020-01-01", "2035-12-31") || [];
    rows.forEach((row) => {
      if (formatOrderId(row && row.order_number) !== want) return;
      (row.delivery_days || []).forEach((day, i) => {
        const code = String((row.cells && row.cells[day]) || "").trim().toUpperCase();
        const label = code === "LC" ? "Latest Courier" : (code === "LD" ? "Latest Delivery" : "Delivery");
        out.push(eventOf({
          id: "delivery-" + want + "-" + day + "-" + i,
          at: day,
          stage: "delivery",
          title: label + " " + day,
          detail: [row.courier, row.waybill].filter(Boolean).join(" · "),
          order_number: order.order_number,
          kind: "delivery"
        }));
      });
    });
  } catch (e) {}
  return out;
}

function sortEvents(events) {
  const seen = {};
  return (events || []).filter((ev) => {
    if (!ev || !ev.title) return false;
    const key = ev.id || (ev.at + "|" + ev.title + "|" + ev.order_number);
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  }).sort((a, b) => {
    const at = String(a.at || "").localeCompare(String(b.at || ""));
    if (at) return at;
    return String(a.id || "").localeCompare(String(b.id || ""));
  });
}

function shopSideEvents(order) {
  if (!order) return [];
  return []
    .concat(orderOpenedEvents(order))
    .concat(jobCardEvents(order))
    .concat(productionLogEvents(order))
    .concat(glassEvents(order))
    .concat(woodEvents(order))
    .concat(paintEvents(order))
    .concat(paymentEvents(order))
    .concat(deliveryEvents(order));
}

function collectOrderLife(orderNumber) {
  const order = findOrder(orderNumber);
  if (!order) throw new Error("Order not found");
  const enquiry = order.enquiry_no ? getEnquiry(order.enquiry_no) : null;
  const events = sortEvents(
    enquiryEvents(enquiry, order.order_number).concat(shopSideEvents(order))
  );
  return {
    order_number: order.order_number,
    enquiry_no: (enquiry && enquiry.enquiry_no) || order.enquiry_no || "",
    status: order.status || "",
    product: order.product || "",
    client_name: order.client_name || "",
    enquiry_status: (enquiry && enquiry.status) || "",
    events
  };
}

function collectEnquiryLife(enquiryNo) {
  const enquiry = getEnquiry(enquiryNo);
  if (!enquiry) throw new Error("Enquiry not found");
  const orders = ordersLinkedToEnquiry(enquiry);
  let events = enquiryEvents(enquiry, enquiry.order_number || "");
  orders.forEach((order) => {
    events = events.concat(shopSideEvents(order));
  });
  return {
    enquiry_no: enquiry.enquiry_no,
    order_numbers: orders.map((o) => o.order_number).filter(Boolean),
    status: enquiry.status || "",
    product: enquiry.product || "",
    client_name: enquiry.client_name || "",
    events: sortEvents(events)
  };
}

module.exports = {
  collectOrderLife,
  collectEnquiryLife,
  sameOrder
};
