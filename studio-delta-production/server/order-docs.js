"use strict";

const { listEnquiries, formatOrderId, drawingIsRequired, drawingFilePresent } = require("./db");
const { normalizeBaseOrderNumber } = require("./create-order-from-enquiry");
const { getBook } = require("./workbook-store");

function qcLabel(process) {
  const s = String(process || "").toLowerCase();
  if (s.indexOf("final") !== -1) return "Final QC PDF";
  if (s.indexOf("pre-powder") !== -1 || s.indexOf("pre powder") !== -1) return "Pre-powder QC PDF";
  return "QC PDF";
}

function addUnique(list, item) {
  if (!item || !item.url) return;
  if (list.some((row) => row.url === item.url)) return;
  list.push(item);
}

function collectQcPdfs() {
  const map = {};
  try {
    const sheet = getBook().getSheetByName("Production_Log");
    if (sheet && sheet.getLastRow() >= 2) {
      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        const order = formatOrderId(values[i][1]);
        if (!order) continue;
        const notes = String(values[i][7] || "");
        const process = String(values[i][3] || "").trim();
        const re = /QC PDF:\s*(\S+)/gi;
        let hit;
        while ((hit = re.exec(notes))) {
          if (!map[order]) map[order] = [];
          addUnique(map[order], { url: hit[1], label: qcLabel(process) });
        }
      }
    }
  } catch (e) {}
  try {
    require("./qc-pdf").listReports().forEach((row) => {
      const order = formatOrderId(row.order_number);
      if (!order) return;
      if (!map[order]) map[order] = [];
      addUnique(map[order], { url: row.url, label: row.label || qcLabel(row.kind) });
    });
  } catch (e) {}
  Object.keys(map).forEach((order) => {
    const list = map[order] || [];
    if (list.some((row) => String(row.url || "").indexOf("/api/qc-pdfs/") === 0)) {
      map[order] = list.filter((row) => String(row.url || "").indexOf("/api/qc-pdfs/") === 0);
    }
  });
  return map;
}

function enquiryIndex() {
  const byNo = new Map();
  const byQuote = new Map();
  const byOrder = new Map();
  listEnquiries().forEach((row) => {
    const no = String(row.enquiry_no || "").trim();
    if (no) byNo.set(no.toLowerCase(), row);
    const quote = String(row.quote_no || row.quote_number || "").trim().toLowerCase();
    if (quote) byQuote.set(quote, row);
    const order = formatOrderId(row.order_number || "");
    if (order) {
      byOrder.set(order, row);
      const base = normalizeBaseOrderNumber(order);
      if (base) byOrder.set(base, row);
    }
  });
  return { byNo, byQuote, byOrder };
}

function findEnquiry(order, index) {
  const no = String((order && order.enquiry_no) || "").trim();
  if (no && index.byNo.has(no.toLowerCase())) return index.byNo.get(no.toLowerCase());
  const quote = String((order && order.quote_number) || "").trim().toLowerCase();
  if (quote && index.byQuote.has(quote)) return index.byQuote.get(quote);
  const id = formatOrderId(order && order.order_number);
  if (id && index.byOrder.has(id)) return index.byOrder.get(id);
  const base = normalizeBaseOrderNumber(id);
  if (base && index.byOrder.has(base)) return index.byOrder.get(base);
  return null;
}

function drawingDoc(enquiry) {
  if (!enquiry) return { drawing_url: "", drawing_required: false };
  const required = drawingIsRequired(enquiry.drawing);
  const stored = drawingFilePresent(enquiry.drawing);
  return {
    drawing_url: stored
      ? "/api/office/enquiries/" + encodeURIComponent(enquiry.enquiry_no) + "/files/drawing"
      : "",
    drawing_required: required
  };
}

function qcForOrder(order, qcMap) {
  const id = formatOrderId(order && order.order_number);
  if (!id) return [];
  const base = normalizeBaseOrderNumber(id);
  const out = [];
  (qcMap[id] || []).forEach((item) => addUnique(out, item));
  if (base && base !== id) (qcMap[base] || []).forEach((item) => addUnique(out, item));
  return out;
}

function attachOrderDocs(rows) {
  const index = enquiryIndex();
  const qcMap = collectQcPdfs();
  (rows || []).forEach((row) => {
    const draw = drawingDoc(findEnquiry(row, index));
    row.drawing_url = draw.drawing_url;
    row.drawing_required = draw.drawing_required;
    row.qc_pdfs = qcForOrder(row, qcMap).slice(0, 2);
  });
  return rows;
}

module.exports = { attachOrderDocs, drawingDoc, collectQcPdfs };
