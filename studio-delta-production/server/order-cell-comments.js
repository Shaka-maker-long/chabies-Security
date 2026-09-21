"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir } = require("./workbook-store");
const { formatOrderId, listOrders } = require("./db");
const pipeline = require("./enquiry-pipeline");

const COMMENTABLE_FIELDS = [
  "quote_number", "order_number", "status", "assigned_operator", "type", "category",
  "product", "variation", "doors", "detailed_description", "dimensions", "powder_coating",
  "client_name", "client_number", "email", "payment_date", "address", "province",
  "price_excl_vat", "price_incl_vat", "amount_paid", "owing", "month_of_sale", "source", "city"
];

const FIELD_LABELS = {
  quote_number: "QUOTE NUMBER",
  order_number: "ORDER NUMBER",
  status: "STATUS",
  assigned_operator: "ASSIGNED OPERATOR",
  type: "TYPE",
  category: "CATERGORY",
  product: "PRODUCT",
  variation: "VARIATION",
  doors: "DOORS",
  detailed_description: "DETAILED DESCRIPTION",
  dimensions: "DIMENSIONS",
  powder_coating: "POWDER COATING",
  client_name: "CLIENT NAME",
  client_number: "CLIENT NUMBER",
  email: "EMAIL ADDRESS",
  payment_date: "PAYMENT DATE",
  address: "ADDRESS",
  province: "PROVINCE",
  price_excl_vat: "PRICE (Excl VAT)",
  price_incl_vat: "PRICE (Incl VAT)",
  amount_paid: "AMOUNT PAID (Incl VAT)",
  owing: "OWING",
  month_of_sale: "MONTH OF SALE",
  source: "SOURCE",
  city: "CITY"
};

function storePath() {
  return path.join(dataDir(), "order-cell-comments.json");
}

function emptyStore() {
  return { comments: [] };
}

function loadStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(storePath(), "utf8"));
    const store = emptyStore();
    store.comments = Array.isArray(parsed.comments) ? parsed.comments.map(normalizeComment).filter(Boolean) : [];
    return store;
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[order-cell-comments] could not read", storePath(), e.message || e);
    }
    return emptyStore();
  }
}

function saveStore(store) {
  const file = storePath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify({ comments: (store.comments || []).map(normalizeComment).filter(Boolean) }));
  fs.renameSync(tmp, file);
  return store;
}

function nowIso() {
  return new Date().toISOString();
}

function newId(prefix) {
  return prefix + "_" + crypto.randomBytes(8).toString("hex");
}

function namesMatch(a, b) {
  return String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase();
}

function requireAssignee(name) {
  const n = String(name || "").trim();
  if (!n) throw new Error("Choose the office person this comment is assigned to");
  const hit = pipeline.officeAssignees().find((x) => namesMatch(x, n));
  if (!hit) throw new Error("Assign to an office Admin from Users");
  return hit;
}

function normalizeReply(raw) {
  if (!raw || typeof raw !== "object") return null;
  const body = String(raw.body || "").trim();
  if (!body) return null;
  return {
    id: String(raw.id || newId("ocr")).trim() || newId("ocr"),
    body,
    by: String(raw.by || "").trim(),
    at: String(raw.at || nowIso())
  };
}

function normalizeComment(raw) {
  if (!raw || typeof raw !== "object") return null;
  const orderNumber = formatOrderId(raw.order_number || raw.orderNumber);
  const field = String(raw.field || "").trim();
  if (!orderNumber || !field || COMMENTABLE_FIELDS.indexOf(field) === -1) return null;
  const body = String(raw.body || "").trim();
  if (!body) return null;
  const replies = Array.isArray(raw.replies)
    ? raw.replies.map(normalizeReply).filter(Boolean)
    : [];
  return {
    id: String(raw.id || newId("occ")).trim() || newId("occ"),
    order_number: orderNumber,
    field,
    field_label: FIELD_LABELS[field] || field,
    body,
    created_by: String(raw.created_by || raw.createdBy || "").trim(),
    created_at: String(raw.created_at || raw.createdAt || nowIso()),
    assignee: String(raw.assignee || "").trim(),
    replies,
    resolved: !!raw.resolved,
    resolved_by: String(raw.resolved_by || raw.resolvedBy || "").trim(),
    resolved_at: String(raw.resolved_at || raw.resolvedAt || "")
  };
}

function assertOrderExists(orderNumber) {
  const want = formatOrderId(orderNumber);
  if (!want) throw new Error("Order number is required.");
  const hit = listOrders().find((o) => formatOrderId(o.order_number) === want);
  if (!hit) throw new Error("Order not found.");
  return want;
}

function assertField(field) {
  const key = String(field || "").trim();
  if (COMMENTABLE_FIELDS.indexOf(key) === -1) {
    throw new Error("Pick a sheet column to comment on.");
  }
  return key;
}

function listComments(query) {
  const store = loadStore();
  const order = formatOrderId(query && (query.order || query.order_number));
  const field = String((query && query.field) || "").trim();
  const assignee = String((query && query.assignee) || "").trim();
  const openOnly = String((query && query.open) || "") === "1"
    || String((query && query.open) || "").toLowerCase() === "true";
  let rows = store.comments.slice();
  if (order) rows = rows.filter((c) => formatOrderId(c.order_number) === order);
  if (field) rows = rows.filter((c) => c.field === field);
  if (assignee) rows = rows.filter((c) => namesMatch(c.assignee, assignee));
  if (openOnly) rows = rows.filter((c) => !c.resolved);
  rows.sort((a, b) => String(b.created_at).localeCompare(String(a.created_at)));
  return {
    rows,
    openCount: store.comments.filter((c) => !c.resolved).length,
    fields: COMMENTABLE_FIELDS.map((f) => ({ id: f, label: FIELD_LABELS[f] || f }))
  };
}

function cellSummary() {
  const open = loadStore().comments.filter((c) => !c.resolved);
  const map = {};
  open.forEach((c) => {
    const key = formatOrderId(c.order_number) + "::" + c.field;
    if (!map[key]) {
      map[key] = {
        order_number: c.order_number,
        field: c.field,
        count: 0,
        assignees: []
      };
    }
    map[key].count += 1;
    if (c.assignee && map[key].assignees.indexOf(c.assignee) === -1) {
      map[key].assignees.push(c.assignee);
    }
  });
  return Object.keys(map).map((k) => map[k]);
}

function createComment(body, actor) {
  const orderNumber = assertOrderExists(body && body.order_number);
  const field = assertField(body && body.field);
  const text = String((body && body.body) || "").trim();
  if (!text) throw new Error("Type a comment.");
  const assignee = requireAssignee((body && body.assignee) || "");
  const actorName = String(actor || "").trim() || "Office";
  const comment = normalizeComment({
    id: newId("occ"),
    order_number: orderNumber,
    field,
    body: text,
    created_by: actorName,
    created_at: nowIso(),
    assignee,
    replies: [],
    resolved: false
  });
  const store = loadStore();
  store.comments.unshift(comment);
  saveStore(store);
  return comment;
}

function findComment(id) {
  const want = String(id || "").trim();
  if (!want) return null;
  const store = loadStore();
  const idx = store.comments.findIndex((c) => c.id === want);
  if (idx < 0) return null;
  return { store, idx, comment: store.comments[idx] };
}

function replyToComment(id, body, actor) {
  const hit = findComment(id);
  if (!hit) throw new Error("Comment not found.");
  if (hit.comment.resolved) throw new Error("This comment is resolved. Re-open it before replying.");
  const text = String((body && body.body) || "").trim();
  if (!text) throw new Error("Type a reply.");
  const reply = normalizeReply({
    id: newId("ocr"),
    body: text,
    by: String(actor || "").trim() || "Office",
    at: nowIso()
  });
  hit.comment.replies.push(reply);
  hit.store.comments[hit.idx] = hit.comment;
  saveStore(hit.store);
  return hit.comment;
}

function resolveComment(id, actor) {
  const hit = findComment(id);
  if (!hit) throw new Error("Comment not found.");
  if (hit.comment.resolved) return hit.comment;
  hit.comment.resolved = true;
  hit.comment.resolved_by = String(actor || "").trim() || "Office";
  hit.comment.resolved_at = nowIso();
  hit.store.comments[hit.idx] = hit.comment;
  saveStore(hit.store);
  return hit.comment;
}

function reopenComment(id, actor) {
  const hit = findComment(id);
  if (!hit) throw new Error("Comment not found.");
  hit.comment.resolved = false;
  hit.comment.resolved_by = "";
  hit.comment.resolved_at = "";
  if (actor) {
    hit.comment.replies.push(normalizeReply({
      id: newId("ocr"),
      body: "Re-opened",
      by: String(actor).trim(),
      at: nowIso()
    }));
  }
  hit.store.comments[hit.idx] = hit.comment;
  saveStore(hit.store);
  return hit.comment;
}

function reassignComment(id, assignee, actor) {
  const hit = findComment(id);
  if (!hit) throw new Error("Comment not found.");
  if (hit.comment.resolved) throw new Error("This comment is resolved.");
  const next = requireAssignee(assignee);
  const prev = hit.comment.assignee;
  hit.comment.assignee = next;
  if (!namesMatch(prev, next)) {
    hit.comment.replies.push(normalizeReply({
      id: newId("ocr"),
      body: "Assigned to " + next + (prev ? " (was " + prev + ")" : ""),
      by: String(actor || "").trim() || "Office",
      at: nowIso()
    }));
  }
  hit.store.comments[hit.idx] = hit.comment;
  saveStore(hit.store);
  return hit.comment;
}

function dropCommentsForOrder(orderNumber) {
  const want = formatOrderId(orderNumber);
  if (!want) return 0;
  const store = loadStore();
  const before = store.comments.length;
  store.comments = store.comments.filter((c) => formatOrderId(c.order_number) !== want);
  const removed = before - store.comments.length;
  if (removed) saveStore(store);
  return removed;
}

function dropAllComments() {
  const store = loadStore();
  const removed = store.comments.length;
  if (removed) saveStore(emptyStore());
  return removed;
}

function assignedTo(name) {
  const who = String(name || "").trim();
  if (!who) return [];
  return listComments({ assignee: who, open: "1" }).rows;
}

module.exports = {
  COMMENTABLE_FIELDS,
  FIELD_LABELS,
  listComments,
  cellSummary,
  createComment,
  replyToComment,
  resolveComment,
  reopenComment,
  reassignComment,
  dropCommentsForOrder,
  dropAllComments,
  assignedTo,
  loadStore,
  saveStore
};
