"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir } = require("./workbook-store");
const db = require("./db");
const staff = require("./staff");

const STATUS = {
  NEEDS_ASSIGNMENT: "Needs assignment",
  READY: "Ready for Rework",
  IN_PROGRESS: "Rework",
  DONE: "Done"
};

function nowIso() {
  return new Date().toISOString();
}

function reworksPath() {
  return path.join(dataDir(), "reworks.json");
}

function emptyStore() {
  return { reworks: [] };
}

function loadStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(reworksPath(), "utf8"));
    const store = emptyStore();
    store.reworks = Array.isArray(parsed.reworks) ? parsed.reworks : [];
    return store;
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[reworks] could not read", reworksPath(), e.message || e);
    }
    return emptyStore();
  }
}

function saveStore(store) {
  const file = reworksPath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(store, null, 2));
  fs.renameSync(tmp, file);
  return store;
}

function canAssignReworks(profile) {
  if (!profile) return false;
  if (staff.canManageUsers(profile)) return true;
  return staff.canSeeIdleAlerts(profile);
}

function isProductionOperator(profile) {
  if (!profile || !String(profile.name || "").trim()) return false;
  if (staff.isMarketing(profile)) return false;
  return true;
}

function listOperators() {
  return staff.listUsers()
    .filter(isProductionOperator)
    .map((u) => u.name)
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
}

function snapshotFromOrder(order) {
  return {
    order_number: String((order && order.order_number) || "").trim(),
    shop_status: String((order && order.status) || "").trim(),
    type: String((order && order.type) || "").trim(),
    category: String((order && order.category) || "").trim(),
    product: String((order && order.product) || "").trim(),
    variation: String((order && order.variation) || "").trim(),
    doors: String((order && order.doors) || "").trim(),
    detailed_description: String((order && order.detailed_description) || "").trim(),
    dimensions: String((order && order.dimensions) || "").trim(),
    powder_coating: String((order && order.powder_coating) || "").trim()
  };
}

function normalizeRework(row) {
  const r = row && typeof row === "object" ? row : {};
  return {
    id: String(r.id || "").trim(),
    order_number: String(r.order_number || "").trim(),
    status: String(r.status || STATUS.NEEDS_ASSIGNMENT).trim() || STATUS.NEEDS_ASSIGNMENT,
    assigned_operator: String(r.assigned_operator || "").trim(),
    issue: String(r.issue || r.comment || "").trim(),
    type: String(r.type || "").trim(),
    category: String(r.category || "").trim(),
    product: String(r.product || "").trim(),
    variation: String(r.variation || "").trim(),
    doors: String(r.doors || "").trim(),
    detailed_description: String(r.detailed_description || "").trim(),
    dimensions: String(r.dimensions || "").trim(),
    powder_coating: String(r.powder_coating || "").trim(),
    shop_status: String(r.shop_status || "").trim(),
    created_at: String(r.created_at || "").trim(),
    created_by: String(r.created_by || "").trim(),
    assigned_at: String(r.assigned_at || "").trim(),
    assigned_by: String(r.assigned_by || "").trim(),
    started_at: String(r.started_at || "").trim(),
    finished_at: String(r.finished_at || "").trim(),
    finished_by: String(r.finished_by || "").trim()
  };
}

function listReworks() {
  return loadStore().reworks.map(normalizeRework);
}

function findRework(id) {
  const want = String(id || "").trim();
  if (!want) return null;
  return listReworks().find((r) => r.id === want) || null;
}

function findOpenReworkForOrder(orderNumber, assignedOnly) {
  const want = String(orderNumber || "").trim().toLowerCase();
  if (!want) return null;
  return listReworks().find((r) => {
    if (String(r.order_number || "").trim().toLowerCase() !== want) return false;
    if (r.status === STATUS.DONE) return false;
    if (assignedOnly && !r.assigned_operator) return false;
    return true;
  }) || null;
}

function listOpenForFloor(workerName) {
  const who = String(workerName || "").trim().toLowerCase();
  return listReworks().filter((r) => {
    if (r.status === STATUS.DONE) return false;
    if (!r.assigned_operator) return false;
    if (who && String(r.assigned_operator).trim().toLowerCase() !== who) return false;
    return true;
  });
}

function createRework(body, actor) {
  const orderNo = String((body && (body.order_number || body.order)) || "").trim();
  if (!orderNo) throw new Error("Pick an order for this rework");
  const issue = String((body && (body.issue || body.comment)) || "").trim();
  if (!issue) throw new Error("Add a comment on what the issue was");
  const order = db.listOrders().find((o) => String(o.order_number || "").trim() === orderNo);
  if (!order) throw new Error("Order not found: " + orderNo);
  const open = findOpenReworkForOrder(orderNo, false);
  if (open) throw new Error("This order already has an open rework (" + open.status + ")");

  const snap = snapshotFromOrder(order);
  const row = normalizeRework(Object.assign({}, snap, {
    id: "RW-" + crypto.randomBytes(4).toString("hex").toUpperCase(),
    status: STATUS.NEEDS_ASSIGNMENT,
    assigned_operator: "",
    issue,
    created_at: nowIso(),
    created_by: String(actor || "").trim()
  }));

  const store = loadStore();
  store.reworks.unshift(row);
  saveStore(store);
  return row;
}

function assignRework(id, body, actor, profile) {
  const whoProfile = profile || staff.listUsers().find((u) =>
    String(u.name || "").trim().toLowerCase() === String(actor || "").trim().toLowerCase()
  ) || { name: actor };
  if (!canAssignReworks(whoProfile)) {
    throw new Error("Only Siya or the Manager can assign a rework");
  }
  const who = String((body && (body.assigned_operator || body.assignee || body.operator)) || "").trim();
  if (!who) throw new Error("Pick the production person for this rework");
  const operators = listOperators();
  if (!operators.some((n) => String(n).toLowerCase() === who.toLowerCase())) {
    throw new Error("Choose someone from the production people list");
  }

  const store = loadStore();
  const idx = store.reworks.findIndex((r) => String(r.id || "").trim() === String(id || "").trim());
  if (idx < 0) throw new Error("Rework not found");
  const row = normalizeRework(store.reworks[idx]);
  if (row.status === STATUS.DONE) throw new Error("This rework is already done");
  row.assigned_operator = who;
  row.status = STATUS.READY;
  row.assigned_at = nowIso();
  row.assigned_by = String(actor || "").trim();
  store.reworks[idx] = row;
  saveStore(store);
  return row;
}

function markReworkStarted(id, workerName) {
  const store = loadStore();
  const idx = store.reworks.findIndex((r) => String(r.id || "").trim() === String(id || "").trim());
  if (idx < 0) return null;
  const row = normalizeRework(store.reworks[idx]);
  if (row.status === STATUS.DONE) return row;
  if (workerName && row.assigned_operator
      && String(row.assigned_operator).toLowerCase() !== String(workerName).toLowerCase()) {
    throw new Error("This rework is assigned to " + row.assigned_operator);
  }
  row.status = STATUS.IN_PROGRESS;
  if (!row.started_at) row.started_at = nowIso();
  store.reworks[idx] = row;
  saveStore(store);
  return row;
}

function markReworkFinished(id, workerName) {
  const store = loadStore();
  const idx = store.reworks.findIndex((r) => String(r.id || "").trim() === String(id || "").trim());
  if (idx < 0) return null;
  const row = normalizeRework(store.reworks[idx]);
  row.status = STATUS.DONE;
  row.finished_at = nowIso();
  row.finished_by = String(workerName || "").trim();
  store.reworks[idx] = row;
  saveStore(store);
  return row;
}

function markReworkFinishedForOrder(orderNumber, workerName) {
  const open = findOpenReworkForOrder(orderNumber, true);
  if (!open) return null;
  return markReworkFinished(open.id, workerName);
}

function pagePayload() {
  const orders = db.listOrders()
    .map((o) => ({
      order_number: String(o.order_number || "").trim(),
      status: String(o.status || "").trim(),
      product: String(o.product || "").trim(),
      client_name: String(o.client_name || "").trim()
    }))
    .filter((o) => o.order_number)
    .sort((a, b) => String(b.order_number).localeCompare(String(a.order_number)));
  return {
    reworks: listReworks(),
    orders,
    operators: listOperators(),
    statuses: STATUS
  };
}

module.exports = {
  STATUS,
  canAssignReworks,
  listReworks,
  listOperators,
  listOpenForFloor,
  findRework,
  findOpenReworkForOrder,
  createRework,
  assignRework,
  markReworkStarted,
  markReworkFinished,
  markReworkFinishedForOrder,
  pagePayload,
  snapshotFromOrder
};
