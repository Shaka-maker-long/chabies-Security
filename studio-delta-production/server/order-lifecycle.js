"use strict";

const fs = require("fs");
const path = require("path");
const { dataDir } = require("./workbook-store");
const db = require("./db");
const shop = require("./shop-status");

const CANCELLED = shop.CANCELLED;
const ON_HOLD = shop.ON_HOLD;

function storePath() {
  return path.join(dataDir(), "order-lifecycle.json");
}

function emptyStore() {
  return { orders: {} };
}

function loadStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(storePath(), "utf8"));
    const store = emptyStore();
    store.orders = parsed && parsed.orders && typeof parsed.orders === "object" ? parsed.orders : {};
    return store;
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[order-lifecycle] could not read", storePath(), e.message || e);
    }
    return emptyStore();
  }
}

function saveStore(store) {
  const file = storePath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(store));
  fs.renameSync(tmp, file);
  return store;
}

function recOf(store, orderNumber) {
  const id = db.formatOrderId(orderNumber);
  if (!id) return null;
  const rec = store.orders[id];
  return rec && typeof rec === "object" ? rec : null;
}

function findLiveOrder(orderNumber) {
  const id = db.formatOrderId(orderNumber);
  if (!id) throw new Error("Order number is required.");
  const live = db.listOrders().find((o) => db.formatOrderId(o.order_number) === id);
  if (!live) throw new Error(id + " is not on Orders.");
  return live;
}

function closeShopWork(orderNumber) {
  try {
    require("./order-correct").closeOpenLogsForOrder(orderNumber);
  } catch (e) {}
  try {
    require("./floor-planning").unscheduleOrder(orderNumber);
  } catch (e) {}
}

function persistPark(live, status, rec) {
  const saved = db.upsertOrder(Object.assign({}, live, { status }));
  try { db.persistOffice(); } catch (e) {}
  try { require("./gas").clearShopCache(); } catch (e) {}
  return attach(saved, rec);
}

function cancelOrder(orderNumber, reason, actor) {
  const why = String(reason || "").trim();
  if (!why) throw new Error("Give a reason for cancelling this order.");
  const live = findLiveOrder(orderNumber);
  const id = db.formatOrderId(live.order_number);
  if (shop.isCancelled(live.status)) throw new Error(id + " is already cancelled.");
  const store = loadStore();
  const prev = recOf(store, id) || {};
  const rec = {
    state: "cancelled",
    reason: why,
    parked_from_status: shop.isOnHold(live.status)
      ? (prev.parked_from_status || "Not Yet Started")
      : (String(live.status || "").trim() || "Not Yet Started"),
    actor: String(actor || "").trim(),
    at: new Date().toISOString(),
    hold_reason: prev.hold_reason || "",
    hold_at: prev.hold_at || "",
    hold_actor: prev.hold_actor || ""
  };
  store.orders[id] = rec;
  saveStore(store);
  closeShopWork(id);
  return persistPark(live, CANCELLED, rec);
}

function holdOrder(orderNumber, reason, actor) {
  const why = String(reason || "").trim();
  if (!why) throw new Error("Give a reason for putting this order on hold.");
  const live = findLiveOrder(orderNumber);
  const id = db.formatOrderId(live.order_number);
  if (shop.isCancelled(live.status)) {
    throw new Error("Resume " + id + " from Cancelled before putting it on hold.");
  }
  if (shop.isOnHold(live.status)) throw new Error(id + " is already on hold.");
  const store = loadStore();
  const rec = {
    state: "on_hold",
    reason: why,
    hold_reason: why,
    parked_from_status: String(live.status || "").trim() || "Not Yet Started",
    actor: String(actor || "").trim(),
    at: new Date().toISOString(),
    hold_at: new Date().toISOString(),
    hold_actor: String(actor || "").trim()
  };
  store.orders[id] = rec;
  saveStore(store);
  closeShopWork(id);
  return persistPark(live, ON_HOLD, rec);
}

function resumeOrder(orderNumber, actor) {
  const live = findLiveOrder(orderNumber);
  const id = db.formatOrderId(live.order_number);
  if (!shop.isParked(live.status)) throw new Error(id + " is not cancelled or on hold.");
  const store = loadStore();
  const prev = recOf(store, id) || {};
  let nextStatus = String(prev.parked_from_status || "").trim();
  if (!nextStatus || shop.isParked(nextStatus) || nextStatus === CANCELLED || nextStatus === ON_HOLD) {
    nextStatus = "Not Yet Started";
  }
  if (shop.normalizeShopStatus(nextStatus)) nextStatus = shop.normalizeShopStatus(nextStatus);
  delete store.orders[id];
  saveStore(store);
  const saved = db.upsertOrder(Object.assign({}, live, {
    status: nextStatus,
    assigned_operator: live.assigned_operator || String(actor || "").trim()
  }));
  try { db.persistOffice(); } catch (e) {}
  try { require("./gas").clearShopCache(); } catch (e) {}
  return attach(saved, null);
}

function attach(order, recOverride) {
  if (!order) return order;
  const id = db.formatOrderId(order.order_number);
  const rec = recOverride !== undefined ? recOverride : recOf(loadStore(), id);
  const cancelled = shop.isCancelled(order.status);
  const held = shop.isOnHold(order.status);
  const owing = db.orderOwing(order);
  return Object.assign({}, order, {
    cancelled: cancelled,
    on_hold: held,
    parked: cancelled || held,
    cancel_reason: cancelled ? String((rec && (rec.reason || rec.cancel_reason)) || "") : "",
    hold_reason: held ? String((rec && (rec.hold_reason || rec.reason)) || "") : "",
    parked_from_status: rec && rec.parked_from_status ? rec.parked_from_status : "",
    parked_at: rec && rec.at ? rec.at : "",
    parked_by: rec && rec.actor ? rec.actor : "",
    owing_amount: owing,
    payment_blocked: owing > 0
  });
}

function listCancelled() {
  return db.listOrders()
    .filter((o) => shop.isCancelled(o.status))
    .map((o) => db.decorateMoney(attach(o)))
    .sort((a, b) => String(b.parked_at || "").localeCompare(String(a.parked_at || ""))
      || String(a.order_number).localeCompare(String(b.order_number)));
}

function listOnHold() {
  return db.listOrders()
    .filter((o) => shop.isOnHold(o.status))
    .map((o) => db.decorateMoney(attach(o)))
    .sort((a, b) => String(b.parked_at || "").localeCompare(String(a.parked_at || ""))
      || String(a.order_number).localeCompare(String(b.order_number)));
}

function isSheetHidden(order) {
  return shop.isCancelled(order && order.status);
}

function parkEvents(order) {
  const rec = recOf(loadStore(), order && order.order_number);
  if (!rec || !rec.at) return [];
  const cancelled = shop.isCancelled(order && order.status);
  const held = shop.isOnHold(order && order.status);
  if (!cancelled && !held) return [];
  return [{
    id: "park-" + db.formatOrderId(order.order_number) + "-" + rec.state,
    at: rec.at,
    at_label: db.formatSastDateTime(rec.at) || "",
    stage: "order",
    title: cancelled ? "Order cancelled" : "Order on hold",
    detail: rec.reason || rec.hold_reason || "",
    actor: rec.actor || "",
    order_number: order.order_number,
    kind: rec.state
  }];
}

function renameOrder(fromRaw, toRaw) {
  const fromId = db.formatOrderId(fromRaw);
  const toId = db.formatOrderId(toRaw);
  if (!fromId || !toId || fromId === toId) return false;
  const store = loadStore();
  if (!store.orders[fromId]) return false;
  store.orders[toId] = store.orders[fromId];
  delete store.orders[fromId];
  saveStore(store);
  return true;
}

module.exports = {
  CANCELLED,
  ON_HOLD,
  cancelOrder,
  holdOrder,
  resumeOrder,
  attach,
  listCancelled,
  listOnHold,
  isSheetHidden,
  parkEvents,
  renameOrder,
  isCancelled: shop.isCancelled,
  isOnHold: shop.isOnHold,
  isParked: shop.isParked
};
