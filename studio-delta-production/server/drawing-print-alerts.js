"use strict";

const fs = require("fs");
const path = require("path");
const { dataDir } = require("./workbook-store");
const { formatOrderId } = require("./db");
const staff = require("./staff");

function alertsPath() {
  return path.join(dataDir(), "drawing-print-alerts.json");
}

function loadAll() {
  try {
    const raw = JSON.parse(fs.readFileSync(alertsPath(), "utf8"));
    return Array.isArray(raw) ? raw : [];
  } catch (e) {
    return [];
  }
}

function saveAll(rows) {
  const tmp = alertsPath() + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(rows || []));
  fs.renameSync(tmp, alertsPath());
}

function nextId(rows) {
  const max = (rows || []).reduce((m, row) => {
    const n = Number(String((row && row.id) || "").replace(/\D/g, "")) || 0;
    return Math.max(m, n);
  }, 0);
  return "dpa" + (max + 1);
}

function pending() {
  return loadAll().filter((row) => row && row.status === "pending");
}

function upsertPending(alert) {
  const orderNumber = formatOrderId(alert && alert.order_number);
  if (!orderNumber) return null;
  const rows = loadAll();
  const existing = rows.find((row) => formatOrderId(row.order_number) === orderNumber && row.status === "pending");
  const rec = existing || {
    id: nextId(rows),
    order_number: orderNumber,
    status: "pending",
    seen_at: "",
    seen_by: ""
  };
  rec.client_name = String((alert && alert.client_name) || rec.client_name || "").trim();
  rec.product = String((alert && alert.product) || rec.product || "").trim();
  rec.drawing_url = String((alert && alert.drawing_url) || rec.drawing_url || "").trim();
  rec.uploaded_at = String((alert && alert.uploaded_at) || rec.uploaded_at || new Date().toISOString());
  rec.uploaded_by = String((alert && alert.uploaded_by) || rec.uploaded_by || "").trim();
  rec.status = "pending";
  rec.seen_at = "";
  rec.seen_by = "";
  if (!existing) rows.push(rec);
  saveAll(rows);
  return rec;
}

function pollFor(name) {
  if (!staff.isDrawingPrintOwnerName(name)) return { ok: true, alerts: [] };
  return { ok: true, alerts: pending() };
}

function acknowledge(id, name) {
  if (!staff.isDrawingPrintOwnerName(name)) {
    throw new Error("Only Siya can confirm he has seen a new drawing");
  }
  const want = String(id || "").trim();
  if (!want) throw new Error("Choose which drawing update you have seen");
  const rows = loadAll();
  const rec = rows.find((row) => row && String(row.id) === want);
  if (!rec) throw new Error("That drawing update is no longer waiting");
  if (rec.status !== "pending") return rec;
  rec.status = "seen";
  rec.seen_at = new Date().toISOString();
  rec.seen_by = String(name || "").trim();
  saveAll(rows);
  return rec;
}

module.exports = {
  pending,
  pollFor,
  acknowledge,
  upsertPending
};
