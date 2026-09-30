"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir } = require("./workbook-store");

function devicesPath() {
  return path.join(dataDir(), "trusted-devices.json");
}

function loadStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(devicesPath(), "utf8"));
    const devices = Array.isArray(parsed.devices) ? parsed.devices.filter((row) => row && row.id) : [];
    return { devices };
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[trusted-devices] could not read", devicesPath(), e.message || e);
    }
    return { devices: [] };
  }
}

function saveStore(store) {
  fs.mkdirSync(path.dirname(devicesPath()), { recursive: true });
  const tmp = devicesPath() + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify({ devices: store.devices || [] }));
  fs.renameSync(tmp, devicesPath());
  return store;
}

function nowIso() {
  return new Date().toISOString();
}

function normalizeDeviceId(raw) {
  const id = String(raw || "").trim();
  if (!id) return "";
  if (!/^[A-Za-z0-9_-]{8,80}$/.test(id)) return "";
  return id;
}

function labelFromUserAgent(ua) {
  const s = String(ua || "").trim();
  if (!s) return "Unknown device";
  let browser = "Browser";
  if (/Edg\//i.test(s)) browser = "Edge";
  else if (/Chrome\//i.test(s) && !/Chromium/i.test(s)) browser = "Chrome";
  else if (/Firefox\//i.test(s)) browser = "Firefox";
  else if (/Safari\//i.test(s) && !/Chrome\//i.test(s)) browser = "Safari";
  let os = "device";
  if (/Windows/i.test(s)) os = "Windows";
  else if (/Android/i.test(s)) os = "Android";
  else if (/iPhone|iPad|iPod/i.test(s)) os = "iPhone";
  else if (/Mac OS X|Macintosh/i.test(s)) os = "Mac";
  else if (/Linux/i.test(s)) os = "Linux";
  return browser + " on " + os;
}

function publicDevice(row) {
  if (!row) return null;
  return {
    id: row.id,
    status: row.status,
    label: row.label || "Unknown device",
    userAgent: row.userAgent || "",
    lastIp: row.lastIp || "",
    requestedBy: row.requestedBy || "",
    lastUser: row.lastUser || "",
    createdAt: row.createdAt || "",
    updatedAt: row.updatedAt || "",
    approvedAt: row.approvedAt || "",
    approvedBy: row.approvedBy || "",
    lastSeenAt: row.lastSeenAt || "",
    revokedAt: row.revokedAt || "",
    revokedBy: row.revokedBy || ""
  };
}

function findDevice(store, deviceId) {
  const want = normalizeDeviceId(deviceId);
  if (!want) return null;
  return (store.devices || []).find((row) => row.id === want) || null;
}

function approvedCount(store) {
  return (store.devices || []).filter((row) => row.status === "approved").length;
}

function touchDevice(row, meta) {
  const now = nowIso();
  row.updatedAt = now;
  row.lastSeenAt = now;
  if (meta && meta.userName) row.lastUser = String(meta.userName).trim();
  if (meta && meta.ip) row.lastIp = String(meta.ip || "").trim();
  if (meta && meta.userAgent) {
    row.userAgent = String(meta.userAgent || "").slice(0, 400);
    if (!row.label || row.label === "Unknown device") row.label = labelFromUserAgent(row.userAgent);
  }
  return row;
}

function upsertPending(store, deviceId, meta) {
  const id = normalizeDeviceId(deviceId);
  if (!id) throw new Error("This browser did not send a device id. Refresh and try again.");
  let row = findDevice(store, id);
  const now = nowIso();
  if (!row) {
    row = {
      id,
      status: "pending",
      label: labelFromUserAgent(meta && meta.userAgent),
      userAgent: String((meta && meta.userAgent) || "").slice(0, 400),
      lastIp: String((meta && meta.ip) || "").trim(),
      requestedBy: String((meta && meta.userName) || "").trim(),
      lastUser: String((meta && meta.userName) || "").trim(),
      createdAt: now,
      updatedAt: now,
      approvedAt: "",
      approvedBy: "",
      lastSeenAt: now,
      revokedAt: "",
      revokedBy: ""
    };
    store.devices.push(row);
  } else {
    row.status = "pending";
    row.requestedBy = String((meta && meta.userName) || row.requestedBy || "").trim();
    touchDevice(row, meta);
    row.revokedAt = "";
    row.revokedBy = "";
    row.approvedAt = "";
    row.approvedBy = "";
  }
  saveStore(store);
  return row;
}

function autoApprove(store, deviceId, meta) {
  const id = normalizeDeviceId(deviceId) || ("d_" + crypto.randomBytes(12).toString("hex"));
  let row = findDevice(store, id);
  const now = nowIso();
  if (!row) {
    row = {
      id,
      status: "approved",
      label: labelFromUserAgent(meta && meta.userAgent),
      userAgent: String((meta && meta.userAgent) || "").slice(0, 400),
      lastIp: String((meta && meta.ip) || "").trim(),
      requestedBy: String((meta && meta.userName) || "").trim(),
      lastUser: String((meta && meta.userName) || "").trim(),
      createdAt: now,
      updatedAt: now,
      approvedAt: now,
      approvedBy: String((meta && meta.userName) || "system").trim() || "system",
      lastSeenAt: now,
      revokedAt: "",
      revokedBy: ""
    };
    store.devices.push(row);
  } else {
    row.status = "approved";
    row.approvedAt = now;
    row.approvedBy = String((meta && meta.userName) || row.approvedBy || "system").trim() || "system";
    row.revokedAt = "";
    row.revokedBy = "";
    touchDevice(row, meta);
  }
  saveStore(store);
  return row;
}

function fallbackDeviceId(meta) {
  const raw = [String((meta && meta.userAgent) || "").trim(), String((meta && meta.ip) || "").trim()].join("|");
  if (!raw.replace(/\|/g, "").trim()) return "";
  return "d_" + crypto.createHash("sha256").update(raw).digest("hex").slice(0, 24);
}

/**
 * After credentials are verified, decide whether this device may receive a session.
 * Bootstrap: if no approved devices exist yet, auto-approve the first successful login.
 */
function assertLoginAllowed(meta) {
  const store = loadStore();
  const deviceId = normalizeDeviceId(meta && meta.deviceId) || fallbackDeviceId(meta);
  const info = {
    userName: meta && meta.userName,
    userAgent: meta && meta.userAgent,
    ip: meta && meta.ip
  };

  if (!approvedCount(store)) {
    const row = autoApprove(store, deviceId || ("d_" + crypto.randomBytes(12).toString("hex")), info);
    return { ok: true, bootstrapped: true, device: publicDevice(row) };
  }

  if (!deviceId) {
    return {
      ok: false,
      pending: true,
      error: "This device is not recognised. Refresh the page, then ask the Manager to approve it under Users → Devices."
    };
  }

  const existing = findDevice(store, deviceId);
  if (existing && existing.status === "approved") {
    touchDevice(existing, info);
    saveStore(store);
    return { ok: true, device: publicDevice(existing) };
  }

  const pending = upsertPending(store, deviceId, info);
  return {
    ok: false,
    pending: true,
    device: publicDevice(pending),
    error: "This device is waiting for approval. Ask the Manager to open Users → Devices and approve “" +
      (pending.label || "this device") + "”."
  };
}

function snapshot() {
  const store = loadStore();
  const devices = (store.devices || [])
    .slice()
    .sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")))
    .map(publicDevice);
  return {
    devices,
    approved: devices.filter((row) => row.status === "approved"),
    pending: devices.filter((row) => row.status === "pending"),
    revoked: devices.filter((row) => row.status === "revoked"),
    approvedCount: devices.filter((row) => row.status === "approved").length,
    pendingCount: devices.filter((row) => row.status === "pending").length
  };
}

function approveDevice(deviceId, actorName) {
  const store = loadStore();
  const row = findDevice(store, deviceId);
  if (!row) throw new Error("Device not found.");
  const now = nowIso();
  row.status = "approved";
  row.approvedAt = now;
  row.approvedBy = String(actorName || "").trim() || "Manager";
  row.updatedAt = now;
  row.revokedAt = "";
  row.revokedBy = "";
  saveStore(store);
  return snapshot();
}

function revokeDevice(deviceId, actorName) {
  const store = loadStore();
  const row = findDevice(store, deviceId);
  if (!row) throw new Error("Device not found.");
  const now = nowIso();
  row.status = "revoked";
  row.revokedAt = now;
  row.revokedBy = String(actorName || "").trim() || "Manager";
  row.updatedAt = now;
  saveStore(store);
  return snapshot();
}

function removeDevice(deviceId) {
  const store = loadStore();
  const want = normalizeDeviceId(deviceId);
  const before = store.devices.length;
  store.devices = store.devices.filter((row) => row.id !== want);
  if (store.devices.length === before) throw new Error("Device not found.");
  saveStore(store);
  return snapshot();
}

function clientIp(req) {
  const xf = String((req && req.headers && req.headers["x-forwarded-for"]) || "").split(",")[0].trim();
  if (xf) return xf;
  return String((req && (req.ip || (req.socket && req.socket.remoteAddress))) || "").trim();
}

function metaFromReq(req, userName, bodyDeviceId) {
  const headerId = req && req.headers ? req.headers["x-sd-device-id"] : "";
  return {
    deviceId: normalizeDeviceId(bodyDeviceId || headerId),
    userAgent: String((req && req.headers && req.headers["user-agent"]) || "").slice(0, 400),
    ip: clientIp(req),
    userName: String(userName || "").trim()
  };
}

module.exports = {
  normalizeDeviceId,
  labelFromUserAgent,
  assertLoginAllowed,
  snapshot,
  approveDevice,
  revokeDevice,
  removeDevice,
  metaFromReq,
  approvedCount: () => approvedCount(loadStore())
};
