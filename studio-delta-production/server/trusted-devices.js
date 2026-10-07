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
    devices.forEach(ensureRowShape);
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

function nameKey(name) {
  return String(name || "").replace(/\s+/g, " ").trim().toLowerCase();
}

function cleanName(name) {
  return String(name || "").replace(/\s+/g, " ").trim();
}

function uniqueNames(list) {
  const out = [];
  const seen = {};
  (list || []).forEach((name) => {
    const n = cleanName(name);
    const k = nameKey(n);
    if (!k || seen[k]) return;
    seen[k] = true;
    out.push(n);
  });
  return out;
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

function ensureRowShape(row) {
  if (!row) return row;
  row.nickname = cleanName(row.nickname);
  row.assignedUsers = uniqueNames(
    Array.isArray(row.assignedUsers) ? row.assignedUsers
      : (row.assignedTo ? [row.assignedTo] : [])
  );
  if (!row.assignedUsers.length) {
    const fallback = cleanName(row.assignedTo || row.requestedBy || row.lastUser || "");
    if (fallback) row.assignedUsers = [fallback];
  }
  row.assignedTo = row.assignedUsers[0] || "";
  row.driverPhone = !!row.driverPhone;
  const rawPending = Array.isArray(row.pendingUsers) ? row.pendingUsers : [];
  const mapped = [];
  const seen = {};
  rawPending.forEach((item) => {
    const name = cleanName(item && item.name != null ? item.name : item);
    const k = nameKey(name);
    if (!k || seen[k]) return;
    seen[k] = true;
    mapped.push({ name, at: (item && item.at) || "" });
  });
  row.pendingUsers = mapped;
  return row;
}

function displayLabel(row) {
  const nick = cleanName(row && row.nickname);
  const auto = (row && row.label) || "Unknown device";
  return nick || auto;
}

function publicDevice(row) {
  if (!row) return null;
  ensureRowShape(row);
  return {
    id: row.id,
    status: row.status,
    label: row.label || "Unknown device",
    nickname: row.nickname || "",
    displayLabel: displayLabel(row),
    userAgent: row.userAgent || "",
    lastIp: row.lastIp || "",
    assignedTo: row.assignedTo || "",
    assignedUsers: row.assignedUsers.slice(),
    pendingUsers: (row.pendingUsers || []).map((item) => ({
      name: item.name,
      at: item.at || ""
    })),
    requestedBy: row.requestedBy || "",
    lastUser: row.lastUser || "",
    createdAt: row.createdAt || "",
    updatedAt: row.updatedAt || "",
    approvedAt: row.approvedAt || "",
    approvedBy: row.approvedBy || "",
    lastSeenAt: row.lastSeenAt || "",
    revokedAt: row.revokedAt || "",
    revokedBy: row.revokedBy || "",
    driverPhone: !!row.driverPhone
  };
}

function findDevice(store, deviceId) {
  const want = normalizeDeviceId(deviceId);
  if (!want) return null;
  const row = (store.devices || []).find((item) => item.id === want) || null;
  if (row) ensureRowShape(row);
  return row;
}

function approvedCount(store) {
  return (store.devices || []).filter((row) => row.status === "approved").length;
}

function isAssignedTo(row, userName) {
  ensureRowShape(row);
  const want = nameKey(userName);
  if (!want) return false;
  return (row.assignedUsers || []).some((name) => nameKey(name) === want);
}

function addPendingUser(row, userName) {
  ensureRowShape(row);
  const name = cleanName(userName);
  if (!name) return row;
  if (isAssignedTo(row, name)) return row;
  if (!(row.pendingUsers || []).some((item) => nameKey(item.name) === nameKey(name))) {
    row.pendingUsers.push({ name, at: nowIso() });
  }
  row.requestedBy = name;
  return row;
}

function touchDevice(row, meta) {
  const now = nowIso();
  row.updatedAt = now;
  row.lastSeenAt = now;
  if (meta && meta.userName) row.lastUser = cleanName(meta.userName);
  if (meta && meta.ip) row.lastIp = String(meta.ip || "").trim();
  if (meta && meta.userAgent) {
    row.userAgent = String(meta.userAgent || "").slice(0, 400);
    if (!row.label || row.label === "Unknown device") row.label = labelFromUserAgent(row.userAgent);
  }
  if (meta && meta.nickname) row.nickname = cleanName(meta.nickname);
  return row;
}

function assignUsers(row, names) {
  row.assignedUsers = uniqueNames(names);
  row.assignedTo = row.assignedUsers[0] || "";
  row.pendingUsers = (row.pendingUsers || []).filter(
    (item) => !row.assignedUsers.some((name) => nameKey(name) === nameKey(item.name))
  );
  return row;
}

function upsertPending(store, deviceId, meta) {
  const id = normalizeDeviceId(deviceId);
  if (!id) throw new Error("This browser did not send a device id. Refresh and try again.");
  let row = findDevice(store, id);
  const now = nowIso();
  const userName = cleanName(meta && meta.userName);
  if (!row) {
    row = {
      id,
      status: "pending",
      label: labelFromUserAgent(meta && meta.userAgent),
      nickname: cleanName(meta && meta.nickname),
      userAgent: String((meta && meta.userAgent) || "").slice(0, 400),
      lastIp: String((meta && meta.ip) || "").trim(),
      assignedTo: "",
      assignedUsers: [],
      pendingUsers: userName ? [{ name: userName, at: now }] : [],
      requestedBy: userName,
      lastUser: userName,
      createdAt: now,
      updatedAt: now,
      approvedAt: "",
      approvedBy: "",
      lastSeenAt: now,
      revokedAt: "",
      revokedBy: ""
    };
    store.devices.push(row);
  } else if (row.status === "approved") {
    touchDevice(row, meta);
    addPendingUser(row, userName);
  } else {
    row.status = "pending";
    touchDevice(row, meta);
    addPendingUser(row, userName);
    row.revokedAt = "";
    row.revokedBy = "";
    row.approvedAt = "";
    row.approvedBy = "";
  }
  saveStore(store);
  return row;
}

function approveNow(store, deviceId, meta, actorName) {
  const id = normalizeDeviceId(deviceId) || ("d_" + crypto.randomBytes(12).toString("hex"));
  let row = findDevice(store, id);
  const now = nowIso();
  const userName = cleanName(meta && meta.userName) || "system";
  const by = cleanName(actorName) || userName;
  if (!row) {
    row = {
      id,
      status: "approved",
      label: labelFromUserAgent(meta && meta.userAgent),
      nickname: cleanName(meta && meta.nickname),
      userAgent: String((meta && meta.userAgent) || "").slice(0, 400),
      lastIp: String((meta && meta.ip) || "").trim(),
      assignedTo: userName,
      assignedUsers: [userName],
      pendingUsers: [],
      requestedBy: userName,
      lastUser: userName,
      createdAt: now,
      updatedAt: now,
      approvedAt: now,
      approvedBy: by,
      lastSeenAt: now,
      revokedAt: "",
      revokedBy: ""
    };
    store.devices.push(row);
  } else {
    row.status = "approved";
    row.approvedAt = now;
    row.approvedBy = by;
    row.revokedAt = "";
    row.revokedBy = "";
    assignUsers(row, [userName].concat(row.assignedUsers || []));
    touchDevice(row, meta);
    row.pendingUsers = [];
  }
  saveStore(store);
  return row;
}

function fallbackDeviceId(meta) {
  const raw = [
    String((meta && meta.userAgent) || "").trim(),
    String((meta && meta.ip) || "").trim(),
    String((meta && meta.userName) || "").trim().toLowerCase()
  ].join("|");
  if (!raw.replace(/\|/g, "").trim()) return "";
  return "d_" + crypto.createHash("sha256").update(raw).digest("hex").slice(0, 24);
}

function deviceEnforceDisabled() {
  return String(process.env.SD_TRUST_DEVICES || "").trim() === "0";
}

function bootstrapCodeConfigured() {
  return String(process.env.DEVICE_BOOTSTRAP_CODE || "").trim();
}

function bootstrapCodeMatches(raw) {
  const want = bootstrapCodeConfigured();
  if (!want) return false;
  const got = String(raw || "").trim();
  if (!got || got.length !== want.length) return false;
  try {
    return crypto.timingSafeEqual(Buffer.from(got), Buffer.from(want));
  } catch (e) {
    return false;
  }
}

function userHasApprovedDevice(store, userName) {
  const want = nameKey(userName);
  if (!want) return false;
  return (store.devices || []).some((row) => row.status === "approved" && isAssignedTo(row, userName));
}

/**
 * After credentials are verified:
 * - device must be approved AND assigned to this person (no first-login auto-approve)
 * - Manager may unlock THEIR first device with DEVICE_BOOTSTRAP_CODE (even if other
 *   people already have approved phones — avoids locking the Manager out)
 * - everyone else waits under Users → Devices until the Manager assigns the phone/computer to them
 */
function assertLoginAllowed(meta) {
  if (deviceEnforceDisabled()) {
    return { ok: true, skipped: true, device: null };
  }

  const store = loadStore();
  const deviceId = normalizeDeviceId(meta && meta.deviceId) || fallbackDeviceId(meta);
  const userName = cleanName(meta && meta.userName);
  const info = {
    userName,
    userAgent: meta && meta.userAgent,
    ip: meta && meta.ip,
    nickname: meta && meta.nickname
  };

  if (!deviceId) {
    return {
      ok: false,
      pending: true,
      error: "This device is not recognised. Refresh the page, then ask the Manager to approve it."
    };
  }

  const existing = findDevice(store, deviceId);
  if (existing && existing.status === "approved" && isAssignedTo(existing, userName)) {
    touchDevice(existing, info);
    saveStore(store);
    return { ok: true, device: publicDevice(existing) };
  }

  // Manager unlock: matching DEVICE_BOOTSTRAP_CODE + this Manager has no approved device yet.
  const wantsBootstrap = bootstrapCodeMatches(meta && meta.bootstrapCode);
  const managerUnlock = !!(meta && meta.canManageUsers) && wantsBootstrap && !userHasApprovedDevice(store, userName);
  if (managerUnlock) {
    const row = approveNow(store, deviceId, info, userName + " (bootstrap)");
    return { ok: true, bootstrapped: true, device: publicDevice(row) };
  }

  const pending = upsertPending(store, deviceId, info);
  const assigned = (pending.assignedUsers || []).join(", ");
  const forWho = userName || "this person";
  const managerNeedsUnlock = !!(meta && meta.canManageUsers) && !userHasApprovedDevice(store, userName);
  let error;
  if (managerNeedsUnlock) {
    if (!bootstrapCodeConfigured()) {
      error = "This device is not approved for " + forWho + " yet. Ask the Manager to approve it.";
    } else if (wantsBootstrap) {
      error = "That unlock code is wrong, or this device still needs Manager approval.";
    } else {
      error = "This device is not approved for " + forWho + " yet. Ask the Manager to approve it.";
    }
  } else if (pending.status === "approved" && assigned) {
    error = "This device is assigned to " + assigned + ". " + forWho +
      " needs Manager approval before logging in here.";
  } else {
    error = forWho + " is not linked to this device yet. Ask the Manager to approve it.";
  }
  return {
    ok: false,
    pending: true,
    needsBootstrap: managerNeedsUnlock,
    bootstrapConfigured: !!bootstrapCodeConfigured(),
    device: publicDevice(pending),
    error
  };
}

function sessionStillValid(userName, deviceId) {
  if (deviceEnforceDisabled()) return { ok: true, skipped: true };
  const id = normalizeDeviceId(deviceId);
  if (!id) {
    return { ok: false, error: "This session is not bound to an approved device. Log in again." };
  }
  const store = loadStore();
  const row = findDevice(store, id);
  if (!row || row.status !== "approved" || !isAssignedTo(row, userName)) {
    return { ok: false, error: "This device is no longer approved for " + (cleanName(userName) || "you") + ". Log in again after the Manager approves it." };
  }
  return { ok: true, device: publicDevice(row) };
}

function snapshot() {
  const store = loadStore();
  const devices = (store.devices || [])
    .slice()
    .map((row) => {
      ensureRowShape(row);
      return publicDevice(row);
    })
    .sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")));
  const pending = [];
  devices.forEach((row) => {
    if (row.status === "pending") pending.push(row);
    else if (row.status === "approved" && (row.pendingUsers || []).length) pending.push(row);
  });
  return {
    devices,
    approved: devices.filter((row) => row.status === "approved"),
    pending,
    revoked: devices.filter((row) => row.status === "revoked"),
    approvedCount: devices.filter((row) => row.status === "approved").length,
    pendingCount: pending.length,
    bootstrapConfigured: !!bootstrapCodeConfigured()
  };
}

function approveDevice(deviceId, actorName, body) {
  const store = loadStore();
  const row = findDevice(store, deviceId);
  if (!row) throw new Error("Device not found.");
  const now = nowIso();
  const assignTo = cleanName(
    (body && (body.assignedTo || body.user || body.name))
      || (row.pendingUsers[0] && row.pendingUsers[0].name)
      || row.requestedBy
      || row.lastUser
  );
  if (!assignTo) throw new Error("Pick who this device belongs to.");
  if (body && Object.prototype.hasOwnProperty.call(body, "nickname")) {
    row.nickname = cleanName(body.nickname);
  }
  const keepExisting = !(body && body.replaceAssignees);
  const nextAssignees = keepExisting
    ? uniqueNames((row.assignedUsers || []).concat([assignTo]))
    : [assignTo];
  assignUsers(row, nextAssignees);
  row.pendingUsers = (row.pendingUsers || []).filter((item) => nameKey(item.name) !== nameKey(assignTo));
  row.status = "approved";
  row.approvedAt = now;
  row.approvedBy = cleanName(actorName) || "Manager";
  row.updatedAt = now;
  row.revokedAt = "";
  row.revokedBy = "";
  row.requestedBy = assignTo;
  saveStore(store);
  return snapshot();
}

function clearDriverPhoneForUser(store, userName, exceptId) {
  const want = nameKey(userName);
  if (!want) return;
  (store.devices || []).forEach((item) => {
    if (!item || item.id === exceptId) return;
    ensureRowShape(item);
    if (!item.driverPhone) return;
    if ((item.assignedUsers || []).some((name) => nameKey(name) === want)) {
      item.driverPhone = false;
    }
  });
}

function setDriverPhone(deviceId, on) {
  const store = loadStore();
  const row = findDevice(store, deviceId);
  if (!row) throw new Error("Device not found.");
  if (row.status !== "approved") throw new Error("Only an approved device can be the driver phone.");
  ensureRowShape(row);
  const owner = row.assignedTo || (row.assignedUsers && row.assignedUsers[0]) || "";
  if (!owner) throw new Error("Assign this device to the driver first.");
  if (on) {
    clearDriverPhoneForUser(store, owner, row.id);
    row.driverPhone = true;
  } else {
    row.driverPhone = false;
  }
  row.updatedAt = nowIso();
  saveStore(store);
  return snapshot();
}

function driverPhoneForUser(userName) {
  const store = loadStore();
  const want = nameKey(userName);
  if (!want) return null;
  const row = (store.devices || []).find((item) => {
    if (!item || item.status !== "approved" || !item.driverPhone) return false;
    ensureRowShape(item);
    return (item.assignedUsers || []).some((name) => nameKey(name) === want);
  });
  return row ? publicDevice(row) : null;
}

function driverPhoneByDeviceId(deviceId) {
  const store = loadStore();
  const row = findDevice(store, deviceId);
  if (!row || row.status !== "approved" || !row.driverPhone) return null;
  return publicDevice(row);
}

function updateDevice(deviceId, body) {
  const store = loadStore();
  const row = findDevice(store, deviceId);
  if (!row) throw new Error("Device not found.");
  if (body && Object.prototype.hasOwnProperty.call(body, "nickname")) {
    row.nickname = cleanName(body.nickname);
  }
  if (body && (body.assignedTo || body.assignedUsers)) {
    const list = Array.isArray(body.assignedUsers)
      ? body.assignedUsers
      : [body.assignedTo];
    if (body.replaceAssignees) assignUsers(row, list);
    else assignUsers(row, (row.assignedUsers || []).concat(list));
  }
  if (body && Object.prototype.hasOwnProperty.call(body, "driverPhone")) {
    const on = body.driverPhone === true || body.driverPhone === "yes" || body.driverPhone === 1;
    if (on) {
      if (row.status !== "approved") throw new Error("Only an approved device can be the driver phone.");
      const owner = row.assignedTo || (row.assignedUsers && row.assignedUsers[0]) || "";
      if (!owner) throw new Error("Assign this device to the driver first.");
      clearDriverPhoneForUser(store, owner, row.id);
      row.driverPhone = true;
    } else {
      row.driverPhone = false;
    }
  }
  row.updatedAt = nowIso();
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
  row.revokedBy = cleanName(actorName) || "Manager";
  row.updatedAt = now;
  row.pendingUsers = [];
  row.driverPhone = false;
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
  const body = (req && req.body) || {};
  return {
    deviceId: normalizeDeviceId(bodyDeviceId || headerId || body.deviceId),
    userAgent: String((req && req.headers && req.headers["user-agent"]) || "").slice(0, 400),
    ip: clientIp(req),
    userName: cleanName(userName),
    nickname: cleanName(body.deviceNickname || body.nickname),
    bootstrapCode: body.bootstrapCode || body.deviceBootstrapCode || "",
    canManageUsers: false
  };
}

module.exports = {
  normalizeDeviceId,
  labelFromUserAgent,
  assertLoginAllowed,
  sessionStillValid,
  snapshot,
  approveDevice,
  updateDevice,
  setDriverPhone,
  driverPhoneForUser,
  driverPhoneByDeviceId,
  revokeDevice,
  removeDevice,
  metaFromReq,
  isAssignedTo,
  bootstrapCodeConfigured,
  approvedCount: () => approvedCount(loadStore())
};
