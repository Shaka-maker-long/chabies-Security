"use strict";

const fs = require("fs");
const path = require("path");
const { dataDir } = require("./workbook-store");
const { FACTORY, formatWhen } = require("./delivery-pod");

const LIVE_MS = 90 * 1000;
const KEEP_MS = 24 * 60 * 60 * 1000;

function storePath() {
  return path.join(dataDir(), "delivery-driver-locations.json");
}

function loadStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(storePath(), "utf8"));
    if (parsed && parsed.drivers && typeof parsed.drivers === "object") return parsed;
  } catch (e) {
    if (e && e.code !== "ENOENT") console.error("[delivery-track] read", e.message || e);
  }
  return { drivers: {} };
}

function saveStore(store) {
  const file = storePath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(store));
  fs.renameSync(tmp, file);
  return store;
}

function nameKey(name) {
  return String(name || "").trim().toLowerCase();
}

function ageMs(iso) {
  const t = Date.parse(iso || "");
  if (!Number.isFinite(t)) return Infinity;
  return Date.now() - t;
}

function publicRow(row) {
  if (!row || !row.name) return null;
  const age = ageMs(row.updated_at);
  const live = !!row.sharing && age <= LIVE_MS;
  return {
    name: row.name,
    lat: row.lat,
    lng: row.lng,
    accuracy: row.accuracy == null ? null : Number(row.accuracy),
    heading: row.heading == null ? null : Number(row.heading),
    speed: row.speed == null ? null : Number(row.speed),
    updated_at: row.updated_at || "",
    updated_label: formatWhen(row.updated_at),
    sharing: !!row.sharing,
    live: live,
    stale: !live,
    age_seconds: Number.isFinite(age) ? Math.max(0, Math.round(age / 1000)) : null,
    device_id: row.device_id || "",
    device_label: row.device_label || "",
    pinned_phone: !!row.pinned_phone,
    factory: {
      lat: FACTORY.lat,
      lng: FACTORY.lng,
      label: FACTORY.label,
      address: FACTORY.address
    }
  };
}

function prune(store) {
  Object.keys(store.drivers || {}).forEach((key) => {
    const row = store.drivers[key];
    if (!row || ageMs(row.updated_at) > KEEP_MS) delete store.drivers[key];
  });
  return store;
}

function resolveActor(actorName, body, req) {
  const trusted = require("./trusted-devices");
  const headerId = req && req.headers ? req.headers["x-sd-device-id"] : "";
  const deviceId = trusted.normalizeDeviceId((body && body.deviceId) || headerId);
  const phone = deviceId ? trusted.driverPhoneByDeviceId(deviceId) : null;
  const sessionName = String(actorName || "").trim();
  if (phone) {
    const owner = phone.assignedTo || (phone.assignedUsers && phone.assignedUsers[0]) || "";
    return {
      name: owner || sessionName,
      deviceId: phone.id,
      deviceLabel: phone.displayLabel || phone.nickname || phone.label || "Driver phone",
      pinned: true
    };
  }
  return {
    name: sessionName,
    deviceId: deviceId || "",
    deviceLabel: "",
    pinned: false
  };
}

function saveLocation(actorName, body, req) {
  const resolved = resolveActor(actorName, body, req);
  const name = String(resolved.name || "").trim();
  if (!name) throw new Error("Driver name is required.");
  const sharing = !(body && (body.sharing === false || body.sharing === "no"));
  const store = prune(loadStore());
  const key = nameKey(name);
  const prev = store.drivers[key] || { name: name };
  if (!resolved.pinned) {
    const trusted = require("./trusted-devices");
    const mine = trusted.driverPhoneForUser(name);
    if (mine) {
      if (prev && prev.lat != null) return publicRow(prev);
      throw new Error("Share location from the pinned driver phone (" +
        (mine.displayLabel || mine.nickname || mine.label || "Users → Devices") + ").");
    }
  }
  if (!sharing) {
    store.drivers[key] = Object.assign({}, prev, {
      name: name,
      sharing: false,
      device_id: resolved.deviceId || prev.device_id || "",
      device_label: resolved.deviceLabel || prev.device_label || "",
      pinned_phone: resolved.pinned || !!prev.pinned_phone,
      updated_at: new Date().toISOString()
    });
    saveStore(store);
    return publicRow(store.drivers[key]);
  }
  const lat = Number(body && body.lat);
  const lng = Number(body && body.lng);
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) throw new Error("Location is required.");
  if (lat < -90 || lat > 90 || lng < -180 || lng > 180) throw new Error("Location looks invalid.");
  const accuracy = body && body.accuracy != null && body.accuracy !== "" ? Number(body.accuracy) : null;
  const heading = body && body.heading != null && body.heading !== "" ? Number(body.heading) : null;
  const speed = body && body.speed != null && body.speed !== "" ? Number(body.speed) : null;
  store.drivers[key] = {
    name: name,
    lat: lat,
    lng: lng,
    accuracy: Number.isFinite(accuracy) ? accuracy : null,
    heading: Number.isFinite(heading) ? heading : null,
    speed: Number.isFinite(speed) ? speed : null,
    sharing: true,
    device_id: resolved.deviceId || "",
    device_label: resolved.deviceLabel || "",
    pinned_phone: !!resolved.pinned,
    updated_at: new Date().toISOString()
  };
  saveStore(store);
  return publicRow(store.drivers[key]);
}

function listLocations() {
  const store = prune(loadStore());
  saveStore(store);
  return Object.keys(store.drivers)
    .map((key) => publicRow(store.drivers[key]))
    .filter(Boolean)
    .sort((a, b) => {
      if (a.live !== b.live) return a.live ? -1 : 1;
      return String(b.updated_at || "").localeCompare(String(a.updated_at || ""));
    });
}

function getLocation(name) {
  const store = loadStore();
  return publicRow(store.drivers[nameKey(name)]);
}

function trackerStatus(profile, req) {
  const trusted = require("./trusted-devices");
  const delivery = require("./delivery-pod");
  const headerId = req && req.headers ? req.headers["x-sd-device-id"] : "";
  const deviceId = trusted.normalizeDeviceId(headerId);
  const phone = deviceId ? trusted.driverPhoneByDeviceId(deviceId) : null;
  const mine = profile && profile.name ? trusted.driverPhoneForUser(profile.name) : null;
  const isDriver = !!(profile && delivery.isDriverProfile(profile));
  const canPost = !!(profile && delivery.canSubmitPod(profile));
  const onPinnedPhone = !!(phone && profile && phone.assignedUsers
    && phone.assignedUsers.some((n) => nameKey(n) === nameKey(profile.name)));
  const deviceIsDriverPhone = !!phone;
  return {
    isDriver: isDriver,
    deviceId: deviceId || "",
    driverPhone: deviceIsDriverPhone,
    onPinnedPhone: onPinnedPhone,
    pinnedForMe: !!mine,
    deviceLabel: deviceIsDriverPhone
      ? (phone.displayLabel || phone.nickname || phone.label || "Driver phone")
      : (mine ? (mine.displayLabel || mine.nickname || mine.label || "") : ""),
    shouldShare: !!(canPost && (deviceIsDriverPhone || (isDriver && !mine))),
    alwaysShare: !!(canPost && deviceIsDriverPhone)
  };
}

module.exports = {
  LIVE_MS,
  saveLocation,
  listLocations,
  getLocation,
  trackerStatus,
  publicRow,
  FACTORY
};
