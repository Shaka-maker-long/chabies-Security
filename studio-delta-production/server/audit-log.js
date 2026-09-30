"use strict";

const fs = require("fs");
const path = require("path");
const { dataDir } = require("./workbook-store");

const MAX_LINES_KEEP = 5000;

function auditPath() {
  return path.join(dataDir(), "audit-log.jsonl");
}

function nowIso() {
  return new Date().toISOString();
}

function clientIp(req) {
  if (!req) return "";
  const xf = String((req.headers && req.headers["x-forwarded-for"]) || "").split(",")[0].trim();
  if (xf) return xf;
  return String(req.ip || (req.socket && req.socket.remoteAddress) || "").trim();
}

function append(event) {
  try {
    const row = Object.assign({ at: nowIso() }, event || {});
    fs.mkdirSync(path.dirname(auditPath()), { recursive: true });
    fs.appendFileSync(auditPath(), JSON.stringify(row) + "\n");
  } catch (e) {
    console.warn("[audit] write failed:", e && e.message ? e.message : e);
  }
}

function record(action, meta, req) {
  const m = meta || {};
  append({
    action: String(action || "").trim() || "event",
    actor: m.actor || m.userName || "",
    target: m.target || "",
    detail: m.detail || "",
    ok: m.ok == null ? true : !!m.ok,
    deviceId: m.deviceId || "",
    ip: m.ip || clientIp(req),
    userAgent: m.userAgent || String((req && req.headers && req.headers["user-agent"]) || "").slice(0, 200)
  });
}

function recent(limit) {
  const n = Math.max(1, Math.min(Number(limit) || 200, 1000));
  try {
    const raw = fs.readFileSync(auditPath(), "utf8");
    const lines = raw.split("\n").filter(Boolean);
    const slice = lines.slice(-n);
    return slice.map((line) => {
      try { return JSON.parse(line); } catch (e) { return { raw: line }; }
    }).reverse();
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.warn("[audit] read failed:", e.message || e);
    }
    return [];
  }
}

function trimIfHuge() {
  try {
    const raw = fs.readFileSync(auditPath(), "utf8");
    const lines = raw.split("\n").filter(Boolean);
    if (lines.length <= MAX_LINES_KEEP) return;
    const keep = lines.slice(-MAX_LINES_KEEP);
    fs.writeFileSync(auditPath(), keep.join("\n") + "\n");
  } catch (e) {}
}

setInterval(trimIfHuge, 6 * 60 * 60 * 1000).unref();

module.exports = {
  append,
  record,
  recent,
  clientIp
};
