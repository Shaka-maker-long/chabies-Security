"use strict";

const WINDOW_MS = 15 * 60 * 1000;
const MAX_FAILURES = 5;
const buckets = new Map();

function keyFor(name, ip) {
  return String(name || "").trim().toLowerCase() + "|" + String(ip || "").trim();
}

function prune(now) {
  buckets.forEach((row, key) => {
    if (!row || now - row.firstAt > WINDOW_MS) buckets.delete(key);
  });
}

function check(name, ip) {
  const now = Date.now();
  prune(now);
  const key = keyFor(name, ip);
  const row = buckets.get(key);
  if (!row) return { ok: true, remaining: MAX_FAILURES };
  if (row.lockedUntil && now < row.lockedUntil) {
    const mins = Math.max(1, Math.ceil((row.lockedUntil - now) / 60000));
    return {
      ok: false,
      remaining: 0,
      retryAfterSec: Math.ceil((row.lockedUntil - now) / 1000),
      error: "Too many failed logins. Try again in about " + mins + " minute" + (mins === 1 ? "" : "s") + "."
    };
  }
  if (row.lockedUntil && now >= row.lockedUntil) {
    buckets.delete(key);
    return { ok: true, remaining: MAX_FAILURES };
  }
  return { ok: true, remaining: Math.max(0, MAX_FAILURES - (row.failures || 0)) };
}

function fail(name, ip) {
  const now = Date.now();
  prune(now);
  const key = keyFor(name, ip);
  let row = buckets.get(key);
  if (!row || now - row.firstAt > WINDOW_MS) {
    row = { failures: 0, firstAt: now, lockedUntil: 0 };
  }
  row.failures += 1;
  if (row.failures >= MAX_FAILURES) {
    row.lockedUntil = now + WINDOW_MS;
  }
  buckets.set(key, row);
  return check(name, ip);
}

function clear(name, ip) {
  buckets.delete(keyFor(name, ip));
}

function resetAll() {
  buckets.clear();
}

module.exports = {
  check,
  fail,
  clear,
  resetAll,
  MAX_FAILURES,
  WINDOW_MS
};
