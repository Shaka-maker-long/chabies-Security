"use strict";

const crypto = require("crypto");

const PREFIX = "scrypt$";
const MIN_LEN = 4;
const MAX_LEN = 72;
const SCRYPT_N = 16384;
const SCRYPT_R = 8;
const SCRYPT_P = 1;
const KEYLEN = 32;

function normalizePlain(value) {
  return String(value == null ? "" : value).trim();
}

function isHashed(stored) {
  return String(stored || "").startsWith(PREFIX);
}

function hashPlain(plain) {
  const code = normalizePlain(plain);
  if (!code) throw new Error("Access code is required");
  if (code.length > MAX_LEN) {
    throw new Error("Access code is too long.");
  }
  const salt = crypto.randomBytes(16).toString("base64url");
  const derived = crypto.scryptSync(code, salt, KEYLEN, { N: SCRYPT_N, r: SCRYPT_R, p: SCRYPT_P });
  return PREFIX + salt + "$" + derived.toString("base64url");
}

function verify(plain, stored) {
  const code = normalizePlain(plain);
  const raw = String(stored == null ? "" : stored).trim();
  if (!code || !raw) return false;
  if (!isHashed(raw)) {
    return code === raw;
  }
  const parts = raw.split("$");
  // scrypt$salt$hash
  if (parts.length !== 3 || parts[0] !== "scrypt") return false;
  const salt = parts[1];
  const want = parts[2];
  let derived;
  try {
    derived = crypto.scryptSync(code, salt, KEYLEN, { N: SCRYPT_N, r: SCRYPT_R, p: SCRYPT_P });
  } catch (e) {
    return false;
  }
  const a = Buffer.from(derived.toString("base64url"));
  const b = Buffer.from(want);
  if (a.length !== b.length) return false;
  return crypto.timingSafeEqual(a, b);
}

function needsRehash(stored) {
  return !!normalizePlain(stored) && !isHashed(stored);
}

function strengthOk(plain) {
  const code = normalizePlain(plain);
  return code.length >= MIN_LEN && code.length <= MAX_LEN;
}

module.exports = {
  PREFIX,
  MIN_LEN,
  MAX_LEN,
  normalizePlain,
  isHashed,
  hashPlain,
  verify,
  needsRehash,
  strengthOk
};
