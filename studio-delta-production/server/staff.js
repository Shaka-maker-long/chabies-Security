const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const { getBook, persistWorkbook, dataDir } = require("./workbook-store");

const FLOOR_TASKS = [
  "Profile Cutting", "Plate Cutting", "Tagging", "Welding", "Grinding",
  "Quality Control", "Paint Preparation", "Painting", "Assembly"
];

const sessions = new Map();
const SESSION_MAX_AGE_MS = 90 * 24 * 60 * 60 * 1000;

function sessionsPath() {
  return path.join(dataDir(), "office-sessions.json");
}

function loadSessions() {
  try {
    const raw = JSON.parse(fs.readFileSync(sessionsPath(), "utf8"));
    const now = Date.now();
    Object.keys(raw || {}).forEach((token) => {
      const row = raw[token];
      if (!row || !row.name) return;
      if (row.savedAt && now - Number(row.savedAt) > SESSION_MAX_AGE_MS) return;
      sessions.set(token, {
        name: row.name,
        access: row.access,
        isAdmin: !!row.isAdmin,
        canSeeOffice: !!row.canSeeOffice,
        canSeeDebtors: !!row.canSeeDebtors,
        tasks: Array.isArray(row.tasks) ? row.tasks : []
      });
    });
  } catch (e) {}
}

function persistSessions() {
  try {
    const out = {};
    const now = Date.now();
    sessions.forEach((safe, token) => {
      out[token] = { ...safe, savedAt: now };
    });
    fs.writeFileSync(sessionsPath(), JSON.stringify(out));
  } catch (e) {
    console.warn("[staff] could not persist office sessions:", e.message);
  }
}

loadSessions();

function usersSheet() {
  const book = getBook();
  let sheet = book.getSheetByName("Users");
  if (!sheet) sheet = book.insertSheet("Users");
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getLastRow() < 1
    ? []
    : (sheet.getRange(1, 1, 1, lastCol).getValues()[0] || []);
  if (sheet.getLastRow() < 1 || String(headers[0] || "").trim() === "") {
    sheet.getRange(1, 1, 1, 6).setValues([["Name", "Role", "Password", "Tasks", "Access", "See Debtors"]]);
    persistWorkbook();
    return sheet;
  }
  const norm = headers.map((h) => String(h || "").trim().toLowerCase());
  if (norm.indexOf("access") === -1) sheet.getRange(1, 5).setValue("Access");
  if (norm.indexOf("see debtors") === -1) sheet.getRange(1, 6).setValue("See Debtors");
  return sheet;
}

function durationsSheet() {
  const book = getBook();
  let sheet = book.getSheetByName("Task_Durations");
  if (!sheet) {
    sheet = book.insertSheet("Task_Durations");
    sheet.getRange(1, 1, 1, 3).setValues([["Product", "Process", "Minutes"]]);
    persistWorkbook();
  }
  return sheet;
}

function parseAccess(accessCell, roleCell) {
  const a = String(accessCell || "").trim().toLowerCase();
  if (a === "admin") return "Admin";
  if (a === "production") return "Production";
  const r = String(roleCell || "").trim().toLowerCase();
  if (r === "admin") return "Admin";
  return "Production";
}

function parseSeeDebtors(body, access) {
  if (access !== "Admin") return "No";
  if (body.seeDebtors === false || body.canSeeDebtors === false) return "No";
  const v = String(body.seeDebtors != null && body.seeDebtors !== "" ? body.seeDebtors : (body.canSeeDebtors != null ? body.canSeeDebtors : "Yes"))
    .trim()
    .toLowerCase();
  if (v === "no" || v === "false" || v === "0") return "No";
  return "Yes";
}

function countdownRemainingMs(order, nowMs) {
  const target = Number(order && order.targetMinutes) || 0;
  if (target <= 0) return null;
  let start = order.startedAt;
  if (start instanceof Date) start = start.getTime();
  else if (typeof start === "string" && start) start = new Date(start).getTime();
  else start = Number(start) || 0;
  if (!start) return null;
  const pauseMs = Number(order.pauseMs) || 0;
  let pausedAt = order.pausedAt;
  if (pausedAt instanceof Date) pausedAt = pausedAt.getTime();
  else if (typeof pausedAt === "string" && pausedAt) pausedAt = new Date(pausedAt).getTime();
  else pausedAt = Number(pausedAt) || 0;
  const end = order.isPaused && pausedAt ? pausedAt : nowMs;
  return target * 60 * 1000 - Math.max(0, end - start - pauseMs);
}

function bumpShopCache() {
  try { require("./gas").clearShopCache(); } catch (e) {}
}

function parseTasks(tasksCell) {
  return String(tasksCell || "")
    .split(/[,/&+|]+/)
    .map((s) => s.trim())
    .filter((s) => FLOOR_TASKS.indexOf(s) !== -1);
}

function rowToUser(row, id) {
  const access = parseAccess(row[4], row[1]);
  const isAdmin = access === "Admin";
  const debtors = String(row[5] || "").trim().toLowerCase();
  return {
    id,
    name: String(row[0] || "").trim(),
    role: String(row[1] || "").trim(),
    tasks: parseTasks(row[3]),
    access,
    isAdmin,
    canSeeOffice: isAdmin,
    canSeeDebtors: isAdmin && debtors !== "no",
    seeDebtors: isAdmin && debtors !== "no" ? "Yes" : "No"
  };
}

function seedLocalAdminIfEmpty() {
  if (listUsers().length) return false;
  upsertUser({
    name: "Admin",
    access: "Admin",
    role: "Admin",
    password: process.env.LOCAL_ADMIN_CODE || "admin",
    seeDebtors: "Yes"
  });
  console.log("[staff] no Users yet — seeded local Admin (access code: admin)");
  return true;
}

function listUsers() {
  const sheet = usersSheet();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const grid = sheet.getRange(2, 1, last - 1, 6).getValues();
  const out = [];
  for (let i = 0; i < grid.length; i++) {
    const u = rowToUser(grid[i], i + 2);
    if (u.name) out.push(u);
  }
  return out;
}

function findUserRow(name) {
  const want = String(name || "").trim().toLowerCase();
  const sheet = usersSheet();
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const values = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim().toLowerCase() === want) return i + 2;
  }
  return 0;
}

function upsertUser(body) {
  const name = String(body.name || "").trim();
  if (!name) throw new Error("Name is required");
  const access = parseAccess(body.access, body.role);
  const role = String(body.role || "").trim() || (access === "Admin" ? "Admin" : "Production");
  const tasks = Array.isArray(body.tasks) ? body.tasks.filter((t) => FLOOR_TASKS.indexOf(t) !== -1) : parseTasks(body.tasks);
  const seeDebtors = parseSeeDebtors(body, access);
  const sheet = usersSheet();
  let rowNum = findUserRow(name);
  let password = String(body.password || "").trim();
  if (!rowNum) {
    if (!password) throw new Error("Access code is required for a new user");
    rowNum = sheet.getLastRow() + 1;
  } else if (!password) {
    password = String(sheet.getRange(rowNum, 3).getValue() || "");
  }
  sheet.getRange(rowNum, 1, 1, 6).setValues([[
    name, role, password, tasks.join(", "), access, seeDebtors
  ]]);
  persistWorkbook();
  bumpShopCache();
  return rowToUser([name, role, password, tasks.join(", "), access, seeDebtors], rowNum);
}

function deleteUser(name) {
  const rowNum = findUserRow(name);
  if (rowNum) {
    usersSheet().deleteRow(rowNum);
    persistWorkbook();
    bumpShopCache();
  }
}

function loginFailureMessage() {
  const users = listUsers();
  if (!users.length) {
    return "No Users on Railway yet. Wait a few seconds for names to copy from the old spreadsheet, then try again.";
  }
  if (users.length === 1 && String(users[0].name).toLowerCase() === "admin") {
    return "Incorrect name or access code. This boot only has Admin — try Admin / admin, or wait a few seconds for the old Users sheet to copy.";
  }
  return "Incorrect name or access code";
}

function verifyUser(name, password) {
  const sheet = usersSheet();
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const grid = sheet.getRange(2, 1, last - 1, 6).getValues();
  let adminPassword = "";
  grid.forEach((row) => {
    if (parseAccess(row[4], row[1]) === "Admin" && !adminPassword) adminPassword = String(row[2] || "");
  });
  const want = String(name || "").trim().toLowerCase();
  const pass = String(password || "");
  for (let i = 0; i < grid.length; i++) {
    if (String(grid[i][0] || "").trim().toLowerCase() !== want) continue;
    const rowPass = String(grid[i][2] || "");
    if (pass === rowPass || (adminPassword && pass === adminPassword)) {
      return rowToUser(grid[i], i + 2);
    }
  }
  return null;
}

function createSession(profile) {
  const token = crypto.randomBytes(16).toString("hex");
  const safe = {
    name: profile.name,
    access: profile.access,
    isAdmin: profile.isAdmin,
    canSeeOffice: profile.canSeeOffice,
    canSeeDebtors: profile.canSeeDebtors,
    tasks: profile.tasks
  };
  sessions.set(token, safe);
  persistSessions();
  return { token, ...safe };
}

function readSession(req) {
  const token = String((req.headers && (req.headers["x-sd-token"] || req.headers["authorization"])) || "")
    .replace(/^Bearer\s+/i, "")
    .trim();
  if (!token) return null;
  return sessions.get(token) || null;
}

function listDurations() {
  const sheet = durationsSheet();
  const last = sheet.getLastRow();
  const rows = [];
  if (last >= 2) {
    const grid = sheet.getRange(2, 1, last - 1, 3).getValues();
    grid.forEach((r) => {
      const product = String(r[0] || "").trim();
      const process = String(r[1] || "").trim();
      const minutes = Number(r[2]) || 0;
      if (product && process) rows.push({ product, process, minutes });
    });
  }
  return rows;
}

function setDurations(rows) {
  const sheet = durationsSheet();
  const last = sheet.getLastRow();
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, 3).clearContent();
  }
  const clean = (rows || []).filter((r) => r && r.product && r.process && Number(r.minutes) > 0)
    .map((r) => [String(r.product).trim(), String(r.process).trim(), Number(r.minutes)]);
  if (clean.length) sheet.getRange(2, 1, clean.length, 3).setValues(clean);
  persistWorkbook();
  bumpShopCache();
  return listDurations();
}

function durationMinutes(product, process) {
  const p = String(product || "").trim().toLowerCase();
  const t = String(process || "").trim().toLowerCase();
  const rows = listDurations();
  const hit = rows.find((r) => r.product.toLowerCase() === p && r.process.toLowerCase() === t);
  return hit ? hit.minutes : 0;
}

module.exports = {
  FLOOR_TASKS,
  listUsers,
  seedLocalAdminIfEmpty,
  upsertUser,
  deleteUser,
  verifyUser,
  loginFailureMessage,
  createSession,
  readSession,
  persistSessions,
  sessionCount: () => sessions.size,
  listDurations,
  setDurations,
  durationMinutes,
  countdownRemainingMs,
  usersSheet
};
