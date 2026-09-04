const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const { getBook, persistWorkbook, dataDir } = require("./workbook-store");

const FLOOR_TASKS = [
  "Profile Cutting", "Plate Cutting", "Tagging", "Welding", "Grinding",
  "Quality Control", "Paint Preparation", "Painting", "Assembly"
];

const ENQUIRY_ROLES = ["Costing", "Quoting", "Approval", "Follow-up"];
const USER_HEADERS = ["Name", "Role", "Password", "Tasks", "Access", "See Debtors", "Enquiry Roles", "Manage Users"];

const sessions = new Map();
const SESSION_MAX_AGE_MS = 90 * 24 * 60 * 60 * 1000;

function sessionsPath() {
  return path.join(dataDir(), "office-sessions.json");
}

function applySessionMap(raw) {
  const now = Date.now();
  Object.keys(raw || {}).forEach((token) => {
    const row = raw[token];
    if (!row || !row.name) return;
    if (row.savedAt && now - Number(row.savedAt) > SESSION_MAX_AGE_MS) return;
    sessions.set(token, {
      name: row.name,
      access: row.access,
      role: String(row.role || "").trim(),
      jobTitle: String(row.jobTitle || row.role || "").trim(),
      isAdmin: !!row.isAdmin,
      canSeeOffice: !!row.canSeeOffice,
      canSeeDebtors: !!row.canSeeDebtors,
      canManageUsers: !!row.canManageUsers,
      tasks: Array.isArray(row.tasks) ? row.tasks : []
    });
  });
}

function loadSessions() {
  try {
    const raw = JSON.parse(fs.readFileSync(sessionsPath(), "utf8"));
    applySessionMap(raw);
    try { require("./sqlite-store").saveSessions(sessions); } catch (e) {}
    return;
  } catch (e) {}
  try {
    const fromSql = require("./sqlite-store").loadSessions();
    if (fromSql) applySessionMap(fromSql);
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
    try { require("./sqlite-store").saveSessions(out); } catch (e) {}
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
    sheet.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS.slice()]);
    persistWorkbook();
    return sheet;
  }
  const norm = headers.map((h) => String(h || "").trim().toLowerCase());
  if (norm.indexOf("access") === -1) sheet.getRange(1, 5).setValue("Access");
  if (norm.indexOf("see debtors") === -1) sheet.getRange(1, 6).setValue("See Debtors");
  if (norm.indexOf("enquiry roles") === -1) sheet.getRange(1, 7).setValue("Enquiry Roles");
  if (norm.indexOf("manage users") === -1) sheet.getRange(1, 8).setValue("Manage Users");
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

function isManagerTitle(role) {
  return String(role || "").trim().toLowerCase() === "manager";
}

function parseAccess(accessCell, roleCell) {
  if (isManagerTitle(roleCell)) return "Admin";
  const a = String(accessCell || "").trim().toLowerCase();
  if (a === "admin") return "Admin";
  if (a === "production") return "Production";
  const r = String(roleCell || "").trim().toLowerCase();
  if (r === "admin") return "Admin";
  return "Production";
}

function parseYesNo(value) {
  const v = String(value == null ? "" : value).trim().toLowerCase();
  return v === "yes" || v === "true" || v === "1";
}

function parseManageUsers(body, access) {
  if (access !== "Admin") return "No";
  if (body && (body.manageUsers === false || body.canManageUsers === false)) return "No";
  if (body && (body.manageUsers === true || body.canManageUsers === true)) return "Yes";
  return parseYesNo(body && (body.manageUsers != null ? body.manageUsers : body.manage_users)) ? "Yes" : "No";
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

function canonicalizeEnquiryRole(raw) {
  const t = String(raw || "").trim().toLowerCase();
  if (t === "costing" || t === "coster") return "Costing";
  if (t === "quoting" || t === "quote" || t === "quoter") return "Quoting";
  if (t === "approval" || t === "approver") return "Approval";
  if (t === "follow-up" || t === "followup" || t === "follow up" || t === "followups") return "Follow-up";
  return "";
}

function parseEnquiryRoles(cell, access) {
  if (access !== "Admin") return [];
  const parts = Array.isArray(cell)
    ? cell
    : String(cell || "").split(/[,/&+|]+/);
  const set = new Set();
  parts.forEach((p) => {
    const role = canonicalizeEnquiryRole(p);
    if (role) set.add(role);
  });
  return ENQUIRY_ROLES.filter((r) => set.has(r));
}

function namesEqual(a, b) {
  return String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase();
}

function enquiryRoleHolders(role) {
  const want = canonicalizeEnquiryRole(role) || String(role || "").trim();
  if (ENQUIRY_ROLES.indexOf(want) === -1) return [];
  return listUsers()
    .filter((u) => u.canSeeOffice && (u.enquiryRoles || []).indexOf(want) >= 0)
    .map((u) => u.name);
}

function defaultEnquiryAssignee(role, preferred) {
  const holders = enquiryRoleHolders(role);
  if (!holders.length) return String(preferred || "").trim();
  const pref = String(preferred || "").trim();
  if (pref) {
    const hit = holders.find((n) => namesEqual(n, pref));
    if (hit) return hit;
  }
  return holders[0];
}

function enquiryRoleDefaults() {
  return {
    costing: defaultEnquiryAssignee("Costing"),
    quoting: defaultEnquiryAssignee("Quoting"),
    approval: defaultEnquiryAssignee("Approval"),
    followup: defaultEnquiryAssignee("Follow-up"),
    followups: enquiryRoleHolders("Follow-up")
  };
}

function rowToUser(row, id) {
  const access = parseAccess(row[4], row[1]);
  const isAdmin = access === "Admin";
  const debtors = String(row[5] || "").trim().toLowerCase();
  const manage = isAdmin && (parseYesNo(row[7]) || isManagerTitle(row[1]));
  return {
    id,
    name: String(row[0] || "").trim(),
    role: String(row[1] || "").trim(),
    jobTitle: String(row[1] || "").trim(),
    tasks: parseTasks(row[3]),
    access,
    isAdmin,
    canSeeOffice: isAdmin,
    canSeeDebtors: isAdmin && debtors !== "no",
    seeDebtors: isAdmin && debtors !== "no" ? "Yes" : "No",
    enquiryRoles: parseEnquiryRoles(row[6], access),
    canManageUsers: manage,
    manageUsers: manage ? "Yes" : "No"
  };
}

function seedLocalAdminIfEmpty() {
  if (listUsers().length) return false;
  upsertUser({
    name: "Admin",
    access: "Admin",
    role: "Manager",
    password: process.env.LOCAL_ADMIN_CODE || "admin",
    seeDebtors: "Yes",
    manageUsers: "Yes"
  });
  console.log("[staff] no Users yet — seeded local Admin (access code: admin)");
  return true;
}

function listUsers() {
  const sheet = usersSheet();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const grid = sheet.getRange(2, 1, last - 1, USER_HEADERS.length).getValues();
  const out = [];
  for (let i = 0; i < grid.length; i++) {
    const u = rowToUser(grid[i], i + 2);
    if (u.name) out.push(u);
  }
  if (!out.some((u) => u.canManageUsers)) {
    const fallback = out.find((u) => String(u.name).toLowerCase() === "admin" && u.access === "Admin")
      || out.find((u) => u.access === "Admin");
    if (fallback) {
      fallback.canManageUsers = true;
      fallback.manageUsers = "Yes";
    }
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
  let role = String(body.role || "").trim();
  const wantsManage = isManagerTitle(role) || parseManageUsers(body, "Admin") === "Yes";
  if (wantsManage) role = "Manager";
  let access = parseAccess(body.access, role);
  if (isManagerTitle(role)) access = "Admin";
  if (!role) role = access === "Admin" ? "Admin" : "Production";
  const tasks = Array.isArray(body.tasks) ? body.tasks.filter((t) => FLOOR_TASKS.indexOf(t) !== -1) : parseTasks(body.tasks);
  const seeDebtors = parseSeeDebtors(body, access);
  const enquiryRoles = parseEnquiryRoles(body.enquiryRoles != null ? body.enquiryRoles : body.enquiry_roles, access);
  const manageUsers = isManagerTitle(role) ? "Yes" : "No";
  const sheet = usersSheet();
  let rowNum = findUserRow(name);
  let password = accessCode(body.password);
  if (!rowNum) {
    if (!password) throw new Error("Access code is required for a new user");
    rowNum = sheet.getLastRow() + 1;
  } else if (!password) {
    password = String(sheet.getRange(rowNum, 3).getValue() || "");
  }
  sheet.getRange(rowNum, 1, 1, USER_HEADERS.length).setValues([[
    name, role, password, tasks.join(", "), access, seeDebtors, enquiryRoles.join(", "), manageUsers
  ]]);
  if (manageUsers === "Yes") setSoleManager(name);
  persistWorkbook();
  bumpShopCache();
  return rowToUser([name, role, password, tasks.join(", "), access, seeDebtors, enquiryRoles.join(", "), manageUsers], rowNum);
}

function setSoleManager(name) {
  const want = String(name || "").trim().toLowerCase();
  const sheet = usersSheet();
  const last = sheet.getLastRow();
  if (last < 2 || !want) return;
  const grid = sheet.getRange(2, 1, last - 1, USER_HEADERS.length).getValues();
  for (let i = 0; i < grid.length; i++) {
    const isThis = String(grid[i][0] || "").trim().toLowerCase() === want;
    const role = String(grid[i][1] || "").trim();
    if (isThis) {
      grid[i][1] = "Manager";
      grid[i][4] = "Admin";
      grid[i][7] = "Yes";
    } else {
      grid[i][7] = "No";
      if (isManagerTitle(role)) grid[i][1] = "Admin";
    }
  }
  sheet.getRange(2, 1, last - 1, USER_HEADERS.length).setValues(grid);
}

function canManageUsers(profile) {
  if (!profile || !profile.name) return false;
  const live = listUsers().find((u) => String(u.name).toLowerCase() === String(profile.name).toLowerCase());
  return !!(live && live.canManageUsers);
}

function changeOwnPassword(name, currentPassword, nextPassword) {
  const want = String(name || "").trim();
  const current = accessCode(currentPassword);
  const next = accessCode(nextPassword);
  if (!want) throw new Error("Name is required");
  if (!next) throw new Error("New access code is required");
  const rowNum = findUserRow(want);
  if (!rowNum) throw new Error("Current access code is wrong");
  const stored = accessCode(usersSheet().getRange(rowNum, 3).getValue());
  if (!stored || stored !== current) throw new Error("Current access code is wrong");
  usersSheet().getRange(rowNum, 3).setValue(next);
  persistWorkbook();
  bumpShopCache();
  return { name: String(usersSheet().getRange(rowNum, 1).getValue() || want) };
}

function setUserPassword(name, nextPassword) {
  const want = String(name || "").trim();
  const next = accessCode(nextPassword);
  if (!want) throw new Error("Name is required");
  if (!next) throw new Error("New access code is required");
  const rowNum = findUserRow(want);
  if (!rowNum) throw new Error("No user named " + want);
  usersSheet().getRange(rowNum, 3).setValue(next);
  persistWorkbook();
  bumpShopCache();
  return { name: String(usersSheet().getRange(rowNum, 1).getValue() || want) };
}

function deleteUser(name) {
  const rowNum = findUserRow(name);
  if (!rowNum) return;
  const users = listUsers();
  const target = users.find((u) => String(u.name).toLowerCase() === String(name || "").trim().toLowerCase());
  if (target && target.canManageUsers && users.filter((u) => u.canManageUsers).length < 2) {
    throw new Error("Give someone else the Manager job title before deleting this person");
  }
  usersSheet().deleteRow(rowNum);
  persistWorkbook();
  bumpShopCache();
}

function loginFailureMessage() {
  const users = listUsers();
  if (!users.length) {
    return "No users yet. Try again in a moment.";
  }
  return "Incorrect name or access code";
}

function accessCode(value) {
  return String(value == null ? "" : value).trim();
}

function verifyUser(name, password) {
  const sheet = usersSheet();
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const grid = sheet.getRange(2, 1, last - 1, USER_HEADERS.length).getValues();
  const want = String(name || "").trim().toLowerCase();
  const pass = accessCode(password);
  for (let i = 0; i < grid.length; i++) {
    if (String(grid[i][0] || "").trim().toLowerCase() !== want) continue;
    if (pass && pass === accessCode(grid[i][2])) {
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
    role: profile.role || profile.jobTitle || "",
    jobTitle: profile.jobTitle || profile.role || "",
    isAdmin: profile.isAdmin,
    canSeeOffice: profile.canSeeOffice,
    canSeeDebtors: profile.canSeeDebtors,
    canManageUsers: canManageUsers(profile),
    tasks: profile.tasks
  };
  sessions.set(token, safe);
  persistSessions();
  return { token, ...safe };
}

function tokensFromReq(req) {
  const out = [];
  const seen = {};
  function add(raw) {
    const token = String(raw || "").replace(/^Bearer\s+/i, "").trim();
    if (!token || seen[token]) return;
    seen[token] = true;
    out.push(token);
  }
  add((req.headers && (req.headers["x-sd-token"] || req.headers["authorization"])) || "");
  const cookie = String((req.headers && req.headers.cookie) || "");
  const office = cookie.match(/(?:^|; )sd_office=([^;]*)/);
  const shop = cookie.match(/(?:^|; )sd_session=([^;]*)/);
  if (office) {
    try { add(decodeURIComponent(office[1].trim())); } catch (e) { add(office[1].trim()); }
  }
  if (shop) {
    try { add(decodeURIComponent(shop[1].trim())); } catch (e) { add(shop[1].trim()); }
  }
  return out;
}

function tokenFromReq(req) {
  const tokens = tokensFromReq(req);
  for (let i = 0; i < tokens.length; i++) {
    if (sessions.has(tokens[i])) return tokens[i];
  }
  return tokens[0] || "";
}

function readSession(req) {
  const token = tokenFromReq(req);
  if (!token) return null;
  const row = sessions.get(token) || null;
  if (!row) return null;
  if (row.isAdmin) row.canSeeOffice = true;
  const live = listUsers().find((u) => String(u.name).toLowerCase() === String(row.name).toLowerCase());
  if (live) {
    row.canManageUsers = !!live.canManageUsers;
    row.jobTitle = live.jobTitle || live.role || "";
    row.role = live.role || "";
    row.access = live.access;
    row.isAdmin = !!live.isAdmin;
    row.canSeeOffice = !!live.canSeeOffice;
    row.canSeeDebtors = !!live.canSeeDebtors;
  } else {
    row.canManageUsers = canManageUsers(row);
    row.jobTitle = String(row.jobTitle || row.role || "").trim();
  }
  return row;
}

function dropSession(req) {
  const token = tokenFromReq(req);
  if (!token) return false;
  const had = sessions.delete(token);
  if (had) persistSessions();
  return had;
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
  ENQUIRY_ROLES,
  listUsers,
  seedLocalAdminIfEmpty,
  upsertUser,
  deleteUser,
  changeOwnPassword,
  setUserPassword,
  canManageUsers,
  verifyUser,
  loginFailureMessage,
  createSession,
  readSession,
  dropSession,
  persistSessions,
  sessionCount: () => sessions.size,
  listDurations,
  setDurations,
  durationMinutes,
  countdownRemainingMs,
  usersSheet,
  enquiryRoleHolders,
  defaultEnquiryAssignee,
  enquiryRoleDefaults
};
