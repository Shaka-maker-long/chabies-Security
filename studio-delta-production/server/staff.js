const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const { getBook, persistWorkbook, dataDir } = require("./workbook-store");
const accessCodes = require("./access-codes");

const FLOOR_TASKS = [
  "Profile Cutting", "Plate Cutting", "Tagging", "Welding", "Grinding",
  "Quality Control", "Paint Preparation", "Painting", "Assembly"
];

const ENQUIRY_ROLES = ["Costing", "Quoting", "Approval", "Follow-up"];
const USER_HEADERS = ["Name", "Role", "Password", "Tasks", "Access", "See Debtors", "Enquiry Roles", "Manage Users"];

const sessions = new Map();
const SESSION_MAX_AGE_MS = 14 * 24 * 60 * 60 * 1000;

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
      isMarketing: !!row.isMarketing || String(row.access || "").toLowerCase() === "marketing",
      canSeeOffice: !!row.canSeeOffice,
      canSeeDrawingDesk: !!row.canSeeDrawingDesk || isDrawingOwnerName(row.name),
      canSeeDebtors: !!row.canSeeDebtors,
      canManageUsers: !!row.canManageUsers,
      canEditMarketingFields: !!row.canEditMarketingFields || String(row.access || "").toLowerCase() === "marketing",
      tasks: Array.isArray(row.tasks) ? row.tasks : [],
      deviceId: String(row.deviceId || "").trim(),
      savedAt: Number(row.savedAt) || now
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

function hoursFromDurationCell(raw, header) {
  const n = Number(raw) || 0;
  if (!(n > 0)) return 0;
  const label = String(header || "").trim().toLowerCase();
  if (label === "hours" || label === "hour") return Math.round(n * 100) / 100;
  if (n > 24 || (n >= 15 && Math.round(n) === n)) return Math.round((n / 60) * 100) / 100;
  return Math.round(n * 100) / 100;
}

function minutesFromDurationHours(hours) {
  const n = Number(hours) || 0;
  if (!(n > 0)) return 0;
  return Math.max(1, Math.round(n * 60));
}

function migrateTaskDurationSheet(sheet) {
  if (!sheet) return sheet;
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, 3).setValues([["Product", "Process", "Hours"]]);
    persistWorkbook();
    return sheet;
  }
  const header = String(sheet.getRange(1, 3).getValue() || "").trim();
  const label = header.toLowerCase();
  if (label === "hours" || label === "hour") return sheet;
  sheet.getRange(1, 3).setValue("Hours");
  if (sheet.getLastRow() >= 2) {
    const n = sheet.getLastRow() - 1;
    const vals = sheet.getRange(2, 3, n, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      vals[i][0] = hoursFromDurationCell(vals[i][0], header || "Minutes");
    }
    sheet.getRange(2, 3, n, 1).setValues(vals);
  }
  persistWorkbook();
  return sheet;
}

function durationsSheet() {
  const book = getBook();
  let sheet = book.getSheetByName("Task_Durations");
  if (!sheet) {
    sheet = book.insertSheet("Task_Durations");
    sheet.getRange(1, 1, 1, 3).setValues([["Product", "Process", "Hours"]]);
    persistWorkbook();
  }
  return migrateTaskDurationSheet(sheet);
}

function isManagerTitle(role) {
  return String(role || "").trim().toLowerCase() === "manager";
}

function isIdleManagerTitle(role) {
  const t = String(role || "").trim().toLowerCase();
  return t === "manager" || t === "production manager" || t === "site manager";
}

function canSeeIdleAlerts(profile) {
  if (!profile) return false;
  if (isMarketing(profile)) return false;
  if (isIdleManagerTitle(profile.jobTitle || profile.role)) return true;
  return String(profile.name || "").trim().toLowerCase() === "siya";
}

function isMarketingTitle(role) {
  return String(role || "").trim().toLowerCase() === "marketing";
}

function isMarketing(profile) {
  if (!profile) return false;
  if (String(profile.access || "").trim().toLowerCase() === "marketing") return true;
  return isMarketingTitle(profile.jobTitle || profile.role);
}

function isProductionFloorUser(profile) {
  if (!profile || !String(profile.name || "").trim()) return false;
  if (profile.isAdmin) return false;
  if (isMarketing(profile)) return false;
  return String(profile.access || "").toLowerCase() !== "admin";
}

function parseAccess(accessCell, roleCell) {
  if (isManagerTitle(roleCell)) return "Admin";
  const a = String(accessCell || "").trim().toLowerCase();
  if (a === "admin") return "Admin";
  if (a === "marketing") return "Marketing";
  if (a === "production") return "Production";
  const r = String(roleCell || "").trim().toLowerCase();
  if (r === "admin") return "Admin";
  if (r === "marketing") return "Marketing";
  return "Production";
}

const MARKETING_EDIT_FIELDS = ["source", "campaign"];
const MARKETING_WRITE_ERROR = "Marketing can only update Source and Campaign on enquiries and orders.";

function marketingEditFields() {
  return MARKETING_EDIT_FIELDS.slice();
}

function canEditMarketingFields(profile) {
  return isMarketing(profile);
}

function isMutatingHttpMethod(method) {
  const m = String(method || "").toUpperCase();
  return m === "POST" || m === "PUT" || m === "PATCH" || m === "DELETE";
}

function marketingWriteAllowed(method, path) {
  if (!isMutatingHttpMethod(method)) return true;
  const p = String(path || "").split("?")[0].replace(/\/+$/, "") || "/";
  if (String(method).toUpperCase() === "PUT" && (p === "/api/office/orders" || p === "/api/office/enquiries")) {
    return true;
  }
  return false;
}

function applyMarketingFieldPatch(existing, body, opts) {
  opts = opts || {};
  if (!existing) {
    throw new Error(opts.createError || "Marketing can only update Source and Campaign on existing rows.");
  }
  const out = Object.assign({}, existing);
  MARKETING_EDIT_FIELDS.forEach((field) => {
    if (body && Object.prototype.hasOwnProperty.call(body, field)) {
      out[field] = body[field] == null ? "" : String(body[field]);
    }
  });
  return out;
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
  if (access !== "Admin" && access !== "Marketing") return "No";
  if (body.seeDebtors === false || body.canSeeDebtors === false) return "No";
  const fallback = access === "Marketing" ? "No" : "Yes";
  const v = String(body.seeDebtors != null && body.seeDebtors !== "" ? body.seeDebtors : (body.canSeeDebtors != null ? body.canSeeDebtors : fallback))
    .trim()
    .toLowerCase();
  if (v === "no" || v === "false" || v === "0") return "No";
  if (access === "Marketing" && v === "") return "No";
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
  const prior = Number(order.priorWorkMs) || 0;
  return target * 60 * 1000 - Math.max(0, end - start - pauseMs) - prior;
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

const DRAWING_OWNER = "Erin";

function isDrawingOwnerName(name) {
  const n = String(name || "").trim();
  if (!n) return false;
  const first = n.split(/\s+/)[0];
  return n.toLowerCase() === "erin" || first.toLowerCase() === "erin";
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
  const marketing = access === "Marketing";
  const debtors = String(row[5] || "").trim().toLowerCase();
  const manage = isAdmin && (parseYesNo(row[7]) || isManagerTitle(row[1]));
  return {
    id,
    name: String(row[0] || "").trim(),
    role: String(row[1] || "").trim(),
    jobTitle: String(row[1] || "").trim(),
    tasks: marketing ? [] : parseTasks(row[3]),
    access,
    isAdmin,
    isMarketing: marketing,
    canSeeOffice: isAdmin || marketing,
    canSeeDrawingDesk: isDrawingOwnerName(String(row[0] || "").trim()),
    canSeeDebtors: isAdmin ? debtors !== "no" : (marketing && debtors === "yes"),
    seeDebtors: (isAdmin && debtors !== "no") || (marketing && debtors === "yes") ? "Yes" : "No",
    enquiryRoles: parseEnquiryRoles(row[6], access),
    canManageUsers: manage,
    manageUsers: manage ? "Yes" : "No",
    canEditMarketingFields: marketing
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
  if (isMarketingTitle(role) && access !== "Admin") access = "Marketing";
  if (access === "Marketing" && !role) role = "Marketing";
  if (!role) role = access === "Admin" ? "Admin" : (access === "Marketing" ? "Marketing" : "Production");
  const tasks = access === "Marketing"
    ? []
    : (Array.isArray(body.tasks) ? body.tasks.filter((t) => FLOOR_TASKS.indexOf(t) !== -1) : parseTasks(body.tasks));
  const seeDebtors = parseSeeDebtors(body, access);
  const enquiryRoles = access === "Marketing" ? [] : parseEnquiryRoles(body.enquiryRoles != null ? body.enquiryRoles : body.enquiry_roles, access);
  const manageUsers = isManagerTitle(role) ? "Yes" : "No";
  const sheet = usersSheet();
  let rowNum = findUserRow(name);
  let password = accessCode(body.password);
  if (!rowNum) {
    if (!password) throw new Error("Access code is required for a new user");
    password = accessCodes.hashPlain(password);
    rowNum = sheet.getLastRow() + 1;
  } else if (!password) {
    password = String(sheet.getRange(rowNum, 3).getValue() || "");
  } else {
    password = accessCodes.hashPlain(password);
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

function canMarkNoPlate(profile) {
  return require("./no-plates").canMarkNoPlate(profile);
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
  if (!accessCodes.strengthOk(next)) {
    throw new Error("Access code must be at least " + accessCodes.MIN_LEN + " characters.");
  }
  const rowNum = findUserRow(want);
  if (!rowNum) throw new Error("Current access code is wrong");
  const stored = String(usersSheet().getRange(rowNum, 3).getValue() || "");
  if (!accessCodes.verify(current, stored)) throw new Error("Current access code is wrong");
  usersSheet().getRange(rowNum, 3).setValue(accessCodes.hashPlain(next));
  persistWorkbook();
  bumpShopCache();
  return { name: String(usersSheet().getRange(rowNum, 1).getValue() || want) };
}

function setUserPassword(name, nextPassword) {
  const want = String(name || "").trim();
  const next = accessCode(nextPassword);
  if (!want) throw new Error("Name is required");
  if (!next) throw new Error("New access code is required");
  if (!accessCodes.strengthOk(next)) {
    throw new Error("Access code must be at least " + accessCodes.MIN_LEN + " characters.");
  }
  const rowNum = findUserRow(want);
  if (!rowNum) throw new Error("No user named " + want);
  usersSheet().getRange(rowNum, 3).setValue(accessCodes.hashPlain(next));
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
    const stored = String(grid[i][2] || "");
    if (pass && accessCodes.verify(pass, stored)) {
      if (accessCodes.needsRehash(stored)) {
        try {
          sheet.getRange(i + 2, 3).setValue(accessCodes.hashPlain(pass));
          persistWorkbook();
          bumpShopCache();
        } catch (e) {
          console.warn("[staff] could not rehash access code:", e && e.message ? e.message : e);
        }
      }
      return rowToUser(grid[i], i + 2);
    }
  }
  return null;
}

function checkRestoreSecrets(actor, secrets) {
  secrets = secrets || {};
  const confirm = String(secrets.confirm || "").trim().toUpperCase();
  const password = String(secrets.password || "");
  const confirmPassword = String(secrets.confirmPassword || secrets.passwordConfirm || "");
  if (confirm !== "RESTORE") throw new Error("Type RESTORE to restore this backup.");
  if (!password || !confirmPassword) throw new Error("Enter your access code twice to restore.");
  if (password !== confirmPassword) throw new Error("The two access codes do not match.");
  if (!actor || !actor.name) throw new Error("Log in as Manager first.");
  if (!canManageUsers(actor)) throw new Error("Only the Manager can restore a backup.");
  const user = verifyUser(actor.name, password);
  if (!user) throw new Error("Access code is wrong. Restore was not started.");
  return true;
}

function createSession(profile, opts) {
  opts = opts || {};
  const token = crypto.randomBytes(24).toString("hex");
  const deviceId = String(opts.deviceId || profile.deviceId || "").trim();
  const safe = {
    name: profile.name,
    access: profile.access,
    role: profile.role || profile.jobTitle || "",
    jobTitle: profile.jobTitle || profile.role || "",
    isAdmin: profile.isAdmin,
    isMarketing: isMarketing(profile),
    canSeeOffice: profile.canSeeOffice,
    canSeeDrawingDesk: !!(profile.canSeeDrawingDesk || isDrawingOwnerName(profile.name)),
    canSeeDebtors: profile.canSeeDebtors,
    canManageUsers: canManageUsers(profile),
    canSeeIdleAlerts: canSeeIdleAlerts(profile),
    canEditMarketingFields: canEditMarketingFields(profile),
    tasks: profile.tasks,
    deviceId,
    savedAt: Date.now()
  };
  sessions.set(token, safe);
  persistSessions();
  return { token, ...safe };
}

function dropSessionsForDevice(deviceId) {
  const want = String(deviceId || "").trim();
  if (!want) return 0;
  let n = 0;
  sessions.forEach((row, token) => {
    if (row && String(row.deviceId || "").trim() === want) {
      sessions.delete(token);
      n += 1;
    }
  });
  if (n) persistSessions();
  return n;
}

function dropSessionsForUser(name) {
  const want = String(name || "").trim().toLowerCase();
  if (!want) return 0;
  let n = 0;
  sessions.forEach((row, token) => {
    if (row && String(row.name || "").trim().toLowerCase() === want) {
      sessions.delete(token);
      n += 1;
    }
  });
  if (n) persistSessions();
  return n;
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

function readSession(req, opts) {
  opts = opts || {};
  const token = tokenFromReq(req);
  if (!token) return null;
  const row = sessions.get(token) || null;
  if (!row) return null;
  if (row.savedAt && Date.now() - Number(row.savedAt) > SESSION_MAX_AGE_MS) {
    sessions.delete(token);
    persistSessions();
    return null;
  }
  if (row.isAdmin) row.canSeeOffice = true;
  const live = listUsers().find((u) => String(u.name).toLowerCase() === String(row.name).toLowerCase());
  if (live) {
    row.canManageUsers = !!live.canManageUsers;
    row.jobTitle = live.jobTitle || live.role || "";
    row.role = live.role || "";
    row.access = live.access;
    row.isAdmin = !!live.isAdmin;
    row.canSeeOffice = !!live.canSeeOffice;
    row.canSeeDrawingDesk = !!live.canSeeDrawingDesk || isDrawingOwnerName(live.name);
    row.canSeeDebtors = !!live.canSeeDebtors;
    row.canSeeIdleAlerts = canSeeIdleAlerts(live);
  } else {
    row.canManageUsers = canManageUsers(row);
    row.jobTitle = String(row.jobTitle || row.role || "").trim();
    row.canSeeDrawingDesk = !!(row.canSeeDrawingDesk || isDrawingOwnerName(row.name));
    row.canSeeIdleAlerts = canSeeIdleAlerts(row);
  }
  if (!opts.skipDeviceCheck && String(process.env.SD_TRUST_DEVICES || "").trim() !== "0") {
    try {
      const trustedDevices = require("./trusted-devices");
      const headerId = req && req.headers ? req.headers["x-sd-device-id"] : "";
      const deviceId = trustedDevices.normalizeDeviceId(headerId) || String(row.deviceId || "").trim();
      const check = trustedDevices.sessionStillValid(row.name, deviceId);
      if (!check.ok) {
        sessions.delete(token);
        persistSessions();
        return null;
      }
      if (deviceId && !row.deviceId) row.deviceId = deviceId;
    } catch (e) {
      console.warn("[staff] device session check failed:", e && e.message ? e.message : e);
    }
  }
  return row;
}

function dropSession(req) {
  const token = tokenFromReq(req);
  if (!token) return false;
  const had = sessions.delete(token);
  persistSessions();
  return had;
}

function keepRequestSession(req) {
  const token = tokenFromReq(req);
  if (!token) return null;
  const profile = sessions.get(token);
  if (!profile) return null;
  return { token, profile: { ...profile } };
}

function reloadSessionsKeeping(kept) {
  sessions.clear();
  loadSessions();
  if (kept && kept.token && kept.profile) {
    sessions.set(kept.token, kept.profile);
    persistSessions();
  }
}

function durationHoursFromRow(r) {
  if (!r) return 0;
  if (r.hours != null && String(r.hours).trim() !== "") {
    const n = Number(r.hours);
    return n > 0 ? Math.round(n * 100) / 100 : 0;
  }
  const mins = Number(r.minutes) || 0;
  if (mins > 0) return Math.round((mins / 60) * 100) / 100;
  return 0;
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
      const hours = Number(r[2]) || 0;
      if (product && process) {
        rows.push({
          product,
          process,
          hours,
          minutes: minutesFromDurationHours(hours)
        });
      }
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
  sheet.getRange(1, 3).setValue("Hours");
  const clean = (rows || []).filter((r) => r && r.product && r.process && durationHoursFromRow(r) > 0)
    .map((r) => [String(r.product).trim(), String(r.process).trim(), durationHoursFromRow(r)]);
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
  return hit ? minutesFromDurationHours(hit.hours) : 0;
}

module.exports = {
  FLOOR_TASKS,
  ENQUIRY_ROLES,
  MARKETING_EDIT_FIELDS,
  MARKETING_WRITE_ERROR,
  listUsers,
  seedLocalAdminIfEmpty,
  upsertUser,
  deleteUser,
  changeOwnPassword,
  setUserPassword,
  canManageUsers,
  canMarkNoPlate,
  canSeeIdleAlerts,
  isDrawingOwnerName,
  DRAWING_OWNER,
  isProductionFloorUser,
  isManagerTitle,
  isMarketingTitle,
  isMarketing,
  canEditMarketingFields,
  marketingEditFields,
  marketingWriteAllowed,
  isMutatingHttpMethod,
  applyMarketingFieldPatch,
  verifyUser,
  checkRestoreSecrets,
  loginFailureMessage,
  createSession,
  readSession,
  dropSession,
  dropSessionsForDevice,
  dropSessionsForUser,
  keepRequestSession,
  reloadSessionsKeeping,
  persistSessions,
  sessionCount: () => sessions.size,
  SESSION_MAX_AGE_MS,
  listDurations,
  setDurations,
  durationMinutes,
  countdownRemainingMs,
  usersSheet,
  enquiryRoleHolders,
  defaultEnquiryAssignee,
  enquiryRoleDefaults
};
