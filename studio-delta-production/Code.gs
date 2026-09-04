var SHEET_ID = "1pdvAFTIyd5sf8Wbf38MSd4cfk3mb3McPqJrYeM8SOYk"; 
var TAB_ORDERS = "ORDERS";
var TAB_USERS = "Users";
var TAB_LOGS = "Production_Log";
var TAB_OVERVIEW = "Overview";
var TAB_RATES = "Rates";
var FOLDER_NAME = "Studio_Delta_QC_Records";
var TAB_STEEL_PROFILES = "Steel_Profiles";
var TAB_STEEL_USAGE = "Steel_Usage";
var TAB_BACKBOARDS = "Backboards";
var TAB_BACKBOARD_USAGE = "Backboard_Usage";

var TEMP_ID_PRE_POWDER = "18gdKTtaJFqG7EALy-OofLoBxcJ873U4sUTUks3_2oEo";
var TEMP_ID_FINISHED   = "1WXW4F_PIjcA5v2ZSqJDptQrKtlikiVuqOeUy706rV7I";
var QC_EMAIL_RECIPIENT = "siyabonga.msiza@studiodelta.co.za,shaka.chabalala@deltabec.com";
var ALERT_EMAIL_RECIPIENT = "siyabonga.msiza@studiodelta.co.za,shaka.chabalala@deltabec.com";
var POWDER_FOLDER_NAME = "Studio_Delta_Powder_Lists";
var POWDER_EMAIL_RECIPIENT = "siyabonga.msiza@studiodelta.co.za";
var QUEUE_FOLDER_ID = "1MRl3nX7-4d8dmrjQU0UrCbzCf6Ilymub";
var TAB_IDLE = "Idle_Alerts";
var TAB_SCHEDULE = "Schedule";
var TZ_JOBURG = "Africa/Johannesburg";
var SAST_OFFSET_MS = 2 * 60 * 60 * 1000; // South Africa has no DST
var STANDARD_DAY_MINS = 7 * 60 + 30; // paid shift: 07:45-15:45 minus 30 min lunch
var SHIFT_START_MINS = 7 * 60 + 45;
var SHIFT_END_MINS = 15 * 60 + 45;
var LUNCH_START_MINS = 12 * 60;
var LUNCH_END_MINS = 12 * 60 + 30;
var SHIFT_DURATION = STANDARD_DAY_MINS;
var IDLE_GRACE_MINS = 10;
var INDIRECT_TASKS = ["Cleaning", "Maintenance", "Material handling", "Waiting for materials", "Waiting for plate", "Meeting", "Training", "Other"];


var KNOWN_FLOOR_TASKS = ["Profile Cutting", "Plate Cutting", "Tagging", "Welding", "Grinding", "Quality Control", "Paint Preparation", "Painting", "Assembly"];

var TASK_ALIAS_MAP = {
  "profile cutting": "Profile Cutting",
  "profile cutter": "Profile Cutting",
  "steelwork": "Profile Cutting",
  "plate cutting": "Plate Cutting",
  "plate cutter": "Plate Cutting",
  "plate": "Plate Cutting",
  "tagging": "Tagging",
  "tagger": "Tagging",
  "welding": "Welding",
  "welder": "Welding",
  "grinding": "Grinding",
  "grinder": "Grinding",
  "quality control": "Quality Control",
  "qc": "Quality Control",
  "final qc": "Quality Control",
  "pre-powder": "Quality Control",
  "powder coating": "Quality Control",
  "assembly": "Assembly",
  "assembler": "Assembly",
  "paint preparation": "Paint Preparation",
  "paint prep": "Paint Preparation",
  "painting preparation": "Paint Preparation",
  "painting": "Painting",
  "painter": "Painting"
};

function ensureUsersSheetTasksColumn() {
  try {
    if (CacheService.getScriptCache().get("usersTasksCol")) return;
  } catch (e) {}
  var ss = getSpreadsheet();
  var sheet = getSheetOrDie(ss, TAB_USERS);
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var foundTasks = false;
  var foundAccess = false;
  var foundDebtors = false;
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim().toLowerCase();
    if (h === "tasks") foundTasks = true;
    if (h === "access") foundAccess = true;
    if (h === "see debtors" || h === "debtors") foundDebtors = true;
  }
  if (!foundTasks) sheet.getRange(1, 4).setValue("Tasks");
  if (!foundAccess) sheet.getRange(1, 5).setValue("Access");
  if (!foundDebtors) sheet.getRange(1, 6).setValue("See Debtors");
  try { CacheService.getScriptCache().put("usersTasksCol", "1", 21600); } catch (e) {}
}

function parseAccessLabel(accessCell, roleCell) {
  var role = String(roleCell || "").trim().toLowerCase();
  if (role === "manager") return "Admin";
  var a = String(accessCell || "").trim().toLowerCase();
  if (a === "admin" || a === "manager") return "Admin";
  if (a === "production") return "Production";
  if (role === "admin") return "Admin";
  return "Production";
}

function canonicalTaskName(raw) {
  var s = String(raw || "").trim().toLowerCase().replace(/\s+/g, " ");
  if (!s) return "";
  if (TASK_ALIAS_MAP[s]) return TASK_ALIAS_MAP[s];
  for (var i = 0; i < KNOWN_FLOOR_TASKS.length; i++) {
    if (KNOWN_FLOOR_TASKS[i].toLowerCase() === s) return KNOWN_FLOOR_TASKS[i];
  }
  return "";
}

function parseUserTasks(roleCell, tasksCell, forceAdmin) {
  var role = String(roleCell || "").trim();
  var extra = String(tasksCell || "").trim();
  var roleLower = role.toLowerCase();
  var isAdmin = !!forceAdmin || roleLower === "admin" || roleLower === "manager";
  if (isAdmin) {
    return { isAdmin: true, isQcOnly: false, tasks: KNOWN_FLOOR_TASKS.slice(), jobTitle: role || "Admin" };
  }

  var found = [];
  function addTask(name) {
    if (!name) return;
    if (found.indexOf(name) === -1) found.push(name);
  }

  var source = extra || role;
  var parts = String(source).split(/[,/&+|]+/);
  for (var p = 0; p < parts.length; p++) {
    addTask(canonicalTaskName(parts[p]));
  }

  if (found.length === 0) {
    var blob = String(source).toLowerCase();
    var keys = Object.keys(TASK_ALIAS_MAP).sort(function(a, b) { return b.length - a.length; });
    for (var k = 0; k < keys.length; k++) {
      var key = keys[k];
      var re = new RegExp("(^|[^a-z])" + key.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "([^a-z]|$)", "i");
      if (re.test(blob)) addTask(TASK_ALIAS_MAP[key]);
    }
  }

  var qcOnlyRole = roleLower === "quality control" || roleLower === "qc";
  if (qcOnlyRole) {
    return { isAdmin: false, isQcOnly: true, tasks: ["Quality Control"], jobTitle: role || "Quality Control" };
  }

  var isQcOnly = found.length === 1 && found[0] === "Quality Control";
  if (found.length === 0 && role) {
    addTask(canonicalTaskName(role));
  }
  return { isAdmin: false, isQcOnly: isQcOnly, tasks: found, jobTitle: role || (found[0] || "") };
}

function readUserRowProfile(row) {
  var name = String(row[0] || "").trim();
  var access = parseAccessLabel(row.length > 4 ? row[4] : "", row[1]);
  var isAdmin = access === "Admin";
  var parsed = parseUserTasks(row[1], row.length > 3 ? row[3] : "", isAdmin);
  var debtorsCell = String(row.length > 5 ? row[5] : "").trim().toLowerCase();
  var manageCell = String(row.length > 7 ? row[7] : "").trim().toLowerCase();
  var isManager = String(row[1] || "").trim().toLowerCase() === "manager";
  return {
    name: name,
    role: String(row[1] || "").trim(),
    jobTitle: parsed.jobTitle,
    access: access,
    isAdmin: isAdmin,
    canSeeOffice: isAdmin,
    canSeeDebtors: isAdmin && debtorsCell !== "no",
    canManageUsers: isAdmin && (isManager || manageCell === "yes" || manageCell === "true" || manageCell === "1"),
    isQcOnly: parsed.isQcOnly,
    tasks: parsed.tasks
  };
}

function getUserProfileByName(workerName) {
  var needle = String(workerName || "").trim().toLowerCase();
  if (!needle) return null;
  var users = getUsersAndRoles();
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].name || "").trim().toLowerCase() === needle) return users[i];
  }
  return null;
}

function workerCanPerformTask(workerName, task) {
  var profile = getUserProfileByName(workerName);
  if (!profile) return false;
  if (profile.isAdmin) return true;
  var want = canonicalTaskName(task) || String(task || "").trim();
  if (profile.tasks.indexOf(want) !== -1) return true;
  // Assemblers and painters may start Paint Preparation without a separate login task.
  if (want === "Paint Preparation" && (profile.tasks.indexOf("Assembly") !== -1 || profile.tasks.indexOf("Painting") !== -1)) return true;
  return false;
}


// --- STEP 1: CENTRALIZED STEEL PROFILE MANAGEMENT ---
/**
 * Returns an array of all steel profile names from the hidden "Steel_Profiles" tab.
 * If the tab is empty or does not exist, seeds it with the default list and returns that.
 */
function getSteelProfiles() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(TAB_STEEL_PROFILES);

  // Create the tab if it does not exist
  if (!sheet) {
    sheet = ss.insertSheet(TAB_STEEL_PROFILES);
    sheet.hideSheet();
    sheet.appendRow(["Category", "Profile Name"]); // Header
    return[];
  }

  var cachedSteel = floorCacheGet("steel");
  if (cachedSteel) return cachedSteel;

  // Read existing profiles
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return[];
  }

  // Skip header row (index 0), return the values as objects
  var profiles =[];
  for (var i = 1; i < data.length; i++) {
    var cat = String(data[i][0]).trim();
    var name = String(data[i][1]).trim();
    
    // Handle legacy 1-column setup gracefully if you have any old rows left
    if (!name && cat) {
        name = cat;
        cat = "Uncategorized";
    }
    if (name) profiles.push({ category: cat, name: name });
  }
  floorCachePut("steel", profiles, CACHE_TTL_STEEL);
  return profiles;
}

function getBackboards() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(TAB_BACKBOARDS);

  if (!sheet) {
    sheet = ss.insertSheet(TAB_BACKBOARDS);
    sheet.hideSheet();
    sheet.appendRow(["Category", "Profile Name"]);
    return [];
  }

  var cached = floorCacheGet("backboards");
  if (cached) return cached;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return [];
  }

  var profiles = [];
  for (var i = 1; i < data.length; i++) {
    var cat = String(data[i][0]).trim();
    var name = String(data[i][1]).trim();
    if (!name && cat) {
      name = cat;
      cat = "Uncategorized";
    }
    if (name) profiles.push({ category: cat, name: name });
  }
  floorCachePut("backboards", profiles, CACHE_TTL_BACKBOARD);
  return profiles;
}

function processNeedsBackboard(role, processName) {
  var r = String(role || "").trim().toLowerCase();
  var p = String(processName || "").trim().toLowerCase();
  return r === "assembly" || p === "assembly" || p === "final qc";
}

function writeBackboardUsage(ss, orderNum, workerName, processName, backboardUsageData) {
  if (!backboardUsageData || backboardUsageData.length === 0) return;
  var usageSheet = ss.getSheetByName(TAB_BACKBOARD_USAGE);
  if (!usageSheet) {
    usageSheet = ss.insertSheet(TAB_BACKBOARD_USAGE);
    usageSheet.hideSheet();
    usageSheet.appendRow(["Timestamp", "Order #", "Worker", "Process", "Type", "Size"]);
  }

  var profilesSheet = ss.getSheetByName(TAB_BACKBOARDS);
  if (!profilesSheet) {
    getBackboards();
    profilesSheet = ss.getSheetByName(TAB_BACKBOARDS);
  }
  var existingProfileData = profilesSheet ? profilesSheet.getDataRange().getValues() : [];
  var existingProfileNames = [];
  for (var p = 1; p < existingProfileData.length; p++) {
    var pv = String(existingProfileData[p][1] || existingProfileData[p][0]).trim().toLowerCase();
    if (pv) existingProfileNames.push(pv);
  }

  backboardUsageData.forEach(function(item) {
    var usageName = item.category && item.category !== "Uncategorized" ? (item.category + " - " + item.type) : item.type;
    usageSheet.appendRow([
      new Date(),
      orderNum,
      workerName,
      processName,
      usageName,
      item.size
    ]);

    if (item.isCustom && profilesSheet) {
      var newProfileLower = String(item.type).trim().toLowerCase();
      if (newProfileLower && existingProfileNames.indexOf(newProfileLower) === -1) {
        profilesSheet.appendRow([item.category || "Custom", String(item.type).trim()]);
        existingProfileNames.push(newProfileLower);
      }
    }
  });
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Studio Delta Production')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0'); 
}

function lazySetup() {
  try { ensureIdleTrigger(); } catch (e) {}
  try { ensureUsersSheetTasksColumn(); } catch (e) {}
}

var _ssMemo = null;
var _sheetMemo = {};

function getSpreadsheet() {
  if (_ssMemo) return _ssMemo;
  if (SHEET_ID && SHEET_ID !== "1pdvAFTIyd5sf8Wbf38MSd4cfk3mb3McPqJrYeM8SOYk") {
    _ssMemo = SpreadsheetApp.openById(SHEET_ID);
  } else {
    _ssMemo = SpreadsheetApp.getActiveSpreadsheet();
  }
  return _ssMemo;
}

function getSheetOrDie(ss, tabName) {
  if (_sheetMemo[tabName]) return _sheetMemo[tabName];
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error("Missing Tab: '" + tabName + "'. Please create it.");
  _sheetMemo[tabName] = sheet;
  return sheet;
}

var _gridMemo = {};
function getSheetGrid(ss, tabName, numCols) {
  var key = tabName + ":" + String(numCols || 0);
  if (_gridMemo[key]) return _gridMemo[key];
  var sheet = getSheetOrDie(ss, tabName);
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  var cols = numCols ? Math.min(numCols, lastCol) : lastCol;
  var values = sheet.getRange(1, 1, lastRow, cols).getValues();
  _gridMemo[key] = values;
  return values;
}

var LOG_SCAN_MAX = 1500;
var _logPackMemo = null;
var _logPackFull = false;

function getLogPack(ss, forceFull) {
  if (_logPackMemo && (!forceFull || _logPackFull)) return _logPackMemo;
  var sheet = getSheetOrDie(ss, TAB_LOGS);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    _logPackMemo = { values: [[]], fromRow: 2 };
    _logPackFull = true;
    return _logPackMemo;
  }
  var cols = Math.min(13, lastCol);
  var header = sheet.getRange(1, 1, 1, cols).getValues();
  if (lastRow === 1) {
    _logPackMemo = { values: header, fromRow: 2 };
    _logPackFull = true;
    return _logPackMemo;
  }
  var from = 2;
  if (!forceFull && (lastRow - 1) > LOG_SCAN_MAX) {
    from = lastRow - LOG_SCAN_MAX + 1;
  }
  var body = sheet.getRange(from, 1, lastRow - from + 1, cols).getValues();
  _logPackMemo = { values: header.concat(body), fromRow: from };
  _logPackFull = (from === 2);
  return _logPackMemo;
}

function invalidateLogPack() {
  _logPackMemo = null;
  _logPackFull = false;
}

function packSheetRow(pack, valuesIndex) {
  return pack.fromRow + valuesIndex - 1;
}
function packValuesIndex(pack, sheetRow) {
  return sheetRow - pack.fromRow + 1;
}

var CACHE_TTL_FLOOR = 90;
var CACHE_TTL_ACTIVITY = 45;
var CACHE_TTL_USERS = 180;
var CACHE_TTL_STEEL = 300;
var CACHE_TTL_BACKBOARD = 300;
var CACHE_TTL_ADMIN = 60;

function floorCacheGen() {
  try { return CacheService.getScriptCache().get("floorGen") || "0"; } catch (e) { return "0"; }
}
function bumpFloorCache() {
  try { CacheService.getScriptCache().put("floorGen", String(new Date().getTime()), 21600); } catch (e) {}
}
function floorCacheGet(key) {
  try {
    var raw = CacheService.getScriptCache().get(floorCacheGen() + ":" + key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}
function floorCachePut(key, value, ttl) {
  try {
    var s = JSON.stringify(value);
    if (s.length > 8000000) return;
    CacheService.getScriptCache().put(floorCacheGen() + ":" + key, s, ttl || CACHE_TTL_FLOOR);
  } catch (e) {}
}

function emptyPlateStatus() {
  return {status: '', assigned: '', isPaused: false, pauseReason: "", logId: "", batchId: "", isBatched: false, startTime: "", pauseMs: 0, pausedAt: "", batchShare: 1};
}

function pauseAccounting(meta, legacyPauseStart) {
  var pauses = (meta && meta.pauses) || [];
  var pauseMs = 0;
  var pausedAt = 0;
  for (var i = 0; i < pauses.length; i++) {
    var ps = pauses[i].start ? new Date(pauses[i].start).getTime() : 0;
    if (!ps) continue;
    if (pauses[i].end) {
      var pe = new Date(pauses[i].end).getTime();
      if (pe > ps) pauseMs += (pe - ps);
    } else {
      pausedAt = ps;
    }
  }
  if (!pausedAt && legacyPauseStart) {
    var legacy = new Date(legacyPauseStart).getTime();
    if (legacy) pausedAt = legacy;
  }
  return { pauseMs: pauseMs, pausedAt: pausedAt || "" };
}

function plateStatusFromLogRow(row) {
  var meta = parseLogMeta(row.length > 12 ? row[12] : "");
  var pauseStart = getOpenPauseStart(meta) || row[9];
  var pauseReason = "";
  if (meta.pauses && meta.pauses.length) pauseReason = meta.pauses[meta.pauses.length - 1].reason || "";
  if (!pauseReason) pauseReason = row.length > 11 ? row[11] : "";
  var acc = pauseAccounting(meta, pauseStart);
  if (!row[6]) {
    return {
      status: 'Plate Cutting',
      assigned: row[2],
      isPaused: !!pauseStart,
      pauseReason: pauseReason || "",
      logId: row[0],
      batchId: meta.batchId || "",
      isBatched: !!(meta.batchId && !meta.batchSplitAt && (meta.batchShare || 1) > 1),
      startTime: row[5] || "",
      pauseMs: acc.pauseMs,
      pausedAt: acc.pausedAt,
      batchShare: meta.batchShare || 1
    };
  }
  return {status: 'Finished', assigned: '', isPaused: false, pauseReason: "", logId: "", batchId: "", isBatched: false, startTime: "", pauseMs: 0, pausedAt: "", batchShare: 1};
}

function buildPlateStatusMap(logData) {
  var map = {};
  for (var i = logData.length - 1; i >= 1; i--) {
    if (String(logData[i][3]).trim() !== 'Plate Cutting') continue;
    var orderNum = String(logData[i][1]);
    if (map.hasOwnProperty(orderNum)) continue;
    map[orderNum] = plateStatusFromLogRow(logData[i]);
  }
  return map;
}

function getPlateCuttingStatus(ss, orderNum) {
  var map = buildPlateStatusMap(getLogPack(ss).values);
  return map[String(orderNum)] || emptyPlateStatus();
}

function setPlateCuttingStatus(ss, orderNum, status, worker) {
  // This function doesn't need to do anything because plate cutting status
  // is automatically derived from Production_Log entries.
  // The status is set when a log entry is created in startOrder().
  // We keep this function for compatibility but it's a no-op.
  return;
}

function clearPlateCuttingStatus(ss, orderNum) {
  // This function doesn't need to do anything because plate cutting completion
  // is automatically tracked when the log entry's end time is set in finishOrder().
  // We keep this function for compatibility but it's a no-op.
  return;
}

// --- CORE: USERS & LOGIN ---
function getUsersAndRoles() {
  var cached = floorCacheGet("users");
  if (cached) return cached;
  var ss = getSpreadsheet();
  var data = getSheetGrid(ss, TAB_USERS, 8);
  if (data.length <= 1) return []; 
  var users = [];
  for (var i = 1; i < data.length; i++) {
    var profile = readUserRowProfile(data[i]);
    users.push({
      name: profile.name,
      role: profile.role,
      jobTitle: profile.jobTitle,
      access: profile.access,
      isAdmin: profile.isAdmin,
      canSeeOffice: profile.canSeeOffice,
      canSeeDebtors: profile.canSeeDebtors,
      canManageUsers: profile.canManageUsers,
      isQcOnly: profile.isQcOnly,
      tasks: profile.tasks
    });
  }
  floorCachePut("users", users, CACHE_TTL_USERS);
  return users;
}

// Helper function for case-insensitive status comparison
function includesStatusCaseInsensitive(statusArray, statusToCheck) {
  if (!statusToCheck) return false;
  var lowerStatus = String(statusToCheck).toLowerCase().trim();
  for (var i = 0; i < statusArray.length; i++) {
    if (String(statusArray[i]).toLowerCase().trim() === lowerStatus) {
      return true;
    }
  }
  return false;
}

function verifyLogin(role, name, password) {
  var ss = getSpreadsheet();
  var sheet = getSheetOrDie(ss, TAB_USERS);
  var data = sheet.getDataRange().getValues();
  var pass = String(password || "").trim();

  for (var i = 1; i < data.length; i++) {
    var rowName = data[i][0];
    var rowRole = data[i][1];
    var rowPass = String(data[i][2] || "").trim();

    if (role === rowRole && name === rowName) {
      if (pass && pass === rowPass) {
        var isAdmin = (role === 'Admin');
        return { success: true, isAdmin: isAdmin };
      }
    }
  }

  return { success: false, error: "Incorrect Access Code" };
}

function verifyGlobalLogin(name, password) {
  var ss = getSpreadsheet();
  var sheet = getSheetOrDie(ss, TAB_USERS);
  var data = getSheetGrid(ss, TAB_USERS, 8);
  if (!data.length) data = sheet.getDataRange().getValues();
  if (!data || data.length <= 1) {
    return { success: false, error: "No Users loaded yet. Wait a few seconds and try again, or use Admin / admin." };
  }

  var lowerName = String(name).trim().toLowerCase();
  var pass = String(password || "").trim();
  for (var i = 1; i < data.length; i++) {
    var rowName = String(data[i][0]).trim().toLowerCase();
    var rowPass = String(data[i][2] || "").trim();

    if (lowerName === rowName && rowName !== "") {
      if (pass && pass === rowPass) {
        var profile = readUserRowProfile(data[i]);
        return {
          success: true,
          name: profile.name,
          role: profile.role,
          jobTitle: profile.jobTitle,
          access: profile.access,
          isAdmin: profile.isAdmin,
          canSeeOffice: profile.canSeeOffice,
          canSeeDebtors: profile.canSeeDebtors,
          canManageUsers: profile.canManageUsers,
          isQcOnly: profile.isQcOnly,
          tasks: profile.tasks
        };
      }
    }
  }

  return { success: false, error: "Incorrect Name or Access Code" };
}

function getOrdersForRole(role, workerName, skipCache) {
  if (workerName && role && role !== "Admin" && !workerCanPerformTask(workerName, role)) {
    return [];
  }
  var cacheKey = "orders:" + String(role || "");
  if (!skipCache) {
    var cached = floorCacheGet(cacheKey);
    if (cached) return cached;
  }
  var ss = getSpreadsheet();
  var data = getSheetGrid(ss, TAB_ORDERS, 7);
  var logData = getLogPack(ss).values;
  var activeAssignments = getActiveAssignmentsFromData(logData);
  var plateMap = (role === 'Plate Cutting') ? buildPlateStatusMap(logData) : {};
  var relevantOrders = [];
  
  var mainVisibilityMap = {
    'Profile Cutting': ['Not Yet Started', 'Ready for Steelwork', 'Profile Cutting'],
    'Tagging': ['Ready for Tagging', 'Tagging'],
    'Welding': ['Ready for Welding', 'Welding'],
    'Grinding': ['Ready for Grinding', 'Grinding'],
    'Quality Control': [
      'Ready for Pre-Powder Coating', 'Pre-Powder Coating', 
      'Ready for Powder Coating', 'Powder Coating', 
      'Ready for Final QC', 'Final QC',
      'Ready for Delivery', 'Out for Delivery'
    ],
    'Assembly': ['Ready for Assembly', 'Assembly', 'Paint Preparation'],
    'Paint Preparation': ['Ready for Assembly', 'Paint Preparation'],
    'Painting': ['Ready for Assembly', 'Paint Preparation', 'Ready for Painting', 'Painting']
  };

  var plateCuttingStages = [
    'not yet started', 
    'ready for steelwork', 'profile cutting', 
    'ready for tagging', 'tagging', 
    'ready for welding', 'welding'
  ];

  for (var i = 1; i < data.length; i++) {
    var orderNum = data[i][1];             
    var mainStatusRaw = data[i][2];        
    var mainStatus = String(mainStatusRaw).trim(); 
    var mainStatusLower = mainStatus.toLowerCase();
    var productName = data[i][6];

    if (!isAllowedStatus(mainStatus)) continue;

    var assignment = activeAssignments[orderNum];
    var assignedWorker = assignment ? assignment.worker : ""; 
    var isPaused = assignment ? assignment.isPaused : false;
    var pauseReason = assignment ? assignment.pauseReason : "";
    var logId = assignment ? assignment.logId : "";
    var batchId = assignment ? assignment.batchId : "";
    var isBatched = assignment ? assignment.isBatched : false;

    if (role === 'Plate Cutting') {
      var plateInfo = plateMap[String(orderNum)] || emptyPlateStatus();
      var isStageValid = plateCuttingStages.indexOf(mainStatusLower) > -1;
      var isAlreadyActive = plateInfo.status === 'Plate Cutting';
      var isNotFinished = plateInfo.status !== 'Finished';

      if ((isStageValid || isAlreadyActive) && isNotFinished) {
        relevantOrders.push({
          rowIndex: i + 1,
          order: orderNum,
          productName: productName,
          status: isAlreadyActive ? 'In Progress' : 'Available',
          assigned: plateInfo.assigned,
          isPaused: plateInfo.isPaused,
          pauseReason: plateInfo.pauseReason,
          isPlateOrder: true,
          logId: plateInfo.logId || "",
          batchId: plateInfo.batchId || "",
          isBatched: !!plateInfo.isBatched,
          startedAt: plateInfo.startTime || "",
          targetMinutes: getTaskDurationMinutes(productName, "Plate Cutting"),
          pauseMs: plateInfo.pauseMs || 0,
          pausedAt: plateInfo.pausedAt || "",
          batchShare: plateInfo.batchShare || 1
        });
      }
      continue; 
    }

    var allowedStatuses = mainVisibilityMap[role] ||[];
    
    if (role === 'Admin' || includesStatusCaseInsensitive(allowedStatuses, mainStatus)) {
      relevantOrders.push({
        rowIndex: i + 1,
        order: orderNum,
        productName: productName,
        status: mainStatus,
        assigned: assignedWorker, 
        isPaused: isPaused,
        pauseReason: pauseReason,
        isPlateOrder: false,
        logId: logId || "",
        batchId: batchId || "",
        isBatched: !!isBatched,
        startedAt: assignment && assignment.startTime ? assignment.startTime : "",
        targetMinutes: getTaskDurationMinutes(productName, role),
        pauseMs: assignment ? (assignment.pauseMs || 0) : 0,
        pausedAt: assignment ? (assignment.pausedAt || "") : "",
        batchShare: assignment ? (assignment.batchShare || 1) : 1
      });
    }
  }
  floorCachePut(cacheKey, relevantOrders, CACHE_TTL_FLOOR);
  return relevantOrders;
}

function getTaskDurationMinutes(product, process) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Task_Durations");
  var rows = [];
  if (sheet && sheet.getLastRow() >= 2) {
    var grid = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < grid.length; i++) {
      var prod = String(grid[i][0] || "").trim();
      var proc = String(grid[i][1] || "").trim();
      var mins = Number(grid[i][2]) || 0;
      if (prod && proc && mins > 0) rows.push({ product: prod, process: proc, minutes: mins });
    }
  }
  var p = String(product || "").trim().toLowerCase();
  var t = String(process || "").trim().toLowerCase();
  for (var r = 0; r < rows.length; r++) {
    if (rows[r].product.toLowerCase() === p && rows[r].process.toLowerCase() === t) return rows[r].minutes;
  }
  return 0;
}

function getTaskDuration(product, process) {
  return { minutes: getTaskDurationMinutes(product, process) };
}

function pollFloor(role, workerName) {
  return {
    orders: getOrdersForRole(role, workerName),
    notice: workerName ? popWorkerNotice(workerName) : null
  };
}

function tallyFloorCounts(orders, allowStatus) {
  var ready = 0;
  var active = 0;
  var paused = 0;
  var allow = null;
  if (allowStatus && allowStatus.length) {
    allow = {};
    for (var a = 0; a < allowStatus.length; a++) {
      allow[String(allowStatus[a] || "").trim().toLowerCase()] = true;
    }
  }
  for (var j = 0; j < (orders || []).length; j++) {
    var o = orders[j];
    if (allow) {
      var st = String(o && o.status || "").trim().toLowerCase();
      if (!allow[st]) continue;
    }
    if (o && o.assigned) {
      if (o.isPaused) paused += 1;
      else active += 1;
    } else {
      ready += 1;
    }
  }
  return { ready: ready, active: active, paused: paused };
}

function getFloorTaskCounts() {
  var tasks = [
    "Profile Cutting", "Plate Cutting", "Tagging", "Welding", "Grinding",
    "Quality Control", "Paint Preparation", "Painting", "Assembly"
  ];
  var out = {};
  for (var i = 0; i < tasks.length; i++) {
    out[tasks[i]] = tallyFloorCounts(getOrdersForRole(tasks[i], "", true), null);
  }
  var qc = getOrdersForRole("Quality Control", "", true);
  out["Pre-Powder Coating QC"] = tallyFloorCounts(qc, [
    "Ready for Pre-Powder Coating", "Pre-Powder Coating"
  ]);
  out["Powder Coating"] = tallyFloorCounts(qc, [
    "Ready for Powder Coating", "Powder Coating"
  ]);
  out["Final QC"] = tallyFloorCounts(qc, [
    "Ready for Final QC", "Final QC"
  ]);
  var fresh = 0;
  var grid = getSheetGrid(getSpreadsheet(), TAB_ORDERS, 3);
  for (var r = 1; r < grid.length; r++) {
    var st = String(grid[r][2] || "").trim().toLowerCase();
    if (!st || st === "not yet started") fresh += 1;
  }
  out["New orders"] = { ready: fresh, active: 0, paused: 0, total: fresh };
  return out;
}

function floorReadyPileId(status) {
  var s = String(status || "").trim().toLowerCase();
  if (!s || s === "not yet started") return "office";
  if (s === "ready for steelwork" || s === "profile cutting") return "steelwork";
  if (s === "ready for tagging" || s === "tagging") return "tagging";
  if (s === "ready for welding" || s === "welding") return "welding";
  if (s === "ready for grinding" || s === "grinding") return "grinding";
  if (s === "ready for pre-powder coating" || s === "pre-powder coating") return "prepowder";
  if (s === "ready for powder coating" || s === "powder coating") return "powder";
  if (s === "paint preparation" || s === "ready for painting") return "prep";
  if (s === "painting") return "painting";
  if (s === "ready for assembly" || s === "assembly") return "assembly";
  if (s === "ready for final qc") return "delivery";
  if (s === "final qc") return "finalqc";
  if (s === "ready for delivery" || s === "out for delivery") return "delivery";
  return "other";
}

function getFloorLayout() {
  var ss = getSpreadsheet();
  var grid = getSheetGrid(ss, TAB_ORDERS, 3);
  var pack = getLogPack(ss);
  var assignments = getActiveAssignmentsFromData(pack.values);
  var plateMap = buildPlateStatusMap(pack.values);
  var office = [];
  var piles = {};
  var workers = {};
  function addWorker(name, row) {
    var key = String(name || "").trim();
    if (!key) return;
    if (!workers[key]) workers[key] = [];
    workers[key].push(row);
  }
  function addPile(id, row) {
    if (!id) id = "other";
    if (!piles[id]) piles[id] = [];
    piles[id].push(row);
  }
  for (var r = 1; r < grid.length; r++) {
    var order = String(grid[r][1] || "").trim();
    if (!order) continue;
    var status = String(grid[r][2] || "").trim();
    var st = status.toLowerCase();
    var asg = assignments[order];
    if (asg && String(asg.worker || "").trim()) {
      addWorker(asg.worker, {
        order: order,
        status: status,
        process: asg.process || "",
        paused: !!asg.isPaused
      });
      continue;
    }
    if (!st || st === "not yet started") {
      office.push({ order: order, status: status || "Not Yet Started" });
      continue;
    }
    if (!isAllowedStatus(status)) continue;
    addPile(floorReadyPileId(status), { order: order, status: status, paused: false });
  }
  var plateOrders = Object.keys(plateMap);
  for (var p = 0; p < plateOrders.length; p++) {
    var info = plateMap[plateOrders[p]];
    if (!info || info.status !== "Plate Cutting" || !info.assigned) continue;
    addWorker(info.assigned, {
      order: plateOrders[p],
      status: "Plate Cutting",
      process: "Plate Cutting",
      paused: !!info.isPaused
    });
  }
  return { office: office, piles: piles, workers: workers };
}

function uniqueOrderNums(arr) {
  var out = [];
  var seen = {};
  for (var i = 0; i < (arr || []).length; i++) {
    var o = String(arr[i] || "").trim();
    if (!o || seen[o]) continue;
    seen[o] = true;
    out.push(o);
  }
  return out;
}

function joinWorkerOrdersTogether(ss, workerName, extraOrderNums, pack) {
  pack = pack || getLogPack(ss);
  var logSheet = getSheetOrDie(ss, TAB_LOGS);
  var overviewSheet = ss.getSheetByName(TAB_OVERVIEW);
  var now = new Date();
  var logs = pack.values;
  var targetMap = {};
  (extraOrderNums || []).forEach(function(o) {
    var n = String(o || "").trim();
    if (n) targetMap[n] = true;
  });
  var open = [];
  for (var i = 1; i < logs.length; i++) {
    if (logs[i][6]) continue;
    if (String(logs[i][2] || "").trim() !== String(workerName || "").trim()) continue;
    var meta = parseLogMeta(logs[i].length > 12 ? logs[i][12] : "");
    if (meta.entryType === "indirect") continue;
    var paused = hasOpenPause(meta.pauses) || !!logs[i][9];
    var ord = String(logs[i][1] || "");
    open.push({ valuesIndex: i, meta: meta, order: ord, paused: paused });
    if (!paused && ord) targetMap[ord] = true;
  }
  var targetOrders = uniqueOrderNums(Object.keys(targetMap));
  var share = targetOrders.length;
  if (share < 2) {
    return { batchId: "", batchShare: 1, handled: {} };
  }

  var batchId = "";
  for (var r = 0; r < open.length; r++) {
    if (open[r].paused) continue;
    if (targetOrders.indexOf(open[r].order) === -1) continue;
    if (open[r].meta.batchId && !open[r].meta.batchSplitAt) {
      batchId = open[r].meta.batchId;
      break;
    }
  }
  if (!batchId) batchId = Utilities.getUuid();

  var handled = {};
  for (var j = 0; j < open.length; j++) {
    var rec = open[j];
    if (targetOrders.indexOf(rec.order) === -1) continue;
    if (rec.paused) {
      rec.meta = closeOpenPauseInMeta(rec.meta, now);
      writeLogPauseState(logSheet, packSheetRow(pack, rec.valuesIndex), rec.meta, logs[rec.valuesIndex][4]);
    }
    var sameBatch = rec.meta.batchId === batchId && !rec.meta.batchSplitAt && Number(rec.meta.batchShare || 1) === share;
    if (sameBatch && !rec.paused) {
      handled[rec.order] = true;
      continue;
    }
    var sheetRow = packSheetRow(pack, rec.valuesIndex);
    logSheet.getRange(sheetRow, 7).setValue(now);
    pack.values[rec.valuesIndex][6] = now;

    var newMeta = defaultLogMeta();
    newMeta.batchId = batchId;
    newMeta.batchShare = share;
    newMeta.entryType = rec.meta.entryType || "production";
    var row = pack.values[rec.valuesIndex];
    var uniqueId = Utilities.getUuid();
    logSheet.appendRow([
      uniqueId, row[1], row[2], row[3], row[4], now, "", "", "",
      "", "", "", JSON.stringify(newMeta)
    ]);
    if (overviewSheet) {
      overviewSheet.appendRow([uniqueId, row[1], row[2], row[4], now, "", ""]);
    }
    handled[rec.order] = true;
  }
  invalidateLogPack();
  return { batchId: batchId, batchShare: share, handled: handled };
}

function startOrder(rowIndex, workerName, role, batchRowIndices, switchReason, workTogether) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000); 
  
  try {
    var ss = getSpreadsheet();
    var sheet = getSheetOrDie(ss, TAB_ORDERS);
    var logSheet = getSheetOrDie(ss, TAB_LOGS);
    var overviewSheet = ss.getSheetByName(TAB_OVERVIEW);
    var orderData = getSheetGrid(ss, TAB_ORDERS, 7);
    var pack = getLogPack(ss);

    var rowsToStart = [];
    if (batchRowIndices && batchRowIndices.length) {
      for (var b = 0; b < batchRowIndices.length; b++) {
        var br = parseInt(batchRowIndices[b], 10);
        if (rowsToStart.indexOf(br) === -1) rowsToStart.push(br);
      }
    }
    if (rowsToStart.indexOf(parseInt(rowIndex, 10)) === -1) {
      rowsToStart.unshift(parseInt(rowIndex, 10));
    }

    if (!workerCanPerformTask(workerName, role)) {
      throw new Error(workerName + " is not assigned to " + role + ". Ask admin to add it on the Users sheet.");
    }

    var exceptOrders = [];
    for (var e = 0; e < rowsToStart.length; e++) {
      var er = rowsToStart[e];
      exceptOrders.push(orderData[er - 1] ? orderData[er - 1][1] : "");
    }

    var batchId = rowsToStart.length > 1 ? Utilities.getUuid() : "";
    var batchShare = rowsToStart.length > 1 ? rowsToStart.length : 1;
    var runningOthers = listRunningOrdersForWorker(pack.values, workerName, exceptOrders);

    if (workTogether) {
      var join = joinWorkerOrdersTogether(ss, workerName, exceptOrders, pack);
      if (join.batchId) {
        batchId = join.batchId;
        batchShare = join.batchShare;
      }
      runningOthers = [];
      pack = getLogPack(ss);
    } else if (runningOthers.length && !isUserPauseReason(switchReason)) {
      return { success: false, needsSwitchReason: true, runningOrders: runningOthers, message: "Choose why you are leaving the current order." };
    }

    closeIndirectTasksForWorker(ss, workerName, pack);
    if (runningOthers.length) {
      autoPauseWorkerOtherJobs(ss, workerName, exceptOrders, switchReason, batchId, pack);
    }

    var started = [];
    var startTime = new Date();
    var plateMap = (role === 'Plate Cutting') ? buildPlateStatusMap(pack.values) : {};
    var activeAssignments = (role === 'Plate Cutting') ? {} : getActiveAssignmentsFromData(pack.values);

    for (var r = 0; r < rowsToStart.length; r++) {
      var thisRow = rowsToStart[r];
      var orderRow = orderData[thisRow - 1] || [];
      var orderNum = orderRow[1];
      var currentStatus = orderRow[2];

      var nextStatus = getStartStatusForRole(currentStatus, role);

      if (role === 'Plate Cutting') {
        var plateInfo = plateMap[String(orderNum)] || emptyPlateStatus();
        if (plateInfo.assigned !== "" && plateInfo.assigned !== workerName) {
          throw new Error("Plate Cutting is already being done by " + plateInfo.assigned);
        }
        if (plateInfo.assigned === workerName) {
          started.push({ order: orderNum, rowIndex: thisRow, logId: plateInfo.logId, newStatus: nextStatus });
          continue;
        }
      } else {
        var currentAssignment = activeAssignments[orderNum];
        var currentAssigned = currentAssignment ? currentAssignment.worker : ""; 
        if (currentAssigned !== "" && currentAssigned !== workerName) {
          throw new Error("Order locked by " + currentAssigned);
        }
        if (currentAssigned === workerName && currentAssignment && !currentAssignment.isPaused) {
          started.push({ order: orderNum, rowIndex: thisRow, logId: currentAssignment.logId, newStatus: currentStatus });
          continue;
        }
        if (currentAssigned === workerName && currentAssignment && currentAssignment.isPaused) {
          resumeWorkerLog(ss, workerName, orderNum, pack);
          started.push({ order: orderNum, rowIndex: thisRow, logId: currentAssignment.logId, newStatus: currentStatus, resumed: true });
          continue;
        }
      }

      var uniqueId = Utilities.getUuid();
      var meta = defaultLogMeta();
      meta.batchId = batchId;
      meta.batchShare = batchShare;
      meta.entryType = "production";

      if (role !== 'Plate Cutting') {
        sheet.getRange(thisRow, 3, 1, 2).setValues([[nextStatus, workerName]]);
      }

      logSheet.appendRow([
        uniqueId, orderNum, workerName, role, nextStatus, startTime, "", "", "",
        "", "", "", JSON.stringify(meta)
      ]);

      if (overviewSheet) {
        overviewSheet.appendRow([uniqueId, orderNum, workerName, nextStatus, startTime, "", ""]);
      }

      started.push({ order: orderNum, rowIndex: thisRow, logId: uniqueId, newStatus: nextStatus });
    }
    
    return {
      success: true,
      newStatus: started[0] ? started[0].newStatus : "",
      logId: started[0] ? started[0].logId : "",
      started: started,
      batchId: batchId
    };
    
  } finally {
    try { bumpFloorCache(); } catch (ignore2) {} lock.releaseLock();
  }
}

// STEP 2: Added steelUsageData as the 7th parameter. orderNumHint is 8th (optional).
function finishOrder(rowIndex, logId, qcData, signatureUrl, filesData, workerName, steelUsageData, orderNumHint, backboardUsageData) {
  var lock = LockService.getScriptLock();
  lock.waitLock(120000); 
  
  try {
    var ss = getSpreadsheet();
    var sheet = getSheetOrDie(ss, TAB_ORDERS);
    var logSheet = getSheetOrDie(ss, TAB_LOGS);
    var overviewSheet = ss.getSheetByName(TAB_OVERVIEW);
    var endTime = new Date(); 
    
    var pack = getLogPack(ss);
    var orderNum = orderNumHint || sheet.getRange(rowIndex, 2).getValue();
    var rowToUpdate = findOpenLogRow(pack, logId, orderNum, workerName);
    if (rowToUpdate === -1) {
      pack = getLogPack(ss, true);
      rowToUpdate = findOpenLogRow(pack, logId, orderNum, workerName);
    }
    
    if (rowToUpdate === -1) {
      throw new Error("Could not find active log entry for " + workerName + " on this order.");
    }

    var logRow = pack.values[packValuesIndex(pack, rowToUpdate)];
    var role = logRow[3];
    var processName = logRow[4];
    var startTime = logRow[5] ? new Date(logRow[5]) : null;
    orderNum = logRow[1];
    var meta = parseLogMeta(logRow.length > 12 ? logRow[12] : "");

    var trueRow = findOrderRowByNumber(sheet, orderNum);
    if (trueRow > 0) rowIndex = trueRow;

    if (role === 'Quality Control') {
      var procLower = String(processName).trim().toLowerCase();
      if (procLower !== 'powder coating' && procLower !== 'out for delivery') {
        if (!qcData || qcData.length === 0) {
          throw new Error("Server rejected: Incomplete QC answers. Please answer all checklist questions.");
        }
        if (!signatureUrl) {
          throw new Error("Server rejected: A signature is required to complete Quality Control.");
        }
      }
    }
    if (role === 'Profile Cutting' || role === 'Plate Cutting') {
      if (!steelUsageData || steelUsageData.length === 0) {
        throw new Error("Server rejected: Steel material usage must be logged before finishing " + role + ".");
      }
    }
    if (processNeedsBackboard(role, processName)) {
      if (!backboardUsageData || backboardUsageData.length === 0) {
        throw new Error("Server rejected: Backboard used must be logged before finishing " + (processName || role) + ".");
      }
    }

    meta = closeOpenPauseInMeta(meta, endTime);
    if (meta.batchId && !meta.batchSplitAt) {
      meta.batchSplitAt = endTime.getTime();
    }

    if(role === 'Plate Cutting') {
        // Plate Cutting Finished -> Do NOT change Order Status
    } else {
        var nextStep = getNextStatus(processName || sheet.getRange(rowIndex, 3).getValue()); 
        sheet.getRange(rowIndex, 3).setValue(nextStep);
        sheet.getRange(rowIndex, 4).setValue("");
    }

    if (steelUsageData && steelUsageData.length > 0) {
      var usageSheet = ss.getSheetByName(TAB_STEEL_USAGE);
      if (!usageSheet) {
        usageSheet = ss.insertSheet(TAB_STEEL_USAGE);
        usageSheet.hideSheet();
        usageSheet.appendRow(["Timestamp", "Order #", "Worker", "Process", "Profile Type", "Size / Length"]);
      }

      var profilesSheet = ss.getSheetByName(TAB_STEEL_PROFILES);
      if (!profilesSheet) {
        getSteelProfiles();
        profilesSheet = ss.getSheetByName(TAB_STEEL_PROFILES);
      }
      var existingProfileData = profilesSheet ? profilesSheet.getDataRange().getValues() : [];
      var existingProfileNames =[];
      for (var p = 1; p < existingProfileData.length; p++) {
        var pv = String(existingProfileData[p][1] || existingProfileData[p][0]).trim().toLowerCase();
        if (pv) existingProfileNames.push(pv);
      }

      steelUsageData.forEach(function(item) {
        var usageName = item.category && item.category !== 'Uncategorized' ? (item.category + " - " + item.type) : item.type;
        usageSheet.appendRow([
          new Date(),
          orderNum,
          workerName,
          processName,
          usageName,
          item.size
        ]);

        if (item.isCustom && profilesSheet) {
          var newProfileLower = String(item.type).trim().toLowerCase();
          if (newProfileLower && existingProfileNames.indexOf(newProfileLower) === -1) {
            profilesSheet.appendRow([item.category || "Custom", String(item.type).trim()]);
            existingProfileNames.push(newProfileLower);
          }
        }
      });
    }

    writeBackboardUsage(ss, orderNum, workerName, processName, backboardUsageData);

    var resultStr = qcData ? qcData.map(function(i){return i.q+": "+i.a}).join("\n") : "Complete";
    logSheet.getRange(rowToUpdate, 7).setValue(endTime);
    logSheet.getRange(rowToUpdate, 8).setValue(resultStr);
    if(signatureUrl) logSheet.getRange(rowToUpdate, 9).setValue(signatureUrl);
    writeLogPauseState(logSheet, rowToUpdate, meta, processName);

    SpreadsheetApp.flush();

    var switched = [];
    if (role === 'Plate Cutting') {
      switched = autoSwitchWeldersAfterPlate(ss, orderNum);
    }

    if (role === 'Quality Control' && qcData && signatureUrl && processName !== 'Powder Coating') {
        try {
            var jobData = {
                logId: logRow[0],
                qcData: qcData,
                signatureUrl: signatureUrl,
                filesData: filesData,
                workerName: workerName,
                orderNum: orderNum, 
                processName: processName,
                rowToUpdate: rowToUpdate
            };
            
            var queueFolder = DriveApp.getFolderById(QUEUE_FOLDER_ID);
            var fileName = "JOB_" + new Date().getTime() + "_" + jobData.orderNum + ".json";
            queueFolder.createFile(fileName, JSON.stringify(jobData), MimeType.PLAIN_TEXT);

        } catch (e) {
            logSheet.getRange(rowToUpdate, 8).setValue(resultStr + "\n\nError adding to PDF Queue: " + e.toString());
        }
    }

    if (overviewSheet) {
      var ovData = overviewSheet.getDataRange().getValues();
      var ovRow = -1;
      for (var k = ovData.length - 1; k >= 0; k--) {
         if (ovData[k][1] == orderNum && 
             ovData[k][3] == processName && 
             ovData[k][5] === "") {
            ovRow = k + 1;
            break;
         }
      }

      if (ovRow > 0) {
        var finishedRow = logSheet.getRange(rowToUpdate, 1, 1, 13).getValues()[0];
        var durationMins = calculateWorkMinutesFromLog(finishedRow);
        var durationStr = formatDurationServer(durationMins);
        overviewSheet.getRange(ovRow, 6).setValue(endTime); 
        overviewSheet.getRange(ovRow, 7).setValue(durationStr);
      }
    }
    
    return { success: true, switched: switched };
    
  } catch(e) {
    return { success: false, error: e.toString() };
  } finally {
    try { try { bumpFloorCache(); } catch (ignore2) {} lock.releaseLock(); } catch (ignore) {}
  }
}

function logWelderSteel(orderNum, workerName, item) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); 
  try {
    var ss = getSpreadsheet();
    var usageSheet = ss.getSheetByName(TAB_STEEL_USAGE);
    if (!usageSheet) {
      usageSheet = ss.insertSheet(TAB_STEEL_USAGE);
      usageSheet.hideSheet();
      usageSheet.appendRow(["Timestamp", "Order #", "Worker", "Process", "Profile Type", "Size / Length"]);
    }

    var profilesSheet = ss.getSheetByName(TAB_STEEL_PROFILES);
    if (!profilesSheet) {
      getSteelProfiles(); // This seeds the sheet on first call
      profilesSheet = ss.getSheetByName(TAB_STEEL_PROFILES);
    }
    
    var existingProfileData = profilesSheet ? profilesSheet.getDataRange().getValues() : [];
    var existingProfileNames = [];
    for (var p = 1; p < existingProfileData.length; p++) {
      var pv = String(existingProfileData[p][1] || existingProfileData[p][0]).trim().toLowerCase();
      if (pv) existingProfileNames.push(pv);
    }

    var usageName = item.category && item.category !== 'Uncategorized' ? (item.category + " - " + item.type) : item.type;
    
    usageSheet.appendRow([
      new Date(),
      orderNum,
      workerName,
      "Welding", // Explicitly log this as the Welding process
      usageName,
      item.size
    ]);

    if (item.isCustom && profilesSheet) {
      var newProfileLower = String(item.type).trim().toLowerCase();
      if (newProfileLower && existingProfileNames.indexOf(newProfileLower) === -1) {
        profilesSheet.appendRow([item.category || "Custom", String(item.type).trim()]);
      }
    }
    SpreadsheetApp.flush(); 
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  } finally {
    try { bumpFloorCache(); } catch (ignore2) {} lock.releaseLock();
  }
}

// --- NEW FUNCTION: Fetch Only Welding Orders ---
function getWeldingOrders() {
  var ss = getSpreadsheet();
  var sheet = getSheetOrDie(ss, TAB_ORDERS);
  var data = sheet.getDataRange().getValues();
  var weldingOrders = [];
  
  // Available to Welders = "Ready for Welding"
  // In Progress for Welders = "Welding"
  var validStatuses = ['ready for welding', 'welding'];

  for (var i = 1; i < data.length; i++) {
    var orderNum = data[i][1];
    var mainStatusRaw = data[i][2];
    var mainStatusLower = String(mainStatusRaw).trim().toLowerCase();
    var productName = data[i][6];

    // If the order is currently at the Welder stage, grab it
    if (validStatuses.indexOf(mainStatusLower) > -1) {
      weldingOrders.push({
        order: orderNum,
        productName: productName
      });
    }
  }
  return weldingOrders;
}

function generateQCPdf(templateId, orderNum, workerName, qcAnswers, sigBase64, photos, shouldEmail) {
  // Default shouldEmail to true if not provided (backward compatibility)
  if (shouldEmail === undefined) {
    shouldEmail = true;
  }
  
  var folder = getFolder();
  var templateFile = DriveApp.getFileById(templateId);
  
  var qcType = (templateId === TEMP_ID_FINISHED) ? "Final QC" : "Pre-powder coating QC";
  var exactFileName = orderNum + "_" + qcType;
  
  var newFile = templateFile.makeCopy(exactFileName, folder);
  var doc = DocumentApp.openById(newFile.getId());
  var body = doc.getBody();

  // 1. Text Replacements
  body.replaceText("{{WorkerName}}", workerName);
  body.replaceText("{{Timestamp}}", new Date().toLocaleString());
  body.replaceText("{{OrderNumber}}", orderNum);

  // 2. Answer Replacements (Y/N -> Yes/No)
  if (qcAnswers) {
    for (var i = 0; i < qcAnswers.length; i++) {
      var answerRaw = qcAnswers[i].a;
      var answerFormatted = answerRaw;
      if (answerRaw === "Y") answerFormatted = "Yes";
      if (answerRaw === "N") answerFormatted = "No";
      body.replaceText("{{Q" + (i+1) + "}}", answerFormatted);
    }
  }

  
  function replaceImageTagWithHeader(tag, base64Data, isOptional, displayName) {
    var r = body.findText(tag);
    if (r) {
      var element = r.getElement();
      var parent = element.getParent();
      
      // Clean up header text
      var headerText = displayName;
      if (!headerText) {
        headerText = tag.replace("{{Image_", "").replace("}}", "").replace("{{", "").replace("}}", "");
        headerText = headerText.replace(/([A-Z])/g, ' $1').trim();
      }

      if (base64Data) {
        // 1. INSERT PAGE BREAK
        // This gives the photo a full page to occupy
        try {
          var parentIndex = body.getChildIndex(parent);
          if (parentIndex > 0) {
            body.insertPageBreak(parentIndex);
          }
        } catch(e) {}

        // 2. SET HEADER TEXT
        var text = element.asText();
        text.setText(headerText + "\r"); 
        text.setBold(true);
        text.setFontSize(16); // Nice big header

        // 3. INSERT IMAGE
        var imgBlob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg');
        var img = parent.insertInlineImage(parent.getChildIndex(element)+1, imgBlob);
        
        // 4. "ASPECT RATIO PRESERVATION" LOGIC
        try {
          // A. Get the original size from the image itself
          var origW = img.getWidth();
          var origH = img.getHeight();
          var aspectRatio = origW / origH; // Example: 1.77 (Landscape) or 0.56 (Portrait)

          // B. Define the "Safe Box" for an A4 page (minus margins)
          // 540 points is roughly the max width for standard margins
          // 720 points is roughly the max height for standard margins + header
          var MAX_WIDTH = 540;  
          var MAX_HEIGHT = 720; 

          // C. Calculate target dimensions
          // First, try to fill the Width
          var finalW = MAX_WIDTH;
          var finalH = finalW / aspectRatio;

          // D. Check for Overflow
          // If filling the width makes the image too tall (Portrait issue), 
          // we switch strategy and fill the Height instead.
          if (finalH > MAX_HEIGHT) {
             finalH = MAX_HEIGHT;
             finalW = finalH * aspectRatio; // Calculate width based on height
          }

          // E. Apply Calculated Dimensions
          // Since finalW and finalH are linked by the aspectRatio, 
          // STRETCHING IS MATHEMATICALLY IMPOSSIBLE here.
          img.setWidth(finalW);
          img.setHeight(finalH);
          
        } catch (e) {
          // Fallback only if image data is corrupt (rare)
          img.setWidth(500); // Sets width, leaves height auto
        }

      } else if (!isOptional) {
        // Required photo missing
        element.asText().setText(headerText + ": No photo provided");
        element.asText().setBold(false);
      } else {
        // Optional photo missing
        parent.removeChild(element);
      }
    }
  }

  // 3. Signature (Special handling for size)
  var sigTag = "{{Signature}}";
  var sigRange = body.findText(sigTag);
  if (sigRange) {
      var sigElem = sigRange.getElement();
      var sigParent = sigElem.getParent();
      
      // Set Header
      sigElem.asText().setText("Signature\r");
      sigElem.asText().setBold(true);

      if (sigBase64) {
        // Note: Signature from canvas usually has "data:image/png;base64," prefix, we must strip it
        var cleanSig = sigBase64.split(',')[1];
        var sigBlob = Utilities.newBlob(Utilities.base64Decode(cleanSig), 'image/png');
        
        var sigImg = sigParent.insertInlineImage(sigParent.getChildIndex(sigElem)+1, sigBlob);
        sigImg.setWidth(200).setHeight(100); // Keep signature smaller/rectangular
      } else {
        sigElem.asText().setText("Signature: Not signed");
      }
  }

  // 4. Process Photos using the new Helper
  // Pre-Powder Coating Template: Front, Left Side, Right Side, Back, Open, Top (optional), Level 1, Level 2 (optional)
  var mapPre = [
    {tag: "{{Image_Front}}", name: "Front", optional: false},
    {tag: "{{Image_LeftSide}}", name: "Left Side", optional: false},
    {tag: "{{Image_RightSide}}", name: "Right Side", optional: false},
    {tag: "{{Image_Back}}", name: "Back", optional: false},
    {tag: "{{Image_Open}}", name: "Open", optional: false},
    {tag: "{{Image_Top}}", name: "Top", optional: true},
    {tag: "{{Image_Level1}}", name: "Level 1", optional: false},
    {tag: "{{Image_Level2}}", name: "Level 2", optional: true}
  ];
  
  // Finished Goods Template: Front, Level 1, Back, Left Side, Right Side, Job Card, Open, Top (optional), Level 2 (optional)
  var mapFin = [
    {tag: "{{Image_Front}}", name: "Front", optional: false},
    {tag: "{{Image_Level1}}", name: "Level 1", optional: false},
    {tag: "{{Image_Back}}", name: "Back", optional: false},
    {tag: "{{Image_LeftSide}}", name: "Left Side", optional: false},
    {tag: "{{Image_RightSide}}", name: "Right Side", optional: false},
    {tag: "{{Image_Card}}", name: "Job Card", optional: false},
    {tag: "{{Image_Open}}", name: "Open", optional: false},
    {tag: "{{Image_Top}}", name: "Top", optional: true},
    {tag: "{{Image_Level2}}", name: "Level 2", optional: true}
  ];
  
  // Detect which template is being used based on tags present
  var useMap = body.findText("{{Image_Card}}") ? mapFin : mapPre;

  // Note: The photos array includes null entries for skipped optional photos
  // The array length matches the number of photo inputs, with null representing skipped optional photos
  for (var j = 0; j < useMap.length; j++) {
    var photoData = null;
    if (photos && j < photos.length && photos[j] !== null) {
      photoData = photos[j].data;
    }
    replaceImageTagWithHeader(useMap[j].tag, photoData, useMap[j].optional, useMap[j].name);
  }

  doc.saveAndClose();
  
  // Convert to PDF
  var pdfBlob = newFile.getAs('application/pdf');
  var pdfFile = folder.createFile(pdfBlob);
  
  // Trash the temp doc
  newFile.setTrashed(true);

  // --- EMAIL SECTION ---
  if (shouldEmail && QC_EMAIL_RECIPIENT && QC_EMAIL_RECIPIENT.trim() !== "") {
    try {
      // Clean, trim whitespace, and parse email addresses into an array
      var emailList = QC_EMAIL_RECIPIENT.split(",")
        .map(function(email) { return email.trim(); })
        .filter(function(email) { return email.length > 0; });

      if (emailList.length > 0) {
        var mainRecipient = emailList[0]; // First email gets 'TO'
        var ccRecipients = emailList.slice(1).join(","); // All remaining emails get 'CC'

        var emailOptions = {
          to: mainRecipient,
          subject: "QC Report: " + orderNum + " (" + workerName + ")",
          htmlBody: "<p>Please find the attached QC report for Order <strong>" + orderNum + "</strong>.</p>" +
                    "<p>Completed by: " + workerName + "<br>Date: " + new Date().toLocaleString() + "</p>",
          attachments: [pdfBlob]
        };

        // Attach CC list if additional recipients exist
        if (ccRecipients.length > 0) {
          emailOptions.cc = ccRecipients;
        }

        MailApp.sendEmail(emailOptions);
      }
    } catch (e) {
      Logger.log("QC Email failed: " + e.toString());
    }
  }
  
  return pdfFile.getUrl();
}

function getFolder() {
  // Forces the system to use YOUR exact folder ID going forward
  return DriveApp.getFolderById("1pyzJ-jcgltJlrCOwjR7c8AIFxwcd2YcK");
}
// --- CALCULATION LOGIC ---
// --- CALCULATION LOGIC ---
function calculateWorkMinutesServer(start, end, taskName, pausedMins, pauseStart) {
  if (!start) return 0;
  var meta = defaultLogMeta();
  if (pauseStart) {
    meta.pauses.push({
      start: new Date(pauseStart).getTime(),
      end: end ? new Date(end).getTime() : null,
      reason: ""
    });
    return calculateWorkMinutesMeta(start, end, taskName, meta, 0);
  }
  var actualEnd = end ? end : new Date();
  var rawMins = calcRawServerMins(start, actualEnd, taskName);
  var pMins = parseFloat(pausedMins) || 0;
  var finalMins = rawMins - pMins;
  return finalMins > 0 ? finalMins : 0;
}

function calcRawServerMins(start, end, taskName) {
  if (!start || !end) return 0;
  if (taskName && String(taskName).trim() === 'Powder Coating') {
    return (end.getTime() - start.getTime()) / 1000 / 60;
  }

  function getWorkingMins(startDate, endDate) {
     var sMins = sastMinsOfDay(startDate);
     var eMins = sastMinsOfDay(endDate);
     
     sMins = Math.max(SHIFT_START_MINS, Math.min(sMins, SHIFT_END_MINS));
     eMins = Math.max(SHIFT_START_MINS, Math.min(eMins, SHIFT_END_MINS));
     
     if (sMins >= eMins) return 0;
     
     var total = eMins - sMins;
     var lunchOverlap = Math.max(0, Math.min(eMins, LUNCH_END_MINS) - Math.max(sMins, LUNCH_START_MINS));
     return total - lunchOverlap;
  }

  var startStamp = sastDayStamp(start);
  var endStamp = sastDayStamp(end);
  var totalMinutes = 0;

  if (startStamp === endStamp) {
    var day = sastDayOfWeek(start);
    if (day !== 0 && day !== 6) {
      totalMinutes = getWorkingMins(start, end);
    }
  } else {
    var startDow = sastDayOfWeek(start);
    if (startDow !== 0 && startDow !== 6) {
      totalMinutes += getWorkingMins(start, sastWallToDate(start, 15, 45));
    }

    var cursor = addSastDays(sastWallToDate(start, 12, 0), 1);
    while (sastDayStamp(cursor) < endStamp) {
      var dow = sastDayOfWeek(cursor);
      if (dow !== 0 && dow !== 6) totalMinutes += SHIFT_DURATION;
      cursor = addSastDays(cursor, 1);
      if (totalMinutes > 200000) break;
    }

    var endDow = sastDayOfWeek(end);
    if (endDow !== 0 && endDow !== 6) {
      totalMinutes += getWorkingMins(sastWallToDate(end, Math.floor(SHIFT_START_MINS / 60), SHIFT_START_MINS % 60), end);
    }
  }

  return totalMinutes;
}

function formatDurationServer(totalMins) {
  var h = Math.floor(totalMins / 60);
  var m = Math.floor(totalMins % 60);
  return (h < 10 ? "0"+h : h) + ":" + (m < 10 ? "0"+m : m);
}

// --- ADMIN: FETCH ALL DATA ---
function weekLabelSast(date) {
  if (!date) return "Unknown Week";
  return Utilities.formatDate(date, TZ_JOBURG, "yyyy - 'Week' ww");
}

function getDashboardDaySlices(row) {
  var start = row[5] ? new Date(row[5]) : null;
  if (start && isNaN(start.getTime())) start = null;
  if (!start) return [];
  var raw = splitWorkByDay(row, null, null);
  var byDay = {};
  var order = [];
  for (var i = 0; i < raw.length; i++) {
    var s = raw[i];
    var stamp = s.dayStamp;
    if (!stamp) continue;
    if (!byDay[stamp]) {
      byDay[stamp] = {
        date: stamp,
        start: s.start ? s.start.getTime() : null,
        end: s.end ? s.end.getTime() : null,
        durationMins: s.mins || 0,
        week: weekLabelSast(s.start || start),
        stillRunning: !!s.stillRunning,
        stopReason: s.stopReason || ""
      };
      order.push(stamp);
    } else {
      var rec = byDay[stamp];
      rec.durationMins += (s.mins || 0);
      if (s.start && (rec.start === null || s.start.getTime() < rec.start)) rec.start = s.start.getTime();
      if (s.stillRunning) {
        rec.stillRunning = true;
        rec.end = null;
      } else if (!rec.stillRunning && s.end && (rec.end === null || s.end.getTime() > rec.end)) {
        rec.end = s.end.getTime();
      }
      if (s.stopReason) rec.stopReason = s.stopReason;
    }
  }
  var out = [];
  for (var j = 0; j < order.length; j++) out.push(byDay[order[j]]);
  return out;
}

function getAdminDashboardData() {
  var cached = floorCacheGet("adminDash:v2");
  if (cached) return cached;
  var ss = getSpreadsheet();
  
  var logData = getLogPack(ss).values.slice(1);
  var logs = logData.map(function(row) {
    var meta = parseLogMeta(row.length > 12 ? row[12] : "");
    var ts = row[6] ? new Date(row[6]) : (row[5] ? new Date(row[5]) : null);
    if (ts && isNaN(ts.getTime())) ts = null;
    var weekStr = ts ? Utilities.formatDate(ts, "Africa/Johannesburg", "yyyy - 'Week' ww") : "Unknown Week";
    
    var startObj = row[5] ? new Date(row[5]) : null;
    if (startObj && isNaN(startObj.getTime())) startObj = null;
    var endObj = row[6] ? new Date(row[6]) : null;
    if (endObj && isNaN(endObj.getTime())) endObj = null;
    var openStart = getOpenPauseStart(meta);
    var durationMins = calculateWorkMinutesFromLog(row);
    
    return {
      order: row[1],
      worker: row[2],
      role: row[3],
      task: row[4],
      start: startObj ? startObj.getTime() : null,
      end: endObj ? endObj.getTime() : null,
      qc: row[7],
      pauseStart: openStart ? new Date(openStart).getTime() : (row[9] ? new Date(row[9]).getTime() : null),
      pausedMins: cumulativePauseMinsFromMeta(meta, row[4]) || parseFloat(row[10]) || 0,
      batchId: meta.batchId || "",
      batchShare: meta.batchShare || 1,
      durationMins: durationMins,
      week: weekStr
    };
  });

  var orderData = getSheetGrid(ss, TAB_ORDERS, 3).slice();
  orderData.shift(); 
  
  var orders = [];
  
  // Only include orders that are in the Allowed Workflow
  for (var i = 0; i < orderData.length; i++) {
    var status = orderData[i][2]; // Col C
    
    if (isAllowedStatus(status)) {
       orders.push({ 
         order: orderData[i][1], // Col B 
         status: status 
       });
    }
  }

  var adminPayload = { logs: logs, orders: orders };
  floorCachePut("adminDash:v2", adminPayload, CACHE_TTL_ADMIN);
  return adminPayload;
}

// --- UTILS ---
function getNextStatus(current) {
  var currentTrimmed = String(current).trim();
  var currentLower = currentTrimmed.toLowerCase();

  // Optional paint loop: Ready for Assembly can go to Paint Preparation instead of Assembly.
  if (currentLower === "paint preparation") return "Ready for Painting";
  if (currentLower === "ready for painting") return "Painting";
  if (currentLower === "painting") return "Ready for Assembly";

  var flow = [
    "Not Yet Started", 
    "Ready for Steelwork", "Profile Cutting", 
    "Ready for Tagging", "Tagging", 
    "Ready for Welding", "Welding", 
    "Ready for Grinding", "Grinding", 
    "Ready for Pre-Powder Coating", "Pre-Powder Coating",
    "Ready for Powder Coating",
    "Powder Coating", 
    "Ready for Assembly", "Assembly", 
    "Ready for Final QC", "Final QC",
    "Ready for Delivery", "Out for Delivery", 
    "Delivered"
  ];

  var idx = -1;
  for (var i = 0; i < flow.length; i++) {
    if (flow[i].toLowerCase() === currentLower) {
      idx = i;
      break;
    }
  }

  return (idx > -1 && idx < flow.length - 1) ? flow[idx + 1] : current; 
}

function getStartStatusForRole(currentStatus, role) {
  var currentLower = String(currentStatus || "").trim().toLowerCase();
  var roleLower = String(role || "").trim().toLowerCase();
  if (roleLower === "plate cutting") return "Plate Cutting";
  if (currentLower === "ready for assembly" && (roleLower === "paint preparation" || roleLower === "painting" || roleLower === "painter")) {
    return "Paint Preparation";
  }
  if (currentLower === "ready for painting" && (roleLower === "painting" || roleLower === "painter")) {
    return "Painting";
  }
  if (currentLower === "paint preparation") return "Paint Preparation";
  if (currentLower === "painting") return "Painting";
  var nextStatus = getNextStatus(currentStatus);
  if (nextStatus && String(nextStatus).indexOf("Ready") === 0) {
    while (nextStatus && String(nextStatus).indexOf("Ready") === 0) {
      var temp = getNextStatus(nextStatus);
      if (!temp || temp === nextStatus) break;
      nextStatus = temp;
    }
  }
  return nextStatus;
}

/**
 * Scans the Production Log to find currently active MAIN FLOW workers.
 * IT IGNORES PLATE CUTTING so parallel work can happen.
 */
function getActiveAssignments(ss) {
  return getActiveAssignmentsFromData(getLogPack(ss).values);
}

function getActiveAssignmentsFromData(logData) {
  var assignments = {};
  for (var i = 1; i < logData.length; i++) {
    if (logData[i][6]) continue;
    var orderNum = logData[i][1];
    if (!orderNum) continue;
    var roleStr = String(logData[i][3]).trim();
    if (roleStr === 'Plate Cutting' || roleStr === 'Out for Delivery' || roleStr === 'Indirect') continue;
    var meta = parseLogMeta(logData[i].length > 12 ? logData[i][12] : "");
    if (meta.entryType === "indirect") continue;
    var pauseStart = getOpenPauseStart(meta) || logData[i][9];
    var pauseReason = "";
    if (meta.pauses && meta.pauses.length) pauseReason = meta.pauses[meta.pauses.length - 1].reason || "";
    if (!pauseReason) pauseReason = logData[i].length > 11 ? logData[i][11] : "";
    var acc = pauseAccounting(meta, pauseStart);
    assignments[orderNum] = {
      worker: logData[i][2],
      process: roleStr,
      isPaused: !!pauseStart,
      pauseReason: pauseReason || "",
      logId: logData[i][0],
      batchId: meta.batchId || "",
      isBatched: !!(meta.batchId && !meta.batchSplitAt && (meta.batchShare || 1) > 1),
      startTime: logData[i][5] || "",
      pauseMs: acc.pauseMs,
      pausedAt: acc.pausedAt,
      batchShare: meta.batchShare || 1
    };
  }
  return assignments;
}

function isAllowedStatus(status) {
  if (!status) return false;
  var s = String(status).trim().toLowerCase();
  
  var allowed = [
    // Pre-Production
    "not yet started", 
    "ready for steelwork", "profile cutting", 
    "ready for tagging", "tagging", 
    "ready for welding", "welding", 
    "ready for grinding", "grinding", 
    // Powder Coating
    "ready for pre-powder coating", "pre-powder coating",
    "ready for powder coating", "powder coating", 
    // Assembly, painting & QC
    "ready for assembly", "paint preparation", "ready for painting", "painting", "assembly", 
    "ready for final qc", "final qc",
    // Delivery
    "ready for delivery", "out for delivery"
    // Note: "delivered" and "completed" are NOT here, so they will be hidden.
  ];
  
  return allowed.indexOf(s) > -1;
}

// --- METRICS DASHBOARD ---
function getMetricsDashboardData() {
  var ss = getSpreadsheet();
  var logSheet = getSheetOrDie(ss, TAB_LOGS);
  var usersSheet = getSheetOrDie(ss, TAB_USERS);
  
  // Get rates data
  var ratesSheet = ss.getSheetByName(TAB_RATES);
  var rates = {};
  if (ratesSheet) {
    var ratesData = ratesSheet.getDataRange().getValues();
    // Skip header row
    for (var i = 1; i < ratesData.length; i++) {
      var role = String(ratesData[i][0]).trim();
      var rate = parseFloat(ratesData[i][1]);
      if (role && !isNaN(rate)) {
        rates[role.toLowerCase()] = rate;
      }
    }
  }
  
  // Get all logs
  var logData = logSheet.getDataRange().getValues();
  
  // Get all workers from Users sheet
  var usersData = usersSheet.getDataRange().getValues();
  var workerRoles = {};
  for (var i = 1; i < usersData.length; i++) {
    var name = usersData[i][0];
    var role = usersData[i][1];
    if (name && role) {
      workerRoles[name] = role;
    }
  }
  
  // Process logs to group by worker and order
  var workerMetrics = {};
  
  for (var i = 1; i < logData.length; i++) {
    var orderNum = logData[i][1];
    var worker = logData[i][2];
    var role = logData[i][3];
    var task = logData[i][4];
    var startTime = logData[i][5] ? new Date(logData[i][5]) : null;
    var endTime = logData[i][6] ? new Date(logData[i][6]) : null;
    
    if (!worker || !orderNum) continue;
    
    // Initialize worker if not exists
    if (!workerMetrics[worker]) {
      workerMetrics[worker] = {
        orders: {}
      };
    }
    
    // Initialize order if not exists
    if (!workerMetrics[worker].orders[orderNum]) {
      workerMetrics[worker].orders[orderNum] = {
        totalMinutes: 0,
        totalCost: 0,
        tasks: []
      };
    }
    
    // Calculate duration
    var durationMins = calculateWorkMinutesFromLog(logData[i]);
    
    // Get hourly rate for this role
    var hourlyRate = rates[role.toLowerCase()] || 0;
    var labourCost = (durationMins / 60) * hourlyRate;
    
    workerMetrics[worker].orders[orderNum].totalMinutes += durationMins;
    workerMetrics[worker].orders[orderNum].totalCost += labourCost;
    workerMetrics[worker].orders[orderNum].tasks.push({
      task: task,
      role: role,
      startTime: startTime ? startTime.getTime() : null,
      endTime: endTime ? endTime.getTime() : null,
      durationMins: durationMins,
      labourCost: labourCost
    });
  }
  
  // Format the data for the frontend
  var result = [];
  for (var worker in workerMetrics) {
    var orders = workerMetrics[worker].orders;
    var orderList = [];
    
    for (var orderNum in orders) {
      var orderData = orders[orderNum];
      orderList.push({
        order: orderNum,
        totalMinutes: orderData.totalMinutes,
        totalCost: orderData.totalCost,
        tasks: orderData.tasks
      });
    }
    
    // Sort by most recent (based on latest task end time or start time)
    orderList.sort(function(a, b) {
      var aTime = 0;
      var bTime = 0;
      
      for (var i = 0; i < a.tasks.length; i++) {
        var t = a.tasks[i].endTime || a.tasks[i].startTime || 0;
        if (t > aTime) aTime = t;
      }
      
      for (var i = 0; i < b.tasks.length; i++) {
        var t = b.tasks[i].endTime || b.tasks[i].startTime || 0;
        if (t > bTime) bTime = t;
      }
      
      return bTime - aTime; // Most recent first
    });
    
    // Take only last 5 orders
    var last5Orders = orderList.slice(0, 5);
    
    // Calculate highest cost from the last 5 orders only
    var maxCost = 0;
    var maxCostOrder = '';
    for (var i = 0; i < last5Orders.length; i++) {
      if (last5Orders[i].totalCost > maxCost) {
        maxCost = last5Orders[i].totalCost;
        maxCostOrder = last5Orders[i].order;
      }
    }
    
    result.push({
      worker: worker,
      role: workerRoles[worker] || 'Unknown',
      orders: last5Orders,
      highestCostOrder: maxCostOrder,
      highestCost: maxCost
    });
  }
  
  return result;
}

/**
 * Sort processes by manufacturing workflow order
 * Returns a sorted array of process names
 */
function sortProcessesByWorkflow(processes) {
  // Define the manufacturing workflow order
  var workflowOrder = [
    'Profile Cutting',
    'Plate Cutting',
    'Tagging',
    'Welding',
    'Grinding',
    'Powder Coating',
    'Paint Preparation',
    'Painting',
    'Assembly'
  ];
  
  // Create a map for quick lookup of order position
  var orderMap = {};
  for (var i = 0; i < workflowOrder.length; i++) {
    orderMap[workflowOrder[i].toLowerCase()] = i;
  }
  
  // Sort processes using the workflow order (create a copy to avoid mutation)
  return processes.slice().sort(function(a, b) {
    var indexA = orderMap[a.toLowerCase()];
    var indexB = orderMap[b.toLowerCase()];
    
    // If both are in the workflow order, sort by their position
    if (indexA !== undefined && indexB !== undefined) {
      return indexA - indexB;
    }
    
    // If only A is in the workflow, it comes first
    if (indexA !== undefined) return -1;
    
    // If only B is in the workflow, it comes first
    if (indexB !== undefined) return 1;
    
    // If neither is in the workflow, sort alphabetically
    return a.localeCompare(b);
  });
}

/**
 * Get order-based metrics with production processes as columns
 * Returns: { processes: [], orders: [{orderNum, productName, processes: {processName: {totalMinutes, totalCost}}}] }
 * Note: Uses calculateWorkMinutesServer() function defined in this file
 */
function getOrderMetrics() {
  var ss = getSpreadsheet();
  var logSheet = getSheetOrDie(ss, TAB_LOGS);
  var ordersSheet = getSheetOrDie(ss, TAB_ORDERS);
  var ratesSheet = ss.getSheetByName(TAB_RATES);
  
  // Get rates data
  var rates = {};
  if (ratesSheet) {
    var ratesData = ratesSheet.getDataRange().getValues();
    for (var i = 1; i < ratesData.length; i++) {
      var role = String(ratesData[i][0]).trim();
      var rate = parseFloat(ratesData[i][1]);
      if (role && !isNaN(rate)) {
        rates[role.toLowerCase()] = rate;
      }
    }
  }
  
  // Get product names
  var ordersData = ordersSheet.getDataRange().getValues();
  var orderProducts = {};
  for (var i = 1; i < ordersData.length; i++) {
    var orderNum = ordersData[i][1]; 
    var productName = ordersData[i][6]; 
    if (orderNum && productName) {
      orderProducts[String(orderNum).trim()] = String(productName).trim();
    }
  }
  
  var logData = logSheet.getDataRange().getValues();
  var processesSet = {};
  var weeksSet = {};
  var orderMetrics = {};
  
  // Process logs - group by order, process, AND week
  for (var i = 1; i < logData.length; i++) {
    var orderNum = logData[i][1];
    var worker = logData[i][2];
    var role = logData[i][3];
    var task = logData[i][4];
    var startTime = logData[i][5] ? new Date(logData[i][5]) : null;
    if (startTime && isNaN(startTime.getTime())) startTime = null;
    var endTime = logData[i][6] ? new Date(logData[i][6]) : null;
    if (endTime && isNaN(endTime.getTime())) endTime = null;
    
    if (!orderNum || !task) continue;
    
    var processName = String(task).trim();
    if (processName.toLowerCase() === 'pre-powder coating' || processName.toLowerCase() === 'final qc') continue;
    
    processesSet[processName] = true;
    
    // Determine the exact week this specific log entry ended
    var ts = endTime ? endTime.getTime() : (startTime ? startTime.getTime() : 0);
    if (isNaN(ts)) ts = 0; // Guard against Invalid Date
    var weekStr = ts ? Utilities.formatDate(new Date(ts), "Africa/Johannesburg", "yyyy - 'Week' ww") : "Unknown Week";
    weeksSet[weekStr] = true;
    
    if (!orderMetrics[orderNum]) {
      orderMetrics[orderNum] = {
        orderNum: orderNum,
        productName: orderProducts[String(orderNum).trim()] || 'Unknown',
        processes: {}
      };
    }
    
    if (!orderMetrics[orderNum].processes[processName]) {
      orderMetrics[orderNum].processes[processName] = {
        totalMinutes: 0,
        totalCost: 0,
        workers:[],
        weeklyData: {}
      };
    }
    
    if (!orderMetrics[orderNum].processes[processName].weeklyData[weekStr]) {
      orderMetrics[orderNum].processes[processName].weeklyData[weekStr] = {
        minutes: 0,
        cost: 0,
        workers:[]
      };
    }
    
    var durationMins = calculateWorkMinutesFromLog(logData[i]);
    var hourlyRate = rates[role.toLowerCase()] || 0;
    var labourCost = (durationMins / 60) * hourlyRate;
    
    // Add to absolute totals
    orderMetrics[orderNum].processes[processName].totalMinutes += durationMins;
    orderMetrics[orderNum].processes[processName].totalCost += labourCost;
    
    // Add to specific week totals
    orderMetrics[orderNum].processes[processName].weeklyData[weekStr].minutes += durationMins;
    orderMetrics[orderNum].processes[processName].weeklyData[weekStr].cost += labourCost;
    
    if (worker) {
      if (orderMetrics[orderNum].processes[processName].workers.indexOf(worker) === -1) {
        orderMetrics[orderNum].processes[processName].workers.push(worker);
      }
      if (orderMetrics[orderNum].processes[processName].weeklyData[weekStr].workers.indexOf(worker) === -1) {
        orderMetrics[orderNum].processes[processName].weeklyData[weekStr].workers.push(worker);
      }
    }
  }
  
  var processes = sortProcessesByWorkflow(Object.keys(processesSet));
  var weeks = Object.keys(weeksSet).sort().reverse();
  var orders =[];
  for (var orderKey in orderMetrics) {
    orders.push(orderMetrics[orderKey]);
  }
  
  return { processes: processes, weeks: weeks, orders: orders };
}

/**
 * Get production trends data for line graph
 * Returns: { processes: [], products: [], orderData: [{orderNum, productName, processes: {}}] }
 * Note: Uses calculateWorkMinutesServer() function defined in this file
 */
function getProductionTrendsData() {
  // It uses the exact same core logic as getOrderMetrics now, but adds the products list
  var data = getOrderMetrics(); 
  
  var productsSet = {};
  for (var i = 0; i < data.orders.length; i++) {
     productsSet[data.orders[i].productName] = true;
  }
  
  return {
    processes: data.processes,
    products: Object.keys(productsSet).sort(),
    weeks: data.weeks,
    orderData: data.orders
  };
}

function reportScratchedGlass(orderNum, workerName) {
  try {
    var emailBody = "<p><strong>URGENT: Scratched Glass Reported</strong></p>" +
                    "<p>Order Number: <strong>" + orderNum + "</strong></p>" +
                    "<p>Reported By: <strong>" + workerName + "</strong></p>" +
                    "<p>Task: Assembly</p>" +
                    "<p>The worker has been blocked from starting the assembly for this order because the glass is scratched.</p>";

    var recipient = (typeof ALERT_EMAIL_RECIPIENT !== 'undefined' && ALERT_EMAIL_RECIPIENT) 
                      ? ALERT_EMAIL_RECIPIENT 
                      : "siyabonga.msiza@studiodelta.co.za,shaka.chabalala@deltabec.com";

    MailApp.sendEmail({
      to: recipient,
      subject: "⚠️ Scratched Glass Alert - Order " + orderNum,
      htmlBody: emailBody
    });

    return { success: true };
  } catch(e) {
    throw new Error("Failed to send email: " + e.toString());
  }
}

function generatePowderCoatingList(listData, workerName) {
  try {
    var dateStr = new Date().toLocaleDateString();
    var timeStr = new Date().toLocaleTimeString();
    var fileName = "Powder_Coating_List_" + new Date().toISOString().slice(0,10);
    
    // Build HTML for the PDF Document
    var html = "<html><head><style>" +
               "body { font-family: Arial, sans-serif; margin: 20px; color: #333; }" +
               "h2 { text-align: center; margin-bottom: 5px; text-transform: uppercase; }" +
               "p { text-align: center; margin-top: 0; font-size: 14px; color: #555; }" +
               "table { width: 100%; border-collapse: collapse; margin-top: 30px; font-size: 11px; }" +
               "th, td { border: 1px solid #999; padding: 10px 8px; text-align: center; vertical-align: middle; }" +
               "th { background-color: #f2f2f2; font-weight: bold; text-transform: uppercase; font-size: 10px; }" +
               ".desc { text-align: left; }" +
               "</style></head><body>" +
               "<h2>Studio Delta - Powder Coating List</h2>" +
               "<p>Generated By: <strong>" + workerName + "</strong> on " + dateStr + " at " + timeStr + "</p>" +
               "<table>" +
               "<thead><tr>" +
               "<th style='width: 12%'>Order #</th>" +
               "<th style='width: 25%'>Item Description</th>" +
               "<th style='width: 15%'>Dimensions<br>(H x W x D)</th>" +
               "<th style='width: 8%'>QTY</th>" +
               "<th style='width: 15%'>Colour</th>" +
               "<th style='width: 25%'>Profiles Used</th>" +
               "</tr></thead><tbody>";
               
    // Loop through selected orders and add rows
    for (var i = 0; i < listData.length; i++) {
      var item = listData[i];
      html += "<tr>" +
              "<td><strong>" + item.order + "</strong></td>" +
              "<td class='desc'>" + item.desc + "</td>" +
              "<td>" + item.dimensions + "</td>" +
              "<td>" + item.qty + "</td>" +
              "<td>" + item.color + "</td>" +
              "<td class='desc'>" + item.profiles + "</td>" +
              "</tr>";
    }
    
    html += "</tbody></table></body></html>";
    
    // Convert the HTML to a PDF Blob
    var blob = HtmlService.createHtmlOutput(html).getAs('application/pdf').setName(fileName + ".pdf");
    
    // Check if folder exists, if not create it
    var folders = DriveApp.getFoldersByName(POWDER_FOLDER_NAME);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(POWDER_FOLDER_NAME);
    
    // Save to Drive
    var file = folder.createFile(blob);
    var fileUrl = file.getUrl();
    
    // Send Email if an address is provided in settings
    if (POWDER_EMAIL_RECIPIENT && POWDER_EMAIL_RECIPIENT.trim() !== "") {
      MailApp.sendEmail({
        to: POWDER_EMAIL_RECIPIENT,
        subject: "Studio Delta - Powder Coating List (" + dateStr + ")",
        htmlBody: "<p>Good day,</p><p>Please find attached the latest Powder Coating List generated by " + workerName + " on " + dateStr + ".</p><p>Regards,<br>Studio Delta</p>",
        attachments: [blob]
      });
    }
    
    return { success: true, url: fileUrl };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function batchStartOrders(rowIndices, workerName, role, switchReason, workTogether) {
  try {
    return startOrder(rowIndices[0], workerName, role, rowIndices, switchReason, workTogether);
  } catch(e) {
    Logger.log("Batch start error: " + e);
    return {success: false, error: e.toString()};
  }
}

function batchFinishOrders(rowIndices, workerName) {
  for(var i = 0; i < rowIndices.length; i++) {
    try {
      // By passing null for qcData, it bypasses checklists (allowed for Powder/Delivery)
      finishOrder(rowIndices[i], null, null, null, null, workerName);
    } catch(e) {
      Logger.log("Batch finish error for row " + rowIndices[i] + ": " + e);
    }
  }
  return {success: true};
}

// --- WORKER PAUSE / RESUME FEATURES ---
function workerPauseOrder(rowIndex, orderNum, workerName, reason) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!isUserPauseReason(reason)) {
      return { success: false, message: "Choose a pause reason: No materials, Touch up (with order number), or Other." };
    }
    var ss = getSpreadsheet();
    var logSheet = getSheetOrDie(ss, TAB_LOGS);
    var pack = getLogPack(ss);
    var logs = pack.values;
    var pausedSomething = false;
    for (var i = logs.length - 1; i >= 1; i--) {
      if (String(logs[i][1]) === String(orderNum) && String(logs[i][2]).trim() === String(workerName).trim() && !logs[i][6]) {
         var meta = parseLogMeta(logs[i].length > 12 ? logs[i][12] : "");
         if (!hasOpenPause(meta.pauses) && !logs[i][9]) {
            var sheetRow = packSheetRow(pack, i);
            meta = addPauseToMeta(meta, reason);
            writeLogPauseState(logSheet, sheetRow, meta, logs[i][4]);
            pausedSomething = true;
         }
      }
    }
    if (!pausedSomething) return {success: false, message: "No active tasks running for you on this order."};
    return {success: true};
  } catch(e) {
    return {success: false, message: e.toString()};
  } finally {
    try { bumpFloorCache(); } catch (ignore2) {} lock.releaseLock();
  }
}

function workerResumeOrder(rowIndex, orderNum, workerName, switchReason, workTogether) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var pack = getLogPack(ss);
    var runningOthers = listRunningOrdersForWorker(pack.values, workerName, [orderNum]);
    if (workTogether) {
      closeIndirectTasksForWorker(ss, workerName, pack);
      var join = joinWorkerOrdersTogether(ss, workerName, [orderNum], pack);
      pack = getLogPack(ss);
      var okTogether = resumeWorkerLog(ss, workerName, orderNum, pack);
      if (!okTogether && !(join.handled && join.handled[String(orderNum)])) {
        return {success: false, message: "No paused tasks found for you on this order."};
      }
      return {success: true, batchId: join.batchId || ""};
    }
    if (runningOthers.length && !isUserPauseReason(switchReason)) {
      return { success: false, needsSwitchReason: true, runningOrders: runningOthers, message: "Choose why you are leaving the current order." };
    }
    if (runningOthers.length) {
      autoPauseWorkerOtherJobs(ss, workerName, [orderNum], switchReason, "", pack);
    }
    closeIndirectTasksForWorker(ss, workerName, pack);
    var ok = resumeWorkerLog(ss, workerName, orderNum, pack);
    if (!ok) return {success: false, message: "No paused tasks found for you on this order."};
    return {success: true};
  } catch(e) {
    return {success: false, message: e.toString()};
  } finally {
    try { bumpFloorCache(); } catch (ignore2) {} lock.releaseLock();
  }
}

function adminPauseOrder(orderNum) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var logSheet = getSheetOrDie(ss, TAB_LOGS);
    var pack = getLogPack(ss);
    var logs = pack.values;
    var pausedSomething = false;
    for (var i = logs.length - 1; i >= 1; i--) {
      if (String(logs[i][1]) === String(orderNum) && !logs[i][6]) {
         var meta = parseLogMeta(logs[i].length > 12 ? logs[i][12] : "");
         if (!hasOpenPause(meta.pauses) && !logs[i][9]) {
            var sheetRow = packSheetRow(pack, i);
            meta = addPauseToMeta(meta, "Admin pause");
            writeLogPauseState(logSheet, sheetRow, meta, logs[i][4]);
            pausedSomething = true;
         }
      }
    }
    if (!pausedSomething) return {success: false, message: "No active tasks running for this order."};
    return {success: true};
  } catch(e) {
    return {success: false, message: e.toString()};
  } finally {
    try { bumpFloorCache(); } catch (ignore2) {} lock.releaseLock();
  }
}

function adminResumeOrder(orderNum) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var logSheet = getSheetOrDie(ss, TAB_LOGS);
    var pack = getLogPack(ss);
    var logs = pack.values;
    var resumedSomething = false;
    for (var i = logs.length - 1; i >= 1; i--) {
      if (String(logs[i][1]) === String(orderNum) && !logs[i][6]) {
         var meta = parseLogMeta(logs[i].length > 12 ? logs[i][12] : "");
         if (hasOpenPause(meta.pauses) || logs[i][9]) {
            var sheetRow = packSheetRow(pack, i);
            meta = closeOpenPauseInMeta(meta, new Date());
            writeLogPauseState(logSheet, sheetRow, meta, logs[i][4]);
            resumedSomething = true;
         }
      }
    }
    if (!resumedSomething) return {success: false, message: "No paused tasks found."};
    return {success: true};
  } catch(e) {
    return {success: false, message: e.toString()};
  } finally {
    try { bumpFloorCache(); } catch (ignore2) {} lock.releaseLock();
  }
}

function processPdfQueue() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return; // Exit if another trigger is already running
  
  try {
    var queueFolder = DriveApp.getFolderById(QUEUE_FOLDER_ID);
    var files = queueFolder.getFilesByType(MimeType.PLAIN_TEXT);
    
    // If no files are waiting, release lock and exit
    if (!files.hasNext()) return; 
    
    var file = files.next();
    var fileContent = file.getBlob().getDataAsString();
    var jobData = JSON.parse(fileContent);
    
    var ss = getSpreadsheet();
    var logSheet = getSheetOrDie(ss, TAB_LOGS);
    
    // --- GENERATE THE PDF ---
    var templateId = TEMP_ID_PRE_POWDER;
    if (jobData.processName === 'Final QC') {
        templateId = TEMP_ID_FINISHED;
    }
    
    // Pass the data to the PDF generator
    var pdfUrl = generateQCPdf(
      templateId, 
      jobData.orderNum, 
      jobData.workerName, 
      jobData.qcData, 
      jobData.signatureUrl, 
      jobData.filesData, 
      true
    );
    
    // --- UPDATE THE LOG SHEET ---
    if (pdfUrl) {
      var currentResult = logSheet.getRange(jobData.rowToUpdate, 8).getValue();
      logSheet.getRange(jobData.rowToUpdate, 8).setValue(currentResult + "\n\nQC PDF: " + pdfUrl);
    }
    
    // --- CLEAN UP ---
    // Delete the job file from the queue folder now that it is complete
    file.setTrashed(true); 
    
  } catch (err) {
    Logger.log("Queue Processing Error: " + err.toString());
    // Rename the file to ERROR_ so it gets skipped next time but isn't lost
    if (file) {
      file.setName("ERROR_" + file.getName());
    }
  } finally {
    try { bumpFloorCache(); } catch (ignore2) {} lock.releaseLock();
  }
}

// --- WEEKLY COMPLETIONS ANALYTICS ---
function getWeeklyAnalyticsData() {
  var ss = getSpreadsheet();
  var logSheet = getSheetOrDie(ss, TAB_LOGS);
  var ordersSheet = getSheetOrDie(ss, TAB_ORDERS);
  
  // Get product names to map to orders
  var ordersData = ordersSheet.getDataRange().getValues();
  var orderProducts = {};
  for (var i = 1; i < ordersData.length; i++) {
    var orderNum = String(ordersData[i][1]).trim();
    var productName = String(ordersData[i][6]).trim(); // Column G is Product Name
    if (orderNum && productName) {
      orderProducts[orderNum] = productName;
    }
  }
  
  var logData = logSheet.getDataRange().getValues();
  
  var weeklyData = {}; 
  var processesSet = {};

  // FIRST PASS: Aggregate completion dates and worker assignments per process per order
  var processAggregates = {};
  // Start from 1 to skip header
  for (var i = 1; i < logData.length; i++) {
    var orderNum = String(logData[i][1]).trim();
    var worker = String(logData[i][2]).trim();
    var task = String(logData[i][4]).trim();
    var endTime = logData[i][6]; // Column G (End Time)
    
    // Only count tasks that actually have an End Time (Completed)
    if (!orderNum || !task || !endTime) continue; 
    
    // Exclude QC tasks from manufacturing throughput (Optional, keeps the board clean)
    var lowerTask = task.toLowerCase();
    if (lowerTask === 'pre-powder coating' || lowerTask === 'final qc') {
       continue; 
    }

    var aggKey = orderNum + "|" + task;
    var ts = new Date(endTime).getTime();
    if (isNaN(ts)) ts = 0;
    
    if (!processAggregates[aggKey]) {
      processAggregates[aggKey] = {
        orderNum: orderNum,
        processName: task,
        maxEndTime: ts,
        workers: worker ? [worker] : []
      };
    } else {
      if (ts > processAggregates[aggKey].maxEndTime) {
        processAggregates[aggKey].maxEndTime = ts;
      }
      if (worker && processAggregates[aggKey].workers.indexOf(worker) === -1) {
        processAggregates[aggKey].workers.push(worker);
      }
    }
  }

  // SECOND PASS: Group aggregated processes by their specific phase completion week
  for (var key in processAggregates) {
    var agg = processAggregates[key];
    var processName = agg.processName;
    processesSet[processName] = true;
    
    // Format the date into "YYYY - Week WW" using the PHASE'S final end time
    var weekStr = "Unknown Week";
    if (agg.maxEndTime && !isNaN(agg.maxEndTime)) {
      weekStr = Utilities.formatDate(new Date(agg.maxEndTime), "Africa/Johannesburg", "yyyy - 'Week' ww");
    }
    
    if (!weeklyData[weekStr]) weeklyData[weekStr] = {};
    if (!weeklyData[weekStr][processName]) weeklyData[weekStr][processName] = [];
    
    // Push the compiled details for the pop-up modal
    weeklyData[weekStr][processName].push({
       orderNum: agg.orderNum,
       productName: orderProducts[agg.orderNum] || 'Unknown Product',
       worker: agg.workers.length > 0 ? agg.workers.join(", ") : "Unknown"
    });
  }
  
  // Use existing helper to sort processes properly (Welding -> Grinding -> Powder)
  var processes = sortProcessesByWorkflow(Object.keys(processesSet));
  
  // Sort weeks descending (newest weeks at the top)
  var weeks = Object.keys(weeklyData).sort().reverse(); 
  
  return {
    weeks: weeks,
    processes: processes,
    data: weeklyData
  };
}

function getQCReportsFast() {
  var folderId = "1pyzJ-jcgltJlrCOwjR7c8AIFxwcd2YcK"; // Your exact QC Folder ID
  var list = [];
  
  try {
    var folder = DriveApp.getFolderById(folderId);
    
    // 1. Instantly unlock the folder for Admins (ignores errors if you aren't the owner)
    try {
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(e) {}
    
    // 2. Get EVERY file directly from this specific folder
    var files = folder.getFiles();
    
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      
      // Ensure we only grab PDFs (checking both name and file type to be safe)
      if (name.toLowerCase().indexOf('.pdf') > -1 || file.getMimeType() === MimeType.PDF) {
        list.push({
          name: name.replace('.pdf', ''),
          url: file.getUrl(),
          dateCreated: file.getDateCreated().getTime()
        });
      }
    }
  } catch(e) {
    throw new Error("Could not read folder. Error: " + e.toString());
  }
  
  // 3. Sort by newest first
  list.sort(function(a, b) { return b.dateCreated - a.dateCreated; });
  
  // Return the most recent 300 so the app stays fast
  return list.slice(0, 300);
}

// =============================================================================
// TIME ENGINE (Johannesburg) — one running clock, pause intervals, batch split
// =============================================================================

function defaultLogMeta() {
  return {
    pauses: [],
    batchId: "",
    batchShare: 1,
    batchSplitAt: null,
    entryType: "production"
  };
}

function parseLogMeta(cell) {
  var meta = defaultLogMeta();
  if (cell === null || cell === undefined || cell === "") return meta;
  var s = String(cell).trim();
  if (!s) return meta;
  try {
    var parsed = JSON.parse(s);
    if (!parsed || typeof parsed !== "object") return meta;
    meta.pauses = parsed.pauses || [];
    meta.batchId = parsed.batchId || "";
    meta.batchShare = parsed.batchShare || 1;
    meta.batchSplitAt = parsed.batchSplitAt || null;
    meta.entryType = parsed.entryType || "production";
    return meta;
  } catch (e) {
    return meta;
  }
}

function writeLogMeta(logSheet, rowNum, meta) {
  logSheet.getRange(rowNum, 13).setValue(JSON.stringify(meta));
}

function writeLogPauseState(logSheet, rowNum, meta, taskName) {
  meta = meta || defaultLogMeta();
  var openStart = getOpenPauseStart(meta);
  var lastReason = "";
  if (meta.pauses && meta.pauses.length) {
    lastReason = meta.pauses[meta.pauses.length - 1].reason || "";
  }
  logSheet.getRange(rowNum, 10, 1, 4).setValues([[
    openStart ? new Date(openStart) : "",
    cumulativePauseMinsFromMeta(meta, taskName),
    openStart ? lastReason : "",
    JSON.stringify(meta)
  ]]);
}

function asSast(date) {
  return new Date(date.getTime() + SAST_OFFSET_MS);
}

function sastDayOfWeek(date) {
  return asSast(date).getUTCDay();
}

function sastMinsOfDay(date) {
  var d = asSast(date);
  return d.getUTCHours() * 60 + d.getUTCMinutes() + d.getUTCSeconds() / 60;
}

function sastDayStamp(date) {
  return Utilities.formatDate(date, TZ_JOBURG, "yyyy-MM-dd");
}

function sastWallToDate(date, hours, minutes) {
  var d = asSast(date);
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), hours - 2, minutes, 0, 0));
}

function addSastDays(date, days) {
  return new Date(date.getTime() + days * 86400000);
}

function isWithinShiftNow() {
  var now = new Date();
  var dow = sastDayOfWeek(now);
  if (dow === 0 || dow === 6) return false;
  var mins = sastMinsOfDay(now);
  return mins >= SHIFT_START_MINS && mins < SHIFT_END_MINS;
}

function getPauseIntervalsForRow(row) {
  var meta = parseLogMeta(row.length > 12 ? row[12] : "");
  var pauses = meta.pauses ? meta.pauses.slice() : [];
  if (pauses.length === 0 && row[9]) {
    pauses.push({
      start: new Date(row[9]).getTime(),
      end: row[6] ? new Date(row[6]).getTime() : null,
      reason: row.length > 11 ? String(row[11] || "") : ""
    });
  }
  return pauses;
}

function sumPauseMinutesInWindow(pauses, wStart, wEnd, taskName) {
  if (!pauses || !wStart || !wEnd) return 0;
  var total = 0;
  var w0 = wStart.getTime();
  var w1 = wEnd.getTime();
  for (var i = 0; i < pauses.length; i++) {
    var ps = new Date(pauses[i].start).getTime();
    var pe = pauses[i].end ? new Date(pauses[i].end).getTime() : w1;
    var a = Math.max(ps, w0);
    var b = Math.min(pe, w1);
    if (b > a) total += calcRawServerMins(new Date(a), new Date(b), taskName);
  }
  return total;
}

function calculateWorkMinutesMeta(start, end, taskName, meta, legacyPausedMins) {
  if (!start) return 0;
  var actualEnd = end ? end : new Date();
  meta = meta || defaultLogMeta();
  var pauses = meta.pauses || [];
  var share = Math.max(1, parseFloat(meta.batchShare) || 1);
  var splitAt = meta.batchSplitAt ? new Date(meta.batchSplitAt) : null;

  if (pauses.length === 0 && legacyPausedMins && !meta.batchId) {
    var rawLegacy = calcRawServerMins(start, actualEnd, taskName);
    return Math.max(0, rawLegacy - (parseFloat(legacyPausedMins) || 0));
  }

  if (splitAt && splitAt.getTime() > start.getTime() && splitAt.getTime() < actualEnd.getTime()) {
    var beforeRaw = calcRawServerMins(start, splitAt, taskName);
    var afterRaw = calcRawServerMins(splitAt, actualEnd, taskName);
    var beforePause = sumPauseMinutesInWindow(pauses, start, splitAt, taskName);
    var afterPause = sumPauseMinutesInWindow(pauses, splitAt, actualEnd, taskName);
    return Math.max(0, beforeRaw - beforePause) / share + Math.max(0, afterRaw - afterPause);
  }

  var raw = calcRawServerMins(start, actualEnd, taskName);
  var pauseMins = sumPauseMinutesInWindow(pauses, start, actualEnd, taskName);
  var net = Math.max(0, raw - pauseMins);
  if (!splitAt && share > 1 && meta.batchId) return net / share;
  return net;
}

function calculateWorkMinutesFromLog(row) {
  var start = row[5] ? new Date(row[5]) : null;
  if (start && isNaN(start.getTime())) start = null;
  var end = row[6] ? new Date(row[6]) : null;
  if (end && isNaN(end.getTime())) end = null;
  var task = row[4];
  var meta = parseLogMeta(row.length > 12 ? row[12] : "");
  if (row[9] && !hasOpenPause(meta.pauses)) {
    meta.pauses = getPauseIntervalsForRow(row);
  }
  return calculateWorkMinutesMeta(start, end, task, meta, row[10]);
}

function getWorkBoutsFromLog(row) {
  var start = row[5] ? new Date(row[5]) : null;
  if (start && isNaN(start.getTime())) start = null;
  if (!start) return [];
  var isOpen = !row[6];
  var end = row[6] ? new Date(row[6]) : new Date();
  if (end && isNaN(end.getTime())) end = new Date();
  if (end.getTime() <= start.getTime()) return [];
  var meta = parseLogMeta(row.length > 12 ? row[12] : "");
  var pauses = (meta.pauses && meta.pauses.length) ? meta.pauses.slice() : getPauseIntervalsForRow(row);
  pauses.sort(function(a, b) { return new Date(a.start).getTime() - new Date(b.start).getTime(); });
  var bouts = [];
  var cursor = start.getTime();
  var endMs = end.getTime();
  for (var i = 0; i < pauses.length; i++) {
    var ps = new Date(pauses[i].start).getTime();
    if (isNaN(ps)) continue;
    var pe = pauses[i].end ? new Date(pauses[i].end).getTime() : endMs;
    if (isNaN(pe)) pe = endMs;
    if (ps > cursor) {
      bouts.push({
        start: new Date(cursor),
        end: new Date(Math.min(ps, endMs)),
        stopReason: String(pauses[i].reason || ""),
        stillRunning: false
      });
    }
    if (pe > cursor) cursor = pe;
    if (cursor >= endMs) break;
  }
  if (endMs > cursor) {
    bouts.push({
      start: new Date(cursor),
      end: new Date(endMs),
      stopReason: "",
      stillRunning: isOpen
    });
  }
  return bouts;
}

function splitWorkByDay(row, rangeFrom, rangeTo) {
  var task = row[4];
  var emptyMeta = defaultLogMeta();
  var slices = [];
  var rangeFromMs = rangeFrom ? rangeFrom.getTime() : 0;
  var rangeToMs = rangeTo ? rangeTo.getTime() : 0;
  var bouts = getWorkBoutsFromLog(row);
  for (var b = 0; b < bouts.length; b++) {
    var bout = bouts[b];
    var boutStartMs = bout.start.getTime();
    var boutEndMs = bout.end.getTime();
    if (rangeToMs && boutStartMs >= rangeToMs) continue;
    if (rangeFromMs && boutEndMs <= rangeFromMs) continue;
    if (rangeFromMs && boutStartMs < rangeFromMs) boutStartMs = rangeFromMs;
    if (rangeToMs && boutEndMs > rangeToMs) boutEndMs = rangeToMs;
    if (boutEndMs <= boutStartMs) continue;
    var clippedStart = new Date(boutStartMs);
    var clippedEnd = new Date(boutEndMs);
    var endStamp = sastDayStamp(clippedEnd);
    var cursor = sastWallToDate(clippedStart, 0, 0);
    var safety = 0;
    var daySlices = [];
    while (sastDayStamp(cursor) <= endStamp && safety < 40) {
      safety++;
      var dayStamp = sastDayStamp(cursor);
      var dayStart = sastWallToDate(cursor, 0, 0);
      var nextMidnight = addSastDays(dayStart, 1);
      var sliceStartMs = Math.max(clippedStart.getTime(), dayStart.getTime());
      var sliceEndMs = Math.min(clippedEnd.getTime(), nextMidnight.getTime());
      if (sliceEndMs > sliceStartMs) {
        var sliceStart = new Date(sliceStartMs);
        var sliceEnd = new Date(sliceEndMs);
        var mins = calculateWorkMinutesMeta(sliceStart, sliceEnd, task, emptyMeta, 0);
        if (mins > 0) {
          daySlices.push({
            dayStamp: dayStamp,
            start: sliceStart,
            end: sliceEnd,
            mins: mins,
            stopReason: "",
            stillRunning: false
          });
        }
      }
      cursor = addSastDays(cursor, 1);
    }
    if (daySlices.length) {
      var pauseInRange = !rangeToMs || bout.end.getTime() <= rangeToMs;
      var still = !!bout.stillRunning && (!rangeToMs || bout.end.getTime() <= rangeToMs);
      daySlices[daySlices.length - 1].stopReason = pauseInRange ? (bout.stopReason || "") : "";
      daySlices[daySlices.length - 1].stillRunning = still;
      if (still) daySlices[daySlices.length - 1].end = null;
    }
    for (var s = 0; s < daySlices.length; s++) slices.push(daySlices[s]);
  }
  return slices;
}

function getActivityPeriodBounds(period, ref) {
  var day0 = sastWallToDate(ref, 0, 0);
  if (period === "day") {
    return { from: day0, to: addSastDays(day0, 1) };
  }
  if (period === "month") {
    var ym = Utilities.formatDate(ref, TZ_JOBURG, "yyyy-MM").split("-");
    var y = Number(ym[0]);
    var m = Number(ym[1]);
    var from = sastWallToDate(new Date(Date.UTC(y, m - 1, 1, 10, 0, 0)), 0, 0);
    var next = m === 12 ? new Date(Date.UTC(y + 1, 0, 1, 10, 0, 0)) : new Date(Date.UTC(y, m, 1, 10, 0, 0));
    return { from: from, to: sastWallToDate(next, 0, 0) };
  }
  return { from: addSastDays(day0, -8), to: addSastDays(day0, 8) };
}

function hasOpenPause(pauses) {
  if (!pauses) return false;
  for (var i = 0; i < pauses.length; i++) {
    if (!pauses[i].end) return true;
  }
  return false;
}

function getOpenPauseStart(meta) {
  var pauses = (meta && meta.pauses) || [];
  for (var i = pauses.length - 1; i >= 0; i--) {
    if (!pauses[i].end) return pauses[i].start;
  }
  return null;
}

function addPauseToMeta(meta, reason) {
  meta = meta || defaultLogMeta();
  if (hasOpenPause(meta.pauses)) return meta;
  meta.pauses.push({
    start: new Date().getTime(),
    end: null,
    reason: reason || ""
  });
  return meta;
}

function isUserPauseReason(reason) {
  var r = String(reason || "").trim();
  if (r === "No materials") return true;
  if (/^Touch up #\S+/i.test(r)) return true;
  if (/^Other:\s+\S/.test(r)) return true;
  return false;
}

function listRunningOrdersForWorker(logs, workerName, exceptOrders) {
  var exceptMap = {};
  (exceptOrders || []).forEach(function(o) { exceptMap[String(o)] = true; });
  var running = [];
  var seen = {};
  for (var i = 1; i < logs.length; i++) {
    if (logs[i][6]) continue;
    if (String(logs[i][2] || "").trim() !== String(workerName || "").trim()) continue;
    var meta = parseLogMeta(logs[i].length > 12 ? logs[i][12] : "");
    if (meta.entryType === "indirect") continue;
    if (hasOpenPause(meta.pauses) || logs[i][9]) continue;
    var ord = String(logs[i][1] || "");
    if (!ord || exceptMap[ord] || seen[ord]) continue;
    seen[ord] = true;
    running.push({ order: ord, task: logs[i][4], role: logs[i][3] });
  }
  return running;
}

function closeOpenPauseInMeta(meta, atTime) {
  meta = meta || defaultLogMeta();
  var endMs = (atTime ? new Date(atTime) : new Date()).getTime();
  for (var i = meta.pauses.length - 1; i >= 0; i--) {
    if (!meta.pauses[i].end) {
      meta.pauses[i].end = endMs;
      break;
    }
  }
  return meta;
}

function cumulativePauseMinsFromMeta(meta, taskName) {
  var pauses = (meta && meta.pauses) || [];
  var total = 0;
  var now = new Date();
  for (var i = 0; i < pauses.length; i++) {
    if (!pauses[i].end) continue;
    total += calcRawServerMins(new Date(pauses[i].start), new Date(pauses[i].end), taskName);
  }
  return total;
}

function syncLegacyPauseCells(logSheet, rowNum, meta, taskName) {
  var openStart = getOpenPauseStart(meta);
  var lastReason = "";
  if (meta.pauses && meta.pauses.length) {
    lastReason = meta.pauses[meta.pauses.length - 1].reason || "";
  }
  if (openStart) {
    logSheet.getRange(rowNum, 10).setValue(new Date(openStart));
    logSheet.getRange(rowNum, 12).setValue(lastReason);
  } else {
    logSheet.getRange(rowNum, 10).clearContent();
    logSheet.getRange(rowNum, 12).clearContent();
  }
  logSheet.getRange(rowNum, 11).setValue(cumulativePauseMinsFromMeta(meta, taskName));
}

function findOrderRowByNumber(orderSheet, orderNum) {
  var last = orderSheet.getLastRow();
  if (last < 2) return -1;
  var col = orderSheet.getRange(2, 2, last - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    if (String(col[i][0]) === String(orderNum)) return i + 2;
  }
  return -1;
}

function findOpenLogRow(pack, logId, orderNum, workerName) {
  var logs = pack.values;
  var valuesIndex = -1;
  if (logId) {
    var normalizedLogId = String(logId).trim();
    for (var i = 1; i < logs.length; i++) {
      if (String(logs[i][0]).trim() === normalizedLogId) {
        valuesIndex = i;
        break;
      }
    }
  }
  if (valuesIndex !== -1) {
    var r = logs[valuesIndex];
    var matchesOrder = !orderNum || String(r[1]) === String(orderNum);
    var matchesWorker = !workerName || String(r[2]).trim() === String(workerName).trim();
    var isOpen = !r[6];
    if (!matchesOrder || !matchesWorker || !isOpen) valuesIndex = -1;
  }
  if (valuesIndex === -1 && orderNum && workerName) {
    for (var j = logs.length - 1; j >= 1; j--) {
      if (String(logs[j][1]) === String(orderNum) &&
          String(logs[j][2]).trim() === String(workerName).trim() &&
          !logs[j][6]) {
        valuesIndex = j;
        break;
      }
    }
  }
  return valuesIndex === -1 ? -1 : packSheetRow(pack, valuesIndex);
}

function closeIndirectTasksForWorker(ss, workerName, pack) {
  pack = pack || getLogPack(ss);
  var logSheet = getSheetOrDie(ss, TAB_LOGS);
  var logs = pack.values;
  var closed = 0;
  var now = new Date();
  for (var i = 1; i < logs.length; i++) {
    var meta = parseLogMeta(logs[i].length > 12 ? logs[i][12] : "");
    var isIndirect = meta.entryType === "indirect" || String(logs[i][3]).trim() === "Indirect";
    if (isIndirect && String(logs[i][2]).trim() === String(workerName).trim() && !logs[i][6]) {
      var sheetRow = packSheetRow(pack, i);
      meta = closeOpenPauseInMeta(meta, now);
      logSheet.getRange(sheetRow, 7).setValue(now);
      writeLogPauseState(logSheet, sheetRow, meta, logs[i][4]);
      closed++;
    }
  }
  return closed;
}

function autoPauseWorkerOtherJobs(ss, workerName, exceptOrders, reason, exceptBatchId, pack) {
  pack = pack || getLogPack(ss);
  var logSheet = getSheetOrDie(ss, TAB_LOGS);
  var logs = pack.values;
  var exceptMap = {};
  (exceptOrders || []).forEach(function(o) { exceptMap[String(o)] = true; });
  var paused = [];
  for (var i = 1; i < logs.length; i++) {
    var meta = parseLogMeta(logs[i].length > 12 ? logs[i][12] : "");
    if (meta.entryType === "indirect") continue;
    if (logs[i][6]) continue;
    if (String(logs[i][2]).trim() !== String(workerName).trim()) continue;
    if (exceptMap[String(logs[i][1])]) continue;
    if (exceptBatchId && meta.batchId && meta.batchId === exceptBatchId) continue;
    if (hasOpenPause(meta.pauses) || logs[i][9]) continue;
    var sheetRow = packSheetRow(pack, i);
    meta = addPauseToMeta(meta, reason || "Switched job");
    writeLogPauseState(logSheet, sheetRow, meta, logs[i][4]);
    paused.push(String(logs[i][1]));
  }
  return paused;
}

function resumeWorkerLog(ss, workerName, orderNum, pack) {
  pack = pack || getLogPack(ss);
  var logSheet = getSheetOrDie(ss, TAB_LOGS);
  var logs = pack.values;
  for (var i = logs.length - 1; i >= 1; i--) {
    if (String(logs[i][1]) === String(orderNum) &&
        String(logs[i][2]).trim() === String(workerName).trim() &&
        !logs[i][6]) {
      var sheetRow = packSheetRow(pack, i);
      var meta = parseLogMeta(logs[i].length > 12 ? logs[i][12] : "");
      meta = closeOpenPauseInMeta(meta, new Date());
      writeLogPauseState(logSheet, sheetRow, meta, logs[i][4]);
      return true;
    }
  }
  return false;
}

function autoSwitchWeldersAfterPlate(ss, orderNum) {
  var pack = getLogPack(ss);
  var logs = pack.values;
  var switched = [];
  for (var i = 1; i < logs.length; i++) {
    var role = String(logs[i][3]).trim();
    if (String(logs[i][1]) !== String(orderNum)) continue;
    if (role !== "Welding") continue;
    if (logs[i][6]) continue;
    var worker = String(logs[i][2]).trim();
    var meta = parseLogMeta(logs[i].length > 12 ? logs[i][12] : "");
    var isPaused = hasOpenPause(meta.pauses) || !!logs[i][9];
    if (!isPaused) continue;

    var pausedOthers = autoPauseWorkerOtherJobs(ss, worker, [orderNum], "Plate ready on " + orderNum, meta.batchId, pack);
    closeIndirectTasksForWorker(ss, worker, pack);
    resumeWorkerLog(ss, worker, orderNum, pack);
    var notice = {
      type: "plate-ready",
      worker: worker,
      resumedOrder: String(orderNum),
      pausedOrders: pausedOthers,
      message: "Plate ready — you are back on order " + orderNum
    };
    setWorkerNotice(worker, notice);
    switched.push(notice);
  }
  return switched;
}

function setWorkerNotice(worker, notice) {
  CacheService.getScriptCache().put("notice_" + String(worker).trim().toLowerCase(), JSON.stringify(notice), 900);
}

function popWorkerNotice(workerName) {
  var cache = CacheService.getScriptCache();
  var key = "notice_" + String(workerName).trim().toLowerCase();
  var v = cache.get(key);
  if (!v) return null;
  cache.remove(key);
  try { return JSON.parse(v); } catch (e) { return null; }
}

function undoAutoSwitch(workerName) {
  var cache = CacheService.getScriptCache();
  var key = "notice_" + String(workerName).trim().toLowerCase();
  var v = cache.get(key);
  if (!v) return { success: false, message: "No recent auto-switch to undo." };
  var notice;
  try { notice = JSON.parse(v); } catch (e) { return { success: false, message: "Invalid notice." }; }
  var ss = getSpreadsheet();
  autoPauseWorkerOtherJobs(ss, workerName, notice.pausedOrders || [], "Still on previous job", "");
  (notice.pausedOrders || []).forEach(function(ord) {
    resumeWorkerLog(ss, workerName, ord);
  });
  cache.remove(key);
  return { success: true };
}

function leaveBatchForOrder(workerName, keepOrder) {
  var ss = getSpreadsheet();
  var logSheet = getSheetOrDie(ss, TAB_LOGS);
  var pack = getLogPack(ss);
  var logs = pack.values;
  var now = new Date();
  var batchId = "";
  for (var i = logs.length - 1; i >= 1; i--) {
    if (String(logs[i][1]) === String(keepOrder) &&
        String(logs[i][2]).trim() === String(workerName).trim() &&
        !logs[i][6]) {
      var meta = parseLogMeta(logs[i].length > 12 ? logs[i][12] : "");
      batchId = meta.batchId;
      break;
    }
  }
  if (!batchId) return { success: true, message: "Not in a batch." };
  for (var j = 1; j < logs.length; j++) {
    var m = parseLogMeta(logs[j].length > 12 ? logs[j][12] : "");
    if (m.batchId !== batchId || logs[j][6]) continue;
    if (String(logs[j][2]).trim() !== String(workerName).trim()) continue;
    if (!m.batchSplitAt) m.batchSplitAt = now.getTime();
    if (String(logs[j][1]) !== String(keepOrder)) {
      m = addPauseToMeta(m, "Batch cut done — assembling " + keepOrder);
    }
    var sheetRow = packSheetRow(pack, j);
    writeLogPauseState(logSheet, sheetRow, m, logs[j][4]);
  }
  bumpFloorCache();
  return { success: true };
}

function ensureIdleTrigger() {
  try {
    if (CacheService.getScriptCache().get("idleTriggerOk")) return;
  } catch (e) {}
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "checkIdleWorkers") {
      try { CacheService.getScriptCache().put("idleTriggerOk", "1", 21600); } catch (e2) {}
      return;
    }
  }
  ScriptApp.newTrigger("checkIdleWorkers").timeBased().everyMinutes(5).create();
  try { CacheService.getScriptCache().put("idleTriggerOk", "1", 21600); } catch (e3) {}
}

function getIdleAlertSheet(ss) {
  var sheet = ss.getSheetByName(TAB_IDLE);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_IDLE);
    sheet.appendRow(["Date", "Worker", "Role", "IdleSince", "AlertedAt", "Status", "AssignedTask"]);
    sheet.hideSheet();
  }
  return sheet;
}

function workerHasRunningJob(logs, workerName) {
  for (var i = 1; i < logs.length; i++) {
    if (String(logs[i][2]).trim() !== String(workerName).trim()) continue;
    if (logs[i][6]) continue;
    var meta = parseLogMeta(logs[i].length > 12 ? logs[i][12] : "");
    if (meta.entryType === "indirect") continue;
    if (hasOpenPause(meta.pauses) || logs[i][9]) continue;
    return true;
  }
  return false;
}

function workerHasOpenIndirect(logs, workerName) {
  for (var i = 1; i < logs.length; i++) {
    if (String(logs[i][2]).trim() !== String(workerName).trim()) continue;
    if (logs[i][6]) continue;
    var meta = parseLogMeta(logs[i].length > 12 ? logs[i][12] : "");
    if (meta.entryType === "indirect" || String(logs[i][3]).trim() === "Indirect") return true;
  }
  return false;
}

function lastActivityMs(logs, workerName) {
  var latest = 0;
  for (var i = 1; i < logs.length; i++) {
    if (String(logs[i][2]).trim() !== String(workerName).trim()) continue;
    var start = logs[i][5] ? new Date(logs[i][5]).getTime() : 0;
    var end = logs[i][6] ? new Date(logs[i][6]).getTime() : 0;
    if (start > latest) latest = start;
    if (end > latest) latest = end;
  }
  return latest;
}

function alreadyAlertedToday(idleSheet, workerName) {
  var today = sastDayStamp(new Date());
  var data = idleSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(workerName).trim() &&
        String(data[i][0]) === today &&
        String(data[i][5]).toLowerCase() !== "resolved") {
      return true;
    }
  }
  return false;
}

function checkIdleWorkers() {
  if (!isWithinShiftNow()) return;
  var ss = getSpreadsheet();
  var logs = getLogPack(ss).values;
  var users = getUsersAndRoles();
  var idleSheet = getIdleAlertSheet(ss);
  var now = new Date();
  var graceMs = IDLE_GRACE_MINS * 60 * 1000;
  var alerted = [];

  for (var u = 0; u < users.length; u++) {
    var name = users[u].name;
    var role = users[u].role;
    if (!name || String(role).toLowerCase() === "admin") continue;
    if (workerHasRunningJob(logs, name)) continue;
    if (workerHasOpenIndirect(logs, name)) continue;
    var last = lastActivityMs(logs, name);
    if (last && (now.getTime() - last) < graceMs) continue;
    if (alreadyAlertedToday(idleSheet, name)) continue;

    idleSheet.appendRow([
      sastDayStamp(now),
      name,
      role,
      last ? new Date(last) : "",
      now,
      "Open",
      ""
    ]);
    alerted.push(name + " (" + role + ")");
  }
}

function getIdleWorkers() {
  var ss = getSpreadsheet();
  var idleSheet = getIdleAlertSheet(ss);
  var data = idleSheet.getDataRange().getValues();
  var today = sastDayStamp(new Date());
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === today && String(data[i][5]).toLowerCase() === "open") {
      list.push({
        row: i + 1,
        worker: data[i][1],
        role: data[i][2],
        idleSince: data[i][3] ? new Date(data[i][3]).getTime() : null,
        alertedAt: data[i][4] ? new Date(data[i][4]).getTime() : null
      });
    }
  }
  return { workers: list, tasks: INDIRECT_TASKS };
}

function heartbeatKey(name) {
  return "hb_" + String(name || "").trim().toLowerCase();
}

function userSeesIdleAlerts(workerName) {
  var profile = getUserProfileByName(workerName);
  if (!profile) return false;
  if (profile.isAdmin || profile.isQcOnly) return true;
  if (profile.tasks && profile.tasks.indexOf("Quality Control") !== -1) return true;
  var role = String(profile.role || "").toLowerCase();
  return role === "qc" || role === "quality control";
}

function markStaffHeartbeat(workerName) {
  if (!workerName || !userSeesIdleAlerts(workerName)) return;
  try {
    var cache = CacheService.getScriptCache();
    cache.put(heartbeatKey(workerName), JSON.stringify({ name: workerName, t: new Date().getTime() }), 120);
  } catch (e) {}
}

function pollIdleAlerts(workerName) {
  lazySetup();
  if (!workerName) return { ok: false, alerts: [], canAssign: false, tasks: INDIRECT_TASKS };
  markStaffHeartbeat(workerName);
  if (!userSeesIdleAlerts(workerName)) {
    return { ok: true, alerts: [], canAssign: false, tasks: INDIRECT_TASKS };
  }
  var data = getIdleWorkers();
  return { ok: true, alerts: data.workers || [], canAssign: true, tasks: data.tasks || INDIRECT_TASKS };
}

function assignIndirectTask(workerName, taskName, assignedBy) {
  if (!workerName || !taskName) return { success: false, message: "Worker and task are required." };
  var ss = getSpreadsheet();
  var logSheet = getSheetOrDie(ss, TAB_LOGS);
  var logs = getLogPack(ss).values;
  if (workerHasRunningJob(logs, workerName)) {
    return { success: false, message: workerName + " already has a running job." };
  }
  closeIndirectTasksForWorker(ss, workerName);

  var uniqueId = Utilities.getUuid();
  var meta = defaultLogMeta();
  meta.entryType = "indirect";
  logSheet.appendRow([
    uniqueId,
    "INDIRECT",
    workerName,
    "Indirect",
    taskName,
    new Date(),
    "",
    "",
    "",
    "",
    "",
    "",
    JSON.stringify(meta)
  ]);

  var idleSheet = getIdleAlertSheet(ss);
  var data = idleSheet.getDataRange().getValues();
  var today = sastDayStamp(new Date());
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(workerName).trim() &&
        String(data[i][0]) === today &&
        String(data[i][5]).toLowerCase() === "open") {
      idleSheet.getRange(i + 1, 6).setValue("Assigned");
      idleSheet.getRange(i + 1, 7).setValue(taskName);
    }
  }
  bumpFloorCache();
  return { success: true };
}

function getActivityReport(period, refDateMs, workerFilter) {
  var ref = refDateMs ? new Date(Number(refDateMs)) : new Date();
  var refStamp = sastDayStamp(ref);
  var refWeek = Utilities.formatDate(ref, TZ_JOBURG, "yyyy-'W'ww");
  var refMonth = Utilities.formatDate(ref, TZ_JOBURG, "yyyy-MM");
  var cacheKey = "activity:" + String(period || "day") + ":" + (period === "day" ? refStamp : (period === "week" ? refWeek : refMonth)) + ":" + String(workerFilter || "").trim().toLowerCase();
  var cached = floorCacheGet(cacheKey);
  if (cached) return cached;
  var ss = getSpreadsheet();
  var logSheet = getSheetOrDie(ss, TAB_LOGS);
  var lastRow = logSheet.getLastRow();
  var lastCol = Math.min(13, logSheet.getLastColumn());
  var logData = lastRow < 2 ? [[]] : logSheet.getRange(1, 1, lastRow, lastCol).getValues();
  var bounds = getActivityPeriodBounds(period, ref);
  var fromMs = bounds.from.getTime();
  var toMs = bounds.to.getTime();

  function inRange(dateObj) {
    if (!dateObj) return false;
    if (period === "day") return sastDayStamp(dateObj) === refStamp;
    if (period === "week") return Utilities.formatDate(dateObj, TZ_JOBURG, "yyyy-'W'ww") === refWeek;
    return Utilities.formatDate(dateObj, TZ_JOBURG, "yyyy-MM") === refMonth;
  }

  function inRangeStamp(stamp) {
    if (!stamp) return false;
    if (period === "day") return stamp === refStamp;
    if (period === "month") return String(stamp).slice(0, 7) === refMonth;
    var parts = String(stamp).split("-");
    if (parts.length < 3) return false;
    var noon = sastWallToDate(new Date(Date.UTC(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 10, 0, 0)), 12, 0);
    return inRange(noon);
  }

  var byWorker = {};
  var allWorkerNames = {};
  for (var i = 1; i < logData.length; i++) {
    var worker = String(logData[i][2] || "").trim();
    if (!worker) continue;
    var start = logData[i][5] ? new Date(logData[i][5]) : null;
    if (!start || isNaN(start.getTime())) continue;
    var end = logData[i][6] ? new Date(logData[i][6]) : new Date();
    if (end && isNaN(end.getTime())) end = new Date();
    if (start.getTime() >= toMs) continue;
    if (end.getTime() <= fromMs) continue;

    var slices = splitWorkByDay(logData[i], bounds.from, bounds.to);
    var anyInRange = false;
    for (var s = 0; s < slices.length; s++) {
      if (inRangeStamp(slices[s].dayStamp)) { anyInRange = true; break; }
    }
    if (!anyInRange) continue;
    allWorkerNames[worker] = true;
    if (workerFilter && String(workerFilter).trim() && String(workerFilter).trim().toLowerCase() !== worker.toLowerCase()) continue;

    var rowMeta = parseLogMeta(logData[i].length > 12 ? logData[i][12] : "");
    if (!byWorker[worker]) byWorker[worker] = { worker: worker, days: {}, tasks: [] };
    for (var s = 0; s < slices.length; s++) {
      var slice = slices[s];
      if (!inRangeStamp(slice.dayStamp)) continue;
      var dayStamp = slice.dayStamp;
      if (!byWorker[worker].days[dayStamp]) {
        byWorker[worker].days[dayStamp] = { date: dayStamp, minutes: 0, overtime: 0, regular: 0, tasks: [] };
      }
      var task = {
        order: logData[i][1],
        role: logData[i][3],
        task: logData[i][4],
        start: slice.start ? slice.start.getTime() : null,
        end: slice.end ? slice.end.getTime() : null,
        minutes: slice.mins,
        date: dayStamp,
        stopReason: slice.stopReason || "",
        entryType: rowMeta.entryType || "production"
      };
      byWorker[worker].days[dayStamp].minutes += slice.mins;
      byWorker[worker].days[dayStamp].tasks.push(task);
      byWorker[worker].tasks.push(task);
    }
  }

  var workers = [];
  var workerNames = Object.keys(byWorker).sort();
  for (var w = 0; w < workerNames.length; w++) {
    var rec = byWorker[workerNames[w]];
    var days = [];
    var dayKeys = Object.keys(rec.days).sort().reverse();
    var totalMins = 0;
    var totalOt = 0;
    for (var d = 0; d < dayKeys.length; d++) {
      var day = rec.days[dayKeys[d]];
      day.overtime = Math.max(0, day.minutes - STANDARD_DAY_MINS);
      day.regular = Math.min(day.minutes, STANDARD_DAY_MINS);
      day.overLimit = day.minutes > STANDARD_DAY_MINS;
      totalMins += day.minutes;
      totalOt += day.overtime;
      days.push(day);
    }
    rec.tasks.sort(function(a, b) { return (b.start || 0) - (a.start || 0); });
    workers.push({
      worker: rec.worker,
      days: days,
      tasks: rec.tasks,
      totalMinutes: totalMins,
      totalRegular: Math.max(0, totalMins - totalOt),
      totalOvertime: totalOt
    });
  }

  var result = {
    period: period,
    refDate: ref.getTime(),
    label: period === "day" ? refStamp : (period === "week" ? refWeek : refMonth),
    standardDayMins: STANDARD_DAY_MINS,
    allWorkerNames: Object.keys(allWorkerNames).sort(),
    workers: workers
  };
  floorCachePut(cacheKey, result, CACHE_TTL_ACTIVITY);
  return result;
}
function getScheduleSheet(ss) {
  ss = ss || getSpreadsheet();
  var sheet = ss.getSheetByName(TAB_SCHEDULE);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_SCHEDULE);
    sheet.appendRow(["Id", "Worker", "Process", "Order", "Product", "Title", "Start", "End", "DurationMins", "Kind", "Seq", "EstimateSource"]);
  }
  return sheet;
}

function normScheduleKey(s) {
  return String(s || "").trim().toLowerCase();
}

function nextWorkInstant(date) {
  var cursor = new Date(date.getTime());
  var safety = 0;
  while (safety++ < 30) {
    var dow = sastDayOfWeek(cursor);
    if (dow === 0) {
      cursor = sastWallToDate(addSastDays(cursor, 1), 7, 45);
      continue;
    }
    if (dow === 6) {
      cursor = sastWallToDate(addSastDays(cursor, 2), 7, 45);
      continue;
    }
    var mins = sastMinsOfDay(cursor);
    if (mins < SHIFT_START_MINS) return sastWallToDate(cursor, 7, 45);
    if (mins >= LUNCH_START_MINS && mins < LUNCH_END_MINS) return sastWallToDate(cursor, 12, 30);
    if (mins >= SHIFT_END_MINS) {
      cursor = sastWallToDate(addSastDays(cursor, 1), 7, 45);
      continue;
    }
    return cursor;
  }
  return date;
}

function workWindowEnd(date) {
  var mins = sastMinsOfDay(date);
  if (mins < LUNCH_START_MINS) return sastWallToDate(date, 12, 0);
  return sastWallToDate(date, 15, 45);
}

function layOutWorkingDuration(startDate, durationMins, meta) {
  var remaining = Math.max(1, Math.round(Number(durationMins) || 0));
  var cursor = nextWorkInstant(startDate);
  var segs = [];
  var safety = 0;
  while (remaining > 0 && safety++ < 80) {
    cursor = nextWorkInstant(cursor);
    var winEnd = workWindowEnd(cursor);
    var avail = (winEnd.getTime() - cursor.getTime()) / 60000;
    if (avail < 0.5) {
      cursor = new Date(winEnd.getTime() + 60000);
      continue;
    }
    var take = Math.min(remaining, avail);
    var end = new Date(cursor.getTime() + take * 60000);
    segs.push({
      worker: meta.worker,
      process: meta.process || "",
      order: meta.order || "",
      product: meta.product || "",
      title: meta.title || "",
      kind: meta.kind || "order",
      seq: meta.seq || 0,
      estimateSource: meta.estimateSource || "",
      start: cursor,
      end: end,
      durationMins: take
    });
    remaining -= take;
    cursor = end;
  }
  return segs;
}

function getHistoricalProcessAverages() {
  var cached = floorCacheGet("schedAvg");
  if (cached) return cached;
  var ss = getSpreadsheet();
  var ordersSheet = getSheetOrDie(ss, TAB_ORDERS);
  var ordersData = ordersSheet.getDataRange().getValues();
  var orderProducts = {};
  for (var i = 1; i < ordersData.length; i++) {
    var on = String(ordersData[i][1] || "").trim();
    var pn = String(ordersData[i][6] || "").trim();
    if (on && pn) orderProducts[on] = pn;
  }
  var logSheet = getSheetOrDie(ss, TAB_LOGS);
  var lastRow = logSheet.getLastRow();
  var lastCol = Math.min(13, logSheet.getLastColumn());
  var logData = lastRow < 2 ? [[]] : logSheet.getRange(1, 1, lastRow, lastCol).getValues();
  var byProduct = {};
  var byProcess = {};
  function addAvg(map, key, mins) {
    if (!key || !(mins > 0)) return;
    if (!map[key]) map[key] = { total: 0, count: 0 };
    map[key].total += mins;
    map[key].count += 1;
  }
  for (var r = 1; r < logData.length; r++) {
    if (!logData[r][6]) continue;
    var meta = parseLogMeta(logData[r].length > 12 ? logData[r][12] : "");
    if (meta.entryType === "indirect") continue;
    var processName = String(logData[r][4] || "").trim();
    if (!processName) continue;
    var mins = calculateWorkMinutesFromLog(logData[r]);
    if (!(mins > 0)) continue;
    if (mins > STANDARD_DAY_MINS * 5) mins = STANDARD_DAY_MINS;
    addAvg(byProcess, normScheduleKey(processName), mins);
    var product = orderProducts[String(logData[r][1] || "").trim()] || "";
    if (product) addAvg(byProduct, normScheduleKey(product) + "||" + normScheduleKey(processName), mins);
  }
  var result = { byProduct: byProduct, byProcess: byProcess };
  floorCachePut("schedAvg", result, 300);
  return result;
}

function estimateScheduleMinutes(product, process, averages) {
  averages = averages || getHistoricalProcessAverages();
  var pk = normScheduleKey(product) + "||" + normScheduleKey(process);
  var prod = averages.byProduct[pk];
  if (prod && prod.count > 0) {
    return {
      mins: Math.max(15, Math.round(prod.total / prod.count)),
      source: "Same product × " + prod.count
    };
  }
  var proc = averages.byProcess[normScheduleKey(process)];
  if (proc && proc.count > 0) {
    return {
      mins: Math.max(15, Math.round(proc.total / proc.count)),
      source: "Typical " + process + " × " + proc.count
    };
  }
  return { mins: 120, source: "Default 2h (no history yet)" };
}

function readScheduleBlocks(ss) {
  var sheet = getScheduleSheet(ss);
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var values = sheet.getRange(2, 1, last - 1, 12).getValues();
  var blocks = [];
  for (var i = 0; i < values.length; i++) {
    var start = values[i][6] ? new Date(values[i][6]) : null;
    var end = values[i][7] ? new Date(values[i][7]) : null;
    if (start && isNaN(start.getTime())) start = null;
    if (end && isNaN(end.getTime())) end = null;
    if (!start || !end) continue;
    blocks.push({
      id: String(values[i][0] || ""),
      worker: String(values[i][1] || ""),
      process: String(values[i][2] || ""),
      order: String(values[i][3] || ""),
      product: String(values[i][4] || ""),
      title: String(values[i][5] || ""),
      start: start,
      end: end,
      startMs: start.getTime(),
      endMs: end.getTime(),
      durationMins: Number(values[i][8]) || Math.round((end.getTime() - start.getTime()) / 60000),
      kind: String(values[i][9] || "order"),
      seq: Number(values[i][10]) || 0,
      estimateSource: String(values[i][11] || ""),
      sheetRow: i + 2
    });
  }
  return blocks;
}

function writeScheduleBlocks(ss, workerName, fromMs, newBlocks) {
  var sheet = getScheduleSheet(ss);
  var existing = readScheduleBlocks(ss);
  var keep = [];
  var toClear = [];
  for (var i = 0; i < existing.length; i++) {
    var b = existing[i];
    if (String(b.worker) !== String(workerName)) {
      keep.push(b);
      continue;
    }
    if (b.endMs <= fromMs) {
      keep.push(b);
      continue;
    }
    toClear.push(b.sheetRow);
  }
  toClear.sort(function(a, b) { return b - a; });
  for (var c = 0; c < toClear.length; c++) {
    sheet.deleteRow(toClear[c]);
  }
  if (!newBlocks.length) return keep;
  var rows = [];
  for (var n = 0; n < newBlocks.length; n++) {
    var bl = newBlocks[n];
    rows.push([
      Utilities.getUuid(),
      workerName,
      bl.process || "",
      bl.order || "",
      bl.product || "",
      bl.title || "",
      bl.start,
      bl.end,
      bl.durationMins || 0,
      bl.kind || "order",
      bl.seq || (n + 1),
      bl.estimateSource || ""
    ]);
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
  return keep.concat(newBlocks);
}

function collapseBlocksToItems(blocks) {
  var items = [];
  var cur = null;
  function flush() {
    if (cur) items.push(cur);
    cur = null;
  }
  for (var i = 0; i < blocks.length; i++) {
    var b = blocks[i];
    var key = b.kind + "|" + (b.order || b.title || "");
    if (cur && cur.key === key && b.kind === "order") {
      cur.durationMins += b.durationMins;
      continue;
    }
    flush();
    cur = {
      key: key,
      worker: b.worker,
      process: b.process,
      order: b.order,
      product: b.product,
      title: b.title,
      kind: b.kind,
      seq: b.seq,
      estimateSource: b.estimateSource,
      durationMins: b.durationMins,
      pinnedStart: b.kind === "other" ? b.start : null
    };
  }
  flush();
  return items;
}

function packItemsFrom(fromMs, items) {
  var cursor = nextWorkInstant(new Date(fromMs));
  var segs = [];
  var pinned = [];
  var floating = [];
  for (var i = 0; i < items.length; i++) {
    if (items[i].kind === "other" && items[i].pinnedStart) pinned.push(items[i]);
    else floating.push(items[i]);
  }
  pinned.sort(function(a, b) { return a.pinnedStart.getTime() - b.pinnedStart.getTime(); });
  var p = 0;
  var f = 0;
  function placeOne(item, startAt) {
    var laid = layOutWorkingDuration(startAt, item.durationMins, item);
    for (var s = 0; s < laid.length; s++) segs.push(laid[s]);
    if (laid.length) cursor = laid[laid.length - 1].end;
    else cursor = startAt;
  }
  while (p < pinned.length || f < floating.length) {
    var pin = p < pinned.length ? pinned[p] : null;
    if (pin) {
      var pinAt = nextWorkInstant(pin.pinnedStart);
      while (f < floating.length && nextWorkInstant(cursor).getTime() + 30000 < pinAt.getTime()) {
        var beforePin = nextWorkInstant(cursor);
        if (beforePin.getTime() >= pinAt.getTime()) break;
        var item = floating[f];
        var winEnd = workWindowEnd(beforePin);
        var room = Math.min(
          (winEnd.getTime() - beforePin.getTime()) / 60000,
          (pinAt.getTime() - beforePin.getTime()) / 60000
        );
        if (room < 0.5) {
          cursor = pinAt;
          break;
        }
        var take = Math.min(item.durationMins, room);
        if (take >= item.durationMins - 0.5) {
          placeOne(item, beforePin);
          f++;
        } else {
          var part = {};
          for (var k in item) part[k] = item[k];
          part.durationMins = take;
          placeOne(part, beforePin);
          item.durationMins -= take;
        }
      }
      var pinStart = pin.pinnedStart;
      if (pinStart.getTime() < cursor.getTime()) pinStart = cursor;
      placeOne(pin, pinStart);
      p++;
    } else {
      placeOne(floating[f], cursor);
      f++;
    }
  }
  return segs;
}

function getScheduleBoard() {
  var ss = getSpreadsheet();
  var orderData = getSheetGrid(ss, TAB_ORDERS, 7);
  var orders = [];
  for (var i = 1; i < orderData.length; i++) {
    var status = String(orderData[i][2] || "").trim();
    var lower = status.toLowerCase();
    if (!isAllowedStatus(status)) continue;
    if (lower !== "not yet started" && lower !== "ready for steelwork") continue;
    orders.push({
      rowIndex: i + 1,
      order: orderData[i][1],
      status: status,
      productName: orderData[i][6] || ""
    });
  }
  var users = getUsersAndRoles();
  var workers = [];
  for (var u = 0; u < users.length; u++) {
    if (!users[u].name || users[u].isAdmin) continue;
    workers.push({
      name: users[u].name,
      tasks: users[u].tasks || [],
      jobTitle: users[u].jobTitle || users[u].role || ""
    });
  }
  var averages = getHistoricalProcessAverages();
  var blocks = readScheduleBlocks(ss).map(function(b) {
    return {
      id: b.id,
      worker: b.worker,
      process: b.process,
      order: b.order,
      product: b.product,
      title: b.title,
      start: b.start.getTime(),
      end: b.end.getTime(),
      durationMins: b.durationMins,
      kind: b.kind,
      seq: b.seq,
      estimateSource: b.estimateSource,
      date: sastDayStamp(b.start)
    };
  });
  return {
    orders: orders,
    workers: workers,
    blocks: blocks,
    processes: ["Profile Cutting", "Plate Cutting", "Tagging", "Welding", "Grinding", "Paint Preparation", "Painting", "Assembly"],
    indirectTasks: INDIRECT_TASKS.slice(),
    averages: averages
  };
}

function generateWorkerSchedule(workerName, processName, orderNums, fromDateMs) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    workerName = String(workerName || "").trim();
    processName = String(processName || "Profile Cutting").trim();
    if (!workerName) return { success: false, message: "Pick a worker." };
    var ss = getSpreadsheet();
    var fromMs = fromDateMs ? Number(fromDateMs) : new Date().getTime();
    var orderData = getSheetGrid(ss, TAB_ORDERS, 7);
    var byNum = {};
    for (var i = 1; i < orderData.length; i++) {
      byNum[String(orderData[i][1]).trim()] = {
        order: orderData[i][1],
        product: orderData[i][6] || "",
        status: orderData[i][2] || ""
      };
    }
    var averages = getHistoricalProcessAverages();
    var existing = readScheduleBlocks(ss).filter(function(b) {
      return String(b.worker) === workerName && b.endMs > fromMs;
    });
    existing.sort(function(a, b) { return a.startMs - b.startMs; });
    var pinnedOthers = collapseBlocksToItems(existing.filter(function(b) { return b.kind === "other"; }));
    var floating = [];
    for (var o = 0; o < (orderNums || []).length; o++) {
      var num = String(orderNums[o] || "").trim();
      var rec = byNum[num];
      if (!rec) continue;
      var est = estimateScheduleMinutes(rec.product, processName, averages);
      floating.push({
        worker: workerName,
        process: processName,
        order: rec.order,
        product: rec.product,
        title: rec.product ? (rec.order + " · " + rec.product) : rec.order,
        kind: "order",
        seq: o + 1,
        estimateSource: est.source,
        durationMins: est.mins,
        pinnedStart: null
      });
    }
    var items = pinnedOthers.concat(floating);
    var segs = packItemsFrom(fromMs, items);
    writeScheduleBlocks(ss, workerName, fromMs, segs);
    return { success: true, blocks: segs.length };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function insertScheduleTask(workerName, title, durationMins, startMs, processName) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    workerName = String(workerName || "").trim();
    title = String(title || "").trim();
    durationMins = Math.max(15, Math.round(Number(durationMins) || 0));
    if (!workerName || !title) return { success: false, message: "Worker and task name are required." };
    var ss = getSpreadsheet();
    var fromMs = startMs ? Number(startMs) : new Date().getTime();
    var existing = readScheduleBlocks(ss).filter(function(b) {
      return String(b.worker) === workerName;
    });
    existing.sort(function(a, b) { return a.startMs - b.startMs; });
    var keepPrefix = [];
    var later = [];
    for (var i = 0; i < existing.length; i++) {
      var b = existing[i];
      if (b.endMs <= fromMs) {
        keepPrefix.push(b);
        continue;
      }
      if (b.startMs < fromMs && b.endMs > fromMs) {
        var doneMins = Math.max(0, Math.round((fromMs - b.startMs) / 60000));
        var left = Math.max(15, b.durationMins - doneMins);
        later.push({
          worker: workerName,
          process: b.process,
          order: b.order,
          product: b.product,
          title: b.title,
          kind: b.kind,
          seq: b.seq,
          estimateSource: b.estimateSource,
          durationMins: left,
          pinnedStart: b.kind === "other" ? new Date(fromMs + durationMins * 60000) : null
        });
        continue;
      }
      later.push({
        worker: workerName,
        process: b.process,
        order: b.order,
        product: b.product,
        title: b.title,
        kind: b.kind,
        seq: b.seq,
        estimateSource: b.estimateSource,
        durationMins: b.durationMins,
        pinnedStart: b.kind === "other" ? b.start : null
      });
    }
    var inserted = {
      worker: workerName,
      process: processName || "",
      order: "",
      product: "",
      title: title,
      kind: "other",
      seq: 0,
      estimateSource: "Inserted",
      durationMins: durationMins,
      pinnedStart: new Date(fromMs)
    };
    var items = [inserted].concat(later);
    var segs = packItemsFrom(fromMs, items);
    writeScheduleBlocks(ss, workerName, fromMs, segs);
    return { success: true, blocks: segs.length };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function clearWorkerScheduleFrom(workerName, fromDateMs) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = getSpreadsheet();
    var fromMs = fromDateMs ? Number(fromDateMs) : new Date().getTime();
    writeScheduleBlocks(ss, String(workerName || "").trim(), fromMs, []);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
