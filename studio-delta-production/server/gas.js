const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { spawnSync } = require("child_process");
const crypto = require("crypto");
const { google } = require("googleapis");
const { Spreadsheet, createSpreadsheetApp } = require("./sheets");

const TZ = process.env.TZ || "Africa/Johannesburg";
const SAST_OFFSET_MS = 2 * 60 * 60 * 1000;

function getAuth() {
  const scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"];
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    return new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes
    });
  }
  if (process.env.GOOGLE_APPLICATION_CREDENTIALS) {
    return new google.auth.GoogleAuth({ keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS, scopes });
  }
  throw new Error("Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_APPLICATION_CREDENTIALS");
}

function driveRpc(payload) {
  const r = spawnSync(process.execPath, [path.join(__dirname, "drive-cli.js")], {
    input: JSON.stringify(payload),
    encoding: "utf8",
    env: process.env,
    maxBuffer: 40 * 1024 * 1024
  });
  if (!r.stdout) throw new Error(r.stderr || "Drive helper failed");
  const parsed = JSON.parse(r.stdout);
  if (!parsed.ok) throw new Error(parsed.error || "Drive helper error");
  return parsed;
}

const cacheStore = new Map();
function cacheGet(key) {
  const hit = cacheStore.get(key);
  if (!hit) return null;
  if (hit.exp && hit.exp < Date.now()) {
    cacheStore.delete(key);
    return null;
  }
  return hit.val;
}
function cachePut(key, val, ttlSec) {
  cacheStore.set(key, { val, exp: ttlSec ? Date.now() + ttlSec * 1000 : 0 });
}

function isoWeek(date) {
  const sast = new Date(date.getTime() + SAST_OFFSET_MS);
  const utc = Date.UTC(sast.getUTCFullYear(), sast.getUTCMonth(), sast.getUTCDate());
  const d = new Date(utc);
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = Date.UTC(d.getUTCFullYear(), 0, 1);
  return Math.ceil(((d.getTime() - yearStart) / 86400000 + 1) / 7);
}

function formatDate(date, _tz, pattern) {
  const sast = new Date(date.getTime() + SAST_OFFSET_MS);
  const y = String(sast.getUTCFullYear());
  const m = String(sast.getUTCMonth() + 1).padStart(2, "0");
  const d = String(sast.getUTCDate()).padStart(2, "0");
  const ww = String(isoWeek(date)).padStart(2, "0");
  return String(pattern)
    .replace(/yyyy/g, y)
    .replace(/MM/g, m)
    .replace(/dd/g, d)
    .replace(/ww/g, ww)
    .replace(/'Week'/g, "Week")
    .replace(/'W'/g, "W");
}

function fileIterator(files) {
  let i = 0;
  return {
    hasNext: () => i < files.length,
    next: () => files[i++]
  };
}

function makeFolder(id) {
  return {
    getId: () => id,
    createFile: function (a, b, c) {
      if (typeof a === "string" && b != null) {
        const r = driveRpc({ op: "createFile", folderId: id, name: a, content: String(b), mimeType: c || "text/plain" });
        return makeFile(r);
      }
      const blob = a;
      const payload = {
        op: "createFile",
        folderId: id,
        name: blob.name || "file",
        mimeType: blob.contentType || "application/octet-stream"
      };
      if (blob._base64) payload.contentBase64 = blob._base64;
      else payload.content = blob._text || "";
      return makeFile(driveRpc(payload));
    },
    getFilesByType: function (mime) {
      const r = driveRpc({ op: "listFiles", folderId: id, mimeType: mime });
      return fileIterator((r.files || []).map(makeFile));
    },
    getFiles: function () {
      const r = driveRpc({ op: "listFiles", folderId: id });
      return fileIterator((r.files || []).map(makeFile));
    },
    setSharing: function () {}
  };
}

function makeFile(info) {
  const file = {
    getId: () => info.id,
    getUrl: () => info.url || info.webViewLink || "",
    getName: () => info.name,
    getMimeType: () => info.mimeType || "",
    getDateCreated: () => (info.createdTime ? new Date(info.createdTime) : new Date(0)),
    setName: function (name) {
      driveRpc({ op: "renameFile", fileId: info.id, name });
      info.name = name;
      return this;
    },
    setTrashed: function () {
      driveRpc({ op: "trashFile", fileId: info.id });
    },
    makeCopy: function (name, folder) {
      const r = driveRpc({ op: "copyFile", fileId: info.id, name, folderId: folder && folder.getId && folder.getId() });
      return makeFile(r);
    },
    getAs: function (mime) {
      if (String(mime).indexOf("pdf") === -1) {
        const r = driveRpc({ op: "getFileText", fileId: info.id });
        return { name: info.name || "file", _text: r.text || "", contentType: "text/plain", setName: function (n) { this.name = n; return this; } };
      }
      const r = driveRpc({ op: "exportPdf", fileId: info.id });
      return {
        name: (info.name || "file") + ".pdf",
        contentType: "application/pdf",
        _base64: r.pdfBase64,
        setName: function (n) { this.name = n; return this; }
      };
    },
    getBlob: function () {
      const r = driveRpc({ op: "getFileText", fileId: info.id });
      return {
        getDataAsString: () => r.text || "",
        setName: function (n) { this.name = n; return this; }
      };
    }
  };
  return file;
}

function makeBlob(bytes, contentType) {
  const buf = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes || []);
  return {
    contentType: contentType || "application/octet-stream",
    name: "file",
    _base64: buf.toString("base64"),
    setName: function (n) { this.name = n; return this; },
    getBytes: () => buf
  };
}

function buildDriveApp() {
  return {
    Access: { ANYONE_WITH_LINK: "ANYONE_WITH_LINK" },
    Permission: { VIEW: "VIEW" },
    getFileById: (id) => makeFile({ id }),
    getFolderById: (id) => makeFolder(id),
    getFoldersByName: (name) => {
      const r = driveRpc({ op: "listFoldersByName", name });
      const folders = (r.files || []).map((f) => makeFolder(f.id));
      let i = 0;
      return { hasNext: () => i < folders.length, next: () => folders[i++] };
    },
    createFolder: (name) => {
      const r = driveRpc({ op: "createFolder", name });
      return makeFolder(r.id);
    }
  };
}

function buildDocumentApp() {
  return {
    openById: (id) => ({
      getId: () => id,
      getBody: () => ({
        replaceText: (tag, value) => {
          driveRpc({ op: "replaceText", docId: id, replacements: [{ tag, value: String(value == null ? "" : value) }] });
        },
        findText: () => null,
        insertPageBreak: () => {},
        getChildIndex: () => 0
      }),
      saveAndClose: () => {}
    })
  };
}

function buildHtmlService() {
  return {
    createHtmlOutput: (html) => ({
      getAs: function (mime) {
        if (String(mime).indexOf("pdf") === -1) {
          return { setName: function (n) { this.name = n; return this; }, name: "file", _text: html, contentType: "text/html" };
        }
        try {
          const r = driveRpc({ op: "htmlToPdf", html, name: "powder-list" });
          return {
            name: "file.pdf",
            contentType: "application/pdf",
            _base64: r.pdfBase64,
            setName: function (n) { this.name = n; return this; }
          };
        } catch (e) {
          return { setName: function (n) { this.name = n; return this; }, name: "file.html", _text: html, contentType: "text/html" };
        }
      }
    }),
    createTemplateFromFile: () => ({ evaluate: () => ({ setTitle: function () { return this; }, setXFrameOptionsMode: function () { return this; }, addMetaTag: function () { return this; } }) }),
    XFrameOptionsMode: { ALLOWALL: "ALLOWALL" }
  };
}

function buildMailApp() {
  return {
    sendEmail: function (opts) {
      const attachments = (opts.attachments || []).map((b) => ({
        name: b.name || "file.pdf",
        mime: b.contentType || "application/pdf",
        base64: b._base64 || Buffer.from(b._text || "", "utf8").toString("base64")
      }));
      driveRpc({
        op: "sendMail",
        to: opts.to,
        cc: opts.cc || "",
        subject: opts.subject || "",
        html: opts.htmlBody || opts.body || "",
        attachments
      });
    }
  };
}

const ALLOWED = new Set([
  "getSteelProfiles", "getBackboards", "getUsersAndRoles", "verifyGlobalLogin", "verifyLogin",
  "pollFloor", "startOrder", "finishOrder", "workerPauseOrder", "workerResumeOrder",
  "batchStartOrders", "batchFinishOrders", "reportScratchedGlass", "getWeldingOrders", "logWelderSteel",
  "getAdminDashboardData", "adminPauseOrder", "adminResumeOrder",
  "getOrderMetrics", "getProductionTrendsData", "getWeeklyAnalyticsData",
  "generatePowderCoatingList", "getQCReportsFast", "processPdfQueue",
  "undoAutoSwitch", "leaveBatchForOrder", "getIdleWorkers", "pollIdleAlerts", "assignIndirectTask",
  "getActivityReport", "getScheduleBoard", "generateWorkerSchedule", "insertScheduleTask", "clearWorkerScheduleFrom",
  "checkIdleWorkers", "lazySetup"
]);

let scriptSource = null;
function loadScript() {
  if (!scriptSource) {
    scriptSource = fs.readFileSync(path.join(__dirname, "..", "Code.gs"), "utf8");
  }
  return scriptSource;
}

const SHEET_CACHE_MS = Number(process.env.SHEET_CACHE_MS || 30000);
let workbookCache = null;

async function getCachedWorkbook() {
  const sheetId = process.env.SHEET_ID;
  if (!sheetId) throw new Error("SHEET_ID is not set");
  const now = Date.now();
  if (
    workbookCache &&
    workbookCache.spreadsheetId === sheetId &&
    now - workbookCache.loadedAt < SHEET_CACHE_MS
  ) {
    return workbookCache.book;
  }
  const auth = getAuth();
  const client = await auth.getClient();
  const book = new Spreadsheet(client, sheetId);
  await book.load();
  workbookCache = { book, spreadsheetId: sheetId, loadedAt: now };
  return book;
}

function jsonSafe(value) {
  if (value instanceof Date) return value.getTime();
  if (Array.isArray(value)) return value.map(jsonSafe);
  if (value && typeof value === "object") {
    const out = {};
    for (const key of Object.keys(value)) out[key] = jsonSafe(value[key]);
    return out;
  }
  return value;
}

async function callShopFunction(fnName, args) {
  if (!/^[A-Za-z0-9_]+$/.test(fnName) || !ALLOWED.has(fnName)) {
    throw new Error("Unknown function: " + fnName);
  }
  const workbook = await getCachedWorkbook();

  const sandbox = {
    SpreadsheetApp: createSpreadsheetApp(workbook),
    CacheService: {
      getScriptCache: () => ({
        get: (k) => cacheGet(k),
        put: (k, v, ttl) => cachePut(k, v, ttl)
      })
    },
    LockService: {
      getScriptLock: () => ({
        waitLock: () => {},
        tryLock: () => true,
        releaseLock: () => {}
      })
    },
    Utilities: {
      formatDate,
      getUuid: () => crypto.randomUUID(),
      newBlob: makeBlob,
      base64Decode: (s) => Buffer.from(String(s).replace(/^data:[^;]+;base64,/, ""), "base64")
    },
    DriveApp: buildDriveApp(),
    DocumentApp: buildDocumentApp(),
    HtmlService: buildHtmlService(),
    MailApp: buildMailApp(),
    ScriptApp: {
      getProjectTriggers: () => [],
      newTrigger: () => ({ timeBased: () => ({ everyMinutes: () => ({ create: () => {} }) }) })
    },
    Logger: { log: (...a) => console.log("[shop]", ...a) },
    MimeType: { PLAIN_TEXT: "text/plain", PDF: "application/pdf" },
    console,
    Date,
    Math,
    JSON,
    String,
    Number,
    Array,
    Object,
    parseInt,
    parseFloat,
    isNaN,
    encodeURIComponent,
    decodeURIComponent,
    Error,
    Buffer
  };

  const ctx = vm.createContext(sandbox);
  vm.runInContext(loadScript() + "\nthis.__fn = " + fnName + ";", ctx);
  ctx.getSpreadsheet = function () {
    if (!ctx._ssMemo) ctx._ssMemo = workbook;
    return workbook;
  };
  ctx._ssMemo = workbook;
  ctx._sheetMemo = {};
  ctx._gridMemo = {};
  ctx._logPackMemo = null;
  ctx._logPackFull = false;

  const fn = ctx.__fn;
  if (typeof fn !== "function") throw new Error("Function not found: " + fnName);
  let result;
  try {
    result = fn.apply(null, args || []);
  } catch (e) {
    await workbook.flush().catch(() => {});
    throw e;
  }
  await workbook.flush();
  if (workbookCache && workbookCache.book === workbook) {
    workbookCache.loadedAt = Date.now();
  }
  return jsonSafe(result);
}

module.exports = { callShopFunction, ALLOWED, jsonSafe, getCachedWorkbook };
