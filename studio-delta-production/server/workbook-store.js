const fs = require("fs");
const path = require("path");
const { Spreadsheet } = require("./sheets");

const ORDER_HEADERS = [
  "QUOTE NUMBER", "ORDER NUMBER", "STATUS", "ASSIGNED OPERATOR", "TYPE", "CATERGORY",
  "PRODUCT", "VARIATION", "DOORS", "DETAILED DESCRIPTION", "DIMENSIONS", "POWDER COATING",
  "CLIENT NAME", "CLIENT NUMBER", "EMAIL ADDRESS", "PAYMENT DATE", "ADDRESS", "PROVINCE",
  "PRICE (Excl VAT)", "PRICE (Incl VAT)", "AMOUNT PAID", "MONTH OF SALE", "SOURCE", "CITY"
];

const SEED_TABS = {
  Users: [["Name", "Role", "Password", "Tasks", "Access", "See Debtors"]],
  ORDERS: [ORDER_HEADERS.slice()],
  Production_Log: [[
    "ID", "Order #", "Worker", "Process", "Status", "Start", "End", "Result", "Signature",
    "Pause Start", "Pause Mins", "Pause Reason", "Meta"
  ]],
  Overview: [["ID", "Order #", "Worker", "Status", "Start", "End", "Duration"]],
  Rates: [["Process", "Rate"]],
  Steel_Profiles: [["Category", "Profile Name"]],
  Steel_Usage: [["Timestamp", "Order #", "Worker", "Process", "Profile Type", "Size / Length"]],
  Backboards: [["Category", "Profile Name"]],
  Backboard_Usage: [["Timestamp", "Order", "Worker", "Process", "Type", "Size"]],
  Idle_Alerts: [["Date", "Worker", "Idle Since", "Minutes", "Noted", "Status", "Task"]],
  Schedule: [["Worker", "Block Start", "Block End", "Process", "Order", "Notes"]],
  Task_Durations: [["Product", "Process", "Minutes"]]
};

function onRailway() {
  return !!(process.env.RAILWAY_ENVIRONMENT || process.env.RAILWAY_PROJECT_ID || process.env.RAILWAY_SERVICE_ID);
}

function dataDir() {
  const volume = String(process.env.RAILWAY_VOLUME_MOUNT_PATH || "").trim();
  const configured = String(process.env.DATA_DIR || "").trim();
  // On Railway the volume is the durable disk. Docker also sets DATA_DIR=/app/data;
  // if those paths differ, writing to DATA_DIR would be wiped on every deploy.
  const dir = (onRailway() && volume) || configured || path.join(__dirname, "..", "data");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function storageInfo() {
  const dir = dataDir();
  const volume = String(process.env.RAILWAY_VOLUME_MOUNT_PATH || "").trim();
  const dataDirEnv = String(process.env.DATA_DIR || "").trim();
  const railway = onRailway();
  const ephemeral = railway && !volume;
  let warning = null;
  if (ephemeral) {
    warning = "No Railway volume is attached. Files in " + dir + " are wiped on every deploy. Add a Volume mounted at /app/data.";
  }
  const bookFile = workbookPath();
  return {
    database: "JSON files on disk (not Google Sheets, not Postgres)",
    dataDir: dir,
    volumeMount: volume || null,
    dataDirEnv: dataDirEnv || null,
    onRailway: railway,
    usingEphemeralDisk: ephemeral,
    warning,
    workbook: bookFile,
    workbookExists: fs.existsSync(bookFile)
  };
}

function workbookPath() {
  return path.join(dataDir(), "floor-workbook.json");
}

let book = null;

function writeBookFile(target) {
  const file = workbookPath();
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(target.toJSON()));
  fs.renameSync(tmp, file);
}

function attachPersist(target) {
  target.spreadsheetId = "railway-local";
  target.onFlush = function () {
    writeBookFile(target);
  };
  return target;
}

function seedMissingTabs(target) {
  Object.keys(SEED_TABS).forEach((title) => {
    if (target.getSheetByName(title)) return;
    const sheet = target.insertSheet(title);
    sheet.getRange(1, 1, 1, SEED_TABS[title][0].length).setValues(SEED_TABS[title]);
  });
}

function createEmptyBook() {
  const target = attachPersist(new Spreadsheet(null, "railway-local"));
  seedMissingTabs(target);
  writeBookFile(target);
  return target;
}

function initWorkbook() {
  if (book) return book;
  const file = workbookPath();
  if (fs.existsSync(file)) {
    try {
      const data = JSON.parse(fs.readFileSync(file, "utf8"));
      book = attachPersist(new Spreadsheet(null, "railway-local"));
      book.loadFromJSON(data);
      seedMissingTabs(book);
      console.log("[workbook] loaded", file);
      return book;
    } catch (e) {
      console.error("[workbook] could not read", file, e && e.message ? e.message : e);
    }
  }
  book = createEmptyBook();
  console.log("[workbook] new", file);
  return book;
}

function getBook() {
  return book || initWorkbook();
}

function persistWorkbook() {
  writeBookFile(getBook());
}

function ordersEmpty(target) {
  const sheet = (target || getBook()).getSheetByName("ORDERS");
  return !sheet || sheet.getLastRow() < 2;
}

function usersEmpty(target) {
  const sheet = (target || getBook()).getSheetByName("Users");
  return !sheet || sheet.getLastRow() < 2;
}

function productionLogHasRows(target) {
  const sheet = (target || getBook()).getSheetByName("Production_Log");
  return !!(sheet && sheet.getLastRow() >= 2);
}

function tabCounts(target) {
  const src = target || getBook();
  const out = {};
  Object.keys(src.sheetsByName || {}).forEach((title) => {
    const sheet = src.getSheetByName(title);
    out[title] = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
  });
  return out;
}

function hasGoogleAuth() {
  return !!(process.env.GOOGLE_SERVICE_ACCOUNT_JSON || process.env.GOOGLE_APPLICATION_CREDENTIALS);
}

function googleMigrateEnabled() {
  const flag = String(process.env.GOOGLE_MIGRATE || "").trim().toLowerCase();
  return (flag === "1" || flag === "true" || flag === "yes")
    && hasGoogleAuth()
    && !!String(process.env.SHEET_ID || "").trim();
}

async function importGoogleWorkbook() {
  const { google } = require("googleapis");
  let credentials;
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  } else if (process.env.GOOGLE_APPLICATION_CREDENTIALS) {
    credentials = require(process.env.GOOGLE_APPLICATION_CREDENTIALS);
  } else {
    throw new Error("Google credentials are not set");
  }
  if (!process.env.SHEET_ID) throw new Error("SHEET_ID is not set");
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"]
  });
  const client = await auth.getClient();
  const remote = new Spreadsheet(client, process.env.SHEET_ID);
  await remote.load();
  const local = getBook();
  local.loadFromJSON(remote.toJSON());
  attachPersist(local);
  seedMissingTabs(local);
  writeBookFile(local);
  console.log("[workbook] copied Google spreadsheet into Railway files");
  return local;
}

async function maybeImportGoogleOnce() {
  // Railway is the live database. Google Sheets is never read on boot.
  return getBook();
}

module.exports = {
  ORDER_HEADERS,
  initWorkbook,
  getBook,
  persistWorkbook,
  ordersEmpty,
  usersEmpty,
  productionLogHasRows,
  tabCounts,
  importGoogleWorkbook,
  maybeImportGoogleOnce,
  googleMigrateEnabled,
  workbookPath,
  dataDir,
  storageInfo,
  onRailway,
  hasGoogleAuth
};
