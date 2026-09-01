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

function dataDir() {
  const dir = process.env.DATA_DIR || path.join(__dirname, "..", "data");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function workbookPath() {
  return path.join(dataDir(), "floor-workbook.json");
}

let book = null;
let googleImportTried = false;

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
  googleImportTried = true;
  console.log("[workbook] imported from Google Sheets");
  return local;
}

async function maybeImportGoogleOnce() {
  const local = getBook();
  if (googleImportTried) return local;
  googleImportTried = true;
  if (!hasGoogleAuth() || !process.env.SHEET_ID) return local;
  if (!usersEmpty(local) || !ordersEmpty(local)) return local;
  if (productionLogHasRows(local)) return local;
  try {
    await importGoogleWorkbook();
  } catch (e) {
    console.error("[workbook] google import failed", e && e.message ? e.message : e);
  }
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
  workbookPath,
  hasGoogleAuth
};
