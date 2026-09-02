const fs = require("fs");
const path = require("path");
const { dataDir } = require("./workbook-store");
const { ENQUIRY_FIELDS } = require("./enquiries-default");

const ORDER_COLUMNS = [
  "quote_number", "order_number", "status", "assigned_operator", "type", "category",
  "product", "variation", "doors", "detailed_description", "dimensions", "powder_coating",
  "client_name", "client_number", "email", "payment_date", "address", "province",
  "price_excl_vat", "price_incl_vat", "amount_paid", "month_of_sale", "source", "city"
];

let opened = null;
let openedPath = "";

function sqlitePath() {
  return path.join(dataDir(), "studio-delta.db");
}

function sqliteAvailable() {
  try {
    require("node:sqlite");
    return true;
  } catch (e) {
    return false;
  }
}

function open() {
  const file = sqlitePath();
  if (opened && openedPath === file) return opened;
  if (opened) {
    try { opened.close(); } catch (e) {}
    opened = null;
  }
  if (!sqliteAvailable()) return null;
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const { DatabaseSync } = require("node:sqlite");
  const db = new DatabaseSync(file);
  db.exec("PRAGMA journal_mode = WAL;");
  db.exec(`
    CREATE TABLE IF NOT EXISTS meta (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS users (
      name TEXT PRIMARY KEY,
      role TEXT,
      password TEXT,
      tasks TEXT,
      access TEXT,
      see_debtors TEXT
    );
    CREATE TABLE IF NOT EXISTS orders (
      order_number TEXT PRIMARY KEY,
      quote_number TEXT,
      status TEXT,
      assigned_operator TEXT,
      type TEXT,
      category TEXT,
      product TEXT,
      variation TEXT,
      doors TEXT,
      detailed_description TEXT,
      dimensions TEXT,
      powder_coating TEXT,
      client_name TEXT,
      client_number TEXT,
      email TEXT,
      payment_date TEXT,
      address TEXT,
      province TEXT,
      price_excl_vat TEXT,
      price_incl_vat TEXT,
      amount_paid TEXT,
      month_of_sale TEXT,
      source TEXT,
      city TEXT,
      extra_json TEXT
    );
    CREATE TABLE IF NOT EXISTS enquiries (
      enquiry_no TEXT PRIMARY KEY,
      date_enquired TEXT,
      month_enquired TEXT,
      enquiry_source TEXT,
      enquiry_type TEXT,
      client_name TEXT,
      source TEXT,
      client_email TEXT,
      client_number TEXT,
      province TEXT,
      category TEXT,
      product TEXT,
      request TEXT,
      status TEXT,
      date_quoted TEXT,
      quote_no TEXT,
      comment TEXT,
      extra_json TEXT
    );
    CREATE TABLE IF NOT EXISTS blobs (
      kind TEXT PRIMARY KEY,
      json TEXT NOT NULL
    );
  `);
  opened = db;
  openedPath = file;
  return db;
}

function counts() {
  const db = open();
  if (!db) return { users: 0, orders: 0, enquiries: 0, hasSqlite: false, path: sqlitePath() };
  const one = (sql) => {
    const row = db.prepare(sql).get();
    return Number(row && (row.n != null ? row.n : Object.values(row)[0])) || 0;
  };
  return {
    users: one("SELECT COUNT(*) AS n FROM users"),
    orders: one("SELECT COUNT(*) AS n FROM orders"),
    enquiries: one("SELECT COUNT(*) AS n FROM enquiries"),
    hasSqlite: true,
    path: sqlitePath(),
    exists: fs.existsSync(sqlitePath())
  };
}

function headerKey(h) {
  return String(h || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");
}

function sheetObjects(book, title) {
  if (!book || typeof book.getSheetByName !== "function") return [];
  const sheet = book.getSheetByName(title);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = (sheet.getRange(1, 1, 1, lastCol).getValues()[0] || []).map(headerKey);
  const grid = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  return grid.map((row) => {
    const obj = {};
    headers.forEach((key, i) => {
      if (!key) return;
      obj[key] = row[i] == null ? "" : row[i];
    });
    return obj;
  });
}

function runReplace(db, deleteSql, insertSql, rows, valuesFn) {
  db.exec("BEGIN");
  try {
    db.exec(deleteSql);
    const stmt = db.prepare(insertSql);
    rows.forEach((row) => stmt.run(...valuesFn(row)));
    db.exec("COMMIT");
  } catch (e) {
    try { db.exec("ROLLBACK"); } catch (err) {}
    throw e;
  }
}

function saveUsersFromBook(book) {
  const db = open();
  if (!db) return 0;
  const rows = sheetObjects(book, "Users").filter((r) => String(r.name || "").trim());
  runReplace(
    db,
    "DELETE FROM users",
    "INSERT INTO users (name, role, password, tasks, access, see_debtors) VALUES (?, ?, ?, ?, ?, ?)",
    rows,
    (r) => [
      String(r.name || "").trim(),
      String(r.role || ""),
      String(r.password || ""),
      String(r.tasks || ""),
      String(r.access || ""),
      String(r.see_debtors || r.see_debtor || "")
    ]
  );
  return rows.length;
}

function saveOrdersFromBook(book) {
  const db = open();
  if (!db) return 0;
  const rows = sheetObjects(book, "ORDERS").filter((r) => {
    return String(r.order_number || r.order || "").trim();
  }).map((r) => {
    if (!r.order_number && r.order) r.order_number = r.order;
    if (!r.category && r.catergory) r.category = r.catergory;
    if (!r.email && r.email_address) r.email = r.email_address;
    if (!r.price_excl_vat && r.price_excl_vat_ !== undefined) r.price_excl_vat = r.price_excl_vat_;
    return r;
  });
  const extraKeys = (r) => {
    const extra = {};
    Object.keys(r).forEach((k) => {
      if (ORDER_COLUMNS.indexOf(k) === -1 && k !== "catergory" && k !== "email_address") extra[k] = r[k];
    });
    return Object.keys(extra).length ? JSON.stringify(extra) : "";
  };
  runReplace(
    db,
    "DELETE FROM orders",
    "INSERT INTO orders (" + ORDER_COLUMNS.join(", ") + ", extra_json) VALUES (" + ORDER_COLUMNS.map(() => "?").join(", ") + ", ?)",
    rows,
    (r) => ORDER_COLUMNS.map((k) => (r[k] == null ? "" : String(r[k]))).concat([extraKeys(r)])
  );
  return rows.length;
}

function saveOffice(state) {
  const db = open();
  if (!db || !state) return 0;
  const enquiries = Array.isArray(state.enquiries) ? state.enquiries : [];
  runReplace(
    db,
    "DELETE FROM enquiries",
    "INSERT INTO enquiries (" + ENQUIRY_FIELDS.join(", ") + ", extra_json) VALUES (" + ENQUIRY_FIELDS.map(() => "?").join(", ") + ", ?)",
    enquiries.filter((r) => r && r.enquiry_no),
    (row) => {
      const extra = {};
      Object.keys(row).forEach((k) => {
        if (ENQUIRY_FIELDS.indexOf(k) === -1) extra[k] = row[k];
      });
      return ENQUIRY_FIELDS.map((k) => (row[k] == null ? "" : String(row[k]))).concat([
        JSON.stringify(extra)
      ]);
    }
  );
  db.prepare("INSERT OR REPLACE INTO blobs (kind, json) VALUES (?, ?)").run(
    "office",
    JSON.stringify({
      dropdowns: state.dropdowns || {},
      enquiry_dropdowns: state.enquiry_dropdowns || {},
      paymentsByOrder: state.paymentsByOrder || {},
      schedule_rows: state.schedule_rows || [],
      schedule_cells: state.schedule_cells || [],
      nextOrderId: state.nextOrderId || 1,
      nextScheduleId: state.nextScheduleId || 1
    })
  );
  db.prepare("INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)").run("office_saved_at", new Date().toISOString());
  return enquiries.length;
}

function saveWorkbook(book) {
  const db = open();
  if (!db || !book) return { users: 0, orders: 0 };
  const users = saveUsersFromBook(book);
  const orders = saveOrdersFromBook(book);
  if (typeof book.toJSON === "function") {
    db.prepare("INSERT OR REPLACE INTO blobs (kind, json) VALUES (?, ?)").run(
      "workbook",
      JSON.stringify(book.toJSON())
    );
  }
  db.prepare("INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)").run("workbook_saved_at", new Date().toISOString());
  db.prepare("INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)").run("engine", "sqlite");
  return { users, orders };
}

function loadOffice() {
  const db = open();
  if (!db) return null;
  const n = counts();
  if (!n.enquiries && !n.orders && !n.users) {
    const blob = db.prepare("SELECT json FROM blobs WHERE kind = ?").get("office");
    if (!blob || !blob.json) return null;
  }
  const enquiryRows = db.prepare("SELECT * FROM enquiries").all();
  const officeBlob = db.prepare("SELECT json FROM blobs WHERE kind = ?").get("office");
  let extras = {};
  try { extras = officeBlob && officeBlob.json ? JSON.parse(officeBlob.json) : {}; } catch (e) { extras = {}; }
  const enquiries = enquiryRows.map((row) => {
    let extra = {};
    try { extra = row.extra_json ? JSON.parse(row.extra_json) : {}; } catch (e) { extra = {}; }
    const out = { ...extra };
    ENQUIRY_FIELDS.forEach((k) => { out[k] = row[k] == null ? "" : row[k]; });
    return out;
  });
  return {
    orders: extras.orders || [],
    schedule_rows: extras.schedule_rows || [],
    schedule_cells: extras.schedule_cells || [],
    nextOrderId: extras.nextOrderId || 1,
    nextScheduleId: extras.nextScheduleId || 1,
    dropdowns: extras.dropdowns || {},
    paymentsByOrder: extras.paymentsByOrder || {},
    enquiries,
    enquiry_dropdowns: extras.enquiry_dropdowns || {}
  };
}

function loadWorkbookJson() {
  const db = open();
  if (!db) return null;
  const row = db.prepare("SELECT json FROM blobs WHERE kind = ?").get("workbook");
  if (!row || !row.json) return null;
  try { return JSON.parse(row.json); } catch (e) { return null; }
}

function importIfEmpty(book, officeState) {
  const n = counts();
  const out = { users: n.users, orders: n.orders, enquiries: n.enquiries, imported: false };
  if (!n.hasSqlite) return out;
  if (!n.users && !n.orders && book) {
    const saved = saveWorkbook(book);
    out.users = saved.users;
    out.orders = saved.orders;
    out.imported = true;
  }
  if (!n.enquiries && officeState) {
    out.enquiries = saveOffice(officeState);
    out.imported = true;
  } else if (officeState && n.enquiries === 0) {
    out.enquiries = saveOffice(officeState);
    out.imported = true;
  }
  return out;
}

function checkpoint() {
  const db = open();
  if (!db) return false;
  try { db.exec("PRAGMA wal_checkpoint(TRUNCATE);"); } catch (e) {}
  return fs.existsSync(sqlitePath());
}

function info() {
  const n = counts();
  return {
    database: n.hasSqlite
      ? "SQLite on the Railway volume (not Google Sheets, not Postgres)"
      : "JSON files on disk (SQLite not available on this Node)",
    sqlitePath: n.path,
    sqliteExists: fs.existsSync(n.path),
    sqliteUsers: n.users,
    sqliteOrders: n.orders,
    sqliteEnquiries: n.enquiries,
    sqliteAvailable: n.hasSqlite
  };
}

module.exports = {
  sqlitePath,
  sqliteAvailable,
  open,
  counts,
  saveOffice,
  saveWorkbook,
  loadOffice,
  loadWorkbookJson,
  importIfEmpty,
  checkpoint,
  info
};
