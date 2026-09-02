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
      see_debtors TEXT,
      enquiry_roles TEXT
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
    CREATE TABLE IF NOT EXISTS sheets (
      title TEXT PRIMARY KEY,
      hidden INTEGER,
      last_row INTEGER,
      last_col INTEGER
    );
    CREATE TABLE IF NOT EXISTS sheet_rows (
      title TEXT NOT NULL,
      row_idx INTEGER NOT NULL,
      json TEXT NOT NULL,
      PRIMARY KEY (title, row_idx)
    );
    CREATE TABLE IF NOT EXISTS dropdowns (
      group_name TEXT NOT NULL,
      value TEXT NOT NULL,
      PRIMARY KEY (group_name, value)
    );
    CREATE TABLE IF NOT EXISTS payments (
      order_number TEXT NOT NULL,
      seq INTEGER NOT NULL,
      at TEXT,
      amount TEXT,
      note TEXT,
      PRIMARY KEY (order_number, seq)
    );
    CREATE TABLE IF NOT EXISTS office_schedule_rows (
      id TEXT PRIMARY KEY,
      json TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS office_schedule_cells (
      row_id TEXT NOT NULL,
      day TEXT NOT NULL,
      value TEXT,
      PRIMARY KEY (row_id, day)
    );
    CREATE TABLE IF NOT EXISTS sessions (
      token TEXT PRIMARY KEY,
      json TEXT NOT NULL
    );
  `);
  opened = db;
  openedPath = file;
  try { db.exec("ALTER TABLE users ADD COLUMN enquiry_roles TEXT"); } catch (e) {}
  return db;
}

function counts() {
  const db = open();
  if (!db) {
    return {
      users: 0, orders: 0, enquiries: 0, sheets: 0, sheetRows: 0,
      dropdowns: 0, payments: 0, officeScheduleRows: 0, officeScheduleCells: 0,
      sessions: 0, hasSqlite: false, path: sqlitePath()
    };
  }
  const one = (sql) => {
    const row = db.prepare(sql).get();
    return Number(row && (row.n != null ? row.n : Object.values(row)[0])) || 0;
  };
  return {
    users: one("SELECT COUNT(*) AS n FROM users"),
    orders: one("SELECT COUNT(*) AS n FROM orders"),
    enquiries: one("SELECT COUNT(*) AS n FROM enquiries"),
    sheets: one("SELECT COUNT(*) AS n FROM sheets"),
    sheetRows: one("SELECT COUNT(*) AS n FROM sheet_rows"),
    dropdowns: one("SELECT COUNT(*) AS n FROM dropdowns"),
    payments: one("SELECT COUNT(*) AS n FROM payments"),
    officeScheduleRows: one("SELECT COUNT(*) AS n FROM office_schedule_rows"),
    officeScheduleCells: one("SELECT COUNT(*) AS n FROM office_schedule_cells"),
    sessions: one("SELECT COUNT(*) AS n FROM sessions"),
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
    "INSERT INTO users (name, role, password, tasks, access, see_debtors, enquiry_roles) VALUES (?, ?, ?, ?, ?, ?, ?)",
    rows,
    (r) => [
      String(r.name || "").trim(),
      String(r.role || ""),
      String(r.password || ""),
      String(r.tasks || ""),
      String(r.access || ""),
      String(r.see_debtors || r.see_debtor || ""),
      String(r.enquiry_roles || "")
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

function saveAllSheets(book) {
  const db = open();
  if (!db || !book || typeof book.toJSON !== "function") return 0;
  const json = book.toJSON();
  const sheets = (json && json.sheets) || {};
  db.exec("BEGIN");
  try {
    db.exec("DELETE FROM sheet_rows");
    db.exec("DELETE FROM sheets");
    const insS = db.prepare("INSERT INTO sheets (title, hidden, last_row, last_col) VALUES (?, ?, ?, ?)");
    const insR = db.prepare("INSERT INTO sheet_rows (title, row_idx, json) VALUES (?, ?, ?)");
    let rows = 0;
    Object.keys(sheets).forEach((title) => {
      const sheet = sheets[title] || {};
      insS.run(title, sheet.hidden ? 1 : 0, Number(sheet.lastRow) || 0, Number(sheet.lastCol) || 0);
      const grid = sheet.grid;
      if (Array.isArray(grid)) {
        grid.forEach((row, i) => {
          insR.run(title, i, JSON.stringify(row || []));
          rows += 1;
        });
      } else if (grid && typeof grid === "object") {
        Object.keys(grid).forEach((key) => {
          insR.run(title, Number(key), JSON.stringify(grid[key] || []));
          rows += 1;
        });
      }
    });
    db.exec("COMMIT");
    return rows;
  } catch (e) {
    try { db.exec("ROLLBACK"); } catch (err) {}
    throw e;
  }
}

function saveOfficeExtras(db, state) {
  db.exec("BEGIN");
  try {
    db.exec("DELETE FROM dropdowns");
    const insD = db.prepare("INSERT OR IGNORE INTO dropdowns (group_name, value) VALUES (?, ?)");
    const addGroups = (prefix, groups) => {
      if (!groups || typeof groups !== "object") return;
      Object.keys(groups).forEach((g) => {
        (Array.isArray(groups[g]) ? groups[g] : []).forEach((v) => {
          const val = String(v == null ? "" : v).trim();
          if (!val) return;
          insD.run(prefix + g, val);
        });
      });
    };
    addGroups("", state.dropdowns);
    addGroups("enquiry:", state.enquiry_dropdowns);

    db.exec("DELETE FROM payments");
    const insP = db.prepare("INSERT INTO payments (order_number, seq, at, amount, note) VALUES (?, ?, ?, ?, ?)");
    const pay = state.paymentsByOrder || {};
    Object.keys(pay).forEach((order) => {
      (Array.isArray(pay[order]) ? pay[order] : []).forEach((p, i) => {
        const item = p && typeof p === "object" ? p : { amount: p };
        insP.run(
          String(order),
          i,
          item.at || item.date || "",
          item.amount == null ? "" : String(item.amount),
          item.note || ""
        );
      });
    });

    db.exec("DELETE FROM office_schedule_rows");
    db.exec("DELETE FROM office_schedule_cells");
    const insR = db.prepare("INSERT INTO office_schedule_rows (id, json) VALUES (?, ?)");
    (state.schedule_rows || []).forEach((row) => {
      if (!row || row.id == null) return;
      insR.run(String(row.id), JSON.stringify(row));
    });
    const insC = db.prepare("INSERT INTO office_schedule_cells (row_id, day, value) VALUES (?, ?, ?)");
    (state.schedule_cells || []).forEach((c) => {
      if (!c || c.row_id == null || c.day == null) return;
      insC.run(String(c.row_id), String(c.day), c.value == null ? "" : String(c.value));
    });

    db.prepare("INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)").run("next_order_id", String(state.nextOrderId || 1));
    db.prepare("INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)").run("next_schedule_id", String(state.nextScheduleId || 1));
    db.exec("COMMIT");
  } catch (e) {
    try { db.exec("ROLLBACK"); } catch (err) {}
    throw e;
  }
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
      nextScheduleId: state.nextScheduleId || 1,
      orders: state.orders || []
    })
  );
  saveOfficeExtras(db, state);
  db.prepare("INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)").run("office_saved_at", new Date().toISOString());
  return enquiries.length;
}

function saveWorkbook(book) {
  const db = open();
  if (!db || !book) return { users: 0, orders: 0, sheetRows: 0 };
  const users = saveUsersFromBook(book);
  const orders = saveOrdersFromBook(book);
  const sheetRows = saveAllSheets(book);
  if (typeof book.toJSON === "function") {
    db.prepare("INSERT OR REPLACE INTO blobs (kind, json) VALUES (?, ?)").run(
      "workbook",
      JSON.stringify(book.toJSON())
    );
  }
  db.prepare("INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)").run("workbook_saved_at", new Date().toISOString());
  db.prepare("INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)").run("engine", "sqlite");
  return { users, orders, sheetRows };
}

function loadDropdowns(db) {
  const dropdowns = {};
  const enquiry_dropdowns = {};
  db.prepare("SELECT group_name, value FROM dropdowns").all().forEach((r) => {
    const name = r.group_name || "";
    if (name.indexOf("enquiry:") === 0) {
      const g = name.slice("enquiry:".length);
      enquiry_dropdowns[g] = enquiry_dropdowns[g] || [];
      enquiry_dropdowns[g].push(r.value);
    } else {
      dropdowns[name] = dropdowns[name] || [];
      dropdowns[name].push(r.value);
    }
  });
  return { dropdowns, enquiry_dropdowns };
}

function loadPayments(db) {
  const paymentsByOrder = {};
  db.prepare("SELECT order_number, seq, at, amount, note FROM payments ORDER BY order_number, seq").all().forEach((r) => {
    paymentsByOrder[r.order_number] = paymentsByOrder[r.order_number] || [];
    paymentsByOrder[r.order_number].push({
      at: r.at || "",
      amount: r.amount || "",
      note: r.note || ""
    });
  });
  return paymentsByOrder;
}

function loadOfficeSchedule(db) {
  const schedule_rows = db.prepare("SELECT json FROM office_schedule_rows").all().map((r) => {
    try { return JSON.parse(r.json); } catch (e) { return null; }
  }).filter(Boolean);
  const schedule_cells = db.prepare("SELECT row_id, day, value FROM office_schedule_cells").all().map((r) => {
    const n = Number(r.row_id);
    return {
      row_id: Number.isFinite(n) ? n : r.row_id,
      day: r.day,
      value: r.value || ""
    };
  });
  return { schedule_rows, schedule_cells };
}

function metaNumber(db, key, fallback) {
  const row = db.prepare("SELECT value FROM meta WHERE key = ?").get(key);
  const v = row && row.value != null ? Number(row.value) : NaN;
  return Number.isFinite(v) && v > 0 ? v : fallback;
}

function loadOffice() {
  const db = open();
  if (!db) return null;
  const n = counts();
  const officeBlob = db.prepare("SELECT json FROM blobs WHERE kind = ?").get("office");
  let extras = {};
  try { extras = officeBlob && officeBlob.json ? JSON.parse(officeBlob.json) : {}; } catch (e) { extras = {}; }
  if (!n.enquiries && !n.dropdowns && !n.payments && !n.officeScheduleRows && !officeBlob) return null;

  const enquiryRows = db.prepare("SELECT * FROM enquiries").all();
  const enquiries = enquiryRows.map((row) => {
    let extra = {};
    try { extra = row.extra_json ? JSON.parse(row.extra_json) : {}; } catch (e) { extra = {}; }
    const out = { ...extra };
    ENQUIRY_FIELDS.forEach((k) => { out[k] = row[k] == null ? "" : row[k]; });
    return out;
  });
  const fromDrop = n.dropdowns ? loadDropdowns(db) : {
    dropdowns: extras.dropdowns || {},
    enquiry_dropdowns: extras.enquiry_dropdowns || {}
  };
  const paymentsByOrder = n.payments ? loadPayments(db) : (extras.paymentsByOrder || {});
  const sched = (n.officeScheduleRows || n.officeScheduleCells)
    ? loadOfficeSchedule(db)
    : {
      schedule_rows: extras.schedule_rows || [],
      schedule_cells: extras.schedule_cells || []
    };
  return {
    orders: extras.orders || [],
    schedule_rows: sched.schedule_rows,
    schedule_cells: sched.schedule_cells,
    nextOrderId: metaNumber(db, "next_order_id", extras.nextOrderId || 1),
    nextScheduleId: metaNumber(db, "next_schedule_id", extras.nextScheduleId || 1),
    dropdowns: Object.keys(fromDrop.dropdowns).length ? fromDrop.dropdowns : (extras.dropdowns || {}),
    paymentsByOrder,
    enquiries,
    enquiry_dropdowns: Object.keys(fromDrop.enquiry_dropdowns).length
      ? fromDrop.enquiry_dropdowns
      : (extras.enquiry_dropdowns || {})
  };
}

function loadWorkbookFromSheets() {
  const db = open();
  if (!db) return null;
  const sheets = db.prepare("SELECT title, hidden, last_row, last_col FROM sheets").all();
  const rows = db.prepare("SELECT title, row_idx, json FROM sheet_rows ORDER BY title, row_idx").all();
  if (!sheets.length && !rows.length) return null;
  const out = { version: 1, sheets: {} };
  sheets.forEach((s) => {
    out.sheets[s.title] = {
      title: s.title,
      hidden: !!s.hidden,
      lastRow: s.last_row || 0,
      lastCol: s.last_col || 0,
      grid: []
    };
  });
  rows.forEach((r) => {
    if (!out.sheets[r.title]) {
      out.sheets[r.title] = { title: r.title, hidden: false, lastRow: 0, lastCol: 0, grid: [] };
    }
    let cells = [];
    try { cells = JSON.parse(r.json); } catch (e) { cells = []; }
    out.sheets[r.title].grid[Number(r.row_idx)] = cells;
  });
  Object.keys(out.sheets).forEach((title) => {
    const sheet = out.sheets[title];
    for (let i = 0; i < sheet.grid.length; i++) {
      if (!sheet.grid[i]) sheet.grid[i] = [];
    }
    if (!sheet.lastRow) sheet.lastRow = sheet.grid.length;
    if (!sheet.lastCol) {
      sheet.lastCol = sheet.grid.reduce((m, row) => Math.max(m, (row || []).length), 0);
    }
  });
  return out;
}

function loadWorkbookJson() {
  const fromSheets = loadWorkbookFromSheets();
  if (fromSheets) return fromSheets;
  const db = open();
  if (!db) return null;
  const row = db.prepare("SELECT json FROM blobs WHERE kind = ?").get("workbook");
  if (!row || !row.json) return null;
  try { return JSON.parse(row.json); } catch (e) { return null; }
}

function saveSessions(mapOrObj) {
  const db = open();
  if (!db) return 0;
  const entries = [];
  if (mapOrObj && typeof mapOrObj.forEach === "function") {
    mapOrObj.forEach((val, token) => entries.push([token, val]));
  } else if (mapOrObj && typeof mapOrObj === "object") {
    Object.keys(mapOrObj).forEach((token) => entries.push([token, mapOrObj[token]]));
  }
  db.exec("BEGIN");
  try {
    db.exec("DELETE FROM sessions");
    const ins = db.prepare("INSERT INTO sessions (token, json) VALUES (?, ?)");
    entries.forEach(([token, val]) => {
      if (!token) return;
      ins.run(String(token), JSON.stringify(val || {}));
    });
    db.exec("COMMIT");
    return entries.length;
  } catch (e) {
    try { db.exec("ROLLBACK"); } catch (err) {}
    throw e;
  }
}

function loadSessions() {
  const db = open();
  if (!db) return null;
  const rows = db.prepare("SELECT token, json FROM sessions").all();
  if (!rows.length) return null;
  const out = {};
  rows.forEach((r) => {
    try { out[r.token] = JSON.parse(r.json); } catch (e) {}
  });
  return out;
}

function importIfEmpty(book, officeState) {
  const n = counts();
  const out = { users: n.users, orders: n.orders, enquiries: n.enquiries, imported: false };
  if (!n.hasSqlite) return out;
  if (book && (!n.users && !n.orders || !n.sheetRows)) {
    const saved = saveWorkbook(book);
    out.users = saved.users;
    out.orders = saved.orders;
    out.imported = true;
  }
  if (officeState && (!n.enquiries || !n.dropdowns)) {
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
    sqliteAvailable: n.hasSqlite,
    sqliteUsers: n.users,
    sqliteOrders: n.orders,
    sqliteEnquiries: n.enquiries,
    sqliteSheets: n.sheets,
    sqliteSheetRows: n.sheetRows,
    sqliteDropdowns: n.dropdowns,
    sqlitePayments: n.payments,
    sqliteOfficeScheduleRows: n.officeScheduleRows,
    sqliteSessions: n.sessions
  };
}

module.exports = {
  sqlitePath,
  sqliteAvailable,
  open,
  counts,
  saveOffice,
  saveWorkbook,
  saveAllSheets,
  saveSessions,
  loadOffice,
  loadWorkbookJson,
  loadWorkbookFromSheets,
  loadSessions,
  importIfEmpty,
  checkpoint,
  info
};
