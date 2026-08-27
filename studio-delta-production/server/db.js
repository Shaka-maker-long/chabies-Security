const fs = require("fs");
const path = require("path");
const Database = require("better-sqlite3");

const dataDir = process.env.DATA_DIR || path.join(__dirname, "..", "data");
fs.mkdirSync(dataDir, { recursive: true });
const dbPath = process.env.SQLITE_PATH || path.join(dataDir, "studio-delta.sqlite");
const db = new Database(dbPath);
db.pragma("journal_mode = WAL");

db.exec(`
CREATE TABLE IF NOT EXISTS users (
  id INTEGER PRIMARY KEY,
  name TEXT NOT NULL UNIQUE,
  role TEXT,
  password TEXT NOT NULL,
  tasks TEXT
);

CREATE TABLE IF NOT EXISTS orders (
  id INTEGER PRIMARY KEY,
  quote_number TEXT,
  order_number TEXT NOT NULL UNIQUE,
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
  price_incl_vat TEXT,
  price_excl_vat TEXT,
  month_of_sale TEXT,
  source TEXT,
  city TEXT,
  updated_at TEXT
);

CREATE TABLE IF NOT EXISTS schedule_rows (
  id INTEGER PRIMARY KEY,
  order_number TEXT NOT NULL,
  item_type TEXT,
  category TEXT,
  product TEXT,
  province TEXT,
  order_date TEXT,
  courier TEXT,
  waybill TEXT,
  status TEXT,
  sort_order INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS schedule_cells (
  row_id INTEGER NOT NULL,
  day TEXT NOT NULL,
  value TEXT,
  PRIMARY KEY (row_id, day),
  FOREIGN KEY (row_id) REFERENCES schedule_rows(id) ON DELETE CASCADE
);
`);

const ORDER_FIELDS = [
  "quote_number", "order_number", "status", "assigned_operator", "type", "category",
  "product", "variation", "doors", "detailed_description", "dimensions", "powder_coating",
  "client_name", "client_number", "email", "payment_date", "address", "province",
  "price_incl_vat", "price_excl_vat", "month_of_sale", "source", "city"
];

function nowIso() {
  return new Date().toISOString();
}

function listOrders() {
  return db.prepare("SELECT * FROM orders ORDER BY id DESC").all();
}

function upsertOrder(row) {
  const orderNumber = String(row.order_number || "").trim();
  if (!orderNumber) throw new Error("Order number is required");
  const existing = db.prepare("SELECT id FROM orders WHERE order_number = ?").get(orderNumber);
  const payload = {};
  for (const f of ORDER_FIELDS) payload[f] = row[f] == null ? "" : String(row[f]);
  payload.order_number = orderNumber;
  payload.updated_at = nowIso();
  if (existing) {
    const sets = ORDER_FIELDS.filter((f) => f !== "order_number").map((f) => f + " = @" + f);
    db.prepare("UPDATE orders SET " + sets.join(", ") + ", updated_at = @updated_at WHERE id = @id").run({
      ...payload,
      id: existing.id
    });
    return db.prepare("SELECT * FROM orders WHERE id = ?").get(existing.id);
  }
  const cols = ORDER_FIELDS.concat("updated_at");
  db.prepare(
    "INSERT INTO orders (" + cols.join(",") + ") VALUES (" + cols.map((c) => "@" + c).join(",") + ")"
  ).run(payload);
  return db.prepare("SELECT * FROM orders WHERE order_number = ?").get(orderNumber);
}

function deleteOrder(orderNumber) {
  db.prepare("DELETE FROM orders WHERE order_number = ?").run(orderNumber);
}

function listSchedule(fromDay, toDay) {
  const rows = db.prepare("SELECT * FROM schedule_rows ORDER BY sort_order, id").all();
  const cells = db.prepare(
    "SELECT row_id, day, value FROM schedule_cells WHERE day >= ? AND day <= ?"
  ).all(fromDay, toDay);
  const byRow = {};
  for (const c of cells) {
    if (!byRow[c.row_id]) byRow[c.row_id] = {};
    byRow[c.row_id][c.day] = c.value;
  }
  return rows.map((r) => ({ ...r, cells: byRow[r.id] || {} }));
}

function upsertScheduleRow(row) {
  const orderNumber = String(row.order_number || "").trim();
  if (!orderNumber) throw new Error("Order number is required");
  let id = row.id;
  if (id) {
    db.prepare(`UPDATE schedule_rows SET
      order_number=@order_number, item_type=@item_type, category=@category, product=@product,
      province=@province, order_date=@order_date, courier=@courier, waybill=@waybill,
      status=@status, sort_order=@sort_order WHERE id=@id`).run({
      id,
      order_number: orderNumber,
      item_type: row.item_type || "",
      category: row.category || "",
      product: row.product || "",
      province: row.province || "",
      order_date: row.order_date || "",
      courier: row.courier || "",
      waybill: row.waybill || "",
      status: row.status || "",
      sort_order: Number(row.sort_order) || 0
    });
  } else {
    const info = db.prepare(`INSERT INTO schedule_rows
      (order_number, item_type, category, product, province, order_date, courier, waybill, status, sort_order)
      VALUES (@order_number, @item_type, @category, @product, @province, @order_date, @courier, @waybill, @status, @sort_order)`).run({
      order_number: orderNumber,
      item_type: row.item_type || "",
      category: row.category || "",
      product: row.product || "",
      province: row.province || "",
      order_date: row.order_date || "",
      courier: row.courier || "",
      waybill: row.waybill || "",
      status: row.status || "",
      sort_order: Number(row.sort_order) || 0
    });
    id = info.lastInsertRowid;
  }
  if (row.cells && typeof row.cells === "object") {
    const ins = db.prepare("INSERT INTO schedule_cells (row_id, day, value) VALUES (?, ?, ?) ON CONFLICT(row_id, day) DO UPDATE SET value = excluded.value");
    const del = db.prepare("DELETE FROM schedule_cells WHERE row_id = ? AND day = ?");
    for (const day of Object.keys(row.cells)) {
      const value = row.cells[day];
      if (value == null || String(value).trim() === "") del.run(id, day);
      else ins.run(id, day, String(value));
    }
  }
  return db.prepare("SELECT * FROM schedule_rows WHERE id = ?").get(id);
}

function setScheduleCell(rowId, day, value) {
  if (value == null || String(value).trim() === "") {
    db.prepare("DELETE FROM schedule_cells WHERE row_id = ? AND day = ?").run(rowId, day);
  } else {
    db.prepare(
      "INSERT INTO schedule_cells (row_id, day, value) VALUES (?, ?, ?) ON CONFLICT(row_id, day) DO UPDATE SET value = excluded.value"
    ).run(rowId, day, String(value));
  }
}

function countOrders() {
  return db.prepare("SELECT COUNT(*) AS n FROM orders").get().n;
}

module.exports = {
  db,
  dbPath,
  ORDER_FIELDS,
  listOrders,
  upsertOrder,
  deleteOrder,
  listSchedule,
  upsertScheduleRow,
  setScheduleCell,
  countOrders
};
