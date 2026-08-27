const fs = require("fs");
const path = require("path");

const ORDER_FIELDS = [
  "quote_number", "order_number", "status", "assigned_operator", "type", "category",
  "product", "variation", "doors", "detailed_description", "dimensions", "powder_coating",
  "client_name", "client_number", "email", "payment_date", "address", "province",
  "price_incl_vat", "price_excl_vat", "month_of_sale", "source", "city"
];

function emptyState() {
  return { orders: [], schedule_rows: [], schedule_cells: [], nextOrderId: 1, nextScheduleId: 1 };
}

function pickDataFile() {
  const dataDir = process.env.DATA_DIR || path.join(__dirname, "..", "data");
  const preferred = process.env.OFFICE_DB_PATH || path.join(dataDir, "studio-delta.json");
  const fallback = path.join("/tmp", "studio-delta.json");
  for (const candidate of [preferred, fallback]) {
    try {
      fs.mkdirSync(path.dirname(candidate), { recursive: true });
      fs.accessSync(path.dirname(candidate), fs.constants.W_OK);
      return candidate;
    } catch (e) {
      console.error("[db] not writable", candidate, e && e.message ? e.message : e);
    }
  }
  return fallback;
}

const dbPath = pickDataFile();
let state = emptyState();
try {
  const raw = fs.readFileSync(dbPath, "utf8");
  const parsed = JSON.parse(raw);
  state = {
    ...emptyState(),
    ...parsed,
    orders: Array.isArray(parsed.orders) ? parsed.orders : [],
    schedule_rows: Array.isArray(parsed.schedule_rows) ? parsed.schedule_rows : [],
    schedule_cells: Array.isArray(parsed.schedule_cells) ? parsed.schedule_cells : []
  };
  console.log("[db] opened", dbPath, "orders", state.orders.length);
} catch (e) {
  if (e && e.code !== "ENOENT") {
    console.error("[db] could not read", dbPath, e && e.message ? e.message : e);
  } else {
    console.log("[db] new file", dbPath);
  }
}

function save() {
  const tmp = dbPath + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(state));
  fs.renameSync(tmp, dbPath);
}

function nowIso() {
  return new Date().toISOString();
}

function listOrders() {
  return state.orders.slice().sort((a, b) => Number(b.id) - Number(a.id));
}

function upsertOrder(row) {
  const orderNumber = String(row.order_number || "").trim();
  if (!orderNumber) throw new Error("Order number is required");
  const payload = {};
  for (const f of ORDER_FIELDS) payload[f] = row[f] == null ? "" : String(row[f]);
  payload.order_number = orderNumber;
  payload.updated_at = nowIso();
  const existing = state.orders.find((o) => o.order_number === orderNumber);
  if (existing) {
    Object.assign(existing, payload);
    save();
    return existing;
  }
  payload.id = state.nextOrderId++;
  state.orders.push(payload);
  save();
  return payload;
}

function deleteOrder(orderNumber) {
  state.orders = state.orders.filter((o) => o.order_number !== orderNumber);
  save();
}

function listSchedule(fromDay, toDay) {
  const rows = state.schedule_rows.slice().sort((a, b) => (a.sort_order - b.sort_order) || (a.id - b.id));
  return rows.map((r) => {
    const cells = {};
    for (const c of state.schedule_cells) {
      if (c.row_id === r.id && c.day >= fromDay && c.day <= toDay) cells[c.day] = c.value;
    }
    return { ...r, cells };
  });
}

function upsertScheduleRow(row) {
  const orderNumber = String(row.order_number || "").trim();
  if (!orderNumber) throw new Error("Order number is required");
  let id = row.id ? Number(row.id) : 0;
  let existing = id ? state.schedule_rows.find((r) => r.id === id) : null;
  const payload = {
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
  };
  if (existing) {
    Object.assign(existing, payload);
  } else {
    id = state.nextScheduleId++;
    existing = { id, ...payload };
    state.schedule_rows.push(existing);
  }
  if (row.cells && typeof row.cells === "object") {
    for (const day of Object.keys(row.cells)) {
      setScheduleCell(id, day, row.cells[day], false);
    }
  }
  save();
  return existing;
}

function setScheduleCell(rowId, day, value, persist = true) {
  const id = Number(rowId);
  state.schedule_cells = state.schedule_cells.filter((c) => !(c.row_id === id && c.day === day));
  if (value != null && String(value).trim() !== "") {
    state.schedule_cells.push({ row_id: id, day, value: String(value) });
  }
  if (persist) save();
}

function countOrders() {
  return state.orders.length;
}

module.exports = {
  db: null,
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
