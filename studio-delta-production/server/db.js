const fs = require("fs");
const path = require("path");
const { DROPDOWN_KEYS, DEFAULT_DROPDOWNS } = require("./dropdowns-default");

const ORDER_FIELDS = [
  "quote_number", "order_number", "status", "assigned_operator", "type", "category",
  "product", "variation", "doors", "detailed_description", "dimensions", "powder_coating",
  "client_name", "client_number", "email", "payment_date", "address", "province",
  "price_excl_vat", "price_incl_vat", "amount_paid", "month_of_sale", "source", "city"
];

const VAT_RATE = 0.15;

function parseMoney(s) {
  const n = Number(String(s || "").replace(/,/g, "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0;
}

function money(n) {
  return (Math.round(Number(n) * 100) / 100).toFixed(2);
}

function inclFromExcl(excl) {
  const n = parseMoney(excl);
  if (!n) return "";
  return money(n * (1 + VAT_RATE));
}

function orderTotal(order) {
  const incl = parseMoney(order && order.price_incl_vat);
  if (incl) return incl;
  return parseMoney(inclFromExcl(order && order.price_excl_vat));
}

function orderPaid(order) {
  return parseMoney(order && order.amount_paid);
}

function orderOwing(order) {
  return Math.max(0, Math.round((orderTotal(order) - orderPaid(order)) * 100) / 100);
}

function applyPriceAndPayments(payload, row, existing) {
  if (payload.price_excl_vat) payload.price_incl_vat = inclFromExcl(payload.price_excl_vat);
  payload.amount_paid = payload.amount_paid === "" || payload.amount_paid == null
    ? (existing && existing.amount_paid) || "0.00"
    : money(parseMoney(payload.amount_paid));
  payload.payments = Array.isArray(row.payments)
    ? row.payments
    : (existing && Array.isArray(existing.payments) ? existing.payments : []);
  return payload;
}

function emptyState() {
  return {
    orders: [],
    schedule_rows: [],
    schedule_cells: [],
    nextOrderId: 1,
    nextScheduleId: 1,
    dropdowns: JSON.parse(JSON.stringify(DEFAULT_DROPDOWNS))
  };
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
  const dropdowns = JSON.parse(JSON.stringify(DEFAULT_DROPDOWNS));
  if (parsed.dropdowns && typeof parsed.dropdowns === "object") {
    for (const key of DROPDOWN_KEYS) {
      if (Array.isArray(parsed.dropdowns[key])) dropdowns[key] = parsed.dropdowns[key];
    }
  }
  state = {
    ...emptyState(),
    ...parsed,
    orders: Array.isArray(parsed.orders) ? parsed.orders : [],
    schedule_rows: Array.isArray(parsed.schedule_rows) ? parsed.schedule_rows : [],
    schedule_cells: Array.isArray(parsed.schedule_cells) ? parsed.schedule_cells : [],
    dropdowns
  };
  console.log("[db] opened", dbPath, "orders", state.orders.length);
  if (!parsed.dropdowns) save();
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
  applyPriceAndPayments(payload, row, existing);
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

function listDropdowns() {
  const out = {};
  for (const key of DROPDOWN_KEYS) {
    out[key] = Array.isArray(state.dropdowns[key]) ? state.dropdowns[key].slice() : DEFAULT_DROPDOWNS[key].slice();
  }
  return out;
}

function addDropdownItem(field, value) {
  if (!DROPDOWN_KEYS.includes(field)) throw new Error("Unknown dropdown");
  const item = String(value || "").trim();
  if (!item) throw new Error("Value is required");
  if (!Array.isArray(state.dropdowns[field])) state.dropdowns[field] = [];
  const exists = state.dropdowns[field].some((v) => String(v).toLowerCase() === item.toLowerCase());
  if (!exists) state.dropdowns[field].push(item);
  save();
  return listDropdowns();
}

function removeDropdownItem(field, value) {
  if (!DROPDOWN_KEYS.includes(field)) throw new Error("Unknown dropdown");
  const item = String(value || "").trim();
  state.dropdowns[field] = (state.dropdowns[field] || []).filter((v) => v !== item);
  save();
  return listDropdowns();
}

function decorateMoney(order) {
  const total = orderTotal(order);
  const paid = orderPaid(order);
  const owing = orderOwing(order);
  return {
    ...order,
    total: money(total),
    paid: money(paid),
    owing: money(owing),
    is_debtor: total > 0 && owing > 0.001
  };
}

function listDebtors() {
  return listOrders().map(decorateMoney).filter((o) => o.is_debtor);
}

function recordPayment(orderNumber, amount, note) {
  const order = state.orders.find((o) => o.order_number === String(orderNumber || "").trim());
  if (!order) throw new Error("Order not found");
  const add = parseMoney(amount);
  if (add <= 0) throw new Error("Payment amount must be more than 0");
  if (!Array.isArray(order.payments)) order.payments = [];
  order.payments.push({
    at: nowIso(),
    amount: money(add),
    note: String(note || "").trim()
  });
  order.amount_paid = money(orderPaid(order) + add);
  if (!order.payment_date) order.payment_date = nowIso().slice(0, 10);
  order.updated_at = nowIso();
  save();
  return decorateMoney(order);
}

module.exports = {
  db: null,
  dbPath,
  ORDER_FIELDS,
  DROPDOWN_KEYS,
  VAT_RATE,
  parseMoney,
  money,
  inclFromExcl,
  listOrders,
  upsertOrder,
  deleteOrder,
  listSchedule,
  upsertScheduleRow,
  setScheduleCell,
  countOrders,
  listDropdowns,
  addDropdownItem,
  removeDropdownItem,
  listDebtors,
  recordPayment,
  decorateMoney
};
