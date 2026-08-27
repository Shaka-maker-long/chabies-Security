const fs = require("fs");
const path = require("path");
const { DROPDOWN_KEYS, DEFAULT_DROPDOWNS } = require("./dropdowns-default");
const { getBook, persistWorkbook, ORDER_HEADERS } = require("./workbook-store");

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
    dropdowns: JSON.parse(JSON.stringify(DEFAULT_DROPDOWNS)),
    paymentsByOrder: {}
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
    dropdowns,
    paymentsByOrder: parsed.paymentsByOrder && typeof parsed.paymentsByOrder === "object" ? parsed.paymentsByOrder : {}
  };
  if (!Object.keys(state.paymentsByOrder).length && Array.isArray(parsed.orders)) {
    parsed.orders.forEach((o) => {
      if (o && o.order_number && Array.isArray(o.payments) && o.payments.length) {
        state.paymentsByOrder[o.order_number] = o.payments;
      }
    });
  }
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

const ORDER_HEADER_MAP = {
  "quote number": "quote_number",
  "order number": "order_number",
  "status": "status",
  "assigned operator": "assigned_operator",
  "type": "type",
  "catergory": "category",
  "category": "category",
  "product": "product",
  "variation": "variation",
  "doors": "doors",
  "detailed description": "detailed_description",
  "dimensions": "dimensions",
  "powder coating": "powder_coating",
  "client name": "client_name",
  "client number": "client_number",
  "email address": "email",
  "email": "email",
  "payment date": "payment_date",
  "address": "address",
  "province": "province",
  "price (incl vat)": "price_incl_vat",
  "price (excl vat)": "price_excl_vat",
  "amount paid": "amount_paid",
  "month of sale": "month_of_sale",
  "source": "source",
  "city": "city"
};

function normHeader(s) {
  return String(s || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function cellStr(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

function ordersSheet() {
  const book = getBook();
  let sheet = book.getSheetByName("ORDERS");
  if (!sheet) sheet = book.insertSheet("ORDERS");
  return sheet;
}

function headerLookup(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), ORDER_HEADERS.length, 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const idx = {};
  headers.forEach((h, i) => {
    const field = ORDER_HEADER_MAP[normHeader(h)];
    if (field) idx[field] = i;
  });
  return { headers, idx, lastCol };
}

function ensureOrderHeaders(sheet) {
  let look = headerLookup(sheet);
  if (look.idx.order_number == null || look.idx.status == null) {
    sheet.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS]);
    look = headerLookup(sheet);
  }
  if (look.idx.amount_paid == null) {
    const col = Math.max(sheet.getLastColumn(), look.headers.length) + 1;
    sheet.getRange(1, col).setValue("AMOUNT PAID");
    look = headerLookup(sheet);
  }
  return look;
}

function rowToOrder(row, idx, id) {
  const o = { id };
  for (const f of ORDER_FIELDS) {
    const i = idx[f];
    o[f] = i == null ? "" : cellStr(row[i]);
  }
  return o;
}

function listOrders() {
  const sheet = ordersSheet();
  const { idx, lastCol } = ensureOrderHeaders(sheet);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const grid = sheet.getRange(2, 1, last - 1, lastCol).getValues();
  const out = [];
  for (let i = 0; i < grid.length; i++) {
    const row = rowToOrder(grid[i], idx, i + 2);
    if (!String(row.order_number || "").trim()) continue;
    if (state.paymentsByOrder && state.paymentsByOrder[row.order_number]) {
      row.payments = state.paymentsByOrder[row.order_number];
    }
    out.push(row);
  }
  return out.reverse();
}

function findOrderSheetRow(sheet, idx, orderNumber) {
  const last = sheet.getLastRow();
  if (last < 2 || idx.order_number == null) return 0;
  const values = sheet.getRange(2, idx.order_number + 1, last - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === orderNumber) return i + 2;
  }
  return 0;
}

function upsertOrder(row) {
  const orderNumber = String(row.order_number || "").trim();
  if (!orderNumber) throw new Error("Order number is required");
  const payload = {};
  for (const f of ORDER_FIELDS) payload[f] = row[f] == null ? "" : String(row[f]);
  payload.order_number = orderNumber;
  payload.updated_at = nowIso();
  const existing = listOrders().find((o) => o.order_number === orderNumber);
  applyPriceAndPayments(payload, row, existing);
  const sheet = ordersSheet();
  const { idx, lastCol } = ensureOrderHeaders(sheet);
  let rowNum = findOrderSheetRow(sheet, idx, orderNumber);
  if (!rowNum) rowNum = sheet.getLastRow() + 1;
  const width = Math.max(lastCol, 1);
  const current = rowNum <= sheet.getLastRow()
    ? sheet.getRange(rowNum, 1, 1, width).getValues()[0]
    : [];
  while (current.length < width) current.push("");
  for (const f of ORDER_FIELDS) {
    if (idx[f] == null) continue;
    current[idx[f]] = payload[f] == null ? "" : payload[f];
  }
  sheet.getRange(rowNum, 1, 1, current.length).setValues([current]);
  if (!state.paymentsByOrder) state.paymentsByOrder = {};
  state.paymentsByOrder[orderNumber] = payload.payments || [];
  save();
  persistWorkbook();
  payload.id = rowNum;
  payload.payments = state.paymentsByOrder[orderNumber];
  return payload;
}

function deleteOrder(orderNumber) {
  const sheet = ordersSheet();
  const { idx } = ensureOrderHeaders(sheet);
  const rowNum = findOrderSheetRow(sheet, idx, String(orderNumber || "").trim());
  if (rowNum) {
    sheet.deleteRow(rowNum);
    persistWorkbook();
  }
  if (state.paymentsByOrder) delete state.paymentsByOrder[orderNumber];
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
  return listOrders().length;
}

function migrateJsonOrdersToWorkbook() {
  const sheet = ordersSheet();
  ensureOrderHeaders(sheet);
  if (sheet.getLastRow() >= 2) return { migrated: 0 };
  const leftover = (state.orders || []).filter((o) => o && String(o.order_number || "").trim());
  leftover.forEach((o) => {
    try {
      upsertOrder(o);
    } catch (e) {
      console.error("[db] migrate order failed", o && o.order_number, e && e.message ? e.message : e);
    }
  });
  if (leftover.length) {
    console.log("[db] migrated", leftover.length, "json orders into Railway workbook");
  }
  return { migrated: leftover.length };
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
  const num = String(orderNumber || "").trim();
  const existing = listOrders().find((o) => o.order_number === num);
  if (!existing) throw new Error("Order not found");
  const add = parseMoney(amount);
  if (add <= 0) throw new Error("Payment amount must be more than 0");
  if (!state.paymentsByOrder) state.paymentsByOrder = {};
  const history = Array.isArray(state.paymentsByOrder[num])
    ? state.paymentsByOrder[num].slice()
    : (Array.isArray(existing.payments) ? existing.payments.slice() : []);
  history.push({
    at: nowIso(),
    amount: money(add),
    note: String(note || "").trim()
  });
  state.paymentsByOrder[num] = history;
  const saved = upsertOrder({
    ...existing,
    amount_paid: money(orderPaid(existing) + add),
    payment_date: existing.payment_date || nowIso().slice(0, 10),
    payments: history
  });
  return decorateMoney(saved);
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
  decorateMoney,
  migrateJsonOrdersToWorkbook
};
