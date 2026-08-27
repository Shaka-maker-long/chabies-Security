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

const ORDER_SHEET_HEADERS = [
  "Quote Number", "Order Number", "Status", "Assigned Operator", "Type", "Catergory",
  "Product", "Variation", "Doors", "Detailed Description", "Dimensions", "Powder Coating",
  "Client Name", "Client Number", "Email Address", "Payment Date", "Address", "Province",
  "Price (Excl VAT)", "Price (Incl VAT)", "Amount Paid", "Month of Sale", "Source", "City"
];

const HEADER_MAP = {
  "quote number": "quote_number",
  "order number": "order_number",
  "order #": "order_number",
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

const REQUIRED_SHEETS = {
  ORDERS: ORDER_SHEET_HEADERS,
  Users: ["Name", "Role", "Password", "Tasks"],
  Production_Log: ["Id", "Order #", "Worker", "Role", "Process", "Start", "End", "Result", "Signature", "PauseStart", "PauseEnd", "PauseReason", "Meta"],
  Overview: ["Id", "Order #", "Worker", "Status", "Start", "End", "Notes"],
  Steel_Profiles: ["Category", "Profile Name"],
  Steel_Usage: ["Timestamp", "Order #", "Worker", "Process", "Profile Type", "Size / Length"],
  Backboards: ["Category", "Profile Name"],
  Backboard_Usage: ["Timestamp", "Order #", "Worker", "Process", "Type", "Size"],
  Idle_Alerts: ["Date", "Worker", "Role", "IdleSince", "AlertedAt", "Status", "AssignedTask"],
  Schedule: ["Id", "Worker", "Process", "Order", "Product", "Title", "Start", "End", "DurationMins", "Kind", "Seq", "EstimateSource"],
  Rates: ["Item", "Rate"]
};

function emptyState() {
  return {
    orders: [],
    schedule_rows: [],
    schedule_cells: [],
    nextOrderId: 1,
    nextScheduleId: 1,
    dropdowns: JSON.parse(JSON.stringify(DEFAULT_DROPDOWNS)),
    workbook: { sheets: {}, importedAt: null }
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
  const workbook = parsed.workbook && typeof parsed.workbook === "object"
    ? parsed.workbook
    : { sheets: {}, importedAt: null };
  if (!workbook.sheets || typeof workbook.sheets !== "object") workbook.sheets = {};
  state = {
    ...emptyState(),
    ...parsed,
    orders: Array.isArray(parsed.orders) ? parsed.orders : [],
    schedule_rows: Array.isArray(parsed.schedule_rows) ? parsed.schedule_rows : [],
    schedule_cells: Array.isArray(parsed.schedule_cells) ? parsed.schedule_cells : [],
    dropdowns,
    workbook
  };
  console.log(
    "[db] opened",
    dbPath,
    "orders",
    state.orders.length,
    "workbookSheets",
    Object.keys(state.workbook.sheets).length
  );
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
    applyOrderToWorkbook(existing);
    save();
    return existing;
  }
  payload.id = state.nextOrderId++;
  state.orders.push(payload);
  applyOrderToWorkbook(payload);
  save();
  return payload;
}

function deleteOrder(orderNumber) {
  state.orders = state.orders.filter((o) => o.order_number !== orderNumber);
  removeOrderFromWorkbook(orderNumber);
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

function cellStr(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

function cellDisplay(v) {
  if (v == null) return "";
  if (v instanceof Date) {
    const t = v.getTime();
    if (!Number.isFinite(t)) return "";
    return v.toISOString();
  }
  return String(v);
}

function normHeader(s) {
  return String(s || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function sheetMaxCol(grid) {
  return (grid || []).reduce((m, row) => Math.max(m, (row || []).length), 0);
}

function ensureSheet(title) {
  const wb = ensureWorkbookBare();
  if (!wb.sheets[title]) {
    const headers = REQUIRED_SHEETS[title] || ["Name"];
    wb.sheets[title] = {
      title,
      hidden: /Usage|Profiles|Backboards|Idle_Alerts/.test(title),
      lastRow: 1,
      lastCol: headers.length,
      grid: [headers.slice()]
    };
  }
  const sheet = wb.sheets[title];
  if (!Array.isArray(sheet.grid) || !sheet.grid.length) {
    const headers = REQUIRED_SHEETS[title] || ["Name"];
    sheet.grid = [headers.slice()];
    sheet.lastRow = 1;
    sheet.lastCol = headers.length;
  }
  return sheet;
}

function ensureWorkbookBare() {
  if (!state.workbook || typeof state.workbook !== "object") {
    state.workbook = { sheets: {}, importedAt: null };
  }
  if (!state.workbook.sheets || typeof state.workbook.sheets !== "object") {
    state.workbook.sheets = {};
  }
  return state.workbook;
}

function ensureWorkbook() {
  const wb = ensureWorkbookBare();
  for (const title of Object.keys(REQUIRED_SHEETS)) {
    if (!wb.sheets[title]) ensureSheet(title);
  }
  return wb;
}

function getWorkbookSnapshot() {
  return ensureWorkbook();
}

function workbookImportedAt() {
  return (state.workbook && state.workbook.importedAt) || null;
}

function hasLocalWorkbook() {
  const sheets = state.workbook && state.workbook.sheets;
  if (!sheets) return false;
  const orders = sheets.ORDERS;
  const users = sheets.Users;
  const hasOrders = !!(orders && Array.isArray(orders.grid) && orders.grid.length > 1);
  const hasUsers = !!(users && Array.isArray(users.grid) && users.grid.length > 1);
  return hasOrders || hasUsers || !!workbookImportedAt();
}

function markWorkbookImported() {
  ensureWorkbook().importedAt = nowIso();
  save();
}

function replaceWorkbook(snapshot, opts) {
  const imported = !opts || opts.imported !== false;
  state.workbook = {
    sheets: (snapshot && snapshot.sheets) || {},
    importedAt: imported ? nowIso() : (snapshot && snapshot.importedAt) || workbookImportedAt()
  };
  ensureWorkbook();
  syncOrdersFromWorkbook();
  save();
  return state.workbook;
}

function persistWorkbook() {
  ensureWorkbook();
  syncOrdersFromWorkbook();
  save();
}

function headerIndexMap(headerRow) {
  const map = {};
  (headerRow || []).forEach((h, i) => {
    const field = HEADER_MAP[normHeader(h)];
    if (field && map[field] == null) map[field] = i;
  });
  return map;
}

function applyOrderToWorkbook(order) {
  if (!order || !order.order_number) return;
  const sheet = ensureSheet("ORDERS");
  const grid = sheet.grid;
  if (!grid[0] || !grid[0].length) grid[0] = ORDER_SHEET_HEADERS.slice();
  const idx = headerIndexMap(grid[0]);
  let rowIndex = -1;
  const orderCol = idx.order_number != null ? idx.order_number : 1;
  for (let i = 1; i < grid.length; i++) {
    if (String((grid[i] || [])[orderCol] || "").trim() === String(order.order_number).trim()) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex < 0) {
    grid.push([]);
    rowIndex = grid.length - 1;
  }
  while (grid[rowIndex].length < grid[0].length) grid[rowIndex].push("");
  for (const field of ORDER_FIELDS) {
    const col = idx[field];
    if (col == null) continue;
    grid[rowIndex][col] = order[field] == null ? "" : String(order[field]);
  }
  sheet.lastRow = grid.length;
  sheet.lastCol = Math.max(sheet.lastCol || 0, sheetMaxCol(grid));
}

function removeOrderFromWorkbook(orderNumber) {
  const sheet = state.workbook && state.workbook.sheets && state.workbook.sheets.ORDERS;
  if (!sheet || !Array.isArray(sheet.grid) || sheet.grid.length < 2) return;
  const idx = headerIndexMap(sheet.grid[0]);
  const orderCol = idx.order_number != null ? idx.order_number : 1;
  const want = String(orderNumber || "").trim();
  sheet.grid = sheet.grid.filter((row, i) => i === 0 || String((row || [])[orderCol] || "").trim() !== want);
  sheet.lastRow = sheet.grid.length;
}

function rowFromSheet(headerRow, row) {
  const out = {};
  (headerRow || []).forEach((h, c) => {
    const field = HEADER_MAP[normHeader(h)];
    if (field) out[field] = cellDisplay((row || [])[c]);
  });
  return out;
}

function minutesBetween(start, end) {
  const a = start instanceof Date ? start : (start ? new Date(start) : null);
  const b = end instanceof Date ? end : (end ? new Date(end) : null);
  if (!a || !b || isNaN(a.getTime()) || isNaN(b.getTime())) return 0;
  return Math.max(0, Math.round((b.getTime() - a.getTime()) / 60000));
}

function pauseMinutes(metaRaw) {
  let meta = metaRaw;
  if (typeof metaRaw === "string" && metaRaw) {
    try { meta = JSON.parse(metaRaw); } catch (e) { return 0; }
  }
  if (!meta || !Array.isArray(meta.pauses)) return 0;
  let mins = 0;
  for (const p of meta.pauses) {
    mins += minutesBetween(p.start || p.from, p.end || p.to);
  }
  return mins;
}

function sheetRows(title) {
  const sheet = state.workbook && state.workbook.sheets && state.workbook.sheets[title];
  if (!sheet || !Array.isArray(sheet.grid) || sheet.grid.length < 2) return [];
  return sheet.grid.slice(1);
}

function syncOrdersFromWorkbook() {
  const sheet = ensureSheet("ORDERS");
  const grid = sheet.grid || [];
  if (grid.length < 2) {
    syncFloorOntoOrders();
    return;
  }
  const headers = grid[0];
  for (let i = 1; i < grid.length; i++) {
    const row = rowFromSheet(headers, grid[i]);
    const orderNumber = String(row.order_number || "").trim();
    if (!orderNumber) continue;
    const existing = state.orders.find((o) => o.order_number === orderNumber);
    const payload = {};
    for (const f of ORDER_FIELDS) payload[f] = row[f] == null ? "" : String(row[f]);
    payload.order_number = orderNumber;
    payload.updated_at = nowIso();
    if (existing) {
      if (!payload.amount_paid && existing.amount_paid) payload.amount_paid = existing.amount_paid;
      if (!payload.price_excl_vat && existing.price_excl_vat) payload.price_excl_vat = existing.price_excl_vat;
      if (!payload.price_incl_vat && existing.price_incl_vat) payload.price_incl_vat = existing.price_incl_vat;
      if (!Array.isArray(payload.payments)) payload.payments = existing.payments || [];
      Object.assign(existing, payload);
    } else {
      payload.id = state.nextOrderId++;
      payload.payments = [];
      if (!payload.amount_paid) payload.amount_paid = "0.00";
      state.orders.push(payload);
    }
  }
  syncFloorOntoOrders();
}

function syncFloorOntoOrders() {
  const logsByOrder = {};
  const steelByOrder = {};
  const backboardByOrder = {};

  for (const row of sheetRows("Production_Log")) {
    const orderNum = String(row[1] || "").trim();
    if (!orderNum) continue;
    const start = row[5] || "";
    const end = row[6] || "";
    const meta = row[12] || "";
    const duration = end ? Math.max(0, minutesBetween(start, end) - pauseMinutes(meta)) : 0;
    if (!logsByOrder[orderNum]) logsByOrder[orderNum] = [];
    logsByOrder[orderNum].push({
      id: cellStr(row[0]),
      worker: cellStr(row[2]),
      role: cellStr(row[3]),
      process: cellStr(row[4]),
      start: cellStr(start),
      end: cellStr(end),
      result: cellStr(row[7]),
      pause_reason: cellStr(row[11]),
      meta: cellStr(meta),
      duration_minutes: duration,
      open: !end
    });
  }

  for (const row of sheetRows("Steel_Usage")) {
    const orderNum = String(row[1] || "").trim();
    if (!orderNum) continue;
    if (!steelByOrder[orderNum]) steelByOrder[orderNum] = [];
    steelByOrder[orderNum].push({
      at: cellStr(row[0]),
      worker: cellStr(row[2]),
      process: cellStr(row[3]),
      type: cellStr(row[4]),
      size: cellStr(row[5])
    });
  }

  for (const row of sheetRows("Backboard_Usage")) {
    const orderNum = String(row[1] || "").trim();
    if (!orderNum) continue;
    if (!backboardByOrder[orderNum]) backboardByOrder[orderNum] = [];
    backboardByOrder[orderNum].push({
      at: cellStr(row[0]),
      worker: cellStr(row[2]),
      process: cellStr(row[3]),
      type: cellStr(row[4]),
      size: cellStr(row[5])
    });
  }

  for (const order of state.orders) {
    const key = String(order.order_number || "").trim();
    const logs = logsByOrder[key] || [];
    const steel = steelByOrder[key] || [];
    const backboard = backboardByOrder[key] || [];
    order.work_logs = logs;
    order.steel_usage = steel;
    order.backboard_usage = backboard;
    order.duration_minutes = logs.reduce((sum, log) => sum + (Number(log.duration_minutes) || 0), 0);
  }
}

module.exports = {
  db: null,
  dbPath,
  ORDER_FIELDS,
  ORDER_SHEET_HEADERS,
  HEADER_MAP,
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
  save,
  ensureWorkbook,
  ensureSheet,
  getWorkbookSnapshot,
  hasLocalWorkbook,
  workbookImportedAt,
  markWorkbookImported,
  replaceWorkbook,
  persistWorkbook,
  applyOrderToWorkbook,
  syncOrdersFromWorkbook,
  syncFloorOntoOrders,
  normHeader
};
