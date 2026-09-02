const fs = require("fs");
const path = require("path");
const { DROPDOWN_KEYS, DEFAULT_DROPDOWNS } = require("./dropdowns-default");
const {
  unique,
  ENQUIRY_FIELDS,
  ENQUIRY_DROPDOWN_KEYS,
  DEFAULT_ENQUIRY_DROPDOWNS,
  NEW_DESIGN_MIN_CHARS
} = require("./enquiries-default");
const { getBook, persistWorkbook, ORDER_HEADERS, dataDir } = require("./workbook-store");

const ORDER_FIELDS = [
  "quote_number", "order_number", "status", "assigned_operator", "type", "category",
  "product", "variation", "doors", "detailed_description", "dimensions", "powder_coating",
  "client_name", "client_number", "email", "payment_date", "address", "province",
  "price_excl_vat", "price_incl_vat", "amount_paid", "month_of_sale", "source", "city"
];

const VAT_RATE = 0.15;

const SAST_OFFSET_MS = 2 * 60 * 60 * 1000;
const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];

function parseMoney(s) {
  const n = Number(String(s || "").replace(/,/g, "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0;
}

function money(n) {
  return (Math.round(Number(n) * 100) / 100).toFixed(2);
}

function formatRand(n) {
  const v = parseMoney(n);
  const neg = v < 0 ? "-" : "";
  const [whole, frac] = money(Math.abs(v)).split(".");
  const grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return neg + "R " + grouped + "." + frac;
}

function asDate(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === "number" && isFinite(v) && v >= 20000 && v <= 120000) {
    const utcMs = Math.round((v - 25569) * 86400000);
    return new Date(utcMs - SAST_OFFSET_MS);
  }
  const s = String(v || "").trim();
  if (!s) return null;
  const dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmy) {
    return new Date(Date.UTC(Number(dmy[3]), Number(dmy[2]) - 1, Number(dmy[1])) - SAST_OFFSET_MS);
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(s) || /^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function sastParts(d) {
  const sast = new Date(d.getTime() + SAST_OFFSET_MS);
  return { y: sast.getUTCFullYear(), m: sast.getUTCMonth(), day: sast.getUTCDate() };
}

function dateToSerial(d) {
  return Math.round((d.getTime() + SAST_OFFSET_MS) / 86400000 + 25569);
}

function looksLikeConvertedDate(v) {
  if (v instanceof Date) return true;
  const s = String(v || "");
  return /^\d{4}-\d{2}-\d{2}T/.test(s) || /^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}/.test(s);
}

function formatOrderId(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number" && isFinite(v)) return String(Math.round(v));
  if (looksLikeConvertedDate(v)) {
    const d = asDate(v);
    if (d) return String(dateToSerial(d));
  }
  return String(v).trim();
}

function formatPaymentDate(v) {
  if (v == null || v === "") return "";
  const d = asDate(v);
  if (!d) return String(v).replace(/T.*$/, "").trim();
  const p = sastParts(d);
  return String(p.day).padStart(2, "0") + "/" + String(p.m + 1).padStart(2, "0") + "/" + p.y;
}

function formatMonthOfSale(v) {
  if (v == null || v === "") return "";
  const s = String(v).trim();
  if (/^[A-Za-z]+ \d{4}$/.test(s)) return s;
  const d = asDate(v);
  if (!d) return s.replace(/T.*$/, "");
  const p = sastParts(d);
  return MONTH_NAMES[p.m] + " " + p.y;
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
  if (payload.price_excl_vat) {
    payload.price_excl_vat = money(parseMoney(payload.price_excl_vat));
    payload.price_incl_vat = inclFromExcl(payload.price_excl_vat);
  }
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
    paymentsByOrder: {},
    enquiries: [],
    enquiry_dropdowns: {}
  };
}

function pickDataFile() {
  const preferred = process.env.OFFICE_DB_PATH || path.join(dataDir(), "studio-delta.json");
  const fallback = path.join("/tmp", "studio-delta.json");
  for (const candidate of [preferred, fallback]) {
    try {
      fs.mkdirSync(path.dirname(candidate), { recursive: true });
      fs.accessSync(path.dirname(candidate), fs.constants.W_OK);
      if (candidate === fallback && preferred !== fallback) {
        console.error("[db] WARNING office data falling back to", candidate, "— this will be lost on deploy. Mount a Railway volume and set DATA_DIR.");
      }
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
    paymentsByOrder: parsed.paymentsByOrder && typeof parsed.paymentsByOrder === "object" ? parsed.paymentsByOrder : {},
    enquiries: Array.isArray(parsed.enquiries) ? parsed.enquiries : [],
    enquiry_dropdowns: parsed.enquiry_dropdowns && typeof parsed.enquiry_dropdowns === "object" ? parsed.enquiry_dropdowns : {}
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
    try {
      const sqlite = require("./sqlite-store");
      const fromSql = sqlite.loadOffice();
      if (fromSql && (fromSql.enquiries || []).length) {
        const dropdowns = JSON.parse(JSON.stringify(DEFAULT_DROPDOWNS));
        if (fromSql.dropdowns && typeof fromSql.dropdowns === "object") {
          for (const key of DROPDOWN_KEYS) {
            if (Array.isArray(fromSql.dropdowns[key])) dropdowns[key] = fromSql.dropdowns[key];
          }
        }
        state = { ...emptyState(), ...fromSql, dropdowns };
        console.log("[db] opened SQLite", sqlite.sqlitePath(), "enquiries", state.enquiries.length);
      } else if (fromSql && (Object.keys(fromSql.dropdowns || {}).length || (fromSql.schedule_rows || []).length)) {
        const dropdowns = JSON.parse(JSON.stringify(DEFAULT_DROPDOWNS));
        if (fromSql.dropdowns && typeof fromSql.dropdowns === "object") {
          for (const key of DROPDOWN_KEYS) {
            if (Array.isArray(fromSql.dropdowns[key])) dropdowns[key] = fromSql.dropdowns[key];
          }
        }
        state = { ...emptyState(), ...fromSql, dropdowns };
        console.log("[db] opened SQLite", sqlite.sqlitePath(), "dropdowns");
      } else {
        console.log("[db] new file", dbPath);
      }
    } catch (err) {
      console.log("[db] new file", dbPath);
    }
  }
}
try {
  require("./sqlite-store").saveOffice(state);
} catch (e) {}

function save() {
  const tmp = dbPath + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(state));
  fs.renameSync(tmp, dbPath);
  try { require("./sqlite-store").saveOffice(state); } catch (e) {
    console.error("[db] sqlite office save failed", e && e.message ? e.message : e);
  }
}

function persistenceInfo() {
  const { storageInfo } = require("./workbook-store");
  const info = storageInfo();
  const preferred = process.env.OFFICE_DB_PATH || path.join(dataDir(), "studio-delta.json");
  const fellBack = dbPath !== preferred;
  let sqlite = {};
  try { sqlite = require("./sqlite-store").info(); } catch (e) { sqlite = {}; }
  return {
    ...info,
    ...sqlite,
    officeDb: dbPath,
    officeDbExists: fs.existsSync(dbPath),
    enquiryCount: (state.enquiries || []).length,
    usingEphemeralDisk: info.usingEphemeralDisk || fellBack,
    warning: info.warning || (fellBack
      ? "Office file could not be written to " + preferred + " and is using " + dbPath + " instead."
      : null)
  };
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
  if (v instanceof Date) return formatPaymentDate(v);
  return String(v);
}

function formatOrderField(field, v) {
  if (field === "quote_number" || field === "order_number") return formatOrderId(v);
  if (field === "payment_date") return formatPaymentDate(v);
  if (field === "month_of_sale") return formatMonthOfSale(v);
  if (field === "price_excl_vat" || field === "price_incl_vat" || field === "amount_paid") {
    if (v === "" || v == null) return "";
    return money(parseMoney(v));
  }
  return cellStr(v);
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
    o[f] = i == null ? "" : formatOrderField(f, row[i]);
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
  const want = formatOrderId(orderNumber);
  const values = sheet.getRange(2, idx.order_number + 1, last - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (formatOrderId(values[i][0]) === want) return i + 2;
  }
  return 0;
}

function normalizeOrdersSheet() {
  const sheet = ordersSheet();
  const { idx, lastCol } = ensureOrderHeaders(sheet);
  const last = sheet.getLastRow();
  if (last < 2) return { rewritten: 0 };
  const grid = sheet.getRange(2, 1, last - 1, lastCol).getValues();
  let rewritten = 0;
  const fields = ["quote_number", "order_number", "payment_date", "month_of_sale"];
  for (let i = 0; i < grid.length; i++) {
    fields.forEach((f) => {
      if (idx[f] == null) return;
      const cur = grid[i][idx[f]];
      const next = formatOrderField(f, cur);
      const same = !(cur instanceof Date) && String(cur || "") === next;
      if (!same) {
        grid[i][idx[f]] = next;
        rewritten++;
      }
    });
  }
  if (rewritten) {
    sheet.getRange(2, 1, last - 1, lastCol).setValues(grid);
    persistWorkbook();
  }
  return { rewritten };
}

function upsertOrder(row) {
  const orderNumber = formatOrderId(row.order_number);
  if (!orderNumber) throw new Error("Order number is required");
  const payload = {};
  for (const f of ORDER_FIELDS) payload[f] = row[f] == null ? "" : String(row[f]);
  payload.quote_number = formatOrderId(payload.quote_number);
  payload.order_number = orderNumber;
  payload.payment_date = payload.payment_date ? formatPaymentDate(payload.payment_date) : "";
  payload.month_of_sale = payload.month_of_sale ? formatMonthOfSale(payload.month_of_sale) : "";
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
    price_excl_vat: order.price_excl_vat ? formatRand(order.price_excl_vat) : "",
    price_incl_vat: order.price_incl_vat ? formatRand(order.price_incl_vat) : "",
    amount_paid: formatRand(order.amount_paid || 0),
    total: formatRand(total),
    paid: formatRand(paid),
    owing: formatRand(owing),
    is_debtor: total > 0 && owing > 0.001
  };
}

function listDebtors() {
  return listOrders().map(decorateMoney).filter((o) => o.is_debtor);
}

const MONTH_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const FIRST_ENQUIRY_NO = 1996;

function enquiryNumberValue(raw) {
  const m = String(raw || "").trim().match(/#?\s*(\d+)/);
  return m ? Number(m[1]) : 0;
}

function formatEnquiryNo(n) {
  return "#" + String(n);
}

function nextEnquiryNo() {
  let max = FIRST_ENQUIRY_NO - 1;
  for (const row of state.enquiries || []) {
    const n = enquiryNumberValue(row && row.enquiry_no);
    if (n > max) max = n;
  }
  return formatEnquiryNo(max + 1);
}

function quoteNumberValue(raw) {
  const m = String(raw || "").trim().toUpperCase().replace(/\s+/g, "").match(/^(?:SOQ)?(\d+)$/);
  return m ? Number(m[1]) : 0;
}

function normalizeQuoteNo(raw, opts) {
  const s = String(raw || "").trim().toUpperCase().replace(/\s+/g, "");
  if (!s) {
    if (opts && opts.allowEmpty) return "";
    throw new Error("Enter a quotation number");
  }
  const n = quoteNumberValue(s);
  if (!n) throw new Error("Quotation number must look like SOQ2361");
  return "SOQ" + n;
}

function collectQuoteNumbers(exceptEnquiryNo) {
  const except = enquiryNumberValue(exceptEnquiryNo);
  const out = [];
  const seen = new Set();
  function push(raw, source) {
    const quoteNo = normalizeQuoteNo(raw, { allowEmpty: true });
    if (!quoteNo) return;
    const key = quoteNo.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    out.push({ quote_no: quoteNo, n: quoteNumberValue(quoteNo), source: source || "" });
  }
  for (const row of state.enquiries || []) {
    const isExcept = except && enquiryNumberValue(row && row.enquiry_no) === except;
    const current = normalizeQuoteNo(row && row.quote_no, { allowEmpty: true });
    if (!isExcept) push(row && row.quote_no, row && row.enquiry_no);
    const quotes = Array.isArray(row && row.quotes) ? row.quotes : [];
    for (const q of quotes) {
      const qn = normalizeQuoteNo(q && q.quote_no, { allowEmpty: true });
      if (isExcept && qn && current && qn === current) continue;
      push(q && q.quote_no, row && row.enquiry_no);
    }
  }
  try {
    const sheet = ordersSheet();
    const { idx } = ensureOrderHeaders(sheet);
    const last = sheet.getLastRow();
    if (idx.quote_number != null && last >= 2) {
      const values = sheet.getRange(2, idx.quote_number + 1, last - 1, 1).getValues();
      for (let i = 0; i < values.length; i++) push(values[i][0], "order");
    }
  } catch (e) {}
  out.sort((a, b) => a.n - b.n);
  return out;
}

function recentQuoteNos(limit) {
  const n = Number(limit) > 0 ? Number(limit) : 5;
  return collectQuoteNumbers().slice(-n).map((x) => x.quote_no);
}

function nextQuoteNo() {
  const nums = collectQuoteNumbers().map((x) => x.n).filter((n) => n > 0);
  const max = nums.length ? Math.max.apply(null, nums) : 0;
  return "SOQ" + (max + 1);
}

function requireUniqueQuoteNo(raw, enquiryNo) {
  const quoteNo = normalizeQuoteNo(raw);
  const used = collectQuoteNumbers(enquiryNo).some((x) => x.quote_no === quoteNo);
  if (used) throw new Error("Quotation number " + quoteNo + " is already used");
  return quoteNo;
}

function quoteNoHint() {
  const recent = recentQuoteNos(5);
  return {
    next: nextQuoteNo(),
    recent
  };
}

function monthFromEnquiryDate(v) {
  const d = asDate(v);
  if (!d) return "";
  return MONTH_SHORT[sastParts(d).m];
}

function listEnquiryDropdowns() {
  const saved = state.enquiry_dropdowns && typeof state.enquiry_dropdowns === "object" ? state.enquiry_dropdowns : {};
  const out = {};
  for (const key of ENQUIRY_DROPDOWN_KEYS) {
    out[key] = unique([
      ...(DEFAULT_ENQUIRY_DROPDOWNS[key] || []),
      ...(Array.isArray(saved[key]) ? saved[key] : [])
    ]);
  }
  return out;
}

function addEnquiryDropdownItem(field, value) {
  if (ENQUIRY_DROPDOWN_KEYS.indexOf(field) === -1) throw new Error("Unknown enquiry dropdown");
  const item = String(value || "").trim();
  if (!item) throw new Error("Value is required");
  if (!state.enquiry_dropdowns || typeof state.enquiry_dropdowns !== "object") state.enquiry_dropdowns = {};
  if (!Array.isArray(state.enquiry_dropdowns[field])) state.enquiry_dropdowns[field] = [];
  const all = listEnquiryDropdowns()[field] || [];
  const exists = all.some((v) => String(v).toLowerCase() === item.toLowerCase());
  if (!exists) state.enquiry_dropdowns[field].push(item);
  save();
  return listEnquiryDropdowns();
}

function emptyCustomSpec() {
  return { kind: "", other: "", detail: "" };
}

function normalizeCustomSpecs(row, existing) {
  let specs = Array.isArray(row && row.custom_specs) ? row.custom_specs : null;
  if (!specs && existing && Array.isArray(existing.custom_specs)) specs = existing.custom_specs;
  if (!specs) return [];
  const cleaned = [];
  for (const spec of specs) {
    let kind = String((spec && spec.kind) || "").trim();
    const other = String((spec && spec.other) || "").trim();
    const detail = String((spec && spec.detail) || "").trim();
    if (/^other$/i.test(kind) && other) kind = other;
    if (!kind && !other && !detail) continue;
    cleaned.push({ kind, other: /^other$/i.test(String((spec && spec.kind) || "")) ? other : "", detail });
  }
  return cleaned;
}

function customSpecSummary(specs) {
  return (specs || [])
    .filter((s) => s.kind && s.detail)
    .map((s) => s.kind + ": " + s.detail)
    .join("; ");
}

function rememberCustomSpecKinds(specs) {
  for (const spec of specs || []) {
    const kind = String(spec.kind || "").trim();
    if (!kind || /^other$/i.test(kind)) continue;
    addEnquiryDropdownItem("custom_spec", kind);
  }
}

function applyEnquiryTypeDetails(payload, row, existing) {
  const type = payload.enquiry_type;
  payload.custom_specs = normalizeCustomSpecs(row, existing);
  payload.design_description = String(
    row.design_description != null ? row.design_description : (existing && existing.design_description) || ""
  ).trim();

  if (type === "Custom") {
    payload.design_description = "";
    if (!payload.custom_specs.length) {
      throw new Error("For Custom, say whether it is Dimensions, Colour, or Other");
    }
    for (const spec of payload.custom_specs) {
      if (!spec.kind || /^other$/i.test(spec.kind)) {
        throw new Error("If it is Other, specify what the custom change is");
      }
      if (!spec.detail) {
        throw new Error("Write the " + spec.kind.toLowerCase() + " for this custom enquiry");
      }
    }
    rememberCustomSpecKinds(payload.custom_specs);
    payload.request = customSpecSummary(payload.custom_specs);
    return;
  }

  if (type === "New Design") {
    payload.custom_specs = [];
    if (payload.design_description.length < NEW_DESIGN_MIN_CHARS) {
      throw new Error("For a New Design, write a full description of what it is");
    }
    payload.request = payload.design_description;
    return;
  }

  payload.custom_specs = [];
  payload.design_description = "";
}

function enquiryQuoteKey(enquiryNo) {
  const n = enquiryNumberValue(enquiryNo);
  if (!n) throw new Error("Missing enquiry number");
  return String(n);
}

function enquiryQuotePdfPath(enquiryNo) {
  const dir = path.join(path.dirname(dbPath), "enquiry-quotes");
  fs.mkdirSync(dir, { recursive: true });
  return path.join(dir, enquiryQuoteKey(enquiryNo) + ".pdf");
}

function enquiryHasQuotePdf(enquiryNo) {
  try {
    return fs.existsSync(enquiryQuotePdfPath(enquiryNo));
  } catch (e) {
    return false;
  }
}

function decodeEnquiryPdf(raw) {
  const s = String(raw || "").trim();
  if (!s) return null;
  const b64 = s.replace(/^data:application\/pdf;base64,/i, "");
  const buf = Buffer.from(b64, "base64");
  if (buf.length < 5 || buf.slice(0, 5).toString("utf8") !== "%PDF-") {
    throw new Error("Quote file must be a PDF");
  }
  return buf;
}

function saveEnquiryQuotePdf(enquiryNo, raw, filename) {
  const buf = Buffer.isBuffer(raw) ? raw : decodeEnquiryPdf(raw);
  if (!buf) throw new Error("Quote PDF is missing");
  if (buf.length < 5 || buf.slice(0, 5).toString("utf8") !== "%PDF-") {
    throw new Error("Quote file must be a PDF");
  }
  fs.writeFileSync(enquiryQuotePdfPath(enquiryNo), buf);
  return {
    quote_pdf_name: String(filename || "quote.pdf").replace(/[^\w.\- ()]/g, "").slice(0, 120) || "quote.pdf",
    quote_pdf_uploaded_at: nowIso()
  };
}

function readEnquiryQuotePdf(enquiryNo) {
  const file = enquiryQuotePdfPath(enquiryNo);
  if (!fs.existsSync(file)) return null;
  const n = enquiryNumberValue(enquiryNo);
  const row = (state.enquiries || []).find((r) => enquiryNumberValue(r.enquiry_no) === n) || {};
  return {
    buffer: fs.readFileSync(file),
    filename: row.quote_pdf_name || (enquiryQuoteKey(enquiryNo) + ".pdf")
  };
}

function removeEnquiryQuotePdf(enquiryNo) {
  try {
    const file = enquiryQuotePdfPath(enquiryNo);
    if (fs.existsSync(file)) fs.unlinkSync(file);
  } catch (e) {}
}

function todayEnquiryDate() {
  const p = sastParts(new Date());
  return String(p.day).padStart(2, "0") + "/" + String(p.m + 1).padStart(2, "0") + "/" + p.y;
}

function formatSastDateTime(v) {
  const d = v instanceof Date ? v : (asDate(v) || (v ? new Date(v) : null));
  if (!d || isNaN(d.getTime())) return "";
  const sast = new Date(d.getTime() + SAST_OFFSET_MS);
  const p = (n) => String(n).padStart(2, "0");
  return p(sast.getUTCDate()) + "/" + p(sast.getUTCMonth() + 1) + "/" + sast.getUTCFullYear() +
    " " + p(sast.getUTCHours()) + ":" + p(sast.getUTCMinutes());
}

function durationLabel(ms) {
  if (!Number.isFinite(ms) || ms < 0) return "";
  const mins = Math.round(ms / 60000);
  if (mins < 60) return Math.max(1, mins) + " min";
  const hours = Math.floor(mins / 60);
  const rem = mins % 60;
  if (hours < 48) return hours + "h" + (rem ? " " + rem + "m" : "");
  const days = Math.floor(hours / 24);
  const h = hours % 24;
  return days + "d" + (h ? " " + h + "h" : "");
}

function nextEventId(events) {
  const max = (events || []).reduce((m, ev) => {
    const n = Number(String(ev && ev.id || "").replace(/\D/g, "")) || 0;
    return Math.max(m, n);
  }, 0);
  return "ev" + (max + 1);
}

function normalizeEnquiryEvents(row) {
  const list = Array.isArray(row && row.events) ? row.events : [];
  return list.map((ev, i) => {
    const at = String((ev && ev.at) || "").trim();
    return {
      id: String((ev && ev.id) || "ev" + (i + 1)),
      at,
      at_label: formatSastDateTime(at) || at,
      kind: String((ev && ev.kind) || "event"),
      actor: String((ev && ev.actor) || "").trim(),
      status: String((ev && ev.status) || "").trim(),
      from_status: String((ev && ev.from_status) || "").trim(),
      label: String((ev && ev.label) || (ev && ev.kind) || "Event"),
      note: String((ev && ev.note) || "").trim()
    };
  }).filter((ev) => ev.at || ev.label);
}

function synthesizeEnquiryEvents(row) {
  const events = [];
  const createdAt = row && row.created_at
    ? row.created_at
    : (row && asDate(row.date_enquired) ? asDate(row.date_enquired).toISOString() : "");
  if (createdAt) {
    events.push({
      id: "ev1",
      at: createdAt,
      kind: "created",
      actor: "",
      status: "New",
      from_status: "",
      label: "Enquiry captured",
      note: ""
    });
  }
  if (row && row.date_quoted) {
    const quotedAt = row.quote_pdf_uploaded_at || (asDate(row.date_quoted) ? asDate(row.date_quoted).toISOString() : "");
    if (quotedAt) {
      events.push({
        id: "ev" + (events.length + 1),
        at: quotedAt,
        kind: "complete_quote",
        actor: "",
        status: "Quoted",
        from_status: "",
        label: "Quote PDF issued" + (row.quote_no ? " " + row.quote_no : ""),
        note: ""
      });
    }
  }
  const status = String((row && row.status) || "");
  if (status === "Ordered" || status === "Rejected" || status === "Not Interested" || status === "Not within scope") {
    const at = (row && row.updated_at) || nowIso();
    events.push({
      id: "ev" + (events.length + 1),
      at,
      kind: status === "Ordered" ? "complete_order" : "close",
      actor: "",
      status,
      from_status: "",
      label: status === "Ordered" ? "Ordered" : ("Closed: " + status),
      note: ""
    });
  }
  return normalizeEnquiryEvents({ events });
}

function appendEnquiryEvent(row, partial) {
  if (!row) return null;
  const events = normalizeEnquiryEvents(row);
  const at = nowIso();
  const event = {
    id: nextEventId(events),
    at,
    kind: String((partial && partial.kind) || "event"),
    actor: String((partial && partial.actor) || "").trim(),
    status: String((partial && partial.status) != null ? partial.status : (row.status || "")),
    from_status: String((partial && partial.from_status) || "").trim(),
    label: String((partial && partial.label) || (partial && partial.kind) || "Event"),
    note: String((partial && partial.note) || "").trim()
  };
  events.push(event);
  row.events = events;
  return event;
}

function captureFieldsChanged(existing, payload) {
  if (!existing) return true;
  const keys = [
    "client_name", "enquiry_type", "enquiry_source", "source", "province",
    "client_email", "client_number", "comment", "date_enquired", "design_description"
  ];
  for (const k of keys) {
    if (String(existing[k] || "").trim() !== String(payload[k] || "").trim()) return true;
  }
  const names = (list) => (list || []).map((p) => String((p && p.product) || "").trim() + "|" + String((p && p.category) || "").trim()).filter((s) => s !== "|").join(";");
  if (names(existing.products) !== names(payload.products)) return true;
  const specs = (list) => JSON.stringify(list || []);
  if (specs(existing.custom_specs) !== specs(payload.custom_specs)) return true;
  return false;
}

function enquiryLifespan(row, events) {
  const list = events && events.length ? events : normalizeEnquiryEvents(row);
  const first = list[0];
  const start = first && first.at ? Date.parse(first.at) : Date.parse(row && row.created_at || "");
  const ordered = list.filter((ev) => ev.kind === "complete_order" || ev.status === "Ordered").slice(-1)[0];
  const closed = list.filter((ev) => ev.kind === "close" || ev.kind === "complete_reject" || /Rejected|Not Interested|Not within scope/.test(ev.status)).slice(-1)[0];
  const endEvent = ordered || closed;
  const end = endEvent && endEvent.at ? Date.parse(endEvent.at) : Date.now();
  const ms = Number.isFinite(start) ? end - start : NaN;
  const label = durationLabel(ms);
  let lifespan_label = "";
  if (label && ordered) lifespan_label = label + " to order";
  else if (label && closed) lifespan_label = label + " to closed";
  else if (label) lifespan_label = label + " open";
  return {
    opened_at: first && first.at ? first.at : (row && row.created_at) || "",
    opened_at_label: first && first.at_label ? first.at_label : formatSastDateTime((row && row.created_at) || "") ,
    ordered_at: ordered && ordered.at ? ordered.at : "",
    ordered_at_label: ordered && ordered.at_label ? ordered.at_label : "",
    lifespan_ms: Number.isFinite(ms) ? ms : 0,
    lifespan_label
  };
}

function emptyEnquiryLine() {
  return { product: "", category: "", value_excl_vat: "" };
}

function normalizeEnquiryLines(row, existing) {
  let lines = Array.isArray(row && row.products) ? row.products : null;
  if (!lines && existing && Array.isArray(existing.products)) lines = existing.products;
  if (!lines) {
    const product = String((row && row.product) || (existing && existing.product) || "").trim();
    const category = String((row && row.category) || (existing && existing.category) || "").trim();
    const value = (row && row.value_excl_vat) || "";
    lines = product || category ? [{ product, category, value_excl_vat: value }] : [emptyEnquiryLine()];
  }
  const cleaned = [];
  for (const line of lines) {
    const product = String((line && line.product) || "").trim();
    const category = String((line && line.category) || "").trim();
    const rawVal = line && line.value_excl_vat != null ? String(line.value_excl_vat).trim() : "";
    const value = rawVal === "" ? "" : money(parseMoney(rawVal));
    if (!product && !category && !value) continue;
    cleaned.push({ product, category, value_excl_vat: value });
  }
  if (!cleaned.length) cleaned.push(emptyEnquiryLine());
  return cleaned;
}

const KEEP_VALUE_STATUSES = [
  "Quoted", "Followed Up", "Ordered"
];
const CAPTURE_STATUSES = [
  "New",
  "Waiting on clients personal details",
  "Waiting on clients specifictions",
  "Waiting on productions confirmation"
];

function cloneJson(v, fallback) {
  if (v == null) return fallback;
  try {
    return JSON.parse(JSON.stringify(v));
  } catch (e) {
    return fallback;
  }
}

function normalizeEnquiryTasks(row) {
  const list = Array.isArray(row && row.tasks) ? row.tasks : [];
  return list.map((t, i) => ({
    id: String((t && t.id) || ("t" + (i + 1))),
    kind: String((t && t.kind) || "").trim(),
    title: String((t && t.title) || "").trim(),
    assignee: String((t && t.assignee) || "").trim(),
    status: String((t && t.status) || "open").trim() || "open",
    created_at: (t && t.created_at) || "",
    completed_at: (t && t.completed_at) || "",
    completed_by: (t && t.completed_by) || "",
    due_at: (t && t.due_at) || "",
    note: String((t && t.note) || "").trim(),
    label: String((t && t.label) || "").trim()
  }));
}

function parseCorrespondenceName(filename) {
  const raw = String(filename || "").trim();
  const base = raw.replace(/\.(msg|eml)$/i, "");
  const m = base.match(/Re[_:]?\s*Order\s*#?\s*([A-Za-z]?\d+)\s*[-–]\s*(.+)/i);
  return {
    title: base || raw,
    order_no: m ? String(m[1]).trim().toUpperCase() : "",
    customer: m ? String(m[2]).trim() : ""
  };
}

function outlookMimeFor(filename, mime) {
  const n = String(filename || "").toLowerCase();
  const m = String(mime || "").toLowerCase();
  if (/\.msg$/.test(n) || m.indexOf("ms-outlook") >= 0) return "application/vnd.ms-outlook";
  if (/\.eml$/.test(n) || m.indexOf("rfc822") >= 0) return "message/rfc822";
  return mime || "application/octet-stream";
}

function restIdFromWebUrl(url) {
  try {
    const u = new URL(String(url || "").trim());
    const item = u.searchParams.get("ItemID") || u.searchParams.get("itemid") || u.searchParams.get("itemId");
    if (item) return item;
    const m = String(u.pathname || "").match(/\/(?:id|deeplink\/read(?:item|m365)?)\/([^/]+)/i);
    return m ? decodeURIComponent(m[1]) : "";
  } catch (e) {
    return "";
  }
}

function sanitizeOutlookOpenUrl(url) {
  const raw = String(url || "").trim();
  if (!raw || /[\u0000-\u001f<>]/.test(raw)) return "";
  if (/^outlook:\/*/i.test(raw)) {
    const id = raw.replace(/^outlook:\/*/i, "").replace(/[^A-Za-z0-9+/=_-]/g, "");
    return id ? "outlook:" + id : "";
  }
  if (/^ms-outlook:/i.test(raw)) {
    if (/\s/.test(raw)) return "";
    return raw;
  }
  if (!/^https:\/\//i.test(raw)) return "";
  try {
    const u = new URL(raw);
    const host = String(u.hostname || "").toLowerCase();
    const ok = host === "outlook.office.com" || host === "outlook.office365.com" || host === "outlook.live.com"
      || host === "outlook.cloud.microsoft" || /\.outlook\.(office|office365|live)\.com$/.test(host);
    if (!ok) return "";
    u.hash = "";
    return u.toString();
  } catch (e) {
    return "";
  }
}

function parseOutlookLinks(text) {
  const raw = String(text || "");
  const found = [];
  const seen = new Set();
  const re = /(ms-outlook:[^\s"'<>]+|outlook:\/?\/?[A-Za-z0-9+/=_-]+|https:\/\/[^\s"'<>]+)/gi;
  let m;
  while ((m = re.exec(raw))) {
    const url = sanitizeOutlookOpenUrl(m[1].replace(/[),.;]+$/, ""));
    if (!url || seen.has(url)) continue;
    seen.add(url);
    found.push(url);
  }
  return found;
}

function outlookDesktopUrl(mail) {
  const item = mail && typeof mail === "object" ? mail : {};
  const entry = String(item.entry_id || "").replace(/^outlook:\/*/i, "").replace(/[^A-Za-z0-9+/=_-]/g, "");
  if (entry) return "outlook:" + entry;
  const rest = String(item.rest_id || "").trim() || restIdFromWebUrl(item.web_url || item.outlook_url || "") || String(item.item_id || "").trim();
  if (rest) return "ms-outlook://emails/message/open?restID=" + encodeURIComponent(rest);
  const open = sanitizeOutlookOpenUrl(item.outlook_url);
  if (open && /^(outlook:|ms-outlook:)/i.test(open)) return open;
  const mid = String(item.internet_message_id || "").trim();
  if (mid) return "ms-outlook://search?querytext=" + encodeURIComponent(mid);
  const title = String(item.title || item.subject || "").trim();
  if (title) return "ms-outlook://search?querytext=" + encodeURIComponent(title);
  return open;
}

function mailDedupeKeys(mail) {
  const item = mail && typeof mail === "object" ? mail : {};
  const keys = [];
  const rest = String(item.rest_id || "").trim().toLowerCase();
  if (rest) keys.push("rest:" + rest);
  const mid = String(item.internet_message_id || "").trim().toLowerCase();
  if (mid) keys.push("mid:" + mid);
  const entry = String(item.entry_id || "").replace(/^outlook:\/*/i, "").toLowerCase();
  if (entry) keys.push("entry:" + entry);
  const name = parseCorrespondenceName(item.title || item.filename || "").title
    .replace(/\.(msg|eml)$/i, "")
    .trim()
    .toLowerCase();
  if (name) keys.push("name:" + name);
  const url = String(item.outlook_url || "").trim().toLowerCase();
  if (url) keys.push("url:" + url);
  return keys;
}

function mailDedupeKey(mail) {
  return mailDedupeKeys(mail)[0] || "";
}

function normalizeOutlookMail(from, index) {
  const src = from && typeof from === "object" ? from : {};
  const title = String(src.title || src.subject || src.filename || "").trim();
  const parsed = parseCorrespondenceName(title);
  const web = sanitizeOutlookOpenUrl(src.web_url);
  const rest = String(src.rest_id || src.restId || "").trim() || restIdFromWebUrl(web) || restIdFromWebUrl(src.outlook_url);
  const entry = String(src.entry_id || src.entryId || "").replace(/^outlook:\/*/i, "").replace(/[^A-Za-z0-9+/=_-]/g, "");
  const mail = {
    id: String(src.id || "").trim() || ("mail_" + (Number(index) + 1 || 1)),
    title: parsed.title || title,
    from: String(src.from_name || (typeof src.from === "string" ? src.from : (src.from && src.from.displayName) || "")).trim(),
    from_email: String(src.from_email || src.fromEmail || (src.from && src.from.emailAddress) || "").trim(),
    sent_at: String(src.sent_at || src.sentAt || src.dateTimeCreated || "").trim(),
    order_no: String(src.order_no || parsed.order_no || "").trim(),
    customer: String(src.customer || parsed.customer || "").trim(),
    internet_message_id: String(src.internet_message_id || src.internetMessageId || "").trim(),
    item_id: String(src.item_id || src.itemId || "").trim(),
    rest_id: rest,
    entry_id: entry,
    web_url: web,
    outlook_url: "",
    kind: String(src.kind || "").trim(),
    stored_as: String(src.stored_as || "").trim(),
    filename: String(src.filename || "").trim(),
    mime: src.mime || ""
  };
  mail.outlook_url = outlookDesktopUrl({ ...mail, outlook_url: src.outlook_url }) || sanitizeOutlookOpenUrl(src.outlook_url);
  if (!mail.outlook_url && !mail.stored_as && !mail.kind) return null;
  return mail;
}

function extractOutlookFromBuffer(buf, filename) {
  const b = Buffer.isBuffer(buf) ? buf : Buffer.from(buf || []);
  if (!b.length) return null;
  const latin = b.toString("latin1");
  function asciiHeader(name) {
    const m = latin.match(new RegExp(name + ":\\s*([^\\r\\n\\x00]+)", "i"));
    return m ? String(m[1]).replace(/[\x00-\x08]/g, "").trim() : "";
  }
  function utf16Header(name) {
    const needle = Buffer.from(name + ":", "utf16le");
    const idx = b.indexOf(needle);
    if (idx < 0) return "";
    let out = "";
    for (let i = idx + needle.length; i + 1 < b.length; i += 2) {
      const c = b[i] | (b[i + 1] << 8);
      if (!c || c === 10 || c === 13) break;
      if (c >= 32) out += String.fromCharCode(c);
      if (out.length > 180) break;
    }
    return out.trim();
  }
  const mid = (latin.match(/Message-ID:\s*(<[^>\s]+>)/i) || [])[1]
    || (utf16Header("Message-ID").match(/<[^>\s]+>/) || [])[0]
    || "";
  const named = parseCorrespondenceName(filename);
  const subject = (asciiHeader("Subject") || utf16Header("Subject")).slice(0, 180) || named.title;
  const from = (asciiHeader("From") || utf16Header("From")).slice(0, 120);
  if (!mid && !asciiHeader("Subject") && !utf16Header("Subject") && !named.order_no) return null;
  return normalizeOutlookMail({
    title: subject || "Outlook email",
    from,
    internet_message_id: mid,
    filename
  }, 0);
}

function extractOutlookFromDataUrl(dataUrl, filename) {
  const s = String(dataUrl || "").trim();
  if (!s) return null;
  const m = s.match(/^data:[^;]*;base64,([\s\S]+)$/i);
  const raw = m ? m[1] : (s.indexOf("base64,") >= 0 ? s.split("base64,").pop() : "");
  if (!raw) return null;
  try {
    return extractOutlookFromBuffer(Buffer.from(raw, "base64"), filename);
  } catch (e) {
    return null;
  }
}

function mailsFromPastedLinks(text) {
  const raw = String(text || "");
  const urls = parseOutlookLinks(raw);
  if (urls.length) {
    return urls.map((url, i) => normalizeOutlookMail({
      title: "Outlook email",
      outlook_url: url,
      web_url: /^https:/i.test(url) ? url : "",
      rest_id: restIdFromWebUrl(url)
    }, i)).filter(Boolean);
  }
  const title = raw.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().slice(0, 180);
  if (title.length < 3) return [];
  const mail = normalizeOutlookMail({ title, subject: title }, 0);
  return mail ? [mail] : [];
}

function normalizeCorrespondence(from) {
  const c = from && from.correspondence;
  if (!c || typeof c !== "object") {
    return { saved_at: "", saved_by: "", mails: [] };
  }
  const mails = [];
  const seen = new Set();
  const list = Array.isArray(c.mails) ? c.mails : [];
  list.forEach((item, i) => {
    const mail = normalizeOutlookMail(item, i);
    if (!mail) return;
    const keys = mailDedupeKeys(mail);
    if (!keys.length || keys.some((k) => seen.has(k))) return;
    keys.forEach((k) => seen.add(k));
    mails.push(mail);
  });
  return {
    saved_at: c.saved_at || "",
    saved_by: String(c.saved_by || "").trim(),
    mails
  };
}

function copyPipeline(from, to) {
  to.tasks = normalizeEnquiryTasks(from);
  to.cost_sheet = cloneJson(from && from.cost_sheet, null);
  to.approval = cloneJson(from && from.approval, null);
  to.follow_ups = Array.isArray(from && from.follow_ups) ? cloneJson(from.follow_ups, []) : [];
  to.quotes = Array.isArray(from && from.quotes) ? cloneJson(from.quotes, []) : [];
  to.follow_up_assignee = String((from && from.follow_up_assignee) || "").trim();
  to.quote_assignee = String((from && from.quote_assignee) || "").trim();
  to.correspondence = normalizeCorrespondence(from);
  to.client_outcome = cloneJson(from && from.client_outcome, null);
  to.drawing = cloneJson(from && from.drawing, null);
  to.ready_for_orders = !!(from && from.ready_for_orders);
  to.custom_specs = normalizeCustomSpecs(from, null);
  to.design_description = String((from && from.design_description) || "").trim();
  to.created_at = String((from && from.created_at) || "").trim();
  to.events = normalizeEnquiryEvents(from);
}

function enquiryFilesDir(enquiryNo) {
  const dir = path.join(path.dirname(dbPath), "enquiry-files", enquiryQuoteKey(enquiryNo));
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function sanitizeUploadName(filename, fallback) {
  const clean = String(filename || "").replace(/[^\w.\- ()#]/g, "").slice(0, 120);
  return clean || fallback || "file";
}

function extFromUpload(filename, mime) {
  const m = String(filename || "").toLowerCase().match(/(\.[a-z0-9]{1,8})$/);
  if (m) return m[1];
  const type = String(mime || "").toLowerCase();
  if (type.indexOf("pdf") >= 0) return ".pdf";
  if (type.indexOf("png") >= 0) return ".png";
  if (type.indexOf("jpeg") >= 0 || type.indexOf("jpg") >= 0) return ".jpg";
  if (type.indexOf("webp") >= 0) return ".webp";
  if (type.indexOf("gif") >= 0) return ".gif";
  if (type.indexOf("outlook") >= 0 || type.indexOf("ms-outlook") >= 0) return ".msg";
  if (type.indexOf("rfc822") >= 0 || type.indexOf("message") >= 0) return ".eml";
  if (type.indexOf("csv") >= 0) return ".csv";
  if (type.indexOf("spreadsheet") >= 0 || type.indexOf("xlsx") >= 0) return ".xlsx";
  if (type.indexOf("excel") >= 0 || type.indexOf("xls") >= 0) return ".xls";
  return ".bin";
}

function decodeDataUrl(raw) {
  const s = String(raw || "").trim();
  if (!s) return null;
  const m = s.match(/^data:([^;]+);base64,([\s\S]+)$/i);
  if (m) {
    return { mime: m[1], buffer: Buffer.from(m[2], "base64") };
  }
  if (s.indexOf("base64,") >= 0) {
    const parts = s.split("base64,");
    return { mime: "application/octet-stream", buffer: Buffer.from(parts[1], "base64") };
  }
  return { mime: "application/octet-stream", buffer: Buffer.from(s, "base64") };
}

function saveEnquiryAttachment(enquiryNo, kind, dataUrl, filename) {
  const decoded = decodeDataUrl(dataUrl);
  if (!decoded || !decoded.buffer || !decoded.buffer.length) throw new Error("Upload a file first");
  const safeKind = String(kind || "file").replace(/[^\w.-]/g, "_");
  if (!safeKind) throw new Error("Missing file kind");
  const ext = extFromUpload(filename, decoded.mime);
  const dir = enquiryFilesDir(enquiryNo);
  const existing = fs.readdirSync(dir);
  for (const name of existing) {
    if (name === safeKind + path.extname(name) || name.indexOf(safeKind + ".") === 0) {
      try { fs.unlinkSync(path.join(dir, name)); } catch (e) {}
    }
  }
  const storedAs = safeKind + ext;
  fs.writeFileSync(path.join(dir, storedAs), decoded.buffer);
  const filenameSafe = sanitizeUploadName(filename, storedAs);
  const parsed = parseCorrespondenceName(filename || filenameSafe);
  return {
    kind: safeKind,
    filename: filenameSafe,
    mime: outlookMimeFor(filenameSafe, decoded.mime),
    stored_as: storedAs,
    uploaded_at: nowIso(),
    size: decoded.buffer.length,
    title: parsed.title,
    order_no: parsed.order_no,
    customer: parsed.customer
  };
}

function readEnquiryAttachment(enquiryNo, kind) {
  const want = String(kind || "").trim();
  if (want === "quote" || want === "quote.pdf") return readEnquiryQuotePdf(enquiryNo);
  const row = getEnquiryRaw(enquiryNo);
  if (!row) return null;
  let meta = null;
  if (want === "cost_sheet") meta = row.cost_sheet;
  else if (want === "pop") meta = row.client_outcome && row.client_outcome.file;
  else if (want === "drawing") meta = row.drawing && row.drawing.file;
  else if (want === "follow_up") {
    const list = Array.isArray(row.follow_ups) ? row.follow_ups : [];
    meta = list.length ? list[list.length - 1].file : null;
  } else if (/^quote_(\d+)$/.test(want)) {
    const n = Number(want.split("_").pop());
    const list = Array.isArray(row.quotes) ? row.quotes : [];
    const hit = list.find((q) => Number(q.n) === n);
    meta = hit && hit.file;
    if (!meta || !meta.stored_as) {
      const latest = list.length ? list[list.length - 1] : null;
      if (latest && Number(latest.n) === n) return readEnquiryQuotePdf(enquiryNo);
    }
  } else if (/^follow_up_(\d+)$/.test(want)) {
    const n = Number(want.split("_").pop());
    const list = Array.isArray(row.follow_ups) ? row.follow_ups : [];
    const hit = list.find((f) => Number(f.n) === n);
    meta = hit && hit.file;
  } else if (/^correspondence_(\d+)$/.test(want)) {
    const mails = (row.correspondence && Array.isArray(row.correspondence.mails)) ? row.correspondence.mails : [];
    const files = (row.correspondence && Array.isArray(row.correspondence.files)) ? row.correspondence.files : [];
    meta = mails.find((f) => f && f.kind === want)
      || files.find((f) => f && f.kind === want)
      || mails[Number(want.split("_").pop()) - 1]
      || files[Number(want.split("_").pop()) - 1]
      || null;
  }
  if (!meta || !meta.stored_as) return null;
  const file = path.join(enquiryFilesDir(enquiryNo), meta.stored_as);
  if (!fs.existsSync(file)) return null;
  return {
    buffer: fs.readFileSync(file),
    filename: meta.filename || meta.stored_as,
    mime: outlookMimeFor(meta.filename || meta.stored_as, meta.mime)
  };
}

function removeEnquiryFiles(enquiryNo) {
  try {
    const dir = path.join(path.dirname(dbPath), "enquiry-files", enquiryQuoteKey(enquiryNo));
    fs.rmSync(dir, { recursive: true, force: true });
  } catch (e) {}
}

function listEnquiryDeliverables(row) {
  const src = row && typeof row === "object" ? row : {};
  const items = [];
  const mails = normalizeCorrespondence(src).mails || [];
  mails.forEach((mail) => {
    items.push({
      group: "correspondence",
      label: "CORRESPONDANCE",
      title: mail.title || mail.filename || "Outlook email",
      filename: mail.filename || ((mail.title || "email") + ".msg"),
      kind: mail.kind || "",
      from: mail.from || mail.from_email || "",
      order_no: mail.order_no || "",
      open: !!(mail.kind && mail.stored_as),
      outlook: !!(mail.kind && mail.stored_as)
    });
  });
  if (src.cost_sheet && src.cost_sheet.stored_as) {
    items.push({
      group: "cost_sheet",
      label: "Cost sheet",
      title: src.cost_sheet.filename || "Cost sheet",
      filename: src.cost_sheet.filename || "cost-sheet",
      kind: "cost_sheet",
      from: "",
      order_no: "",
      open: true,
      outlook: false
    });
  }
  const quotes = Array.isArray(src.quotes) ? src.quotes : [];
  if (quotes.length) {
    quotes.forEach((item, i) => {
      const file = item && item.file;
      const n = item && item.n ? item.n : i + 1;
      const quoteNo = (item && item.quote_no) || "";
      const kind = (file && file.kind) || (n === quotes.length && enquiryHasQuotePdf(src.enquiry_no) ? "quote" : ("quote_" + n));
      items.push({
        group: "quote",
        label: quotes.length === 1
          ? ("Quote PDF" + (quoteNo ? " · " + quoteNo : ""))
          : ("Quote " + n + (quoteNo ? " · " + quoteNo : "")),
        title: (file && (file.filename || file.title)) || item.quote_pdf_name || src.quote_pdf_name || "quote.pdf",
        filename: (file && file.filename) || src.quote_pdf_name || "quote.pdf",
        kind,
        from: (item && item.by) || "",
        order_no: quoteNo,
        open: !!(file && file.stored_as) || (kind === "quote" && enquiryHasQuotePdf(src.enquiry_no)),
        outlook: false
      });
    });
  } else if (enquiryHasQuotePdf(src.enquiry_no)) {
    items.push({
      group: "quote",
      label: "Quote PDF" + (src.quote_no ? " · " + src.quote_no : ""),
      title: src.quote_pdf_name || "quote.pdf",
      filename: src.quote_pdf_name || "quote.pdf",
      kind: "quote",
      from: "",
      order_no: src.quote_no || "",
      open: true,
      outlook: false
    });
  }
  const followUps = Array.isArray(src.follow_ups) ? src.follow_ups : [];
  followUps.forEach((item, i) => {
    const file = item && item.file;
    if (!file || !file.stored_as) return;
    items.push({
      group: "follow_up",
      label: item.label || ("Follow up " + (item.n || i + 1)),
      title: file.filename || item.label || "Follow-up",
      filename: file.filename || "follow-up",
      kind: file.kind || ("follow_up_" + (item.n || i + 1)),
      from: item.by || "",
      order_no: "",
      open: true,
      outlook: false
    });
  });
  const pop = src.client_outcome && src.client_outcome.file;
  if (pop && pop.stored_as) {
    items.push({
      group: "pop",
      label: "Proof of payment",
      title: pop.filename || "POP",
      filename: pop.filename || "pop",
      kind: pop.kind || "pop",
      from: src.client_outcome.decided_by || "",
      order_no: "",
      open: true,
      outlook: false
    });
  }
  const drawing = src.drawing && src.drawing.file;
  if (drawing && drawing.stored_as) {
    items.push({
      group: "drawing",
      label: "Drawing",
      title: drawing.filename || "Drawing",
      filename: drawing.filename || "drawing",
      kind: drawing.kind || "drawing",
      from: src.drawing.assignee || "",
      order_no: "",
      open: true,
      outlook: false
    });
  }
  return items;
}

function decorateEnquiry(row) {
  const products = normalizeEnquiryLines(row, null);
  const named = products.filter((p) => p.product);
  const productsTotal = named.reduce((sum, p) => sum + parseMoney(p.value_excl_vat), 0);
  const delivery = parseMoney(row.delivery_excl_vat);
  const hasPdf = enquiryHasQuotePdf(row.enquiry_no);
  const tasks = normalizeEnquiryTasks(row);
  const openTasks = tasks.filter((t) => t.status === "open");
  const customSpecs = normalizeCustomSpecs(row, null);
  const enquiryType = String(row.enquiry_type || "").trim();
  const deliverables = listEnquiryDeliverables(row);
  const storedEvents = normalizeEnquiryEvents(row);
  const events = storedEvents.length ? storedEvents : synthesizeEnquiryEvents(row);
  const life = enquiryLifespan(row, events);
  return {
    ...row,
    products,
    tasks,
    product: named.map((p) => p.product).join(", "),
    category: named.map((p) => p.category).filter(Boolean)[0] || row.category || "",
    delivery_excl_vat: row.delivery_excl_vat === "" || row.delivery_excl_vat == null ? "" : money(delivery),
    products_total_excl_vat: money(productsTotal),
    quote_total_excl_vat: money(productsTotal + delivery),
    quotes: Array.isArray(row.quotes) ? row.quotes : [],
    quote_count: Array.isArray(row.quotes) ? row.quotes.length : (hasPdf ? 1 : 0),
    has_quote_pdf: hasPdf,
    quote_pdf_name: hasPdf ? (row.quote_pdf_name || "quote.pdf") : "",
    ready_for_orders: !!row.ready_for_orders,
    open_task_count: openTasks.length,
    assigned_to: openTasks.map((t) => t.assignee).filter(Boolean).join(", "),
    custom_specs: customSpecs,
    design_description: String(row.design_description || "").trim(),
    deliverables,
    deliverable_count: deliverables.length,
    spec_summary: enquiryType === "New Design"
      ? String(row.design_description || row.request || "").trim()
      : customSpecSummary(customSpecs),
    events,
    created_at: row.created_at || life.opened_at || "",
    opened_at: life.opened_at,
    opened_at_label: life.opened_at_label,
    ordered_at: life.ordered_at,
    ordered_at_label: life.ordered_at_label,
    lifespan_ms: life.lifespan_ms,
    lifespan_label: life.lifespan_label
  };
}

function listEnquiries() {
  return (state.enquiries || [])
    .slice()
    .sort((a, b) => enquiryNumberValue(b.enquiry_no) - enquiryNumberValue(a.enquiry_no))
    .map(decorateEnquiry);
}

function getEnquiryRaw(enquiryNo) {
  const n = enquiryNumberValue(enquiryNo);
  if (!n) return null;
  return (state.enquiries || []).find((r) => enquiryNumberValue(r.enquiry_no) === n) || null;
}

function getEnquiry(enquiryNo) {
  const row = getEnquiryRaw(enquiryNo);
  return row ? decorateEnquiry(row) : null;
}

function saveEnquiryRecord(row) {
  if (!row || !row.enquiry_no) throw new Error("Missing enquiry number");
  const existing = getEnquiryRaw(row.enquiry_no);
  if (!existing) throw new Error("Enquiry not found");
  existing.month_enquired = monthFromEnquiryDate(existing.date_enquired);
  Object.assign(existing, row);
  existing.updated_at = nowIso();
  save();
  return decorateEnquiry(existing);
}

function upsertEnquiry(row, opts) {
  const fromMigrate = !!(opts && opts.fromMigrate);
  const fromPipeline = !!(opts && opts.fromPipeline) || fromMigrate;
  const payload = {};
  for (const f of ENQUIRY_FIELDS) payload[f] = row[f] == null ? "" : String(row[f]).trim();
  if (!payload.enquiry_no) payload.enquiry_no = nextEnquiryNo();
  if (!/^#\d+$/.test(payload.enquiry_no)) {
    const n = enquiryNumberValue(payload.enquiry_no);
    payload.enquiry_no = n ? formatEnquiryNo(n) : nextEnquiryNo();
  }
  payload.month_enquired = monthFromEnquiryDate(payload.date_enquired);
  const existing = getEnquiryRaw(payload.enquiry_no);
  if (!payload.enquiry_type && existing && existing.enquiry_type) payload.enquiry_type = existing.enquiry_type;
  payload.products = normalizeEnquiryLines(row, existing);
  payload.product = payload.products.filter((p) => p.product).map((p) => p.product).join(", ");
  payload.category = payload.products.map((p) => p.category).filter(Boolean)[0] || payload.category;
  const deliveryRaw = row.delivery_excl_vat != null ? String(row.delivery_excl_vat).trim() : (existing && existing.delivery_excl_vat) || "";
  payload.delivery_excl_vat = deliveryRaw === "" ? "" : money(parseMoney(deliveryRaw));
  payload.quote_pdf_name = String((row.quote_pdf_name != null ? row.quote_pdf_name : (existing && existing.quote_pdf_name)) || "").trim();
  payload.quote_pdf_uploaded_at = (row.quote_pdf_uploaded_at != null ? row.quote_pdf_uploaded_at : (existing && existing.quote_pdf_uploaded_at)) || "";
  copyPipeline(existing || row, payload);
  if (!fromPipeline) {
    applyEnquiryTypeDetails(payload, row, existing);
  } else {
    payload.custom_specs = normalizeCustomSpecs(row, existing);
    payload.design_description = String(
      row.design_description != null ? row.design_description : (existing && existing.design_description) || ""
    ).trim();
  }

  if (!fromPipeline) {
    if (existing) {
      payload.status = existing.status || "New";
      payload.date_quoted = existing.date_quoted || payload.date_quoted;
      payload.quote_no = existing.quote_no || "";
    } else if (CAPTURE_STATUSES.indexOf(payload.status) === -1) {
      payload.status = "New";
    }
  }

  const keepValues = KEEP_VALUE_STATUSES.indexOf(payload.status) >= 0;
  if (!keepValues) {
    payload.products = payload.products.map((p) => ({ ...p, value_excl_vat: "" }));
    payload.delivery_excl_vat = "";
  }

  if (fromPipeline && !fromMigrate && payload.status === "Quoted") {
    const named = payload.products.filter((p) => p.product);
    if (!named.length) throw new Error("Add at least one product before marking Quoted");
    if (named.some((p) => p.value_excl_vat === "")) {
      throw new Error("Enter a value excluding VAT for each product when quoting");
    }
    if (payload.delivery_excl_vat === "") {
      throw new Error("Delivery excluding VAT is required when quoting");
    }
    if (!enquiryHasQuotePdf(payload.enquiry_no) && !row.quote_pdf_base64) {
      throw new Error("Upload and confirm the quote PDF before marking Quoted");
    }
    payload.quote_no = requireUniqueQuoteNo(payload.quote_no || row.quote_no, payload.enquiry_no);
  }

  if (fromPipeline && row.quote_pdf_base64) {
    if (!row.quote_pdf_confirmed) throw new Error("Preview the quote PDF and confirm it is the correct file before saving");
    const savedPdf = saveEnquiryQuotePdf(payload.enquiry_no, row.quote_pdf_base64, row.quote_pdf_name || "quote.pdf");
    payload.quote_pdf_name = savedPdf.quote_pdf_name;
    payload.quote_pdf_uploaded_at = savedPdf.quote_pdf_uploaded_at;
    payload.date_quoted = todayEnquiryDate();
  }

  if (payload.status === "Quoted" && !payload.date_quoted) payload.date_quoted = todayEnquiryDate();

  payload.updated_at = nowIso();
  if (!state.enquiries) state.enquiries = [];
  const actor = opts && opts.actor ? String(opts.actor).trim() : "";
  if (existing) {
    payload.id = existing.id;
    payload.created_at = existing.created_at || payload.created_at || nowIso();
    payload.events = normalizeEnquiryEvents(existing).length
      ? normalizeEnquiryEvents(existing)
      : synthesizeEnquiryEvents({ ...existing, created_at: payload.created_at });
    if (!fromPipeline && captureFieldsChanged(existing, payload)) {
      appendEnquiryEvent(payload, {
        kind: "edited",
        actor,
        from_status: existing.status || "",
        status: payload.status || existing.status || "",
        label: "Enquiry details edited"
      });
    }
    Object.assign(existing, payload);
    save();
    return decorateEnquiry(existing);
  }
  payload.id = (state.enquiries.reduce((m, o) => Math.max(m, Number(o.id) || 0), 0) || 0) + 1;
  payload.created_at = nowIso();
  payload.events = [];
  appendEnquiryEvent(payload, {
    kind: "created",
    actor,
    status: payload.status || "New",
    label: "Enquiry captured"
  });
  state.enquiries.push(payload);
  save();
  return decorateEnquiry(payload);
}

function enquiryHeaderField(header) {
  const h = String(header || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");
  const aliases = {
    enquiry_no: ["enquiry_no", "enquiry_number", "enquiry", "no"],
    date_enquired: ["date_enquired", "date", "date_enquire"],
    month_enquired: ["month_enquired", "month"],
    enquiry_source: ["enquiry_source"],
    enquiry_type: ["enquiry_type"],
    client_name: ["client_name", "name", "customer", "customer_name"],
    source: ["source"],
    client_email: ["client_email", "email", "email_address"],
    client_number: ["client_number", "number", "phone", "cell"],
    province: ["province"],
    category: ["category", "catergory"],
    product: ["product"],
    request: ["request"],
    status: ["status"],
    date_quoted: ["date_quoted"],
    quote_no: ["quote_no", "quote_number"],
    comment: ["comment", "comments"]
  };
  for (const field of Object.keys(aliases)) {
    if (aliases[field].indexOf(h) !== -1) return field;
  }
  if (ENQUIRY_FIELDS.indexOf(h) !== -1) return h;
  return "";
}

function copyEnquiriesFromWorkbook(book) {
  if (!book || typeof book.getSheetByName !== "function") return { imported: 0 };
  const sheet = book.getSheetByName("Enquiries")
    || book.getSheetByName("ENQUIRIES")
    || book.getSheetByName("Enquiry Log");
  if (!sheet || sheet.getLastRow() < 2) return { imported: 0 };
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = (sheet.getRange(1, 1, 1, lastCol).getValues()[0] || []).map(enquiryHeaderField);
  const grid = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  let imported = 0;
  grid.forEach((row) => {
    const payload = {};
    headers.forEach((field, i) => {
      if (!field) return;
      payload[field] = row[i];
    });
    if (!payload.enquiry_no && !payload.client_name && !payload.product) return;
    if (!payload.status) payload.status = "New";
    try {
      upsertEnquiry(payload, { fromMigrate: true });
      imported += 1;
    } catch (e) {
      console.error("[db] enquiry migrate failed", payload.enquiry_no, e && e.message ? e.message : e);
    }
  });
  return { imported };
}

function railwayBackup() {
  const workbook = getBook().toJSON();
  let office = {};
  try {
    office = JSON.parse(fs.readFileSync(dbPath, "utf8"));
  } catch (e) {
    office = state;
  }
  return {
    version: 1,
    database: "railway",
    exportedAt: nowIso(),
    workbook,
    office
  };
}

function deleteEnquiry(enquiryNo) {
  const want = String(enquiryNo || "").trim();
  try { removeEnquiryQuotePdf(want); } catch (e) {}
  try { removeEnquiryFiles(want); } catch (e) {}
  state.enquiries = (state.enquiries || []).filter((o) => o.enquiry_no !== want);
  save();
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
  persist: save,
  persistenceInfo,
  railwayBackup,
  copyEnquiriesFromWorkbook,
  ORDER_FIELDS,
  DROPDOWN_KEYS,
  VAT_RATE,
  parseMoney,
  money,
  formatRand,
  formatOrderId,
  formatPaymentDate,
  formatMonthOfSale,
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
  migrateJsonOrdersToWorkbook,
  normalizeOrdersSheet,
  ENQUIRY_FIELDS,
  KEEP_VALUE_STATUSES,
  CAPTURE_STATUSES,
  listEnquiries,
  getEnquiry,
  getEnquiryRaw,
  saveEnquiryRecord,
  upsertEnquiry,
  deleteEnquiry,
  nextEnquiryNo,
  nextQuoteNo,
  recentQuoteNos,
  normalizeQuoteNo,
  requireUniqueQuoteNo,
  quoteNoHint,
  monthFromEnquiryDate,
  listEnquiryDropdowns,
  addEnquiryDropdownItem,
  normalizeCustomSpecs,
  normalizeCorrespondence,
  parseCorrespondenceName,
  parseOutlookLinks,
  sanitizeOutlookOpenUrl,
  outlookDesktopUrl,
  normalizeOutlookMail,
  mailsFromPastedLinks,
  mailDedupeKey,
  mailDedupeKeys,
  extractOutlookFromBuffer,
  extractOutlookFromDataUrl,
  decodeDataUrl,
  outlookMimeFor,
  readEnquiryQuotePdf,
  saveEnquiryQuotePdf,
  enquiryHasQuotePdf,
  saveEnquiryAttachment,
  readEnquiryAttachment,
  listEnquiryDeliverables,
  normalizeEnquiryLines,
  todayEnquiryDate,
  nowIso,
  asDate,
  appendEnquiryEvent,
  formatSastDateTime
};
