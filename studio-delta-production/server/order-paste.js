"use strict";

const { ORDER_FIELDS, parseMoney, money, formatOrderId } = require("./db");
const { SHOP_STATUSES, isShopStatus } = require("./shop-status");
const { ORDER_TYPES } = require("./create-order-from-enquiry");

const PROVINCES = [
  "Eastern Cape", "Free State", "Gauteng", "KwaZulu-Natal", "Limpopo", "Mpumalanga",
  "North West", "Northern Cape", "Western Cape"
];

const MONTHS = {
  jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3, apr: 4, april: 4,
  may: 5, jun: 6, june: 6, jul: 7, july: 7, aug: 8, august: 8, sep: 9, sept: 9,
  september: 9, oct: 10, october: 10, nov: 11, november: 11, dec: 12, december: 12
};

const HEADER_MAP = {
  "quote": "quote_number",
  "quote number": "quote_number",
  "quote no": "quote_number",
  "order": "order_number",
  "order number": "order_number",
  "order no": "order_number",
  "status": "status",
  "assigned operator": "assigned_operator",
  "operator": "assigned_operator",
  "type": "type",
  "catergory": "category",
  "category": "category",
  "product": "product",
  "variation": "variation",
  "doors": "doors",
  "detailed description": "detailed_description",
  "description": "detailed_description",
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
  "price incl vat": "price_incl_vat",
  "price (excl vat)": "price_excl_vat",
  "price excl vat": "price_excl_vat",
  "amount paid": "amount_paid",
  "month of sale": "month_of_sale",
  "source": "source",
  "city": "city",
  "enquiry no": "enquiry_no",
  "enquiry number": "enquiry_no"
};

const LEAD_FIELDS = [
  "quote_number", "order_number", "status", "assigned_operator", "type", "category",
  "product", "variation", "doors", "detailed_description", "dimensions", "powder_coating",
  "client_name", "client_number", "email", "payment_date", "address"
];

function spacesToTabs(s) {
  if (s.indexOf("\t") !== -1) return s;
  let out = "";
  let inQuotes = false;
  let i = 0;
  while (i < s.length) {
    const c = s[i];
    if (inQuotes) {
      out += c;
      if (c === "\"" && s[i + 1] === "\"") {
        out += s[i + 1];
        i += 2;
        continue;
      }
      if (c === "\"") inQuotes = false;
      i += 1;
      continue;
    }
    if (c === "\"") {
      inQuotes = true;
      out += c;
      i += 1;
      continue;
    }
    if (c === " " && s[i + 1] === " ") {
      while (s[i] === " ") i += 1;
      out += "\t";
      continue;
    }
    out += c;
    i += 1;
  }
  return out;
}

function parseTsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let i = 0;
  let inQuotes = false;
  const s = spacesToTabs(String(text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n"));
  while (i < s.length) {
    const c = s[i];
    if (inQuotes) {
      if (c === "\"") {
        if (s[i + 1] === "\"") {
          cell += "\"";
          i += 2;
          continue;
        }
        inQuotes = false;
        i += 1;
        continue;
      }
      cell += c;
      i += 1;
      continue;
    }
    if (c === "\"") {
      inQuotes = true;
      i += 1;
      continue;
    }
    if (c === "\t") {
      row.push(cell);
      cell = "";
      i += 1;
      continue;
    }
    if (c === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      i += 1;
      continue;
    }
    cell += c;
    i += 1;
  }
  if (cell !== "" || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows.filter((r) => r.some((c) => String(c || "").trim()));
}

function normHeader(s) {
  return String(s || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function looksLikeHeader(cells) {
  const norms = cells.map(normHeader);
  const joined = norms.join(" | ");
  if (/order number/.test(joined) || /quote number/.test(joined) || /catergory/.test(joined)) return true;
  const quoteHead = norms[0] === "quote" || norms[0] === "quote no";
  const orderHead = norms[1] === "order" || norms[1] === "order no";
  return quoteHead && orderHead;
}

function headerFields(cells) {
  return cells.map((c) => HEADER_MAP[normHeader(c)] || "");
}

function isMoneyCell(s) {
  const t = String(s || "").trim();
  if (!t) return false;
  if (!/^R?\s*[\d][\d\s,]*(?:\.\d+)?$/.test(t)) return false;
  return parseMoney(t) > 0 || /^R?\s*0/.test(t);
}

function matchProvince(s) {
  const t = String(s || "").trim().toLowerCase();
  if (!t) return "";
  return PROVINCES.find((p) => p.toLowerCase() === t) || "";
}

function matchType(s) {
  const t = String(s || "").trim();
  return ORDER_TYPES.find((x) => x.toLowerCase() === t.toLowerCase()) || "";
}

function parseMonthOfSale(s) {
  const t = String(s || "").trim().replace(/\s+/g, " ");
  const m = t.match(/^([A-Za-z]+)\s+(\d{4})$/);
  if (!m) return t;
  const month = MONTHS[m[1].toLowerCase()];
  if (!month) return t;
  const names = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  return names[month - 1] + " " + m[2];
}

function parseSheetDate(s, monthOfSale) {
  const t = String(s || "").trim();
  if (!t) return "";
  const iso = t.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[1] + "-" + iso[2] + "-" + iso[3];
  const dmy = t.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (dmy) {
    const y = dmy[3].length === 2 ? "20" + dmy[3] : dmy[3];
    return y + "-" + String(dmy[2]).padStart(2, "0") + "-" + String(dmy[1]).padStart(2, "0");
  }
  const named = t.match(/^(\d{1,2})[ \-]([A-Za-z]+)$/);
  if (named) {
    const month = MONTHS[named[2].toLowerCase()];
    const yearHit = String(monthOfSale || "").match(/(\d{4})/);
    const y = yearHit ? yearHit[1] : String(new Date().getFullYear());
    if (month) {
      return y + "-" + String(month).padStart(2, "0") + "-" + String(named[1]).padStart(2, "0");
    }
  }
  return t;
}

function emptyRow() {
  const row = {};
  ORDER_FIELDS.forEach((f) => { row[f] = ""; });
  return row;
}

function applyLead(row, cells) {
  LEAD_FIELDS.forEach((field, i) => {
    if (cells[i] != null) row[field] = String(cells[i] || "").trim();
  });
  return cells.slice(LEAD_FIELDS.length);
}

function applyTail(row, tail) {
  const leftover = [];
  tail.forEach((raw) => {
    const cell = String(raw || "").trim();
    if (!cell) return;
    if (!row.province && matchProvince(cell)) {
      row.province = matchProvince(cell);
      return;
    }
    if (isMoneyCell(cell)) {
      leftover.push(cell);
      return;
    }
    if (!row.month_of_sale && /[A-Za-z]+\s+\d{4}/.test(cell)) {
      row.month_of_sale = parseMonthOfSale(cell);
      return;
    }
    leftover.push(cell);
  });
  const moneys = leftover.filter(isMoneyCell).map((c) => parseMoney(c));
  const others = leftover.filter((c) => !isMoneyCell(c));
  if (moneys.length >= 2) {
    const a = moneys[0];
    const b = moneys[1];
    row.price_incl_vat = money(Math.max(a, b));
    row.price_excl_vat = money(Math.min(a, b));
  } else if (moneys.length === 1) {
    row.price_incl_vat = money(moneys[0]);
  }
  others.forEach((cell) => {
    if (!row.source) {
      row.source = cell;
      return;
    }
    if (!row.city) {
      row.city = cell;
      return;
    }
    if (!row.enquiry_no && /^#?\d+/.test(cell)) row.enquiry_no = cell;
  });
}

function rowFromMapped(cells, fields) {
  const row = emptyRow();
  const moneys = [];
  fields.forEach((field, i) => {
    const cell = String(cells[i] == null ? "" : cells[i]).trim();
    if (!field) return;
    if (field === "price_incl_vat" || field === "price_excl_vat") {
      if (isMoneyCell(cell)) moneys.push({ field, value: parseMoney(cell) });
      return;
    }
    row[field] = cell;
  });
  if (moneys.length >= 2 && !row.price_incl_vat) {
    const vals = moneys.map((m) => m.value);
    row.price_incl_vat = money(Math.max(vals[0], vals[1]));
    row.price_excl_vat = money(Math.min(vals[0], vals[1]));
  } else {
    moneys.forEach((m) => { row[m.field] = money(m.value); });
  }
  return row;
}

function normalizePastedRow(raw, paidInFull) {
  const row = Object.assign(emptyRow(), raw || {});
  row.order_number = formatOrderId(row.order_number);
  row.quote_number = formatOrderId(row.quote_number);
  row.product = String(row.product || "").trim();
  row.client_name = String(row.client_name || "").trim();
  const typed = matchType(row.type);
  row.type = typed || (row.product ? "Standard" : "");
  const status = String(row.status || "").trim();
  row.status = isShopStatus(status) ? status : "Not Yet Started";
  row.month_of_sale = parseMonthOfSale(row.month_of_sale);
  row.payment_date = parseSheetDate(row.payment_date, row.month_of_sale);
  const incl = parseMoney(row.price_incl_vat);
  const excl = parseMoney(row.price_excl_vat);
  if (incl > 0 && excl > 0 && incl < excl) {
    row.price_incl_vat = money(excl);
    row.price_excl_vat = money(incl);
  }
  if (paidInFull !== false) {
    const due = parseMoney(row.price_incl_vat) || parseMoney(row.price_excl_vat) * 1.15;
    if (due > 0) row.amount_paid = money(parseMoney(row.price_incl_vat) || due);
  }
  const errors = [];
  if (!row.order_number) errors.push("Order number is missing");
  if (!row.product) errors.push((row.order_number || "Row") + " needs a product");
  if (!row.client_name) errors.push((row.order_number || "Row") + " needs a client name");
  return { row, errors };
}

function parseOrderPaste(text, opts) {
  const paidInFull = !opts || opts.paidInFull !== false;
  const grid = parseTsv(text);
  if (!grid.length) return { rows: [], errors: ["Paste the order rows from the old sheet."] };
  let start = 0;
  let fields = null;
  if (looksLikeHeader(grid[0])) {
    fields = headerFields(grid[0]);
    start = 1;
  }
  const rows = [];
  const errors = [];
  for (let i = start; i < grid.length; i++) {
    const raw = fields ? rowFromMapped(grid[i], fields) : (() => {
      const row = emptyRow();
      const tail = applyLead(row, grid[i]);
      applyTail(row, tail);
      return row;
    })();
    const parsed = normalizePastedRow(raw, paidInFull);
    parsed.errors.forEach((e) => errors.push(e));
    if (parsed.errors.length) continue;
    rows.push(parsed.row);
  }
  if (!rows.length && !errors.length) errors.push("No order rows were found in that paste.");
  return { rows, errors };
}

module.exports = {
  parseTsv,
  parseOrderPaste,
  normalizePastedRow,
  SHOP_STATUSES,
  ORDER_TYPES
};
