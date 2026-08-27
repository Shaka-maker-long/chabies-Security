const { google } = require("googleapis");
const {
  ORDER_FIELDS,
  listOrders,
  upsertOrder,
  deleteOrder,
  listSchedule,
  upsertScheduleRow,
  setScheduleCell,
  countOrders
} = require("./db");

function workdays(fromIso, days) {
  const out = [];
  const d = new Date(fromIso + "T12:00:00");
  while (out.length < days) {
    const dow = d.getDay();
    if (dow !== 0 && dow !== 6) {
      out.push(d.toISOString().slice(0, 10));
    }
    d.setDate(d.getDate() + 1);
  }
  return out;
}

function mondayOf(dateIso) {
  const d = new Date((dateIso || new Date().toISOString().slice(0, 10)) + "T12:00:00");
  const day = d.getDay() || 7;
  d.setDate(d.getDate() - day + 1);
  return d.toISOString().slice(0, 10);
}

function normHeader(s) {
  return String(s || "").trim().toLowerCase().replace(/\s+/g, " ");
}

const HEADER_MAP = {
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
  "month of sale": "month_of_sale",
  "source": "source",
  "city": "city"
};

async function importOrdersFromSheets() {
  let credentials;
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  } else if (process.env.GOOGLE_APPLICATION_CREDENTIALS) {
    credentials = require(process.env.GOOGLE_APPLICATION_CREDENTIALS);
  } else {
    throw new Error("Google credentials are not set");
  }
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"]
  });
  const sheets = google.sheets({ version: "v4", auth });
  const spreadsheetId = process.env.SHEET_ID;
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: "ORDERS"
  });
  const values = res.data.values || [];
  if (values.length < 2) return { imported: 0 };
  const headers = values[0].map(normHeader);
  let n = 0;
  for (let i = 1; i < values.length; i++) {
    const row = {};
    headers.forEach((h, c) => {
      const field = HEADER_MAP[h];
      if (field) row[field] = values[i][c] || "";
    });
    if (row.order_number) {
      upsertOrder(row);
      n++;
    }
  }
  return { imported: n };
}

function mountOffice(app) {
  app.get("/api/office/orders", (_req, res) => {
    res.json({ ok: true, rows: listOrders(), fields: ORDER_FIELDS });
  });

  app.put("/api/office/orders", (req, res) => {
    try {
      const row = upsertOrder(req.body || {});
      res.json({ ok: true, row });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.delete("/api/office/orders/:orderNumber", (req, res) => {
    deleteOrder(req.params.orderNumber);
    res.json({ ok: true });
  });

  app.get("/api/office/schedule", (req, res) => {
    const start = mondayOf(req.query.start);
    const days = workdays(start, 15);
    const fromDay = days[0];
    const toDay = days[days.length - 1];
    res.json({
      ok: true,
      start,
      days,
      rows: listSchedule(fromDay, toDay)
    });
  });

  app.put("/api/office/schedule/row", (req, res) => {
    try {
      const row = upsertScheduleRow(req.body || {});
      res.json({ ok: true, row });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.put("/api/office/schedule/cell", (req, res) => {
    const { rowId, day, value } = req.body || {};
    if (!rowId || !day) {
      res.status(400).json({ ok: false, error: "rowId and day are required" });
      return;
    }
    setScheduleCell(Number(rowId), day, value);
    res.json({ ok: true });
  });

  app.post("/api/office/import-sheets", async (_req, res) => {
    try {
      const result = await importOrdersFromSheets();
      res.json({ ok: true, ...result });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
}

module.exports = { mountOffice };
