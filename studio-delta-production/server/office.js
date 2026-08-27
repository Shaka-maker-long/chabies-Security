const {
  ORDER_FIELDS,
  DROPDOWN_KEYS,
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
  VAT_RATE,
  normalizeOrdersSheet
} = require("./db");
const { importGoogleWorkbook, tabCounts } = require("./workbook-store");

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

function mountOffice(app) {
  app.get("/api/office/orders", (_req, res) => {
    res.json({ ok: true, rows: listOrders().map(decorateMoney), fields: ORDER_FIELDS, vatRate: VAT_RATE });
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
      const book = await importGoogleWorkbook();
      normalizeOrdersSheet();
      const tabs = tabCounts(book);
      res.json({
        ok: true,
        imported: tabs.ORDERS || 0,
        tabs,
        message: "Copied Users, orders, production logs, steel and backboards from Google into Railway."
      });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/dropdowns", (_req, res) => {
    res.json({ ok: true, dropdowns: listDropdowns(), keys: DROPDOWN_KEYS });
  });

  app.post("/api/office/dropdowns/:field", (req, res) => {
    try {
      const dropdowns = addDropdownItem(req.params.field, (req.body && req.body.value) || "");
      res.json({ ok: true, dropdowns });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.delete("/api/office/dropdowns/:field", (req, res) => {
    try {
      const value = (req.body && req.body.value) || req.query.value || "";
      const dropdowns = removeDropdownItem(req.params.field, value);
      res.json({ ok: true, dropdowns });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/debtors", (_req, res) => {
    res.json({ ok: true, rows: listDebtors(), vatRate: VAT_RATE });
  });

  app.post("/api/office/orders/:orderNumber/payments", (req, res) => {
    try {
      const row = recordPayment(
        req.params.orderNumber,
        req.body && req.body.amount,
        req.body && req.body.note
      );
      res.json({ ok: true, row, debtors: listDebtors() });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
}

module.exports = { mountOffice };
