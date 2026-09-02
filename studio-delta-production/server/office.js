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
  normalizeOrdersSheet,
  listEnquiries,
  upsertEnquiry,
  deleteEnquiry,
  nextEnquiryNo,
  listEnquiryDropdowns,
  ENQUIRY_FIELDS,
  readEnquiryQuotePdf,
  readEnquiryAttachment,
  railwayBackup,
  copyEnquiriesFromWorkbook,
  persistenceInfo
} = require("./db");
const { importGoogleWorkbook, tabCounts, googleMigrateEnabled } = require("./workbook-store");
const staff = require("./staff");
const pipeline = require("./enquiry-pipeline");
const fs = require("fs");
const sqlite = require("./sqlite-store");

function requireOffice(req, res, next) {
  const profile = staff.readSession(req);
  if (!profile) {
    res.status(401).json({ ok: false, error: "Log in as Admin first." });
    return;
  }
  if (!profile.canSeeOffice) {
    res.status(403).json({ ok: false, error: "Production users can only use the floor." });
    return;
  }
  req.office = profile;
  next();
}

function requireDebtors(req, res, next) {
  if (!req.office || !req.office.canSeeDebtors) {
    res.status(403).json({ ok: false, error: "You cannot see Debtors." });
    return;
  }
  next();
}

async function migrateFromGoogle(_req, res) {
  if (!googleMigrateEnabled()) {
    res.status(400).json({
      ok: false,
      error: "Google Sheets is not the database. Set GOOGLE_MIGRATE=1, SHEET_ID, and Google credentials on Railway only to copy the old spreadsheet once, then remove GOOGLE_MIGRATE."
    });
    return;
  }
  try {
    const book = await importGoogleWorkbook();
    normalizeOrdersSheet();
    const enquiries = copyEnquiriesFromWorkbook(book);
    try { require("./gas").clearShopCache(); } catch (e) {}
    const tabs = tabCounts(book);
    res.json({
      ok: true,
      imported: tabs.ORDERS || 0,
      tabs,
      enquiries: enquiries.imported,
      message: "Copied the old Google spreadsheet into Railway. Remove GOOGLE_MIGRATE so Google cannot overwrite this again."
    });
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message || String(e) });
  }
}

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

function sendEnquiryFile(res, enquiryNo, kind, download) {
  const file = kind === "quote" || kind === "quote.pdf"
    ? readEnquiryQuotePdf(enquiryNo)
    : readEnquiryAttachment(enquiryNo, kind);
  if (!file) {
    res.status(404).json({ ok: false, error: "No file saved for this enquiry" });
    return;
  }
  const name = file.filename || "file";
  const mime = file.mime || (kind === "quote" || kind === "quote.pdf" ? "application/pdf" : "application/octet-stream");
  const outlook = /\.msg$/i.test(name) || mime === "application/vnd.ms-outlook" || /\.eml$/i.test(name);
  res.setHeader("Content-Type", outlook ? (/\.eml$/i.test(name) ? "message/rfc822" : "application/vnd.ms-outlook") : mime);
  res.setHeader(
    "Content-Disposition",
    (outlook || download ? "attachment" : "inline") + "; filename=\"" + String(name).replace(/"/g, "") + "\""
  );
  res.send(file.buffer);
}

function mountOffice(app) {
  app.post("/api/office/login", (req, res) => {
    const profile = staff.verifyUser((req.body && req.body.name) || "", (req.body && req.body.password) || "");
    if (!profile) {
      res.status(401).json({ ok: false, error: staff.loginFailureMessage() });
      return;
    }
    if (!profile.canSeeOffice) {
      res.status(403).json({ ok: false, error: "Production users can only use the floor." });
      return;
    }
    res.json({ ok: true, ...staff.createSession(profile) });
  });

  app.get("/api/office/me", (req, res) => {
    const profile = staff.readSession(req);
    if (!profile) {
      res.status(401).json({ ok: false, error: "Log in as Admin first." });
      return;
    }
    res.json({ ok: true, profile });
  });

  app.get("/api/office/users", requireOffice, (_req, res) => {
    res.json({ ok: true, rows: staff.listUsers(), tasks: staff.FLOOR_TASKS });
  });
  app.put("/api/office/users", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, row: staff.upsertUser(req.body || {}) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
  app.delete("/api/office/users/:name", requireOffice, (req, res) => {
    staff.deleteUser(req.params.name);
    res.json({ ok: true });
  });

  app.get("/api/office/durations", requireOffice, (_req, res) => {
    res.json({ ok: true, rows: staff.listDurations(), tasks: staff.FLOOR_TASKS });
  });
  app.put("/api/office/durations", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, rows: staff.setDurations((req.body && req.body.rows) || []) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/orders", requireOffice, (_req, res) => {
    res.json({ ok: true, rows: listOrders().map(decorateMoney), fields: ORDER_FIELDS, vatRate: VAT_RATE });
  });

  app.put("/api/office/orders", requireOffice, (req, res) => {
    try {
      const row = upsertOrder(req.body || {});
      res.json({ ok: true, row });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.delete("/api/office/orders/:orderNumber", requireOffice, (req, res) => {
    deleteOrder(req.params.orderNumber);
    res.json({ ok: true });
  });

  app.get("/api/office/enquiries", requireOffice, (_req, res) => {
    res.json({
      ok: true,
      rows: listEnquiries(),
      fields: ENQUIRY_FIELDS,
      nextEnquiryNo: nextEnquiryNo(),
      dropdowns: listEnquiryDropdowns(),
      vatRate: VAT_RATE
    });
  });

  app.put("/api/office/enquiries", requireOffice, (req, res) => {
    try {
      const row = upsertEnquiry(req.body || {});
      res.json({ ok: true, row, nextEnquiryNo: nextEnquiryNo(), dropdowns: listEnquiryDropdowns() });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.delete("/api/office/enquiries/:enquiryNo", requireOffice, (req, res) => {
    deleteEnquiry(req.params.enquiryNo);
    res.json({ ok: true, nextEnquiryNo: nextEnquiryNo() });
  });

  app.get("/api/office/enquiries/:enquiryNo/quote.pdf", requireOffice, (req, res) => {
    sendEnquiryFile(res, req.params.enquiryNo, "quote", String(req.query.download || "") === "1");
  });

  app.get("/api/office/enquiries/:enquiryNo/files/:kind", requireOffice, (req, res) => {
    sendEnquiryFile(res, req.params.enquiryNo, req.params.kind, String(req.query.download || "") === "1");
  });

  app.get("/api/office/assignees", requireOffice, (_req, res) => {
    res.json({ ok: true, rows: pipeline.officeAssignees() });
  });

  app.get("/api/office/my-tasks", requireOffice, (req, res) => {
    res.json({ ok: true, rows: pipeline.listMyTasks(req.office.name) });
  });

  app.get("/api/office/enquiries/:enquiryNo/process", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, me: req.office.name, ...pipeline.processSnapshot(req.params.enquiryNo) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/office/enquiries/:enquiryNo/process", requireOffice, (req, res) => {
    try {
      const snap = pipeline.applyAction(req.params.enquiryNo, req.office.name, req.body || {});
      res.json({ ok: true, me: req.office.name, ...snap });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/office/enquiries/:enquiryNo/outlook-mail", requireOffice, (req, res) => {
    try {
      const snap = pipeline.applyAction(req.params.enquiryNo, req.office.name, {
        action: "add_correspondence",
        correspondence_mails: [req.body || {}]
      });
      const mails = (snap.row && snap.row.correspondence && snap.row.correspondence.mails) || [];
      res.json({ ok: true, me: req.office.name, mail_count: mails.length, ...snap });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/schedule", requireOffice, (req, res) => {
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

  app.put("/api/office/schedule/row", requireOffice, (req, res) => {
    try {
      const row = upsertScheduleRow(req.body || {});
      res.json({ ok: true, row });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.put("/api/office/schedule/cell", requireOffice, (req, res) => {
    const { rowId, day, value } = req.body || {};
    if (!rowId || !day) {
      res.status(400).json({ ok: false, error: "rowId and day are required" });
      return;
    }
    setScheduleCell(Number(rowId), day, value);
    res.json({ ok: true });
  });

  app.post("/api/office/import-sheets", requireOffice, async (req, res) => {
    migrateFromGoogle(req, res);
  });

  app.post("/api/office/migrate-from-google", requireOffice, migrateFromGoogle);

  app.get("/api/office/database", requireOffice, (_req, res) => {
    res.json({
      ok: true,
      live: "railway",
      sheetsLive: false,
      migrateAvailable: googleMigrateEnabled(),
      persist: persistenceInfo()
    });
  });

  app.get("/api/office/backup", requireOffice, (_req, res) => {
    const stamp = new Date().toISOString().slice(0, 10);
    res.setHeader("Content-Disposition", "attachment; filename=\"studio-delta-railway-" + stamp + ".json\"");
    res.json(railwayBackup());
  });

  app.get("/api/office/backup.db", requireOffice, (_req, res) => {
    try { require("./db").persist(); } catch (e) {}
    sqlite.checkpoint();
    const file = sqlite.sqlitePath();
    if (!fs.existsSync(file)) {
      res.status(404).json({ ok: false, error: "SQLite is not on this volume yet. Use the app once, then download again." });
      return;
    }
    const stamp = new Date().toISOString().slice(0, 10);
    res.setHeader("Content-Type", "application/vnd.sqlite3");
    res.setHeader("Content-Disposition", "attachment; filename=\"studio-delta-" + stamp + ".db\"");
    fs.createReadStream(file).pipe(res);
  });

  app.get("/api/office/backups", requireOffice, (_req, res) => {
    const backup = require("./backup");
    res.json({
      ok: true,
      ...backup.info(),
      last: backup.loadStatus(),
      snapshots: backup.listLocalSnapshots()
    });
  });

  app.post("/api/office/backups/run", requireOffice, (_req, res) => {
    try {
      const status = require("./backup").runBackup("manual");
      res.json({ ok: !!status.ok, ...status });
    } catch (e) {
      res.status(500).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/backups/file/:name", requireOffice, (req, res) => {
    const file = require("./backup").safeBackupName(req.params.name);
    if (!file) {
      res.status(404).json({ ok: false, error: "That backup file is not on this volume." });
      return;
    }
    res.setHeader("Content-Type", file.endsWith(".json") ? "application/json" : "application/octet-stream");
    res.setHeader("Content-Disposition", "attachment; filename=\"" + require("path").basename(file) + "\"");
    fs.createReadStream(file).pipe(res);
  });

  app.get("/api/office/dropdowns", requireOffice, (_req, res) => {
    res.json({ ok: true, dropdowns: listDropdowns(), keys: DROPDOWN_KEYS });
  });

  app.post("/api/office/dropdowns/:field", requireOffice, (req, res) => {
    try {
      const dropdowns = addDropdownItem(req.params.field, (req.body && req.body.value) || "");
      res.json({ ok: true, dropdowns });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.delete("/api/office/dropdowns/:field", requireOffice, (req, res) => {
    try {
      const value = (req.body && req.body.value) || req.query.value || "";
      const dropdowns = removeDropdownItem(req.params.field, value);
      res.json({ ok: true, dropdowns });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/debtors", requireOffice, requireDebtors, (_req, res) => {
    res.json({ ok: true, rows: listDebtors(), vatRate: VAT_RATE });
  });

  app.post("/api/office/orders/:orderNumber/payments", requireOffice, requireDebtors, (req, res) => {
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
