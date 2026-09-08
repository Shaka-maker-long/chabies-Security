const {
  ORDER_FIELDS,
  DROPDOWN_KEYS,
  listOrders,
  upsertOrder,
  deleteOrder,
  listSchedule,
  listDeliveryItems,
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
  getEnquiry,
  upsertEnquiry,
  deleteEnquiry,
  deleteAllEnquiries,
  deleteAllOrders,
  nextEnquiryNo,
  listEnquiryDropdowns,
  ENQUIRY_FIELDS,
  readEnquiryQuotePdf,
  readEnquiryAttachment,
  railwayBackup,
  copyEnquiriesFromWorkbook,
  persistenceInfo,
  findOpenEnquiryDuplicates,
  createOrderFromEnquiry,
  createOrderDraftFromEnquiry,
  createOrdersFromEnquiryForm,
  nextStudioOrderNumber,
  formatOrderId
} = require("./db");
const fromEnquiry = require("./create-order-from-enquiry");
const jobCard = require("./job-card");
const { importGoogleWorkbook, tabCounts, googleMigrateEnabled, dataDir } = require("./workbook-store");
const staff = require("./staff");
const pipeline = require("./enquiry-pipeline");
const desk = require("./enquiry-desk");
const paintShop = require("./powder-shop");
const glassPo = require("./glass-po");
const floorPlanning = require("./floor-planning");
const fs = require("fs");
const sqlite = require("./sqlite-store");
const {
  SCHEDULE_CODES,
  SCHEDULE_WORKDAYS,
  mondayOf,
  workdays,
  isoWeekInfo,
  weekKey,
  weekOptions
} = require("./office-schedule");

function officeCookie(token, clear) {
  if (clear) return "sd_office=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0";
  return "sd_office=" + encodeURIComponent(token) + "; Path=/; HttpOnly; SameSite=Lax";
}

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
      error: "Google copy is not available."
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
      message: "Copied."
    });
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message || String(e) });
  }
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
    (download ? "attachment" : "inline") + "; filename=\"" + String(name).replace(/"/g, "") + "\""
  );
  res.send(file.buffer);
}

function floorLayoutPath() {
  return require("path").join(dataDir(), "floor-layout.json");
}
function readFloorLayout() {
  try {
    const raw = fs.readFileSync(floorLayoutPath(), "utf8");
    const j = JSON.parse(raw);
    return j && typeof j === "object" ? j : {};
  } catch (e) {
    return {};
  }
}
function writeFloorLayout(layout) {
  const body = layout && typeof layout === "object" ? layout : {};
  fs.writeFileSync(floorLayoutPath(), JSON.stringify(body));
  return body;
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
    const session = staff.createSession(profile);
    res.setHeader("Set-Cookie", officeCookie(session.token));
    res.json({ ok: true, canManageUsers: staff.canManageUsers(profile), ...session });
  });

  app.post("/api/office/logout", (req, res) => {
    staff.dropSession(req);
    res.setHeader("Set-Cookie", officeCookie("", true));
    res.json({ ok: true });
  });

  app.get("/api/office/me", (req, res) => {
    const profile = staff.readSession(req);
    if (!profile) {
      res.status(401).json({ ok: false, error: "Log in as Admin first." });
      return;
    }
    res.json({ ok: true, profile: Object.assign({}, profile, { canManageUsers: staff.canManageUsers(profile) }) });
  });

  app.get("/api/office/users", requireOffice, (req, res) => {
    const manage = staff.canManageUsers(req.office);
    const rows = staff.listUsers();
    res.json({
      ok: true,
      rows: manage ? rows : rows.filter((u) => String(u.name).toLowerCase() === String(req.office.name).toLowerCase()),
      tasks: staff.FLOOR_TASKS,
      canManageUsers: manage
    });
  });
  app.put("/api/office/users", requireOffice, (req, res) => {
    try {
      if (!staff.canManageUsers(req.office)) {
        res.status(403).json({ ok: false, error: "Only the Manager can add people or assign roles." });
        return;
      }
      res.json({ ok: true, row: staff.upsertUser(req.body || {}) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
  app.delete("/api/office/users/:name", requireOffice, (req, res) => {
    if (!staff.canManageUsers(req.office)) {
      res.status(403).json({ ok: false, error: "Only the Manager can delete people." });
      return;
    }
    try {
      staff.deleteUser(req.params.name);
      res.json({ ok: true });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
  app.post("/api/office/password", (req, res) => {
    const session = staff.readSession(req);
    const body = req.body || {};
    const requested = String(body.name || "").trim();
    const self = String((session && session.name) || "").trim();
    const next = body.new_password || body.newPassword;
    try {
      if (self && requested && self.toLowerCase() !== requested.toLowerCase()) {
        if (!staff.canManageUsers(session)) {
          res.status(403).json({ ok: false, error: "Only the Manager can change someone else's access code." });
          return;
        }
        res.json({ ok: true, name: staff.setUserPassword(requested, next).name });
        return;
      }
      const name = self || requested;
      if (!name) throw new Error("Name is required");
      staff.changeOwnPassword(name, body.current_password || body.currentPassword || body.password, next);
      res.json({ ok: true, name });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/materials-to-order", requireOffice, async (_req, res) => {
    try {
      const { callShopFunction } = require("./gas");
      const data = await callShopFunction("listMaterialsToOrder", []);
      const glass = glassPo.snapshot();
      res.json({
        ok: true,
        glass: glass.glass,
        wood: (data && data.wood) || [],
        toOrder: glass.toOrder,
        outstanding: glass.outstanding,
        received: glass.received,
        toOrderCount: glass.toOrderCount,
        outstandingCount: glass.outstandingCount
      });
    } catch (e) {
      res.status(500).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/office/glass-po", requireOffice, (req, res) => {
    try {
      const result = glassPo.createPurchaseOrder(
        (req.body && (req.body.lineIds || req.body.ids)) || [],
        req.office && req.office.name
      );
      res.json({ ok: true, ...result });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/office/glass-po/receive", requireOffice, (req, res) => {
    try {
      const result = glassPo.receiveGlass(req.body || {}, req.office && req.office.name);
      res.json({ ok: true, ...result });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/glass-po/invoices/:invoiceId", requireOffice, (req, res) => {
    const file = glassPo.readInvoiceFile(req.params.invoiceId);
    if (!file) {
      res.status(404).json({ ok: false, error: "Invoice not found." });
      return;
    }
    const download = String((req.query && req.query.download) || "") === "1";
    res.setHeader("Content-Type", file.mime || "application/octet-stream");
    res.setHeader(
      "Content-Disposition",
      (download ? "attachment" : "inline") + "; filename=\"" + String(file.filename || "invoice").replace(/"/g, "") + "\""
    );
    res.send(file.buffer);
  });

  app.put("/api/office/materials-to-order", requireOffice, async (req, res) => {
    try {
      const { callShopFunction } = require("./gas");
      const result = await callShopFunction("markMaterialOrdered", [
        req.body && req.body.kind,
        req.body && req.body.id,
        req.body && req.body.status
      ]);
      if (!result || result.success === false) {
        return res.status(400).json({ ok: false, error: (result && result.message) || "Could not update" });
      }
      res.json({ ok: true, result });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
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

  app.get("/api/office/planning", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, board: floorPlanning.getBoard(req.query && req.query.week) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
  app.post("/api/office/planning/schedule", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, result: floorPlanning.scheduleSelected(req.body || {}) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
  app.delete("/api/office/planning/orders/:orderNumber", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, result: floorPlanning.unscheduleOrder(req.params.orderNumber) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
  app.post("/api/office/planning/move", requireOffice, (req, res) => {
    try {
      const body = req.body || {};
      res.json({ ok: true, result: floorPlanning.moveBlock(body.blockId || body.id, body.start) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
  app.post("/api/office/planning/other", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, result: floorPlanning.insertOtherTask(req.body || {}) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
  app.delete("/api/office/planning/blocks/:id", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, result: floorPlanning.removeBlock(req.params.id) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/orders", requireOffice, (_req, res) => {
    const rows = listOrders().map((o) => {
      const copy = decorateMoney(o);
      delete copy.payments;
      return copy;
    });
    res.json({
      ok: true,
      rows,
      fields: ORDER_FIELDS,
      vatRate: VAT_RATE,
      nextOrderNumber: nextStudioOrderNumber(),
      operators: staff.listUsers().map((u) => u.name).filter(Boolean),
      canManageUsers: staff.canManageUsers(_req.office)
    });
  });

  app.put("/api/office/orders", requireOffice, (req, res) => {
    try {
      const body = req.body || {};
      const existing = listOrders().find((o) => o.order_number === formatOrderId(body.order_number)) || null;
      const row = upsertOrder(jobCard.applyOfficeOrderStatusLock(body, existing));
      res.json({ ok: true, row });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/job-cards", requireOffice, (_req, res) => {
    res.json({ ok: true, rows: jobCard.listGeneratedJobCards() });
  });

  app.get("/api/office/job-cards/eligible", requireOffice, (_req, res) => {
    res.json({ ok: true, rows: jobCard.listEligibleOrders() });
  });

  app.get("/api/office/job-cards/catalog", requireOffice, (req, res) => {
    const product = (req.query && req.query.product) || "";
    const found = require("./product-catalog").lookupProduct(product);
    res.json({
      ok: true,
      product: found,
      catalog: product ? undefined : require("./product-catalog").listCatalog()
    });
  });

  app.get("/api/office/job-cards/:orderNumber", requireOffice, (req, res) => {
    const rec = jobCard.getJobCard(req.params.orderNumber);
    if (!rec) {
      res.status(404).json({ ok: false, error: "No job card saved for this order yet." });
      return;
    }
    res.json({ ok: true, record: rec });
  });

  app.get("/api/office/job-cards/:orderNumber/pdf", requireOffice, (req, res) => {
    const file = jobCard.readJobCardPdf(req.params.orderNumber);
    if (!file) {
      res.status(404).json({ ok: false, error: "No job card PDF for this order." });
      return;
    }
    const download = String((req.query && req.query.download) || "") === "1";
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", (download ? "attachment" : "inline") + "; filename=\"" + file.filename + "\"");
    res.send(file.buffer);
  });

  app.post("/api/office/job-cards/parse", requireOffice, (req, res) => {
    const cutting = jobCard.parsePastedCuttingList((req.body && (req.body.text || req.body.cutting_text)) || "");
    res.json({ ok: true, cutting, count: jobCard.cuttingCount(cutting) });
  });

  app.post("/api/office/job-cards", requireOffice, async (req, res) => {
    try {
      const result = await jobCard.generateJobCard(req.body || {});
      res.json({ ok: true, ...result });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.delete("/api/office/orders/:orderNumber", requireOffice, (req, res) => {
    deleteOrder(req.params.orderNumber);
    res.json({ ok: true });
  });

  app.post("/api/office/orders/clear-all", requireOffice, (req, res) => {
    if (!staff.canManageUsers(req.office)) {
      res.status(403).json({ ok: false, error: "Only the Manager can clear all orders." });
      return;
    }
    const confirm = String((req.body && (req.body.confirm || req.body.confirmation)) || "").trim();
    if (confirm.toUpperCase() !== "CLEAR") {
      res.status(400).json({ ok: false, error: "Type CLEAR to delete every order. Enquiries stay." });
      return;
    }
    try {
      const removed = deleteAllOrders();
      try { jobCard.deleteAllJobCards(); } catch (e) {}
      res.json({ ok: true, removed, nextOrderNumber: nextStudioOrderNumber() });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/enquiries", requireOffice, (req, res) => {
    res.json({
      ok: true,
      rows: listEnquiries(),
      fields: ENQUIRY_FIELDS,
      nextEnquiryNo: nextEnquiryNo(),
      dropdowns: listEnquiryDropdowns(),
      vatRate: VAT_RATE,
      canManageUsers: staff.canManageUsers(req.office),
      assignees: pipeline.officeAssignees(),
      enquiryRoles: staff.enquiryRoleDefaults(),
      onboardStatuses: pipeline.ONBOARD_STATUSES
    });
  });

  app.get("/api/office/enquiries/duplicates", requireOffice, (req, res) => {
    try {
      res.json({
        ok: true,
        rows: findOpenEnquiryDuplicates({
          client_email: req.query.email || req.query.client_email,
          client_number: req.query.number || req.query.client_number,
          enquiry_no: req.query.except || req.query.enquiry_no
        }, req.query.except || req.query.enquiry_no)
      });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/enquiries/dashboard", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, ...require("./enquiry-dashboard").buildDashboard(req.query || {}) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/enquiries/dashboard/drill", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, ...require("./enquiry-dashboard").buildDrill(req.query || {}) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/enquiries/replies", requireOffice, (_req, res) => {
    try {
      res.json({
        ok: true,
        replies: desk.loadReplies(),
        topics: desk.TOPICS,
        placeholders: desk.PLACEHOLDERS,
        enquiries: desk.enquiryOptions()
      });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.put("/api/office/enquiries/replies", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, ...desk.upsertReply(req.body || {}) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/office/enquiries/replies/restore", requireOffice, (_req, res) => {
    try {
      res.json({ ok: true, replies: desk.restoreReplies() });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.delete("/api/office/enquiries/replies/:id", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, replies: desk.deleteReply(req.params.id) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/enquiries/bookings", requireOffice, (_req, res) => {
    try {
      res.json({
        ok: true,
        bookings: desk.loadBookings(),
        statuses: desk.BOOKING_STATUSES
      });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.put("/api/office/enquiries/bookings", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, ...desk.upsertBooking(req.body || {}, req.office && req.office.name) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/office/enquiries/bookings/:id/enquiry", requireOffice, (req, res) => {
    try {
      const actor = req.office && req.office.name;
      const out = desk.transferVisit(req.params.id, actor);
      let row = out.enquiry;
      if (out.created && actor && pipeline.isAutoCaptureStatus(row && row.status)) {
        row = pipeline.applyCaptureRoute(row.enquiry_no, actor).row;
      }
      res.json({
        ok: true,
        booking: out.booking,
        bookings: out.bookings,
        enquiry: row,
        created: out.created,
        linkedExisting: out.linkedExisting
      });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.delete("/api/office/enquiries/bookings/:id", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, bookings: desk.deleteBooking(req.params.id) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/office/enquiries/clear-all", requireOffice, (req, res) => {
    if (!staff.canManageUsers(req.office)) {
      res.status(403).json({ ok: false, error: "Only the Manager can clear all enquiries." });
      return;
    }
    const confirm = String((req.body && (req.body.confirm || req.body.confirmation)) || "").trim();
    if (confirm.toUpperCase() !== "CLEAR") {
      res.status(400).json({ ok: false, error: "Type CLEAR to delete every enquiry. Orders stay." });
      return;
    }
    try {
      const removed = deleteAllEnquiries();
      res.json({ ok: true, removed, nextEnquiryNo: nextEnquiryNo() });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/office/enquiries/onboard", requireOffice, (req, res) => {
    try {
      const snap = pipeline.onboardEnquiry(req.office.name, req.body || {});
      res.json({
        ok: true,
        row: snap.row,
        nextEnquiryNo: nextEnquiryNo(),
        dropdowns: listEnquiryDropdowns(),
        duplicates: findOpenEnquiryDuplicates(snap.row, snap.row.enquiry_no),
        ...snap
      });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.put("/api/office/enquiries", requireOffice, (req, res) => {
    try {
      const actor = req.office && req.office.name;
      const incoming = req.body || {};
      const saved = upsertEnquiry(incoming, {
        actor,
        createOnly: !!(incoming.create_only || incoming.createOnly),
        previousEnquiryNo: incoming.previous_enquiry_no || incoming.previousEnquiryNo || ""
      });
      const pasted = String(incoming.correspondence_links || incoming.correspondenceLinks || "").trim();
      if (actor && pasted) {
        pipeline.applyAction(saved.enquiry_no, actor, {
          action: "add_correspondence",
          correspondence_links: pasted
        });
      }
      const row = actor && pipeline.isAutoCaptureStatus(saved.status)
        ? pipeline.applyCaptureRoute(saved.enquiry_no, actor).row
        : saved;
      res.json({
        ok: true,
        row,
        nextEnquiryNo: nextEnquiryNo(),
        dropdowns: listEnquiryDropdowns(),
        duplicates: findOpenEnquiryDuplicates(row, row.enquiry_no)
      });
    } catch (e) {
      const incoming = req.body || {};
      const msg = e.message || String(e);
      const existing = /already/.test(msg) ? getEnquiry(incoming.enquiry_no) : null;
      if (existing && (incoming.create_only || incoming.createOnly)) {
        const sameClient = String(existing.client_name || "").trim().toLowerCase()
          === String(incoming.client_name || "").trim().toLowerCase();
        if (sameClient) {
          return res.json({
            ok: true,
            row: existing,
            nextEnquiryNo: nextEnquiryNo(),
            dropdowns: listEnquiryDropdowns(),
            replayed: true
          });
        }
      }
      res.status(400).json({
        ok: false,
        error: msg,
        existing: existing || undefined,
        nextEnquiryNo: nextEnquiryNo()
      });
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
    const done = String(req.query.done || "") === "1" || String(req.query.done || "").toLowerCase() === "true";
    const manager = staff.canManageUsers(req.office);
    const all = manager && String(req.query.scope || "all").toLowerCase() !== "mine";
    res.json({
      ok: true,
      done,
      scope: all ? "all" : "mine",
      canViewAll: manager,
      rows: done
        ? pipeline.listMyCompletedTasks(req.office.name, { all })
        : pipeline.listMyTasks(req.office.name, { all })
    });
  });

  app.get("/api/office/enquiries/:enquiryNo/process", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, ...pipeline.processSnapshot(req.params.enquiryNo, req.office.name) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/office/enquiries/:enquiryNo/process", requireOffice, (req, res) => {
    try {
      const snap = pipeline.applyAction(req.params.enquiryNo, req.office.name, req.body || {});
      res.json({ ok: true, ...snap });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/office/enquiries/:enquiryNo/create-order", requireOffice, (req, res) => {
    try {
      const out = createOrderFromEnquiry(req.params.enquiryNo);
      res.json({ ok: true, ...out });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/enquiries/:enquiryNo/create-order-draft", requireOffice, (req, res) => {
    try {
      const out = createOrderDraftFromEnquiry(req.params.enquiryNo);
      res.json({ ok: true, ...out });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/office/enquiries/:enquiryNo/create-orders", requireOffice, (req, res) => {
    try {
      const out = createOrdersFromEnquiryForm(req.params.enquiryNo, req.body || {});
      res.json({ ok: true, ...out });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/delivery-estimate", requireOffice, (req, res) => {
    try {
      const types = [].concat(req.query.type || req.query.types || []).filter(Boolean);
      res.json({
        ok: true,
        delivery: fromEnquiry.estimateDelivery({
          types,
          province: req.query.province || "",
          date: req.query.date || "",
          now: req.query.now || ""
        })
      });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/access-requests", requireOffice, (req, res) => {
    res.json({ ok: true, rows: pipeline.listAccessInbox(req.office.name) });
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
    const days = workdays(start, SCHEDULE_WORKDAYS);
    const fromDay = days[0];
    const toDay = days[days.length - 1];
    res.json({
      ok: true,
      start,
      days,
      rows: listSchedule(fromDay, toDay),
      codes: SCHEDULE_CODES
    });
  });

  app.get("/api/office/schedule/delivery", requireOffice, (_req, res) => {
    const { items, categories } = listDeliveryItems();
    const current = isoWeekInfo(mondayOf());
    res.json({
      ok: true,
      items,
      categories,
      weeks: weekOptions(items),
      currentWeekKey: weekKey(current),
      codes: SCHEDULE_CODES
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

  app.get("/api/office/floor-layout", (_req, res) => {
    res.json({ ok: true, layout: readFloorLayout() });
  });
  app.put("/api/office/floor-layout", requireOffice, (req, res) => {
    try {
      res.json({ ok: true, layout: writeFloorLayout((req.body && req.body.layout) || {}) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

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

  app.get("/api/office/paint-shop", requireOffice, (_req, res) => {
    try {
      res.json({ ok: true, ...paintShop.snapshot() });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/office/paint-shop/send", requireOffice, (req, res) => {
    try {
      const result = paintShop.sendToPaintShop(
        (req.body && (req.body.orderNumbers || req.body.orders)) || [],
        req.office && req.office.name
      );
      res.json({ ok: true, ...result });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/office/paint-shop/receive", requireOffice, (req, res) => {
    try {
      const result = paintShop.receiveFromPaintShop(req.body || {}, req.office && req.office.name);
      res.json({ ok: true, ...result });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/office/paint-shop/invoices/:invoiceId", requireOffice, (req, res) => {
    const file = paintShop.readInvoiceFile(req.params.invoiceId);
    if (!file) {
      res.status(404).json({ ok: false, error: "Invoice not found." });
      return;
    }
    const download = String((req.query && req.query.download) || "") === "1";
    res.setHeader("Content-Type", file.mime || "application/octet-stream");
    res.setHeader(
      "Content-Disposition",
      (download ? "attachment" : "inline") + "; filename=\"" + String(file.filename || "invoice").replace(/"/g, "") + "\""
    );
    res.send(file.buffer);
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
