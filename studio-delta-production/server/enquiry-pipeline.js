const db = require("./db");
const staff = require("./staff");

const FOLLOW_UP_DAYS = 7;

const WAITING_STATUSES = [
  "Waiting on clients personal details",
  "Waiting on clients specifictions",
  "Waiting on productions confirmation"
];

const CLOSED_STATUSES = ["Not within scope", "Not Interested", "Rejected"];

const TASK_TITLES = {
  chase_info: "Chase missing information",
  cost_sheet: "Complete costing",
  supplier: "Waiting on supplier",
  approval: "Approve cost sheet",
  quote: "Issue quote PDF",
  follow_up: "Follow up with client",
  pop: "Record client outcome",
  drawing: "Upload drawing"
};

function officeAssignees() {
  return staff.listUsers()
    .filter((u) => u.canSeeOffice && u.name)
    .map((u) => u.name)
    .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
}

function optionalAssignee(name) {
  const n = String(name || "").trim();
  if (!n) return "";
  return requireAssignee(n);
}

function requireAssignee(name) {
  const n = String(name || "").trim();
  if (!n) throw new Error("Choose the office person this task is assigned to");
  const hit = officeAssignees().find((x) => namesMatch(x, n));
  if (!hit) throw new Error("Assign to an office Admin from Users");
  return hit;
}

function namesMatch(a, b) {
  return String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase();
}

function namedProducts(row) {
  return (row.products || []).filter((p) => String(p.product || "").trim());
}

function requirePricedProducts(row) {
  const named = namedProducts(row);
  if (!named.length) throw new Error("Add at least one product first");
  if (named.some((p) => String(p.value_excl_vat || "").trim() === "")) {
    throw new Error("Enter a value excluding VAT for each product");
  }
  if (row.delivery_excl_vat === "" || row.delivery_excl_vat == null) {
    throw new Error("Delivery excluding VAT is required");
  }
}

function applyPricedBody(row, body) {
  if (Array.isArray(body.products) && body.products.length) {
    row.products = db.normalizeEnquiryLines({ products: body.products }, row);
  }
  if (body.delivery_excl_vat != null && String(body.delivery_excl_vat).trim() !== "") {
    row.delivery_excl_vat = db.money(db.parseMoney(body.delivery_excl_vat));
  }
  row.product = namedProducts(row).map((p) => p.product).join(", ");
  row.category = namedProducts(row).map((p) => p.category).filter(Boolean)[0] || row.category || "";
}

function nextTaskId(row) {
  const max = (row.tasks || []).reduce((m, t) => {
    const n = Number(String(t.id || "").replace(/\D/g, "")) || 0;
    return Math.max(m, n);
  }, 0);
  return "t" + (max + 1);
}

function addTask(row, kind, assignee, extra) {
  if (!Array.isArray(row.tasks)) row.tasks = [];
  const task = {
    id: nextTaskId(row),
    kind,
    title: (extra && extra.title) || TASK_TITLES[kind] || kind,
    assignee,
    status: "open",
    created_at: db.nowIso(),
    completed_at: "",
    completed_by: "",
    due_at: (extra && extra.due_at) || "",
    note: (extra && extra.note) || "",
    label: (extra && extra.label) || ""
  };
  row.tasks.push(task);
  return task;
}

function closeOpenKind(row, kind, actor, note) {
  for (const t of row.tasks || []) {
    if (t.kind === kind && t.status === "open") {
      t.status = "done";
      t.completed_at = db.nowIso();
      t.completed_by = actor;
      if (note) t.note = note;
    }
  }
}

function cancelOpenKind(row, kind) {
  for (const t of row.tasks || []) {
    if (t.kind === kind && t.status === "open") t.status = "cancelled";
  }
}

function openOfKind(row, kind) {
  return (row.tasks || []).find((t) => t.kind === kind && t.status === "open") || null;
}

function parseDmy(s) {
  return db.asDate(s);
}

function addDaysIso(fromDate, days) {
  const d = fromDate instanceof Date ? new Date(fromDate.getTime()) : parseDmy(fromDate);
  if (!d) return "";
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString();
}

function isOverdue(dueAt) {
  if (!dueAt) return false;
  const d = new Date(dueAt);
  if (isNaN(d.getTime())) {
    const parsed = parseDmy(dueAt);
    if (!parsed) return false;
    return Date.now() >= parsed.getTime();
  }
  return Date.now() >= d.getTime();
}

function followUpDueAt(row) {
  const list = Array.isArray(row.follow_ups) ? row.follow_ups : [];
  if (list.length) {
    const last = list[list.length - 1];
    return addDaysIso(last.uploaded_at || last.at || db.nowIso(), FOLLOW_UP_DAYS);
  }
  if (row.date_quoted) return addDaysIso(row.date_quoted, FOLLOW_UP_DAYS);
  return "";
}

function nextFollowUpLabel(count) {
  const n = Number(count) || 0;
  if (n <= 0) return "Follow up";
  return "Follow up x" + (n + 1);
}

function requireFile(body, message) {
  const raw = body.file_base64 || body.fileBase64 || "";
  if (!String(raw).trim()) throw new Error(message || "Upload a file, check the preview, then confirm it");
  if (!body.file_confirmed && !body.fileConfirmed) {
    throw new Error("Tick that this is the correct file before saving");
  }
  return raw;
}

function isSpreadsheet(name, mime) {
  const n = String(name || "").toLowerCase();
  const m = String(mime || "").toLowerCase();
  return /\.(xlsx|xls|csv)$/.test(n) || /spreadsheet|excel|csv/.test(m);
}

function isPdf(name, mime, buf) {
  const n = String(name || "").toLowerCase();
  const m = String(mime || "").toLowerCase();
  if (/\.pdf$/.test(n) || m.indexOf("pdf") >= 0) return true;
  return buf && buf.length >= 5 && buf.slice(0, 5).toString("utf8") === "%PDF-";
}

function isImage(name, mime) {
  const n = String(name || "").toLowerCase();
  const m = String(mime || "").toLowerCase();
  return /\.(png|jpe?g|webp|gif)$/.test(n) || m.indexOf("image/") === 0;
}

function statusAllows(row, list) {
  return list.indexOf(row.status) >= 0;
}

function availableActions(row) {
  const status = row.status || "New";
  const actions = [];
  const closed = CLOSED_STATUSES.indexOf(status) >= 0;
  if (closed) return actions;

  if (statusAllows(row, ["New"].concat(WAITING_STATUSES))) {
    actions.push({ id: "assign_waiting", label: "Assign someone to chase missing information" });
    if (namedProducts(row).length) {
      actions.push({ id: "assign_costing", label: "Assign costing" });
    }
    actions.push({ id: "close", label: "Close enquiry" });
  }
  if (openOfKind(row, "chase_info")) {
    actions.push({ id: "complete_chase", label: "Update chased information" });
  }
  if (!closed) {
    actions.push({ id: "add_correspondence", label: "Save email from Outlook" });
  }
  if (statusAllows(row, ["Costing", "Re-Cost"])) {
    actions.push({ id: "assign_costing", label: "Change costing person" });
    actions.push({ id: "complete_cost_sheet", label: "Upload cost sheet" });
    actions.push({ id: "supplier_wait", label: "Waiting on supplier" });
  }
  if (status === "Waiting on Supplier") {
    actions.push({ id: "complete_supplier", label: "Supplier answered — back to costing" });
  }
  if (openOfKind(row, "approval") || (status === "Costed" && row.approval && row.approval.status === "pending")) {
    actions.push({ id: "complete_approval", label: "Approve or reject costing" });
  }
  if (status === "Costed" && row.approval && row.approval.status === "approved") {
    actions.push({ id: "complete_quote", label: "Upload quote PDF" });
  }
  if (statusAllows(row, ["Quoted", "Followed Up"])) {
    actions.push({ id: "complete_followup", label: "Log a follow-up screenshot" });
    actions.push({ id: "complete_order", label: "Client approved — attach POP" });
    actions.push({ id: "complete_reject", label: "Client rejected" });
  }
  if (status === "Ordered" && row.drawing && row.drawing.required && !(row.drawing.file && row.drawing.file.stored_as)) {
    actions.push({ id: "complete_drawing", label: "Upload drawing" });
  }
  if (status === "Ordered" && row.drawing && row.drawing.required === false && !row.ready_for_orders) {
    row.ready_for_orders = true;
  }
  return actions;
}

function listMyTasks(userName) {
  const me = String(userName || "").trim();
  const out = [];
  for (const row of db.listEnquiries()) {
    const tasks = Array.isArray(row.tasks) ? row.tasks : [];
    for (const task of tasks) {
      if (task.status !== "open") continue;
      if (!namesMatch(task.assignee, me)) continue;
      const dueAt = task.due_at || (task.kind === "follow_up" ? followUpDueAt(row) : "");
      out.push(decorateTask(row, task, dueAt));
    }
    if (statusAllows(row, ["Quoted", "Followed Up"]) && namesMatch(row.follow_up_assignee, me) && !openOfKind(row, "follow_up")) {
      const dueAt = followUpDueAt(row);
      if (dueAt && isOverdue(dueAt)) {
        out.push(decorateTask(row, {
          id: "follow-due",
          kind: "follow_up",
          title: nextFollowUpLabel((row.follow_ups || []).length),
          assignee: row.follow_up_assignee,
          status: "open",
          created_at: row.date_quoted || "",
          due_at: dueAt,
          note: "Quote or last follow-up is 7 or more days old"
        }, dueAt));
      }
    }
  }
  out.sort((a, b) => {
    if (a.overdue !== b.overdue) return a.overdue ? -1 : 1;
    return String(b.created_at || "").localeCompare(String(a.created_at || ""));
  });
  return out;
}

function decorateTask(row, task, dueAt) {
  const correspondence = db.normalizeCorrespondence(row);
  return {
    ...task,
    due_at: dueAt || task.due_at || "",
    overdue: isOverdue(dueAt || task.due_at),
    enquiry_no: row.enquiry_no,
    client_name: row.client_name || "",
    product: row.product || "",
    enquiry_status: row.status || "",
    date_quoted: row.date_quoted || "",
    correspondence_mails: correspondence.mails.length
  };
}

function processSnapshot(enquiryNo) {
  const row = db.getEnquiry(enquiryNo);
  if (!row) throw new Error("Enquiry not found");
  return {
    row,
    assignees: officeAssignees(),
    actions: availableActions(row),
    waitingStatuses: WAITING_STATUSES,
    closedStatuses: CLOSED_STATUSES.filter((s) => s !== "Rejected"),
    followUpDays: FOLLOW_UP_DAYS,
    quoteNo: db.quoteNoHint(),
    outlookAddin: { manifest: "/outlook-addin/manifest.xml", install: "/outlook-addin" }
  };
}

function applyAction(enquiryNo, actorName, body) {
  const actor = String(actorName || "").trim();
  if (!actor) throw new Error("Not signed in");
  const raw = db.getEnquiryRaw(enquiryNo);
  if (!raw) throw new Error("Enquiry not found. Save the enquiry first.");
  if (!Array.isArray(raw.tasks)) raw.tasks = [];
  if (!Array.isArray(raw.follow_ups)) raw.follow_ups = [];
  const action = String((body && body.action) || "").trim();
  const handlers = {
    assign_waiting: assignWaiting,
    assign_costing: assignCosting,
    add_correspondence: addCorrespondence,
    complete_chase: completeChase,
    supplier_wait: supplierWait,
    complete_supplier: completeSupplier,
    complete_cost_sheet: completeCostSheet,
    complete_approval: completeApproval,
    complete_quote: completeQuote,
    complete_followup: completeFollowup,
    complete_reject: completeReject,
    complete_order: completeOrder,
    complete_drawing: completeDrawing,
    close: closeEnquiry,
    reassign: reassignTask
  };
  const fn = handlers[action];
  if (!fn) throw new Error("Unknown process action");
  fn(raw, actor, body || {});
  raw.updated_at = db.nowIso();
  db.saveEnquiryRecord(raw);
  return processSnapshot(raw.enquiry_no);
}

function assignWaiting(row, actor, body) {
  if (CLOSED_STATUSES.indexOf(row.status) >= 0) throw new Error("This enquiry is closed");
  if (!statusAllows(row, ["New"].concat(WAITING_STATUSES))) {
    throw new Error("Missing-info waiting is only used during capture, before costing");
  }
  const waiting = WAITING_STATUSES.find((s) => s === body.waiting_status) || WAITING_STATUSES.find((s) => namesMatch(s, body.waiting_status));
  if (!waiting) throw new Error("Choose what you are waiting on");
  const assignee = requireAssignee(body.assignee);
  cancelOpenKind(row, "chase_info");
  row.status = waiting;
  addTask(row, "chase_info", assignee, { note: waiting });
}

function assignCosting(row, actor, body) {
  if (!statusAllows(row, ["New"].concat(WAITING_STATUSES).concat(["Costing", "Re-Cost"]))) {
    throw new Error("Costing is assigned from capture, or changed while the enquiry is still in costing");
  }
  if (!namedProducts(row).length) throw new Error("Add at least one product name before assigning costing");
  const assignee = requireAssignee(body.assignee);
  archiveCorrespondence(row, actor, body);
  const open = openOfKind(row, "cost_sheet");
  if (open && statusAllows(row, ["Costing", "Re-Cost"])) {
    open.assignee = assignee;
    return;
  }
  cancelOpenKind(row, "chase_info");
  cancelOpenKind(row, "cost_sheet");
  row.status = "Costing";
  addTask(row, "cost_sheet", assignee);
}

function archiveCorrespondence(row, actor, body) {
  const existing = db.normalizeCorrespondence(row);
  const next = {
    saved_at: existing.saved_at,
    saved_by: existing.saved_by,
    mails: existing.mails.slice()
  };
  const incoming = [];
  if (Array.isArray(body && body.correspondence_mails)) {
    incoming.push.apply(incoming, body.correspondence_mails);
  }
  incoming.push.apply(incoming, db.mailsFromPastedLinks((body && (body.correspondence_links || body.correspondenceLinks)) || ""));
  const seen = new Set(next.mails.map((m) => db.mailDedupeKey(m)));
  let added = 0;
  incoming.forEach((item, i) => {
    const mail = db.normalizeOutlookMail(item, next.mails.length + i);
    if (!mail) return;
    const key = db.mailDedupeKey(mail);
    if (!key || seen.has(key)) return;
    seen.add(key);
    mail.id = "mail_" + (next.mails.length + 1);
    next.mails.push(mail);
    added += 1;
  });
  if (added) {
    next.saved_at = db.nowIso();
    next.saved_by = actor;
  }
  if (next.mails.length) row.correspondence = next;
  return added;
}

function addCorrespondence(row, actor, body) {
  const added = archiveCorrespondence(row, actor, body);
  if (!added) {
    throw new Error("Paste Outlook’s Copy as link. Do not upload .msg files.");
  }
}

function completeChase(row, actor, body) {
  const task = openOfKind(row, "chase_info");
  if (!task) throw new Error("No open chase task");
  const next = String(body.next || "").trim();
  closeOpenKind(row, "chase_info", actor, body.comments || "");
  if (next === "costing") {
    if (!namedProducts(row).length) throw new Error("Add the product name(s) on the enquiry before sending to costing");
    archiveCorrespondence(row, actor, body);
    row.status = "Costing";
    addTask(row, "cost_sheet", requireAssignee(body.assignee));
    return;
  }
  if (next === "waiting") {
    const waiting = WAITING_STATUSES.find((s) => s === body.waiting_status) || WAITING_STATUSES[0];
    row.status = waiting;
    addTask(row, "chase_info", requireAssignee(body.assignee || task.assignee), { note: waiting });
    return;
  }
  throw new Error("Choose whether this can now go to costing, or is still waiting");
}

function supplierWait(row, _actor, body) {
  if (!statusAllows(row, ["Costing", "Re-Cost"])) {
    throw new Error("Waiting on Supplier is only used during costing");
  }
  const assignee = requireAssignee(body.assignee || (openOfKind(row, "cost_sheet") || {}).assignee);
  cancelOpenKind(row, "cost_sheet");
  row.status = "Waiting on Supplier";
  addTask(row, "supplier", assignee);
}

function completeSupplier(row, actor, body) {
  if (row.status !== "Waiting on Supplier") throw new Error("This enquiry is not waiting on a supplier");
  closeOpenKind(row, "supplier", actor, body.comments || "");
  row.status = "Costing";
  addTask(row, "cost_sheet", requireAssignee(body.assignee || lastAssignee(row, "cost_sheet")));
}

function completeCostSheet(row, actor, body) {
  if (!statusAllows(row, ["Costing", "Re-Cost"])) throw new Error("Upload the cost sheet from Costing");
  const filename = body.file_name || body.filename || "cost-sheet.xlsx";
  const raw = requireFile(body, "Upload the Excel cost sheet, check the preview, then confirm it");
  if (!isSpreadsheet(filename, "") && !isPdf(filename, "", null) && !/\.csv$/i.test(filename)) {
    throw new Error("Cost sheet must be Excel (xlsx / xls), CSV, or PDF");
  }
  const approver = optionalAssignee(body.assignee);
  const quoter = optionalAssignee(body.quote_assignee);
  if (!quoter) throw new Error("Choose the quoting person");
  row.cost_sheet = db.saveEnquiryAttachment(row.enquiry_no, "cost_sheet", raw, filename);
  row.quote_assignee = quoter;
  closeOpenKind(row, "cost_sheet", actor);
  cancelOpenKind(row, "approval");
  cancelOpenKind(row, "quote");
  row.status = "Costed";
  if (approver) {
    row.approval = {
      requested_from: approver,
      requested_at: db.nowIso(),
      requested_by: actor,
      status: "pending",
      comments: "",
      decided_by: "",
      decided_at: ""
    };
    addTask(row, "approval", approver, { note: "Approve or reject the cost sheet" });
    return;
  }
  row.approval = {
    requested_from: "",
    requested_at: db.nowIso(),
    requested_by: actor,
    status: "approved",
    comments: "Approval skipped",
    decided_by: actor,
    decided_at: db.nowIso()
  };
  addTask(row, "quote", quoter);
}

function completeApproval(row, actor, body) {
  if (row.status !== "Costed" || !row.approval || row.approval.status !== "pending") {
    throw new Error("There is no cost sheet waiting for approval");
  }
  const decision = String(body.decision || "").trim().toLowerCase();
  const comments = String(body.comments || "").trim();
  closeOpenKind(row, "approval", actor, comments);
  if (decision === "reject" || decision === "rejected") {
    if (!comments) throw new Error("Comments are required when costing is rejected");
    row.approval.status = "rejected";
    row.approval.comments = comments;
    row.approval.decided_by = actor;
    row.approval.decided_at = db.nowIso();
    row.status = "Re-Cost";
    const coster = requireAssignee(body.assignee || lastAssignee(row, "cost_sheet"));
    addTask(row, "cost_sheet", coster, { note: comments });
    return;
  }
  if (decision !== "approve" && decision !== "approved") {
    throw new Error("Choose approve or reject");
  }
  row.approval.status = "approved";
  row.approval.comments = comments;
  row.approval.decided_by = actor;
  row.approval.decided_at = db.nowIso();
  const quotePerson = requireAssignee(body.assignee || row.quote_assignee);
  row.quote_assignee = quotePerson;
  cancelOpenKind(row, "quote");
  addTask(row, "quote", quotePerson);
}

function lastAssignee(row, kind) {
  const list = (row.tasks || []).filter((t) => t.kind === kind && t.assignee);
  return list.length ? list[list.length - 1].assignee : "";
}

function completeQuote(row, actor, body) {
  if (row.status !== "Costed" || !row.approval || row.approval.status !== "approved") {
    throw new Error("The cost sheet must be approved before a quote PDF is issued");
  }
  applyPricedBody(row, body);
  requirePricedProducts(row);
  const quoteNo = db.requireUniqueQuoteNo(body.quote_no, row.enquiry_no);
  row.quote_no = quoteNo;
  const followPerson = requireAssignee(body.follow_up_assignee || body.assignee);
  const payload = {
    ...row,
    status: "Quoted",
    quote_pdf_base64: body.file_base64 || body.quote_pdf_base64,
    quote_pdf_name: body.file_name || body.quote_pdf_name || "quote.pdf",
    quote_pdf_confirmed: !!(body.file_confirmed || body.quote_pdf_confirmed)
  };
  if (!payload.quote_pdf_base64) throw new Error("Upload the quote PDF, check the preview, then confirm it is the correct file");
  db.upsertEnquiry(payload, { fromPipeline: true });
  const saved = db.getEnquiryRaw(row.enquiry_no);
  Object.assign(row, saved);
  closeOpenKind(row, "quote", actor);
  row.follow_up_assignee = followPerson;
  cancelOpenKind(row, "follow_up");
  addTask(row, "follow_up", followPerson, {
    title: "Follow up",
    due_at: addDaysIso(row.date_quoted || db.todayEnquiryDate(), FOLLOW_UP_DAYS)
  });
  addTask(row, "pop", followPerson, { title: "Record client outcome" });
}

function completeFollowup(row, actor, body) {
  if (!statusAllows(row, ["Quoted", "Followed Up"])) throw new Error("Follow-ups start after the quote PDF is issued");
  const filename = body.file_name || "follow-up.png";
  const raw = requireFile(body, "Upload a screenshot of the follow-up");
  if (!isImage(filename, "") && !isPdf(filename, "", null)) {
    throw new Error("Follow-up proof must be a screenshot (image) or PDF");
  }
  const n = (row.follow_ups || []).length + 1;
  const file = db.saveEnquiryAttachment(row.enquiry_no, "follow_up_" + n, raw, filename);
  row.follow_ups.push({
    n,
    label: n === 1 ? "Follow up" : "Follow up x" + n,
    uploaded_at: db.nowIso(),
    by: actor,
    file
  });
  closeOpenKind(row, "follow_up", actor);
  row.status = "Followed Up";
  const assignee = requireAssignee(body.assignee || row.follow_up_assignee || actor);
  row.follow_up_assignee = assignee;
  addTask(row, "follow_up", assignee, {
    title: nextFollowUpLabel(n),
    label: n === 1 ? "x2" : "x" + (n + 1),
    due_at: addDaysIso(db.nowIso(), FOLLOW_UP_DAYS)
  });
}

function completeReject(row, actor, body) {
  if (!statusAllows(row, ["Quoted", "Followed Up"])) throw new Error("Client rejection is recorded after a quote");
  const reason = String(body.reason || body.comments || "").trim();
  if (!reason) throw new Error("A rejection reason is required");
  row.status = "Rejected";
  row.client_outcome = { kind: "rejected", reason, decided_at: db.nowIso(), decided_by: actor };
  cancelOpenKind(row, "follow_up");
  cancelOpenKind(row, "pop");
  closeOpenKind(row, "pop", actor, reason);
}

function completeOrder(row, actor, body) {
  if (!statusAllows(row, ["Quoted", "Followed Up"])) throw new Error("Attach proof of payment after the client approves the quote");
  const filename = body.file_name || "pop.pdf";
  const raw = requireFile(body, "Upload proof of payment (screenshot or PDF)");
  if (!isImage(filename, "") && !isPdf(filename, "", null)) {
    throw new Error("Proof of payment must be a screenshot or PDF");
  }
  const file = db.saveEnquiryAttachment(row.enquiry_no, "pop", raw, filename);
  row.status = "Ordered";
  row.client_outcome = {
    kind: "approved",
    reason: "",
    decided_at: db.nowIso(),
    decided_by: actor,
    file
  };
  cancelOpenKind(row, "follow_up");
  closeOpenKind(row, "pop", actor);
  const drawingRaw = body.drawing_required;
  const needsDrawing = drawingRaw === true || drawingRaw === "yes" || drawingRaw === "true";
  const noDrawing = drawingRaw === false || drawingRaw === "no" || drawingRaw === "false";
  if (!needsDrawing && !noDrawing) throw new Error("Say whether this order requires a drawing");
  if (!needsDrawing) {
    row.drawing = { required: false, file: null };
    row.ready_for_orders = true;
    return;
  }
  row.drawing = { required: true, file: null, assignee: requireAssignee(body.assignee) };
  row.ready_for_orders = false;
  addTask(row, "drawing", row.drawing.assignee);
}

function completeDrawing(row, actor, body) {
  if (row.status !== "Ordered" || !row.drawing || !row.drawing.required) {
    throw new Error("A drawing is only required when the office said this order needs one");
  }
  const filename = body.file_name || "drawing.pdf";
  const raw = requireFile(body, "Upload the drawing, check the preview, then confirm it");
  if (!isImage(filename, "") && !isPdf(filename, "", null)) {
    throw new Error("Drawing must be a PDF or image");
  }
  row.drawing.file = db.saveEnquiryAttachment(row.enquiry_no, "drawing", raw, filename);
  row.drawing.uploaded_at = db.nowIso();
  row.drawing.uploaded_by = actor;
  closeOpenKind(row, "drawing", actor);
  row.ready_for_orders = true;
}

function closeEnquiry(row, actor, body) {
  if (!statusAllows(row, ["New"].concat(WAITING_STATUSES))) {
    throw new Error("Not within scope / Not Interested is used from capture");
  }
  const status = CLOSED_STATUSES.filter((s) => s !== "Rejected").find((s) => s === body.status || namesMatch(s, body.status));
  if (!status) throw new Error("Choose Not within scope or Not Interested");
  row.status = status;
  row.client_outcome = { kind: "closed", reason: String(body.reason || body.comments || "").trim(), decided_at: db.nowIso(), decided_by: actor };
  (row.tasks || []).forEach((t) => {
    if (t.status === "open") t.status = "cancelled";
  });
}

function reassignTask(row, _actor, body) {
  const id = String(body.task_id || "").trim();
  const task = (row.tasks || []).find((t) => t.id === id && t.status === "open");
  if (!task) throw new Error("Open task not found");
  task.assignee = requireAssignee(body.assignee);
  if (task.kind === "follow_up") row.follow_up_assignee = task.assignee;
}

module.exports = {
  FOLLOW_UP_DAYS,
  WAITING_STATUSES,
  CLOSED_STATUSES,
  officeAssignees,
  listMyTasks,
  processSnapshot,
  applyAction,
  availableActions
};
