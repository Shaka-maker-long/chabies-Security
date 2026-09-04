const db = require("./db");
const staff = require("./staff");
const access = require("./enquiry-access");

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

function assigneeFromRole(role, requested, preferred) {
  const chosen = optionalAssignee(requested);
  if (chosen) return chosen;
  return optionalAssignee(staff.defaultEnquiryAssignee(role, preferred));
}

function requireRoleAssignee(role, requested, preferred) {
  const n = assigneeFromRole(role, requested, preferred);
  if (!n && canonicalizeRoleLabel(role) === "Quoting") {
    throw new Error("Choose the quoting person");
  }
  return requireAssignee(n);
}

function canonicalizeRoleLabel(role) {
  const t = String(role || "").trim().toLowerCase();
  if (t === "quoting" || t === "quote" || t === "quoter") return "Quoting";
  if (t === "costing" || t === "coster") return "Costing";
  if (t === "approval" || t === "approver") return "Approval";
  return String(role || "").trim();
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
  if (named.some((p) => String(p.value_incl_vat || p.value_excl_vat || "").trim() === "")) {
    throw new Error("Enter a value including VAT for each product");
  }
  if (
    (row.delivery_incl_vat === "" || row.delivery_incl_vat == null) &&
    (row.delivery_excl_vat === "" || row.delivery_excl_vat == null)
  ) {
    throw new Error("Delivery including VAT is required");
  }
}

function applyPricedBody(row, body) {
  if (Array.isArray(body.products) && body.products.length) {
    row.products = db.normalizeEnquiryLines({ products: body.products }, row);
  }
  if (
    (body.delivery_incl_vat != null && String(body.delivery_incl_vat).trim() !== "") ||
    (body.delivery_excl_vat != null && String(body.delivery_excl_vat).trim() !== "")
  ) {
    const pair = db.vatPair(body.delivery_incl_vat, body.delivery_excl_vat);
    row.delivery_excl_vat = pair.excl;
    row.delivery_incl_vat = pair.incl;
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

const DELIVERABLE_ACTIONS = {
  complete_chase: "chase_info",
  complete_cost_sheet: "cost_sheet",
  supplier_wait: "cost_sheet",
  complete_supplier: "supplier",
  complete_approval: "approval",
  complete_quote: "quote",
  complete_followup: "follow_up",
  complete_reject: "pop",
  complete_order: "pop",
  complete_drawing: "drawing"
};

function isManagerName(name) {
  return staff.canManageUsers({ name: String(name || "").trim() });
}

function actionKind(actionId) {
  return DELIVERABLE_ACTIONS[actionId] || "";
}

function actionOwner(row, actionId) {
  const kind = actionKind(actionId);
  if (!kind) return "";
  const open = openOfKind(row, kind);
  if (open && open.assignee) return open.assignee;
  if (kind === "quote") return row.quote_assignee || "";
  if (kind === "follow_up" || kind === "pop") {
    return row.follow_up_assignee || row.quote_assignee || lastAssignee(row, kind) || "";
  }
  if (kind === "drawing") return (row.drawing && row.drawing.assignee) || lastAssignee(row, "drawing") || "";
  if (kind === "approval") return (row.approval && row.approval.requested_from) || lastAssignee(row, "approval") || "";
  return lastAssignee(row, kind) || "";
}

function canAct(row, actor, actionId) {
  if (!DELIVERABLE_ACTIONS[actionId]) return true;
  const owner = actionOwner(row, actionId);
  if (!owner) return true;
  const who = String(actor || "").trim();
  if (!who) return false;
  if (namesMatch(owner, who)) return true;
  if (isManagerName(who)) return true;
  return access.grantedFor(row, who, actionKind(actionId));
}

function assertCanAct(row, actor, actionId) {
  if (canAct(row, actor, actionId)) return;
  const owner = actionOwner(row, actionId) || "the assigned person";
  throw new Error(
    owner + " is assigned this step. Ask them or the Manager to grant you access, then upload the deliverable."
  );
}

function decorateActions(row, actor) {
  return availableActions(row).map((action) => {
    const kind = actionKind(action.id);
    const owner = actionOwner(row, action.id);
    const locked = !!(kind && owner);
    const pending = locked ? access.pendingFor(row, actor, kind) : null;
    return {
      ...action,
      kind,
      assignee: owner,
      can_act: !locked || canAct(row, actor, action.id),
      request_pending: !!(pending && pending.id),
      grant_id: pending ? pending.id : ""
    };
  });
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
    actions.push({ id: "complete_quote", label: "Issue another quote" });
    actions.push({ id: "assign_costing", label: "Client wants changes — recost" });
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
    correspondence_mails: correspondence.mails.length,
    deliverable_count: (row.deliverable_count != null ? row.deliverable_count : correspondence.mails.length),
    completed_at_label: task.completed_at ? db.formatSastDateTime(task.completed_at) : ""
  };
}

function listMyCompletedTasks(userName) {
  const me = String(userName || "").trim();
  const out = [];
  for (const row of db.listEnquiries()) {
    const tasks = Array.isArray(row.tasks) ? row.tasks : [];
    for (const task of tasks) {
      if (task.status !== "done") continue;
      if (!namesMatch(task.assignee, me) && !namesMatch(task.completed_by, me)) continue;
      out.push(decorateTask(row, task, task.due_at));
    }
  }
  out.sort((a, b) => String(b.completed_at || "").localeCompare(String(a.completed_at || "")));
  return out;
}

function processSnapshot(enquiryNo, actorName) {
  const row = db.getEnquiry(enquiryNo);
  if (!row) throw new Error("Enquiry not found");
  const actor = String(actorName || "").trim();
  const manager = actor ? isManagerName(actor) : false;
  return {
    row,
    me: actor,
    is_manager: manager,
    assignees: officeAssignees(),
    enquiryRoles: staff.enquiryRoleDefaults(),
    actions: decorateActions(row, actor),
    access: access.snapshotFor(row, actor, manager),
    waitingStatuses: WAITING_STATUSES,
    closedStatuses: CLOSED_STATUSES.filter((s) => s !== "Rejected"),
    followUpDays: FOLLOW_UP_DAYS,
    quoteNo: db.quoteNoHint(),
    outlookAddin: { manifest: "/outlook-addin/manifest.xml", install: "/outlook-addin" }
  };
}

function listAccessInbox(userName) {
  const actor = String(userName || "").trim();
  return access.inboxFor(db.listEnquiries(), actor, actor ? isManagerName(actor) : false);
}

function applyAction(enquiryNo, actorName, body) {
  const actor = String(actorName || "").trim();
  if (!actor) throw new Error("Not signed in");
  const raw = db.getEnquiryRaw(enquiryNo);
  if (!raw) throw new Error("Enquiry not found. Save the enquiry first.");
  if (!Array.isArray(raw.tasks)) raw.tasks = [];
  if (!Array.isArray(raw.follow_ups)) raw.follow_ups = [];
  if (!Array.isArray(raw.quotes)) raw.quotes = [];
  const action = String((body && body.action) || "").trim();
  if (action === "request_access") {
    const forAction = String((body && body.for_action) || "").trim();
    const kind = String((body && body.kind) || actionKind(forAction) || "").trim();
    const owner = (forAction ? actionOwner(raw, forAction) : "")
      || (openOfKind(raw, kind) && openOfKind(raw, kind).assignee)
      || "";
    const grant = access.requestAccess(raw, actor, kind, owner);
    db.appendEnquiryEvent(raw, {
      kind: "request_access",
      actor,
      status: raw.status || "",
      label: "Asked " + (grant.assignee || "assignee") + " for access to " + access.kindLabel(kind),
      note: ""
    });
    raw.updated_at = db.nowIso();
    db.saveEnquiryRecord(raw);
    return processSnapshot(raw.enquiry_no, actor);
  }
  if (action === "grant_access" || action === "deny_access") {
    const grant = action === "grant_access"
      ? access.grantAccess(raw, actor, body && body.grant_id, isManagerName(actor))
      : access.denyAccess(raw, actor, body && body.grant_id, isManagerName(actor));
    db.appendEnquiryEvent(raw, {
      kind: action,
      actor,
      status: raw.status || "",
      label: (action === "grant_access" ? "Granted " : "Refused ") + (grant.requester || "access") + " — " + access.kindLabel(grant.kind),
      note: ""
    });
    raw.updated_at = db.nowIso();
    db.saveEnquiryRecord(raw);
    return processSnapshot(raw.enquiry_no, actor);
  }
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
  assertCanAct(raw, actor, action);
  if (!Array.isArray(raw.events) || !raw.events.length) {
    raw.created_at = raw.created_at || db.nowIso();
    db.appendEnquiryEvent(raw, {
      kind: "created",
      actor: "",
      status: "New",
      label: "Enquiry captured"
    });
  }
  const fromStatus = raw.status || "New";
  fn(raw, actor, body || {});
  db.appendEnquiryEvent(raw, {
    kind: action,
    actor,
    from_status: fromStatus,
    status: raw.status || fromStatus,
    label: eventLabel(action, raw, fromStatus, body || {}),
    note: String((body && (body.comments || body.reason || body.note)) || "").trim()
  });
  const kind = actionKind(action);
  if (kind) access.consumeKind(raw, kind);
  raw.updated_at = db.nowIso();
  db.saveEnquiryRecord(raw);
  return processSnapshot(raw.enquiry_no, actor);
}

function eventLabel(action, row, fromStatus, body) {
  const status = row.status || "";
  const coster = (openOfKind(row, "cost_sheet") || {}).assignee || "";
  if (action === "assign_waiting") return "Waiting: " + status;
  if (action === "assign_costing") {
    return (fromStatus === "Costing" || fromStatus === "Re-Cost") && status === fromStatus
      ? "Costing assigned to " + coster
      : "Assigned costing → " + coster;
  }
  if (action === "add_correspondence") return "Correspondance link saved";
  if (action === "complete_chase") {
    return String(body.next || "") === "costing"
      ? "Chase complete — sent to costing"
      : "Still waiting: " + status;
  }
  if (action === "supplier_wait") return "Waiting on supplier";
  if (action === "complete_supplier") return "Supplier answered — back to costing";
  if (action === "complete_cost_sheet") return "Cost sheet uploaded";
  if (action === "complete_approval") {
    const d = String(body.decision || "").toLowerCase();
    return d.indexOf("reject") >= 0 ? "Costing rejected — Re-Cost" : "Costing approved";
  }
  if (action === "complete_quote") {
    const n = Array.isArray(row.quotes) ? row.quotes.length : 0;
    return (n > 1 ? "Quote " + n + " issued" : "Quote PDF issued") + (row.quote_no ? " " + row.quote_no : "");
  }
  if (action === "complete_followup") return "Follow-up logged";
  if (action === "complete_reject") return "Client rejected";
  if (action === "complete_order") {
    return row.drawing && row.drawing.required ? "POP saved — drawing required" : "POP saved — ready for Orders";
  }
  if (action === "complete_drawing") return "Drawing uploaded — ready for Orders";
  if (action === "close") return "Closed: " + status;
  if (action === "reassign") return "Task reassigned";
  return action;
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
  access.cancelKind(row, "chase_info");
  row.status = waiting;
  addTask(row, "chase_info", assignee, { note: waiting });
}

function assignCosting(row, actor, body) {
  if (statusAllows(row, ["Quoted", "Followed Up"])) {
    const assignee = requireRoleAssignee("Costing", body.assignee, lastAssignee(row, "cost_sheet"));
    cancelOpenKind(row, "follow_up");
    cancelOpenKind(row, "pop");
    cancelOpenKind(row, "quote");
    access.cancelKind(row, "cost_sheet");
    row.status = "Re-Cost";
    addTask(row, "cost_sheet", assignee, { note: "Client wants changes — recost for another quote" });
    return;
  }
  if (!statusAllows(row, ["New"].concat(WAITING_STATUSES).concat(["Costing", "Re-Cost"]))) {
    throw new Error("Costing is assigned from capture, or changed while the enquiry is still in costing");
  }
  if (!namedProducts(row).length) throw new Error("Add at least one product name before assigning costing");
  const assignee = requireRoleAssignee("Costing", body.assignee, (openOfKind(row, "cost_sheet") || {}).assignee);
  archiveCorrespondence(row, actor, body);
  const open = openOfKind(row, "cost_sheet");
  if (open && statusAllows(row, ["Costing", "Re-Cost"])) {
    if (!namesMatch(open.assignee, assignee)) access.cancelKind(row, "cost_sheet");
    open.assignee = assignee;
    return;
  }
  cancelOpenKind(row, "chase_info");
  cancelOpenKind(row, "cost_sheet");
  access.cancelKind(row, "cost_sheet");
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
  if (Array.isArray(body && body.correspondence_files)) {
    for (const item of body.correspondence_files) {
      const raw = item && (item.file_base64 || item.fileBase64);
      if (!raw) continue;
      const filename = String(item.file_name || item.filename || "outlook-email.msg").trim() || "outlook-email.msg";
      const extracted = db.extractOutlookFromDataUrl(raw, filename);
      const decoded = db.decodeDataUrl(raw);
      const big = decoded && decoded.buffer && decoded.buffer.length > 64;
      if (!extracted && !big) continue;
      const n = next.mails.length + incoming.length + 1;
      const saved = db.saveEnquiryAttachment(row.enquiry_no, "correspondence_" + n, raw, filename);
      incoming.push(Object.assign({}, extracted || {}, saved, {
        title: (extracted && extracted.title) || saved.title || filename
      }));
    }
  }
  if (Array.isArray(body && body.correspondence_mails)) {
    incoming.push.apply(incoming, body.correspondence_mails);
  }
  incoming.push.apply(incoming, db.mailsFromPastedLinks((body && (body.correspondence_links || body.correspondenceLinks)) || ""));
  const indexByKey = new Map();
  function indexMail(mail, idx) {
    (db.mailDedupeKeys ? db.mailDedupeKeys(mail) : [db.mailDedupeKey(mail)]).forEach((k) => {
      if (k) indexByKey.set(k, idx);
    });
  }
  next.mails.forEach((m, i) => indexMail(m, i));
  let added = 0;
  incoming.forEach((item, i) => {
    const mail = db.normalizeOutlookMail(item, next.mails.length + i);
    if (!mail) return;
    const keys = db.mailDedupeKeys ? db.mailDedupeKeys(mail) : [db.mailDedupeKey(mail)];
    if (!keys.length) return;
    let idx = -1;
    for (let k = 0; k < keys.length; k++) {
      if (indexByKey.has(keys[k])) {
        idx = indexByKey.get(keys[k]);
        break;
      }
    }
    if (idx >= 0) {
      const prev = next.mails[idx];
      if (mail.stored_as) {
        next.mails[idx] = Object.assign({}, prev, mail, { id: prev.id });
        indexMail(next.mails[idx], idx);
        added += 1;
      }
      return;
    }
    mail.id = "mail_" + (next.mails.length + 1);
    indexMail(mail, next.mails.length);
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
    throw new Error("Paste the Correspondance link, then Save update.");
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
    addTask(row, "cost_sheet", requireRoleAssignee("Costing", body.assignee));
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

function existingCostSheets(row) {
  if (Array.isArray(row.cost_sheets) && row.cost_sheets.length) {
    return row.cost_sheets.map((s, i) => ({
      n: Number(s && s.n) || i + 1,
      product: String((s && s.product) || "").trim(),
      kind: (s && s.kind) || ("cost_sheet_" + (Number(s && s.n) || i + 1)),
      stored_as: s && s.stored_as,
      filename: s && s.filename,
      mime: (s && s.mime) || "",
      uploaded_at: (s && s.uploaded_at) || "",
      uploaded_by: (s && s.uploaded_by) || "",
      size: (s && s.size) || 0
    })).filter((s) => s.stored_as);
  }
  if (row.cost_sheet && row.cost_sheet.stored_as) {
    const first = namedProducts(row)[0];
    return [{
      n: 1,
      product: (first && first.product) || "",
      kind: row.cost_sheet.kind || "cost_sheet",
      stored_as: row.cost_sheet.stored_as,
      filename: row.cost_sheet.filename,
      mime: row.cost_sheet.mime || "",
      uploaded_at: row.cost_sheet.uploaded_at || "",
      uploaded_by: row.cost_sheet.uploaded_by || "",
      size: row.cost_sheet.size || 0
    }];
  }
  return [];
}

function asCostGroups(row, body) {
  const named = namedProducts(row).map((p) => p.product);
  if (Array.isArray(body && body.cost_sheets) && body.cost_sheets.length) {
    return body.cost_sheets.map((group) => ({
      product: String((group && group.product) || "").trim(),
      files: Array.isArray(group && group.files)
        ? group.files
        : (group && (group.file_base64 || group.fileBase64) ? [group] : [])
    }));
  }
  return [{
    product: named[0] || "",
    files: (body && (body.file_base64 || body.fileBase64)) ? [body] : []
  }];
}

function completeCostSheet(row, actor, body) {
  if (!statusAllows(row, ["Costing", "Re-Cost"])) throw new Error("Upload the cost sheet from Costing");
  const named = namedProducts(row).map((p) => p.product);
  if (!named.length) throw new Error("Add at least one product name before uploading cost sheets");
  const groups = asCostGroups(row, body);
  const byProduct = new Map(groups.map((g) => [g.product, g]));
  const uploads = [];
  for (const product of named) {
    const group = byProduct.get(product)
      || (named.length === 1 && groups[0] && !groups[0].product ? groups[0] : null);
    const files = ((group && group.files) || []).filter((f) => f && (f.file_base64 || f.fileBase64));
    if (!files.length) throw new Error("Upload a cost sheet for " + product);
    for (const file of files) {
      if (!file.file_confirmed && !file.fileConfirmed) {
        throw new Error("Tick that this is the correct file before saving");
      }
      const filename = file.file_name || file.filename || "cost-sheet.xlsx";
      if (!isSpreadsheet(filename, file.file_type || "") && !isPdf(filename, file.file_type || "", null) && !/\.csv$/i.test(filename)) {
        throw new Error("Cost sheet must be Excel (xlsx / xls), CSV, or PDF");
      }
      uploads.push({ product, raw: file.file_base64 || file.fileBase64, filename });
    }
  }
  const approver = assigneeFromRole("Approval", body.assignee);
  const quoter = requireRoleAssignee("Quoting", body.quote_assignee, row.quote_assignee);
  const existing = existingCostSheets(row);
  let n = existing.reduce((m, s) => Math.max(m, Number(s.n) || 0), 0);
  for (const item of uploads) {
    n += 1;
    const kind = "cost_sheet_" + n;
    const saved = db.saveEnquiryAttachment(row.enquiry_no, kind, item.raw, item.filename);
    existing.push({
      n,
      product: item.product,
      kind,
      stored_as: saved.stored_as,
      filename: saved.filename,
      mime: saved.mime,
      uploaded_at: saved.uploaded_at || db.nowIso(),
      uploaded_by: actor,
      size: saved.size
    });
  }
  row.cost_sheets = existing;
  row.cost_sheet = existing[existing.length - 1] || row.cost_sheet;
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
    const coster = requireRoleAssignee("Costing", body.assignee, lastAssignee(row, "cost_sheet"));
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
  const quotePerson = requireRoleAssignee("Quoting", body.quote_assignee || body.assignee, row.quote_assignee);
  row.quote_assignee = quotePerson;
  cancelOpenKind(row, "quote");
  addTask(row, "quote", quotePerson);
}

function lastAssignee(row, kind) {
  const list = (row.tasks || []).filter((t) => t.kind === kind && t.assignee);
  return list.length ? list[list.length - 1].assignee : "";
}

function snapshotQuoteLines(row) {
  return namedProducts(row).map((p) => ({
    product: p.product || "",
    category: p.category || "",
    value_excl_vat: p.value_excl_vat || "",
    value_incl_vat: p.value_incl_vat || ""
  }));
}

function archiveLegacyQuote(row) {
  if (!Array.isArray(row.quotes)) row.quotes = [];
  if (row.quotes.length) return;
  if (!db.enquiryHasQuotePdf(row.enquiry_no)) return;
  recordIssuedQuote(row, "");
}

function recordIssuedQuote(row, actor) {
  if (!Array.isArray(row.quotes)) row.quotes = [];
  const quoteNo = String(row.quote_no || "").trim();
  const existing = quoteNo ? row.quotes.find((q) => String(q.quote_no || "") === quoteNo) : null;
  const pdf = db.readEnquiryQuotePdf(row.enquiry_no);
  let file = existing && existing.file && existing.file.stored_as ? existing.file : null;
  if (!file && pdf && pdf.buffer) {
    const n = existing ? existing.n : row.quotes.length + 1;
    file = db.saveEnquiryAttachment(
      row.enquiry_no,
      "quote_" + n,
      "data:application/pdf;base64," + pdf.buffer.toString("base64"),
      row.quote_pdf_name || pdf.filename || "quote.pdf"
    );
  }
  if (existing) {
    if (file) existing.file = file;
    if (actor && !existing.by) existing.by = actor;
    return;
  }
  row.quotes.push({
    n: row.quotes.length + 1,
    quote_no: quoteNo,
    date_quoted: row.date_quoted || db.todayEnquiryDate(),
    uploaded_at: row.quote_pdf_uploaded_at || db.nowIso(),
    by: actor || "",
    products: snapshotQuoteLines(row),
    delivery_excl_vat: row.delivery_excl_vat || "",
    delivery_incl_vat: row.delivery_incl_vat || "",
    file
  });
}

function completeQuote(row, actor, body) {
  const revision = statusAllows(row, ["Quoted", "Followed Up"]);
  if (!revision && (row.status !== "Costed" || !row.approval || row.approval.status !== "approved")) {
    throw new Error("The cost sheet must be approved before a quote PDF is issued");
  }
  applyPricedBody(row, body);
  requirePricedProducts(row);
  archiveLegacyQuote(row);
  const quoteNo = db.requireUniqueQuoteNo(body.quote_no, revision ? "" : row.enquiry_no);
  const followPerson = requireAssignee(body.follow_up_assignee || body.assignee);
  const payload = {
    ...row,
    quote_no: quoteNo,
    status: "Quoted",
    quote_pdf_base64: body.file_base64 || body.quote_pdf_base64,
    quote_pdf_name: body.file_name || body.quote_pdf_name || "quote.pdf",
    quote_pdf_confirmed: !!(body.file_confirmed || body.quote_pdf_confirmed)
  };
  if (!payload.quote_pdf_base64) throw new Error("Upload the quote PDF, check the preview, then confirm it is the correct file");
  db.upsertEnquiry(payload, { fromPipeline: true });
  const saved = db.getEnquiryRaw(row.enquiry_no);
  Object.assign(row, saved);
  recordIssuedQuote(row, actor);
  closeOpenKind(row, "quote", actor);
  row.follow_up_assignee = followPerson;
  cancelOpenKind(row, "follow_up");
  addTask(row, "follow_up", followPerson, {
    title: "Follow up",
    due_at: addDaysIso(row.date_quoted || db.todayEnquiryDate(), FOLLOW_UP_DAYS)
  });
  if (!openOfKind(row, "pop")) {
    addTask(row, "pop", followPerson, { title: "Record client outcome" });
  }
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
  const next = requireAssignee(body.assignee);
  if (!namesMatch(task.assignee, next)) access.cancelKind(row, task.kind);
  task.assignee = next;
  if (task.kind === "follow_up") row.follow_up_assignee = task.assignee;
  if (task.kind === "quote") row.quote_assignee = task.assignee;
  if (task.kind === "drawing" && row.drawing) row.drawing.assignee = task.assignee;
}

module.exports = {
  FOLLOW_UP_DAYS,
  WAITING_STATUSES,
  CLOSED_STATUSES,
  officeAssignees,
  listMyTasks,
  listMyCompletedTasks,
  listAccessInbox,
  processSnapshot,
  applyAction,
  availableActions,
  canAct,
  actionOwner,
  actionKind
};
