function namesMatch(a, b) {
  return String(a || "").trim().toLowerCase() === String(b || "").trim().toLowerCase();
}

function listGrants(row) {
  return Array.isArray(row && row.access_grants) ? row.access_grants : [];
}

function nextGrantId(row) {
  const max = listGrants(row).reduce((m, g) => {
    const n = Number(String((g && g.id) || "").replace(/\D/g, "")) || 0;
    return Math.max(m, n);
  }, 0);
  return "g" + (max + 1);
}

function kindLabel(kind) {
  return ({
    chase_info: "chase missing information",
    cost_sheet: "upload the cost sheet",
    supplier: "upload the supplier quotation",
    approval: "approve or reject costing",
    quote: "upload the quote PDF",
    follow_up: "log a follow-up",
    pop: "attach proof of payment or the client outcome",
    drawing: "upload the drawing"
  })[kind] || String(kind || "this step");
}

function grantedFor(row, actor, kind) {
  if (!actor || !kind) return false;
  return listGrants(row).some((g) => {
    return g && g.status === "granted"
      && g.kind === kind
      && namesMatch(g.requester, actor);
  });
}

function pendingFor(row, actor, kind) {
  if (!actor || !kind) return null;
  return listGrants(row).find((g) => {
    return g && g.status === "pending"
      && g.kind === kind
      && namesMatch(g.requester, actor);
  }) || null;
}

function cancelKind(row, kind) {
  if (!row || !kind) return;
  listGrants(row).forEach((g) => {
    if (g && g.kind === kind && (g.status === "pending" || g.status === "granted")) {
      g.status = "cancelled";
    }
  });
}

function consumeKind(row, kind) {
  if (!row || !kind) return;
  listGrants(row).forEach((g) => {
    if (g && g.kind === kind && g.status === "granted") g.status = "used";
    if (g && g.kind === kind && g.status === "pending") g.status = "cancelled";
  });
}

function requestAccess(row, actor, kind, owner) {
  const requester = String(actor || "").trim();
  const assignee = String(owner || "").trim();
  if (!requester) throw new Error("Not signed in");
  if (!kind) throw new Error("Choose which step you need access to");
  if (!assignee) throw new Error("This step is not assigned to anyone yet");
  if (namesMatch(requester, assignee)) {
    throw new Error("This step is already assigned to you");
  }
  if (grantedFor(row, requester, kind)) {
    throw new Error("Access is already granted for this step");
  }
  const existing = pendingFor(row, requester, kind);
  if (existing) return existing;
  if (!Array.isArray(row.access_grants)) row.access_grants = [];
  const grant = {
    id: nextGrantId(row),
    enquiry_no: row.enquiry_no || "",
    kind,
    requester,
    assignee,
    requested_at: new Date().toISOString(),
    status: "pending",
    granted_by: "",
    granted_at: "",
    note: ""
  };
  row.access_grants.push(grant);
  return grant;
}

function canDecide(grant, actor, isManager) {
  if (!grant) return false;
  if (isManager) return true;
  return namesMatch(grant.assignee, actor);
}

function grantAccess(row, actor, grantId, isManager) {
  const grant = listGrants(row).find((g) => g && g.id === String(grantId || "").trim());
  if (!grant) throw new Error("Access request not found");
  if (grant.status !== "pending") throw new Error("That request is no longer waiting");
  if (!canDecide(grant, actor, isManager)) {
    throw new Error("Only " + (grant.assignee || "the assigned person") + " or the Manager can grant this");
  }
  grant.status = "granted";
  grant.granted_by = String(actor || "").trim();
  grant.granted_at = new Date().toISOString();
  return grant;
}

function denyAccess(row, actor, grantId, isManager) {
  const grant = listGrants(row).find((g) => g && g.id === String(grantId || "").trim());
  if (!grant) throw new Error("Access request not found");
  if (grant.status !== "pending") throw new Error("That request is no longer waiting");
  if (!canDecide(grant, actor, isManager)) {
    throw new Error("Only " + (grant.assignee || "the assigned person") + " or the Manager can refuse this");
  }
  grant.status = "denied";
  grant.granted_by = String(actor || "").trim();
  grant.granted_at = new Date().toISOString();
  return grant;
}

function snapshotFor(row, actor, isManager) {
  const me = String(actor || "").trim();
  const grants = listGrants(row);
  return {
    pendingForMe: grants.filter((g) => g && g.status === "pending" && (isManager || namesMatch(g.assignee, me))),
    mine: grants.filter((g) => g && namesMatch(g.requester, me) && (g.status === "pending" || g.status === "granted")),
    kindLabel
  };
}

function inboxFor(enquiries, actor, isManager) {
  const me = String(actor || "").trim();
  const out = [];
  (enquiries || []).forEach((row) => {
    listGrants(row).forEach((g) => {
      if (!g || g.status !== "pending") return;
      if (!isManager && !namesMatch(g.assignee, me)) return;
      out.push({
        ...g,
        enquiry_no: row.enquiry_no,
        client_name: row.client_name || "",
        enquiry_status: row.status || "",
        kind_label: kindLabel(g.kind)
      });
    });
  });
  out.sort((a, b) => String(b.requested_at || "").localeCompare(String(a.requested_at || "")));
  return out;
}

module.exports = {
  namesMatch,
  listGrants,
  kindLabel,
  grantedFor,
  pendingFor,
  cancelKind,
  consumeKind,
  requestAccess,
  grantAccess,
  denyAccess,
  snapshotFor,
  inboxFor
};
