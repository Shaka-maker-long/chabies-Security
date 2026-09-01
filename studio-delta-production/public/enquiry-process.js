(function () {
  const KEEP_VALUES = {
    Costing: 1, Costed: 1, Quoted: 1, "Followed Up": 1, Ordered: 1, "Re-Cost": 1, "Waiting on Supplier": 1
  };
  let state = { open: false, enquiryNo: "", focusTaskId: "", snap: null, file: { base64: "", name: "", url: "", confirmed: false, mime: "" } };

  function esc(s) {
    return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
  }
  function optionList(names, selected) {
    return ["<option value=\"\"></option>"].concat((names || []).map((n) => {
      return "<option" + (n === selected ? " selected" : "") + ">" + esc(n) + "</option>";
    })).join("");
  }
  function assigneeSelect(id, selected) {
    const names = (state.snap && state.snap.assignees) || [];
    return "<select id=\"" + id + "\">" + optionList(names, selected || "") + "</select>";
  }
  function namedLines(row) {
    return ((row && row.products) || []).filter((p) => String(p.product || "").trim());
  }
  function revokeFile() {
    if (state.file.url && String(state.file.url).indexOf("blob:") === 0) URL.revokeObjectURL(state.file.url);
    state.file = { base64: "", name: "", url: "", confirmed: false, mime: "" };
  }
  function fileUrl(kind, download) {
    const no = encodeURIComponent(state.enquiryNo);
    const extra = download ? "?download=1" : "";
    if (kind === "quote") return "/api/office/enquiries/" + no + "/quote.pdf" + extra;
    return "/api/office/enquiries/" + no + "/files/" + encodeURIComponent(kind) + extra;
  }

  function ensureDom() {
    if (document.getElementById("sdProcessMask")) return;
    const wrap = document.createElement("div");
    wrap.id = "sdProcessMask";
    wrap.className = "sd-process-mask";
    wrap.innerHTML =
      "<div class=\"sd-process-sheet\" role=\"dialog\" aria-modal=\"true\">" +
      "<header><div><h1 id=\"sdProcessTitle\">Enquiry process</h1><p id=\"sdProcessSub\" class=\"sd-process-sub\"></p></div>" +
      "<div class=\"sd-process-tools\"><a id=\"sdProcessSheetLink\" href=\"/enquiries\">Enquiries sheet</a>" +
      "<button type=\"button\" class=\"ghost\" id=\"sdProcessClose\">Close</button></div></header>" +
      "<div class=\"sd-process-body\" id=\"sdProcessBody\"></div></div>";
    document.body.appendChild(wrap);
    if (!document.getElementById("sdProcessCss")) {
      const css = document.createElement("style");
      css.id = "sdProcessCss";
      css.textContent =
        ".sd-process-mask{position:fixed;inset:0;background:rgba(16,24,40,.45);z-index:50;display:none;align-items:flex-start;justify-content:center;padding:24px 12px;overflow:auto}" +
        ".sd-process-mask.open{display:flex}" +
        ".sd-process-sheet{width:min(860px,96vw);background:#fff;border-radius:12px;border:1px solid #d0d5dd;margin:12px auto;font-family:Inter,system-ui,sans-serif;color:#1d2939}" +
        ".sd-process-sheet header{display:flex;gap:12px;align-items:flex-start;padding:16px;border-bottom:1px solid #d0d5dd}" +
        ".sd-process-sheet h1{font-size:18px;margin:0;font-family:Outfit,Inter,sans-serif}" +
        ".sd-process-sub{margin:4px 0 0;color:#667085;font-size:13px}" +
        ".sd-process-tools{margin-left:auto;display:flex;gap:8px;align-items:center}" +
        ".sd-process-tools a{font-size:13px;font-weight:600;color:#344054}" +
        ".sd-process-body{padding:16px;display:flex;flex-direction:column;gap:14px}" +
        ".sd-process-card{border:1px solid #d0d5dd;border-radius:10px;padding:12px;background:#f8fafc}" +
        ".sd-process-card h2{margin:0 0 8px;font-size:14px}" +
        ".sd-process-meta{display:flex;gap:12px;flex-wrap:wrap;font-size:13px;color:#475467}" +
        ".sd-process-actions{display:flex;flex-direction:column;gap:10px}" +
        ".sd-process-form label{display:block;font-size:12px;font-weight:600;margin:8px 0 4px}" +
        ".sd-process-form input,.sd-process-form select,.sd-process-form textarea{width:100%;border:1px solid #d0d5dd;border-radius:6px;padding:8px;font:inherit}" +
        ".sd-process-form textarea{min-height:72px}" +
        ".sd-process-preview{width:100%;min-height:220px;max-height:360px;border:1px solid #d0d5dd;border-radius:8px;background:#fff;overflow:auto}" +
        ".sd-process-preview iframe,.sd-process-preview img{width:100%;height:320px;border:0;object-fit:contain}" +
        ".sd-process-preview table{border-collapse:collapse;font-size:11px;width:100%}" +
        ".sd-process-preview th,.sd-process-preview td{border:1px solid #d0d5dd;padding:4px 6px}" +
        ".sd-process-err{color:#b42318;font-size:12px;min-height:16px}" +
        ".sd-process-form button{margin-top:10px}" +
        ".sd-process-sheet button{border:1px solid #1d2939;background:#1d2939;color:#fff;border-radius:6px;padding:8px 12px;font-weight:600;cursor:pointer}" +
        ".sd-process-sheet button.ghost{background:#fff;color:#1d2939}" +
        ".sd-task-pill{display:inline-block;background:#fff;border:1px solid #d0d5dd;border-radius:999px;padding:2px 8px;font-size:12px;margin:0 6px 6px 0}" +
        ".sd-task-pill.overdue{border-color:#fda29b;color:#b42318}" +
        ".sd-lines{width:100%;border-collapse:collapse;font-size:12px;background:#fff}" +
        ".sd-lines th,.sd-lines td{border:1px solid #d0d5dd;padding:6px}";
      document.head.appendChild(css);
    }
    wrap.addEventListener("click", (e) => { if (e.target.id === "sdProcessMask") closeProcess(); });
    document.getElementById("sdProcessClose").onclick = closeProcess;
  }

  function closeProcess() {
    document.getElementById("sdProcessMask").classList.remove("open");
    revokeFile();
    state.open = false;
    if (typeof window.sdOnEnquiryProcessClose === "function") window.sdOnEnquiryProcessClose();
  }

  async function loadXlsx() {
    if (window.XLSX) return window.XLSX;
    await new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";
      s.onload = resolve;
      s.onerror = () => reject(new Error("Could not load spreadsheet preview"));
      document.head.appendChild(s);
    });
    return window.XLSX;
  }

  async function previewFromFile(file) {
    revokeFile();
    if (!file) return;
    state.file.name = file.name;
    state.file.mime = file.type || "";
    state.file.url = URL.createObjectURL(file);
    state.file.base64 = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
    const box = document.getElementById("sdFilePreview");
    if (!box) return;
    const name = file.name.toLowerCase();
    if (file.type.indexOf("pdf") >= 0 || /\.pdf$/.test(name)) {
      box.innerHTML = "<iframe title=\"Preview\"></iframe>";
      box.querySelector("iframe").src = state.file.url;
      return;
    }
    if (file.type.indexOf("image/") === 0 || /\.(png|jpe?g|webp|gif)$/.test(name)) {
      box.innerHTML = "<img alt=\"Preview\">";
      box.querySelector("img").src = state.file.url;
      return;
    }
    if (/\.(xlsx|xls|csv)$/.test(name) || /spreadsheet|csv|excel/.test(file.type)) {
      try {
        const XLSX = await loadXlsx();
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }).slice(0, 24);
        if (!rows.length) {
          box.innerHTML = "<p class=\"sd-process-sub\">The spreadsheet is empty. Download it to check, then confirm.</p>";
          return;
        }
        box.innerHTML = "<table>" + rows.map((r, i) => {
          const cells = (r || []).slice(0, 12).map((c) => (i === 0 ? "<th>" : "<td>") + esc(c) + (i === 0 ? "</th>" : "</td>")).join("");
          return "<tr>" + cells + "</tr>";
        }).join("") + "</table>";
      } catch (e) {
        box.innerHTML = "<p class=\"sd-process-sub\">Could not preview this spreadsheet. Download it, check it, then confirm it is the correct file.</p>";
      }
      return;
    }
    box.innerHTML = "<p class=\"sd-process-sub\">Preview is not available for this file type. Check it, then confirm it is the correct file.</p>";
  }

  async function showSaved(kind) {
    const box = document.getElementById("sdFilePreview");
    if (!box) return;
    const r = await sdOfficeFetch(fileUrl(kind));
    if (!r.ok) return;
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    if ((blob.type || "").indexOf("pdf") >= 0 || kind === "quote") {
      box.innerHTML = "<iframe title=\"Saved file\"></iframe>";
      box.querySelector("iframe").src = url;
    } else if ((blob.type || "").indexOf("image/") === 0) {
      box.innerHTML = "<img alt=\"Saved file\">";
      box.querySelector("img").src = url;
    } else {
      box.innerHTML = "<p class=\"sd-process-sub\">Saved file is attached. <a href=\"" + fileUrl(kind, true) + "\" download>Download</a> to review it.</p>";
    }
  }

  function valuesTable(row, editable) {
    const lines = namedLines(row);
    if (!lines.length) return "<p class=\"sd-process-sub\">Add product names on the Enquiries sheet first.</p>";
    return "<table class=\"sd-lines\"><thead><tr><th>Product</th><th>Value excl VAT</th></tr></thead><tbody>" +
      lines.map((l, i) => "<tr><td>" + esc(l.product) + "</td><td>" +
        (editable
          ? "<input data-val=\"" + i + "\" value=\"" + esc(l.value_excl_vat || "") + "\" inputmode=\"decimal\">"
          : esc(l.value_excl_vat || "—")) +
        "</td></tr>").join("") +
      "</tbody></table>" +
      "<label>Delivery excl VAT *" + (editable
        ? "<input id=\"sdDelivery\" value=\"" + esc(row.delivery_excl_vat || "") + "\" inputmode=\"decimal\">"
        : "</label><div>" + esc(row.delivery_excl_vat || "—") + "</div>") +
      (editable ? "</label>" : "");
  }

  function readValues(row) {
    const lines = namedLines(row).map((l, i) => {
      const input = document.querySelector('input[data-val="' + i + '"]');
      return { product: l.product, category: l.category || "", value_excl_vat: input ? input.value : l.value_excl_vat };
    });
    const delivery = document.getElementById("sdDelivery");
    return { products: lines, delivery_excl_vat: delivery ? delivery.value : row.delivery_excl_vat };
  }

  function fileBlock(accept, hint) {
    return "<label>File<input id=\"sdFileInput\" type=\"file\" accept=\"" + esc(accept) + "\"></label>" +
      "<p class=\"sd-process-sub\">" + esc(hint) + "</p>" +
      "<div class=\"sd-process-preview\" id=\"sdFilePreview\"></div>" +
      "<label><input id=\"sdFileOk\" type=\"checkbox\"> This is the correct file</label>";
  }

  function bindFile() {
    const input = document.getElementById("sdFileInput");
    if (!input) return;
    input.onchange = async (e) => {
      const file = e.target.files && e.target.files[0];
      const ok = document.getElementById("sdFileOk");
      if (ok) ok.checked = false;
      state.file.confirmed = false;
      await previewFromFile(file);
    };
    const ok = document.getElementById("sdFileOk");
    if (ok) ok.onchange = (e) => { state.file.confirmed = !!e.target.checked; };
  }

  function filePayload() {
    return {
      file_base64: state.file.base64,
      file_name: state.file.name,
      file_confirmed: !!(document.getElementById("sdFileOk") && document.getElementById("sdFileOk").checked)
    };
  }

  function formFor(action, row) {
    const waiting = (state.snap.waitingStatuses || []).map((s) => "<option>" + esc(s) + "</option>").join("");
    const closed = (state.snap.closedStatuses || []).map((s) => "<option>" + esc(s) + "</option>").join("");
    if (action.id === "assign_waiting") {
      return "<label>Waiting on<select id=\"sdWaiting\">" + waiting + "</select></label>" +
        "<label>Assign to</label>" + assigneeSelect("sdAssignee");
    }
    if (action.id === "assign_costing" || action.id === "supplier_wait" || action.id === "complete_supplier") {
      return "<label>Assign to</label>" + assigneeSelect("sdAssignee");
    }
    if (action.id === "complete_chase") {
      return "<label>What next?<select id=\"sdNext\"><option value=\"costing\">Enough to cost — assign costing</option><option value=\"waiting\">Still waiting</option></select></label>" +
        "<label>Waiting on<select id=\"sdWaiting\">" + waiting + "</select></label>" +
        "<label>Assign to</label>" + assigneeSelect("sdAssignee") +
        "<label>Comment<textarea id=\"sdComments\"></textarea></label>";
    }
    if (action.id === "complete_cost_sheet") {
      return valuesTable(row, true) + fileBlock(".xlsx,.xls,.csv,application/pdf,.pdf", "Upload the Excel cost sheet. Check the preview, then confirm it.") +
        "<label>Request approval from</label>" + assigneeSelect("sdAssignee");
    }
    if (action.id === "complete_approval") {
      return (row.cost_sheet ? "<p class=\"sd-process-sub\">Cost sheet: " + esc(row.cost_sheet.filename || "cost sheet") + " — <a href=\"" + fileUrl("cost_sheet", true) + "\">download</a></p><div class=\"sd-process-preview\" id=\"sdFilePreview\"></div>" : "") +
        valuesTable(row, false) +
        "<label>Decision<select id=\"sdDecision\"><option value=\"approve\">Approve — send to quote</option><option value=\"reject\">Reject — back to costing</option></select></label>" +
        "<label>Comments (required if rejected)<textarea id=\"sdComments\"></textarea></label>" +
        "<label>Next person (quote person if approved, costing if rejected)</label>" + assigneeSelect("sdAssignee");
    }
    if (action.id === "complete_quote") {
      return valuesTable(row, true) + fileBlock("application/pdf,.pdf", "Upload the quote PDF, check the preview, then confirm it is the correct file. DATE QUOTED is saved with the PDF.") +
        "<label>Who follows up after 7 days?</label>" + assigneeSelect("sdFollow");
    }
    if (action.id === "complete_followup") {
      return fileBlock("image/*,.png,.jpg,.jpeg,.webp,.gif,application/pdf,.pdf", "Upload a screenshot of the follow-up.") +
        "<label>Who owns the next follow-up?</label>" + assigneeSelect("sdAssignee", row.follow_up_assignee);
    }
    if (action.id === "complete_reject") {
      return "<label>Rejection reason *<textarea id=\"sdComments\"></textarea></label>";
    }
    if (action.id === "complete_order") {
      return fileBlock("image/*,.png,.jpg,.jpeg,.webp,.gif,application/pdf,.pdf", "Upload proof of payment (screenshot or PDF).") +
        "<label>Requires drawing?<select id=\"sdDrawing\"><option value=\"\"></option><option value=\"no\">No — ready for Orders</option><option value=\"yes\">Yes — assign drawing</option></select></label>" +
        "<label>Drawing assigned to</label>" + assigneeSelect("sdAssignee");
    }
    if (action.id === "complete_drawing") {
      return fileBlock("application/pdf,.pdf,image/*,.png,.jpg,.jpeg,.webp", "Upload the drawing, check the preview, then confirm it.");
    }
    if (action.id === "close") {
      return "<label>Close as<select id=\"sdCloseStatus\">" + closed + "</select></label>" +
        "<label>Note<textarea id=\"sdComments\"></textarea></label>";
    }
    return "";
  }

  function collect(action, row) {
    const body = { action: action.id };
    const assignee = document.getElementById("sdAssignee");
    const waiting = document.getElementById("sdWaiting");
    const comments = document.getElementById("sdComments");
    if (assignee) body.assignee = assignee.value;
    if (waiting) body.waiting_status = waiting.value;
    if (comments) {
      body.comments = comments.value;
      body.reason = comments.value;
    }
    if (action.id === "complete_chase") body.next = (document.getElementById("sdNext") || {}).value;
    if (action.id === "complete_approval") body.decision = (document.getElementById("sdDecision") || {}).value;
    if (action.id === "complete_quote") body.follow_up_assignee = (document.getElementById("sdFollow") || {}).value;
    if (action.id === "complete_order") {
      const drawing = (document.getElementById("sdDrawing") || {}).value;
      body.drawing_required = drawing;
    }
    if (action.id === "close") body.status = (document.getElementById("sdCloseStatus") || {}).value;
    if (action.id === "complete_cost_sheet" || action.id === "complete_quote") Object.assign(body, readValues(row));
    if (/complete_cost_sheet|complete_quote|complete_followup|complete_order|complete_drawing/.test(action.id)) {
      Object.assign(body, filePayload());
    }
    return body;
  }

  function renderBody() {
    const snap = state.snap;
    const row = snap.row;
    const openTasks = (row.tasks || []).filter((t) => t.status === "open");
    const body = document.getElementById("sdProcessBody");
    document.getElementById("sdProcessTitle").textContent = row.enquiry_no + " · " + (row.client_name || "Enquiry");
    document.getElementById("sdProcessSub").textContent = (row.status || "New") + (row.product ? " · " + row.product : "");
    document.getElementById("sdProcessSheetLink").href = "/enquiries";
    let html = "<div class=\"sd-process-card\"><div class=\"sd-process-meta\">" +
      "<span>Status <b>" + esc(row.status || "New") + "</b></span>" +
      (row.date_quoted ? "<span>Quoted " + esc(row.date_quoted) + "</span>" : "") +
      (row.ready_for_orders ? "<span>Ready for Orders</span>" : "") +
      "</div>";
    if (openTasks.length) {
      html += "<h2>Assigned now</h2>" + openTasks.map((t) => {
        return "<span class=\"sd-task-pill\">" + esc(t.title) + " → " + esc(t.assignee) + "</span>";
      }).join("");
    } else {
      html += "<p class=\"sd-process-sub\">No open assigned task. Capture and assign the next person from here.</p>";
    }
    html += "</div>";
    const actions = snap.actions || [];
    if (!actions.length) {
      html += "<p class=\"sd-process-sub\">This enquiry has no further process steps.</p>";
    } else {
      html += "<div class=\"sd-process-actions\">";
      actions.forEach((action, i) => {
        html += "<form class=\"sd-process-form sd-process-card\" data-action-i=\"" + i + "\">" +
          "<h2>" + esc(action.label) + "</h2>" + formFor(action, row) +
          "<div class=\"sd-process-err\" data-err></div>" +
          "<button type=\"submit\">Save update</button></form>";
      });
      html += "</div>";
    }
    body.innerHTML = html;
    bindFile();
    if (document.getElementById("sdFilePreview") && row.cost_sheet && !document.getElementById("sdFileInput")) {
      showSaved("cost_sheet");
    }
    body.querySelectorAll("form").forEach((form) => {
      form.onsubmit = async (e) => {
        e.preventDefault();
        const i = Number(form.getAttribute("data-action-i"));
        const action = actions[i];
        const err = form.querySelector("[data-err]");
        err.textContent = "";
        const r = await sdOfficeFetch("/api/office/enquiries/" + encodeURIComponent(state.enquiryNo) + "/process", {
          method: "POST",
          body: JSON.stringify(collect(action, row))
        });
        const j = await r.json();
        if (!j.ok) { err.textContent = j.error || "Could not save"; return; }
        state.snap = j;
        revokeFile();
        renderBody();
      };
    });
  }

  window.sdOpenEnquiryProcess = async function (enquiryNo, focusTaskId) {
    ensureDom();
    state.enquiryNo = enquiryNo;
    state.focusTaskId = focusTaskId || "";
    state.open = true;
    revokeFile();
    document.getElementById("sdProcessMask").classList.add("open");
    document.getElementById("sdProcessBody").innerHTML = "<p class=\"sd-process-sub\">Loading…</p>";
    const r = await sdOfficeFetch("/api/office/enquiries/" + encodeURIComponent(enquiryNo) + "/process");
    const j = await r.json();
    if (!j.ok) {
      document.getElementById("sdProcessBody").innerHTML = "<p class=\"sd-process-err\">" + esc(j.error || "Could not load") + "</p>";
      return;
    }
    state.snap = j;
    renderBody();
  };
  window.sdEnquiryKeepValues = KEEP_VALUES;
})();
