(function () {
  const KEEP_VALUES = { Quoted: 1, "Followed Up": 1, Ordered: 1 };
  let state = { open: false, enquiryNo: "", focusTaskId: "", snap: null, file: { base64: "", name: "", url: "", confirmed: false, mime: "" } };

  function esc(s) {
    return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
  }
  function optionList(names, selected) {
    return ["<option value=\"\"></option>"].concat((names || []).map((n) => {
      return "<option" + (n === selected ? " selected" : "") + ">" + esc(n) + "</option>";
    })).join("");
  }
  function assigneeSelect(selected, fieldName) {
    const names = (state.snap && state.snap.assignees) || [];
    const me = (state.snap && state.snap.me) || "";
    const current = String(selected || "");
    const field = fieldName || "assignee";
    let html = "<select name=\"" + esc(field) + "\" class=\"sd-assignee\">";
    html += "<option value=\"\"></option>";
    names.forEach((n) => {
      const label = n === me ? n + " (you)" : n;
      html += "<option value=\"" + esc(n) + "\"" + (n === current ? " selected" : "") + ">" + esc(label) + "</option>";
    });
    return html + "</select>";
  }
  function openAssignee(row, kind) {
    const t = ((row && row.tasks) || []).find((task) => task.kind === kind && task.status === "open");
    return (t && t.assignee) || "";
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
        ".sd-assignee{max-width:280px;width:100%}" +
        ".sd-quote-no{max-width:180px;width:100%;border:1px solid #d0d5dd;border-radius:6px;padding:8px;font:inherit}" +
        ".sd-path{width:100%;border:1px solid #d0d5dd;border-radius:6px;padding:8px;font:inherit;font-size:12px}" +
        ".sd-correspondence{border:1px dashed #98a2b3;border-radius:10px;padding:10px 12px;background:#fff;margin-top:8px}" +
        ".sd-correspondence h2{margin:0 0 6px;font-size:13px}" +
        ".sd-mail{width:100%;border-collapse:collapse;font-size:12px;background:#fff;margin-top:8px}" +
        ".sd-mail th,.sd-mail td{border:1px solid #d0d5dd;padding:6px 8px;text-align:left}" +
        ".sd-mail button,.sd-mail a.sd-open-mail{margin:0;display:inline-block;text-decoration:none}" +
        ".sd-path-row{display:flex;gap:8px;align-items:center;flex-wrap:wrap}" +
        ".sd-path-row code{font-size:12px;word-break:break-all;flex:1;min-width:160px}" +
        ".sd-process-form textarea{min-height:72px}" +
        ".sd-process-preview{width:100%;min-height:220px;max-height:360px;border:1px solid #d0d5dd;border-radius:8px;background:#fff;overflow:auto}" +
        ".sd-process-preview iframe,.sd-process-preview img{width:100%;height:320px;border:0;object-fit:contain}" +
        ".sd-process-preview table{border-collapse:collapse;font-size:11px;width:100%}" +
        ".sd-process-preview th,.sd-process-preview td{border:1px solid #d0d5dd;padding:4px 6px}" +
        ".sd-process-err{color:#b42318;font-size:12px;min-height:16px}" +
        ".sd-process-form button{margin-top:10px}" +
        ".sd-process-card summary{cursor:pointer;list-style:none;display:flex;align-items:center;gap:8px}" +
        ".sd-process-card summary h2{margin:0}" +
        ".sd-process-card summary::-webkit-details-marker{display:none}" +
        ".sd-process-card:not([open]) .sd-process-form{display:none}" +
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

  async function previewFromFile(file, box) {
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

  async function showSaved(kind, box) {
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

  function productNamesLine(row) {
    const names = namedLines(row).map((l) => l.product);
    if (!names.length) return "<p class=\"sd-process-sub\">Add product names on the Enquiries sheet first.</p>";
    return "<p class=\"sd-process-sub\">Products: " + esc(names.join(", ")) + "</p>";
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
        ? "<input name=\"delivery_excl_vat\" value=\"" + esc(row.delivery_excl_vat || "") + "\" inputmode=\"decimal\">"
        : "</label><div>" + esc(row.delivery_excl_vat || "—") + "</div>") +
      (editable ? "</label>" : "");
  }

  function readValues(form, row) {
    const lines = namedLines(row).map((l, i) => {
      const input = form.querySelector('input[data-val="' + i + '"]');
      return { product: l.product, category: l.category || "", value_excl_vat: input ? input.value : l.value_excl_vat };
    });
    const delivery = form.querySelector('[name="delivery_excl_vat"]');
    return { products: lines, delivery_excl_vat: delivery ? delivery.value : row.delivery_excl_vat };
  }

  function fileBlock(accept, hint) {
    return "<label>File<input name=\"file\" type=\"file\" accept=\"" + esc(accept) + "\"></label>" +
      "<p class=\"sd-process-sub\">" + esc(hint) + "</p>" +
      "<div class=\"sd-process-preview\" data-preview></div>" +
      "<label><input name=\"file_ok\" type=\"checkbox\"> This is the correct file</label>";
  }

  function bindFile(form) {
    const input = form && form.querySelector('[name="file"]');
    if (!input) return;
    input.onchange = async (e) => {
      const file = e.target.files && e.target.files[0];
      const ok = form.querySelector('[name="file_ok"]');
      if (ok) ok.checked = false;
      state.file.confirmed = false;
      await previewFromFile(file, form.querySelector("[data-preview]"));
    };
    const ok = form.querySelector('[name="file_ok"]');
    if (ok) ok.onchange = (e) => { state.file.confirmed = !!e.target.checked; };
  }

  function filePayload(form) {
    return {
      file_base64: state.file.base64,
      file_name: state.file.name,
      file_confirmed: !!(form.querySelector('[name="file_ok"]') && form.querySelector('[name="file_ok"]').checked)
    };
  }

  function outlookHref(mail) {
    return String((mail && mail.outlook_url) || "").trim();
  }
  function correspondenceFields() {
    const origin = (location && location.origin) || "";
    return "<p class=\"sd-process-sub\">Home in Outlook will not show a Studio Delta button. You attach the email from here. Do not save a .msg file.</p>" +
      "<p class=\"sd-process-sub\">In Outlook, click the email once in Inbox or Sent. Right-click it → <b>Copy as link</b>. If you do not see that, click the <b>…</b> on the far right of Home, or open the email and look on the Message tab. Paste the link below.</p>" +
      "<label>Outlook link<textarea class=\"sd-path\" name=\"correspondence_links\" placeholder=\"Paste the Outlook link here. One email per line.\"></textarea></label>" +
      "<p class=\"sd-process-sub\">Optional, only if your PC allows add-ins: <a href=\"" + origin + "/outlook-addin\" target=\"_blank\" rel=\"noopener\">File → Get Add-ins</a>.</p>";
  }
  function correspondenceCard(row) {
    const c = (row && row.correspondence) || {};
    const mails = c.mails || [];
    if (!mails.length) return "";
    let html = "<div class=\"sd-correspondence\"><h2>CORRESPONDANCE</h2>" +
      "<p class=\"sd-process-sub\">These emails stay in Outlook. Open in Outlook launches the desktop app — nothing is downloaded.</p>" +
      "<table class=\"sd-mail\"><thead><tr><th>Email</th><th>From</th><th>Order</th><th></th></tr></thead><tbody>";
    html += mails.map((f) => {
      const href = outlookHref(f);
      return "<tr>" +
        "<td>" + esc(f.title || "Outlook email") + "</td>" +
        "<td>" + esc(f.from || f.from_email || "") + "</td>" +
        "<td>" + esc(f.order_no || "") + "</td>" +
        "<td>" + (href
          ? "<a class=\"sd-open-mail\" href=\"" + esc(href) + "\">Open in Outlook</a>"
          : "") + "</td>" +
        "</tr>";
    }).join("");
    return html + "</tbody></table></div>";
  }
  function formFor(action, row) {
    const waiting = (state.snap.waitingStatuses || []).map((s) => "<option>" + esc(s) + "</option>").join("");
    const closed = (state.snap.closedStatuses || []).map((s) => "<option>" + esc(s) + "</option>").join("");
    if (action.id === "assign_waiting") {
      return "<label>Waiting on<select name=\"waiting_status\">" + waiting + "</select></label>" +
        "<label>Assign to</label>" + assigneeSelect();
    }
    if (action.id === "assign_costing") {
      return "<p class=\"sd-process-sub\">Any office Admin, including yourself.</p><label>Assign to</label>" + assigneeSelect(openAssignee(row, "cost_sheet")) +
        correspondenceFields();
    }
    if (action.id === "add_correspondence") {
      return correspondenceFields();
    }
    if (action.id === "supplier_wait" || action.id === "complete_supplier") {
      return "<label>Assign to</label>" + assigneeSelect(openAssignee(row, action.id === "supplier_wait" ? "cost_sheet" : "supplier"));
    }
    if (action.id === "complete_chase") {
      return "<label>What next?<select name=\"next\"><option value=\"costing\">Enough to cost — assign costing</option><option value=\"waiting\">Still waiting</option></select></label>" +
        "<label>Waiting on<select name=\"waiting_status\">" + waiting + "</select></label>" +
        "<label>Assign to</label>" + assigneeSelect() +
        "<label>Comment<textarea name=\"comments\"></textarea></label>" +
        correspondenceFields();
    }
    if (action.id === "complete_cost_sheet") {
      return productNamesLine(row) + fileBlock(".xlsx,.xls,.csv,application/pdf,.pdf", "Upload the Excel cost sheet. Check the preview, then confirm it.") +
        "<label>Request approval from (optional)</label>" + assigneeSelect("", "assignee") +
        "<p class=\"sd-process-sub\">Leave this empty to skip approval and send the enquiry to the quoting person.</p>" +
        "<label>Quoting person *</label>" + assigneeSelect(row.quote_assignee || "", "quote_assignee");
    }
    if (action.id === "complete_approval") {
      return (row.cost_sheet ? "<p class=\"sd-process-sub\">Cost sheet: " + esc(row.cost_sheet.filename || "cost sheet") + " — <a href=\"" + fileUrl("cost_sheet", true) + "\">download</a></p><div class=\"sd-process-preview\" data-preview></div>" : "") +
        productNamesLine(row) +
        "<label>Decision<select name=\"decision\"><option value=\"approve\">Approve — send to quote</option><option value=\"reject\">Reject — back to costing</option></select></label>" +
        "<label>Comments (required if rejected)<textarea name=\"comments\"></textarea></label>" +
        "<label>Next person (quote person if approved, costing if rejected)</label>" +
        assigneeSelect(row.quote_assignee || openAssignee(row, "quote"));
    }
    if (action.id === "complete_quote") {
      const hint = (state.snap && state.snap.quoteNo) || {};
      const recent = (hint.recent || []).slice();
      const next = row.quote_no || hint.next || "";
      const recentLine = recent.length
        ? "Last quotation numbers: " + recent.join(", ") + "."
        : "No quotation numbers yet.";
      return valuesTable(row, true) +
        "<label>Quotation number *<input class=\"sd-quote-no\" name=\"quote_no\" value=\"" + esc(next) + "\" autocomplete=\"off\"></label>" +
        "<p class=\"sd-process-sub\">" + esc(recentLine) + " Default is the next number (" + esc(hint.next || next) + "). You can change it, but it cannot match an existing quotation.</p>" +
        fileBlock("application/pdf,.pdf", "Enter each product value and delivery excluding VAT here, then upload the quote PDF. DATE QUOTED is saved with the PDF.") +
        "<label>Who follows up after 7 days?</label>" + assigneeSelect();
    }
    if (action.id === "complete_followup") {
      return fileBlock("image/*,.png,.jpg,.jpeg,.webp,.gif,application/pdf,.pdf", "Upload a screenshot of the follow-up.") +
        "<label>Who owns the next follow-up?</label>" + assigneeSelect(row.follow_up_assignee);
    }
    if (action.id === "complete_reject") {
      return "<label>Rejection reason *<textarea name=\"comments\"></textarea></label>";
    }
    if (action.id === "complete_order") {
      return fileBlock("image/*,.png,.jpg,.jpeg,.webp,.gif,application/pdf,.pdf", "Upload proof of payment (screenshot or PDF).") +
        "<label>Requires drawing?<select name=\"drawing_required\"><option value=\"\"></option><option value=\"no\">No — ready for Orders</option><option value=\"yes\">Yes — assign drawing</option></select></label>" +
        "<label>Drawing assigned to</label>" + assigneeSelect();
    }
    if (action.id === "complete_drawing") {
      return fileBlock("application/pdf,.pdf,image/*,.png,.jpg,.jpeg,.webp", "Upload the drawing, check the preview, then confirm it.");
    }
    if (action.id === "close") {
      return "<label>Close as<select name=\"status\">" + closed + "</select></label>" +
        "<label>Note<textarea name=\"comments\"></textarea></label>";
    }
    return "";
  }

  function field(form, name) {
    const el = form.querySelector('[name="' + name + '"]');
    return el ? String(el.value || "").trim() : "";
  }

  function collect(form, action, row) {
    const body = { action: action.id };
    body.assignee = field(form, "assignee");
    body.quote_assignee = field(form, "quote_assignee");
    body.waiting_status = field(form, "waiting_status");
    body.comments = field(form, "comments");
    body.reason = body.comments;
    if (action.id === "complete_chase") body.next = field(form, "next");
    if (action.id === "complete_approval") body.decision = field(form, "decision");
    if (action.id === "complete_quote") body.follow_up_assignee = field(form, "assignee");
    if (action.id === "complete_quote") body.quote_no = field(form, "quote_no");
    if (action.id === "assign_costing" || action.id === "complete_chase" || action.id === "add_correspondence") {
      body.correspondence_links = field(form, "correspondence_links");
    }
    if (action.id === "complete_order") body.drawing_required = field(form, "drawing_required");
    if (action.id === "close") body.status = field(form, "status");
    if (action.id === "complete_quote") Object.assign(body, readValues(form, row));
    if (/complete_cost_sheet|complete_quote|complete_followup|complete_order|complete_drawing/.test(action.id)) {
      Object.assign(body, filePayload(form));
    }
    return body;
  }

  function actionForTaskKind(kind) {
    return {
      chase_info: "complete_chase",
      cost_sheet: "complete_cost_sheet",
      supplier: "complete_supplier",
      approval: "complete_approval",
      quote: "complete_quote",
      follow_up: "complete_followup",
      pop: "complete_order",
      drawing: "complete_drawing"
    }[kind] || "";
  }

  function shouldExpandAction(action, i, row, actions) {
    const focus = state.focusTaskId;
    if (focus) {
      const task = (row.tasks || []).find((t) => t.id === focus);
      if (task) return action.id === actionForTaskKind(task.kind);
      return action.id === focus;
    }
    if (actions.length === 1) return true;
    if (actions.some((a) => a.id === "assign_costing")) return action.id === "assign_costing";
    return i === 0;
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
      (row.date_quoted ? "<span>Quoted " + esc(row.date_quoted) + (row.quote_no ? " · " + esc(row.quote_no) : "") + "</span>" : "") +
      (row.ready_for_orders ? "<span>Ready for Orders</span>" : "") +
      "</div>" + correspondenceCard(row);
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
        const open = shouldExpandAction(action, i, row, actions);
        html += "<details class=\"sd-process-card\" data-action-i=\"" + i + "\"" + (open ? " open" : "") + ">" +
          "<summary><h2>" + esc(action.label) + "</h2></summary>" +
          "<form class=\"sd-process-form\">" + formFor(action, row) +
          "<div class=\"sd-process-err\" data-err></div>" +
          "<button type=\"submit\">Save update</button></form></details>";
      });
      html += "</div>";
    }
    body.innerHTML = html;
    body.querySelectorAll("form").forEach((form) => {
      bindFile(form);
      const card = form.closest("[data-action-i]");
      const i = Number((card || form).getAttribute("data-action-i"));
      const preview = form.querySelector("[data-preview]");
      if (preview && row.cost_sheet && !form.querySelector('[name="file"]')) showSaved("cost_sheet", preview);
      form.onsubmit = async (e) => {
        e.preventDefault();
        const action = actions[i];
        const err = form.querySelector("[data-err]");
        err.textContent = "";
        const r = await sdOfficeFetch("/api/office/enquiries/" + encodeURIComponent(state.enquiryNo) + "/process", {
          method: "POST",
          body: JSON.stringify(collect(form, action, row))
        });
        const j = await r.json();
        if (!j.ok) { err.textContent = j.error || "Could not save"; return; }
        state.snap = j;
        revokeFile();
        renderBody();
      };
    });
    body.querySelectorAll("a.sd-open-mail").forEach((a) => {
      a.onclick = (e) => {
        const href = a.getAttribute("href") || "";
        if (/^(outlook:|ms-outlook:)/i.test(href)) {
          e.preventDefault();
          window.location.href = href;
        }
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
