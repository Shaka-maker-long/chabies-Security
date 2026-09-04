(function () {
  const KEEP_VALUES = { Quoted: 1, "Followed Up": 1, Ordered: 1 };
  let state = { open: false, enquiryNo: "", focusTaskId: "", snap: null, file: { base64: "", name: "", url: "", confirmed: false, mime: "" }, costFiles: {} };

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
  function rolePerson(kind) {
    const roles = (state.snap && state.snap.enquiryRoles) || {};
    return String(roles[kind] || "").trim();
  }
  function openAssignee(row, kind) {
    const t = ((row && row.tasks) || []).find((task) => task.kind === kind && task.status === "open");
    return (t && t.assignee) || "";
  }
  function lastCosting(row) {
    const list = ((row && row.tasks) || []).filter((t) => t.kind === "cost_sheet" && t.assignee);
    return list.length ? list[list.length - 1].assignee : "";
  }
  function namedLines(row) {
    return ((row && row.products) || []).filter((p) => String(p.product || "").trim());
  }
  function emptyFileSlot() {
    return { base64: "", name: "", url: "", confirmed: false, mime: "" };
  }
  function revokeSlot(slot) {
    if (slot && slot.url && String(slot.url).indexOf("blob:") === 0) {
      URL.revokeObjectURL(slot.url);
      slot.url = "";
    }
  }
  function revokeAllCostFiles() {
    const map = state.costFiles || {};
    Object.keys(map).forEach((product) => {
      (map[product] || []).forEach(revokeSlot);
    });
    state.costFiles = {};
  }
  function revokeFile() {
    revokeSlot(state.file);
    state.file = emptyFileSlot();
    revokeAllCostFiles();
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
        ".sd-process-sheet{width:min(860px,96vw);max-width:100%;background:#fff;border-radius:12px;border:1px solid #d0d5dd;margin:12px auto;font-family:Inter,system-ui,sans-serif;color:#1d2939;overflow:hidden}" +
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
        ".sd-drop{position:relative;border:1px dashed #98a2b3;border-radius:8px;padding:18px 14px;background:#fff;text-align:center;font-size:13px;color:#475467;margin:8px 0;min-height:72px}" +
        ".sd-drop.over{border-color:#1d2939;background:#eef2f6}" +
        ".sd-drop input[type=file]{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%;font-size:0}" +
        ".sd-drop span{pointer-events:none;display:block}" +
        ".sd-correspondence h2,.sd-files h2{margin:0 0 6px;font-size:13px}" +
        ".sd-process-sheet table{min-width:0;width:100%;max-width:100%;table-layout:fixed}" +
        ".sd-file-list{display:flex;flex-direction:column;gap:8px;margin-top:8px}" +
        ".sd-file-row{display:flex;gap:10px;align-items:center;justify-content:space-between;flex-wrap:wrap;background:#fff;border:1px solid #d0d5dd;border-radius:8px;padding:8px 10px}" +
        ".sd-file-row .sd-file-meta{min-width:0;flex:1}" +
        ".sd-file-row .sd-file-type{font-size:11px;font-weight:600;letter-spacing:.02em;color:#667085}" +
        ".sd-file-row .sd-file-name{font-size:13px;word-break:break-word}" +
        ".sd-file-row button{margin:0;flex:0 0 auto}" +
        ".sd-mail{width:100%;border-collapse:collapse;font-size:12px;background:#fff;margin-top:8px}" +
        ".sd-mail th,.sd-mail td{border:1px solid #d0d5dd;padding:6px 8px;text-align:left;white-space:normal;word-break:break-word}" +
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
        ".sd-lines th,.sd-lines td{border:1px solid #d0d5dd;padding:6px}" +
        ".sd-quote-totals{display:flex;flex-wrap:wrap;gap:10px 18px;margin-top:10px;padding:10px 12px;background:#fff;border:1px solid #d0d5dd;border-radius:8px;font-size:13px}" +
        ".sd-quote-totals b{font-family:Outfit,Inter,sans-serif}" +
        ".sd-quote-totals .sd-quote-total{font-weight:700}" +
        ".sd-quote-totals .sd-quote-total b{font-size:15px}" +
        ".sd-quote-totals .sd-process-sub{flex:1 1 100%;margin:0}" +
        ".sd-life{font-size:13px;color:#1d2939}" +
        ".sd-timeline{list-style:none;margin:8px 0 0;padding:0;border-left:2px solid #d0d5dd}" +
        ".sd-timeline li{position:relative;padding:0 0 12px 16px;font-size:13px}" +
        ".sd-timeline li:last-child{padding-bottom:0}" +
        ".sd-timeline li::before{content:\"\";position:absolute;left:-5px;top:6px;width:8px;height:8px;border-radius:50%;background:#1d2939}" +
        ".sd-timeline time{display:block;font-size:11px;font-weight:600;color:#667085;letter-spacing:.02em}" +
        ".sd-timeline .sd-tl-actor{color:#667085;font-size:12px}" +
        ".sd-product-cost{background:#fff;border:1px solid #d0d5dd;border-radius:8px;padding:10px 12px;margin:10px 0}" +
        ".sd-cost-product{margin:0 0 8px;font-size:14px}" +
        ".sd-cost-slot{margin:0 0 12px;padding:0 0 12px;border-bottom:1px solid #eaecf0}" +
        ".sd-cost-slot:last-child{border-bottom:0;margin-bottom:0;padding-bottom:0}" +
        ".sd-locked{background:#fff7ed;border-color:#fdc5a3}" +
        ".sd-grant-row{display:flex;gap:8px;align-items:center;flex-wrap:wrap;background:#fff;border:1px solid #d0d5dd;border-radius:8px;padding:8px 10px;margin:6px 0}" +
        ".sd-create-order{margin-top:10px}";
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

  async function fillSlotFromFile(slot, file, box) {
    revokeSlot(slot);
    slot.base64 = "";
    slot.name = "";
    slot.confirmed = false;
    slot.mime = "";
    slot.url = "";
    if (!file) {
      if (box) box.innerHTML = "";
      return;
    }
    slot.name = file.name;
    slot.mime = file.type || "";
    slot.url = URL.createObjectURL(file);
    slot.base64 = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
    if (!box) return;
    const name = file.name.toLowerCase();
    if (file.type.indexOf("pdf") >= 0 || /\.pdf$/.test(name)) {
      box.innerHTML = "<iframe title=\"Preview\"></iframe>";
      box.querySelector("iframe").src = slot.url;
      return;
    }
    if (file.type.indexOf("image/") === 0 || /\.(png|jpe?g|webp|gif)$/.test(name)) {
      box.innerHTML = "<img alt=\"Preview\">";
      box.querySelector("img").src = slot.url;
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

  async function previewFromFile(file, box) {
    await fillSlotFromFile(state.file, file, box);
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

  function parseMoney(s) {
    const n = Number(String(s || "").replace(/,/g, "").replace(/[^0-9.-]/g, ""));
    return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0;
  }
  function money(n) {
    return (Math.round(Number(n || 0) * 100) / 100).toFixed(2);
  }
  function formatRand(n) {
    const v = parseMoney(n);
    const neg = v < 0 ? "-" : "";
    const [whole, frac] = money(Math.abs(v)).split(".");
    return neg + "R " + whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + "." + frac;
  }
  function exclFromIncl(raw) {
    if (String(raw || "").trim() === "") return "";
    return money(parseMoney(raw) / 1.15);
  }
  function inclFromExcl(raw) {
    if (String(raw || "").trim() === "") return "";
    return money(parseMoney(raw) * 1.15);
  }
  function displayIncl(line) {
    const incl = String((line && line.value_incl_vat) || "").trim();
    if (incl) return incl;
    return inclFromExcl(line && line.value_excl_vat);
  }
  function displayDeliveryIncl(row) {
    const incl = String((row && row.delivery_incl_vat) || "").trim();
    if (incl) return incl;
    return inclFromExcl(row && row.delivery_excl_vat);
  }

  function productNamesLine(row) {
    const names = namedLines(row).map((l) => l.product);
    if (!names.length) return "<p class=\"sd-process-sub\">Add product names on the Enquiries sheet first.</p>";
    return "<p class=\"sd-process-sub\">Products: " + esc(names.join(", ")) + "</p>";
  }

  function ensureCostSlots(row) {
    state.costFiles = state.costFiles || {};
    namedLines(row).forEach((l) => {
      const product = l.product;
      if (!state.costFiles[product] || !state.costFiles[product].length) {
        state.costFiles[product] = [emptyFileSlot()];
      }
    });
  }

  function costFileSlotHtml(product, fi) {
    const slot = ((state.costFiles[product] || [])[fi]) || emptyFileSlot();
    return "<div class=\"sd-cost-slot\" data-product=\"" + esc(product) + "\" data-slot=\"" + fi + "\">" +
      "<label>Cost sheet<input class=\"cost-file-input\" type=\"file\" accept=\".xlsx,.xls,.csv,application/pdf,.pdf\" data-product=\"" + esc(product) + "\" data-slot=\"" + fi + "\"></label>" +
      (slot.name ? "<p class=\"sd-process-sub\">Selected: " + esc(slot.name) + "</p>" : "") +
      "<p class=\"sd-process-sub\">Excel, CSV, or PDF. Check the preview, then confirm it.</p>" +
      "<div class=\"sd-process-preview\" data-cost-preview></div>" +
      "<label><input class=\"cost-file-ok\" type=\"checkbox\" data-product=\"" + esc(product) + "\" data-slot=\"" + fi + "\"" + (slot.confirmed ? " checked" : "") + "> This is the correct file</label>" +
      (fi > 0 ? "<button type=\"button\" class=\"ghost remove-cost-file\" data-product=\"" + esc(product) + "\" data-slot=\"" + fi + "\">Remove this sheet</button>" : "") +
      "</div>";
  }

  function existingCostSheetNote(row) {
    const sheets = Array.isArray(row.cost_sheets) && row.cost_sheets.length
      ? row.cost_sheets
      : (row.cost_sheet && row.cost_sheet.stored_as ? [row.cost_sheet] : []);
    if (!sheets.length) {
      return "<p class=\"sd-process-sub\">Upload at least one cost sheet for each product. You can add more than one sheet per item.</p>";
    }
    const list = sheets.map((s) => {
      return esc((s.product ? s.product + " · " : "") + (s.filename || "cost sheet"));
    }).join("; ");
    return "<p class=\"sd-process-sub\">Already on file: " + list + ". Those stay in Files. Upload at least one sheet for each product this time — add extra sheets per item if you need them.</p>";
  }

  function productCostBlocks(row) {
    const named = namedLines(row);
    if (!named.length) return "<p class=\"sd-process-sub\">Add product names on the Enquiries sheet first.</p>";
    ensureCostSlots(row);
    return existingCostSheetNote(row) + named.map((l) => {
      const product = l.product;
      const slots = state.costFiles[product] || [emptyFileSlot()];
      return "<section class=\"sd-product-cost\" data-product=\"" + esc(product) + "\">" +
        "<h3 class=\"sd-cost-product\">" + esc(product) + "</h3>" +
        "<div class=\"cost-slots\">" + slots.map((_, fi) => costFileSlotHtml(product, fi)).join("") + "</div>" +
        "<button type=\"button\" class=\"ghost add-cost-file\" data-product=\"" + esc(product) + "\">Add another cost sheet</button>" +
        "</section>";
    }).join("");
  }

  function costSheetLinks(row) {
    const sheets = Array.isArray(row.cost_sheets) && row.cost_sheets.length
      ? row.cost_sheets
      : (row.cost_sheet && row.cost_sheet.stored_as ? [row.cost_sheet] : []);
    if (!sheets.length) return "";
    return sheets.map((s) => {
      const kind = s.kind || "cost_sheet";
      const label = s.product ? (s.filename || "cost sheet") + " · " + s.product : (s.filename || "cost sheet");
      return "<p class=\"sd-process-sub\">Cost sheet: " + esc(label) + " — <a href=\"" + fileUrl(kind, true) + "\">download</a></p>";
    }).join("") + "<div class=\"sd-process-preview\" data-preview></div>";
  }

  function bindCostSheets(form) {
    if (!form || !form.querySelector(".sd-product-cost")) return;
    form.querySelectorAll(".cost-file-input").forEach((input) => {
      if (input.dataset.bound) return;
      input.dataset.bound = "1";
      input.onchange = async () => {
        const product = input.getAttribute("data-product") || "";
        const fi = Number(input.getAttribute("data-slot") || 0);
        const slots = state.costFiles[product] || (state.costFiles[product] = []);
        while (slots.length <= fi) slots.push(emptyFileSlot());
        const slot = slots[fi];
        slot.confirmed = false;
        const wrap = input.closest(".sd-cost-slot");
        const ok = wrap && wrap.querySelector(".cost-file-ok");
        if (ok) ok.checked = false;
        const box = wrap && wrap.querySelector("[data-cost-preview]");
        await fillSlotFromFile(slot, input.files && input.files[0], box);
      };
    });
    form.querySelectorAll(".cost-file-ok").forEach((ok) => {
      if (ok.dataset.bound) return;
      ok.dataset.bound = "1";
      ok.onchange = () => {
        const product = ok.getAttribute("data-product") || "";
        const fi = Number(ok.getAttribute("data-slot") || 0);
        const slot = ((state.costFiles[product] || [])[fi]);
        if (slot) slot.confirmed = !!ok.checked;
      };
    });
    form.querySelectorAll(".add-cost-file").forEach((btn) => {
      if (btn.dataset.bound) return;
      btn.dataset.bound = "1";
      btn.onclick = (e) => {
        e.preventDefault();
        const product = btn.getAttribute("data-product") || "";
        if (!state.costFiles[product]) state.costFiles[product] = [emptyFileSlot()];
        state.costFiles[product].push(emptyFileSlot());
        const fi = state.costFiles[product].length - 1;
        const wrap = btn.closest(".sd-product-cost").querySelector(".cost-slots");
        wrap.insertAdjacentHTML("beforeend", costFileSlotHtml(product, fi));
        bindCostSheets(form);
      };
    });
    form.querySelectorAll(".remove-cost-file").forEach((btn) => {
      if (btn.dataset.bound) return;
      btn.dataset.bound = "1";
      btn.onclick = (e) => {
        e.preventDefault();
        const product = btn.getAttribute("data-product") || "";
        const fi = Number(btn.getAttribute("data-slot") || 0);
        const slots = state.costFiles[product] || [];
        if (slots[fi]) revokeSlot(slots[fi]);
        slots.splice(fi, 1);
        if (!slots.length) slots.push(emptyFileSlot());
        const section = btn.closest(".sd-product-cost");
        const wrap = section.querySelector(".cost-slots");
        wrap.innerHTML = slots.map((_, i) => costFileSlotHtml(product, i)).join("");
        bindCostSheets(form);
      };
    });
  }

  function valuesTable(row, editable) {
    const lines = namedLines(row);
    if (!lines.length) return "<p class=\"sd-process-sub\">Add product names on the Enquiries sheet first.</p>";
    return "<table class=\"sd-lines\"><thead><tr><th>Product</th><th>Value incl VAT</th></tr></thead><tbody>" +
      lines.map((l, i) => "<tr><td>" + esc(l.product) + "</td><td>" +
        (editable
          ? "<input data-val=\"" + i + "\" value=\"" + esc(displayIncl(l) || "") + "\" inputmode=\"decimal\" placeholder=\"0.00\">"
          : esc(displayIncl(l) ? formatRand(displayIncl(l)) : "—")) +
        "</td></tr>").join("") +
      "</tbody></table>" +
      "<label>Delivery incl VAT *" + (editable
        ? "<input name=\"delivery_incl_vat\" value=\"" + esc(displayDeliveryIncl(row) || "") + "\" inputmode=\"decimal\" placeholder=\"0.00\">"
        : "</label><div>" + esc(displayDeliveryIncl(row) ? formatRand(displayDeliveryIncl(row)) : "—") + "</div>") +
      (editable ? "</label>" : "") +
      (editable
        ? "<div class=\"sd-quote-totals\" data-quote-totals>" +
          "<div>Products incl VAT <b data-tot=\"products\">R 0.00</b></div>" +
          "<div>Delivery incl VAT <b data-tot=\"delivery\">R 0.00</b></div>" +
          "<div class=\"sd-quote-total\">Total incl VAT <b data-tot=\"incl\">R 0.00</b></div>" +
          "<div>VAT 15% <b data-tot=\"vat\">R 0.00</b></div>" +
          "<div>Total excl VAT <b data-tot=\"excl\">R 0.00</b></div>" +
          "<p class=\"sd-process-sub\">Check Total incl VAT against the quote PDF as you type. Exclusive VAT is saved too.</p>" +
          "</div>"
        : "");
  }

  function paintQuoteTotals(form) {
    const box = form && form.querySelector("[data-quote-totals]");
    if (!box) return;
    let productsIncl = 0;
    let productsExcl = 0;
    form.querySelectorAll("input[data-val]").forEach((input) => {
      productsIncl += parseMoney(input.value);
      productsExcl += parseMoney(exclFromIncl(input.value) || 0);
    });
    const deliveryEl = form.querySelector('[name="delivery_incl_vat"]');
    const deliveryIncl = parseMoney(deliveryEl ? deliveryEl.value : "");
    const deliveryExcl = parseMoney(exclFromIncl(deliveryEl ? deliveryEl.value : "") || 0);
    const incl = Math.round((productsIncl + deliveryIncl) * 100) / 100;
    const excl = Math.round((productsExcl + deliveryExcl) * 100) / 100;
    const vat = Math.round((incl - excl) * 100) / 100;
    const set = (key, n) => {
      const el = box.querySelector('[data-tot="' + key + '"]');
      if (el) el.textContent = formatRand(n);
    };
    set("products", productsIncl);
    set("delivery", deliveryIncl);
    set("incl", incl);
    set("vat", vat);
    set("excl", excl);
  }

  function bindQuoteTotals(form) {
    if (!form || !form.querySelector("[data-quote-totals]")) return;
    form.addEventListener("input", () => paintQuoteTotals(form));
    paintQuoteTotals(form);
  }

  function readValues(form, row) {
    const lines = namedLines(row).map((l, i) => {
      const input = form.querySelector('input[data-val="' + i + '"]');
      const incl = input ? input.value : displayIncl(l);
      return {
        product: l.product,
        category: l.category || "",
        value_incl_vat: incl,
        value_excl_vat: exclFromIncl(incl)
      };
    });
    const delivery = form.querySelector('[name="delivery_incl_vat"]');
    const deliveryIncl = delivery ? delivery.value : displayDeliveryIncl(row);
    return {
      products: lines,
      delivery_incl_vat: deliveryIncl,
      delivery_excl_vat: exclFromIncl(deliveryIncl)
    };
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
  async function openSavedFile(kind, filename, outlook) {
    const r = await sdOfficeFetch(fileUrl(kind, !!outlook));
    if (!r.ok) throw new Error("Could not open that file");
    const blob = await r.blob();
    const name = filename || "file";
    const type = outlook
      ? (/\.eml$/i.test(name) ? "message/rfc822" : "application/vnd.ms-outlook")
      : (blob.type || "application/octet-stream");
    const file = new File([blob], name, { type });
    const url = URL.createObjectURL(file);
    const download = !!outlook || /\.(msg|eml|xlsx|xls|csv)$/i.test(name);
    if (download) {
      const a = document.createElement("a");
      a.href = url;
      a.download = name;
      a.rel = "noopener";
      document.body.appendChild(a);
      a.click();
      a.remove();
    } else {
      window.open(url, "_blank", "noopener");
    }
    setTimeout(() => URL.revokeObjectURL(url), download ? 4000 : 60000);
  }
  async function openSavedOutlookMail(kind, filename) {
    return openSavedFile(kind, filename, true);
  }
  function correspondenceFields() {
    return "<p class=\"sd-process-sub\">Paste the file link. It is saved on the server as Correspondance link — do not attach a file. Any link or path is fine, including those with //.</p>" +
      "<label>Correspondance link<textarea class=\"sd-path\" name=\"correspondence_links\" rows=\"3\" placeholder=\"Paste any link or path\" autocomplete=\"off\"></textarea></label>";
  }
  function fileHref(kind) {
    return "/api/office/enquiries/" + encodeURIComponent(state.enquiryNo) + "/files/" + encodeURIComponent(kind);
  }
  function absoluteHref(pathOrUrl) {
    const raw = String(pathOrUrl || "").trim();
    if (!raw) return "";
    if (/^\/[^/]/.test(raw)) {
      try { return new URL(raw, window.location.origin).toString(); } catch (e) { return raw; }
    }
    return raw;
  }
  function filesCard(row) {
    const items = (row && Array.isArray(row.deliverables)) ? row.deliverables : [];
    let html = "<div class=\"sd-files sd-correspondence\"><h2>Files</h2>" +
      "<p class=\"sd-process-sub\">CORRESPONDANCE is a pasted file link saved on the server as Correspondance link — any link or path, including those with //. Cost sheets are per product (more than one sheet per item is fine). Quote PDF (including earlier quotes), follow-up screenshot, proof of payment, and drawing stay here too. Copy link or Open — you do not attach a file for Correspondance.</p>";
    if (!items.length) {
      return html + "<p class=\"sd-process-sub\">No files on this enquiry yet.</p></div>";
    }
    html += "<div class=\"sd-file-list\">";
    html += items.map((f) => {
      const server = f.kind ? absoluteHref(fileHref(f.kind)) : "";
      const href = (f.kind && (f.open || f.stored_as) && server) ? server : (f.href || "");
      const open = href
        ? "<a class=\"sd-open-mail\" href=\"" + esc(href) + "\" target=\"_blank\" rel=\"noopener\">Open</a>"
        : "<span class=\"sd-process-sub\">No link yet</span>";
      const copy = href
        ? "<button type=\"button\" class=\"ghost\" data-copy-link=\"" + esc(href) + "\">Copy link</button>"
        : "";
      return "<div class=\"sd-file-row\">" +
        "<div class=\"sd-file-meta\"><div class=\"sd-file-type\">" + esc(f.label || "File") + "</div>" +
        "<div class=\"sd-file-name\">" + esc(f.group === "correspondence" ? "Correspondance link" : (f.title || f.filename || "File")) + "</div>" +
        (f.from ? "<div class=\"sd-process-sub\">" + esc(f.from) + "</div>" : "") +
        (href ? "<div class=\"sd-process-sub\" style=\"word-break:break-all\">" + esc(href) + "</div>" : "") +
        "</div><div class=\"row-actions\">" + copy + open + "</div></div>";
    }).join("");
    return html + "</div></div>";
  }
  function correspondenceCard(row) {
    return filesCard(row);
  }
  function formFor(action, row) {
    const waiting = (state.snap.waitingStatuses || []).map((s) => "<option>" + esc(s) + "</option>").join("");
    const closed = (state.snap.closedStatuses || []).map((s) => "<option>" + esc(s) + "</option>").join("");
    if (action.id === "assign_waiting") {
      return "<label>Waiting on<select name=\"waiting_status\">" + waiting + "</select></label>" +
        "<label>Assign to</label>" + assigneeSelect();
    }
    if (action.id === "assign_costing") {
      const recost = /Quoted|Followed Up/.test(row.status || "");
      const coster = openAssignee(row, "cost_sheet") || rolePerson("costing");
      return "<p class=\"sd-process-sub\">" +
        (recost
          ? "The client stays on this enquiry. Costing runs again, then you issue another quote PDF. Previous quotes stay in Files."
          : "Any office Admin, including yourself. Defaults to the Costing person on Users.") +
        "</p><label>Assign to</label>" + assigneeSelect(coster) +
        (recost ? "" : correspondenceFields());
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
        "<label>Assign to</label>" + assigneeSelect(rolePerson("costing")) +
        "<label>Comment<textarea name=\"comments\"></textarea></label>" +
        correspondenceFields();
    }
    if (action.id === "complete_cost_sheet") {
      return productCostBlocks(row) +
        "<label>Request approval from (optional)</label>" + assigneeSelect(rolePerson("approval"), "assignee") +
        "<p class=\"sd-process-sub\">Defaults to the Approval person on Users. Untick Approval on Users if cost sheets should skip approval.</p>" +
        "<label>Quoting person *</label>" + assigneeSelect(row.quote_assignee || rolePerson("quoting"), "quote_assignee");
    }
    if (action.id === "complete_approval") {
      return costSheetLinks(row) +
        productNamesLine(row) +
        "<label>Decision<select name=\"decision\"><option value=\"approve\">Approve — send to quote</option><option value=\"reject\">Reject — back to costing</option></select></label>" +
        "<label>Comments (required if rejected)<textarea name=\"comments\"></textarea></label>" +
        "<label>Quoting person if approved</label>" +
        assigneeSelect(row.quote_assignee || rolePerson("quoting"), "quote_assignee") +
        "<label>Costing person if rejected</label>" +
        assigneeSelect(rolePerson("costing") || lastCosting(row));
    }
    if (action.id === "complete_quote") {
      const hint = (state.snap && state.snap.quoteNo) || {};
      const recent = (hint.recent || []).slice();
      const quotedAlready = /Quoted|Followed Up/.test(row.status || "") || (row.quotes || []).length > 0 || !!row.quote_no;
      const next = quotedAlready ? (hint.next || "") : (row.quote_no || hint.next || "");
      const recentLine = recent.length
        ? "Last quotation numbers: " + recent.join(", ") + "."
        : "No quotation numbers yet.";
      const prev = (row.quotes || []).map((q) => q.quote_no).filter(Boolean);
      return (quotedAlready
        ? "<p class=\"sd-process-sub\">Previous quote" + (prev.length === 1 ? "" : "s") +
          (prev.length ? " (" + prev.join(", ") + ")" : "") +
          " stay on this enquiry. The sheet shows the latest quotation number.</p>"
        : "") +
        valuesTable(row, true) +
        "<label>Quotation number *<input class=\"sd-quote-no\" name=\"quote_no\" value=\"" + esc(next) + "\" autocomplete=\"off\"></label>" +
        "<p class=\"sd-process-sub\">" + esc(recentLine) + " Default is the next number (" + esc(hint.next || next) + "). You can change it, but it cannot match an existing quotation.</p>" +
        fileBlock("application/pdf,.pdf", quotedAlready
          ? "Upload the revised quote PDF. DATE QUOTED updates to today. Enter the new values including VAT."
          : "Enter each product value and delivery including VAT here, then upload the quote PDF. DATE QUOTED is saved with the PDF.") +
        "<label>Who follows up after 7 days?</label>" + assigneeSelect(row.follow_up_assignee || openAssignee(row, "follow_up"));
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

  function accessKindLabel(kind) {
    const labels = (state.snap && state.snap.access && state.snap.access.kindLabel) || {};
    if (typeof labels === "function") return labels(kind);
    return ({
      chase_info: "chase missing information",
      cost_sheet: "upload the cost sheet",
      supplier: "record the supplier answer",
      approval: "approve or reject costing",
      quote: "upload the quote PDF",
      follow_up: "log a follow-up",
      pop: "attach proof of payment or the client outcome",
      drawing: "upload the drawing"
    })[kind] || String(kind || "this step");
  }

  async function postProcess(body) {
    const r = await sdOfficeFetch("/api/office/enquiries/" + encodeURIComponent(state.enquiryNo) + "/process", {
      method: "POST",
      body: JSON.stringify(body || {})
    });
    return r.json();
  }

  function grantInboxHtml(snap) {
    const pending = (snap && snap.access && snap.access.pendingForMe) || [];
    if (!pending.length) return "";
    return "<div class=\"sd-process-card\"><h2>Access requests</h2>" +
      "<p class=\"sd-process-sub\">Someone else needs to upload a deliverable that is assigned to you" +
      (snap.is_manager ? " (or you are the Manager)" : "") + ".</p>" +
      pending.map((g) => {
        return "<div class=\"sd-grant-row\" data-grant=\"" + esc(g.id) + "\">" +
          "<div class=\"grow\"><b>" + esc(g.requester) + "</b> wants to " + esc(accessKindLabel(g.kind)) +
          ".<div class=\"sd-process-sub\">Assigned to " + esc(g.assignee) + "</div></div>" +
          "<button type=\"button\" data-grant-act=\"grant_access\" data-grant-id=\"" + esc(g.id) + "\">Grant</button>" +
          "<button type=\"button\" class=\"ghost\" data-grant-act=\"deny_access\" data-grant-id=\"" + esc(g.id) + "\">Refuse</button>" +
          "</div>";
      }).join("") +
      "</div>";
  }

  function lockedActionHtml(action) {
    const owner = action.assignee || "the assigned person";
    const mine = ((state.snap && state.snap.access && state.snap.access.mine) || [])
      .filter((g) => g.kind === action.kind);
    const pending = action.request_pending || mine.some((g) => g.status === "pending");
    const granted = mine.some((g) => g.status === "granted");
    let extra = "";
    if (granted) extra = "<p class=\"sd-process-sub\">Access is granted. Reload this card if Save is still locked.</p>";
    else if (pending) extra = "<p class=\"sd-process-sub\">Waiting for " + esc(owner) + " or the Manager to grant access.</p>";
    else extra = "<p class=\"sd-process-sub\">Ask " + esc(owner) + " or the Manager. They will see the request here and on My tasks.</p>";
    return "<p class=\"sd-process-sub\"><b>" + esc(owner) + "</b> is assigned this step. Only they (or the Manager) can save the deliverable and move STATUS.</p>" +
      extra +
      (pending || granted ? "" : "<button type=\"button\" data-request-access=\"" + esc(action.id) + "\" data-request-kind=\"" + esc(action.kind || "") + "\">Request access</button>") +
      "<div class=\"sd-process-err\" data-err></div>";
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
    if (action.id === "complete_cost_sheet") {
      body.cost_sheets = namedLines(row).map((l) => {
        const product = l.product;
        const section = Array.from(form.querySelectorAll(".sd-product-cost")).find((el) => el.getAttribute("data-product") === product);
        const slots = (state.costFiles && state.costFiles[product]) || [];
        return {
          product,
          files: slots.map((slot, fi) => {
            const ok = section && section.querySelector('.cost-file-ok[data-slot="' + fi + '"]');
            return {
              file_name: slot.name,
              file_type: slot.mime,
              file_base64: slot.base64,
              file_confirmed: !!(ok && ok.checked) || !!slot.confirmed
            };
          }).filter((f) => f.file_base64)
        };
      });
    } else if (/complete_quote|complete_followup|complete_order|complete_drawing/.test(action.id)) {
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
    if (/Costing|Re-Cost/.test(row.status || "") && actions.some((a) => a.id === "add_correspondence")) {
      return action.id === "add_correspondence";
    }
    if (actions.some((a) => a.id === "assign_costing") && !/Quoted|Followed Up/.test(row.status || "")) {
      return action.id === "assign_costing";
    }
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
      (row.date_quoted ? "<span>Quoted " + esc(row.date_quoted) + (row.quote_no ? " · " + esc(row.quote_no) : "") +
        ((row.quotes || []).length > 1 ? " · " + (row.quotes.length) + " quotes" : "") + "</span>" : "") +
      (row.ready_for_orders ? "<span>Ready for Orders</span>" : "") +
      (row.lifespan_label ? "<span class=\"sd-life\">Lifespan <b>" + esc(row.lifespan_label) + "</b></span>" : "") +
      "</div>";
    const events = row.events || [];
    if (events.length) {
      html += "<h2>Timeline</h2><ol class=\"sd-timeline\">" + events.map((ev) => {
        return "<li><time>" + esc(ev.at_label || ev.at) + "</time>" +
          esc(ev.label || ev.kind) +
          (ev.status ? " · " + esc(ev.status) : "") +
          (ev.actor ? "<div class=\"sd-tl-actor\">" + esc(ev.actor) + "</div>" : "") +
          (ev.note ? "<div class=\"sd-tl-actor\">" + esc(ev.note) + "</div>" : "") +
          "</li>";
      }).join("") + "</ol>";
    }
    html += correspondenceCard(row);
    if (openTasks.length) {
      html += "<h2>Assigned now</h2>" + openTasks.map((t) => {
        return "<span class=\"sd-task-pill\">" + esc(t.title) + " → " + esc(t.assignee) + "</span>";
      }).join("");
    } else {
      html += "<p class=\"sd-process-sub\">No open assigned task. Capture and assign the next person from here.</p>";
    }
    if (row.ready_for_orders) {
      html += "<p class=\"sd-process-sub\">POP" + (row.drawing && row.drawing.required ? " and drawing" : "") +
        " are on file. Create the Orders row with this enquiry number, client, quote, products, and values.</p>" +
        (row.order_number
          ? "<p class=\"sd-process-sub\">Order <b>" + esc(row.order_number) + "</b> is already on Orders.</p>" +
            "<a href=\"/orders\"><button type=\"button\" class=\"ghost sd-create-order\">Open Orders</button></a>"
          : "<button type=\"button\" class=\"sd-create-order\" id=\"sdCreateOrder\">Create order from this enquiry</button>") +
        "<div class=\"sd-process-err\" id=\"sdCreateOrderErr\"></div>";
    }
    html += "</div>";
    html += grantInboxHtml(snap);
    const actions = snap.actions || [];
    if (!actions.length) {
      html += "<p class=\"sd-process-sub\">This enquiry has no further process steps.</p>";
    } else {
      html += "<div class=\"sd-process-actions\">";
      actions.forEach((action, i) => {
        const open = shouldExpandAction(action, i, row, actions);
        const locked = action.can_act === false;
        html += "<details class=\"sd-process-card" + (locked ? " sd-locked" : "") + "\" data-action-i=\"" + i + "\"" + (open ? " open" : "") + ">" +
          "<summary><h2>" + esc(action.label) + (locked ? " · assigned to " + esc(action.assignee || "") : "") + "</h2></summary>";
        if (locked) {
          html += "<div class=\"sd-process-form\">" + lockedActionHtml(action) + "</div>";
        } else {
          html += "<form class=\"sd-process-form\">" + formFor(action, row) +
            "<div class=\"sd-process-err\" data-err></div>" +
            "<button type=\"submit\">Save update</button></form>";
        }
        html += "</details>";
      });
      html += "</div>";
    }
    body.innerHTML = html;
    body.querySelectorAll("form").forEach((form) => {
      bindFile(form);
      bindCostSheets(form);
      bindQuoteTotals(form);
      const card = form.closest("[data-action-i]");
      const i = Number((card || form).getAttribute("data-action-i"));
      const preview = form.querySelector("[data-preview]");
      if (preview && row.cost_sheet && !form.querySelector('[name="file"]') && !form.querySelector(".cost-file-input")) {
        showSaved(row.cost_sheet.kind || "cost_sheet", preview);
      }
      form.onsubmit = async (e) => {
        e.preventDefault();
        const action = actions[i];
        const err = form.querySelector("[data-err]");
        err.textContent = "";
        const r = await postProcess(collect(form, action, row));
        const j = r;
        if (!j.ok) { err.textContent = j.error || "Could not save"; return; }
        state.snap = j;
        revokeFile();
        renderBody();
      };
    });
    body.querySelectorAll("[data-copy-link]").forEach((btn) => {
      btn.onclick = async (e) => {
        e.preventDefault();
        const href = btn.getAttribute("data-copy-link") || "";
        try {
          if (navigator.clipboard && navigator.clipboard.writeText) await navigator.clipboard.writeText(href);
          else {
            const ta = document.createElement("textarea");
            ta.value = href;
            document.body.appendChild(ta);
            ta.select();
            document.execCommand("copy");
            ta.remove();
          }
          btn.textContent = "Copied";
          setTimeout(() => { btn.textContent = "Copy link"; }, 1600);
        } catch (err) {
          btn.textContent = "Copy failed";
        }
      };
    });
    body.querySelectorAll("[data-open-file],[data-open-mail]").forEach((btn) => {
      btn.onclick = async (e) => {
        e.preventDefault();
        try {
          const kind = btn.getAttribute("data-open-file") || btn.getAttribute("data-open-mail");
          const outlook = btn.getAttribute("data-file-outlook") === "1" || btn.hasAttribute("data-open-mail");
          await openSavedFile(kind, btn.getAttribute("data-mail-name"), outlook);
        } catch (err) {
          btn.textContent = "Could not open";
        }
      };
    });
    body.querySelectorAll("[data-grant-act]").forEach((btn) => {
      btn.onclick = async (e) => {
        e.preventDefault();
        const j = await postProcess({ action: btn.getAttribute("data-grant-act"), grant_id: btn.getAttribute("data-grant-id") });
        if (!j.ok) {
          btn.textContent = j.error || "Could not save";
          return;
        }
        state.snap = j;
        renderBody();
      };
    });
    body.querySelectorAll("[data-request-access]").forEach((btn) => {
      btn.onclick = async (e) => {
        e.preventDefault();
        const card = btn.closest("[data-action-i]");
        const err = card && card.querySelector("[data-err]");
        const j = await postProcess({
          action: "request_access",
          for_action: btn.getAttribute("data-request-access"),
          kind: btn.getAttribute("data-request-kind")
        });
        if (!j.ok) {
          if (err) err.textContent = j.error || "Could not request";
          return;
        }
        state.snap = j;
        renderBody();
      };
    });
    const createBtn = document.getElementById("sdCreateOrder");
    if (createBtn) {
      createBtn.onclick = async (e) => {
        e.preventDefault();
        const err = document.getElementById("sdCreateOrderErr");
        if (err) err.textContent = "";
        const r = await sdOfficeFetch("/api/office/enquiries/" + encodeURIComponent(state.enquiryNo) + "/create-order", { method: "POST", body: "{}" });
        const j = await r.json();
        if (!j.ok) {
          if (err) err.textContent = j.error || "Could not create the order";
          return;
        }
        window.location.href = "/orders";
      };
    }
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
