function sdPasteEsc(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
}
function sdPasteMoney(s) {
  const n = Number(String(s || "").replace(/,/g, "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0;
}
function sdPasteRand(n) {
  const v = sdPasteMoney(n);
  const neg = v < 0 ? "-" : "";
  const abs = (Math.round(Math.abs(v) * 100) / 100).toFixed(2);
  const [whole, frac] = abs.split(".");
  return neg + "R " + whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + "." + frac;
}
function sdPasteSetErr(msg) {
  const el = document.getElementById("pasteErr");
  if (el) el.textContent = msg || "";
}
function sdPastePersist() {
  if (typeof sdWriteFormDraft !== "function") return;
  const text = document.getElementById("pasteText");
  const paid = document.getElementById("pastePaid");
  if (!text) return;
  sdWriteFormDraft("order-paste", "box", {
    text: text.value,
    paid: !(paid && paid.checked === false)
  });
}
function sdPasteRestore() {
  if (typeof sdReadFormDraft !== "function") return;
  const draft = sdReadFormDraft("order-paste", "box");
  if (!draft) return;
  const text = document.getElementById("pasteText");
  const paid = document.getElementById("pastePaid");
  if (text && draft.text) text.value = draft.text;
  if (paid && draft.paid === false) paid.checked = false;
}
function sdPasteRenderResult(j, saved) {
  const host = document.getElementById("pasteResult");
  if (!host) return;
  const added = j.added || [];
  const skipped = j.skipped || [];
  const errors = j.errors || [];
  if (!added.length && !skipped.length && !errors.length) {
    host.hidden = true;
    host.innerHTML = "";
    return;
  }
  host.hidden = false;
  let html = "";
  if (added.length) {
    html += "<p class=\"hint\">" + (saved ? "Added " : "Ready to add ") + added.length + (added.length === 1 ? " order." : " orders.") + "</p>";
    html += "<div class=\"paste-table-wrap\"><table class=\"paste-preview\"><thead><tr><th>Order</th><th>Product</th><th>Client</th><th>Status</th><th>Operator</th><th>Incl VAT</th><th>Paid</th></tr></thead><tbody>";
    html += added.map((row) => {
      return "<tr><td>" + sdPasteEsc(row.order_number) + "</td><td>" + sdPasteEsc(row.product) + "</td><td>" + sdPasteEsc(row.client_name) + "</td><td>" + sdPasteEsc(row.status) + "</td><td>" + sdPasteEsc(row.assigned_operator || "—") + "</td><td>" + sdPasteEsc(sdPasteRand(row.price_incl_vat)) + "</td><td>" + sdPasteEsc(sdPasteRand(row.amount_paid)) + "</td></tr>";
    }).join("");
    html += "</tbody></table></div>";
  }
  if (skipped.length) {
    html += "<p class=\"hint\">Already on Orders: " + skipped.map((s) => sdPasteEsc(s.order_number)).join(", ") + "</p>";
  }
  if (errors.length) {
    html += "<p class=\"err\">" + errors.map((e) => sdPasteEsc(e)).join(" · ") + "</p>";
  }
  host.innerHTML = html;
}
async function sdPastePost(preview) {
  sdPasteSetErr("");
  const text = document.getElementById("pasteText");
  const paid = document.getElementById("pastePaid");
  if (!text || !String(text.value || "").trim()) {
    sdPasteSetErr("Paste the order rows from the old sheet.");
    return null;
  }
  const r = await sdOfficeFetch("/api/office/orders/paste", {
    method: "POST",
    body: JSON.stringify({
      text: text.value,
      paid_in_full: !(paid && paid.checked === false),
      preview: !!preview
    })
  });
  const j = await r.json().catch(function () { return {}; });
  if (!j.ok) {
    sdPasteSetErr(j.error || "Could not read that paste.");
    return null;
  }
  sdPasteRenderResult(j, !preview);
  if (!(j.added || []).length && !(j.skipped || []).length) {
    sdPasteSetErr((j.errors && j.errors[0]) || "No order rows were found in that paste.");
  }
  return j;
}
function sdInitOrderPaste(opts) {
  opts = opts || {};
  sdPasteRestore();
  const text = document.getElementById("pasteText");
  const paid = document.getElementById("pastePaid");
  if (text) {
    text.addEventListener("input", sdPastePersist);
    text.addEventListener("paste", function () { setTimeout(sdPastePersist, 0); });
  }
  if (paid) paid.addEventListener("change", sdPastePersist);
  window.previewPaste = function () { return sdPastePost(true); };
  window.confirmPaste = async function () {
    const j = await sdPastePost(false);
    if (!j) return;
    if ((j.added || []).length) {
      if (text) text.value = "";
      if (typeof sdDropFormDraft === "function") sdDropFormDraft("order-paste", "box");
      if (typeof opts.onAdded === "function") opts.onAdded(j);
    }
  };
}
