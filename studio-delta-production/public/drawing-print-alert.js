(function () {
  const STYLE_ID = "sdDrawingPrintStyle";
  const ROOT_ID = "sdDrawingPrintRoot";
  let timer = null;
  let busy = false;

  function esc(s) {
    return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
  }

  function profileName() {
    try {
      const office = JSON.parse(sessionStorage.getItem("sd-office") || "null");
      if (office && office.name) return String(office.name).trim();
    } catch (e) {}
    try {
      const shop = JSON.parse(sessionStorage.getItem("sd-session") || "null");
      if (shop && shop.name) return String(shop.name).trim();
    } catch (e) {}
    return "";
  }

  function token() {
    try {
      const office = JSON.parse(sessionStorage.getItem("sd-office") || "null");
      if (office && office.token) return String(office.token);
    } catch (e) {}
    try {
      const shop = JSON.parse(sessionStorage.getItem("sd-session") || "null");
      if (shop && shop.token) return String(shop.token);
    } catch (e) {}
    return "";
  }

  function isSiya(name) {
    const n = String(name || profileName()).trim();
    if (!n) return false;
    const first = n.split(/\s+/)[0];
    return n.toLowerCase() === "siya" || first.toLowerCase() === "siya";
  }

  function ensureStyle() {
    if (document.getElementById(STYLE_ID)) return;
    const css = document.createElement("style");
    css.id = STYLE_ID;
    css.textContent =
      "#" + ROOT_ID + "{position:fixed;inset:0;z-index:2147483000;display:none;align-items:stretch;justify-content:center;background:#111827;color:#fff;font-family:Outfit,Inter,system-ui,sans-serif}" +
      "#" + ROOT_ID + ".open{display:flex}" +
      "#" + ROOT_ID + " .sd-dpa-card{width:100%;min-height:100%;padding:28px 20px 40px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:18px;text-align:center;box-sizing:border-box}" +
      "#" + ROOT_ID + " h1{margin:0;font-size:clamp(28px,6vw,48px);line-height:1.1;max-width:16ch}" +
      "#" + ROOT_ID + " p{margin:0;font-size:18px;line-height:1.45;max-width:42em;color:#e5e7eb}" +
      "#" + ROOT_ID + " .sd-dpa-order{font-size:clamp(22px,4vw,32px);font-weight:800;color:#fbbf24}" +
      "#" + ROOT_ID + " .sd-dpa-list{width:min(720px,100%);display:flex;flex-direction:column;gap:14px}" +
      "#" + ROOT_ID + " .sd-dpa-item{border:1px solid #374151;border-radius:16px;padding:16px;background:#1f2937;text-align:left}" +
      "#" + ROOT_ID + " a{color:#93c5fd;font-weight:700}" +
      "#" + ROOT_ID + " button{border:0;background:#fbbf24;color:#111827;font:800 18px Outfit,Inter,system-ui,sans-serif;border-radius:12px;padding:16px 20px;cursor:pointer;width:100%}" +
      "#" + ROOT_ID + " button:disabled{opacity:.6;cursor:wait}" +
      "#" + ROOT_ID + " .sd-dpa-err{color:#fca5a5;min-height:1.2em}";
    document.head.appendChild(css);
  }

  function root() {
    let el = document.getElementById(ROOT_ID);
    if (el) return el;
    el = document.createElement("div");
    el.id = ROOT_ID;
    el.setAttribute("role", "alertdialog");
    el.setAttribute("aria-modal", "true");
    el.setAttribute("aria-labelledby", "sdDrawingPrintTitle");
    document.body.appendChild(el);
    return el;
  }

  function hide() {
    const el = document.getElementById(ROOT_ID);
    if (el) el.classList.remove("open");
    document.documentElement.style.overflow = "";
  }

  function show(alerts) {
    ensureStyle();
    const el = root();
    const rows = Array.isArray(alerts) ? alerts : [];
    if (!rows.length) {
      hide();
      return;
    }
    el.innerHTML = "<div class=\"sd-dpa-card\">" +
      "<h1 id=\"sdDrawingPrintTitle\">New drawing — print for production</h1>" +
      "<p>A drawing was updated after the job card was generated. Print the new drawing for the production team before you continue.</p>" +
      "<div class=\"sd-dpa-list\">" + rows.map((row) => {
        const href = String(row.drawing_url || "").trim();
        return "<div class=\"sd-dpa-item\" data-dpa-id=\"" + esc(row.id) + "\">" +
          "<div class=\"sd-dpa-order\">" + esc(row.order_number || "Order") + "</div>" +
          "<p>" + esc(row.client_name || "") + (row.product ? " · " + esc(row.product) : "") + "</p>" +
          (href ? "<p><a href=\"" + esc(href) + "\" target=\"_blank\" rel=\"noopener\">Open drawing</a></p>" : "") +
          "<button type=\"button\" data-dpa-ack=\"" + esc(row.id) + "\">I have seen this — I will print " + esc(row.order_number || "this drawing") + "</button>" +
          "<div class=\"sd-dpa-err\" data-dpa-err></div></div>";
      }).join("") +
      "</div></div>";
    el.classList.add("open");
    document.documentElement.style.overflow = "hidden";
    el.querySelectorAll("[data-dpa-ack]").forEach((btn) => {
      btn.onclick = function () { ack(btn.getAttribute("data-dpa-ack"), btn); };
    });
  }

  function headers() {
    const h = { "Content-Type": "application/json" };
    const t = token();
    if (t) h["x-sd-token"] = t;
    return h;
  }

  async function poll() {
    if (!isSiya() || busy || document.hidden) return;
    try {
      const r = await fetch("/api/drawing-print-alerts", { credentials: "same-origin", headers: headers() });
      const j = await r.json().catch(function () { return {}; });
      if (j && Array.isArray(j.alerts)) show(j.alerts);
    } catch (e) {}
  }

  async function ack(id, btn) {
    if (!id || busy) return;
    busy = true;
    if (btn) {
      btn.disabled = true;
      btn.textContent = "Saving…";
    }
    try {
      const r = await fetch("/api/drawing-print-alerts/" + encodeURIComponent(id) + "/ack", {
        method: "POST",
        credentials: "same-origin",
        headers: headers(),
        body: "{}"
      });
      const j = await r.json().catch(function () { return {}; });
      if (!j || !j.ok) {
        const err = btn && btn.parentElement ? btn.parentElement.querySelector("[data-dpa-err]") : null;
        if (err) err.textContent = (j && j.error) || "Could not save";
        if (btn) {
          btn.disabled = false;
          btn.textContent = "I have seen this — I will print the drawing";
        }
        busy = false;
        return;
      }
      busy = false;
      poll();
    } catch (e) {
      busy = false;
      if (btn) {
        btn.disabled = false;
        const err = btn.parentElement ? btn.parentElement.querySelector("[data-dpa-err]") : null;
        if (err) err.textContent = "Could not save";
      }
    }
  }

  function start(profile) {
    stop();
    const name = (profile && profile.name) || profileName();
    if (!isSiya(name)) return;
    poll();
    timer = setInterval(poll, 8000);
  }

  function stop() {
    if (timer) clearInterval(timer);
    timer = null;
  }

  window.sdStartDrawingPrintAlerts = start;
  window.sdStopDrawingPrintAlerts = stop;

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () { start(); });
  } else {
    start();
  }
})();
