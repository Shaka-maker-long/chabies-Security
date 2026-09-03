const SD_ORDER_STATUSES = [
  "Not Yet Started",
  "Ready for Steelwork", "Profile Cutting",
  "Ready for Tagging", "Tagging",
  "Ready for Welding", "Welding",
  "Ready for Grinding", "Grinding",
  "Ready for Pre-Powder Coating", "Pre-Powder Coating",
  "Ready for Powder Coating", "Powder Coating",
  "Ready for Assembly", "Paint Preparation", "Ready for Painting", "Painting", "Assembly",
  "Ready for Final QC", "Final QC",
  "Ready for Delivery", "Out for Delivery",
  "Delivered"
];

function sdOfficeProfile() {
  try { return JSON.parse(sessionStorage.getItem("sd-office") || "null"); } catch (e) { return null; }
}
function sdSaveOffice(profile) {
  try { sessionStorage.setItem("sd-office", JSON.stringify(profile)); } catch (e) {}
  try { localStorage.removeItem("sd-office"); } catch (e) {}
}
function sdClearOfficeProfile() {
  try { sessionStorage.removeItem("sd-office"); } catch (e) {}
  try { localStorage.removeItem("sd-office"); } catch (e) {}
}
function sdOfficeFetch(url, opts) {
  const p = sdOfficeProfile() || {};
  opts = opts || {};
  opts.headers = Object.assign({ "Content-Type": "application/json", "x-sd-token": p.token || "" }, opts.headers || {});
  return fetch(url, opts);
}
function sdHideDebtorsLinks(canSee) {
  document.querySelectorAll("[data-nav='debtors']").forEach((a) => {
    a.style.display = canSee ? "" : "none";
  });
}
function sdNavCollapsed() {
  return localStorage.getItem("sd-nav-collapsed") === "1";
}
function sdApplyNavCollapsed() {
  const on = sdNavCollapsed();
  document.body.classList.toggle("nav-collapsed", on);
  document.documentElement.classList.toggle("nav-collapsed", on);
  const btn = document.getElementById("sdCollapseBtn");
  if (btn) {
    const icon = btn.querySelector("i");
    const label = btn.querySelector(".sd-link-text");
    if (icon) icon.className = on ? "bi bi-chevron-right" : "bi bi-chevron-left";
    if (label) label.textContent = on ? "Show menu" : "Hide menu";
    btn.title = on ? "Show menu" : "Hide menu";
  }
}
function sdToggleNavCollapsed() {
  localStorage.setItem("sd-nav-collapsed", sdNavCollapsed() ? "0" : "1");
  sdApplyNavCollapsed();
}
function sdEnsureSheet(rel, extra) {
  const bare = rel.split("?")[0];
  if (document.querySelector('link[href*="' + bare + '"]')) return;
  const l = document.createElement("link");
  l.rel = "stylesheet";
  l.href = rel;
  if (extra) Object.keys(extra).forEach((k) => l.setAttribute(k, extra[k]));
  document.head.appendChild(l);
}
function sdEnsureScript(src) {
  return new Promise((resolve) => {
    const bare = src.split("?")[0];
    const found = Array.prototype.find.call(document.scripts || [], (s) => (s.src || "").indexOf(bare) !== -1);
    if (found) {
      if (typeof sdShowWelcome === "function") return resolve();
      found.addEventListener("load", () => resolve());
      found.addEventListener("error", () => resolve());
      return;
    }
    const s = document.createElement("script");
    s.src = src;
    s.onload = () => resolve();
    s.onerror = () => resolve();
    document.head.appendChild(s);
  });
}
function sdLoadBrand() {
  sdEnsureSheet("/sd-brand.css?v=erp-one");
  sdEnsureSheet("/office-shell.css?v=erp-shell");
  return sdEnsureScript("/sd-splash.js?v=erp-shell");
}
function sdForgetOffice() {
  sdClearOfficeProfile();
  return fetch("/api/office/logout", { method: "POST", credentials: "same-origin" }).catch(function () {});
}
function sdOfficeLogout() {
  sdForgetOffice().finally(function () { location.href = "/"; });
}
function sdMountOfficeShell(active) {
  if (document.getElementById("sdSidebar")) {
    sdApplyNavCollapsed();
    return;
  }
  sdEnsureSheet("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css");
  sdEnsureSheet("/office-shell.css?v=erp-shell");
  document.body.classList.add("office-app");
  const items = [
    ["/", "home", "bi-house-door", "Home"],
    ["/?view=floor", "floor", "bi-tools", "Floor"],
    ["/orders", "orders", "bi-table", "Orders"],
    ["/enquiries", "enquiries", "bi-journal-text", "Enquiries"],
    ["/tasks", "tasks", "bi-check2-square", "My tasks"],
    ["/schedule", "schedule", "bi-calendar2-week", "Office schedule"],
    ["/dropdowns", "dropdowns", "bi-list-ul", "Dropdowns"],
    ["/users", "users", "bi-person-plus", "Users"],
    ["/durations", "durations", "bi-hourglass-split", "Task times"],
    ["/debtors", "debtors", "bi-cash-coin", "Debtors"],
    ["/?view=production", "production", "bi-clipboard-data", "Production"],
    ["/?view=workers", "workers", "bi-people", "Workers"],
    ["/?view=metrics", "metrics", "bi-bar-chart", "Metrics"],
    ["/?view=qc", "qc", "bi-file-earmark-pdf", "QC Reports"],
    ["/?view=activity", "activity", "bi-calendar3", "Activity"],
    ["/?view=schedule", "floorschedule", "bi-calendar-week", "Schedule"]
  ];
  const nav = document.createElement("nav");
  nav.className = "sd-sidebar";
  nav.id = "sdSidebar";
  nav.setAttribute("aria-label", "Main navigation");
  const mark = typeof sdSMark === "function" ? sdSMark() : "";
  nav.innerHTML =
    '<div class="sd-sidebar-header">' + mark + '<h5 class="sd-brand-text">STUDIO DELTA</h5></div>' +
    '<div class="sd-sidebar-scroll">' +
    items.map(([href, id, icon, label]) => {
      const on = id === active ? " active" : "";
      const debt = id === "debtors" ? " data-nav=\"debtors\"" : "";
      const section = id === "orders" ? '<div class="sd-nav-label">Office</div>' : (id === "production" ? '<div class="sd-nav-label">Shop</div>' : "");
      return section + '<a class="sd-link' + on + '" href="' + href + '"' + debt + '><i class="bi ' + icon + '"></i><span class="sd-link-text">' + label + "</span></a>";
    }).join("") +
    "</div>" +
    '<div class="sd-sidebar-footer">' +
    '<button type="button" class="sd-collapse-btn" id="sdCollapseBtn" title="Hide menu"><i class="bi bi-chevron-left"></i><span class="sd-link-text">Hide menu</span></button>' +
    '<button type="button" class="sd-logout-btn" id="sdLogoutBtn"><i class="bi bi-box-arrow-right"></i><span class="sd-link-text">Log Out</span></button>' +
    "</div>";
  const burger = document.createElement("button");
  burger.type = "button";
  burger.className = "sd-nav-burger";
  burger.id = "sdNavBurger";
  burger.setAttribute("aria-label", "Open menu");
  burger.innerHTML = '<i class="bi bi-list"></i>';
  const backdrop = document.createElement("div");
  backdrop.className = "sd-backdrop";
  backdrop.id = "sdBackdrop";
  document.body.insertBefore(backdrop, document.body.firstChild);
  document.body.insertBefore(nav, document.body.firstChild);
  document.body.insertBefore(burger, document.body.firstChild);
  burger.onclick = () => document.body.classList.toggle("nav-open");
  backdrop.onclick = () => document.body.classList.remove("nav-open");
  document.getElementById("sdCollapseBtn").onclick = sdToggleNavCollapsed;
  document.getElementById("sdLogoutBtn").onclick = sdOfficeLogout;
  nav.querySelectorAll("a").forEach((a) => {
    a.addEventListener("click", () => document.body.classList.remove("nav-open"));
  });
  sdApplyNavCollapsed();
}
function sdShowLogin(message) {
  return new Promise((resolve) => {
    const wrap = document.createElement("div");
    wrap.className = "sd-login-mask";
    const mark = typeof sdSMark === "function" ? sdSMark("sd-s-mark sd-s-login") : "";
    wrap.innerHTML = '<form class="sd-login-card">' +
      '<div class="sd-login-mark">' + mark + "</div>" +
      "<h2>Studio Delta</h2>" +
      "<p data-login-hint>" + (message || "Office pages are for Admin. Use the same name and access code as the floor.") + "</p>" +
      "<label>Name</label><input name='name' autocomplete='username'>" +
      "<label>Access code</label><input name='password' type='password' autocomplete='current-password'>" +
      "<button type='submit'>Log in</button>" +
      "<p style='margin:14px 0 0;text-align:center'><a href='/'>Back to floor</a></p></form>";
    document.body.appendChild(wrap);
    fetch("/health").then((r) => r.json()).then((j) => {
      const hint = wrap.querySelector("[data-login-hint]");
      if (!hint || message) return;
      if (j && j.usingEphemeralDisk) {
        hint.textContent = "This Railway service has no volume, so logins reset on every deploy. Use your Users name and code. If that fails right after a deploy, wait a few seconds or try Admin / admin.";
      }
    }).catch(function () {});
    wrap.querySelector("form").onsubmit = async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const r = await fetch("/api/office/login", {
        method: "POST",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: fd.get("name"), password: fd.get("password") })
      });
      const j = await r.json();
      if (!j.ok) { alert(j.error || "Login failed"); return; }
      sdSaveOffice(j);
      wrap.remove();
      if (typeof sdShowWelcome === "function") await sdShowWelcome(5000);
      resolve(j);
    };
  });
}
async function sdRequireOffice(page) {
  await sdLoadBrand();
  let profile = sdOfficeProfile();
  if (profile && profile.token) {
    const r = await fetch("/api/office/me", { headers: { "x-sd-token": profile.token } });
    const j = await r.json();
    if (j.ok) {
      profile = Object.assign({}, profile, j.profile);
      sdSaveOffice(profile);
    } else profile = null;
  } else {
    profile = null;
  }
  if (!profile || !profile.canSeeOffice) {
    profile = await sdShowLogin("Log in with your name and access code.");
  }
  if (!profile || !profile.canSeeOffice) {
    location.replace("/");
    return null;
  }
  if (page === "debtors" && !profile.canSeeDebtors) {
    location.replace("/orders");
    return null;
  }
  sdMountOfficeShell(page);
  sdHideDebtorsLinks(!!profile.canSeeDebtors);
  sdWarnPersistence();
  return profile;
}
function sdWarnPersistence() {
  fetch("/health").then((r) => r.json()).then((j) => {
    if (!j || (!j.usingEphemeralDisk && !j.warning)) return;
    if (document.getElementById("sdPersistBanner")) return;
    const bar = document.createElement("div");
    bar.id = "sdPersistBanner";
    bar.className = "sd-persist-banner";
    bar.setAttribute("role", "alert");
    bar.textContent = j.warning || "This Railway service has no volume. Enquiries and shop data are wiped on every deploy. In Railway, add a Volume mounted at /app/data.";
    document.body.appendChild(bar);
    document.body.classList.add("has-persist-warning");
  }).catch(function () {});
}
