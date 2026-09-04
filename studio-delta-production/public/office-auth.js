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
function sdHideUsersLink(canManage) {
  document.querySelectorAll("[data-nav='users']").forEach((a) => {
    a.style.display = canManage ? "" : "none";
  });
}
function sdShowChangePassword() {
  if (document.getElementById("sdPassMask")) return;
  const wrap = document.createElement("div");
  wrap.id = "sdPassMask";
  wrap.className = "sd-login-mask";
  wrap.innerHTML = "<form class=\"sd-login-card\" id=\"sdPassForm\">" +
    "<h2>Change access code</h2>" +
    "<p>Anyone can update their own code. Use the name you log in with.</p>" +
    "<label>Name</label><input name=\"name\" autocomplete=\"username\">" +
    "<label>Current access code</label><input name=\"current_password\" type=\"password\" autocomplete=\"current-password\">" +
    "<label>New access code</label><input name=\"new_password\" type=\"password\" autocomplete=\"new-password\">" +
    "<p class=\"sd-pass-err\" id=\"sdPassErr\" style=\"color:#b42318;font-size:13px;min-height:16px\"></p>" +
    "<button type=\"submit\">Save access code</button>" +
    "<button type=\"button\" class=\"ghost\" id=\"sdPassCancel\" style=\"margin-top:8px;background:#fff;color:#1d2939;border:1px solid #1d2939\">Cancel</button>" +
    "</form>";
  document.body.appendChild(wrap);
  const profile = sdOfficeProfile() || {};
  if (profile.name) wrap.querySelector("[name=name]").value = profile.name;
  wrap.querySelector("#sdPassCancel").onclick = () => wrap.remove();
  wrap.addEventListener("click", (e) => { if (e.target.id === "sdPassMask") wrap.remove(); });
  wrap.querySelector("form").onsubmit = async (e) => {
    e.preventDefault();
    const err = document.getElementById("sdPassErr");
    err.textContent = "";
    const fd = new FormData(e.target);
    const r = await fetch("/api/office/password", {
      method: "POST",
      credentials: "same-origin",
      headers: Object.assign({ "Content-Type": "application/json" }, profile.token ? { "x-sd-token": profile.token } : {}),
      body: JSON.stringify({
        name: fd.get("name"),
        current_password: fd.get("current_password"),
        new_password: fd.get("new_password")
      })
    });
    const j = await r.json().catch(function () { return {}; });
    if (!j.ok) { err.textContent = j.error || "Could not save"; return; }
    wrap.remove();
    alert("Access code updated.");
  };
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
  sdEnsureSheet("/sd-brand.css?v=logged-in-2");
  sdEnsureSheet("/office-shell.css?v=logged-in");
  return sdEnsureScript("/sd-splash.js?v=erp-shell");
}
function sdForgetOffice() {
  sdClearOfficeProfile();
  return fetch("/api/office/logout", { method: "POST", credentials: "same-origin" }).catch(function () {});
}
function sdOfficeLogout() {
  sdForgetOffice().finally(function () { location.href = "/"; });
}
function sdPaintOfficeWho(profile) {
  const nameEl = document.getElementById("officeWhoName");
  const titleEl = document.getElementById("officeWhoTitle");
  const who = document.getElementById("officeWho");
  const name = String(profile && profile.name || "").trim();
  const title = String(profile && (profile.jobTitle || profile.role) || "").trim();
  if (nameEl) nameEl.textContent = name;
  if (titleEl) {
    titleEl.textContent = title;
    titleEl.hidden = !title;
  }
  if (who) who.hidden = !name;
}
function sdMountOfficeShell(active) {
  if (document.getElementById("sdSidebar")) {
    sdApplyNavCollapsed();
    return;
  }
  sdEnsureSheet("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css");
  sdEnsureSheet("/office-shell.css?v=logged-in");
  document.body.classList.add("office-app");
  const items = [
    ["/", "home", "bi-house-door", "Home"],
    ["/orders", "orders", "bi-table", "Orders"],
    ["/enquiries", "enquiries", "bi-journal-text", "Enquiries"],
    ["/tasks", "tasks", "bi-check2-square", "My tasks"],
    ["/schedule", "schedule", "bi-calendar2-week", "Office schedule"],
    ["/dropdowns", "dropdowns", "bi-list-ul", "Dropdowns"],
    ["/users", "users", "bi-person-plus", "Users", "users"],
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
    items.map(([href, id, icon, label, nav]) => {
      const on = id === active ? " active" : "";
      const extra = nav ? " data-nav=\"" + nav + "\"" : (id === "debtors" ? " data-nav=\"debtors\"" : "");
      const section = id === "orders" ? '<div class="sd-nav-label">Office</div>' : (id === "production" ? '<div class="sd-nav-label">Shop</div>' : "");
      return section + '<a class="sd-link' + on + '" href="' + href + '"' + extra + '><i class="bi ' + icon + '"></i><span class="sd-link-text">' + label + "</span></a>";
    }).join("") +
    "</div>" +
    '<div class="sd-sidebar-footer">' +
    '<div class="office-who" id="officeWho">' +
    '<span class="office-who-label">Logged in as</span>' +
    '<strong class="office-who-name" id="officeWhoName"></strong>' +
    '<span class="office-who-title" id="officeWhoTitle"></span>' +
    "</div>" +
    '<button type="button" class="sd-collapse-btn" id="sdCollapseBtn" title="Hide menu"><i class="bi bi-chevron-left"></i><span class="sd-link-text">Hide menu</span></button>' +
    '<button type="button" class="sd-logout-btn" id="sdPasswordBtn"><i class="bi bi-key"></i><span class="sd-link-text">Change access code</span></button>' +
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
  const passBtn = document.getElementById("sdPasswordBtn");
  if (passBtn) passBtn.onclick = sdShowChangePassword;
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
    let r = await fetch("/api/office/me", {
      credentials: "same-origin",
      headers: { "x-sd-token": profile.token }
    });
    let j = await r.json().catch(function () { return {}; });
    if (!(j && j.ok)) {
      r = await fetch("/api/office/me", { credentials: "same-origin" });
      j = await r.json().catch(function () { return {}; });
    }
    if (j && j.ok) {
      profile = Object.assign({}, profile, j.profile);
      if (profile.isAdmin) profile.canSeeOffice = true;
      sdSaveOffice(profile);
    } else profile = null;
  } else {
    profile = null;
  }
  const hadOffice = !!(profile && profile.canSeeOffice);
  if (!profile || !profile.canSeeOffice) {
    profile = await sdShowLogin("Log in with your name and access code.");
  }
  if (!profile || !profile.canSeeOffice) {
    location.replace("/");
    return null;
  }
  if (!hadOffice) {
    location.replace("/");
    return null;
  }
  if (page === "debtors" && !profile.canSeeDebtors) {
    location.replace("/orders");
    return null;
  }
  sdMountOfficeShell(page);
  sdPaintOfficeWho(profile);
  sdHideDebtorsLinks(!!profile.canSeeDebtors);
  sdHideUsersLink(!!profile.canManageUsers);
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
