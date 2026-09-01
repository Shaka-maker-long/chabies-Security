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
  try { return JSON.parse(localStorage.getItem("sd-office") || "null"); } catch (e) { return null; }
}
function sdSaveOffice(profile) {
  localStorage.setItem("sd-office", JSON.stringify(profile));
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
  if (document.querySelector('link[href="' + rel + '"]')) return;
  const l = document.createElement("link");
  l.rel = "stylesheet";
  l.href = rel;
  if (extra) Object.keys(extra).forEach((k) => l.setAttribute(k, extra[k]));
  document.head.appendChild(l);
}
function sdOfficeLogout() {
  try { localStorage.removeItem("sd-office"); } catch (e) {}
  location.href = "/";
}
function sdMountOfficeShell(active) {
  if (document.getElementById("sdSidebar")) {
    sdApplyNavCollapsed();
    return;
  }
  sdEnsureSheet("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css");
  sdEnsureSheet("/office-shell.css?v=fullmenu");
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
  nav.innerHTML =
    '<div class="sd-sidebar-header"><h5 class="sd-brand-text">STUDIO DELTA</h5></div>' +
    '<div class="sd-sidebar-scroll">' +
    items.map(([href, id, icon, label]) => {
      const on = id === active ? " active" : "";
      const debt = id === "debtors" ? " data-nav=\"debtors\"" : "";
      return '<a class="sd-link' + on + '" href="' + href + '"' + debt + '><i class="bi ' + icon + '"></i><span class="sd-link-text">' + label + "</span></a>";
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
    wrap.style.cssText = "position:fixed;inset:0;background:#f8f9fc;display:flex;align-items:center;justify-content:center;z-index:2000;font-family:Inter,system-ui,sans-serif";
    wrap.innerHTML = '<form style="background:#fff;border:1px solid #d0d5dd;border-radius:12px;padding:24px;width:min(360px,92vw)">' +
      "<h2 style='margin:0 0 8px;font-size:18px'>Admin login</h2>" +
      "<p style='margin:0 0 16px;color:#667085;font-size:13px'>" + (message || "Office pages are for Admin only.") + "</p>" +
      "<label style='font-size:12px'>Name</label><input name='name' style='width:100%;margin:4px 0 12px;padding:8px;border:1px solid #d0d5dd;border-radius:6px'>" +
      "<label style='font-size:12px'>Access code</label><input name='password' type='password' style='width:100%;margin:4px 0 16px;padding:8px;border:1px solid #d0d5dd;border-radius:6px'>" +
      "<button style='width:100%;padding:10px;border:0;border-radius:6px;background:#1d2939;color:#fff;font-weight:600'>Log in</button>" +
      "<p style='margin:12px 0 0;text-align:center'><a href='/'>Back to floor</a></p></form>";
    document.body.appendChild(wrap);
    wrap.querySelector("form").onsubmit = async (e) => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const r = await fetch("/api/office/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: fd.get("name"), password: fd.get("password") })
      });
      const j = await r.json();
      if (!j.ok) { alert(j.error || "Login failed"); return; }
      sdSaveOffice(j);
      wrap.remove();
      resolve(j);
    };
  });
}
async function sdRequireOffice(page) {
  let profile = sdOfficeProfile();
  if (profile && profile.token) {
    const r = await fetch("/api/office/me", { headers: { "x-sd-token": profile.token } });
    const j = await r.json();
    if (j.ok) {
      profile = Object.assign({}, profile, j.profile);
      sdSaveOffice(profile);
    } else profile = null;
  }
  if (!profile) profile = await sdShowLogin();
  if (!profile.canSeeOffice) {
    document.body.innerHTML = "<p style='font-family:sans-serif;padding:40px'>Production users can only use the <a href='/'>floor</a>.</p>";
    return null;
  }
  if (page === "debtors" && !profile.canSeeDebtors) {
    document.body.innerHTML = "<p style='font-family:sans-serif;padding:40px'>You do not have access to Debtors. <a href='/orders'>Orders</a></p>";
    return null;
  }
  sdMountOfficeShell(page);
  sdHideDebtorsLinks(!!profile.canSeeDebtors);
  return profile;
}
