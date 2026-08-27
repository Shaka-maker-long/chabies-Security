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
function sdShowLogin(message) {
  return new Promise((resolve) => {
    const wrap = document.createElement("div");
    wrap.style.cssText = "position:fixed;inset:0;background:#f8f9fc;display:flex;align-items:center;justify-content:center;z-index:99;font-family:Inter,system-ui,sans-serif";
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
  sdHideDebtorsLinks(!!profile.canSeeDebtors);
  return profile;
}
