(function (root) {
  var THEME = "#1c1917";

  function ensureHead() {
    var head = document.head || document.getElementsByTagName("head")[0];
    if (!head) return;
    if (!document.querySelector('link[rel="manifest"]')) {
      var man = document.createElement("link");
      man.rel = "manifest";
      man.href = "/manifest.webmanifest";
      head.appendChild(man);
    }
    if (!document.querySelector('meta[name="theme-color"]')) {
      var theme = document.createElement("meta");
      theme.name = "theme-color";
      theme.content = THEME;
      head.appendChild(theme);
    }
    if (!document.querySelector('link[rel="apple-touch-icon"]')) {
      var apple = document.createElement("link");
      apple.rel = "apple-touch-icon";
      apple.href = "/icons/apple-touch-icon.png";
      head.appendChild(apple);
    }
    if (!document.querySelector('meta[name="apple-mobile-web-app-capable"]')) {
      var cap = document.createElement("meta");
      cap.name = "apple-mobile-web-app-capable";
      cap.content = "yes";
      head.appendChild(cap);
    }
    if (!document.querySelector('meta[name="mobile-web-app-capable"]')) {
      var mcap = document.createElement("meta");
      mcap.name = "mobile-web-app-capable";
      mcap.content = "yes";
      head.appendChild(mcap);
    }
    if (!document.querySelector('meta[name="apple-mobile-web-app-title"]')) {
      var title = document.createElement("meta");
      title.name = "apple-mobile-web-app-title";
      title.content = "Studio Delta";
      head.appendChild(title);
    }
    if (!document.querySelector('meta[name="apple-mobile-web-app-status-bar-style"]')) {
      var bar = document.createElement("meta");
      bar.name = "apple-mobile-web-app-status-bar-style";
      bar.content = "black-translucent";
      head.appendChild(bar);
    }
  }

  function standalone() {
    return (root.matchMedia && root.matchMedia("(display-mode: standalone)").matches)
      || root.navigator.standalone === true;
  }

  function registerWorker() {
    if (!("serviceWorker" in navigator)) return;
    navigator.serviceWorker.register("/sw.js", { scope: "/" }).catch(function () {});
  }

  function showInstall(deferred) {
    if (standalone()) return;
    try {
      if (localStorage.getItem("sd-pwa-install-dismissed") === "1") return;
    } catch (e) {}
    if (document.getElementById("sdPwaInstall")) return;
    var bar = document.createElement("div");
    bar.id = "sdPwaInstall";
    bar.setAttribute("role", "dialog");
    bar.setAttribute("aria-label", "Install Studio Delta");
    bar.innerHTML = '<span>Install Studio Delta on this device</span>' +
      '<span class="sd-pwa-actions">' +
      '<button type="button" data-pwa="install">Install</button>' +
      '<button type="button" class="ghost" data-pwa="later">Not now</button>' +
      "</span>";
    var css = document.createElement("style");
    css.textContent = "#sdPwaInstall{position:fixed;left:12px;right:12px;bottom:12px;z-index:11000;display:flex;gap:12px;align-items:center;justify-content:space-between;flex-wrap:wrap;background:#1c1917;color:#f3efe6;border:1px solid #b08948;border-radius:10px;padding:12px 14px;font:600 13px Inter,system-ui,sans-serif;box-shadow:0 10px 28px rgba(28,25,23,.28)}" +
      "#sdPwaInstall .sd-pwa-actions{display:flex;gap:8px}" +
      "#sdPwaInstall button{border:0;background:#b08948;color:#1c1917;border-radius:6px;padding:8px 12px;font:700 12px Inter,system-ui,sans-serif;cursor:pointer}" +
      "#sdPwaInstall button.ghost{background:transparent;color:#f3efe6;border:1px solid #d7d1c6}";
    document.head.appendChild(css);
    document.body.appendChild(bar);
    bar.addEventListener("click", function (e) {
      var act = e.target && e.target.getAttribute("data-pwa");
      if (act === "later") {
        try { localStorage.setItem("sd-pwa-install-dismissed", "1"); } catch (err) {}
        bar.remove();
      }
      if (act === "install" && deferred) {
        deferred.prompt();
        deferred.userChoice.finally(function () { bar.remove(); });
      }
    });
  }

  function bootInstall() {
    var deferred = null;
    root.addEventListener("beforeinstallprompt", function (e) {
      e.preventDefault();
      deferred = e;
      showInstall(deferred);
    });
    if (standalone()) document.documentElement.classList.add("sd-standalone");
  }

  ensureHead();
  registerWorker();
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bootInstall);
  } else {
    bootInstall();
  }
})(typeof window !== "undefined" ? window : this);
