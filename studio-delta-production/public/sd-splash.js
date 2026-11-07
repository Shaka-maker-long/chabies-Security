(function (root) {
  var S_PATH = "M92 36C92 18 76 8 56 8C30 8 14 22 14 42C14 62 28 72 56 82C84 92 100 104 100 126C100 148 82 160 56 160C28 160 10 144 10 124";

  function sMark(className) {
    return '<svg class="' + (className || "sd-s-mark") + '" viewBox="0 0 120 172" aria-hidden="true">' +
      '<path d="' + S_PATH + '"></path></svg>';
  }

  function paintPath(path) {
    if (!path || typeof path.getTotalLength !== "function") return;
    var len = Math.ceil(path.getTotalLength());
    path.style.strokeDasharray = String(len);
    path.style.strokeDashoffset = String(len);
    path.getBoundingClientRect();
    path.style.transition = "stroke-dashoffset 2.4s ease";
    path.style.strokeDashoffset = "0";
  }

  function envLabel(isStaging) {
    return isStaging ? "Staging" : "Production";
  }

  function readStagingHint() {
    try {
      if (document.documentElement.classList.contains("sd-staging")) return true;
      if (document.getElementById("sdStagingBanner")) return true;
      if (String(document.title || "").indexOf("STAGING") !== -1) return true;
    } catch (e) {}
    return false;
  }

  function fetchIsStaging() {
    return fetch("/health", { credentials: "same-origin", cache: "no-store" })
      .then(function (r) { return r.json(); })
      .then(function (j) { return !!(j && j.isStaging); })
      .catch(function () { return readStagingHint(); });
  }

  function showWelcome(ms) {
    ms = Number(ms) > 0 ? Number(ms) : 5000;
    return new Promise(function (resolve) {
      var existing = document.getElementById("sdWelcome");
      if (existing) {
        resolve();
        return;
      }
      var wrap = document.createElement("div");
      wrap.id = "sdWelcome";
      wrap.className = "sd-welcome";
      wrap.setAttribute("role", "dialog");
      wrap.setAttribute("aria-label", "Welcome to Studio Delta");
      wrap.innerHTML = sMark("sd-s") +
        "<h1>Welcome to Studio Delta</h1>" +
        "<p data-sd-env>" + envLabel(readStagingHint()) + "</p>" +
        '<div class="sd-welcome-rule" aria-hidden="true"></div>';
      document.body.appendChild(wrap);
      var path = wrap.querySelector("path");
      requestAnimationFrame(function () { paintPath(path); });
      fetchIsStaging().then(function (isStaging) {
        var label = wrap.querySelector("[data-sd-env]");
        if (label) label.textContent = envLabel(isStaging);
        if (isStaging) wrap.classList.add("sd-welcome-staging");
      });
      setTimeout(function () {
        wrap.classList.add("hide");
        setTimeout(function () {
          if (wrap.parentNode) wrap.parentNode.removeChild(wrap);
          resolve();
        }, 450);
      }, ms);
    });
  }

  root.sdSMark = sMark;
  root.sdShowWelcome = showWelcome;
})(typeof window !== "undefined" ? window : this);
