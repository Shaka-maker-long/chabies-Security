/**
 * google.script.run shim for Railway / any host that is not Apps Script.
 * On Apps Script the native google.script.run already exists, so this file is a no-op.
 */
(function (global) {
  if (global.google && global.google.script && global.google.script.run) return;

  function deviceId() {
    try {
      if (typeof global.sdDeviceId === "function") return global.sdDeviceId();
      var id = String(global.localStorage.getItem("sd-device-id") || "").trim();
      if (/^[A-Za-z0-9_-]{8,80}$/.test(id)) return id;
      id = "d_";
      for (var i = 0; i < 24; i++) id += Math.floor(Math.random() * 16).toString(16);
      global.localStorage.setItem("sd-device-id", id);
      return id;
    } catch (e) {
      return "";
    }
  }

  function runner() {
    var success = null;
    var failure = null;
    var chain = {
      withSuccessHandler: function (fn) {
        success = fn;
        return proxy;
      },
      withFailureHandler: function (fn) {
        failure = fn;
        return proxy;
      }
    };
    var proxy = new Proxy(chain, {
      get: function (target, prop) {
        if (prop in target) return target[prop];
        if (typeof prop !== "string") return undefined;
        if (prop === "then" || prop === "toJSON") return undefined;
        return function () {
          var args = Array.prototype.slice.call(arguments);
          var headers = {
            "Content-Type": "application/json",
            "x-sd-device-id": deviceId()
          };
          var body = { fn: prop, args: args, deviceId: deviceId() };
          fetch("/api/run", {
            method: "POST",
            credentials: "same-origin",
            headers: headers,
            body: JSON.stringify(body)
          })
            .then(function (r) {
              return r.json().then(function (j) {
                return { http: r, j: j };
              });
            })
            .then(function (pack) {
              var j = pack.j || {};
              if (!j.ok) {
                var err = j.error || ("HTTP " + pack.http.status);
                if (failure) failure(err);
                else console.error(err);
                return;
              }
              if (success) success(j.result);
            })
            .catch(function (e) {
              var msg = e && e.message ? e.message : String(e);
              if (failure) failure(msg);
              else console.error(msg);
            });
          return proxy;
        };
      }
    });
    return proxy;
  }

  global.google = global.google || {};
  Object.defineProperty(global.google, "script", {
    configurable: true,
    get: function () {
      return { run: runner() };
    }
  });
})(window);
