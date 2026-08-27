/**
 * google.script.run shim for Railway / any host that is not Apps Script.
 * On Apps Script the native google.script.run already exists, so this file is a no-op.
 */
(function (global) {
  if (global.google && global.google.script && global.google.script.run) return;

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
          fetch("/api/run", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ fn: prop, args: args })
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
