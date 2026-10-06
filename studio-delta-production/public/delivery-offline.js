"use strict";

(function (root) {
  var DB_NAME = "sd-delivery-offline";
  var DB_VERSION = 1;
  var RUN_KEY = "run";

  function queuedOrderKeys(queue) {
    var keys = {};
    (queue || []).forEach(function (item) {
      (item.order_numbers || []).forEach(function (n) {
        keys[String(n || "").trim().toUpperCase()] = true;
      });
      (item.bases || []).forEach(function (b) {
        keys[String(b || "").trim().toUpperCase()] = true;
      });
      var payload = item.payload || {};
      if (payload.order_number) keys[String(payload.order_number).trim().toUpperCase()] = true;
    });
    return keys;
  }

  function splitRunOrders(orders, queue) {
    var keys = queuedOrderKeys(queue);
    var live = [];
    var waiting = [];
    (orders || []).forEach(function (row) {
      var num = String((row && row.order_number) || "").trim().toUpperCase();
      var base = String((row && row.base) || "").trim().toUpperCase();
      if (keys[num] || keys[base]) waiting.push(row);
      else live.push(row);
    });
    return { live: live, waiting: waiting };
  }

  function isNetworkError(err, status) {
    if (status === 0 || status >= 500) return true;
    var m = String((err && err.message) || err || "").toLowerCase();
    return /failed to fetch|networkerror|load failed|offline|internet|network request failed|the internet connection appears to be offline/.test(m);
  }

  function openDb() {
    return new Promise(function (resolve, reject) {
      if (!root.indexedDB) {
        reject(new Error("This phone cannot save delivery forms offline."));
        return;
      }
      var req = root.indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function () {
        var db = req.result;
        if (!db.objectStoreNames.contains("kv")) db.createObjectStore("kv");
        if (!db.objectStoreNames.contains("queue")) db.createObjectStore("queue", { keyPath: "id" });
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error || new Error("Could not open offline storage")); };
    });
  }

  function idbGet(store, key) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(store, "readonly");
        var req = tx.objectStore(store).get(key);
        req.onsuccess = function () { resolve(req.result); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function idbPut(store, value, key) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(store, "readwrite");
        var req = key != null ? tx.objectStore(store).put(value, key) : tx.objectStore(store).put(value);
        req.onsuccess = function () { resolve(value); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function idbDelete(store, key) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(store, "readwrite");
        var req = tx.objectStore(store).delete(key);
        req.onsuccess = function () { resolve(); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function idbGetAll(store) {
    return openDb().then(function (db) {
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(store, "readonly");
        var req = tx.objectStore(store).getAll();
        req.onsuccess = function () { resolve(req.result || []); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function saveRun(snapshot) {
    return idbPut("kv", Object.assign({}, snapshot, { saved_at: new Date().toISOString() }), RUN_KEY);
  }

  function loadRun() {
    return idbGet("kv", RUN_KEY).then(function (row) { return row || null; });
  }

  function listQueue() {
    return idbGetAll("queue").then(function (rows) {
      return (rows || []).sort(function (a, b) {
        return String(a.created_at || "").localeCompare(String(b.created_at || ""));
      });
    });
  }

  function enqueue(item) {
    var row = Object.assign({
      id: item.id || (root.crypto && crypto.randomUUID ? crypto.randomUUID() : String(Date.now())),
      created_at: item.created_at || new Date().toISOString(),
      error: "",
      tries: 0
    }, item);
    return idbPut("queue", row);
  }

  function removeItem(id) {
    return idbDelete("queue", id);
  }

  function markItem(id, patch) {
    return idbGet("queue", id).then(function (row) {
      if (!row) return null;
      return idbPut("queue", Object.assign({}, row, patch || {}));
    });
  }

  function alreadySavedResult(json) {
    return !!(json && json.ok && (json.already || json.id || json.url));
  }

  function flushQueue(opts) {
    var fetchFn = (opts && opts.fetchFn) || root.fetch;
    if (typeof fetchFn !== "function") return Promise.resolve({ sent: 0, left: 0, error: "no fetch" });
    if (root.navigator && root.navigator.onLine === false) {
      return listQueue().then(function (rows) {
        return { sent: 0, left: rows.length, offline: true };
      });
    }
    return listQueue().then(function (rows) {
      var sent = 0;
      var auth = false;
      function next(i) {
        if (i >= rows.length || auth) {
          return listQueue().then(function (left) {
            return { sent: sent, left: left.length, auth: auth };
          });
        }
        var item = rows[i];
        return fetchFn("/api/delivery/pod", {
          method: "POST",
          credentials: "same-origin",
          headers: {
            "Content-Type": "application/json",
            "x-sd-offline-queue": "1"
          },
          body: JSON.stringify(item.payload || {})
        }).then(function (r) {
          if (r.status === 401) {
            auth = true;
            return markItem(item.id, { error: "Log in again to send saved forms.", tries: (item.tries || 0) + 1 })
              .then(function () { return next(i + 1); });
          }
          return r.json().catch(function () { return {}; }).then(function (j) {
            if (alreadySavedResult(Object.assign({ ok: r.ok }, j)) || (r.ok && j && j.ok)) {
              sent += 1;
              return removeItem(item.id).then(function () { return next(i + 1); });
            }
            var msg = (j && j.error) || ("Could not send (" + r.status + ")");
            if (/not loaded on the truck/i.test(msg) && r.status === 400) {
              sent += 1;
              return removeItem(item.id).then(function () { return next(i + 1); });
            }
            if (r.status >= 500 || r.status === 0) {
              return markItem(item.id, { error: "Waiting for a better signal.", tries: (item.tries || 0) + 1 })
                .then(function () { return next(i + 1); });
            }
            return markItem(item.id, { error: msg, tries: (item.tries || 0) + 1 })
              .then(function () { return next(i + 1); });
          });
        }).catch(function (err) {
          return markItem(item.id, {
            error: isNetworkError(err) ? "Waiting for a better signal." : String(err && err.message || err),
            tries: (item.tries || 0) + 1
          }).then(function () { return next(i + 1); });
        });
      }
      return next(0);
    });
  }

  function requestBackgroundSync() {
    if (!root.navigator || !root.navigator.serviceWorker || !root.ServiceWorkerRegistration) {
      return Promise.resolve(false);
    }
    if (!("sync" in root.ServiceWorkerRegistration.prototype)) return Promise.resolve(false);
    return root.navigator.serviceWorker.ready.then(function (reg) {
      return reg.sync.register("sd-delivery-pod").then(function () { return true; });
    }).catch(function () { return false; });
  }

  var api = {
    queuedOrderKeys: queuedOrderKeys,
    splitRunOrders: splitRunOrders,
    isNetworkError: isNetworkError,
    saveRun: saveRun,
    loadRun: loadRun,
    listQueue: listQueue,
    enqueue: enqueue,
    removeItem: removeItem,
    markItem: markItem,
    flushQueue: flushQueue,
    requestBackgroundSync: requestBackgroundSync
  };

  root.sdDeliveryOffline = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof self !== "undefined" ? self : this);
