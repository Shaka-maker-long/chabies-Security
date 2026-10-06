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

  var SESSION_KEY = "sd-delivery-session";
  var memoryStore = {};

  function storage() {
    try {
      if (root.localStorage) return root.localStorage;
    } catch (e) {}
    return {
      getItem: function (k) { return Object.prototype.hasOwnProperty.call(memoryStore, k) ? memoryStore[k] : null; },
      setItem: function (k, v) { memoryStore[k] = String(v); },
      removeItem: function (k) { delete memoryStore[k]; }
    };
  }

  function normalizeLoginName(name) {
    return String(name || "").trim().toLowerCase();
  }

  function pinMaterial(name, pin) {
    return "sd-delivery-v1|" + normalizeLoginName(name) + "|" + String(pin || "");
  }

  function fnv1aHex(str) {
    var h = 2166136261;
    for (var i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = Math.imul(h, 16777619) >>> 0;
    }
    return ("00000000" + h.toString(16)).slice(-8);
  }

  function hashPin(name, pin) {
    var material = pinMaterial(name, pin);
    function fnv() { return Promise.resolve("fnv:" + fnv1aHex(material)); }
    if (!(root.crypto && root.crypto.subtle && typeof root.TextEncoder === "function")) return fnv();
    try {
      return root.crypto.subtle.digest("SHA-256", new TextEncoder().encode(material)).then(function (buf) {
        var bytes = new Uint8Array(buf);
        var hex = "";
        for (var i = 0; i < bytes.length; i++) hex += ("0" + bytes[i].toString(16)).slice(-2);
        return "sha256:" + hex;
      }).catch(fnv);
    } catch (e) {
      return fnv();
    }
  }

  function hashesEqual(a, b) {
    a = String(a || "");
    b = String(b || "");
    if (a.length !== b.length) return false;
    var diff = 0;
    for (var i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
    return diff === 0;
  }

  function isDriverProfile(profile) {
    if (!profile) return false;
    if (profile.canSeeOffice) return false;
    var title = String(profile.jobTitle || profile.role || "").trim().toLowerCase();
    if (title === "driver") return true;
    return (profile.tasks || []).some(function (t) {
      return String(t).trim().toLowerCase() === "delivery";
    });
  }

  function profileSnapshot(profile) {
    return {
      success: true,
      token: profile.token || "",
      name: profile.name,
      jobTitle: profile.jobTitle || profile.role || "Driver",
      isAdmin: false,
      canSeeOffice: false,
      canSeeDebtors: !!profile.canSeeDebtors,
      canSeeIdleAlerts: !!profile.canSeeIdleAlerts,
      canManageUsers: !!profile.canManageUsers,
      access: profile.access || "Production",
      isQcOnly: !!profile.isQcOnly,
      tasks: Array.isArray(profile.tasks) && profile.tasks.length ? profile.tasks.slice() : ["Delivery"],
      role: profile.role || null
    };
  }

  function loadDeliverySession() {
    try {
      var raw = JSON.parse(storage().getItem(SESSION_KEY) || "null");
      if (!raw || !raw.pinHash || !raw.profile || !raw.profile.name) return null;
      return raw;
    } catch (e) {
      return null;
    }
  }

  function clearDeliverySession() {
    try { storage().removeItem(SESSION_KEY); } catch (e) {}
  }

  function saveDeliverySession(profile, pin) {
    if (!isDriverProfile(profile) || !String(pin || "").length) return Promise.resolve(false);
    return hashPin(profile.name, pin).then(function (pinHash) {
      var payload = {
        v: 1,
        pinHash: pinHash,
        nameKey: normalizeLoginName(profile.name),
        savedAt: Date.now(),
        profile: profileSnapshot(profile)
      };
      storage().setItem(SESSION_KEY, JSON.stringify(payload));
      return true;
    }).catch(function () { return false; });
  }

  function unlockDeliverySession(name, pin) {
    var saved = loadDeliverySession();
    if (!saved || !saved.pinHash) return Promise.resolve(null);
    return hashPin(name, pin).then(function (hash) {
      if (!hashesEqual(hash, saved.pinHash)) return null;
      if (saved.nameKey && saved.nameKey !== normalizeLoginName(name)) return null;
      return saved.profile || null;
    }).catch(function () { return null; });
  }

  function offlineLoginMessage(opts) {
    opts = opts || {};
    if (opts.hasSaved) {
      return "No signal. Use the same name and access code as the last driver login on this phone.";
    }
    return "No signal. Log in once with data on this phone. After that a driver can open the run without a connection. Floor clocks still need signal.";
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
    requestBackgroundSync: requestBackgroundSync,
    isDriverProfile: isDriverProfile,
    hashPin: hashPin,
    loadDeliverySession: loadDeliverySession,
    saveDeliverySession: saveDeliverySession,
    clearDeliverySession: clearDeliverySession,
    unlockDeliverySession: unlockDeliverySession,
    offlineLoginMessage: offlineLoginMessage
  };

  root.sdDeliveryOffline = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof self !== "undefined" ? self : this);
