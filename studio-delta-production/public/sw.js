/* Studio Delta PWA — live office/floor data is never cached.
   Driver run HTML/assets are cached so PODs can be filled with no signal. */
importScripts("/delivery-offline.js");

const CACHE = "sd-pwa-v22-driver-pin";
const PRECACHE = [
  "/",
  "/offline.html",
  "/delivery-run",
  "/delivery-offline.js?v=del-off2",
  "/delivery-offline.js",
  "/vendor/leaflet/leaflet.css",
  "/vendor/leaflet/leaflet.js",
  "/manifest.webmanifest",
  "/sd-pwa.js?v=pwa-driverpin",
  "/sd-brand.css?v=logged-in-2",
  "/sd-splash.js?v=erp-shell",
  "/office-auth.js?v=login-msg-3",
  "/facility-floor.png",
  "/icons/icon-192.png",
  "/icons/icon-512.png",
  "/icons/apple-touch-icon.png",
  "/icons/maskable-512.png",
  "/icons/icon.svg",
  "/vendor/bootstrap/bootstrap.min.css",
  "/vendor/bootstrap/bootstrap.bundle.min.js",
  "/vendor/bootstrap-icons/bootstrap-icons.css",
  "/vendor/bootstrap-icons/fonts/bootstrap-icons.woff2",
  "/vendor/bootstrap-icons/fonts/bootstrap-icons.woff",
  "/vendor/chartjs/chart.umd.min.js",
  "/vendor/sheetjs/xlsx.full.min.js"
];

function isApi(url) {
  return url.pathname.indexOf("/api/") === 0;
}

function isHtml(request) {
  if (request.mode === "navigate") return true;
  const accept = request.headers.get("accept") || "";
  return accept.indexOf("text/html") >= 0;
}

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE).then((cache) => cache.addAll(PRECACHE)).then(() => self.skipWaiting())
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
      .then(() => self.clients.matchAll({ type: "window", includeUncontrolled: true }))
      .then((list) => {
        list.forEach((client) => {
          try { client.postMessage({ type: "sd-sw-updated", cache: CACHE }); } catch (e) {}
        });
      })
  );
});

self.addEventListener("fetch", (event) => {
  const request = event.request;
  if (request.method !== "GET") return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;
  if (isApi(url) || url.pathname.indexOf("/outlook-addin") === 0) return;
  if (url.pathname === "/sw.js") return;

  if (isHtml(request)) {
    event.respondWith(
      fetch(request).then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((cache) => cache.put(request, copy)).catch(() => {});
        return res;
      }).catch(() => caches.match(request).then((hit) => hit || caches.match("/offline.html")))
    );
    return;
  }

  event.respondWith(
    caches.match(request).then((hit) => {
      const fresh = fetch(request).then((res) => {
        if (res && res.ok) {
          const copy = res.clone();
          caches.open(CACHE).then((cache) => cache.put(request, copy)).catch(() => {});
        }
        return res;
      }).catch(() => hit);
      return hit || fresh;
    })
  );
});

function notifyDeliveryClients(payload) {
  return self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((list) => {
    list.forEach((client) => {
      try { client.postMessage(Object.assign({ type: "sd-delivery-offline" }, payload || {})); } catch (e) {}
    });
  });
}

self.addEventListener("sync", (event) => {
  if (event.tag !== "sd-delivery-pod") return;
  event.waitUntil(
    self.sdDeliveryOffline.flushQueue().then((result) => notifyDeliveryClients({ flushed: result }))
  );
});

self.addEventListener("message", (event) => {
  const data = event.data || {};
  if (data.type === "SKIP_WAITING") {
    self.skipWaiting();
    return;
  }
  if (data.type === "sd-delivery-flush") {
    event.waitUntil(
      self.sdDeliveryOffline.flushQueue().then((result) => notifyDeliveryClients({ flushed: result }))
    );
  }
});
