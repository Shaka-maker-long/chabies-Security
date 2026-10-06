const fs = require("fs");
const path = require("path");
const assert = require("assert");

const root = path.join(__dirname, "..");
const publicDir = path.join(root, "public");

const manifest = JSON.parse(fs.readFileSync(path.join(publicDir, "manifest.webmanifest"), "utf8"));
assert.strictEqual(manifest.name, "Studio Delta");
assert.strictEqual(manifest.short_name, "Studio Delta");
assert.strictEqual(manifest.start_url, "/");
assert.strictEqual(manifest.scope, "/");
assert.strictEqual(manifest.display, "standalone");
assert.strictEqual(manifest.theme_color, "#1c1917");
assert.strictEqual(manifest.background_color, "#1c1917");
assert.ok(manifest.icons.some((i) => i.sizes === "192x192"));
assert.ok(manifest.icons.some((i) => i.sizes === "512x512" && i.purpose === "any"));
assert.ok(manifest.icons.some((i) => i.purpose === "maskable"));
assert.ok(manifest.shortcuts.some((s) => s.url === "/enquiries"));
assert.ok(manifest.shortcuts.some((s) => s.url === "/orders"));

["icon-192.png", "icon-512.png", "apple-touch-icon.png", "maskable-512.png", "icon.svg"].forEach((name) => {
  const file = path.join(publicDir, "icons", name);
  assert.ok(fs.existsSync(file), "missing icon " + name);
  assert.ok(fs.statSync(file).size > 400, name + " is too small");
});

const sw = fs.readFileSync(path.join(publicDir, "sw.js"), "utf8");
assert.ok(sw.indexOf('"/api/"') !== -1 || sw.indexOf("/api/") !== -1);
assert.ok(sw.indexOf("isApi") !== -1);
assert.ok(sw.indexOf("/offline.html") !== -1);
assert.ok(sw.indexOf("skipWaiting") !== -1);
assert.ok(sw.indexOf("/outlook-addin") !== -1, "do not intercept the Outlook add-in");

const pwaJs = fs.readFileSync(path.join(publicDir, "sd-pwa.js"), "utf8");
assert.ok(pwaJs.indexOf('rel="manifest"') !== -1);
assert.ok(pwaJs.indexOf("/sw.js") !== -1);
assert.ok(pwaJs.indexOf("serviceWorker") !== -1);
assert.ok(pwaJs.indexOf("beforeinstallprompt") !== -1);

const offline = fs.readFileSync(path.join(publicDir, "offline.html"), "utf8");
assert.ok(/offline/i.test(offline));
assert.ok(offline.indexOf("Studio Delta") !== -1);
assert.ok(fs.existsSync(path.join(publicDir, "facility-floor.png")), "Home twin needs facility-floor.png");
assert.ok(sw.indexOf("facility-floor.png") !== -1, "PWA precaches the facility floor image");
assert.ok(sw.indexOf("/delivery-run") !== -1, "PWA precaches the driver run for no-signal drops");
assert.ok(sw.indexOf('\n  "/",\n') !== -1 || sw.indexOf('"/",') !== -1, "PWA precaches Home so login is available offline");
assert.ok(sw.indexOf("sd-pwa-v20-delivery-login") !== -1, "new cache so phones pick up offline login");
assert.ok(sw.indexOf("sd-delivery-pod") !== -1, "background sync flushes saved PODs");
assert.ok(sw.indexOf("/vendor/leaflet/leaflet.js") !== -1);
assert.ok(fs.existsSync(path.join(publicDir, "vendor/leaflet/leaflet.js")));
assert.ok(fs.existsSync(path.join(publicDir, "delivery-offline.js")));

const indexJs = fs.readFileSync(path.join(__dirname, "index.js"), "utf8");
assert.ok(indexJs.indexOf("/manifest.webmanifest") !== -1);
assert.ok(indexJs.indexOf("/sw.js") !== -1);
assert.ok(indexJs.indexOf("Service-Worker-Allowed") !== -1);
assert.ok(indexJs.indexOf("/sd-pwa.js") !== -1);
assert.ok(indexJs.indexOf("/delivery-offline.js") !== -1);
assert.ok(indexJs.indexOf("/icons") !== -1);

const pages = [
  "index.html",
  "public/enquiries.html",
  "public/enquiries-dashboard.html",
  "public/enquiries-replies.html",
  "public/enquiries-bookings.html",
  "public/tasks.html",
  "public/orders.html",
  "public/orders-dashboard.html",
  "public/orders-schedule.html",
  "public/orders-delivery.html",
  "public/job-card.html",
  "public/job-card-existing.html",
  "public/dropdowns.html",
  "public/debtors.html",
  "public/users.html",
  "public/users-devices.html",
  "public/users-backup.html",
  "public/durations.html",
  "public/planning.html",
  "public/orders-to-order.html",
  "public/inventory.html",
  "public/orders-glass-rates.html",
  "public/orders-products.html",
  "public/marketing-dashboard.html",
  "public/orders-cost.html",
  "public/orders-correct.html"
];
pages.forEach((rel) => {
  const html = fs.readFileSync(path.join(root, rel), "utf8");
  assert.ok(html.indexOf("/sd-pwa.js?v=pwa-stg1") !== -1 || html.indexOf("/sd-pwa.js") !== -1, rel + " must boot the PWA");
});

const outlook = fs.readFileSync(path.join(publicDir, "outlook-addin/taskpane.html"), "utf8");
assert.ok(outlook.indexOf("sd-pwa.js") === -1, "Outlook add-in must not register the shop PWA");

console.log("pwa.test.js ok");
