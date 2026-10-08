process.env.TZ = process.env.TZ || "Africa/Johannesburg";

const express = require("express");
const path = require("path");
const { initWorkbook, persistWorkbook, hasGoogleAuth, storageInfo, googleMigrateEnabled, maybeImportGoogleOnce } = require("./workbook-store");
const { migrateJsonOrdersToWorkbook, normalizeOrdersSheet, persistenceInfo } = require("./db");

initWorkbook();

const app = express();
app.disable("x-powered-by");
app.use(express.json({ limit: "80mb" }));
const zlib = require("zlib");
app.use((req, res, next) => {
  const origJson = res.json.bind(res);
  res.json = function (body) {
    const accept = String(req.headers["accept-encoding"] || "");
    if (accept.indexOf("gzip") === -1) return origJson(body);
    let raw;
    try { raw = Buffer.from(JSON.stringify(body)); } catch (e) { return origJson(body); }
    if (raw.length < 16384) return origJson(body);
    const gz = zlib.gzipSync(raw);
    res.set("Content-Type", "application/json; charset=utf-8");
    res.set("Content-Encoding", "gzip");
    res.set("Vary", "Accept-Encoding");
    return res.send(gz);
  };
  next();
});

function health(_req, res) {
  let persist = {};
  try { persist = persistenceInfo(); } catch (e) {
    try { persist = storageInfo(); } catch (err) { persist = { error: String(err && err.message || err) }; }
  }
  let appEnvInfo = { appEnv: "production", isStaging: false, stagingBanner: null };
  try { appEnvInfo = require("./app-env"); } catch (e) {}
  const staging = !!(appEnvInfo.isStaging && appEnvInfo.isStaging());
  const payload = {
    ok: true,
    tz: process.env.TZ,
    appEnv: appEnvInfo.appEnv ? appEnvInfo.appEnv() : (staging ? "staging" : "production"),
    isStaging: staging,
    stagingBanner: staging && appEnvInfo.stagingBannerText ? appEnvInfo.stagingBannerText() : null,
    db: persist.usingEphemeralDisk ? "ephemeral" : "railway",
    database: persist.database || "SQLite on the Railway volume (not Google Sheets, not Postgres)",
    dataDir: persist.dataDir || null,
    volumeMount: persist.volumeMount || null,
    usingEphemeralDisk: !!persist.usingEphemeralDisk,
    warning: persist.warning || null,
    officeDb: persist.officeDb || null,
    officeDbExists: !!persist.officeDbExists,
    enquiryCount: persist.enquiryCount != null ? persist.enquiryCount : null,
    sqliteUsers: persist.sqliteUsers != null ? persist.sqliteUsers : null,
    sqliteOrders: persist.sqliteOrders != null ? persist.sqliteOrders : null,
    sqliteEnquiries: persist.sqliteEnquiries != null ? persist.sqliteEnquiries : null,
    sqliteSheets: persist.sqliteSheets != null ? persist.sqliteSheets : null,
    sqliteSheetRows: persist.sqliteSheetRows != null ? persist.sqliteSheetRows : null,
    sqliteDropdowns: persist.sqliteDropdowns != null ? persist.sqliteDropdowns : null,
    sqlitePayments: persist.sqlitePayments != null ? persist.sqlitePayments : null,
    sqliteOfficeScheduleRows: persist.sqliteOfficeScheduleRows != null ? persist.sqliteOfficeScheduleRows : null,
    sqliteSessions: persist.sqliteSessions != null ? persist.sqliteSessions : null,
    workbookExists: !!persist.workbookExists,
    googleMigrateAvailable: googleMigrateEnabled(),
    googleDriveOptional: hasGoogleAuth(),
    gmailLinked: !!String(process.env.GMAIL_SENDER || "").trim(),
    sheetsLive: false
  };
  try { Object.assign(payload, require("./backup").info()); } catch (e) {}
  res.status(200).json(payload);
}

app.get("/health", health);
app.head("/health", (_req, res) => res.status(200).end());
app.get("/healthz", health);

const publicDir = path.join(__dirname, "..", "public");
const indexHtml = path.join(__dirname, "..", "index.html");

app.get("/orders", (_req, res) => {
  res.sendFile(path.join(publicDir, "orders.html"));
});
app.get("/orders/dashboard", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "orders-dashboard.html"));
});
app.get("/orders/schedule", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "orders-schedule.html"));
});
app.get("/orders/delivery", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "orders-delivery.html"));
});
app.get("/orders/job-card/existing", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "job-card-existing.html"));
});
app.get("/orders/job-card", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "job-card.html"));
});
app.get("/orders/to-order", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "orders-to-order.html"));
});
app.get("/orders/consumables", (_req, res) => {
  res.redirect(302, "/inventory");
});
app.get("/inventory/purchases", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "inventory.html"));
});
app.get("/inventory/steel", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "inventory.html"));
});
app.get("/inventory/glass", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "inventory.html"));
});
app.get("/inventory", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "inventory.html"));
});
app.get("/orders/glass-rates", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "orders-glass-rates.html"));
});
app.get("/orders/products", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "orders-products.html"));
});
app.get("/orders/cost", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "orders-cost.html"));
});
app.get("/orders/correct", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "orders-correct.html"));
});
app.get("/orders/paint-shop", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "orders-paint-shop.html"));
});
app.get("/reworks", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "reworks.html"));
});
app.get("/orders/from-enquiry", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "orders-from-enquiry.html"));
});
app.get("/orders/new", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "orders-from-enquiry.html"));
});
app.get("/orders/onboard", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "orders-onboard.html"));
});
app.get("/orders/paste", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "orders-paste.html"));
});
app.get("/enquiries", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "enquiries.html"));
});
app.get("/enquiries/dashboard", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "enquiries-dashboard.html"));
});
app.get("/marketing", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "marketing-dashboard.html"));
});
app.get("/marketing/dashboard", (_req, res) => {
  noStore(res);
  res.redirect(302, "/marketing");
});
app.get("/enquiries/replies", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "enquiries-replies.html"));
});
app.get("/enquiries/bookings", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "enquiries-bookings.html"));
});
app.get("/tasks", (_req, res) => {
  res.sendFile(path.join(publicDir, "tasks.html"));
});
app.get("/tasks/team", (_req, res) => {
  res.sendFile(path.join(publicDir, "tasks.html"));
});
app.get("/tasks/completed", (_req, res) => {
  res.sendFile(path.join(publicDir, "tasks.html"));
});
app.get("/schedule", (_req, res) => {
  res.redirect(302, "/orders/schedule");
});
app.get("/dropdowns", (_req, res) => {
  res.sendFile(path.join(publicDir, "dropdowns.html"));
});
app.get("/debtors/history", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "debtors.html"));
});
app.get("/debtors", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "debtors.html"));
});
app.get("/users", (_req, res) => {
  res.sendFile(path.join(publicDir, "users.html"));
});
app.get("/users/devices", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "users-devices.html"));
});
app.get("/users/backup", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "users-backup.html"));
});
app.get("/durations", (_req, res) => {
  res.sendFile(path.join(publicDir, "durations.html"));
});
app.get("/planning", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "planning.html"));
});
app.get("/gas-client.js", (_req, res) => {
  res.type("application/javascript").sendFile(path.join(publicDir, "gas-client.js"));
});
function noStore(res) {
  res.set("Cache-Control", "no-store, max-age=0");
}
app.get("/api/powder-lists/:id/pdf", (req, res) => {
  noStore(res);
  try {
    const file = require("./powder-list").readPdf(req.params.id);
    if (!file) {
      res.status(404).json({ ok: false, error: "No powder coating list for that id." });
      return;
    }
    const download = String((req.query && req.query.download) || "") === "1";
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", (download ? "attachment" : "inline") + "; filename=\"" + file.filename + "\"");
    res.send(file.buffer);
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message || String(e) });
  }
});
app.get("/api/qc-pdfs/:id/pdf", (req, res) => {
  noStore(res);
  try {
    const file = require("./qc-pdf").readPdf(req.params.id);
    if (!file) {
      res.status(404).json({ ok: false, error: "No QC PDF for that report." });
      return;
    }
    const download = String((req.query && req.query.download) || "") === "1";
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", (download ? "attachment" : "inline") + "; filename=\"" + file.filename + "\"");
    res.send(file.buffer);
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message || String(e) });
  }
});
app.get("/delivery-run", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "delivery-run.html"));
});
app.get("/delivery-forms", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "delivery-forms.html"));
});
app.get("/delivery-forms/cost", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "delivery-forms.html"));
});
app.get("/driver-tracker", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "driver-tracker.html"));
});
function shopProfile(req, res) {
  const staff = require("./staff");
  const profile = staff.readSession(req);
  if (!profile) {
    res.status(401).json({ ok: false, error: "Log in first." });
    return null;
  }
  return profile;
}
app.get("/api/delivery/run", (req, res) => {
  serialize(async () => {
    try {
      const profile = shopProfile(req, res);
      if (!profile) return;
      const delivery = require("./delivery-pod");
      if (!delivery.canSubmitPod(profile) && !delivery.canLoadTruck(profile)) {
        res.status(403).json({ ok: false, error: "Delivery is for the driver, QC, or office." });
        return;
      }
      res.json({
        ok: true,
        driver: profile.name,
        isDriver: delivery.isDriverProfile(profile),
        canLoad: delivery.canLoadTruck(profile),
        orders: delivery.listLoaded(),
        ready: delivery.canLoadTruck(profile) ? delivery.listReadyToLoad() : []
      });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
});
app.get("/api/delivery/route", (req, res) => {
  serialize(async () => {
    try {
      const profile = shopProfile(req, res);
      if (!profile) return;
      const delivery = require("./delivery-pod");
      if (!delivery.canSubmitPod(profile) && !delivery.canLoadTruck(profile)) {
        res.status(403).json({ ok: false, error: "Delivery is for the driver, QC, or office." });
        return;
      }
      const origin = (req.query && req.query.lat != null && req.query.lng != null)
        ? { lat: Number(req.query.lat), lng: Number(req.query.lng) }
        : null;
      const route = await delivery.buildRoute({ origin: origin && Number.isFinite(origin.lat) ? origin : null });
      res.json({ ok: true, route: route });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
});
app.post("/api/delivery/load", (req, res) => {
  serialize(async () => {
    try {
      const profile = shopProfile(req, res);
      if (!profile) return;
      const delivery = require("./delivery-pod");
      if (!delivery.canLoadTruck(profile)) {
        res.status(403).json({ ok: false, error: "Only QC, the driver, or office can load the truck." });
        return;
      }
      const body = req.body || {};
      const numbers = body.order_numbers || body.orders || body.order_number;
      const result = delivery.loadOnTruck(numbers, profile.name);
      try { require("./gas").clearShopCache(); } catch (e) {}
      res.json(Object.assign({ ok: true }, result));
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
});
app.post("/api/delivery/pod", (req, res) => {
  serialize(async () => {
    try {
      const profile = shopProfile(req, res);
      if (!profile) return;
      const delivery = require("./delivery-pod");
      if (!delivery.canSubmitPod(profile)) {
        res.status(403).json({ ok: false, error: "Only the driver, QC, or office can submit a delivery form." });
        return;
      }
      const result = await delivery.submitPod(req.body || {}, profile.name);
      try { require("./gas").clearShopCache(); } catch (e) {}
      res.json(Object.assign({ ok: true }, result));
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
});
app.get("/api/delivery/forms", (req, res) => {
  serialize(async () => {
    try {
      const profile = shopProfile(req, res);
      if (!profile) return;
      const delivery = require("./delivery-pod");
      if (!profile.canSeeOffice && !profile.isAdmin && !delivery.canSubmitPod(profile)) {
        res.status(403).json({ ok: false, error: "Delivery forms are for office, QC, and the driver." });
        return;
      }
      res.json({ ok: true, rows: delivery.listForms() });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
});
app.delete("/api/delivery/forms/:id", (req, res) => {
  serialize(async () => {
    try {
      const profile = shopProfile(req, res);
      if (!profile) return;
      const delivery = require("./delivery-pod");
      if (!profile.canSeeOffice && !profile.isAdmin) {
        res.status(403).json({ ok: false, error: "Only office can delete delivery forms." });
        return;
      }
      const result = delivery.deleteForm(req.params.id);
      res.json({ ok: true, ...result });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
});
app.get("/api/delivery/cost", (req, res) => {
  serialize(async () => {
    try {
      const profile = shopProfile(req, res);
      if (!profile) return;
      const delivery = require("./delivery-pod");
      if (!profile.canSeeOffice && !profile.isAdmin && !delivery.canLoadTruck(profile)) {
        res.status(403).json({ ok: false, error: "Delivery cost is for office and QC." });
        return;
      }
      const day = String((req.query && req.query.day) || "").trim()
        || delivery.deliveryDayKey(new Date().toISOString());
      const rate = req.query && req.query.rate != null && req.query.rate !== ""
        ? Number(req.query.rate)
        : undefined;
      const cost = await delivery.deliveryCostForDay(day, { rate: rate });
      res.json({ ok: true, cost: cost });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
});
app.post("/api/delivery/location", (req, res) => {
  serialize(async () => {
    try {
      const profile = shopProfile(req, res);
      if (!profile) return;
      const delivery = require("./delivery-pod");
      if (!delivery.canSubmitPod(profile)) {
        res.status(403).json({ ok: false, error: "Only the driver, QC, or office can share a delivery location." });
        return;
      }
      const track = require("./delivery-track");
      const row = track.saveLocation(profile.name, req.body || {}, req);
      res.json({ ok: true, location: row });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
});
app.get("/api/delivery/locations", (req, res) => {
  serialize(async () => {
    try {
      const profile = shopProfile(req, res);
      if (!profile) return;
      const delivery = require("./delivery-pod");
      if (!profile.canSeeOffice && !profile.isAdmin && !delivery.canLoadTruck(profile)) {
        res.status(403).json({ ok: false, error: "Driver tracker is for office and QC." });
        return;
      }
      const track = require("./delivery-track");
      res.json({
        ok: true,
        factory: track.FACTORY,
        drivers: track.listLocations()
      });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
});
app.get("/api/delivery/tracker-status", (req, res) => {
  serialize(async () => {
    try {
      const profile = shopProfile(req, res);
      if (!profile) return;
      const track = require("./delivery-track");
      res.json({ ok: true, ...track.trackerStatus(profile, req) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
});
app.get("/api/delivery-forms/:id/pdf", (req, res) => {
  noStore(res);
  serialize(async () => {
    try {
      const profile = shopProfile(req, res);
      if (!profile) return;
      const file = require("./delivery-pod").readPdf(req.params.id);
      if (!file) {
        res.status(404).json({ ok: false, error: "No delivery form for that id." });
        return;
      }
      const download = String((req.query && req.query.download) || "") === "1";
      res.setHeader("Content-Type", "application/pdf");
      res.setHeader("Content-Disposition", (download ? "attachment" : "inline") + "; filename=\"" + file.filename + "\"");
      res.send(file.buffer);
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });
});
app.get("/office-auth.js", (_req, res) => {
  noStore(res);
  res.type("application/javascript").sendFile(path.join(publicDir, "office-auth.js"));
});
app.get("/enquiry-process.js", (_req, res) => {
  noStore(res);
  res.type("application/javascript").sendFile(path.join(publicDir, "enquiry-process.js"));
});
app.get("/order-paste-ui.js", (_req, res) => {
  noStore(res);
  res.type("application/javascript").sendFile(path.join(publicDir, "order-paste-ui.js"));
});
app.get("/office-shell.css", (_req, res) => {
  noStore(res);
  res.type("text/css").sendFile(path.join(publicDir, "office-shell.css"));
});
app.get("/sd-brand.css", (_req, res) => {
  noStore(res);
  res.type("text/css").sendFile(path.join(publicDir, "sd-brand.css"));
});
app.get("/sd-splash.js", (_req, res) => {
  noStore(res);
  res.type("application/javascript").sendFile(path.join(publicDir, "sd-splash.js"));
});
app.get("/sd-pwa.js", (_req, res) => {
  noStore(res);
  res.type("application/javascript").sendFile(path.join(publicDir, "sd-pwa.js"));
});
app.get("/delivery-offline.js", (_req, res) => {
  noStore(res);
  res.type("application/javascript").sendFile(path.join(publicDir, "delivery-offline.js"));
});
app.get("/ar-measure.js", (_req, res) => {
  res.type("application/javascript").sendFile(path.join(publicDir, "ar-measure.js"));
});
app.get("/manifest.webmanifest", (_req, res) => {
  res.set("Cache-Control", "no-store, max-age=0");
  res.type("application/manifest+json").sendFile(path.join(publicDir, "manifest.webmanifest"));
});
app.get("/sw.js", (_req, res) => {
  res.set("Cache-Control", "no-store, max-age=0");
  res.set("Service-Worker-Allowed", "/");
  res.type("application/javascript").sendFile(path.join(publicDir, "sw.js"));
});
app.get("/offline.html", (_req, res) => {
  noStore(res);
  res.sendFile(path.join(publicDir, "offline.html"));
});
app.use("/icons", express.static(path.join(publicDir, "icons"), { maxAge: "7d" }));
app.use("/vendor", express.static(path.join(publicDir, "vendor"), { maxAge: "7d" }));
app.get("/facility-floor.png", (_req, res) => {
  noStore(res);
  res.type("image/png").sendFile(path.join(publicDir, "facility-floor.png"));
});
try {
  require("./outlook-addin").mountOutlookAddin(app, publicDir);
} catch (e) {
  console.error("[boot] Outlook add-in failed", e && e.stack ? e.stack : e);
}
app.get("/", (_req, res) => {
  noStore(res);
  res.sendFile(indexHtml);
});

try {
  const { mountOffice } = require("./office");
  mountOffice(app);
} catch (e) {
  console.error("[boot] office pages failed", e && e.stack ? e.stack : e);
}

const PORT = Number(process.env.PORT) || 8080;
let server;

async function boot() {
  // Listen first so Railway healthchecks pass while migrations finish.
  try {
    const staff = require("./staff");
    staff.usersSheet();
    staff.seedLocalAdminIfEmpty();
  } catch (e) {}
  server = app.listen(PORT, "0.0.0.0", () => {
    let envLabel = "production";
    try { envLabel = require("./app-env").appEnv(); } catch (e) {}
    console.log("Studio Delta " + envLabel + " listening on " + PORT + " (" + process.env.TZ + ")");
  });

  try {
    migrateJsonOrdersToWorkbook();
  } catch (e) {
    console.error("[boot] migrate orders failed", e && e.message ? e.message : e);
  }
  try {
    normalizeOrdersSheet();
  } catch (e) {
    console.error("[boot] normalize orders failed", e && e.message ? e.message : e);
  }
  try {
    await Promise.race([
      maybeImportGoogleOnce(),
      new Promise((resolve) => setTimeout(resolve, 25000))
    ]);
  } catch (e) {
    console.error("[boot] google copy failed", e && e.message ? e.message : e);
  }
  try {
    const info = persistenceInfo();
    if (info.warning) console.error("[persist]", info.warning);
    else console.log("[persist] dataDir", info.dataDir, "enquiries", info.enquiryCount, "workbook", info.workbookExists);
    try {
      console.log("[boot] users", require("./staff").listUsers().length);
    } catch (err) {}
    try {
      require("./enquiry-pipeline").syncDrawingQueue();
    } catch (err) {
      console.error("[boot] drawing queue", err && err.message ? err.message : err);
    }
  } catch (e) {
    console.error("[persist] could not read storage info", e && e.message ? e.message : e);
  }
}
boot();

let chain = Promise.resolve();
function serialize(work) {
  const run = chain.then(work, work);
  chain = run.catch(() => {});
  return run;
}

let callShopFunction = null;

function loadFloor() {
  if (callShopFunction) return callShopFunction;
  callShopFunction = require("./gas").callShopFunction;
  return callShopFunction;
}

app.post("/api/run", (req, res) => {
  const fn = req.body && req.body.fn;
  const args = (req.body && req.body.args) || [];
  const PUBLIC_RUN = new Set(["verifyGlobalLogin", "verifyLogin", "getUsersAndRoles"]);
  serialize(async () => {
    try {
      if (!fn) {
        res.status(400).json({ ok: false, error: "Missing fn" });
        return;
      }
      const staff = require("./staff");
      const auditLog = require("./audit-log");
      const loginThrottle = require("./login-throttle");
      if (!PUBLIC_RUN.has(fn)) {
        const profile = staff.readSession(req);
        if (!profile) {
          res.status(401).json({ ok: false, error: "Log in first." });
          return;
        }
        req.shop = profile;
      }
      if (fn === "verifyGlobalLogin" || fn === "verifyLogin") {
        const name = fn === "verifyGlobalLogin" ? (args && args[0]) : (args && args[1]);
        const ip = auditLog.clientIp(req);
        const throttle = loginThrottle.check(name, ip);
        if (!throttle.ok) {
          auditLog.record("shop.login.blocked", { actor: name, ok: false, detail: "throttled", ip }, req);
          res.status(429).json({ ok: false, error: throttle.error });
          return;
        }
      }
      const result = await loadFloor()(fn, args);
      if (fn === "verifyGlobalLogin" && result && result.success) {
        const trustedDevices = require("./trusted-devices");
        const deviceId = (req.body && req.body.deviceId)
          || (args && args[2] && args[2].deviceId)
          || (req.headers && req.headers["x-sd-device-id"]);
        const deviceMeta = trustedDevices.metaFromReq(req, result.name, deviceId);
        if (req.body && req.body.bootstrapCode) deviceMeta.bootstrapCode = req.body.bootstrapCode;
        if (args && args[2] && args[2].bootstrapCode) deviceMeta.bootstrapCode = args[2].bootstrapCode;
        deviceMeta.canManageUsers = !!result.canManageUsers;
        const deviceCheck = trustedDevices.assertLoginAllowed(deviceMeta);
        if (!deviceCheck.ok) {
          auditLog.record("shop.login.device_pending", {
            actor: result.name,
            ok: false,
            detail: deviceCheck.error || "pending",
            deviceId: deviceMeta.deviceId
          }, req);
          res.json({
            ok: true,
            result: {
              success: false,
              pendingDevice: !!deviceCheck.pending,
              needsBootstrap: !!deviceCheck.needsBootstrap,
              bootstrapConfigured: trustedDevices.bootstrapCodeConfigured(),
              device: deviceCheck.device || null,
              error: deviceCheck.error || "This device is not approved yet."
            }
          });
          return;
        }
        loginThrottle.clear(result.name, auditLog.clientIp(req));
        const session = staff.createSession({
          name: result.name,
          access: result.access,
          role: result.role || result.jobTitle,
          jobTitle: result.jobTitle || result.role,
          isAdmin: result.isAdmin,
          isMarketing: result.isMarketing,
          canSeeOffice: result.canSeeOffice,
          canSeeDrawingDesk: result.canSeeDrawingDesk,
          canSeeDebtors: result.canSeeDebtors,
          canManageUsers: result.canManageUsers,
          canEditMarketingFields: result.canEditMarketingFields,
          tasks: result.tasks || []
        }, { deviceId: deviceCheck.device && deviceCheck.device.id });
        result.token = session.token;
        result.device = deviceCheck.device || null;
        result.deviceBootstrapped = !!deviceCheck.bootstrapped;
        const secure = String(process.env.COOKIE_SECURE || "").trim() === "0" ? ""
          : ((String((req.headers && req.headers["x-forwarded-proto"]) || "").indexOf("https") !== -1
            || String(process.env.RAILWAY_ENVIRONMENT || "").trim()
            || String(process.env.NODE_ENV || "").toLowerCase() === "production") ? "; Secure" : "");
        const maxAge = Math.floor((staff.SESSION_MAX_AGE_MS || (14 * 24 * 60 * 60 * 1000)) / 1000);
        res.setHeader(
          "Set-Cookie",
          "sd_session=" + encodeURIComponent(session.token) + "; Path=/; HttpOnly; SameSite=Lax; Max-Age=" + maxAge + secure
        );
        auditLog.record("shop.login.ok", {
          actor: result.name,
          ok: true,
          deviceId: session.deviceId,
          detail: deviceCheck.bootstrapped ? "bootstrap" : "approved"
        }, req);
      } else if (fn === "verifyGlobalLogin" && result && !result.success) {
        const name = args && args[0];
        loginThrottle.fail(name, auditLog.clientIp(req));
        auditLog.record("shop.login.failed", { actor: name, ok: false, detail: result.error || "" }, req);
      }
      res.json({ ok: true, result });
    } catch (e) {
      const msg = (e && e.message) || String(e);
      console.error("[api/run]", fn, e && e.stack ? e.stack : e);
      if (!res.headersSent) {
        const quota = /quota exceeded/i.test(msg);
        res.status(quota ? 429 : 400).json({
          ok: false,
          error: quota
            ? "The shop is busy. Wait 60 seconds, then try again. Do not keep tapping."
            : msg
        });
      }
    }
  });
});

setTimeout(() => {
  try {
    const run = loadFloor();
    serialize(() => run("lazySetup", []).catch((e) => console.error("[lazySetup]", e.message || e)));
  } catch (e) {
    console.error("[boot] floor failed to load", e && e.stack ? e.stack : e);
  }
}, 2000);

const FIVE_MIN = 5 * 60 * 1000;
setInterval(() => {
  if (!callShopFunction) return;
  serialize(() =>
    callShopFunction("enforceShiftHours", []).catch((e) => console.error("[enforceShiftHours]", e.message || e))
  );
  serialize(() =>
    callShopFunction("checkIdleWorkers", []).catch((e) => console.error("[checkIdleWorkers]", e.message || e))
  );
  if (hasGoogleAuth()) {
    serialize(() =>
      callShopFunction("processPdfQueue", []).catch((e) => console.error("[processPdfQueue]", e.message || e))
    );
  }
  serialize(() => {
    try { require("./backup").tick(); } catch (e) {
      console.error("[backup]", e && e.message ? e.message : e);
    }
  });
}, FIVE_MIN);

setTimeout(() => {
  serialize(() => {
    try { require("./backup").tick(); } catch (e) {
      console.error("[backup]", e && e.message ? e.message : e);
    }
  });
}, 120000).unref();

function shutdown() {
  try { persistWorkbook(); } catch (e) {}
  try { require("./db").persist(); } catch (e) {}
  try { require("./staff").persistSessions(); } catch (e) {}
  if (server) server.close(() => process.exit(0));
  else process.exit(0);
  setTimeout(() => process.exit(0), 5000).unref();
}
process.on("SIGTERM", shutdown);
process.on("SIGINT", shutdown);
process.on("uncaughtException", (e) => {
  console.error("[uncaughtException]", e && e.stack ? e.stack : e);
});
process.on("unhandledRejection", (e) => {
  console.error("[unhandledRejection]", e && e.stack ? e.stack : e);
});
