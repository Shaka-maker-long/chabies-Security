process.env.TZ = process.env.TZ || "Africa/Johannesburg";

const express = require("express");
const path = require("path");
const { initWorkbook, persistWorkbook, hasGoogleAuth, storageInfo, googleMigrateEnabled, maybeImportGoogleOnce } = require("./workbook-store");
const { migrateJsonOrdersToWorkbook, normalizeOrdersSheet, persistenceInfo } = require("./db");

initWorkbook();
migrateJsonOrdersToWorkbook();
normalizeOrdersSheet();

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
  const payload = {
    ok: true,
    tz: process.env.TZ,
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
app.get("/enquiries", (_req, res) => {
  res.sendFile(path.join(publicDir, "enquiries.html"));
});
app.get("/enquiries/dashboard", (_req, res) => {
  res.sendFile(path.join(publicDir, "enquiries-dashboard.html"));
});
app.get("/tasks", (_req, res) => {
  res.sendFile(path.join(publicDir, "tasks.html"));
});
app.get("/tasks/completed", (_req, res) => {
  res.sendFile(path.join(publicDir, "tasks.html"));
});
app.get("/schedule", (_req, res) => {
  res.sendFile(path.join(publicDir, "schedule.html"));
});
app.get("/dropdowns", (_req, res) => {
  res.sendFile(path.join(publicDir, "dropdowns.html"));
});
app.get("/debtors", (_req, res) => {
  res.sendFile(path.join(publicDir, "debtors.html"));
});
app.get("/users", (_req, res) => {
  res.sendFile(path.join(publicDir, "users.html"));
});
app.get("/durations", (_req, res) => {
  res.sendFile(path.join(publicDir, "durations.html"));
});
app.get("/gas-client.js", (_req, res) => {
  res.type("application/javascript").sendFile(path.join(publicDir, "gas-client.js"));
});
function noStore(res) {
  res.set("Cache-Control", "no-store, max-age=0");
}
app.get("/office-auth.js", (_req, res) => {
  noStore(res);
  res.type("application/javascript").sendFile(path.join(publicDir, "office-auth.js"));
});
app.get("/enquiry-process.js", (_req, res) => {
  noStore(res);
  res.type("application/javascript").sendFile(path.join(publicDir, "enquiry-process.js"));
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
try {
  require("./outlook-addin").mountOutlookAddin(app, publicDir);
} catch (e) {
  console.error("[boot] Outlook add-in failed", e && e.stack ? e.stack : e);
}
app.get("/", (_req, res) => {
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
  try {
    await Promise.race([
      maybeImportGoogleOnce(),
      new Promise((resolve) => setTimeout(resolve, 25000))
    ]);
  } catch (e) {
    console.error("[boot] google copy failed", e && e.message ? e.message : e);
  }
  try {
    const staff = require("./staff");
    staff.usersSheet();
    staff.seedLocalAdminIfEmpty();
  } catch (e) {}
  server = app.listen(PORT, "0.0.0.0", () => {
    console.log("Studio Delta production listening on " + PORT + " (" + process.env.TZ + ")");
    try {
      const info = persistenceInfo();
      if (info.warning) console.error("[persist]", info.warning);
      else console.log("[persist] dataDir", info.dataDir, "enquiries", info.enquiryCount, "workbook", info.workbookExists);
      try {
        console.log("[boot] users", require("./staff").listUsers().length);
      } catch (err) {}
    } catch (e) {
      console.error("[persist] could not read storage info", e && e.message ? e.message : e);
    }
  });
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
  serialize(async () => {
    try {
      if (!fn) {
        res.status(400).json({ ok: false, error: "Missing fn" });
        return;
      }
      const result = await loadFloor()(fn, args);
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
