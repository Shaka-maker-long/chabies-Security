process.env.TZ = process.env.TZ || "Africa/Johannesburg";
process.env.SHEET_ID = process.env.SHEET_ID || "1pdvAFTIyd5sf8Wbf38MSd4cfk3mb3McPqJrYeM8SOYk";

const express = require("express");
const path = require("path");
const { callShopFunction } = require("./gas");
const { mountOffice } = require("./office");
require("./db");

const app = express();
app.disable("x-powered-by");
app.use(express.json({ limit: "50mb" }));

const publicDir = path.join(__dirname, "..", "public");
const indexHtml = path.join(__dirname, "..", "index.html");

let chain = Promise.resolve();
function serialize(work) {
  const run = chain.then(work, work);
  chain = run.catch(() => {});
  return run;
}

function hasGoogleAuth() {
  return !!(process.env.GOOGLE_SERVICE_ACCOUNT_JSON || process.env.GOOGLE_APPLICATION_CREDENTIALS);
}

app.get("/orders", (_req, res) => {
  res.sendFile(path.join(publicDir, "orders.html"));
});
app.get("/schedule", (_req, res) => {
  res.sendFile(path.join(publicDir, "schedule.html"));
});

mountOffice(app);

app.get("/health", (_req, res) => {
  res.json({
    ok: true,
    tz: process.env.TZ,
    sheetsConfigured: hasGoogleAuth() && !!process.env.SHEET_ID
  });
});

app.get("/gas-client.js", (_req, res) => {
  res.type("application/javascript").sendFile(path.join(publicDir, "gas-client.js"));
});

app.get("/", (_req, res) => {
  res.sendFile(indexHtml);
});

app.post("/api/run", (req, res) => {
  const fn = req.body && req.body.fn;
  const args = (req.body && req.body.args) || [];
  serialize(async () => {
    try {
      if (!fn) {
        res.status(400).json({ ok: false, error: "Missing fn" });
        return;
      }
      if (!hasGoogleAuth()) {
        res.status(503).json({
          ok: false,
          error: "Google credentials are not set. Add GOOGLE_SERVICE_ACCOUNT_JSON on Railway."
        });
        return;
      }
      const result = await callShopFunction(fn, args);
      res.json({ ok: true, result });
    } catch (e) {
      const msg = (e && e.message) || String(e);
      console.error("[api/run]", fn, e && e.stack ? e.stack : e);
      if (!res.headersSent) {
        const quota = /quota exceeded/i.test(msg);
        res.status(quota ? 429 : 400).json({
          ok: false,
          error: quota
            ? "Google Sheets is busy (too many reads this minute). Wait 60 seconds, then try again. Do not keep tapping."
            : msg
        });
      }
    }
  });
});

const PORT = Number(process.env.PORT) || 8080;
const server = app.listen(PORT, () => {
  console.log("Studio Delta production listening on " + PORT + " (" + process.env.TZ + ")");
  if (!hasGoogleAuth()) {
    console.warn("Google credentials are not set. Floor API calls will return 503 until Railway env is configured.");
    return;
  }
  serialize(() =>
    callShopFunction("lazySetup", []).catch((e) => console.error("[lazySetup]", e.message || e))
  );
});

const FIVE_MIN = 5 * 60 * 1000;
setInterval(() => {
  if (!hasGoogleAuth()) return;
  serialize(() =>
    callShopFunction("checkIdleWorkers", []).catch((e) => console.error("[checkIdleWorkers]", e.message || e))
  );
  serialize(() =>
    callShopFunction("processPdfQueue", []).catch((e) => console.error("[processPdfQueue]", e.message || e))
  );
}, FIVE_MIN);

function shutdown() {
  server.close(() => process.exit(0));
  setTimeout(() => process.exit(0), 5000).unref();
}
process.on("SIGTERM", shutdown);
process.on("SIGINT", shutdown);
