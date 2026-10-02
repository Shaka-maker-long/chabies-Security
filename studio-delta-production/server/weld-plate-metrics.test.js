"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");
const crypto = require("crypto");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-weld-plate-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const { callShopFunction, ALLOWED } = require("./gas");

assert.ok(ALLOWED.has("getWeldPlateOverlapMetrics"));

initWorkbook();
persistWorkbook();

db.upsertOrder({
  order_number: "WP-DELAY",
  status: "Welding",
  product: "Ella Cabinet",
  client_name: "Delay Client"
});
db.upsertOrder({
  order_number: "WP-OK",
  status: "Ready for Grinding",
  product: "Air Chair",
  client_name: "Ok Client"
});

const book = getBook();
const logs = book.getSheetByName("Production_Log");
assert.ok(logs);

function addLog(order, worker, role, task, start, end, meta) {
  logs.appendRow([
    crypto.randomUUID(),
    order,
    worker,
    role,
    task,
    start,
    end || "",
    "Complete",
    "",
    "",
    "",
    "",
    JSON.stringify(meta || { pauses: [], entryType: "production" })
  ]);
}

// Delay case: welding starts while plate cutting is still running
addLog(
  "WP-DELAY",
  "Willard",
  "Plate Cutting",
  "Plate Cutting",
  new Date("2026-09-01T07:00:00.000Z"),
  new Date("2026-09-01T11:00:00.000Z"),
  { pauses: [{ start: "2026-09-01T08:00:00.000Z", end: "2026-09-01T08:30:00.000Z", reason: "No materials" }], entryType: "production" }
);
addLog(
  "WP-DELAY",
  "Sipho",
  "Welding",
  "Welding",
  new Date("2026-09-01T09:00:00.000Z"),
  new Date("2026-09-01T12:00:00.000Z"),
  { pauses: [], entryType: "production" }
);

// OK case: plate finished before welding started
addLog(
  "WP-OK",
  "Willard",
  "Plate Cutting",
  "Plate Cutting",
  new Date("2026-09-02T07:00:00.000Z"),
  new Date("2026-09-02T09:00:00.000Z"),
  { pauses: [], entryType: "production" }
);
addLog(
  "WP-OK",
  "Sipho",
  "Welding",
  "Welding",
  new Date("2026-09-02T10:00:00.000Z"),
  new Date("2026-09-02T12:00:00.000Z"),
  { pauses: [], entryType: "production" }
);

persistWorkbook();

async function main() {
  const data = await callShopFunction("getWeldPlateOverlapMetrics", []);
  assert.ok(data);
  assert.ok(Array.isArray(data.rows));
  assert.strictEqual(data.orderCount, 2);
  assert.strictEqual(data.delayCount, 1);

  const delay = data.rows.find((r) => r.orderNum === "WP-DELAY");
  const ok = data.rows.find((r) => r.orderNum === "WP-OK");
  assert.ok(delay, "delay order present");
  assert.ok(ok, "ok order present");
  assert.strictEqual(delay.potentialDelay, true);
  assert.strictEqual(ok.potentialDelay, false);
  assert.ok(delay.weldStartLabel);
  assert.ok(delay.plateStartLabel);
  assert.ok(!/HH:mm/.test(delay.weldStartLabel), "clock time must render, got " + delay.weldStartLabel);
  assert.ok(/\d{2}:\d{2}/.test(delay.weldStartLabel), delay.weldStartLabel);
  assert.ok(delay.weldActualHours > 0);
  assert.ok(delay.plateActualHours > 0);
  assert.ok(delay.platePauseMinutes > 0, "plate pause minutes recorded");
  assert.ok(delay.platePauseHours > 0, "plate pause hours recorded");
  assert.ok(delay.overlapHours > 0, "overlap while plate still running");
  assert.ok(/while plate cutting was still running|before plate cutting/i.test(delay.note));
  assert.strictEqual(delay.productName, "Ella Cabinet");
  assert.strictEqual(ok.productName, "Air Chair");

  const floor = fs.readFileSync(path.join(__dirname, "../index.html"), "utf8");
  assert.ok(floor.indexOf("tab-weld-plate") !== -1);
  assert.ok(floor.indexOf("Weld vs Plate") !== -1);
  assert.ok(floor.indexOf("getWeldPlateOverlapMetrics") !== -1);
  assert.ok(floor.indexOf("Weld pauses (h)") !== -1);
  assert.ok(floor.indexOf("Plate pauses (h)") !== -1);
  assert.ok(floor.indexOf("Weld pauses (min)") === -1);

  console.log("weld-plate-metrics.test.js ok");
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
