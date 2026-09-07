const assert = require("assert");
const fs = require("fs");
const os = require("os");
const path = require("path");

const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-floor-"));
process.env.DATA_DIR = dataDir;
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const db = require("./db");
const { callShopFunction, hasGoogleAuth } = require("./gas");

assert.strictEqual(hasGoogleAuth(), false);

const users = db.ensureSheet("Users");
users.grid = [
  ["Name", "Role", "Password", "Tasks"],
  ["Sipho", "Welder Tagger", "1234", "Welding, Profile Cutting, Assembly"],
  ["Admin", "Admin", "admin", ""]
];
users.lastRow = 3;
users.lastCol = 4;

db.upsertOrder({
  order_number: "SD-WELD",
  status: "Ready for Welding",
  product: "Gate",
  client_name: "Test Client"
});
db.upsertOrder({
  order_number: "SD-CUT",
  status: "Ready for Steelwork",
  product: "Gate",
  client_name: "Test Client"
});
db.upsertOrder({
  order_number: "SD-ASM",
  status: "Ready for Assembly",
  product: "Gate",
  client_name: "Test Client"
});

function orderRow(orderNumber) {
  const sheet = db.ensureSheet("ORDERS");
  const headers = sheet.grid[0];
  const col = headers.findIndex((h) => String(h).toLowerCase().indexOf("order") >= 0);
  return sheet.grid.findIndex((row, i) => i > 0 && String(row[col]) === orderNumber) + 1;
}

const CONFIRM = { understood: true, highlights: [] };

async function main() {
  await callShopFunction("grantOvertime", ["Sipho", "", "Admin", "test"]);
  const login = await callShopFunction("verifyGlobalLogin", ["Sipho", "1234"]);
  assert.strictEqual(login.success, true, JSON.stringify(login));

  const weldRow = orderRow("SD-WELD");
  const started = await callShopFunction("startOrder", [weldRow, "Sipho", "Welding", [], "", false, null, CONFIRM]);
  assert.strictEqual(started.success, true, JSON.stringify(started));
  assert.ok(started.logId);

  const paused = await callShopFunction("workerPauseOrder", [weldRow, "SD-WELD", "Sipho", "No materials"]);
  assert.strictEqual(paused.success, true, JSON.stringify(paused));

  const resumed = await callShopFunction("workerResumeOrder", [weldRow, "SD-WELD", "Sipho"]);
  assert.strictEqual(resumed.success, true, JSON.stringify(resumed));

  const finished = await callShopFunction("finishOrder", [
    weldRow, started.logId, null, "", [], "Sipho", [], "SD-WELD", []
  ]);
  assert.ok(finished && (finished.success !== false), JSON.stringify(finished));

  const cutRow = orderRow("SD-CUT");
  const cutStart = await callShopFunction("startOrder", [cutRow, "Sipho", "Profile Cutting", [], "", false, null, CONFIRM]);
  assert.strictEqual(cutStart.success, true, JSON.stringify(cutStart));
  const cutFinish = await callShopFunction("finishOrder", [
    cutRow,
    cutStart.logId,
    null,
    "",
    [],
    "Sipho",
    [{ category: "Square tube", type: "25x25x2", size: "6m", isCustom: true }],
    "SD-CUT",
    []
  ]);
  assert.ok(cutFinish && cutFinish.success !== false, JSON.stringify(cutFinish));

  const asmRow = orderRow("SD-ASM");
  const asmStart = await callShopFunction("startOrder", [asmRow, "Sipho", "Assembly", [], "", false, null, CONFIRM]);
  assert.strictEqual(asmStart.success, true, JSON.stringify(asmStart));
  const asmFinish = await callShopFunction("finishOrder", [
    asmRow,
    asmStart.logId,
    null,
    "",
    [],
    "Sipho",
    [],
    "SD-ASM",
    [{ category: "Board", type: "12mm", size: "1 sheet", isCustom: true }]
  ]);
  assert.ok(asmFinish && asmFinish.success !== false, JSON.stringify(asmFinish));

  const saved = JSON.parse(fs.readFileSync(db.dbPath, "utf8"));
  assert.ok(saved.workbook && saved.workbook.sheets.ORDERS);
  assert.ok(saved.workbook.sheets.Production_Log.grid.length > 1);
  assert.ok(saved.workbook.sheets.Steel_Usage.grid.length > 1);
  assert.ok(saved.workbook.sheets.Backboard_Usage.grid.length > 1);

  const weld = saved.orders.find((o) => o.order_number === "SD-WELD");
  const cut = saved.orders.find((o) => o.order_number === "SD-CUT");
  const asm = saved.orders.find((o) => o.order_number === "SD-ASM");
  assert.ok(weld.work_logs && weld.work_logs.length, "weld logs missing");
  assert.ok(weld.work_logs[0].start, "weld start missing");
  assert.ok(weld.work_logs[0].end, "weld end missing");
  assert.ok(weld.work_logs[0].meta.indexOf("No materials") >= 0 || weld.work_logs[0].pause_reason, "pause not saved");
  assert.ok(cut.steel_usage && cut.steel_usage.length, "steel usage missing");
  assert.strictEqual(cut.steel_usage[0].type.indexOf("25x25x2") >= 0, true);
  assert.ok(asm.backboard_usage && asm.backboard_usage.length, "backboard usage missing");
  assert.ok(String(asm.status || "").length, "assembly status missing");

  const listed = db.listOrders();
  const listedWeld = listed.find((o) => o.order_number === "SD-WELD");
  assert.ok(listedWeld.work_logs.length);
  assert.ok(typeof listedWeld.duration_minutes === "number");

  console.log("floor-store.test.js ok");
}

main().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
