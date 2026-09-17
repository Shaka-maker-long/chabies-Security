const assert = require("assert");
const fs = require("fs");
const os = require("os");
const path = require("path");

const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-floor-"));
process.env.DATA_DIR = dataDir;
process.env.OFFICE_DB_PATH = path.join(dataDir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const { callShopFunction } = require("./gas");

initWorkbook();
const book = getBook();
book.getSheetByName("Users").appendRow([
  "Sipho", "Welder Tagger", "1234", "Welding, Profile Cutting, Plate Cutting, Assembly", "Production", "No"
]);
book.getSheetByName("Users").appendRow(["Thabo", "Welder", "1234", "Welding", "Production", "No"]);
book.getSheetByName("Users").appendRow(["Admin", "Manager", "admin", "", "Admin", "Yes"]);
book.getSheetByName("Users").appendRow(["Siya", "Production Manager", "siya", "", "Admin", "Yes"]);
book.getSheetByName("Users").appendRow(["QC Admin", "Admin", "qcadm", "Quality Control", "Admin", "Yes"]);
persistWorkbook();

const weldOrder = db.upsertOrder({
  order_number: "SD-WELD",
  status: "Ready for Welding",
  product: "Gate",
  client_name: "Test Client"
});
const cutOrder = db.upsertOrder({
  order_number: "SD-CUT",
  status: "Ready for Steelwork",
  product: "Gate",
  client_name: "Test Client"
});
const plateOrder = db.upsertOrder({
  order_number: "SD-PLATE",
  status: "Ready for Welding",
  product: "Gate",
  client_name: "Test Client"
});
const asmOrder = db.upsertOrder({
  order_number: "SD-ASM",
  status: "Ready for Assembly",
  product: "Gate",
  client_name: "Test Client"
});

const CONFIRM = { understood: true, highlights: [] };

async function main() {
  await callShopFunction("grantOvertime", ["Sipho", "", "Admin", "test"]);
  const login = await callShopFunction("verifyGlobalLogin", ["Sipho", "1234"]);
  assert.strictEqual(login.success, true, JSON.stringify(login));

  const weldRow = weldOrder.id;
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

  const cutRow = cutOrder.id;
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

  const asmRow = asmOrder.id;
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

  const logs = book.getSheetByName("Production_Log").getDataRange().getValues();
  assert.ok(logs.length > 1, "production log missing");
  const steelAtFinish = book.getSheetByName("Steel_Usage").getDataRange().getValues();
  assert.ok(steelAtFinish.length > 1, "steel usage missing");
  assert.ok(steelAtFinish.some((row, i) => i > 0 && String(row[4]).indexOf("25x25x2") >= 0), "cut steel not on sheet");
  const boards = book.getSheetByName("Backboard_Usage").getDataRange().getValues();
  assert.ok(boards.length > 1, "backboard usage missing");

  const listed = db.listOrders();
  const listedWeld = listed.find((o) => o.order_number === "SD-WELD");
  assert.ok(listedWeld, "weld order missing from list");
  assert.ok(String(listedWeld.status || "").length, "weld status missing");

  const completed = await callShopFunction("getMyCompletedWork", ["Sipho"]);
  const cutDone = (completed.items || []).find((i) => i.order === "SD-CUT" && i.process === "Profile Cutting");
  assert.ok(cutDone, "completed profile cutting missing");
  assert.ok(cutDone.canEditSteel, "profile cutting should allow steel edit");
  assert.ok(cutDone.steelUsage && cutDone.steelUsage.length, "completed work should show logged steel");
  assert.ok(String(cutDone.steelUsage[0].type).indexOf("25x25x2") >= 0, "logged profile type missing");

  const weldSteel = await callShopFunction("logWelderSteel", [
    "SD-WELD",
    "Sipho",
    { category: "Angle", type: "40x40x3", size: "1 m", isCustom: true }
  ]);
  assert.ok(weldSteel && weldSteel.success !== false, JSON.stringify(weldSteel));

  const emptyEdit = await callShopFunction("updateCompletedSteelUsage", ["Sipho", "SD-CUT", "Profile Cutting", []]);
  assert.strictEqual(emptyEdit.success, false, "empty steel replace must fail");

  const weldEdit = await callShopFunction("updateCompletedSteelUsage", [
    "Sipho", "SD-WELD", "Welding",
    [{ category: "Angle", type: "50x50x3", size: "2 m", isCustom: true }]
  ]);
  assert.strictEqual(weldEdit.success, false, "welding steel must not be edited here");

  const replaced = await callShopFunction("updateCompletedSteelUsage", [
    "Sipho",
    "SD-CUT",
    "Profile Cutting",
    [{ category: "Round tube", type: "38x2", size: "3 m", isCustom: true }]
  ]);
  assert.ok(replaced && replaced.success !== false, JSON.stringify(replaced));

  const completed2 = await callShopFunction("getMyCompletedWork", ["Sipho"]);
  const cutDone2 = (completed2.items || []).find((i) => i.order === "SD-CUT" && i.process === "Profile Cutting");
  assert.strictEqual((cutDone2.steelUsage || []).length, 1, "old profile steel should be replaced");
  assert.ok(String(cutDone2.steelUsage[0].type).indexOf("38x2") >= 0, "updated profile type missing");
  assert.ok(String(cutDone2.steelUsage[0].type).indexOf("25x25x2") === -1, "old profile steel still present");

  const plateRow = plateOrder.id;
  const plateStart = await callShopFunction("startOrder", [plateRow, "Sipho", "Plate Cutting", [], "", false, null, CONFIRM]);
  assert.strictEqual(plateStart.success, true, JSON.stringify(plateStart));
  const plateFinish = await callShopFunction("finishOrder", [
    plateRow,
    plateStart.logId,
    null,
    "",
    [],
    "Sipho",
    [{ category: "Plate", type: "4.5mm", size: "1.2 m²", isCustom: true }],
    "SD-PLATE",
    []
  ]);
  assert.ok(plateFinish && plateFinish.success !== false, JSON.stringify(plateFinish));

  const completedPlate = await callShopFunction("getMyCompletedWork", ["Sipho"]);
  const plateDone = (completedPlate.items || []).find((i) => i.order === "SD-PLATE" && i.process === "Plate Cutting");
  assert.ok(plateDone, "completed plate cutting missing");
  assert.ok(plateDone.steelUsage && plateDone.steelUsage.length, "plate steel missing on completed work");
  assert.ok(String(plateDone.steelUsage[0].type).indexOf("4.5mm") >= 0);

  const plateReplaced = await callShopFunction("updateCompletedSteelUsage", [
    "Sipho",
    "SD-PLATE",
    "Plate Cutting",
    [{ category: "Plate", type: "6mm", size: "0.8 m²", isCustom: true }]
  ]);
  assert.ok(plateReplaced && plateReplaced.success !== false, JSON.stringify(plateReplaced));

  const completedPlate2 = await callShopFunction("getMyCompletedWork", ["Sipho"]);
  const plateDone2 = (completedPlate2.items || []).find((i) => i.order === "SD-PLATE" && i.process === "Plate Cutting");
  assert.strictEqual((plateDone2.steelUsage || []).length, 1);
  assert.ok(String(plateDone2.steelUsage[0].type).indexOf("6mm") >= 0);
  assert.ok(String(plateDone2.steelUsage[0].type).indexOf("4.5mm") === -1);
  assert.strictEqual(plateDone2.worker, "Sipho");

  const siphoOnly = await callShopFunction("getMyCompletedWork", ["Sipho"]);
  assert.ok((siphoOnly.items || []).every((i) => i.worker === "Sipho"), "floor worker still sees only own completed");
  assert.ok(!siphoOnly.viewAll, "own completed list is not the full shop");

  const siphoWeldBoard = await callShopFunction("getMyCompletedWork", ["Sipho", "Welding"]);
  assert.ok(siphoWeldBoard.viewAll, "anyone on a task board sees every completion time for that task");
  assert.ok((siphoWeldBoard.items || []).some((i) => i.order === "SD-WELD"), JSON.stringify(siphoWeldBoard));
  assert.ok(!(siphoWeldBoard.items || []).some((i) => i.order === "SD-CUT" || i.order === "SD-PLATE" || i.order === "SD-ASM"), "welding board hides other processes");

  const blockedSteel = await callShopFunction("updateCompletedSteelUsage", [
    "Sipho",
    "SD-CUT",
    "Profile Cutting",
    [{ category: "Round tube", type: "hack", size: "1 m", isCustom: true }],
    "Thabo"
  ]);
  assert.strictEqual(blockedSteel.success, false, JSON.stringify(blockedSteel));
  assert.ok(/admin/i.test(String(blockedSteel.error || "")), JSON.stringify(blockedSteel));

  const adminSteel = await callShopFunction("updateCompletedSteelUsage", [
    "Sipho",
    "SD-CUT",
    "Profile Cutting",
    [{ category: "Round tube", type: "38x2-admin", size: "3 m", isCustom: true }],
    "Admin"
  ]);
  assert.ok(adminSteel && adminSteel.success !== false, JSON.stringify(adminSteel));

  const dayWork = await callShopFunction("getWorkerDayWork", ["Sipho"]);
  assert.strictEqual(dayWork.worker, "Sipho");
  assert.ok(dayWork.shift && dayWork.shift.id, "day work includes facility shift");
  assert.ok((dayWork.items || []).some((i) => i.order === "SD-WELD" && i.process === "Welding"), JSON.stringify(dayWork));
  assert.ok((dayWork.items || []).some((i) => i.order === "SD-CUT"), JSON.stringify(dayWork));
  assert.ok(typeof dayWork.totalMinutes === "number");
  const adminDay = await callShopFunction("getWorkerDayWork", ["Admin"]);
  assert.ok(!(adminDay.items || []).some((i) => i.order === "SD-WELD"), "manager click must not use overseer completed list");
  const siyaDay = await callShopFunction("getWorkerDayWork", ["Siya"]);
  assert.ok(!(siyaDay.items || []).some((i) => i.order === "SD-WELD"), "production manager click must not use overseer completed list");

  const managerAll = await callShopFunction("getMyCompletedWork", ["Admin"]);
  assert.ok(managerAll.oversees, "manager oversees shop completed work");
  assert.ok((managerAll.items || []).some((i) => i.order === "SD-WELD" && i.worker === "Sipho"), "manager sees Sipho weld");
  assert.ok((managerAll.items || []).some((i) => i.order === "SD-CUT"), "manager sees cutting completed");

  const managerWeld = await callShopFunction("getMyCompletedWork", ["Admin", "Welding"]);
  assert.ok((managerWeld.items || []).some((i) => i.order === "SD-WELD"), "manager welding board shows weld");
  assert.ok(!(managerWeld.items || []).some((i) => i.order === "SD-CUT" || i.order === "SD-PLATE" || i.order === "SD-ASM"), "manager welding board hides other processes");

  const siyaWeld = await callShopFunction("getMyCompletedWork", ["Siya", "Welding"]);
  assert.ok(siyaWeld.oversees, "production manager oversees shop completed work");
  assert.ok((siyaWeld.items || []).some((i) => i.order === "SD-WELD" && i.worker === "Sipho"), "production manager sees Sipho weld");

  const waitingDraw = db.upsertOrder({
    order_number: "SD-DRAW",
    status: "Waiting for drawing",
    product: "Gate",
    client_name: "Test Client"
  });
  const blockedDraw = await callShopFunction("startOrder", [waitingDraw.id, "Sipho", "Profile Cutting", [], "", false, null, CONFIRM]);
  assert.strictEqual(blockedDraw.success, false, JSON.stringify(blockedDraw));
  assert.ok(/drawing/i.test(String(blockedDraw.message || "")), JSON.stringify(blockedDraw));

  const adminLogin = await callShopFunction("verifyGlobalLogin", ["Admin", "admin"]);
  assert.strictEqual(adminLogin.success, true, JSON.stringify(adminLogin));
  assert.strictEqual(adminLogin.isAdmin, true);
  assert.ok((adminLogin.tasks || []).indexOf("Quality Control") === -1, "look-only admin must not inherit every floor task");

  const qcAdminLogin = await callShopFunction("verifyGlobalLogin", ["QC Admin", "qcadm"]);
  assert.strictEqual(qcAdminLogin.success, true, JSON.stringify(qcAdminLogin));
  assert.strictEqual(qcAdminLogin.isAdmin, true);
  assert.ok(qcAdminLogin.tasks.indexOf("Quality Control") !== -1, JSON.stringify(qcAdminLogin));

  await callShopFunction("grantOvertime", ["QC Admin", "", "Admin", "test"]);
  const qcOrder = db.upsertOrder({
    order_number: "SD-QC-ADMIN",
    status: "Ready for Final QC",
    product: "Gate",
    client_name: "Test Client"
  });
  const lookOnlyQc = await callShopFunction("startOrder", [qcOrder.id, "Admin", "Quality Control", [], "", false, null, CONFIRM]);
  assert.strictEqual(lookOnlyQc.success, false, JSON.stringify(lookOnlyQc));
  assert.ok(/not assigned/i.test(String(lookOnlyQc.message || lookOnlyQc.error || "")), JSON.stringify(lookOnlyQc));

  const qcAdminStart = await callShopFunction("startOrder", [qcOrder.id, "QC Admin", "Quality Control", [], "", false, null, CONFIRM]);
  assert.strictEqual(qcAdminStart.success, true, JSON.stringify(qcAdminStart));

  const extraWeld = db.upsertOrder({
    order_number: "SD-WELD-ADMIN",
    status: "Ready for Welding",
    product: "Gate",
    client_name: "Test Client"
  });
  const qcAdminWeld = await callShopFunction("startOrder", [extraWeld.id, "QC Admin", "Welding", [], "", false, null, CONFIRM]);
  assert.strictEqual(qcAdminWeld.success, false, JSON.stringify(qcAdminWeld));
  assert.ok(/not assigned/i.test(String(qcAdminWeld.message || qcAdminWeld.error || "")), JSON.stringify(qcAdminWeld));

  const steelGrid = book.getSheetByName("Steel_Usage").getDataRange().getValues();
  const weldRows = steelGrid.filter((row, i) => i > 0 && String(row[1]) === "SD-WELD");
  assert.ok(weldRows.length >= 1, "welder steel must stay after profile edit");
  assert.ok(weldRows.some((row) => String(row[4]).indexOf("40x40x3") >= 0), "welder steel row missing");

  console.log("floor-store.test.js ok");
}

main().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
