const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-mat-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const { callShopFunction } = require("./gas");

initWorkbook();
const book = getBook();
book.getSheetByName("Users").appendRow(["Nomsa", "Quality Control", "1234", "Quality Control", "Production", "No"]);
persistWorkbook();

const CONFIRM = { understood: true, highlights: [] };
const QC = [{ q: "Is the frame square?", a: "Y" }];
const SIG = "data:image/png;base64,aaa";

(async function main() {
  await callShopFunction("grantOvertime", ["Nomsa", "", "Admin", "test"]);
  const types = await callShopFunction("getGlassTypes", []);
  assert.ok(types.types.indexOf("Reeded") !== -1, JSON.stringify(types));
  assert.ok(types.components.indexOf("backboard") !== -1);
  assert.ok(types.thickness.indexOf("8mm") !== -1);

  const order = db.upsertOrder({
    order_number: "S-PRE-GLASS",
    status: "Ready for Pre-Powder Coating",
    type: "Standard",
    product: "Slider",
    price_excl_vat: "100.00"
  });
  const started = await callShopFunction("startOrder", [
    order.id, "Nomsa", "Quality Control", [], "", false, null, CONFIRM
  ]);
  assert.strictEqual(started.success, true, JSON.stringify(started));

  const noGlass = await callShopFunction("finishOrder", [
    order.id, started.logId, QC, SIG, [], "Nomsa", [], "S-PRE-GLASS", [], [], []
  ]);
  assert.ok(noGlass && noGlass.success === false, JSON.stringify(noGlass));
  assert.ok(/glass/i.test(noGlass.error || noGlass.message || ""), JSON.stringify(noGlass));

  const glass = [{
    component: "Door",
    type: "Frosted ripple",
    thickness: "6",
    height: 1800,
    width: 500,
    quantity: 2
  }];
  const wood = [{
    component: "Shelf",
    type: "Oak veneer",
    thickness: "16mm",
    height: 400,
    width: 800,
    quantity: 3
  }];
  const finished = await callShopFunction("finishOrder", [
    order.id, started.logId, QC, SIG, [], "Nomsa", [], "S-PRE-GLASS", [], glass, wood
  ]);
  assert.strictEqual(finished.success, true, JSON.stringify(finished));

  const listed = await callShopFunction("listMaterialsToOrder", []);
  assert.strictEqual(listed.glass.length, 1, JSON.stringify(listed.glass));
  assert.strictEqual(listed.glass[0].order, "S-PRE-GLASS");
  assert.strictEqual(listed.glass[0].type, "Frosted ripple");
  assert.strictEqual(listed.glass[0].thickness, "6mm");
  assert.strictEqual(listed.glass[0].status, "To order");
  assert.strictEqual(listed.wood.length, 1);
  assert.strictEqual(listed.wood[0].type, "Oak veneer");
  assert.ok(listed.glassTypes.types.indexOf("Frosted ripple") !== -1, JSON.stringify(listed.glassTypes));
  assert.ok(listed.woodTypes.types.indexOf("Oak veneer") !== -1, JSON.stringify(listed.woodTypes));

  const marked = await callShopFunction("markMaterialOrdered", ["glass", listed.glass[0].id, "Ordered"]);
  assert.strictEqual(marked.success, true, JSON.stringify(marked));
  const after = await callShopFunction("listMaterialsToOrder", []);
  assert.strictEqual(after.glass[0].status, "Ordered");

  const briefAfter = await callShopFunction("getOrderJobBrief", ["S-PRE-GLASS", "Pre-Powder Coating"]);
  assert.ok(briefAfter.standardGlass, JSON.stringify(briefAfter));
  assert.strictEqual(briefAfter.standardGlass.noGlass, false);
  assert.strictEqual(briefAfter.standardGlass.lines.length, 1);
  assert.strictEqual(briefAfter.standardGlass.lines[0].type, "Frosted ripple");
  assert.strictEqual(briefAfter.standardGlass.from_order_number, "S-PRE-GLASS");

  const noneOrder = db.upsertOrder({
    order_number: "S-PRE-NONE",
    status: "Ready for Pre-Powder Coating",
    type: "Standard",
    product: "Steel bench",
    price_excl_vat: "80.00"
  });
  const startedNone = await callShopFunction("startOrder", [
    noneOrder.id, "Nomsa", "Quality Control", [], "", false, null, CONFIRM
  ]);
  assert.strictEqual(startedNone.success, true, JSON.stringify(startedNone));
  const finishedNone = await callShopFunction("finishOrder", [
    noneOrder.id, startedNone.logId, QC, SIG, [], "Nomsa", [], "S-PRE-NONE", [], { noGlass: true }, []
  ]);
  assert.strictEqual(finishedNone.success, true, JSON.stringify(finishedNone));
  const listedNone = await callShopFunction("listMaterialsToOrder", []);
  assert.ok(!(listedNone.glass || []).some((g) => g.order === "S-PRE-NONE"));
  const noneBrief = await callShopFunction("getOrderJobBrief", ["S-PRE-NONE", "Pre-Powder Coating"]);
  assert.ok(noneBrief.standardGlass);
  assert.strictEqual(noneBrief.standardGlass.noGlass, true);

  console.log("materials-order.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
