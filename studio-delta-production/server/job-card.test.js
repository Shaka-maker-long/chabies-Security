const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-jobcard-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const db = require("./db");
const catalog = require("./product-catalog");
const jobCard = require("./job-card");

initWorkbook();

const samplePaste = [
  "13mm SQ Square Tube\tWidth\t600\t\t2\t2",
  "13mm SQ Square Tube\tHeight\t600\t\t4\t2",
  "1.6mm Plate\tTop (Bend)\t587\t401\t1\t0",
  "18mm MDF Board\tShelf\t400\t350\t2\t0"
].join("\n");

const parsed = jobCard.parsePastedCuttingList(samplePaste);
assert.strictEqual(parsed.tubes.length, 2, "tubes from paste");
assert.strictEqual(parsed.plates.length, 1, "plates from paste");
assert.strictEqual(parsed.wood.length, 1, "wood from paste");
assert.strictEqual(parsed.bars.length, 0);
assert.strictEqual(parsed.tubes[0].length, 600);
assert.strictEqual(parsed.tubes[0].quantity, 2);
assert.strictEqual(parsed.tubes[0].cutAngle45, 2);
assert.strictEqual(parsed.plates[0].height, 587);
assert.strictEqual(parsed.plates[0].width, 401);
assert.ok(jobCard.cuttingCount(parsed) === 4);

const skipped = jobCard.parsePastedCuttingList("13mm SQ Square Tube\tWidth\t0\t\t2\t2");
assert.strictEqual(skipped.tubes.length, 0, "skip zero length");

assert.strictEqual(
  jobCard.sameShopProduct(
    {
      product: "Thandi Display Cabinet",
      type: "New Design",
      variation: "Top & bottom shelves steel",
      doors: "Clear glass",
      powder_coating: "As per website",
      detailed_description: "Unit with ⟦additional shelf⟧ on the right"
    },
    {
      product: "Thandi Display Cabinet",
      type: "New Design",
      variation: "Top & bottom shelves steel",
      doors: "Clear glass",
      powder_coating: "As per website",
      detailed_description: "Unit with additional shelf on the right"
    }
  ),
  true,
  "important marks do not split a match"
);
assert.strictEqual(
  jobCard.sameShopProduct(
    { product: "Thandi Display Cabinet", type: "New Design", variation: "x", doors: "Clear glass", powder_coating: "Black", detailed_description: "Same" },
    { product: "Thandi Display Cabinet", type: "Custom", variation: "x", doors: "Clear glass", powder_coating: "Black", detailed_description: "Same" }
  ),
  false,
  "different type is not the same product"
);
assert.strictEqual(
  jobCard.sameShopProduct(
    { product: "Thandi Display Cabinet", type: "Standard", variation: "", doors: "N/A", powder_coating: "Black", detailed_description: "Same", dimensions: "Height: 1800mm" },
    { product: "Thandi Display Cabinet", type: "Standard", variation: "", doors: "N/A", powder_coating: "Black", detailed_description: "Same", dimensions: "" }
  ),
  true,
  "blank dimensions on one unit still match"
);

const roundTripPaste = jobCard.formatCuttingPaste(parsed);
const reparsed = jobCard.parsePastedCuttingList(roundTripPaste);
assert.strictEqual(reparsed.tubes.length, parsed.tubes.length);
assert.strictEqual(reparsed.plates.length, parsed.plates.length);
assert.strictEqual(reparsed.wood.length, parsed.wood.length);

assert.strictEqual(jobCard.jobCardEligibility("Not Yet Started").mode, "create");
assert.strictEqual(jobCard.jobCardEligibility("").mode, "create");
assert.strictEqual(jobCard.jobCardEligibility("Ready for Steelwork").mode, "regenerate");
assert.strictEqual(jobCard.jobCardEligibility("Profile Cutting").mode, "regenerate");
assert.strictEqual(jobCard.jobCardEligibility("Ready for Tagging").ok, false);
assert.strictEqual(jobCard.jobCardEligibility("Welding").ok, false);

assert.strictEqual(jobCard.applyOfficeOrderStatusLock({ status: "Delivered" }, null).status, "Not Yet Started");
assert.strictEqual(
  jobCard.applyOfficeOrderStatusLock({ status: "Delivered" }, { status: "Profile Cutting" }).status,
  "Profile Cutting"
);

const talitha = catalog.lookupProduct("Talitha Bookshelf");
assert.ok(talitha);
assert.strictEqual(talitha.height, 1460);
assert.strictEqual(talitha.width, 560);
assert.ok(talitha.imageUrl.indexOf("TALITHA") !== -1);

const felicity = catalog.lookupProduct("Felicity Floating Shelf - Large");
assert.ok(felicity);
assert.strictEqual(felicity.name, "Felicity Open Shelf Large");
assert.strictEqual(felicity.height, 300);

const officeJs = fs.readFileSync(path.join(__dirname, "office.js"), "utf8");
assert.ok(officeJs.indexOf("applyOfficeOrderStatusLock") !== -1);
assert.ok(officeJs.indexOf("upsertOrder(req.body") === -1, "office PUT must not write the form status as-is");
assert.ok(officeJs.indexOf("/api/office/job-cards") !== -1);

const page = fs.readFileSync(path.join(__dirname, "../public/job-card.html"), "utf8");
assert.ok(page.indexOf("Builder") === -1, "job card UI must not ask for a builder");
assert.ok(page.indexOf("BOM") === -1);
assert.ok(page.indexOf("materialsTable") === -1);
assert.ok(page.indexOf("Add All Cutting List Items to BOM") === -1);
assert.ok(page.indexOf("Paste cutting list") !== -1);
assert.ok(page.indexOf("Generate job card") !== -1);
assert.ok(page.indexOf("Mark selected as important") !== -1);
assert.ok(page.indexOf("textarea id=\"description\" readonly") === -1, "shop description is typed on the job card");
assert.ok(page.indexOf("Ready for Steelwork") !== -1);
assert.ok(page.indexOf("Printed job cards") !== -1);
assert.ok(page.indexOf("Open / print") !== -1);
assert.ok(page.indexOf("/api/office/job-cards") !== -1);
assert.ok(page.indexOf("name=\"dimCheck\"") !== -1);
assert.ok(page.indexOf("Dimensions did not change") !== -1);
assert.ok(page.indexOf("Update to the actual dimensions") !== -1);
assert.ok(page.indexOf("dimension_check") !== -1);
assert.ok(page.indexOf("persistJobCardDraft") !== -1, "job card form keeps unsaved progress");
assert.ok(officeJs.indexOf("listGeneratedJobCards") !== -1);
assert.strictEqual(jobCard.isStandardType("Standard"), true);
assert.strictEqual(jobCard.isStandardType("Custom"), false);
assert.strictEqual(jobCard.isTypeOnlyDescription("Standard"), true);
assert.strictEqual(jobCard.isTypeOnlyDescription("New Design"), true);
assert.strictEqual(jobCard.isTypeOnlyDescription("Unit with extra shelf"), false);
assert.strictEqual(jobCard.shopDescription({ type: "New Design", detailed_description: "Standard" }), "");
assert.strictEqual(jobCard.shopDescription({ detailed_description: "Talitha Bookshelf" }), "Talitha Bookshelf");
assert.strictEqual(
  jobCard.shopDescription({ detailed_description: "Standard" }, "Extra hanging rail"),
  "Extra hanging rail"
);
assert.deepStrictEqual(
  jobCard.descriptionSegments("Unit with ⟦additional shelf⟧ on the right").map((s) => s.important),
  [false, true, false]
);
assert.strictEqual(jobCard.hasUsableDimensions({ height: 1, width: 1, depth: 1 }), false);
assert.strictEqual(jobCard.hasUsableDimensions({ height: 1460, width: 560, depth: 560 }), true);

db.upsertOrder({
  order_number: "S260193",
  status: "Not Yet Started",
  type: "Standard",
  product: "Talitha Bookshelf",
  client_name: "Tariq Koor",
  doors: "N/A",
  powder_coating: "Ferrograin Black",
  variation: "Top & bottom shelves steel, Middles shelves: Steel",
  detailed_description: "Talitha Bookshelf",
  dimensions: "Standard",
  province: "Gauteng"
});
db.upsertOrder({
  order_number: "S260200",
  status: "Ready for Tagging",
  product: "Talitha Bookshelf",
  client_name: "Blocked"
});

const eligible = jobCard.listEligibleOrders();
assert.ok(eligible.some((r) => r.order_number === "S260193"));
assert.ok(!eligible.some((r) => r.order_number === "S260200"));

(async function main() {
  const created = await jobCard.generateJobCard({
    order_number: "S260193",
    cutting_text: samplePaste
  });
  assert.strictEqual(created.status, "Ready for Steelwork");
  assert.strictEqual(created.record.start_date, created.record.created_date);
  assert.strictEqual(created.record.start_date, jobCard.todayIso());
  assert.strictEqual(created.record.builder, "");
  const front = jobCard.jobCardFrontRows(created.record);
  assert.strictEqual(front[0][0].label, "BUILDER:");
  assert.strictEqual(front[0][1].label, "DESIGN TYPE:");
  assert.strictEqual(front[1][0].label, "ASSEMBLER:");
  assert.strictEqual(front[1][1].label, "ORDER NUMBER:");
  assert.strictEqual(front[1][1].value, "S260193");
  assert.strictEqual(front[2][0].label, "DATE START:");
  assert.strictEqual(front[2][1].label, "CUSTOMER NAME:");
  assert.strictEqual(front[2][1].value, "Tariq Koor");
  assert.strictEqual(front[3][1].label, "PRODUCT:");
  assert.strictEqual(front[4][1].label, "GLASS TYPE:");
  assert.strictEqual(front[5][1].label, "VARIATIONS:");
  assert.strictEqual(created.record.cutting.tubes.length, 2);
  assert.ok(fs.existsSync(created.record.pdf_path));
  const pdf = fs.readFileSync(created.record.pdf_path);
  assert.ok(pdf.slice(0, 4).toString() === "%PDF");
  assert.ok(pdf.toString("latin1").indexOf("S260193") !== -1);
  assert.strictEqual(created.record.description, "Talitha Bookshelf");
  assert.ok(pdf.toString("latin1").indexOf("BOM") === -1);

  const listed = jobCard.listGeneratedJobCards();
  assert.ok(listed.some((r) => r.order_number === "S260193" && r.has_pdf && r.pdf_url));
  assert.ok(listed[0].order_number === "S260193");
  assert.ok(listed[0].download_url.indexOf("download=1") !== -1);

  const after = db.listOrders().find((o) => o.order_number === "S260193");
  assert.strictEqual(after.status, "Ready for Steelwork");

  const again = await jobCard.generateJobCard({
    order_number: "S260193",
    cutting_text: "13mm SQ Square Tube\tWidth\t500\t\t1\t0"
  });
  assert.strictEqual(again.status, "Ready for Steelwork", "regenerate must keep steelwork status");
  assert.strictEqual(again.record.cutting.tubes[0].length, 500);

  db.upsertOrder(Object.assign({}, after, { status: "Profile Cutting" }));
  const fromProfile = await jobCard.generateJobCard({
    order_number: "S260193",
    cutting: { bars: [], plates: [], tubes: [], wood: [] }
  });
  assert.strictEqual(fromProfile.status, "Profile Cutting");

  db.upsertOrder(Object.assign({}, after, { status: "Ready for Tagging" }));
  let blocked = null;
  try {
    await jobCard.generateJobCard({ order_number: "S260193", cutting_text: samplePaste });
  } catch (e) {
    blocked = e.message;
  }
  assert.ok(blocked && /Profile Cutting|Not Yet Started|Ready for Steelwork/.test(blocked));

  let missing = null;
  try {
    await jobCard.generateJobCard({ order_number: "S999999", cutting_text: samplePaste });
  } catch (e) {
    missing = e.message;
  }
  assert.ok(missing);

  db.upsertOrder({
    order_number: "S260210",
    status: "Not Yet Started",
    type: "Custom",
    product: "Talitha Bookshelf",
    client_name: "Custom Client",
    doors: "N/A",
    powder_coating: "Ferrograin Black",
    variation: "",
    detailed_description: "Talitha Bookshelf custom",
    dimensions: "",
    province: "Gauteng"
  });
  db.upsertOrder({
    order_number: "S260211",
    status: "Not Yet Started",
    type: "New Design",
    product: "New Design",
    client_name: "New Design Client",
    doors: "N/A",
    powder_coating: "Ferrograin Black",
    variation: "",
    detailed_description: "A new bench",
    dimensions: "",
    province: "Western Cape"
  });

  let noCheck = null;
  try {
    await jobCard.generateJobCard({ order_number: "S260210", cutting_text: samplePaste });
  } catch (e) {
    noCheck = e.message;
  }
  assert.ok(noCheck && /not Standard|did not change|actual/i.test(noCheck));

  const customOk = await jobCard.generateJobCard({
    order_number: "S260210",
    cutting_text: samplePaste,
    dimension_check: "unchanged"
  });
  assert.strictEqual(customOk.status, "Ready for Steelwork");
  assert.strictEqual(customOk.record.dimension_check, "unchanged");
  const customOrder = db.listOrders().find((o) => o.order_number === "S260210");
  assert.ok(/Height:\s*1460mm/.test(customOrder.dimensions));
  assert.ok(/Width:\s*560mm/.test(customOrder.dimensions));

  let placeholder = null;
  try {
    await jobCard.generateJobCard({
      order_number: "S260211",
      cutting_text: samplePaste,
      dimension_check: "unchanged"
    });
  } catch (e) {
    placeholder = e.message;
  }
  assert.ok(placeholder && /1 × 1 × 1|actual dimensions|millimetres/i.test(placeholder));

  const newDesignOk = await jobCard.generateJobCard({
    order_number: "S260211",
    cutting_text: samplePaste,
    dimension_check: "updated",
    dimensions: { height: 1500, width: 600, depth: 400 }
  });
  assert.strictEqual(newDesignOk.status, "Ready for Steelwork");
  assert.strictEqual(newDesignOk.record.dimension_check, "updated");
  assert.strictEqual(newDesignOk.record.dimensions.height, 1500);
  const newDesignOrder = db.listOrders().find((o) => o.order_number === "S260211");
  assert.ok(/Height:\s*1500mm/.test(newDesignOrder.dimensions));
  assert.ok(/Width:\s*600mm/.test(newDesignOrder.dimensions));

  db.upsertOrder({
    order_number: "S260212",
    status: "Not Yet Started",
    type: "New Design",
    product: "Thandi Display Cabinet",
    client_name: "Winelands Design Studio",
    doors: "Clear glass",
    powder_coating: "As per website",
    variation: "Top & bottom shelves steel, Middles shelves: Glass",
    detailed_description: "Standard",
    dimensions: "",
    province: "Gauteng"
  });
  let typeInDescription = null;
  try {
    await jobCard.generateJobCard({
      order_number: "S260212",
      cutting_text: samplePaste,
      dimension_check: "unchanged"
    });
  } catch (e) {
    typeInDescription = e.message;
  }
  assert.ok(typeInDescription && /detailed description|Design type/i.test(typeInDescription));

  const typedDesc = await jobCard.generateJobCard({
    order_number: "S260212",
    cutting_text: samplePaste,
    dimension_check: "unchanged",
    description: "Unit with ⟦additional shelf⟧ on the right"
  });
  assert.strictEqual(typedDesc.record.design_type, "New Design");
  assert.strictEqual(typedDesc.record.description, "Unit with ⟦additional shelf⟧ on the right");
  assert.ok(typedDesc.record.description.indexOf("Standard") === -1);
  assert.ok(String(typedDesc.order.detailed_description).indexOf("additional") !== -1);
  const savedDesc = db.listOrders().find((o) => o.order_number === "S260212");
  assert.ok(savedDesc);
  assert.strictEqual(savedDesc.detailed_description, typedDesc.order.detailed_description);
  assert.ok(savedDesc.detailed_description !== "Standard");

  const twin = {
    status: "Not Yet Started",
    type: "Standard",
    product: "Thandi Display Cabinet",
    client_name: "Split Client",
    doors: "Clear glass",
    powder_coating: "As per website",
    variation: "Top & bottom shelves steel, Middles shelves: Glass",
    detailed_description: "Winelands cabinet with extra shelf",
    dimensions: "Height: 1800mm, Width: 900mm",
    province: "Western Cape"
  };
  db.upsertOrder(Object.assign({}, twin, { order_number: "S260001 A" }));
  db.upsertOrder(Object.assign({}, twin, { order_number: "S260001 B" }));
  db.upsertOrder(Object.assign({}, twin, {
    order_number: "S260001 C",
    product: "Talitha Bookshelf",
    detailed_description: "Different product in the same split"
  }));

  const generatedA = await jobCard.generateJobCard({
    order_number: "S260001 A",
    cutting_text: samplePaste
  });
  assert.ok(jobCard.cuttingCount(generatedA.record.cutting) > 0);

  const eligibleAfterA = jobCard.listEligibleOrders();
  const rowA = eligibleAfterA.find((r) => r.order_number === "S260001 A");
  const rowB = eligibleAfterA.find((r) => r.order_number === "S260001 B");
  const rowC = eligibleAfterA.find((r) => r.order_number === "S260001 C");
  assert.ok(rowB && rowB.suggested_cutting, "B should be offered A's cutting list");
  assert.strictEqual(rowB.suggested_cutting.from_order_number, "S260001 A");
  assert.strictEqual(rowB.suggested_cutting.same_split, true);
  assert.ok(rowB.suggested_cutting.paste_text);
  assert.strictEqual(rowB.suggested_cutting.cutting.tubes.length, generatedA.record.cutting.tubes.length);
  assert.ok(!rowA || !rowA.suggested_cutting, "do not suggest a card its own cutting list");
  assert.ok(!rowC || !rowC.suggested_cutting, "different product in the same split is not a match");

  const suggestedParse = jobCard.parsePastedCuttingList(rowB.suggested_cutting.paste_text);
  assert.strictEqual(suggestedParse.tubes.length, generatedA.record.cutting.tubes.length);
  assert.strictEqual(suggestedParse.plates.length, generatedA.record.cutting.plates.length);

  await jobCard.generateJobCard({
    order_number: "S260001 B",
    cutting: rowB.suggested_cutting.cutting
  });
  const afterB = jobCard.listEligibleOrders().find((r) => r.order_number === "S260001 B");
  assert.ok(!afterB || !afterB.suggested_cutting, "B with its own cutting is not offered a suggestion");
  assert.strictEqual(jobCard.getJobCard("S260001 A").cutting.tubes[0].length, generatedA.record.cutting.tubes[0].length);

  const beforeClear = db.listOrders().length;
  assert.ok(beforeClear >= 1);
  assert.ok(jobCard.getJobCard("S260211"));
  const wipedOrders = db.deleteAllOrders();
  assert.ok(wipedOrders >= 1);
  assert.strictEqual(db.listOrders().length, 0);
  jobCard.deleteAllJobCards();
  assert.ok(!jobCard.getJobCard("S260211"));

  console.log("job-card.test.js ok");
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
