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
assert.ok(page.indexOf("Ready for Steelwork") !== -1);

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
  assert.ok(!Object.prototype.hasOwnProperty.call(created.record, "builder"));
  assert.ok(!created.record.builder);
  assert.strictEqual(created.record.cutting.tubes.length, 2);
  assert.ok(fs.existsSync(created.record.pdf_path));
  const pdf = fs.readFileSync(created.record.pdf_path);
  assert.ok(pdf.slice(0, 4).toString() === "%PDF");
  assert.ok(pdf.toString("latin1").indexOf("BUILDER") === -1);
  assert.ok(pdf.toString("latin1").indexOf("S260193") !== -1);
  assert.ok(pdf.toString("latin1").indexOf("BOM") === -1);

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

  console.log("job-card.test.js ok");
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
