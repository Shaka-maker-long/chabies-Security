const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const assert = require("assert");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-paste-ord-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const staff = require("./staff");
const db = require("./db");
const plan = require("./floor-planning");
const { parseOrderPaste } = require("./order-paste");
const { remainingPlanForStatus } = require("./shop-status");
const { mountOffice } = require("./office");

initWorkbook();
staff.upsertUser({
  name: "Office Boss",
  access: "Admin",
  role: "Manager",
  password: "admin",
  seeDebtors: "Yes",
  canManageUsers: true
});
staff.upsertUser({
  name: "Willard",
  access: "Production",
  role: "Cutter",
  password: "1234",
  tasks: ["Profile Cutting", "Plate Cutting"]
});

const SHEET = [
  "SOQ2913\tS260226 D\tReady for Tagging\t\tStandard\tCabinet\tThandi Display Cabinet\tTop & bottom shelves steel, Middles shelves: Steel\tReeded Glass\t\"Thandi Display Cabinet \n\n820mm W x 450mm D x 1800mm H\nCombination of 19mm and 13mm tubular steel frame, 4mm toughened reeded glass doors and sides, all shelves 1.5mm sheet metal\nsitting flush with frame. MDF backboard. Powder coated ferrograin black\"\tStandard\tFerrograin Black\tWinelands Design Studio (Pty) Ltd\t(076) 440 5425\tchante@winelandsdesignstudio.com\t03-Sep\t\"1310 La Vue\nVal de Vie Estate\nPaarl\"\t\tWestern Cape\tR13,949.00\tR12,129.57\t Sep 2026\tRecurring Client\tPaarl\t\t",
  "SOQ2892\tS260227\t\t\tCustom\tBedframe\tBlaire Platform Bed\tBlaire Platform Bed - King Extra Length: 1870mm W x 2020mm L x 300mm H\tN/A\t\"Custom Blaire Platform Bed \n\n1870mm W x 2020mm L x 300mm H (KING XL)\nCombination of 50mm x 25mm rectangular tubing and 25mm round tubing frame. Plywood slats below to support frame. Powdercoate\nd ferrograin black. Unit to be made to be assembled on site by the client. Refer to drawings folder for approved design. Bolts and screws for assembly to be sent with and clearly marked with packaging.\"\t1870mm W x 2020mm L x 300mm H (KING XL)\tFerrograin Black\tSophie Fischer-Linley\t(071) 471 9090\tlinleymax@gmail.com\t04-Sep\t\"39 Rosmead Avenue\nUnit 9 Queensbury Mansions\nGardens\nCape Town\n8001\"\t\tWestern Cape\tR5,919.00\tR5,146.96\t Sep 2026\tNo Trace\tCape Town\t\t",
  "SOQ2927\tS260228 A\tWelding\tWillard\tStandard\tWardrobe\tBrooke Open Wardrobe\tTop & bottom shelves steel, Middles shelves: Steel\tN/A\t\"Brooke Open Wardrobe\n\n1200mm W x 440mm D x 1800mm H\n19mm tubular steel frame, all shelves 1,5mm steel. Powder coated ferrograin black.\"\tStandard\tFerrograin Black\tBianca Van der Merwe\t(072) 439 0435\tvandermerwe.bianca@gmail.com\t04-Sep\t\"360 Marais St\nPretoria\nGauteng\n0011\"\t\tGauteng\tR3,449.00\tR2,999.13\t Sep 2026\tRecurring Client\tPretoria\t\t",
  "SOQ2927\tS260228 B\tReady for Grinding\t\tStandard\tMirror\tZahara Arched Mirror\tN/A\tN/A\t\"Zahara Arched Mirror \n\n600mm W x 25mm D x 1800mm\n19mm arched tubular steel frame with 4mm arched mirror sitting flush with the back of the unit. Freestanding unit. Powder coated smooth black.\"\tStandard\tSmooth Matt Black\tBianca Van der Merwe\t(072) 439 0435\tvandermerwe.bianca@gmail.com\t04-Sep\t\"360 Marais St\nPretoria\nGauteng\n0011\"\t\tGauteng\tR3,499.00\tR3,042.61\t Sep 2026\tRecurring Client\tPretoria\t\t"
].join("\n");

const parsed = parseOrderPaste(SHEET, { paidInFull: true });
assert.strictEqual(parsed.errors.length, 0, parsed.errors.join(" · "));
assert.strictEqual(parsed.rows.length, 4);

const byNo = {};
parsed.rows.forEach((row) => { byNo[row.order_number] = row; });

assert.strictEqual(byNo["S260226 D"].status, "Ready for Tagging");
assert.strictEqual(byNo["S260226 D"].product, "Thandi Display Cabinet");
assert.strictEqual(byNo["S260226 D"].client_name, "Winelands Design Studio (Pty) Ltd");
assert.strictEqual(byNo["S260226 D"].price_incl_vat, "13949.00");
assert.strictEqual(byNo["S260226 D"].price_excl_vat, "12129.57");
assert.strictEqual(byNo["S260226 D"].amount_paid, "13949.00");
assert.strictEqual(byNo["S260226 D"].payment_date, "2026-09-03");
assert.strictEqual(byNo["S260226 D"].month_of_sale, "September 2026");
assert.ok(byNo["S260226 D"].detailed_description.indexOf("reeded glass") !== -1);
assert.ok(byNo["S260226 D"].address.indexOf("Val de Vie") !== -1);

assert.strictEqual(byNo.S260227.status, "Not Yet Started");
assert.strictEqual(byNo.S260227.type, "Custom");
assert.strictEqual(byNo.S260227.product, "Blaire Platform Bed");
assert.strictEqual(byNo.S260227.source, "No Trace");
assert.strictEqual(byNo.S260227.city, "Cape Town");
assert.strictEqual(byNo.S260227.amount_paid, "5919.00");

assert.strictEqual(byNo["S260228 A"].status, "Welding");
assert.strictEqual(byNo["S260228 A"].assigned_operator, "Willard");
assert.strictEqual(byNo["S260228 A"].product, "Brooke Open Wardrobe");
assert.strictEqual(byNo["S260228 A"].amount_paid, "3449.00");

assert.strictEqual(byNo["S260228 B"].status, "Ready for Grinding");
assert.strictEqual(byNo["S260228 B"].assigned_operator, "");
assert.strictEqual(byNo["S260228 B"].powder_coating, "Smooth Matt Black");
assert.strictEqual(byNo["S260228 B"].amount_paid, "3499.00");

const headed = parseOrderPaste(
  "QUOTE NUMBER\tORDER NUMBER\tSTATUS\tASSIGNED OPERATOR\tTYPE\tCATERGORY\tPRODUCT\tVARIATION\tDOORS\tDETAILED DESCRIPTION\tDIMENSIONS\tPOWDER COATING\tCLIENT NAME\tCLIENT NUMBER\tEMAIL ADDRESS\tPAYMENT DATE\tADDRESS\tPROVINCE\tPRICE (Excl VAT)\tPRICE (Incl VAT)\tAMOUNT PAID\tMONTH OF SALE\tSOURCE\tCITY\n" +
  "SOQ1\tS260300\t\t\tStandard\tChair\tAir Chair\t\t\tA chair\tStandard\tBlack\tAda\t011\tada@test.com\t2026-09-01\t1 Road\tGauteng\t1000\t1150\t\tSeptember 2026\tWebsite\tJohannesburg",
  { paidInFull: true }
);
assert.strictEqual(headed.rows.length, 1, headed.errors.join(" · "));
assert.strictEqual(headed.rows[0].order_number, "S260300");
assert.strictEqual(headed.rows[0].status, "Not Yet Started");
assert.strictEqual(headed.rows[0].price_incl_vat, "1150.00");
assert.strictEqual(headed.rows[0].amount_paid, "1150.00");

const preview = db.pasteOrdersFromSheet({ text: SHEET, paid_in_full: true, preview: true });
assert.strictEqual(preview.added.length, 4);
assert.strictEqual(db.listOrders().length, 0, "preview must not write");

const created = db.pasteOrdersFromSheet({ text: SHEET, paid_in_full: true });
assert.strictEqual(created.added.length, 4);
assert.strictEqual(created.skipped.length, 0);
assert.strictEqual(created.errors.length, 0);

const saved = {};
db.listOrders().forEach((row) => { saved[row.order_number] = row; });
assert.strictEqual(saved["S260226 D"].status, "Ready for Tagging");
assert.strictEqual(saved.S260227.status, "Not Yet Started");
assert.strictEqual(saved["S260228 A"].status, "Welding");
assert.strictEqual(saved["S260228 A"].assigned_operator, "Willard");
assert.strictEqual(saved["S260228 B"].status, "Ready for Grinding");
assert.strictEqual(db.parseMoney(saved["S260226 D"].amount_paid), 13949);
assert.ok(Math.abs(db.parseMoney(saved["S260226 D"].price_incl_vat) - db.parseMoney(saved["S260226 D"].amount_paid)) <= 0.01);
assert.ok(Math.abs(db.parseMoney(saved.S260227.price_incl_vat) - db.parseMoney(saved.S260227.amount_paid)) <= 0.01);
assert.ok(String(saved["S260226 D"].payment_date).indexOf("03") !== -1);

const again = db.pasteOrdersFromSheet({ text: SHEET, paid_in_full: true });
assert.strictEqual(again.added.length, 0);
assert.strictEqual(again.skipped.length, 4);
assert.ok(again.skipped.every((row) => /already on Orders/.test(row.reason)));

const weld = remainingPlanForStatus(saved["S260228 A"].status);
assert.deepStrictEqual(weld.processes, ["Plate Cutting", "Grinding", "Assembly"]);
const q = plan.queueOrders().find((row) => row.order_number === "S260228 A");
assert.ok(q, "pasted in-progress orders stay in Planning for remaining work");
assert.ok(!q.processes.some((p) => p.process === "Welding"));
assert.ok(q.processes.some((p) => p.process === "Grinding"));

const bed = plan.queueOrders().find((row) => row.order_number === "S260227");
assert.ok(bed, "blank-status paste stays in Planning as Not Yet Started");

(async function main() {
  const app = express();
  app.use(express.json({ limit: "2mb" }));
  mountOffice(app);
  const server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const base = "http://127.0.0.1:" + server.address().port;
  const login = await fetch(base + "/api/office/login", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name: "Office Boss", password: "admin" })
  });
  const session = await login.json();
  assert.ok(session.ok, JSON.stringify(session));

  const previewRes = await fetch(base + "/api/office/orders/paste", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({
      text: "SOQ9\tS260310\t\t\tStandard\tChair\tAir Chair\t\t\tA chair\tStandard\tBlack\tPat\t012\tpat@test.com\t05-Sep\t1 Road\t\tGauteng\tR1,150.00\tR1,000.00\t Sep 2026\tWebsite\tPretoria",
      paid_in_full: true,
      preview: true
    })
  });
  const previewJson = await previewRes.json();
  assert.ok(previewJson.ok, JSON.stringify(previewJson));
  assert.strictEqual(previewJson.preview, true);
  assert.strictEqual(previewJson.added.length, 1);
  assert.strictEqual(previewJson.added[0].status, "Not Yet Started");
  assert.strictEqual(db.parseMoney(previewJson.added[0].amount_paid), 1150);
  assert.ok(!db.listOrders().some((o) => o.order_number === "S260310"));

  const addRes = await fetch(base + "/api/office/orders/paste", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({
      text: "SOQ9\tS260310\tWelding\tWillard\tStandard\tChair\tAir Chair\t\t\tA chair\tStandard\tBlack\tPat\t012\tpat@test.com\t05-Sep\t1 Road\t\tGauteng\tR1,150.00\tR1,000.00\t Sep 2026\tWebsite\tPretoria",
      paid_in_full: true
    })
  });
  const addJson = await addRes.json();
  assert.ok(addJson.ok, JSON.stringify(addJson));
  assert.strictEqual(addJson.added.length, 1);
  assert.strictEqual(addJson.added[0].status, "Welding");
  assert.strictEqual(addJson.added[0].assigned_operator, "Willard");
  const live = db.listOrders().find((o) => o.order_number === "S260310");
  assert.ok(live);
  assert.strictEqual(live.status, "Welding");

  const skipRes = await fetch(base + "/api/office/orders/paste", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": session.token },
    body: JSON.stringify({ text: SHEET, paid_in_full: true })
  });
  const skipJson = await skipRes.json();
  assert.ok(skipJson.ok);
  assert.strictEqual(skipJson.added.length, 0);
  assert.strictEqual(skipJson.skipped.length, 4);

  server.close();
  console.log("order-paste.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
