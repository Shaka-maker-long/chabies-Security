const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const assert = require("assert");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-noplate-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const staff = require("./staff");
const db = require("./db");
const noPlates = require("./no-plates");
const { callShopFunction } = require("./gas");
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
  name: "Prod Boss",
  access: "Admin",
  role: "Production Manager",
  password: "admin",
  seeDebtors: "No"
});
staff.upsertUser({
  name: "Thabile",
  access: "Production",
  role: "Plate Cutter",
  password: "1234",
  tasks: ["Plate Cutting"]
});
staff.upsertUser({
  name: "Lesedi",
  access: "Admin",
  role: "Admin",
  password: "office",
  seeDebtors: "No"
});

function seed(num, status) {
  return db.upsertOrder({
    order_number: num,
    status: status,
    product: "Air Chair",
    client_name: "Bea"
  });
}

assert.strictEqual(noPlates.canMarkNoPlate(staff.listUsers().find((u) => u.name === "Office Boss")), true);
assert.strictEqual(noPlates.canMarkNoPlate(staff.listUsers().find((u) => u.name === "Prod Boss")), true);
assert.strictEqual(noPlates.canMarkNoPlate(staff.listUsers().find((u) => u.name === "Thabile")), true);
assert.strictEqual(noPlates.canMarkNoPlate(staff.listUsers().find((u) => u.name === "Lesedi")), false);

const tag = seed("SD-TAG", "Tagging");
seed("SD-RFW", "Ready for Welding");
seed("SD-WELD", "Welding");
seed("SD-NYS", "Not Yet Started");
seed("SD-CUT", "Profile Cutting");
seed("SD-GRIND", "Ready for Grinding");

(async () => {
  const listed = await callShopFunction("pollFloor", ["Plate Cutting", "Thabile"]);
  const nums = ((listed && listed.orders) || []).map((o) => o.order);
  assert.ok(nums.indexOf("SD-TAG") !== -1, "Tagging must show for plate cutting");
  assert.ok(nums.indexOf("SD-RFW") !== -1, "Ready for Welding must show for plate cutting");
  assert.ok(nums.indexOf("SD-WELD") !== -1, "Welding must show for plate cutting");
  assert.ok(nums.indexOf("SD-NYS") === -1, "Not Yet Started must not show");
  assert.ok(nums.indexOf("SD-CUT") === -1, "Profile Cutting must not show");
  assert.ok(nums.indexOf("SD-GRIND") === -1, "Ready for Grinding must not show");

  const marked = await callShopFunction("markOrderNoPlate", ["SD-TAG", "Thabile"]);
  assert.ok(marked.success, JSON.stringify(marked));
  const after = await callShopFunction("pollFloor", ["Plate Cutting", "Thabile"]);
  assert.ok(!((after && after.orders) || []).some((o) => o.order === "SD-TAG"), "no-plate order must leave the plate list");

  const blocked = await callShopFunction("startOrder", [
    tag.id, "Thabile", "Plate Cutting", [], "", false, null, { understood: true, highlights: [] }
  ]);
  assert.strictEqual(blocked.success, false, JSON.stringify(blocked));
  assert.ok(/no plates/i.test(blocked.message || blocked.error || ""), JSON.stringify(blocked));

  const early = seed("SD-EARLY", "Ready for Steelwork");
  const tooSoon = await callShopFunction("startOrder", [
    early.id, "Thabile", "Plate Cutting", [], "", false, null, { understood: true, highlights: [] }
  ]);
  assert.strictEqual(tooSoon.success, false, JSON.stringify(tooSoon));
  assert.ok(/Tagging|Welding/i.test(tooSoon.message || tooSoon.error || ""), JSON.stringify(tooSoon));

  const denied = await callShopFunction("markOrderNoPlate", ["SD-WELD", "Lesedi"]);
  assert.strictEqual(denied.success, false);

  const app = express();
  app.use(express.json());
  mountOffice(app);
  const server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const base = "http://127.0.0.1:" + server.address().port;

  async function login(name, password) {
    const r = await fetch(base + "/api/office/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: name, password: password })
    });
    return r.json();
  }

  const prod = await login("Prod Boss", "admin");
  assert.ok(prod.ok, JSON.stringify(prod));
  assert.ok(prod.canMarkNoPlate);
  const prodMark = await fetch(base + "/api/office/orders/no-plate", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": prod.token },
    body: JSON.stringify({ order_number: "SD-RFW", no_plate: true })
  });
  const prodJson = await prodMark.json();
  assert.ok(prodJson.ok, JSON.stringify(prodJson));
  assert.ok(noPlates.isNoPlate("SD-RFW"));

  const mgr = await login("Office Boss", "admin");
  const mgrClear = await fetch(base + "/api/office/orders/no-plate", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": mgr.token },
    body: JSON.stringify({ order_number: "SD-RFW", no_plate: false })
  });
  assert.ok((await mgrClear.json()).ok);
  assert.ok(!noPlates.isNoPlate("SD-RFW"));

  const clerk = await login("Lesedi", "office");
  assert.ok(clerk.ok);
  assert.ok(!clerk.canMarkNoPlate);
  const clerkMark = await fetch(base + "/api/office/orders/no-plate", {
    method: "POST",
    headers: { "Content-Type": "application/json", "x-sd-token": clerk.token },
    body: JSON.stringify({ order_number: "SD-WELD", no_plate: true })
  });
  assert.strictEqual(clerkMark.status, 403);

  server.close();
  console.log("no-plates.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
