"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const assert = require("assert");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-mkt-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
process.env.SD_TRUST_DEVICES = "0";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const staff = require("./staff");
const { mountOffice } = require("./office");
const { upsertOrder, upsertEnquiry, listOrders, getEnquiry } = require("./db");

initWorkbook();

staff.upsertUser({
  name: "Boss",
  access: "Admin",
  role: "Manager",
  password: "boss",
  seeDebtors: "Yes"
});
const marketer = staff.upsertUser({
  name: "Mia Market",
  access: "Marketing",
  role: "Marketing",
  password: "mkt"
});
assert.strictEqual(marketer.access, "Marketing");
assert.strictEqual(marketer.canSeeOffice, true);
assert.strictEqual(marketer.isAdmin, false);
assert.strictEqual(marketer.isMarketing, true);
assert.strictEqual(marketer.canEditMarketingFields, true);
assert.deepStrictEqual(marketer.tasks, []);
assert.deepStrictEqual(marketer.enquiryRoles, []);
assert.ok(staff.isMarketing(marketer));
assert.ok(!staff.isProductionFloorUser(marketer));
assert.strictEqual(staff.canSeeIdleAlerts(marketer), false, "Marketing must not get idle notifications");
assert.strictEqual(staff.canSeeIdleAlerts({
  name: "Mia Market",
  access: "Marketing",
  role: "Marketing",
  jobTitle: "Marketing"
}), false);

const order = upsertOrder({
  order_number: "S260901",
  quote_number: "SOQ9001",
  status: "Not Yet Started",
  client_name: "Test Client",
  source: "Google",
  campaign: "",
  price_excl_vat: "100.00",
  price_incl_vat: "115.00"
});
assert.strictEqual(order.source, "Google");

const enquiry = upsertEnquiry({
  enquiry_no: "#9001",
  client_name: "Test Client",
  status: "New",
  source: "Instagram",
  campaign: "",
  create_only: true
}, { actor: "Boss", createOnly: true });
assert.strictEqual(enquiry.source, "Instagram");

const app = express();
app.use(express.json({ limit: "2mb" }));
mountOffice(app);

(async () => {
  const server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;

  async function login(name, password) {
    const r = await fetch("http://127.0.0.1:" + port + "/api/office/login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, password })
    });
    const j = await r.json();
    assert.ok(j.ok, j.error || "login failed");
    return j;
  }

  async function api(token, method, url, body) {
    const r = await fetch("http://127.0.0.1:" + port + url, {
      method,
      headers: {
        "Content-Type": "application/json",
        "x-sd-token": token
      },
      body: body ? JSON.stringify(body) : undefined
    });
    const j = await r.json().catch(() => ({}));
    return { status: r.status, body: j };
  }

  const session = await login("Mia Market", "mkt");
  assert.strictEqual(session.isMarketing, true);
  assert.strictEqual(session.canEditMarketingFields, true);
  assert.strictEqual(session.canSeeOffice, true);
  assert.strictEqual(session.canSeeIdleAlerts, false, "Marketing login must not unlock idle alerts");

  const tasksBefore = await api(session.token, "GET", "/api/office/my-tasks");
  assert.strictEqual(tasksBefore.status, 200, JSON.stringify(tasksBefore.body));
  assert.ok((tasksBefore.body.rows || []).some((t) => t.kind === "campaign" && t.order_number === "S260901"),
    "orders without a campaign must sit on Marketing My tasks");
  assert.ok((tasksBefore.body.rows || []).every((t) => t.kind !== "campaign" || t.open_order),
    "campaign tasks open the order sheet");

  staff.upsertUser({
    name: "Sam Ads",
    access: "Marketing",
    role: "Marketing",
    password: "mkt2"
  });
  const sam = await login("Sam Ads", "mkt2");
  const samTasks = await api(sam.token, "GET", "/api/office/my-tasks");
  assert.ok((samTasks.body.rows || []).some((t) => t.kind === "campaign" && t.order_number === "S260901"),
    "every Marketing login sees orders missing a campaign");

  const okPatch = await api(session.token, "PUT", "/api/office/orders", {
    order_number: "S260901",
    source: "Billboards",
    campaign: "Spring Push",
    client_name: "HACKED",
    status: "Delivered",
    price_incl_vat: "1.00"
  });
  assert.strictEqual(okPatch.status, 200, JSON.stringify(okPatch.body));
  assert.ok(okPatch.body.ok);
  assert.strictEqual(okPatch.body.row.source, "Billboards");
  assert.strictEqual(okPatch.body.row.campaign, "Spring Push");
  assert.strictEqual(okPatch.body.row.client_name, "Test Client");
  assert.strictEqual(okPatch.body.row.status, "Not Yet Started");

  const listed = listOrders().find((o) => o.order_number === "S260901");
  assert.strictEqual(listed.client_name, "Test Client");
  assert.strictEqual(listed.campaign, "Spring Push");

  const tasksAfter = await api(session.token, "GET", "/api/office/my-tasks");
  assert.ok(!(tasksAfter.body.rows || []).some((t) => t.kind === "campaign" && t.order_number === "S260901"),
    "campaign task clears once Marketing saves a campaign");

  const enqPatch = await api(session.token, "PUT", "/api/office/enquiries", {
    enquiry_no: "#9001",
    source: "Magazine",
    campaign: "Decor Fair",
    client_name: "HACKED",
    status: "Ordered"
  });
  assert.strictEqual(enqPatch.status, 200, JSON.stringify(enqPatch.body));
  assert.strictEqual(enqPatch.body.row.source, "Magazine");
  assert.strictEqual(enqPatch.body.row.campaign, "Decor Fair");
  assert.strictEqual(enqPatch.body.row.client_name, "Test Client");
  assert.strictEqual(enqPatch.body.row.status, "New");

  const blockedCreate = await api(session.token, "PUT", "/api/office/enquiries", {
    enquiry_no: "#9999",
    client_name: "Nope",
    status: "New",
    create_only: true
  });
  assert.strictEqual(blockedCreate.status, 403);

  const blockedUsers = await api(session.token, "PUT", "/api/office/users", {
    name: "Someone",
    access: "Admin",
    password: "x"
  });
  assert.strictEqual(blockedUsers.status, 403);

  const blockedNewOrder = await api(session.token, "PUT", "/api/office/orders", {
    order_number: "S260999",
    client_name: "Nope",
    source: "Google"
  });
  assert.strictEqual(blockedNewOrder.status, 400);

  server.close();
  console.log("marketing-role.test.js ok");
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
