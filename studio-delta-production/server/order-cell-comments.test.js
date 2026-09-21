const fs = require("fs");
const os = require("os");
const path = require("path");
const http = require("http");
const assert = require("assert");
const express = require("express");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-occ-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook } = require("./workbook-store");
const db = require("./db");
const staff = require("./staff");
const { mountOffice } = require("./office");
const comments = require("./order-cell-comments");

initWorkbook();
staff.upsertUser({
  name: "Shaka",
  access: "Admin",
  role: "manager",
  password: "admin",
  seeDebtors: "Yes"
});
staff.upsertUser({
  name: "Office Ada",
  access: "Admin",
  role: "Admin",
  password: "ada",
  seeDebtors: "Yes"
});

db.upsertOrder({
  order_number: "S260301",
  status: "Not Yet Started",
  type: "Cabinet",
  category: "Bedroom",
  product: "Blaire Dresser",
  powder_coating: "Ferrograin Black",
  client_name: "Pat Client",
  price_excl_vat: "100.00"
});

(async function main() {
  const created = comments.createComment({
    order_number: "S260301",
    field: "powder_coating",
    body: "Confirm colour with paint shop",
    assignee: "Office Ada"
  }, "Shaka");
  assert.ok(created.id);
  assert.strictEqual(created.field, "powder_coating");
  assert.strictEqual(created.assignee, "Office Ada");
  assert.strictEqual(created.resolved, false);

  let failed = false;
  try {
    comments.createComment({
      order_number: "S260301",
      field: "powder_coating",
      body: "No assignee",
      assignee: ""
    }, "Shaka");
  } catch (e) {
    failed = /assign/i.test(e.message);
  }
  assert.ok(failed, "assignee is required");

  failed = false;
  try {
    comments.createComment({
      order_number: "S260301",
      field: "_edit",
      body: "Bad field",
      assignee: "Office Ada"
    }, "Shaka");
  } catch (e) {
    failed = /column|field|Pick/i.test(e.message);
  }
  assert.ok(failed, "icon column is not commentable");

  const summary = comments.cellSummary();
  assert.strictEqual(summary.length, 1);
  assert.strictEqual(summary[0].field, "powder_coating");
  assert.strictEqual(summary[0].count, 1);

  comments.replyToComment(created.id, { body: "Called them" }, "Office Ada");
  const resolved = comments.resolveComment(created.id, "Shaka");
  assert.strictEqual(resolved.resolved, true);
  assert.strictEqual(comments.cellSummary().length, 0);

  comments.reopenComment(created.id, "Shaka");
  assert.strictEqual(comments.cellSummary().length, 1);

  const app = express();
  app.use(express.json({ limit: "2mb" }));
  mountOffice(app);
  const server = http.createServer(app);
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;

  function req(method, urlPath, body, token) {
    return new Promise((resolve, reject) => {
      const data = body == null ? null : JSON.stringify(body);
      const r = http.request({
        hostname: "127.0.0.1",
        port,
        path: urlPath,
        method,
        headers: Object.assign(
          { "Content-Type": "application/json", "x-sd-token": token || "" },
          data ? { "Content-Length": Buffer.byteLength(data) } : {}
        )
      }, (res) => {
        const chunks = [];
        res.on("data", (c) => chunks.push(c));
        res.on("end", () => {
          const text = Buffer.concat(chunks).toString("utf8");
          let json = {};
          try { json = JSON.parse(text); } catch (e) {}
          resolve({ status: res.statusCode, json });
        });
      });
      r.on("error", reject);
      if (data) r.write(data);
      r.end();
    });
  }

  const login = await req("POST", "/api/office/login", { name: "Shaka", password: "admin" });
  assert.ok(login.json.ok, "login");
  const token = login.json.token || (login.json.profile && login.json.profile.token);
  assert.ok(token, "token");

  const list = await req("GET", "/api/office/order-comments?order=S260301&field=powder_coating", null, token);
  assert.ok(list.json.ok);
  assert.ok(list.json.rows.length >= 1);

  const posted = await req("POST", "/api/office/order-comments", {
    order_number: "S260301",
    field: "status",
    body: "Move after drawing lands",
    assignee: "Office Ada"
  }, token);
  assert.ok(posted.json.ok, posted.json.error || "post comment");
  assert.strictEqual(posted.json.comment.field, "status");

  const sumApi = await req("GET", "/api/office/order-comments/summary", null, token);
  assert.ok(sumApi.json.ok);
  assert.ok(sumApi.json.cells.some((c) => c.field === "status"));

  const del = await req("DELETE", "/api/office/orders/S260301", null, token);
  assert.ok(del.json.ok, del.json.error || "delete order");
  assert.strictEqual(comments.listComments({ order: "S260301" }).rows.length, 0);

  const ordersHtml = fs.readFileSync(path.join(__dirname, "../public/orders.html"), "utf8");
  assert.ok(ordersHtml.indexOf("cell-commentable") !== -1);
  assert.ok(ordersHtml.indexOf("order-comments") !== -1);
  assert.ok(ordersHtml.indexOf("cellCtxComment") !== -1);
  assert.ok(ordersHtml.indexOf("data-open-comments") !== -1);
  assert.ok(ordersHtml.indexOf("Assign to") !== -1);

  server.close();
  console.log("order-cell-comments.test.js ok");
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
