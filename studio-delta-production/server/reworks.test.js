"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sd-rework-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";

const db = require("./db");
const staff = require("./staff");
const reworks = require("./reworks");

staff.upsertUser({ name: "Manager", access: "Admin", role: "Manager", password: "m", manageUsers: "Yes" });
staff.upsertUser({ name: "Siya", access: "Admin", role: "Production Manager", password: "s" });
staff.upsertUser({ name: "Pat", access: "Admin", role: "Office", password: "p" });
staff.upsertUser({
  name: "Willard",
  access: "Production",
  role: "Welder",
  password: "w",
  tasks: ["Welding", "Rework"]
});
staff.upsertUser({
  name: "Sam",
  access: "Production",
  role: "Cutter",
  password: "x",
  tasks: ["Profile Cutting"]
});

db.upsertOrder({
  order_number: "S260901",
  status: "Ready for Delivery",
  type: "Catalogue",
  category: "Cabinet",
  product: "Ella Arched Cabinet",
  variation: "Left",
  doors: "2",
  detailed_description: "Arched doors with brass pulls",
  dimensions: "1800 x 600",
  powder_coating: "Matt black"
});

assert.ok(reworks.canAssignReworks(staff.listUsers().find((u) => u.name === "Manager")));
assert.ok(reworks.canAssignReworks(staff.listUsers().find((u) => u.name === "Siya")));
assert.ok(!reworks.canAssignReworks(staff.listUsers().find((u) => u.name === "Pat")));

assert.throws(
  () => reworks.createRework({ order_number: "S260901" }, "Pat"),
  /comment|issue/i
);

const created = reworks.createRework({
  order_number: "S260901",
  issue: "Left door out of square — reweld hinge side"
}, "Pat");
assert.ok(created.id);
assert.strictEqual(created.status, "Needs assignment");
assert.strictEqual(created.product, "Ella Arched Cabinet");
assert.strictEqual(created.category, "Cabinet");
assert.strictEqual(created.powder_coating, "Matt black");
assert.strictEqual(created.assigned_operator, "");
assert.ok(created.issue.indexOf("out of square") !== -1);

assert.throws(
  () => reworks.createRework({ order_number: "S260901", issue: "Again" }, "Pat"),
  /already has an open rework/i
);

assert.throws(
  () => reworks.assignRework(created.id, { assigned_operator: "Willard" }, "Pat", staff.listUsers().find((u) => u.name === "Pat")),
  /Siya or the Manager/i
);

const assigned = reworks.assignRework(
  created.id,
  { assigned_operator: "Willard" },
  "Siya",
  staff.listUsers().find((u) => u.name === "Siya")
);
assert.strictEqual(assigned.status, "Ready for Rework");
assert.strictEqual(assigned.assigned_operator, "Willard");

const forWillard = reworks.listOpenForFloor("Willard");
assert.strictEqual(forWillard.length, 1);
assert.strictEqual(forWillard[0].order_number, "S260901");
assert.strictEqual(reworks.listOpenForFloor("Sam").length, 0);
assert.ok(reworks.listOpenForFloor("").some((r) => r.order_number === "S260901"));

const started = reworks.markReworkStarted(created.id, "Willard");
assert.strictEqual(started.status, "Rework");

const done = reworks.markReworkFinished(created.id, "Willard");
assert.strictEqual(done.status, "Done");
assert.strictEqual(reworks.listOpenForFloor("Willard").length, 0);

const page = reworks.pagePayload();
assert.ok(page.orders.some((o) => o.order_number === "S260901"));
assert.ok(page.operators.indexOf("Willard") !== -1);
assert.ok(page.reworks.some((r) => r.id === created.id && r.status === "Done"));

assert.ok(staff.FLOOR_TASKS.indexOf("Rework") !== -1);

console.log("reworks.test.js ok");
