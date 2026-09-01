const fs = require("fs");
const path = require("path");
const assert = require("assert");

const js = fs.readFileSync(path.join(__dirname, "../public/office-auth.js"), "utf8");
const css = fs.readFileSync(path.join(__dirname, "../public/office-shell.css"), "utf8");
const floor = fs.readFileSync(path.join(__dirname, "../index.html"), "utf8");

const labels = [
  "Home", "Floor", "Orders", "Enquiries", "My tasks", "Office schedule", "Dropdowns", "Users",
  "Task times", "Debtors", "Production", "Workers", "Metrics",
  "QC Reports", "Activity", "Schedule", "Log Out"
];
labels.forEach((label) => {
  assert.ok(js.indexOf('"' + label + '"') !== -1 || js.indexOf(">" + label + "<") !== -1, "office menu missing " + label);
});

assert.ok(js.indexOf("/enquiries") !== -1);
assert.ok(js.indexOf("/tasks") !== -1);
assert.ok(js.indexOf("/?view=production") !== -1);
assert.ok(js.indexOf("/?view=workers") !== -1);
assert.ok(js.indexOf("/?view=metrics") !== -1);
assert.ok(js.indexOf("/?view=qc") !== -1);
assert.ok(js.indexOf("/?view=activity") !== -1);
assert.ok(js.indexOf("/?view=schedule") !== -1);
assert.ok(js.indexOf("sd-sidebar-scroll") !== -1);
assert.ok(js.indexOf("sdLogoutBtn") !== -1);
assert.ok(css.indexOf(".sd-sidebar-footer") !== -1);
assert.ok(floor.indexOf('href="/enquiries"') !== -1);
assert.ok(floor.indexOf('href="/tasks"') !== -1);
assert.ok(floor.indexOf("resumeOfficeSession") !== -1);
assert.ok(floor.indexOf("requestedView") !== -1);
assert.ok(floor.indexOf("sidebar-scroll") !== -1);

console.log("office-nav.test.js ok");
