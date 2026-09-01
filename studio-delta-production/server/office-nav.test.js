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
assert.ok(js.indexOf("sdWarnPersistence") !== -1);
assert.ok(css.indexOf(".sd-persist-banner") !== -1);

const processJs = fs.readFileSync(path.join(__dirname, "../public/enquiry-process.js"), "utf8");
assert.ok(processJs.indexOf("productNamesLine") !== -1);
assert.ok(processJs.indexOf("Enter each product value") !== -1);
assert.ok(processJs.indexOf('complete_cost_sheet') !== -1);
assert.ok(!/complete_cost_sheet[\s\S]{0,200}valuesTable/.test(processJs), "cost sheet must not ask for product values");
assert.ok(processJs.indexOf("quote_assignee") !== -1);
assert.ok(processJs.indexOf("Request approval from (optional)") !== -1);
assert.ok(processJs.indexOf("Quoting person") !== -1);
assert.ok(processJs.indexOf("quote_no") !== -1);
assert.ok(processJs.indexOf("Last quotation numbers") !== -1);

const enquiriesHtml = fs.readFileSync(path.join(__dirname, "../public/enquiries.html"), "utf8");
assert.ok(enquiriesHtml.indexOf('id="captureMask"') !== -1, "New enquiry must open a capture modal");
assert.ok(enquiriesHtml.indexOf("openCapture") !== -1);

const ordersHtml = fs.readFileSync(path.join(__dirname, "../public/orders.html"), "utf8");
assert.ok(ordersHtml.indexOf("Import from Sheets") === -1, "Orders must not import from Google Sheets");
const usersHtml = fs.readFileSync(path.join(__dirname, "../public/users.html"), "utf8");
assert.ok(usersHtml.indexOf("Download backup") !== -1);
assert.ok(usersHtml.indexOf("migrate-from-google") !== -1);
assert.ok(floor.indexOf('href="/enquiries"') !== -1);
assert.ok(floor.indexOf('href="/tasks"') !== -1);
assert.ok(floor.indexOf("resumeOfficeSession") !== -1);
assert.ok(floor.indexOf("requestedView") !== -1);
assert.ok(floor.indexOf("sidebar-scroll") !== -1);

console.log("office-nav.test.js ok");
