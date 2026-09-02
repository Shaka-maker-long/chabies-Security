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

const officeJs = fs.readFileSync(path.join(__dirname, "office.js"), "utf8");
assert.ok(officeJs.indexOf("listMyCompletedTasks") !== -1);
assert.ok(officeJs.indexOf("req.query.done") !== -1);
const indexJs = fs.readFileSync(path.join(__dirname, "index.js"), "utf8");
assert.ok(indexJs.indexOf("/tasks/completed") !== -1);
const tasksHtml = fs.readFileSync(path.join(__dirname, "../public/tasks.html"), "utf8");
assert.ok(tasksHtml.indexOf("/tasks/completed") !== -1);
assert.ok(tasksHtml.indexOf(">To do<") !== -1);
assert.ok(tasksHtml.indexOf(">Completed<") !== -1);
assert.ok(tasksHtml.indexOf("done=1") !== -1);
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
assert.ok(processJs.indexOf("Issue another quote") !== -1 || processJs.indexOf("Previous quote") !== -1);
assert.ok(processJs.indexOf("Client wants changes") !== -1 || processJs.indexOf("recost") !== -1);
assert.ok(processJs.indexOf("Last quotation numbers") !== -1);
assert.ok(processJs.indexOf("CORRESPONDANCE") !== -1);
assert.ok(processJs.indexOf("outlook_drop") !== -1);
assert.ok(processJs.indexOf("Drop the Outlook email") !== -1);
assert.ok(processJs.indexOf("data-open-file") !== -1);
assert.ok(processJs.indexOf("Files") !== -1);
assert.ok(processJs.indexOf("Every CORRESPONDANCE email") !== -1);
assert.ok(processJs.indexOf("earlier quotes") !== -1);

assert.ok(processJs.indexOf("sd-file-row") !== -1);
assert.ok(processJs.indexOf("sd-timeline") !== -1);
assert.ok(processJs.indexOf("Lifespan") !== -1);
const enquiriesHtml = fs.readFileSync(path.join(__dirname, "../public/enquiries.html"), "utf8");
assert.ok(enquiriesHtml.indexOf('id="captureMask"') !== -1, "New enquiry must open a capture modal");
assert.ok(enquiriesHtml.indexOf("#grid {") !== -1 || enquiriesHtml.indexOf("#grid{") !== -1);
assert.ok(enquiriesHtml.indexOf("table { border-collapse:separate; border-spacing:0; min-width:3200px") === -1, "enquiry grid min-width must not apply to all tables");
assert.ok(enquiriesHtml.indexOf("openCapture") !== -1);
assert.ok(enquiriesHtml.indexOf("data-edit-enquiry") !== -1);
assert.ok(enquiriesHtml.indexOf("Edit enquiry") !== -1);
assert.ok(enquiriesHtml.indexOf("onclick=\"saveAll()\"") === -1, "enquiries sheet must not save by editing cells");
assert.ok(enquiriesHtml.indexOf("table.oninput") === -1, "grid must not edit cells in place");
assert.ok(enquiriesHtml.indexOf("OPENED") !== -1);
assert.ok(enquiriesHtml.indexOf("LIFESPAN") !== -1);
assert.ok(enquiriesHtml.indexOf("/enquiries/dashboard") !== -1, "sheet must link to the dashboard");
assert.ok(enquiriesHtml.indexOf(">Dashboard<") !== -1);

const dashHtml = fs.readFileSync(path.join(__dirname, "../public/enquiries-dashboard.html"), "utf8");
assert.ok(dashHtml.indexOf("/enquiries") !== -1);
assert.ok(dashHtml.indexOf("Weekly") !== -1);
assert.ok(dashHtml.indexOf("Monthly") !== -1);
assert.ok(dashHtml.indexOf("Revenue") !== -1);
assert.ok(dashHtml.indexOf("chart.js") !== -1);
assert.ok(dashHtml.indexOf("/api/office/enquiries/dashboard") !== -1);
assert.ok(dashHtml.indexOf("sdOfficeFetch") !== -1);
assert.ok(dashHtml.indexOf("CATERGORY") !== -1);
assert.ok(dashHtml.indexOf("id=\"month\"") !== -1);
assert.ok(dashHtml.indexOf("dashboard/drill") !== -1);
assert.ok(dashHtml.indexOf("sdOpenEnquiryProcess") !== -1);
assert.ok(dashHtml.indexOf("enquiry-process.js") !== -1);
assert.ok(dashHtml.indexOf("Quote value by type") !== -1);
assert.ok(dashHtml.indexOf("Week of month") !== -1);
assert.ok(dashHtml.indexOf("prodSearch") !== -1);
assert.ok(dashHtml.indexOf(">Outlook<") === -1, "Outlook pie must be replaced by quote value by type");

const ordersHtml = fs.readFileSync(path.join(__dirname, "../public/orders.html"), "utf8");
assert.ok(ordersHtml.indexOf("Import from Sheets") === -1, "Orders must not import from Google Sheets");
const usersHtml = fs.readFileSync(path.join(__dirname, "../public/users.html"), "utf8");
assert.ok(usersHtml.indexOf("Download JSON backup") !== -1);
assert.ok(usersHtml.indexOf("Download SQLite") !== -1);
assert.ok(usersHtml.indexOf("Backup now") !== -1);
assert.ok(usersHtml.indexOf("/api/office/backups/run") !== -1);
assert.ok(usersHtml.indexOf("backup.db") !== -1);
assert.ok(usersHtml.indexOf("migrate-from-google") !== -1);
assert.ok(floor.indexOf('href="/enquiries"') !== -1);
assert.ok(floor.indexOf('href="/tasks"') !== -1);
assert.ok(floor.indexOf("resumeOfficeSession") !== -1);
assert.ok(floor.indexOf("requestedView") !== -1);
assert.ok(floor.indexOf("sidebar-scroll") !== -1);
assert.ok(floor.indexOf("sdShowWelcome") !== -1);
assert.ok(floor.indexOf("welcome: true") !== -1);
assert.ok(floor.indexOf("sd-brand.css") !== -1);
assert.ok(floor.indexOf("sd-splash.js") !== -1);
assert.ok(floor.indexOf("paintBrandMarks") !== -1);
assert.ok(js.indexOf("sdShowWelcome") !== -1);
assert.ok(js.indexOf("sdSMark") !== -1);
assert.ok(js.indexOf("sdLoadBrand") !== -1);
assert.ok(js.indexOf("sd-login-card") !== -1);
assert.ok(indexJs.indexOf("/sd-brand.css") !== -1);
assert.ok(indexJs.indexOf("/sd-splash.js") !== -1);
const splash = fs.readFileSync(path.join(__dirname, "../public/sd-splash.js"), "utf8");
assert.ok(splash.indexOf("Welcome to Studio Delta") !== -1);
assert.ok(splash.indexOf("sd-s") !== -1);
const brand = fs.readFileSync(path.join(__dirname, "../public/sd-brand.css"), "utf8");
assert.ok(brand.indexOf(".sd-welcome") !== -1);
assert.ok(brand.indexOf("#b08948") !== -1);

console.log("office-nav.test.js ok");
