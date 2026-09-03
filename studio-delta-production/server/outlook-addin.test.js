const fs = require("fs");
const path = require("path");
const assert = require("assert");
const { manifestXml } = require("./outlook-addin");

const xml = manifestXml("https://example.test");
assert.ok(xml.indexOf("xsi:type=\"MailApp\"") !== -1);
assert.ok(xml.indexOf("https://example.test/outlook-addin/taskpane.html") !== -1);
assert.ok(xml.indexOf("ReadItem") !== -1);
assert.ok(xml.indexOf("Studio Delta") !== -1);

const pane = fs.readFileSync(path.join(__dirname, "../public/outlook-addin/taskpane.html"), "utf8");
assert.ok(pane.indexOf("appsforoffice.microsoft.com") !== -1);
assert.ok(pane.indexOf("/api/office/enquiries/") !== -1);
assert.ok(pane.indexOf("outlook-mail") !== -1);
assert.ok(pane.indexOf("sdOfficeProfile") !== -1);
assert.ok(pane.indexOf("sdShowLogin") !== -1);
assert.ok(pane.indexOf("sdForgetOffice") === -1);

const processJs = fs.readFileSync(path.join(__dirname, "../public/enquiry-process.js"), "utf8");
assert.ok(processJs.indexOf("Paste the file link") !== -1);
assert.ok(processJs.indexOf("Correspondance link") !== -1);
assert.ok(processJs.indexOf("Drop the Outlook email") === -1);
assert.ok(processJs.indexOf("correspondence_files") === -1);

console.log("outlook-addin.test.js ok");
