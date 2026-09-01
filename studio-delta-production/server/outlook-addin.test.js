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
assert.ok(pane.indexOf("Nothing is downloaded") !== -1 || pane.indexOf("nothing is downloaded") !== -1);

const processJs = fs.readFileSync(path.join(__dirname, "../public/enquiry-process.js"), "utf8");
assert.ok(processJs.indexOf("a.download") === -1);
assert.ok(processJs.indexOf("application/vnd.ms-outlook") === -1);

console.log("outlook-addin.test.js ok");
