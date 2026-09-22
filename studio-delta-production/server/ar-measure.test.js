const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const src = fs.readFileSync(path.join(__dirname, "../public/ar-measure.js"), "utf8");
const sandbox = { console };
sandbox.window = sandbox;
sandbox.globalThis = sandbox;
vm.runInNewContext(src, sandbox);

assert.strictEqual(typeof sandbox.sdArRoundMm, "function");
assert.strictEqual(typeof sandbox.sdArDistMeters, "function");
assert.strictEqual(typeof sandbox.sdOpenArMeasure, "function");
assert.strictEqual(sandbox.sdArRoundMm(1.2), 1200);
assert.strictEqual(sandbox.sdArRoundMm(0.001), 1);
assert.strictEqual(sandbox.sdArRoundMm(0), 0);
assert.ok(Math.abs(sandbox.sdArDistMeters({ x: 0, y: 0, z: 0 }, { x: 1, y: 0, z: 0 }) - 1) < 1e-9);
assert.strictEqual(sandbox.sdArRoundMm(sandbox.sdArDistMeters({ x: 0, y: 0, z: 0 }, { x: 1.234, y: 0, z: 0 })), 1234);

const indexHtml = fs.readFileSync(path.join(__dirname, "../index.html"), "utf8");
assert.ok(indexHtml.indexOf("/ar-measure.js") !== -1);
assert.ok(indexHtml.indexOf("openQcArMeasure") !== -1);
assert.ok(indexHtml.indexOf("Overall size check (phone AR)") !== -1);
assert.ok(indexHtml.indexOf("openQcArMeasure('glassHeight'") === -1);
assert.ok(indexHtml.indexOf("openQcArMeasure('woodHeight'") === -1);

const indexJs = fs.readFileSync(path.join(__dirname, "index.js"), "utf8");
assert.ok(indexJs.indexOf("/ar-measure.js") !== -1);

console.log("ar-measure.test.js ok");
