"use strict";

const assert = require("assert");
const off = require("../public/delivery-offline");

assert.ok(off.isNetworkError(new Error("Failed to fetch")));
assert.ok(off.isNetworkError(new Error("NetworkError when attempting to fetch resource.")));
assert.ok(off.isNetworkError(null, 503));
assert.ok(off.isNetworkError(null, 0));
assert.ok(!off.isNetworkError("At least one photo is required.", 400));
assert.ok(!off.isNetworkError("Log in first.", 401));

const split = off.splitRunOrders([
  { order_number: "S260401 A", base: "S260401" },
  { order_number: "S260401 B", base: "S260401" },
  { order_number: "S260404", base: "S260404" }
], [{
  order_numbers: ["S260401 A", "S260401 B"],
  bases: ["S260401"]
}]);
assert.strictEqual(split.waiting.length, 2);
assert.strictEqual(split.live.length, 1);
assert.strictEqual(split.live[0].order_number, "S260404");

const keys = off.queuedOrderKeys([{ payload: { order_number: "s260410" } }]);
assert.ok(keys["S260410"]);

console.log("delivery-offline.test.js ok");
