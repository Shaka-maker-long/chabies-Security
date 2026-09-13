"use strict";

const assert = require("assert");
const {
  normalizeShopStatus,
  remainingPlanForStatus,
  isShopStatus
} = require("./shop-status");

assert.strictEqual(normalizeShopStatus("welding"), "Welding");
assert.strictEqual(normalizeShopStatus("READY FOR TAGGING"), "Ready for Tagging");
assert.ok(isShopStatus("assembly"));
assert.ok(!isShopStatus("In Progress"));

assert.deepStrictEqual(
  remainingPlanForStatus("Not Yet Started").processes,
  ["Profile Cutting", "Tagging", "Plate Cutting", "Welding", "Grinding", "Assembly"]
);

assert.ok(remainingPlanForStatus("Profile Cutting").processes.indexOf("Profile Cutting") !== -1);
assert.ok(remainingPlanForStatus("Welding").processes.indexOf("Welding") !== -1);
assert.ok(remainingPlanForStatus("Welding").processes.indexOf("Plate Cutting") !== -1);
assert.ok(remainingPlanForStatus("Welding").processes.indexOf("Profile Cutting") === -1);
assert.ok(remainingPlanForStatus("Welding").processes.indexOf("Tagging") === -1);

assert.ok(remainingPlanForStatus("Ready for Grinding").processes.indexOf("Welding") === -1);
assert.ok(remainingPlanForStatus("Ready for Grinding").processes.indexOf("Grinding") !== -1);

assert.deepStrictEqual(remainingPlanForStatus("Assembly").processes, ["Assembly"]);
assert.ok(!remainingPlanForStatus("Assembly").paintWait, "already at assembly means paint already happened");
assert.deepStrictEqual(remainingPlanForStatus("Paint Preparation").processes, []);
assert.deepStrictEqual(remainingPlanForStatus("Final QC").processes, []);

assert.ok(remainingPlanForStatus("welding").processes.indexOf("Welding") !== -1);

console.log("shop-status.test.js ok");
