"use strict";

const assert = require("assert");
const {
  normalizeShopStatus,
  remainingPlanForStatus,
  isShopStatus,
  waitingStatusIfProfileCuttingIdle
} = require("./shop-status");

assert.strictEqual(normalizeShopStatus("welding"), "Welding");
assert.strictEqual(normalizeShopStatus("READY FOR TAGGING"), "Ready for Tagging");
assert.ok(isShopStatus("assembly"));
assert.ok(isShopStatus("Waiting for drawing"));
assert.ok(!isShopStatus("In Progress"));

assert.deepStrictEqual(
  remainingPlanForStatus("Not Yet Started").processes,
  ["Profile Cutting", "Tagging", "Plate Cutting", "Welding", "Grinding", "Pre-Powder Coating", "Assembly", "Final QC"]
);
assert.deepStrictEqual(
  remainingPlanForStatus("Waiting for drawing").processes,
  remainingPlanForStatus("Not Yet Started").processes
);

assert.ok(remainingPlanForStatus("Profile Cutting").processes.indexOf("Profile Cutting") !== -1);
assert.ok(remainingPlanForStatus("Welding").processes.indexOf("Welding") !== -1);
assert.ok(remainingPlanForStatus("Welding").processes.indexOf("Plate Cutting") === -1, "plate finishes before welding");
assert.ok(remainingPlanForStatus("Ready for Welding").processes.indexOf("Plate Cutting") !== -1);
assert.ok(remainingPlanForStatus("Welding").processes.indexOf("Tagging") === -1);

assert.ok(remainingPlanForStatus("Ready for Grinding").processes.indexOf("Welding") === -1);
assert.ok(remainingPlanForStatus("Ready for Grinding").processes.indexOf("Grinding") !== -1);

assert.deepStrictEqual(remainingPlanForStatus("Assembly").processes, ["Assembly", "Final QC"]);
assert.ok(!remainingPlanForStatus("Assembly").paintWait, "already at assembly means paint already happened");
assert.deepStrictEqual(remainingPlanForStatus("Paint Preparation").processes, []);
assert.deepStrictEqual(remainingPlanForStatus("Final QC").processes, ["Final QC"]);
assert.ok(!remainingPlanForStatus("Not Yet Started", { Upholstery: 0 }).processes.includes("Upholstery"));
assert.ok(remainingPlanForStatus("Not Yet Started", { Upholstery: 60 }).processes.includes("Upholstery"));

assert.ok(remainingPlanForStatus("welding").processes.indexOf("Welding") !== -1);
assert.strictEqual(waitingStatusIfProfileCuttingIdle("Profile Cutting", false), "Ready for Steelwork");
assert.strictEqual(waitingStatusIfProfileCuttingIdle("profile cutting", true), "Profile Cutting");

assert.strictEqual(normalizeShopStatus("Paint Shop"), "Paint Shop");
assert.strictEqual(normalizeShopStatus("paint shop"), "Paint Shop");
assert.strictEqual(normalizeShopStatus("at the paint shop"), "Paint Shop");
assert.strictEqual(normalizeShopStatus("Sent to Paint Shop"), "Sent to Paint Shop");
assert.ok(isShopStatus("Paint Shop"));
assert.ok(isShopStatus("At couriers"));
assert.strictEqual(normalizeShopStatus("at courier"), "At couriers");
assert.strictEqual(normalizeShopStatus("At the couriers"), "At couriers");
assert.deepStrictEqual(remainingPlanForStatus("At couriers").processes, []);
assert.ok(!remainingPlanForStatus("Paint Shop").paintWait);
assert.ok(!remainingPlanForStatus("Sent to Paint Shop").paintWait);
assert.ok(remainingPlanForStatus("Paint Shop").processes.indexOf("Assembly") !== -1);

console.log("shop-status.test.js ok");
