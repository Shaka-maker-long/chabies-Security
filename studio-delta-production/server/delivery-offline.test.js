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

assert.ok(off.isDriverProfile({ jobTitle: "Driver", tasks: ["Delivery"] }));
assert.ok(off.isDriverProfile({ tasks: ["Delivery"] }));
assert.ok(!off.isDriverProfile({ jobTitle: "Driver", canSeeOffice: true, tasks: ["Delivery"] }));
assert.ok(!off.isDriverProfile({ jobTitle: "Welder", tasks: ["Welding"] }));

off.clearDeliverySession();
assert.strictEqual(off.loadDeliverySession(), null);

async function runOfflineLoginTests() {
  off.clearDeliverySession();
  const savedOffice = await off.saveDeliverySession({
    name: "Admin",
    jobTitle: "Driver",
    canSeeOffice: true,
    tasks: ["Delivery"],
    token: "office"
  }, "admin");
  assert.strictEqual(savedOffice, false);
  assert.strictEqual(off.loadDeliverySession(), null);

  const savedWelder = await off.saveDeliverySession({
    name: "John",
    jobTitle: "Welder",
    tasks: ["Welding"],
    token: "w1"
  }, "1234");
  assert.strictEqual(savedWelder, false);

  const saved = await off.saveDeliverySession({
    name: "Sipho",
    jobTitle: "Driver",
    tasks: ["Delivery"],
    token: "tok-1"
  }, "1234");
  assert.strictEqual(saved, true);
  const stored = off.loadDeliverySession();
  assert.ok(stored);
  assert.ok(stored.pinHash);
  assert.ok(!JSON.stringify(stored).includes("1234"));
  assert.strictEqual(stored.nameKey, "sipho");

  const unlocked = await off.unlockDeliverySession("sipho", "1234");
  assert.ok(unlocked);
  assert.strictEqual(unlocked.name, "Sipho");
  assert.strictEqual(unlocked.token, "tok-1");
  assert.strictEqual(unlocked.canSeeOffice, false);

  assert.strictEqual(await off.unlockDeliverySession("Sipho", "9999"), null);
  assert.strictEqual(await off.unlockDeliverySession("Lebo", "1234"), null);

  const hashA = await off.hashPin("Sipho", "1234");
  const hashB = await off.hashPin("sipho", "1234");
  assert.strictEqual(hashA, hashB);

  assert.ok(/No signal/.test(off.offlineLoginMessage({ hasSaved: true })));
  assert.ok(/Log in once with data/.test(off.offlineLoginMessage({})));
  assert.ok(!/Connection error/.test(off.offlineLoginMessage({})));

  off.clearDeliverySession();
  assert.strictEqual(off.loadDeliverySession(), null);
}

runOfflineLoginTests().then(() => {
  console.log("delivery-offline.test.js ok");
}).catch((err) => {
  console.error(err);
  process.exit(1);
});
