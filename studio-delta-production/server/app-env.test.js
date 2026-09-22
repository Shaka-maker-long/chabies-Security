"use strict";

const assert = require("assert");
const fs = require("fs");
const path = require("path");

delete process.env.APP_ENV;
delete process.env.SD_APP_ENV;
delete process.env.RAILWAY_ENVIRONMENT_NAME;
delete process.env.RAILWAY_ENVIRONMENT;
delete process.env.STAGING;
delete require.cache[require.resolve("./app-env")];
let appEnv = require("./app-env");
assert.strictEqual(appEnv.isStaging(), false);
assert.ok(appEnv.appEnv() === "production" || appEnv.appEnv() === "development");

process.env.APP_ENV = "staging";
delete require.cache[require.resolve("./app-env")];
appEnv = require("./app-env");
assert.strictEqual(appEnv.isStaging(), true);
assert.strictEqual(appEnv.appEnv(), "staging");
assert.ok(/STAGING/.test(appEnv.stagingBannerText()));
assert.ok(/Admin/.test(appEnv.stagingBannerText()));

process.env.APP_ENV = "production";
process.env.STAGING = "1";
delete require.cache[require.resolve("./app-env")];
appEnv = require("./app-env");
assert.strictEqual(appEnv.isStaging(), true);

const healthJs = fs.readFileSync(path.join(__dirname, "index.js"), "utf8");
assert.ok(healthJs.indexOf("isStaging") !== -1);
assert.ok(healthJs.indexOf("appEnv") !== -1);

const pwa = fs.readFileSync(path.join(__dirname, "..", "public", "sd-pwa.js"), "utf8");
assert.ok(pwa.indexOf("sdStagingBanner") !== -1);
assert.ok(pwa.indexOf("bootStagingBanner") !== -1);

console.log("app-env.test.js ok");
