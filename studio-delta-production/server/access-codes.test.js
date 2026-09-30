"use strict";

const assert = require("assert");
const codes = require("./access-codes");

const hashed = codes.hashPlain("secret1");
assert.ok(codes.isHashed(hashed));
assert.ok(codes.verify("secret1", hashed));
assert.ok(!codes.verify("secret2", hashed));
assert.ok(codes.verify("legacy", "legacy"));
assert.ok(codes.needsRehash("legacy"));
assert.ok(!codes.needsRehash(hashed));
assert.throws(() => codes.hashPlain(""), /required/);
assert.ok(codes.hashPlain("ab"));
console.log("access-codes.test.js ok");
