const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-hours-lock-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
process.env.WORK_LOCKS_DISABLED = "false";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const { callShopFunction, clearShopCache } = require("./gas");

initWorkbook();
const book = getBook();
const users = book.getSheetByName("Users");
users.appendRow(["Sam", "Welding", "1234", "Welding", "Production", "No"]);
users.appendRow(["John", "Welding", "1234", "Welding", "Production", "No"]);
users.appendRow(["Willard", "Welding", "1234", "Welding", "Production", "No"]);
persistWorkbook();

const morning = new Date("2026-09-15T07:45:00+02:00");
const lunch = new Date("2026-09-15T12:00:00+02:00");
const joinAt = new Date("2026-09-15T10:00:00+02:00");
const afternoon = new Date("2026-09-15T13:00:00+02:00");
const eightHoursEnd = new Date("2026-09-15T16:15:00+02:00");

function togetherMeta(extra) {
  return JSON.stringify(Object.assign({
    entryType: "production",
    batchId: "batch-hours-lock",
    batchShare: 2,
    pauses: []
  }, extra));
}

function soloMeta(extra) {
  return JSON.stringify(Object.assign({
    entryType: "production",
    batchId: "",
    batchShare: 1,
    pauses: [],
    overtimeContinue: false
  }, extra));
}

const logs = book.getSheetByName("Production_Log");
logs.appendRow([
  "log-together-a", "S-LOCK-A", "Sam", "Welding", "Welding",
  morning, lunch, "Complete", "", "", "", "",
  togetherMeta({ batchJoinedAt: morning.getTime(), batchSplitAt: lunch.getTime() })
]);
logs.appendRow([
  "log-together-b", "S-LOCK-B", "Sam", "Welding", "Welding",
  morning, lunch, "Complete", "", "", "", "",
  togetherMeta({ batchJoinedAt: morning.getTime(), batchSplitAt: lunch.getTime() })
]);
logs.appendRow([
  "log-prejoin-a", "S-JOIN-A", "John", "Welding", "Welding",
  morning, lunch, "Complete", "", "", "", "",
  togetherMeta({
    batchId: "batch-prejoin",
    batchJoinedAt: joinAt.getTime(),
    batchSplitAt: lunch.getTime()
  })
]);
logs.appendRow([
  "log-prejoin-b", "S-JOIN-B", "John", "Welding", "Welding",
  joinAt, lunch, "Complete", "", "", "", "",
  togetherMeta({
    batchId: "batch-prejoin",
    batchJoinedAt: joinAt.getTime(),
    batchSplitAt: lunch.getTime()
  })
]);
logs.appendRow([
  "log-eight", "S-EIGHT-1", "Willard", "Welding", "Welding",
  morning, eightHoursEnd, "Complete", "", "", "", "",
  soloMeta({ overtimeContinue: true })
]);
persistWorkbook();
clearShopCache();

(async function main() {
  const samMins = await callShopFunction("workerMinutesToday", ["Sam", afternoon]);
  assert.ok(samMins > 250 && samMins < 260, "two together morning jobs are one morning: " + samMins);
  assert.ok(samMins < 480, "together jobs must not trip the 8-hour counter: " + samMins);

  const samGate = await callShopFunction("floorChangeGate", ["Sam", "start", afternoon]);
  assert.strictEqual(samGate.ok, true, "Sam can start another job after together morning: " + JSON.stringify(samGate));
  assert.ok(!samGate.locked, JSON.stringify(samGate));

  const johnMins = await callShopFunction("workerMinutesToday", ["John", afternoon]);
  assert.ok(johnMins > 250 && johnMins < 260, "solo then together still counts one morning: " + johnMins);

  const johnGate = await callShopFunction("floorChangeGate", ["John", "start", afternoon]);
  assert.strictEqual(johnGate.ok, true, "John can start after joining a second job: " + JSON.stringify(johnGate));

  const willardMins = await callShopFunction("workerMinutesToday", ["Willard", afternoon]);
  assert.ok(willardMins >= 480, "a real 8-hour day still fills the counter: " + willardMins);
  const willardGate = await callShopFunction("floorChangeGate", ["Willard", "start", afternoon]);
  assert.strictEqual(willardGate.ok, false, JSON.stringify(willardGate));
  assert.ok(/8 hours/i.test(willardGate.message || ""), JSON.stringify(willardGate));

  console.log("together-hours-lock.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
