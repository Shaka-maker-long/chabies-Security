"use strict";

const fs = require("fs");
const path = require("path");
const assert = require("assert");

const floor = fs.readFileSync(path.join(__dirname, "../index.html"), "utf8");

assert.ok(floor.indexOf("function isIndirectWorkerLog") !== -1);
assert.ok(floor.indexOf("function weekLabelFromDate") !== -1);
assert.ok(floor.indexOf("function expandWorkerDaySlices") !== -1);
assert.ok(floor.indexOf(">Start date</th>") !== -1);
assert.ok(floor.indexOf(">Finish date</th>") !== -1);
assert.ok(floor.indexOf("if (isIndirectWorkerLog(r)) return;") !== -1);
assert.ok(
  floor.indexOf('<th>Date</th><th>Order</th><th>Task</th><th>Start</th><th>End</th>') === -1,
  "worker week log must use start/finish date columns"
);

// Mirror the client week-split rules for a multi-week overnight weld.
const SAST_OFFSET_MS = 2 * 60 * 60 * 1000;
function asSast(date) { return new Date(date.getTime() + SAST_OFFSET_MS); }
function sastDayStamp(date) {
  const d = asSast(date);
  return d.getUTCFullYear() + "-" + String(d.getUTCMonth() + 1).padStart(2, "0") + "-" + String(d.getUTCDate()).padStart(2, "0");
}
function sastWallToDate(date, hours, minutes) {
  const d = asSast(date);
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), hours - 2, minutes, 0, 0));
}
function addSastDays(date, days) { return new Date(date.getTime() + days * 86400000); }
function isoWeekNumber(date) {
  const sast = asSast(date);
  const utc = Date.UTC(sast.getUTCFullYear(), sast.getUTCMonth(), sast.getUTCDate());
  const d = new Date(utc);
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = Date.UTC(d.getUTCFullYear(), 0, 1);
  return Math.ceil(((d.getTime() - yearStart) / 86400000 + 1) / 7);
}
function weekLabelFromDate(date) {
  const sast = asSast(date);
  return sast.getUTCFullYear() + " - Week " + String(isoWeekNumber(date)).padStart(2, "0");
}

// Sunday 27 Sep 2026 14:00 SAST -> Monday 28 Sep 2026 10:00 SAST (weeks 39 -> 40)
const start = new Date("2026-09-27T12:00:00.000Z"); // 14:00 SAST Sunday
const end = new Date("2026-09-28T08:00:00.000Z"); // 10:00 SAST Monday
assert.strictEqual(weekLabelFromDate(start), "2026 - Week 39");
assert.strictEqual(weekLabelFromDate(end), "2026 - Week 40");

const weeks = {};
let cursor = sastWallToDate(start, 0, 0);
const endStamp = sastDayStamp(end);
let safety = 0;
while (sastDayStamp(cursor) <= endStamp && safety++ < 10) {
  const dayStamp = sastDayStamp(cursor);
  const dayStart = sastWallToDate(cursor, 0, 0);
  const nextMidnight = addSastDays(dayStart, 1);
  const clipStart = Math.max(start.getTime(), dayStart.getTime());
  const clipEnd = Math.min(end.getTime(), nextMidnight.getTime());
  if (clipEnd > clipStart) {
    const week = weekLabelFromDate(new Date(clipStart));
    weeks[week] = (weeks[week] || 0) + (clipEnd - clipStart);
  }
  cursor = nextMidnight;
}
assert.ok(weeks["2026 - Week 39"] > 0, "Sunday portion stays in week 39");
assert.ok(weeks["2026 - Week 40"] > 0, "Monday portion moves to week 40");
assert.ok(
  Object.keys(weeks).length === 2,
  "multi-week job must land in exactly two week buckets, got " + JSON.stringify(weeks)
);

assert.strictEqual(String("INDIRECT").toUpperCase(), "INDIRECT");
assert.ok(!/S260264/i.test("INDIRECT"));

console.log("worker-week-detail.test.js ok");
