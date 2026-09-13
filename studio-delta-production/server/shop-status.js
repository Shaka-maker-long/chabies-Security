"use strict";

const SHOP_STATUSES = [
  "Not Yet Started",
  "Ready for Steelwork", "Profile Cutting",
  "Ready for Tagging", "Tagging",
  "Ready for Welding", "Welding",
  "Ready for Grinding", "Grinding",
  "Ready for Pre-Powder Coating", "Pre-Powder Coating",
  "Ready for Powder Coating", "Sent to Paint Shop", "Paint Shop", "Powder Coating",
  "Ready for Assembly", "Assembly", "Paint Preparation", "Ready for Painting", "Painting",
  "Ready for Final QC", "Final QC",
  "Ready for Delivery", "Out for Delivery",
  "Delivered"
];

const PLANNED_PROCESSES = [
  "Profile Cutting",
  "Tagging",
  "Plate Cutting",
  "Welding",
  "Grinding",
  "Assembly"
];

const STATUS_ALIASES = {
  "paint shop": "Paint Shop",
  "at paint shop": "Paint Shop",
  "at the paint shop": "Paint Shop",
  "paintshop": "Paint Shop",
  "powder coaters": "Paint Shop",
  "at powder coaters": "Paint Shop",
  "at the powder coaters": "Paint Shop"
};

function normalizeShopStatus(status) {
  const raw = String(status == null ? "" : status).trim();
  if (!raw) return "";
  const compact = raw.toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ");
  if (STATUS_ALIASES[compact]) return STATUS_ALIASES[compact];
  const hit = SHOP_STATUSES.find((name) => name.toLowerCase() === compact);
  return hit || raw;
}

function shopStatusIndex(status) {
  const i = SHOP_STATUSES.indexOf(normalizeShopStatus(status));
  return i < 0 ? 0 : i;
}

function atOrAfter(status, name) {
  return shopStatusIndex(status) >= shopStatusIndex(name);
}

function remainingPlanForStatus(status) {
  const s = normalizeShopStatus(status) || "Not Yet Started";
  const skip = {};
  function done() {
    Array.prototype.forEach.call(arguments, (name) => { skip[name] = true; });
  }
  // Only skip a station after the floor has left it. The current
  // in-progress status (Welding, Assembly, …) still needs a plan.
  if (atOrAfter(s, "Ready for Tagging")) done("Profile Cutting");
  if (atOrAfter(s, "Ready for Welding")) done("Tagging");
  if (atOrAfter(s, "Ready for Grinding")) done("Welding");
  if (atOrAfter(s, "Ready for Grinding")) done("Plate Cutting");
  if (atOrAfter(s, "Ready for Pre-Powder Coating")) done("Grinding");
  if (atOrAfter(s, "Paint Preparation")) done("Assembly");
  const processes = PLANNED_PROCESSES.filter((p) => !skip[p]);
  const paintWait = processes.indexOf("Assembly") !== -1 && !atOrAfter(s, "Ready for Powder Coating");
  return { processes, paintWait, status: s };
}

function isShopStatus(status) {
  return SHOP_STATUSES.indexOf(normalizeShopStatus(status)) !== -1;
}

function isAtPaintShop(status) {
  const s = normalizeShopStatus(status);
  return s === "Paint Shop" || s === "Sent to Paint Shop";
}

const PROFILE_CUTTING_WAITING = "Ready for Steelwork";

function isProfileCuttingInProgress(status, hasOpenProfileClock) {
  return normalizeShopStatus(status) === "Profile Cutting" && !!hasOpenProfileClock;
}

function waitingStatusIfProfileCuttingIdle(status, hasOpenProfileClock) {
  if (normalizeShopStatus(status) !== "Profile Cutting") return normalizeShopStatus(status) || status;
  if (hasOpenProfileClock) return "Profile Cutting";
  return PROFILE_CUTTING_WAITING;
}

module.exports = {
  SHOP_STATUSES,
  PLANNED_PROCESSES,
  STATUS_ALIASES,
  PROFILE_CUTTING_WAITING,
  normalizeShopStatus,
  shopStatusIndex,
  remainingPlanForStatus,
  isShopStatus,
  isAtPaintShop,
  isProfileCuttingInProgress,
  waitingStatusIfProfileCuttingIdle
};
