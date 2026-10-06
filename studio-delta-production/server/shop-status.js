"use strict";

const WAITING_FOR_DRAWING = "Waiting for drawing";

const SHOP_STATUSES = [
  WAITING_FOR_DRAWING,
  "Not Yet Started",
  "Ready for Steelwork", "Profile Cutting",
  "Ready for Tagging", "Tagging",
  "Ready for Welding", "Welding",
  "Ready for Grinding", "Grinding",
  "Ready for Pre-Powder Coating", "Pre-Powder Coating",
  "Ready for Powder Coating", "Sent to Paint Shop", "Paint Shop", "Powder Coating",
  "Ready for Assembly", "Assembly", "Paint Preparation", "Ready for Painting", "Painting",
  "Ready for Final QC", "Final QC",
  "Ready for Delivery", "Out for Delivery", "At couriers",
  "Delivered"
];

const PLANNED_PROCESSES = [
  "Profile Cutting",
  "Tagging",
  "Plate Cutting",
  "Welding",
  "Grinding",
  "Pre-Powder Coating",
  "Powder coating",
  "Upholstery",
  "Assembly",
  "Final QC"
];

const STATUS_ALIASES = {
  "paint shop": "Paint Shop",
  "at paint shop": "Paint Shop",
  "at the paint shop": "Paint Shop",
  "paintshop": "Paint Shop",
  "powder coaters": "Paint Shop",
  "at powder coaters": "Paint Shop",
  "at the powder coaters": "Paint Shop",
  "at courier": "At couriers",
  "at couriers": "At couriers",
  "at the courier": "At couriers",
  "at the couriers": "At couriers",
  "with courier": "At couriers",
  "with couriers": "At couriers",
  "at 3rd party": "At couriers",
  "at third party": "At couriers",
  "third party": "At couriers",
  "3rd party": "At couriers"
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

function isWaitingForDrawing(status) {
  return normalizeShopStatus(status) === WAITING_FOR_DRAWING;
}

function remainingPlanForStatus(status, productMinutes) {
  const raw = normalizeShopStatus(status) || "Not Yet Started";
  const s = isWaitingForDrawing(raw) ? "Not Yet Started" : raw;
  const skip = {};
  function done() {
    Array.prototype.forEach.call(arguments, (name) => { skip[name] = true; });
  }
  // Only skip a station after the floor has left it. The current
  // in-progress status (Welding, Assembly, …) still needs a plan.
  if (atOrAfter(s, "Ready for Tagging")) done("Profile Cutting");
  if (atOrAfter(s, "Ready for Welding")) done("Tagging");
  if (atOrAfter(s, "Welding")) done("Plate Cutting");
  if (atOrAfter(s, "Ready for Grinding")) done("Welding");
  if (atOrAfter(s, "Ready for Grinding")) done("Plate Cutting");
  if (atOrAfter(s, "Ready for Pre-Powder Coating")) done("Grinding");
  if (atOrAfter(s, "Ready for Powder Coating")) done("Pre-Powder Coating");
  if (atOrAfter(s, "Ready for Assembly")) done("Powder coating");
  if (atOrAfter(s, "Ready for Final QC")) done("Assembly");
  if (atOrAfter(s, "Ready for Final QC")) done("Upholstery");
  if (atOrAfter(s, "Ready for Delivery")) done("Final QC");
  // Optional paint loop after Ready for Assembly — only clears assembly when that loop is active.
  if (s === "Paint Preparation" || s === "Ready for Painting" || s === "Painting") {
    done("Assembly");
    done("Upholstery");
    done("Final QC");
  }

  const mins = productMinutes && typeof productMinutes === "object" ? productMinutes : null;
  const processes = PLANNED_PROCESSES.filter((p) => {
    if (skip[p]) return false;
    if (p === "Powder coating") return false; // modelled via paintWait + Monday PC block
    if (p === "Upholstery") {
      return !!(mins && Number(mins.Upholstery) > 0);
    }
    return true;
  });
  const paintWait = (
    processes.indexOf("Assembly") !== -1
    || processes.indexOf("Final QC") !== -1
    || processes.indexOf("Upholstery") !== -1
  ) && !atOrAfter(s, "Ready for Powder Coating");
  return { processes, paintWait, status: raw };
}

function isShopStatus(status) {
  return SHOP_STATUSES.indexOf(normalizeShopStatus(status)) !== -1;
}

function isAtPaintShop(status) {
  const s = normalizeShopStatus(status);
  return s === "Paint Shop" || s === "Sent to Paint Shop";
}

function isAtCouriers(status) {
  return normalizeShopStatus(status) === "At couriers";
}

function isHandedOff(status) {
  const s = normalizeShopStatus(status);
  return s === "Out for Delivery" || s === "At couriers" || s === "Delivered";
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
  WAITING_FOR_DRAWING,
  normalizeShopStatus,
  shopStatusIndex,
  remainingPlanForStatus,
  isShopStatus,
  isAtPaintShop,
  isAtCouriers,
  isHandedOff,
  isWaitingForDrawing,
  isProfileCuttingInProgress,
  waitingStatusIfProfileCuttingIdle
};
