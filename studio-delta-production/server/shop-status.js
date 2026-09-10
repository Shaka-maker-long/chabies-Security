"use strict";

const SHOP_STATUSES = [
  "Not Yet Started",
  "Ready for Steelwork", "Profile Cutting",
  "Ready for Tagging", "Tagging",
  "Ready for Welding", "Welding",
  "Ready for Grinding", "Grinding",
  "Ready for Pre-Powder Coating", "Pre-Powder Coating",
  "Ready for Powder Coating", "Sent to Paint Shop", "Powder Coating",
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

function shopStatusIndex(status) {
  const i = SHOP_STATUSES.indexOf(String(status || "").trim());
  return i < 0 ? 0 : i;
}

function atOrAfter(status, name) {
  return shopStatusIndex(status) >= shopStatusIndex(name);
}

function remainingPlanForStatus(status) {
  const s = String(status || "").trim() || "Not Yet Started";
  const skip = {};
  function done() {
    Array.prototype.forEach.call(arguments, (name) => { skip[name] = true; });
  }
  if (s === "Profile Cutting" || atOrAfter(s, "Ready for Tagging")) done("Profile Cutting");
  if (s === "Tagging" || atOrAfter(s, "Ready for Welding")) done("Tagging");
  if (s === "Welding" || atOrAfter(s, "Ready for Grinding")) done("Welding");
  if (atOrAfter(s, "Ready for Grinding")) done("Plate Cutting");
  if (s === "Grinding" || atOrAfter(s, "Ready for Pre-Powder Coating")) done("Grinding");
  if (s === "Assembly" || atOrAfter(s, "Paint Preparation")) done("Assembly");
  const processes = PLANNED_PROCESSES.filter((p) => !skip[p]);
  const paintWait = processes.indexOf("Assembly") !== -1 && !atOrAfter(s, "Ready for Powder Coating");
  return { processes, paintWait, status: s };
}

function isShopStatus(status) {
  return SHOP_STATUSES.indexOf(String(status || "").trim()) !== -1;
}

module.exports = {
  SHOP_STATUSES,
  PLANNED_PROCESSES,
  shopStatusIndex,
  remainingPlanForStatus,
  isShopStatus
};
