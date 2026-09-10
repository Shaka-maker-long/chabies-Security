const { isoWeekInfo } = require("./office-schedule");
const quoteOptions = require("./quote-options");

const ORDER_TYPES = ["Standard", "Custom", "New Design"];

const LEAD_WEEKS = {
  Standard: { gauteng: 5, other: 7 },
  Custom: { gauteng: 7, other: 8 },
  "New Design": { gauteng: 8, other: 10 }
};

function ymd(d) {
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
}

function dateAtNoon(iso) {
  const s = String(iso || "").slice(0, 10);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

function todayIso(now) {
  const d = now ? new Date(now) : new Date();
  if (Number.isNaN(d.getTime())) return ymd(new Date());
  return ymd(d);
}

function addDaysIso(iso, days) {
  const d = dateAtNoon(iso);
  if (!d) return "";
  d.setDate(d.getDate() + Number(days) || 0);
  return ymd(d);
}

function weekdayNum(iso) {
  const d = dateAtNoon(iso);
  return d ? d.getDay() : -1;
}

function isTueOrThu(iso) {
  const dow = weekdayNum(iso);
  return dow === 2 || dow === 4;
}

function isMonday(iso) {
  return weekdayNum(iso) === 1;
}

function mondayOfIsoWeek(iso) {
  const dow = weekdayNum(iso);
  if (dow < 0) return "";
  const back = dow === 0 ? -6 : 1 - dow;
  return addDaysIso(iso, back);
}

function isNextIsoWeek(fromIso, dateIso) {
  const fromMon = mondayOfIsoWeek(fromIso);
  const dateMon = mondayOfIsoWeek(dateIso);
  if (!fromMon || !dateMon) return false;
  return dateMon === addDaysIso(fromMon, 7);
}

function isNonGautengNextWeek(iso, province, today) {
  return !isGauteng(province) && isNextIsoWeek(today, iso);
}

function nextTueOrThu(iso) {
  let cur = String(iso || "").slice(0, 10);
  if (!dateAtNoon(cur)) return "";
  for (let i = 0; i < 14; i++) {
    if (isTueOrThu(cur)) return cur;
    cur = addDaysIso(cur, 1);
  }
  return "";
}

function snapToScheduleDay(iso, province, today) {
  const day = String(iso || "").slice(0, 10);
  const from = todayIso(today);
  if (!dateAtNoon(day)) return "";
  if (isNonGautengNextWeek(day, province, from)) return mondayOfIsoWeek(day);
  const snapped = nextTueOrThu(day);
  if (isNonGautengNextWeek(snapped, province, from)) return mondayOfIsoWeek(snapped);
  return snapped;
}

function prettyDate(iso) {
  const d = dateAtNoon(iso);
  if (!d) return "";
  const weekday = d.toLocaleDateString("en-GB", { weekday: "long" });
  const month = d.toLocaleDateString("en-GB", { month: "long" });
  return weekday + " " + d.getDate() + " " + month + " " + d.getFullYear();
}

function studioOrderSeq(orderNumber) {
  const m = String(orderNumber || "").trim().toUpperCase().match(/^S(\d+)(?:\s+[A-Z]+)?$/);
  return m ? Number(m[1]) : 0;
}

function normalizeBaseOrderNumber(raw) {
  const s = String(raw || "").trim().toUpperCase().replace(/\s+/g, " ");
  const m = s.match(/^S(\d+)(?:\s+[A-Z]+)?$/);
  return m ? "S" + m[1] : "";
}

function nextStudioOrderNumberFrom(orders, now) {
  let max = 0;
  (orders || []).forEach((o) => {
    max = Math.max(max, studioOrderSeq(o && o.order_number));
  });
  if (max < 1) {
    const y = (now ? new Date(now) : new Date()).getFullYear() % 100;
    max = y * 10000;
  }
  return "S" + String(max + 1);
}

function orderBaseTaken(base, orders) {
  const want = normalizeBaseOrderNumber(base);
  if (!want) return false;
  return (orders || []).some((o) => normalizeBaseOrderNumber(o && o.order_number) === want);
}

function splitLetter(index) {
  let n = Number(index) + 1;
  let s = "";
  while (n > 0) {
    n -= 1;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

function orderTypeFromEnquiryType(raw) {
  const s = String(raw || "").trim();
  if (/new\s*design/i.test(s)) return "New Design";
  if (/custom|variation/i.test(s)) return "Custom";
  return "Standard";
}

function strictestOrderType(types) {
  const mapped = (types || []).map(orderTypeFromEnquiryType);
  if (mapped.indexOf("New Design") !== -1) return "New Design";
  if (mapped.indexOf("Custom") !== -1) return "Custom";
  return "Standard";
}

function isGauteng(province) {
  return /gauteng/i.test(String(province || ""));
}

function scheduleCodeForProvince(province) {
  return isGauteng(province) ? "LD" : "LC";
}

function leadWeeksFor(type, province) {
  const used = ORDER_TYPES.indexOf(type) !== -1 ? type : orderTypeFromEnquiryType(type);
  const row = LEAD_WEEKS[used] || LEAD_WEEKS.Standard;
  return isGauteng(province) ? row.gauteng : row.other;
}

function parseMoney(s) {
  const n = Number(String(s || "").replace(/,/g, "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0;
}

function money(n) {
  return (Math.round(Number(n) * 100) / 100).toFixed(2);
}

function isFlag(value) {
  return value === true || value === "true" || value === 1 || value === "1" || value === "on";
}

function lineInclVat(p) {
  const incl = String((p && p.price_incl_vat) == null ? "" : p.price_incl_vat).trim();
  if (incl) return money(parseMoney(incl));
  const excl = String((p && p.price_excl_vat) == null ? "" : p.price_excl_vat).trim();
  if (!excl) return "";
  return money(parseMoney(excl) * 1.15);
}

function lineAmountPaid(p, incl) {
  if (isFlag(p && p.paid_in_full)) return incl;
  return p && p.amount_paid;
}

function splitCents(total, qty, index) {
  const cents = Math.round(parseMoney(total) * 100);
  const n = Math.max(1, Number(qty) || 1);
  const base = Math.floor(cents / n);
  const rem = cents - base * n;
  const share = base + (index === n - 1 ? rem : 0);
  return (share / 100).toFixed(2);
}

function assertTueThu(iso) {
  const day = String(iso || "").trim().slice(0, 10);
  if (!dateAtNoon(day)) throw new Error("Choose a delivery date");
  if (!isTueOrThu(day)) throw new Error("Deliveries are Tuesday and Thursday only.");
  return day;
}

function assertDeliveryDay(iso, province, now) {
  const day = String(iso || "").trim().slice(0, 10);
  const today = todayIso(now);
  if (!dateAtNoon(day)) throw new Error("Choose a delivery date");
  if (isNonGautengNextWeek(day, province, today)) {
    if (!isMonday(day)) {
      throw new Error("If Latest Courier falls in next week, put LC on Monday — not Tuesday or Thursday.");
    }
    return day;
  }
  return assertTueThu(day);
}

function typeLeadLabel(typeUsed) {
  if (typeUsed === "New Design") return "New Design";
  if (typeUsed === "Custom") return "Custom";
  return "Standard";
}

function deliveryExplanation(opts) {
  const today = todayIso(opts && opts.now);
  const date = String((opts && opts.date) || "").slice(0, 10);
  const weeks = Number((opts && opts.weeks) || 0);
  const typeUsed = typeLeadLabel(opts && opts.typeUsed);
  const province = String((opts && opts.province) || "").trim();
  const gauteng = isGauteng(province);
  const provinceBit = gauteng
    ? "the province is Gauteng, so this is Latest Delivery (LD) on the production schedule"
    : (province
      ? "the province is " + province + " (not Gauteng), so this is Latest Courier (LC) on the production schedule"
      : "the province is not Gauteng, so this is Latest Courier (LC) on the production schedule");
  const pretty = prettyDate(date);
  const todayPretty = prettyDate(today);
  const info = dateAtNoon(date) ? isoWeekInfo(date) : { week: "" };
  const weekBit = pretty + " in week " + info.week;
  const nextWeekLc = isNonGautengNextWeek(date, province, today);
  const weekWord = weeks === 1 ? "week" : "weeks";
  if (opts && opts.dateTouched) {
    let why = "You chose " + weekBit + " because " + provinceBit + ".";
    if (nextWeekLc) {
      why += " That week is next week, so LC is Monday — not Tuesday or Thursday.";
    } else {
      why += " Deliveries leave on Tuesday and Thursday only. Change this date if you need a different week — the date you save is the one that goes on the schedule.";
    }
    return why;
  }
  let why = "Based on today (" + todayPretty + "), lead time is " + weeks + " " + weekWord +
    " because the strictest item is " + typeUsed + ", and " + provinceBit + ".";
  if (nextWeekLc) {
    why += " The calculated date falls in next week, so LC is Monday " + weekBit + " — not Tuesday or Thursday.";
  } else {
    why += " That lands on " + weekBit + ". Deliveries leave on Tuesday and Thursday only.";
  }
  why += " Change this date if you need a different week — the date you save is the one that goes on the schedule.";
  return why;
}

function estimateDelivery(opts) {
  const types = (opts && opts.types) || [];
  const province = (opts && opts.province) || "";
  const used = strictestOrderType(types);
  const weeks = leadWeeksFor(used, province);
  const today = todayIso(opts && opts.now);
  const chosen = opts && opts.date
    ? assertDeliveryDay(opts.date, province, today)
    : snapToScheduleDay(addDaysIso(today, weeks * 7), province, today);
  const info = isoWeekInfo(chosen);
  const pretty = prettyDate(chosen);
  const explanation = deliveryExplanation({
    now: today,
    date: chosen,
    weeks,
    typeUsed: used,
    province,
    dateTouched: !!(opts && opts.dateTouched)
  });
  return {
    date: chosen,
    week: info.week,
    year: info.year,
    weekday: pretty.split(" ")[0] || "",
    weeks,
    typeUsed: used,
    scheduleCode: scheduleCodeForProvince(province),
    pretty,
    nextWeekLc: isNonGautengNextWeek(chosen, province, today),
    explanation,
    label: explanation
  };
}

function namedEnquiryLines(enquiry) {
  const chosen = quoteOptions.chosenQuote(enquiry);
  if (chosen && Array.isArray(chosen.products) && chosen.products.some((p) => String((p && p.product) || "").trim())) {
    return chosen.products.filter((p) => String((p && p.product) || "").trim());
  }
  const products = Array.isArray(enquiry && enquiry.products) ? enquiry.products : [];
  const named = products.filter((p) => String((p && p.product) || "").trim());
  if (named.length) return named;
  const product = String((enquiry && enquiry.product) || "").trim();
  if (!product) return [];
  return [{
    product,
    category: String((enquiry && enquiry.category) || "").trim(),
    variation: "",
    value_incl_vat: (enquiry && (enquiry.quote_total_incl_vat || enquiry.value_incl_vat)) || ""
  }];
}

function buildDraft(enquiry, nextOrderNumber, now) {
  const typeDefault = orderTypeFromEnquiryType(enquiry && enquiry.enquiry_type);
  const chosen = quoteOptions.chosenQuote(enquiry);
  const products = namedEnquiryLines(enquiry).map((p) => ({
    category: String(p.category || (enquiry && enquiry.category) || "").trim(),
    product: String(p.product || "").trim(),
    type: typeDefault,
    variation: String(p.variation || "").trim(),
    doors: "",
    detailed_description: "",
    dimensions: "",
    powder_coating: "",
    price_incl_vat: p.value_incl_vat || "",
    amount_paid: "",
    quantity: 1
  }));
  return {
    enquiry_no: enquiry && enquiry.enquiry_no || "",
    quote_number: (chosen && chosen.quote_no) || (enquiry && enquiry.quote_no) || "",
    chosen_option: (enquiry && enquiry.chosen_option) || (chosen && chosen.option) || "",
    order_number: nextOrderNumber,
    shared: {
      client_name: (enquiry && enquiry.client_name) || "",
      client_number: (enquiry && enquiry.client_number) || "",
      email: (enquiry && enquiry.client_email) || "",
      address: (enquiry && enquiry.address) || "",
      province: (enquiry && enquiry.province) || "",
      city: "",
      source: (enquiry && enquiry.source) || ""
    },
    products,
    delivery: estimateDelivery({
      types: products.map((p) => p.type),
      province: (enquiry && enquiry.province) || "",
      now
    })
  };
}

function planUnits(products) {
  const units = [];
  (products || []).forEach((p) => {
    const qty = Math.floor(Number(p && p.quantity) || 0);
    if (qty < 1) return;
    const product = String((p && p.product) || "").trim();
    if (!product) return;
    let incl = lineInclVat(p);
    const paid = lineAmountPaid(p, incl);
    for (let i = 0; i < qty; i++) {
      units.push({
        product,
        category: String((p && p.category) || "").trim(),
        type: orderTypeFromEnquiryType(p && p.type),
        variation: String((p && p.variation) || "").trim(),
        doors: String((p && p.doors) || "").trim(),
        detailed_description: String((p && p.detailed_description) || "").trim(),
        dimensions: String((p && p.dimensions) || "").trim(),
        powder_coating: String((p && p.powder_coating) || "").trim(),
        price_incl_vat: splitCents(incl, qty, i),
        amount_paid: splitCents(paid, qty, i)
      });
    }
  });
  return units;
}

function requireText(value, message) {
  const s = String(value == null ? "" : value).trim();
  if (!s) throw new Error(message);
  return s;
}

function isQtyPriceConfirmed(value) {
  return value === true || value === "true" || value === 1 || value === "1" || value === "on";
}

function assertProductReady(product, index) {
  const qty = Math.floor(Number(product && product.quantity) || 0);
  if (qty < 1) return;
  const label = String((product && product.product) || "").trim() || ("product " + (index + 1));
  const incl = lineInclVat(product);
  const excl = String((product && product.price_excl_vat) || "").trim();
  if (!incl && !excl) throw new Error("Price excl VAT is required for " + label);
  if (!isQtyPriceConfirmed(product && product.qty_price_confirmed)) {
    throw new Error("Confirm quantity and price excl VAT for " + label);
  }
  if (!isFlag(product && product.paid_in_full) && String(product && product.amount_paid == null ? "" : product.amount_paid).trim() === "") {
    throw new Error("Amount paid is required for " + label + " unless you tick Paid in full");
  }
  requireText(product && product.variation, "Variation is required for " + label);
  requireText(product && product.doors, "Doors is required for " + label);
  requireText(product && product.powder_coating, "Powder coating is required for " + label);
  requireText(product && product.dimensions, "Dimensions are required for " + label);
  requireText(product && product.detailed_description, "Detailed description is required for " + label);
}

function planCreate(enquiry, body, existingOrders) {
  const sharedIn = (body && body.shared) || {};
  const client = String(sharedIn.client_name || "").trim();
  if (!client) throw new Error("Client name is required");
  requireText(sharedIn.address, "Address is required");
  requireText(sharedIn.city, "City is required");
  const base = normalizeBaseOrderNumber(body && body.order_number);
  if (!base) throw new Error("Order number must look like S260100");
  if (orderBaseTaken(base, existingOrders)) {
    throw new Error("Order number " + base + " is already used");
  }
  const products = Array.isArray(body && body.products) ? body.products : [];
  products.forEach(assertProductReady);
  const units = planUnits(products);
  if (!units.length) throw new Error("Add a quantity of at least 1 for a product");
  const province = String(sharedIn.province || "").trim();
  const delivery = estimateDelivery({
    types: products.map((p) => p.type),
    province,
    now: body && body.now,
    date: body && body.delivery_date,
    dateTouched: !!(body && body.dateTouched)
  });
  const total = units.length;
  return {
    base,
    quote_number: (enquiry && enquiry.quote_no) || "",
    enquiry_no: (enquiry && enquiry.enquiry_no) || "",
    shared: {
      client_name: client,
      client_number: String(sharedIn.client_number || "").trim(),
      email: String(sharedIn.email || "").trim(),
      address: String(sharedIn.address || "").trim(),
      province,
      city: String(sharedIn.city || "").trim(),
      source: String(sharedIn.source || "").trim()
    },
    units: units.map((u, i) => Object.assign({}, u, {
      order_number: total > 1 ? base + " " + splitLetter(i) : base
    })),
    delivery
  };
}

module.exports = {
  ORDER_TYPES,
  LEAD_WEEKS,
  todayIso,
  addDaysIso,
  isTueOrThu,
  isMonday,
  mondayOfIsoWeek,
  isNextIsoWeek,
  snapToScheduleDay,
  nextTueOrThu,
  prettyDate,
  studioOrderSeq,
  normalizeBaseOrderNumber,
  nextStudioOrderNumberFrom,
  orderBaseTaken,
  splitLetter,
  orderTypeFromEnquiryType,
  strictestOrderType,
  isGauteng,
  scheduleCodeForProvince,
  leadWeeksFor,
  splitCents,
  assertTueThu,
  assertDeliveryDay,
  deliveryExplanation,
  estimateDelivery,
  buildDraft,
  planUnits,
  planCreate
};
