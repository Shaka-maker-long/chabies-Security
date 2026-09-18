"use strict";

const SAST_OFFSET_MS = 2 * 60 * 60 * 1000;
const DAY_MS = 86400000;
const MONTH_NAMES = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];
const MONTH_IX = {
  jan: 0, january: 0, feb: 1, february: 1, mar: 2, march: 2,
  apr: 3, april: 3, may: 4, jun: 5, june: 5, jul: 6, july: 6,
  aug: 7, august: 7, sep: 8, sept: 8, september: 8, oct: 9, october: 9,
  nov: 10, november: 10, dec: 11, december: 11
};

function isSheetError(s) {
  return /^#(NUM|VALUE|REF|N\/A|DIV\/0|NAME|NULL)!/i.test(String(s || "").trim());
}

function expandYear(y) {
  const n = Number(y);
  if (!Number.isFinite(n)) return 0;
  return n < 100 ? 2000 + n : n;
}

function daysInMonth(y, m) {
  return new Date(Date.UTC(Number(y), Number(m), 0)).getUTCDate();
}

function sastDate(y, m, day) {
  return new Date(Date.UTC(Number(y), Number(m) - 1, Number(day)) - SAST_OFFSET_MS);
}

function makeSast(y, month1to12, day) {
  const yy = Number(y);
  const mm = Number(month1to12);
  const dd = Number(day);
  if (!Number.isFinite(yy) || !Number.isFinite(mm) || !Number.isFinite(dd)) return null;
  if (mm < 1 || mm > 12 || dd < 1 || dd > daysInMonth(yy, mm)) return null;
  return sastDate(yy, mm, dd);
}

function sastParts(d) {
  const sast = new Date(d.getTime() + SAST_OFFSET_MS);
  return { y: sast.getUTCFullYear(), m: sast.getUTCMonth(), day: sast.getUTCDate() };
}

function isoFromDate(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return "";
  const p = sastParts(d);
  return p.y + "-" + String(p.m + 1).padStart(2, "0") + "-" + String(p.day).padStart(2, "0");
}

function padDmy(d) {
  const p = sastParts(d);
  return String(p.day).padStart(2, "0") + "/" + String(p.m + 1).padStart(2, "0") + "/" + p.y;
}

function monthHintIndex(hint) {
  if (hint == null || hint === "") return null;
  if (hint instanceof Date && !isNaN(hint.getTime())) return sastParts(hint).m;
  const s = String(hint).trim();
  if (!s) return null;
  const named = s.match(/^([A-Za-z]{3,})(?:\s+(\d{2,4}))?$/);
  if (named && MONTH_IX[named[1].toLowerCase()] != null) return MONTH_IX[named[1].toLowerCase()];
  const short = s.match(/^([A-Za-z]{3,})\s+\d{4}$/);
  if (short && MONTH_IX[short[1].toLowerCase()] != null) return MONTH_IX[short[1].toLowerCase()];
  return null;
}

function startOfSastDay(d) {
  const p = sastParts(d instanceof Date && !isNaN(d.getTime()) ? d : new Date());
  return makeSast(p.y, p.m + 1, p.day);
}

function preferMdyIfDmyFuture(dmy, mdy, now) {
  if (!dmy || !mdy) return false;
  const today = startOfSastDay(now);
  const dmyDay = startOfSastDay(dmy);
  const mdyDay = startOfSastDay(mdy);
  const dmyFuture = dmyDay.getTime() > today.getTime() + 7 * DAY_MS;
  const mdyRecent = mdyDay.getTime() <= today.getTime() + 7 * DAY_MS
    && mdyDay.getTime() >= today.getTime() - 60 * DAY_MS;
  return dmyFuture && mdyRecent;
}

function slashPair(first, second, year, hint, now) {
  const y = expandYear(year);
  const dmy = makeSast(y, second, first);
  const mdy = makeSast(y, first, second);
  if (dmy && !mdy) return { date: dmy, dmy, mdy: null, ambiguous: false, kind: "dmy" };
  if (mdy && !dmy) return { date: mdy, dmy: null, mdy, ambiguous: false, kind: "mdy" };
  if (!dmy && !mdy) return { date: null, dmy: null, mdy: null, ambiguous: false, kind: "none" };
  if (Number(first) === Number(second)) {
    return { date: dmy, dmy, mdy, ambiguous: false, kind: "same" };
  }
  const hm = monthHintIndex(hint);
  if (hm != null) {
    const dmyM = Number(second) - 1;
    const mdyM = Number(first) - 1;
    if (mdyM === hm && dmyM !== hm) {
      return { date: mdy, dmy, mdy, ambiguous: false, kind: "hint-mdy" };
    }
  }
  if (preferMdyIfDmyFuture(dmy, mdy, now)) {
    return { date: mdy, dmy, mdy, ambiguous: false, kind: "future-mdy" };
  }
  return { date: dmy, dmy, mdy, ambiguous: true, kind: "dmy-default" };
}

function slashToDate(first, second, year, hint, now) {
  return slashPair(first, second, year, hint, now).date;
}

function isAmbiguousSlash(s) {
  const m = String(s || "").trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (!m) return false;
  return slashPair(Number(m[1]), Number(m[2]), m[3]).ambiguous;
}

function orderSeq(orderNumber) {
  const m = String(orderNumber || "").trim().match(/^S(\d{6})(?:\s+[A-Za-z].*)?$/i);
  return m ? Number(m[1]) : null;
}

function asDate(v, hint, now) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === "number" && isFinite(v) && v >= 20000 && v <= 120000) {
    const utcMs = Math.round((v - 25569) * 86400000);
    return new Date(utcMs - SAST_OFFSET_MS);
  }
  const s = String(v || "").trim();
  if (!s || isSheetError(s)) return null;
  if (/^\d{4}-\d{2}-\d{2}T/.test(s) || /^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d;
  }
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return sastDate(iso[1], iso[2], iso[3]);
  const slash = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (slash) return slashToDate(Number(slash[1]), Number(slash[2]), slash[3], hint, now);
  const named = s.match(/^(\d{1,2})[\/\-\s]+([A-Za-z]{3,})(?:[\/\-\s,]+(\d{2,4}))?$/);
  if (named) {
    const m = MONTH_IX[named[2].toLowerCase()];
    if (m != null) {
      let y = named[3] ? expandYear(named[3]) : sastParts(new Date()).y;
      if (!named[3] && hint) {
        const hintYear = String(hint).match(/(\d{4})/);
        const hm = monthHintIndex(hint);
        if (hintYear && (hm == null || hm === m)) y = Number(hintYear[1]);
      }
      return makeSast(y, m + 1, Number(named[1]));
    }
  }
  return null;
}

function formatPaymentDate(v, hint, now) {
  if (v == null || v === "") return "";
  if (isSheetError(v)) return "";
  const d = asDate(v, hint, now);
  if (!d) return "";
  return padDmy(d);
}

function formatMonthOfSale(v, hint, now) {
  if (v == null || v === "") return "";
  if (isSheetError(v)) return "";
  const d = asDate(v, hint, now);
  if (d) {
    const p = sastParts(d);
    return MONTH_NAMES[p.m] + " " + p.y;
  }
  const s = String(v).trim().replace(/\s+/g, " ");
  const named = s.match(/^([A-Za-z]+)\s+(\d{4})$/);
  if (!named) return "";
  const m = MONTH_IX[named[1].toLowerCase()];
  if (m == null) return "";
  return MONTH_NAMES[m] + " " + named[2];
}

function monthOfSaleFromPayment(payment, hint) {
  const pay = formatPaymentDate(payment, hint);
  return pay ? formatMonthOfSale(pay) : "";
}

function slashAnchorTime(v) {
  if (v == null || v === "") return null;
  if (isSheetError(v)) return null;
  if (v instanceof Date && !isNaN(v.getTime())) {
    if (isAmbiguousSlash(padDmy(v))) return null;
    return v.getTime();
  }
  const s = String(v).trim();
  if (isAmbiguousSlash(s)) return null;
  const d = asDate(s);
  return d ? d.getTime() : null;
}

function pickBySequence(dmy, mdy, seq, anchors) {
  if (seq == null || !anchors || !anchors.length) return dmy;
  if (!dmy || !mdy) return dmy || mdy;
  const prev = anchors.filter((a) => a.seq < seq).sort((a, b) => b.seq - a.seq)[0];
  const next = anchors.filter((a) => a.seq > seq).sort((a, b) => a.seq - b.seq)[0];
  const CLUSTER = 31 * DAY_MS;
  const FAR = 90 * DAY_MS;
  const dmyT = dmy.getTime();
  const mdyT = mdy.getTime();
  if (prev) {
    const farBeforePrev = dmyT < prev.t - FAR;
    const mdyNearAfterPrev = mdyT >= prev.t - 2 * DAY_MS && mdyT <= prev.t + CLUSTER;
    if (farBeforePrev && mdyNearAfterPrev) return mdy;
  }
  if (next) {
    const farAfterNext = dmyT > next.t + FAR;
    const mdyNearBeforeNext = mdyT <= next.t + 2 * DAY_MS && mdyT >= next.t - CLUSTER;
    if (farAfterNext && mdyNearBeforeNext) return mdy;
  }
  return dmy;
}

function resolvePaymentDate(value, orderNumber, hint, anchors, now) {
  if (value == null || value === "") return "";
  if (isSheetError(value)) return "";
  const raw = value instanceof Date ? padDmy(value) : String(value).trim();
  const slash = raw.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (slash) {
    const pair = slashPair(Number(slash[1]), Number(slash[2]), slash[3], hint, now);
    if (!pair.date) return "";
    if (pair.ambiguous) {
      const picked = pickBySequence(pair.dmy, pair.mdy, orderSeq(orderNumber), anchors || []);
      return padDmy(picked);
    }
    return padDmy(pair.date);
  }
  return formatPaymentDate(value, hint, now);
}

function paymentAnchors(rows, exceptOrder) {
  const skip = String(exceptOrder || "").trim().toLowerCase();
  const anchors = [];
  (rows || []).forEach((row) => {
    const orderNumber = row.orderNumber != null ? row.orderNumber : row.order_number;
    if (skip && String(orderNumber || "").trim().toLowerCase() === skip) return;
    const seq = orderSeq(orderNumber);
    const payment = row.payment != null ? row.payment : row.payment_date;
    const t = slashAnchorTime(payment);
    if (seq != null && t != null) anchors.push({ seq, t });
  });
  return anchors;
}

module.exports = {
  SAST_OFFSET_MS,
  MONTH_NAMES,
  MONTH_IX,
  isSheetError,
  sastDate,
  sastParts,
  isoFromDate,
  asDate,
  slashToDate,
  isAmbiguousSlash,
  orderSeq,
  formatPaymentDate,
  formatMonthOfSale,
  monthOfSaleFromPayment,
  slashAnchorTime,
  resolvePaymentDate,
  paymentAnchors
};
