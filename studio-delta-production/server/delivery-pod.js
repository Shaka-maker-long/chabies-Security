"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir, getBook, persistWorkbook } = require("./workbook-store");
const db = require("./db");
const catalog = require("./product-catalog");
const staff = require("./staff");
const { normalizeBaseOrderNumber } = require("./create-order-from-enquiry");
const { toJpeg } = require("./qc-pdf");

const PAGE_W = 595.28;
const PAGE_H = 841.89;
const MARGIN = 36;
const INK = "#1c1917";
const BRASS = "#b08948";
const MUTED = "#6b645b";
const RULE = "#d7d1c6";
const LAYOUT = 1;
const AVG_KMH = 25;
const STOP_MINUTES = 8;
const FACTORY = {
  lat: Number(process.env.STUDIO_DELTA_LAT) || -25.7254893,
  lng: Number(process.env.STUDIO_DELTA_LNG) || 28.2948254,
  label: "Studio Delta",
  address: "309 Derdepoort Rd, Silverton, Pretoria, 0184"
};
const THIRD_PARTY_DEPOT = {
  address: "26 Milkyway Ave",
  city: "Frankenwald",
  province: "Gauteng",
  full_address: "26 Milkyway Ave, Frankenwald, Sandton",
  label: "3rd party depot",
  lat: -26.0674,
  lng: 28.1112
};

function isGautengProvince(province) {
  const p = String(province || "").trim().toLowerCase();
  if (!p) return false;
  return p === "gauteng" || p === "gp" || p === "gauteng province";
}

function dropForOrder(order) {
  const clientAddress = fullAddress(order);
  if (isGautengProvince(order && order.province)) {
    return {
      drop_kind: "client",
      drop_label: "Client",
      drop_address: clientAddress,
      client_address: clientAddress
    };
  }
  return {
    drop_kind: "third_party",
    drop_label: THIRD_PARTY_DEPOT.label,
    drop_address: THIRD_PARTY_DEPOT.full_address,
    client_address: clientAddress
  };
}

function storePath() {
  return path.join(dataDir(), "delivery-forms.json");
}

function geoPath() {
  return path.join(dataDir(), "delivery-geocode.json");
}

function formsDir() {
  const dir = path.join(dataDir(), "delivery-forms");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function photosDir(id) {
  const dir = path.join(formsDir(), String(id || ""));
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function newId() {
  return crypto.randomBytes(12).toString("hex");
}

function pdfUrlFor(id) {
  return "/api/delivery-forms/" + encodeURIComponent(id) + "/pdf";
}

function loadStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(storePath(), "utf8"));
    if (parsed && Array.isArray(parsed.records)) return parsed;
  } catch (e) {
    if (e && e.code !== "ENOENT") console.error("[delivery-pod] read", e.message || e);
  }
  return { records: [] };
}

function saveStore(store) {
  const file = storePath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(store));
  fs.renameSync(tmp, file);
  return store;
}

function loadGeo() {
  try {
    const parsed = JSON.parse(fs.readFileSync(geoPath(), "utf8"));
    if (parsed && parsed.places && typeof parsed.places === "object") return parsed;
  } catch (e) {}
  return { places: {} };
}

function saveGeo(store) {
  const file = geoPath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(store));
}

function formatWhen(iso) {
  try {
    return new Intl.DateTimeFormat("en-GB", {
      timeZone: process.env.TZ || "Africa/Johannesburg",
      day: "2-digit",
      month: "short",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false
    }).format(iso ? new Date(iso) : new Date());
  } catch (e) {
    return String(iso || "");
  }
}

function isDriverProfile(profile) {
  if (!profile) return false;
  const title = String(profile.jobTitle || profile.role || "").trim().toLowerCase();
  if (title === "driver") return true;
  const tasks = profile.tasks || [];
  return tasks.some((t) => String(t).trim().toLowerCase() === "delivery");
}

function canLoadTruck(profile) {
  if (!profile) return false;
  if (profile.isAdmin || profile.canSeeOffice) return true;
  const tasks = profile.tasks || [];
  if (tasks.indexOf("Quality Control") !== -1) return true;
  if (isDriverProfile(profile)) return true;
  return String(profile.name || "").trim().toLowerCase() === "siya";
}

function canSubmitPod(profile) {
  if (!profile) return false;
  if (isDriverProfile(profile)) return true;
  if (profile.isAdmin || profile.canSeeOffice) return true;
  const tasks = profile.tasks || [];
  return tasks.indexOf("Quality Control") !== -1
    || String(profile.name || "").trim().toLowerCase() === "siya";
}

function orderBase(orderNumber) {
  return normalizeBaseOrderNumber(orderNumber) || String(orderNumber || "").trim().toUpperCase();
}

function fullAddress(order) {
  return [order.address, order.city, order.province].map((s) => String(s || "").trim()).filter(Boolean).join(", ");
}

function decorateOrder(order) {
  if (!order) return null;
  const product = catalog.lookupProduct(order.product) || {};
  const drop = dropForOrder(order);
  return {
    order_number: db.formatOrderId(order.order_number),
    base: orderBase(order.order_number),
    status: String(order.status || "").trim(),
    product: String(order.product || ""),
    client_name: String(order.client_name || ""),
    client_number: String(order.client_number || ""),
    address: String(order.address || ""),
    city: String(order.city || ""),
    province: String(order.province || ""),
    full_address: fullAddress(order),
    client_address: drop.client_address,
    drop_kind: drop.drop_kind,
    drop_label: drop.drop_label,
    drop_address: drop.drop_address,
    imageUrl: product.imageUrl || "",
    assigned_operator: String(order.assigned_operator || "")
  };
}

function relatedUnits(orderNumber, statusWanted) {
  const base = orderBase(orderNumber);
  const want = String(statusWanted || "").trim().toLowerCase();
  return db.listOrders()
    .filter((row) => orderBase(row.order_number) === base)
    .filter((row) => !want || String(row.status || "").trim().toLowerCase() === want)
    .map(decorateOrder)
    .filter(Boolean)
    .sort((a, b) => String(a.order_number).localeCompare(String(b.order_number)));
}

function listReadyToLoad() {
  return db.listOrders()
    .filter((row) => String(row.status || "").trim().toLowerCase() === "ready for delivery")
    .map(decorateOrder)
    .filter(Boolean);
}

function listLoaded() {
  return db.listOrders()
    .filter((row) => String(row.status || "").trim().toLowerCase() === "out for delivery")
    .map(decorateOrder)
    .filter(Boolean);
}

function setOrderStatus(order, status, assigned) {
  const payload = Object.assign({}, order, {
    status,
    assigned_operator: assigned == null ? (order.assigned_operator || "") : assigned
  });
  return db.upsertOrder(payload);
}

function loadOnTruck(orderNumbers, actorName) {
  const names = Array.isArray(orderNumbers) ? orderNumbers : [orderNumbers];
  const picked = names.map((n) => db.formatOrderId(n)).filter(Boolean);
  if (!picked.length) throw new Error("Pick at least one order to load.");
  const actor = String(actorName || "").trim();
  const loaded = [];
  const seen = {};
  picked.forEach((num) => {
    relatedUnits(num, "ready for delivery").forEach((unit) => {
      if (seen[unit.order_number]) return;
      seen[unit.order_number] = true;
      const live = db.listOrders().find((row) => db.formatOrderId(row.order_number) === unit.order_number);
      if (!live) return;
      setOrderStatus(live, "Out for Delivery", actor);
      loaded.push(decorateOrder(Object.assign({}, live, { status: "Out for Delivery", assigned_operator: actor })));
    });
  });
  if (!loaded.length) throw new Error("Those orders are not Ready for Delivery.");
  try { db.persistOffice(); } catch (e) {}
  return { count: loaded.length, orders: loaded };
}

function haversineKm(a, b) {
  if (!a || !b || a.lat == null || b.lat == null) return 0;
  const R = 6371;
  const dLat = (Number(b.lat) - Number(a.lat)) * Math.PI / 180;
  const dLon = (Number(b.lng) - Number(a.lng)) * Math.PI / 180;
  const s = Math.sin(dLat / 2) ** 2
    + Math.cos(Number(a.lat) * Math.PI / 180) * Math.cos(Number(b.lat) * Math.PI / 180) * Math.sin(dLon / 2) ** 2;
  return 2 * R * Math.asin(Math.min(1, Math.sqrt(s)));
}

function defaultGeocode(query) {
  const q = String(query || "").toLowerCase();
  if (!q) return null;
  if (/milkyway|milky way|frankenwald/.test(q)) {
    return { lat: THIRD_PARTY_DEPOT.lat, lng: THIRD_PARTY_DEPOT.lng };
  }
  if (/derdepoort|silverton|silvertondale/.test(q)) {
    return { lat: FACTORY.lat, lng: FACTORY.lng };
  }
  if (/sandton|johannesburg|loop street/.test(q)) return { lat: -26.1076, lng: 28.0567 };
  if (/cape town|stellenbosch|western cape|beach road/.test(q)) return { lat: -33.9249, lng: 18.4241 };
  if (/durban|kwazulu/.test(q)) return { lat: -29.8587, lng: 31.0218 };
  if (/polokwane|limpopo/.test(q)) return { lat: -23.9045, lng: 29.4689 };
  if (/hermanus/.test(q)) return { lat: -34.4187, lng: 19.2345 };
  return { lat: -26.2041, lng: 28.0473 };
}

async function nominatimGeocode(query) {
  const url = "https://nominatim.openstreetmap.org/search?format=json&limit=1&q=" + encodeURIComponent(query);
  const res = await fetch(url, {
    headers: { "User-Agent": "StudioDeltaDelivery/1.0 (shop floor)" },
    signal: AbortSignal.timeout(2500)
  });
  if (!res.ok) return null;
  const rows = await res.json();
  if (!rows || !rows[0]) return null;
  return { lat: Number(rows[0].lat), lng: Number(rows[0].lon) };
}

async function geocodeAddress(query, geocodeFn) {
  const key = String(query || "").trim().toLowerCase();
  if (!key) return null;
  const geo = loadGeo();
  if (geo.places[key]) return geo.places[key];
  let hit = null;
  try {
    hit = geocodeFn ? await geocodeFn(query) : await nominatimGeocode(query);
  } catch (e) {
    hit = null;
  }
  if (!hit || !Number.isFinite(Number(hit.lat))) hit = defaultGeocode(query);
  geo.places[key] = hit;
  saveGeo(geo);
  return hit;
}

function travelMinutes(km) {
  return Math.max(4, Math.round((km / AVG_KMH) * 60));
}

function osrmBaseUrl() {
  return String(process.env.OSRM_URL || "https://router.project-osrm.org").replace(/\/+$/, "");
}

function wazeUrl(lat, lng) {
  if (!Number.isFinite(Number(lat)) || !Number.isFinite(Number(lng))) return "";
  return "https://waze.com/ul?ll=" + Number(lat).toFixed(6) + "," + Number(lng).toFixed(6) + "&navigate=yes";
}

function parseOsrmRoute(json, pointCount) {
  if (!json || json.code !== "Ok" || !json.routes || !json.routes[0]) return null;
  const route = json.routes[0];
  const coords = (((route.geometry || {}).coordinates) || []).map((pair) => [Number(pair[1]), Number(pair[0])]);
  const legs = Array.isArray(route.legs) ? route.legs : [];
  if (pointCount && legs.length !== pointCount - 1) return null;
  return {
    path: coords.filter((p) => Number.isFinite(p[0]) && Number.isFinite(p[1])),
    legs: legs.map((leg) => ({
      km: Math.round((Number(leg.distance) || 0) / 100) / 10,
      minutes: Math.max(1, Math.round((Number(leg.duration) || 0) / 60))
    }))
  };
}

async function fetchOsrmDrive(points) {
  const pins = (points || []).filter((p) => p && Number.isFinite(Number(p.lat)) && Number.isFinite(Number(p.lng)));
  if (pins.length < 2) return null;
  const loc = pins.map((p) => Number(p.lng).toFixed(6) + "," + Number(p.lat).toFixed(6)).join(";");
  const url = osrmBaseUrl() + "/route/v1/driving/" + loc + "?overview=full&geometries=geojson&steps=false";
  const res = await fetch(url, {
    headers: { "User-Agent": "StudioDeltaDelivery/1.0 (shop floor)" },
    signal: AbortSignal.timeout(4000)
  });
  if (!res.ok) return null;
  return parseOsrmRoute(await res.json(), pins.length);
}

function routeStopKey(row) {
  if (row && row.drop_kind === "third_party") return "third_party";
  return String((row && (row.drop_address || row.full_address || row.base)) || "").trim().toLowerCase();
}

async function buildRoute(opts) {
  const loaded = listLoaded();
  const groups = {};
  loaded.forEach((row) => {
    const key = routeStopKey(row) || (row.base || row.order_number);
    if (!groups[key]) {
      groups[key] = {
        base: key,
        drop_kind: row.drop_kind,
        drop_label: row.drop_label,
        address: row.drop_address || row.full_address,
        client_name: row.client_name,
        orders: []
      };
    }
    groups[key].orders.push(row);
  });
  const stops = Object.keys(groups).map((k) => groups[k]);
  const origin = (opts && opts.origin && Number.isFinite(Number(opts.origin.lat)))
    ? { lat: Number(opts.origin.lat), lng: Number(opts.origin.lng), label: "You" }
    : FACTORY;
  const geocodeFn = opts && opts.geocodeFn;
  for (let i = 0; i < stops.length; i++) {
    const pin = await geocodeAddress(stops[i].address, geocodeFn);
    stops[i].lat = pin && pin.lat;
    stops[i].lng = pin && pin.lng;
  }
  const remaining = stops.slice();
  const ordered = [];
  let cursor = origin;
  let elapsed = 0;
  const now = opts && opts.now ? new Date(opts.now) : new Date();
  while (remaining.length) {
    let best = 0;
    let bestKm = Infinity;
    remaining.forEach((stop, i) => {
      const km = haversineKm(cursor, stop);
      if (km < bestKm) {
        bestKm = km;
        best = i;
      }
    });
    const next = remaining.splice(best, 1)[0];
    elapsed += travelMinutes(bestKm);
    const eta = new Date(now.getTime() + elapsed * 60000);
    next.km = Math.round(bestKm * 10) / 10;
    next.eta = eta.toISOString();
    next.eta_label = formatWhen(eta);
    next.minutes = elapsed;
    ordered.push(next);
    cursor = next;
    elapsed += STOP_MINUTES;
  }
  ordered.forEach((stop) => {
    stop.waze_url = wazeUrl(stop.lat, stop.lng);
  });
  let path = [];
  let onRoads = false;
  const driveFn = opts && Object.prototype.hasOwnProperty.call(opts, "driveFn")
    ? opts.driveFn
    : fetchOsrmDrive;
  if (typeof driveFn === "function") {
    try {
      const drive = await driveFn([origin].concat(ordered));
      if (drive && Array.isArray(drive.path) && drive.path.length > 1) {
        path = drive.path;
        onRoads = true;
      }
      const legs = drive && Array.isArray(drive.legs) ? drive.legs : [];
      if (legs.length === ordered.length) {
        elapsed = 0;
        cursor = origin;
        ordered.forEach((stop, i) => {
          const leg = legs[i] || {};
          elapsed += Number(leg.minutes) || travelMinutes(haversineKm(cursor, stop));
          const eta = new Date(now.getTime() + elapsed * 60000);
          stop.km = Number.isFinite(Number(leg.km)) ? Number(leg.km) : stop.km;
          stop.eta = eta.toISOString();
          stop.eta_label = formatWhen(eta);
          stop.minutes = elapsed;
          cursor = stop;
          elapsed += STOP_MINUTES;
        });
      }
    } catch (e) {
      onRoads = false;
    }
  }
  if (!path.length) {
    path = [origin].concat(ordered).filter((p) => p && p.lat != null).map((p) => [Number(p.lat), Number(p.lng)]);
  }
  return {
    origin,
    stops: ordered,
    orders: loaded,
    path,
    on_roads: onRoads,
    average_kmh: AVG_KMH
  };
}

function closeOpenDeliveryLogs(orderNumbers, when, driver) {
  try {
    const sheet = getBook().getSheetByName("Production_Log");
    if (!sheet || sheet.getLastRow() < 2) return;
    const lastCol = Math.max(sheet.getLastColumn(), 7);
    const grid = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
    const want = {};
    (orderNumbers || []).forEach((n) => { want[db.formatOrderId(n)] = true; });
    const end = when instanceof Date ? when : new Date(when || Date.now());
    for (let i = 1; i < grid.length; i++) {
      const row = grid[i] || [];
      if (row[6]) continue;
      if (!want[db.formatOrderId(row[1])]) continue;
      const proc = String(row[4] || "").toLowerCase();
      const role = String(row[3] || "").toLowerCase();
      if (proc.indexOf("delivery") === -1 && role.indexOf("delivery") === -1 && role.indexOf("quality control") === -1) continue;
      sheet.getRange(i + 1, 7).setValue(end);
      if (driver) sheet.getRange(i + 1, 3).setValue(driver);
    }
    persistWorkbook();
  } catch (e) {
    console.error("[delivery-pod] close logs", e && e.message ? e.message : e);
  }
}

function starLine(n) {
  const v = Math.max(0, Math.min(5, Number(n) || 0));
  if (!v) return "—";
  return String(v) + " / 5";
}

function photoRaw(file) {
  if (!file) return "";
  if (Buffer.isBuffer(file)) return file;
  if (typeof file === "string") return file;
  if (file.data != null) return file.data;
  if (file.dataUrl) return file.dataUrl;
  if (file.base64) return file.base64;
  if (file.buf) return file.buf;
  return "";
}

async function encodePhotos(files) {
  const out = [];
  const list = Array.isArray(files) ? files : [];
  for (let i = 0; i < list.length; i++) {
    const jpeg = await toJpeg(photoRaw(list[i]));
    if (!jpeg) continue;
    out.push({ name: "Photo " + (out.length + 1), buf: jpeg });
  }
  return out;
}

async function logoPathForPdf() {
  const dest = path.join(dataDir(), "pdf-images", "studio-delta-logo.jpg");
  if (fs.existsSync(dest) && fs.statSync(dest).size > 400) return dest;
  return null;
}

function drawPdfBox(doc, x, y, w, h) {
  doc.save().lineWidth(0.7).strokeColor(INK).rect(x, y, w, h).stroke().restore();
}

function drawReportHeader(doc, record, logoPath) {
  const inner = PAGE_W - MARGIN * 2;
  drawPdfBox(doc, MARGIN, MARGIN, 72, 72);
  if (logoPath) {
    try {
      doc.image(logoPath, MARGIN + 4, MARGIN + 4, { fit: [64, 64] });
    } catch (e) {
      doc.fillColor(INK).font("Helvetica-Bold").fontSize(9).text("STUDIO\nDELTA", MARGIN, MARGIN + 26, { width: 72, align: "center" });
    }
  } else {
    doc.fillColor(INK).font("Helvetica-Bold").fontSize(9).text("STUDIO\nDELTA", MARGIN, MARGIN + 26, { width: 72, align: "center" });
  }
  doc.fillColor(INK).font("Helvetica-Bold").fontSize(16).text("STUDIO DELTA", MARGIN + 88, MARGIN + 8);
  doc.fillColor(MUTED).font("Helvetica").fontSize(9).text("Furniture  ·  Steel  ·  Glass", MARGIN + 88, MARGIN + 28);
  doc.fillColor(MUTED).font("Helvetica").fontSize(8).text("studiodelta.co.za", MARGIN + 88, MARGIN + 42);
  doc.fillColor(INK).font("Helvetica-Bold").fontSize(18).text("DELIVERY", MARGIN, MARGIN + 8, { width: inner, align: "right" });
  doc.fillColor(BRASS).font("Helvetica-Bold").fontSize(11).text(record.order_label || "", MARGIN, MARGIN + 34, { width: inner, align: "right" });
  doc.save().strokeColor(BRASS).lineWidth(2).moveTo(MARGIN, MARGIN + 84).lineTo(MARGIN + inner, MARGIN + 84).stroke().restore();
}

function kv(doc, label, value, x, y, w) {
  doc.fillColor(MUTED).font("Helvetica-Bold").fontSize(8).text(label, x, y, { width: w });
  doc.fillColor(INK).font("Helvetica").fontSize(11).text(value || "—", x, y + 12, { width: w });
}

async function renderPdf(record, dest) {
  const PDFDocument = require("pdfkit");
  const logoPath = await logoPathForPdf();
  const inner = PAGE_W - MARGIN * 2;
  const photos = (record.photos || []).filter((photo) => photo && photo.buf);
  if (!photos.length) throw new Error("Delivery form needs at least one photo.");
  await new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: "A4",
      margin: MARGIN,
      compress: false,
      info: {
        Title: (record.order_label || "Delivery") + " delivery form",
        Author: "Studio Delta"
      }
    });
    const stream = fs.createWriteStream(dest);
    doc.pipe(stream);
    stream.on("finish", resolve);
    stream.on("error", reject);
    drawReportHeader(doc, record, logoPath);
    let y = MARGIN + 96;
    kv(doc, "CLIENT", record.client_name, MARGIN, y, inner / 2 - 8);
    kv(doc, "RECEIVED BY", record.receiver_name, MARGIN + inner / 2, y, inner / 2);
    y += 40;
    if (record.drop_kind === "third_party") {
      kv(doc, "DROPPED AT  3rd party depot", record.address, MARGIN, y, inner);
      y += 40;
      kv(doc, "CLIENT ADDRESS", record.client_address || "", MARGIN, y, inner);
      y += 40;
    } else {
      kv(doc, "ADDRESS", record.address, MARGIN, y, inner);
      y += 40;
    }
    kv(doc, "DATE / TIME", formatWhen(record.delivered_at), MARGIN, y, inner / 2 - 8);
    kv(doc, "DRIVER", record.driver, MARGIN + inner / 2, y, inner / 2);
    y += 40;
    const geo = record.lat != null
      ? Number(record.lat).toFixed(5) + ", " + Number(record.lng).toFixed(5)
      : "—";
    kv(doc, "LOCATION", geo, MARGIN, y, inner);
    y += 44;
    doc.fillColor(MUTED).font("Helvetica-Bold").fontSize(8).text("RATINGS", MARGIN, y);
    y += 14;
    doc.fillColor(INK).font("Helvetica").fontSize(10)
      .text("Delivery team  " + starLine(record.rating_delivery), MARGIN, y)
      .text("Sales team  " + starLine(record.rating_sales), MARGIN, y + 16)
      .text("Craftsmanship  " + starLine(record.rating_craft), MARGIN, y + 32);
    y += 58;
    if (record.comments) {
      doc.fillColor(MUTED).font("Helvetica-Bold").fontSize(8).text("COMMENTS", MARGIN, y);
      doc.fillColor(INK).font("Helvetica").fontSize(10).text(record.comments, MARGIN, y + 14, { width: inner });
    }
    photos.forEach((photo, i) => {
      doc.addPage();
      doc.fillColor(INK).font("Helvetica-Bold").fontSize(11).text("STUDIO DELTA", MARGIN, MARGIN);
      doc.fillColor(MUTED).font("Helvetica").fontSize(8).text(
        (record.order_label || "") + "  ·  DELIVERY",
        MARGIN, MARGIN + 16
      );
      doc.fillColor(INK).font("Helvetica-Bold").fontSize(12).text(photo.name || "Photo", MARGIN, MARGIN, { width: inner, align: "right" });
      doc.save().strokeColor(BRASS).lineWidth(2).moveTo(MARGIN, MARGIN + 36).lineTo(MARGIN + inner, MARGIN + 36).stroke().restore();
      try {
        doc.image(photo.buf, MARGIN, MARGIN + 48, { fit: [inner, PAGE_H - MARGIN * 2 - 48] });
      } catch (e) {
        doc.fillColor(MUTED).text("Photo could not be placed.", MARGIN, MARGIN + 60);
      }
    });
    doc.addPage();
    drawReportHeader(doc, record, logoPath);
    y = MARGIN + 96;
    doc.fillColor(MUTED).font("Helvetica-Bold").fontSize(8).text("SIGN-OFF", MARGIN, y);
    doc.fillColor(INK).font("Helvetica-Bold").fontSize(14).text(record.receiver_name || "—", MARGIN, y + 16);
    const sigH = PAGE_H - (y + 58) - MARGIN;
    drawPdfBox(doc, MARGIN, y + 48, inner, sigH);
    doc.fillColor(MUTED).font("Helvetica-Bold").fontSize(8).text("SIGNATURE", MARGIN + 12, y + 60);
    if (record.signatureBuf) {
      try {
        doc.image(record.signatureBuf, MARGIN + 24, y + 88, { fit: [inner - 48, sigH - 72] });
      } catch (e) {
        doc.fillColor(MUTED).text("Signature could not be placed.", MARGIN + 16, y + 88);
      }
    } else {
      doc.save().strokeColor(RULE).lineWidth(0.6).moveTo(MARGIN + 24, y + sigH).lineTo(MARGIN + inner - 24, y + sigH).stroke().restore();
    }
    doc.end();
  });
}

function writePhotoFiles(id, photos, signatureBuf) {
  const dir = photosDir(id);
  (photos || []).forEach((photo, i) => {
    if (!photo || !photo.buf) return;
    const name = String(i + 1).padStart(2, "0") + "-photo.jpg";
    fs.writeFileSync(path.join(dir, name), photo.buf);
  });
  if (signatureBuf && signatureBuf.length) {
    fs.writeFileSync(path.join(dir, "signature.jpg"), signatureBuf);
  }
}

async function submitPod(body, actorName) {
  const orderNumber = db.formatOrderId(body && body.order_number);
  if (!orderNumber) throw new Error("Order is required.");
  const units = relatedUnits(orderNumber, "out for delivery");
  if (!units.length) throw new Error("That order is not loaded on the truck.");
  const clientName = String((body && body.client_name) || units[0].client_name || "").trim();
  const clientIsReceiver = body && (body.client_is_receiver === true || body.client_is_receiver === "yes");
  const receiverName = clientIsReceiver
    ? clientName
    : String((body && body.receiver_name) || "").trim();
  if (!receiverName) throw new Error("Name of the person receiving is required.");
  const photos = await encodePhotos(body && body.photos);
  if (!photos.length) throw new Error("At least one photo is required.");
  const signatureBuf = await toJpeg(photoRaw(body && body.signature));
  if (!signatureBuf) throw new Error("Signature is required.");
  const lat = body && body.lat != null && body.lat !== "" ? Number(body.lat) : null;
  const lng = body && body.lng != null && body.lng !== "" ? Number(body.lng) : null;
  const deliveredAt = (body && body.delivered_at) || new Date().toISOString();
  const driver = String(actorName || (body && body.driver) || "").trim();
  const id = newId();
  const orderLabel = units.map((u) => u.order_number).join(", ");
  const record = {
    id,
    kind: "Delivery",
    order_numbers: units.map((u) => u.order_number),
    order_label: orderLabel,
    base: units[0].base,
    client_name: clientName,
    receiver_name: receiverName,
    client_is_receiver: !!clientIsReceiver,
    drop_kind: units[0].drop_kind,
    drop_label: units[0].drop_label,
    address: units[0].drop_address || units[0].full_address,
    client_address: units[0].client_address || units[0].full_address,
    driver,
    lat: Number.isFinite(lat) ? lat : null,
    lng: Number.isFinite(lng) ? lng : null,
    delivered_at: deliveredAt,
    rating_delivery: Number(body && body.rating_delivery) || 0,
    rating_sales: Number(body && body.rating_sales) || 0,
    rating_craft: Number(body && body.rating_craft) || 0,
    comments: String((body && body.comments) || "").trim(),
    pdf_url: pdfUrlFor(id),
    created_at: new Date().toISOString(),
    layout: LAYOUT
  };
  const dest = path.join(formsDir(), id + ".pdf");
  await renderPdf(Object.assign({}, record, { photos, signatureBuf }), dest);
  writePhotoFiles(id, photos, signatureBuf);
  record.pdf_path = dest;
  const store = loadStore();
  store.records.unshift(record);
  saveStore(store);
  units.forEach((unit) => {
    const live = db.listOrders().find((row) => db.formatOrderId(row.order_number) === unit.order_number);
    if (live) setOrderStatus(live, "Delivered", driver);
  });
  closeOpenDeliveryLogs(record.order_numbers, deliveredAt, driver);
  try { db.persistOffice(); } catch (e) {}
  return {
    id,
    url: record.pdf_url,
    order_numbers: record.order_numbers,
    receiver_name: receiverName
  };
}

function listForms() {
  return loadStore().records.map((row) => ({
    id: row.id,
    order_label: row.order_label,
    order_numbers: row.order_numbers,
    client_name: row.client_name,
    receiver_name: row.receiver_name,
    address: row.address,
    client_address: row.client_address || "",
    drop_kind: row.drop_kind || "client",
    drop_label: row.drop_label || "",
    driver: row.driver,
    delivered_at: row.delivered_at,
    delivered_label: formatWhen(row.delivered_at),
    pdf_url: row.pdf_url || pdfUrlFor(row.id),
    rating_delivery: row.rating_delivery || 0,
    rating_sales: row.rating_sales || 0,
    rating_craft: row.rating_craft || 0
  }));
}

function readPdf(id) {
  const rec = loadStore().records.find((row) => row && row.id === String(id || ""));
  if (!rec) return null;
  const file = rec.pdf_path && fs.existsSync(rec.pdf_path) ? rec.pdf_path : path.join(formsDir(), rec.id + ".pdf");
  if (!fs.existsSync(file)) return null;
  return {
    buffer: fs.readFileSync(file),
    filename: "Studio-Delta-Delivery-" + String(rec.order_label || rec.id).replace(/[^\w.-]+/g, "-") + ".pdf"
  };
}

function bundleForOrder(orderNumber) {
  const loaded = relatedUnits(orderNumber, "out for delivery");
  if (loaded.length) return { status: "Out for Delivery", orders: loaded };
  const ready = relatedUnits(orderNumber, "ready for delivery");
  if (ready.length) return { status: "Ready for Delivery", orders: ready };
  return { status: "", orders: relatedUnits(orderNumber, "") };
}

module.exports = {
  FACTORY,
  THIRD_PARTY_DEPOT,
  isGautengProvince,
  isDriverProfile,
  canLoadTruck,
  canSubmitPod,
  orderBase,
  decorateOrder,
  relatedUnits,
  listReadyToLoad,
  listLoaded,
  loadOnTruck,
  buildRoute,
  geocodeAddress,
  defaultGeocode,
  haversineKm,
  wazeUrl,
  parseOsrmRoute,
  fetchOsrmDrive,
  submitPod,
  listForms,
  readPdf,
  pdfUrlFor,
  bundleForOrder,
  formatWhen
};
