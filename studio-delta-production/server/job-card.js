"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const { dataDir } = require("./workbook-store");
const { listOrders, upsertOrder, formatOrderId } = require("./db");
const catalog = require("./product-catalog");

const FIRST_STATUSES = new Set(["Not Yet Started", ""]);
const REGENERATE_STATUSES = new Set(["Ready for Steelwork", "Profile Cutting"]);
const AFTER_GENERATE_STATUS = "Ready for Steelwork";

function todayIso() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: process.env.TZ || "Africa/Johannesburg",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(new Date());
  const y = parts.find((p) => p.type === "year").value;
  const m = parts.find((p) => p.type === "month").value;
  const d = parts.find((p) => p.type === "day").value;
  return y + "-" + m + "-" + d;
}

function emptyCutting() {
  return { bars: [], plates: [], tubes: [], wood: [] };
}

function parsePastedCuttingList(text) {
  const out = emptyCutting();
  const lines = String(text || "").split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    const columns = trimmed.includes("\t")
      ? trimmed.split("\t").map((col) => col.trim())
      : trimmed.split(/\s{2,}/).map((col) => col.trim());
    if (columns.length < 5) continue;
    if (columns.some((c) => /^n\/a$/i.test(c))) continue;
    const materialName = columns[0];
    const description = columns[1];
    const dim1 = parseFloat(columns[2]) || 0;
    const dim2 = parseFloat(columns[3]) || 0;
    const quantity = parseInt(columns[4], 10) || 0;
    const cutAngle45 = parseFloat(columns[5]) || 0;
    if (!materialName || quantity <= 0) continue;

    const hay = (materialName + " " + description).toLowerCase();
    let category = "bars";
    if (/\b(plate|sheet)\b/.test(hay)) category = "plates";
    else if (/\btube\b/.test(hay)) category = "tubes";
    else if (/\b(wood|board|plank|mdf|plywood|timber)\b/.test(hay)) category = "wood";
    else if (/\b(bar|angle|flat|round|channel|beam)\b/.test(hay)) category = "bars";

    if ((category === "bars" || category === "tubes") && dim1 <= 0) continue;
    if ((category === "plates" || category === "wood") && (dim1 <= 0 || dim2 <= 0)) continue;

    if (category === "plates" || category === "wood") {
      out[category].push({
        name: materialName,
        description,
        height: dim1,
        width: dim2,
        quantity,
        totalArea: (dim1 * dim2) / 1000000
      });
    } else {
      out[category].push({
        name: materialName,
        description,
        length: dim1,
        quantity,
        cutAngle45
      });
    }
  }
  return out;
}

function cuttingCount(cutting) {
  const c = cutting || emptyCutting();
  return (c.bars || []).length + (c.plates || []).length + (c.tubes || []).length + (c.wood || []).length;
}

function jobCardEligibility(status) {
  const s = String(status || "").trim();
  if (FIRST_STATUSES.has(s)) return { ok: true, mode: "create" };
  if (REGENERATE_STATUSES.has(s)) return { ok: true, mode: "regenerate" };
  return {
    ok: false,
    mode: "blocked",
    error: "A job card can only be generated when the order is Not Yet Started, Ready for Steelwork, or Profile Cutting."
  };
}

function applyOfficeOrderStatusLock(body, existing) {
  const next = Object.assign({}, body || {});
  if (!existing) {
    next.status = "Not Yet Started";
  } else {
    next.status = existing.status || "Not Yet Started";
  }
  return next;
}

function safeOrderFile(orderNumber) {
  return String(orderNumber || "").replace(/[^A-Za-z0-9._-]+/g, "_") || "job";
}

function jobCardsDir() {
  const dir = path.join(dataDir(), "job-cards");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function imageCacheDir() {
  const dir = path.join(dataDir(), "job-card-images");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function recordsPath() {
  return path.join(dataDir(), "job-cards.json");
}

function loadRecords() {
  try {
    const raw = JSON.parse(fs.readFileSync(recordsPath(), "utf8"));
    return raw && typeof raw === "object" ? raw : {};
  } catch (e) {
    return {};
  }
}

function saveRecords(map) {
  const tmp = recordsPath() + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(map));
  fs.renameSync(tmp, recordsPath());
}

function deleteAllJobCards() {
  saveRecords({});
  try { fs.rmSync(path.join(dataDir(), "job-cards"), { recursive: true, force: true }); } catch (e) {}
}

function getJobCard(orderNumber) {
  const key = formatOrderId(orderNumber);
  return loadRecords()[key] || null;
}

function pdfUrlFor(orderNumber, download) {
  const base = "/api/office/job-cards/" + encodeURIComponent(orderNumber) + "/pdf";
  return download ? base + "?download=1" : base;
}

function listGeneratedJobCards() {
  const map = loadRecords();
  const byNumber = {};
  listOrders().forEach((o) => {
    if (o && o.order_number) byNumber[o.order_number] = o;
  });
  return Object.keys(map).map((key) => {
    const rec = map[key] || {};
    const order = byNumber[key] || {};
    const orderNumber = rec.order_number || key;
    const hasPdf = !!(rec.pdf_path && fs.existsSync(rec.pdf_path));
    return {
      order_number: orderNumber,
      customer: rec.customer || order.client_name || "",
      product: rec.product || order.product || "",
      created_date: rec.created_date || "",
      created_at: rec.created_at || "",
      status: order.status || "",
      province: rec.province || order.province || "",
      colour: rec.colour || order.powder_coating || "",
      has_pdf: hasPdf,
      pdf_url: pdfUrlFor(orderNumber),
      download_url: pdfUrlFor(orderNumber, true)
    };
  }).sort((a, b) => {
    const at = String(b.created_at || "").localeCompare(String(a.created_at || ""));
    if (at) return at;
    return String(b.order_number || "").localeCompare(String(a.order_number || ""));
  });
}

function listEligibleOrders() {
  return listOrders()
    .filter((o) => jobCardEligibility(o.status).ok)
    .map((o) => {
      const elig = jobCardEligibility(o.status);
      const found = catalog.lookupProduct(o.product);
      const saved = getJobCard(o.order_number);
      return {
        order_number: o.order_number,
        status: o.status || "Not Yet Started",
        product: o.product || "",
        client_name: o.client_name || "",
        type: o.type || "",
        doors: o.doors || "",
        powder_coating: o.powder_coating || "",
        variation: o.variation || "",
        detailed_description: o.detailed_description || "",
        dimensions: o.dimensions || "",
        province: o.province || "",
        mode: elig.mode,
        has_job_card: !!saved,
        image_url: (found && found.imageUrl) || "",
        catalog_dimensions: found
          ? {
            height: found.height,
            width: found.width,
            depth: found.depth,
            diameter: found.diameter
          }
          : null
      };
    });
}

function normalizeDimensions(input, productName) {
  const found = catalog.lookupProduct(productName) || {};
  const src = input && typeof input === "object" ? input : {};
  const pick = (key) => {
    const raw = src[key] != null && src[key] !== "" ? src[key] : found[key];
    if (raw == null || raw === "") return null;
    const n = Number(raw);
    return Number.isFinite(n) ? n : null;
  };
  return {
    height: pick("height"),
    width: pick("width"),
    depth: pick("depth"),
    diameter: pick("diameter")
  };
}

function dimensionsRows(dims) {
  const rows = [];
  if (dims.height != null) rows.push({ name: "Height", value: dims.height });
  if (dims.width != null) rows.push({ name: "Width", value: dims.width });
  if (dims.depth != null) rows.push({ name: "Depth", value: dims.depth });
  if (dims.diameter != null) rows.push({ name: "Diameter", value: dims.diameter });
  return rows;
}

function dimensionsLabel(dims, fallback) {
  const rows = dimensionsRows(dims);
  if (!rows.length) return fallback || "Standard";
  return rows.map((r) => r.name + ": " + r.value + "mm").join(", ");
}

function isStandardType(type) {
  return String(type || "").trim().toLowerCase() === "standard";
}

function hasUsableDimensions(dims) {
  const values = [dims && dims.height, dims && dims.width, dims && dims.depth, dims && dims.diameter]
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value) && value > 1);
  return values.length > 0;
}

function assertNonStandardDimensions(order, body, dims) {
  if (isStandardType(order && order.type)) return "standard";
  const check = String(body && body.dimension_check || "").trim().toLowerCase();
  if (check !== "unchanged" && check !== "updated") {
    throw new Error(
      "This order is not Standard. Confirm that dimensions did not change, or type the actual Height, Width, Depth, or Diameter in millimetres before generating the job card."
    );
  }
  if (!hasUsableDimensions(dims)) {
    throw new Error(
      "Actual dimensions are required for " + String((order && order.type) || "this type") +
      ". Enter Height, Width, Depth, or Diameter in millimetres — catalog placeholders such as 1 × 1 × 1 are not accepted."
    );
  }
  return check;
}

function guessExt(url, contentType) {
  const type = String(contentType || "").toLowerCase();
  if (type.indexOf("png") !== -1) return ".png";
  if (type.indexOf("jpeg") !== -1 || type.indexOf("jpg") !== -1) return ".jpg";
  if (type.indexOf("webp") !== -1) return ".webp";
  const clean = String(url || "").split("?")[0].toLowerCase();
  if (clean.endsWith(".png")) return ".png";
  if (clean.endsWith(".jpg") || clean.endsWith(".jpeg")) return ".jpg";
  if (clean.endsWith(".webp")) return ".webp";
  if (clean.endsWith(".jpg.webp") || clean.endsWith(".jpeg.webp")) return ".webp";
  return ".img";
}

async function fetchToCache(url) {
  if (!url) return null;
  const key = crypto.createHash("sha1").update(url).digest("hex");
  const existing = fs.readdirSync(imageCacheDir()).find((name) => name.startsWith(key + "."));
  if (existing) return path.join(imageCacheDir(), existing);
  try {
    const res = await fetch(url, { redirect: "follow", signal: AbortSignal.timeout(8000) });
    if (!res.ok) return null;
    const buf = Buffer.from(await res.arrayBuffer());
    const ext = guessExt(url, res.headers.get("content-type"));
    const dest = path.join(imageCacheDir(), key + ext);
    fs.writeFileSync(dest, buf);
    return dest;
  } catch (e) {
    return null;
  }
}

async function imageForPdf(url) {
  if (!url) return null;
  const cached = await fetchToCache(url);
  if (!cached) return null;
  const ext = path.extname(cached).toLowerCase();
  if (ext === ".jpg" || ext === ".jpeg" || ext === ".png") return cached;
  const png = cached.replace(/\.[^.]+$/, "") + ".png";
  if (fs.existsSync(png)) return png;
  try {
    const sharp = require("sharp");
    await sharp(cached).png().toFile(png);
    return png;
  } catch (e) {
    const stripped = String(url).replace(/\.webp$/i, "");
    if (stripped !== url && /\.(jpe?g|png)$/i.test(stripped.split("?")[0])) {
      return imageForPdf(stripped);
    }
    return null;
  }
}

function drawBox(doc, x, y, w, h) {
  doc.save().lineWidth(0.8).strokeColor("#111").rect(x, y, w, h).stroke().restore();
}

function drawLabeled(doc, x, y, w, h, label, value, opts) {
  opts = opts || {};
  drawBox(doc, x, y, w, h);
  doc.save();
  doc.font("Helvetica-Bold").fontSize(8).fillColor("#111").text(String(label || "").toUpperCase(), x + 4, y + 4, {
    width: w - 8,
    lineBreak: false
  });
  doc.font(opts.bold === false ? "Helvetica" : "Helvetica-Bold").fontSize(opts.size || 10)
    .text(String(value == null || value === "" ? "" : value), x + 4, y + 16, {
      width: w - 8,
      height: h - 20
    });
  doc.restore();
}

function drawCheckRow(doc, x, y, w, h, label) {
  drawBox(doc, x, y, w, h);
  const yesW = 42;
  const noW = 42;
  const qW = w - yesW - noW;
  doc.save();
  doc.font("Helvetica-Bold").fontSize(9).fillColor("#111")
    .text(label, x + 6, y + (h - 10) / 2, { width: qW - 10, lineBreak: false });
  drawBox(doc, x + qW, y, yesW, h);
  drawBox(doc, x + qW + yesW, y, noW, h);
  doc.rect(x + qW + 14, y + 6, 12, 12).stroke();
  doc.rect(x + qW + yesW + 14, y + 6, 12, 12).stroke();
  doc.restore();
}

function drawTable(doc, title, headers, rows, colWeights) {
  const margin = 36;
  const pageW = doc.page.width - margin * 2;
  let y = doc.y;
  if (y > doc.page.height - 120) {
    doc.addPage();
    y = 36;
  }
  doc.font("Helvetica-Bold").fontSize(12).fillColor("#111").text(title, margin, y);
  y = doc.y + 8;
  const widths = colWeights.map((w) => pageW * w);
  const headerH = 18;
  let x = margin;
  headers.forEach((h, i) => {
    drawBox(doc, x, y, widths[i], headerH);
    doc.font("Helvetica-Bold").fontSize(8).text(h, x + 3, y + 5, { width: widths[i] - 6, lineBreak: false });
    x += widths[i];
  });
  y += headerH;
  rows.forEach((row) => {
    const rowH = 18;
    if (y + rowH > doc.page.height - 36) {
      doc.addPage();
      y = 36;
    }
    x = margin;
    row.forEach((cell, i) => {
      drawBox(doc, x, y, widths[i], rowH);
      doc.font("Helvetica").fontSize(8).text(String(cell == null ? "" : cell), x + 3, y + 5, {
        width: widths[i] - 6,
        lineBreak: false
      });
      x += widths[i];
    });
    y += rowH;
  });
  doc.y = y + 16;
}

async function writeJobCardPdf(record, dest) {
  const PDFDocument = require("pdfkit");
  const margin = 36;
  const pageW = 595.28;
  const inner = pageW - margin * 2;
  await new Promise(async (resolve, reject) => {
    const doc = new PDFDocument({ size: "A4", margin: 28, info: {
      Title: "Job Card " + record.order_number,
      Author: "Studio Delta"
    } });
    const stream = fs.createWriteStream(dest);
    doc.pipe(stream);
    stream.on("finish", resolve);
    stream.on("error", reject);

    const logoPath = await imageForPdf(catalog.COMPANY_LOGO_URL);
    const logoSize = 78;
    drawBox(doc, margin, margin, logoSize, logoSize);
    if (logoPath) {
      try { doc.image(logoPath, margin + 4, margin + 4, { fit: [70, 70] }); } catch (e) {}
    } else {
      doc.font("Helvetica-Bold").fontSize(8).text("STUDIO\nDELTA", margin + 8, margin + 28, { width: 62, align: "center" });
    }
    const rightX = margin + logoSize;
    const rightW = inner - logoSize;
    drawBox(doc, rightX, margin, rightW, 42);
    doc.font("Helvetica-Bold").fontSize(22).text("JOB CARD", rightX, margin + 11, { width: rightW, align: "center" });
    drawLabeled(doc, rightX, margin + 42, rightW, 36, "Date job card created:", record.created_date);

    let y = margin + logoSize + 10;
    const rowH = 28;
    const col1 = inner * 0.30;
    const col2 = inner * 0.42;
    const col3 = inner - col1 - col2;

    drawLabeled(doc, margin, y, col1, rowH, "Assembler:", "");
    drawLabeled(doc, margin + col1, y, col2, rowH, "Design type:", record.design_type);
    drawLabeled(doc, margin + col1 + col2, y, col3, rowH, "Colour:", record.colour);
    y += rowH;
    drawLabeled(doc, margin, y, col1, rowH, "Date start:", record.start_date);
    drawLabeled(doc, margin + col1, y, col2, rowH, "Order number:", record.order_number);
    drawLabeled(doc, margin + col1 + col2, y, col3, rowH, "", "");
    y += rowH;
    drawLabeled(doc, margin, y, col1, rowH, "Time started:", "");
    drawLabeled(doc, margin + col1, y, col2, rowH, "Customer name:", record.customer);
    drawLabeled(doc, margin + col1 + col2, y, col3, rowH, "", "");
    y += rowH;
    drawLabeled(doc, margin, y, col1, rowH, "Date finished:", "");
    drawLabeled(doc, margin + col1, y, col2 + col3, rowH, "Product:", record.product);
    y += rowH;
    drawLabeled(doc, margin, y, col1, rowH, "Time finished:", "");
    drawLabeled(doc, margin + col1, y, col2 + col3, rowH, "Glass type:", record.glass_type);
    y += rowH;
    drawLabeled(doc, margin, y, inner, 40, "Variations:", record.variation);
    y += 40;
    drawLabeled(doc, margin, y, inner, 40, "Description:", record.description);
    y += 40;
    const dimW = inner * 0.62;
    drawLabeled(doc, margin, y, dimW, rowH, "Dimensions:", record.dimensions_string);
    drawLabeled(doc, margin + dimW, y, inner - dimW, rowH, "Province:", record.province);
    y += rowH + 12;

    drawBox(doc, margin, y, inner * 0.62, 20);
    drawBox(doc, margin + inner * 0.62, y, inner * 0.19, 20);
    drawBox(doc, margin + inner * 0.81, y, inner * 0.19, 20);
    doc.font("Helvetica-Bold").fontSize(9)
      .text("QC QUESTIONS:", margin + 6, y + 6)
      .text("YES", margin + inner * 0.62, y + 6, { width: inner * 0.19, align: "center" })
      .text("NO", margin + inner * 0.81, y + 6, { width: inner * 0.19, align: "center" });
    y += 20;
    ["DIMENSIONS", "OVERALL SQUARE", "GLASS SIZES", "ALL WELDS POLISHED", "HOLES DRILLED"].forEach((q) => {
      drawCheckRow(doc, margin, y, inner, 22, q);
      y += 22;
    });
    y += 16;
    doc.font("Helvetica-Bold").fontSize(12).text("SIGNATURE", margin, y, { width: inner, align: "center" });
    y += 20;
    ["EMPLOYEE:", "SUPERVISOR:", "QC SIGNATURE:"].forEach((label) => {
      drawLabeled(doc, margin, y, inner, 32, label, "", { bold: false });
      y += 32;
    });

    doc.addPage();
    doc.font("Helvetica-Bold").fontSize(14).text("Product", margin, margin);
    const productImage = await imageForPdf(record.image_url);
    if (productImage) {
      try {
        doc.image(productImage, margin, margin + 24, { fit: [260, 260] });
      } catch (e) {
        doc.font("Helvetica").fontSize(10).text("Product image could not be placed.", margin, margin + 24);
      }
    } else {
      drawBox(doc, margin, margin + 24, 260, 200);
      doc.font("Helvetica").fontSize(10).text("No product image saved for this item.", margin + 12, margin + 110, { width: 236 });
    }
    const dimRows = dimensionsRows(record.dimensions);
    doc.font("Helvetica-Bold").fontSize(12).text("Product dimensions (mm)", margin + 280, margin + 24);
    let dy = margin + 48;
    [["Dimension", "Value (mm)"], ...dimRows.map((r) => [r.name, r.value])].forEach((pair, i) => {
      drawBox(doc, margin + 280, dy, 120, 22);
      drawBox(doc, margin + 400, dy, 120, 22);
      doc.font(i === 0 ? "Helvetica-Bold" : "Helvetica").fontSize(9)
        .text(String(pair[0]), margin + 286, dy + 6, { width: 108 })
        .text(String(pair[1] == null ? "" : pair[1]), margin + 406, dy + 6, { width: 108 });
      dy += 22;
    });
    if (!dimRows.length) {
      doc.font("Helvetica").fontSize(9).text("No catalog dimensions for this product. Values on the card use the order dimensions text.", margin + 280, dy + 8, { width: 240 });
    }

    const cutting = record.cutting || emptyCutting();
    const sections = [
      ["Bars", cutting.bars, ["Description", "Name", "Length (mm)", "Qty", "Cut angle 45"], (item) => [item.description, item.name, item.length, item.quantity, item.cutAngle45 || 0], [0.22, 0.34, 0.16, 0.12, 0.16]],
      ["Tubes", cutting.tubes, ["Description", "Name", "Length (mm)", "Qty", "Cut angle 45"], (item) => [item.description, item.name, item.length, item.quantity, item.cutAngle45 || 0], [0.22, 0.34, 0.16, 0.12, 0.16]],
      ["Plates", cutting.plates, ["Description", "Name", "Height (mm)", "Width (mm)", "Qty"], (item) => [item.description, item.name, item.height, item.width, item.quantity], [0.22, 0.30, 0.16, 0.16, 0.16]],
      ["Wood", cutting.wood, ["Description", "Name", "Height (mm)", "Width (mm)", "Qty"], (item) => [item.description, item.name, item.height, item.width, item.quantity], [0.22, 0.30, 0.16, 0.16, 0.16]]
    ];
    let startedLists = false;
    sections.forEach(([title, items, headers, mapRow, weights]) => {
      if (!items || !items.length) return;
      if (!startedLists) {
        doc.addPage();
        doc.font("Helvetica-Bold").fontSize(14).text("Cutting list", 36, 36);
        doc.y = 58;
        startedLists = true;
      }
      drawTable(doc, title, headers, items.map(mapRow), weights);
    });
    if (!startedLists) {
      doc.addPage();
      doc.font("Helvetica-Bold").fontSize(14).text("Cutting list", 36, 36);
      doc.font("Helvetica").fontSize(10).text("No cutting list was pasted for this job card.", 36, 60);
    }

    doc.end();
  });
  return dest;
}

function findOrder(orderNumber) {
  const want = formatOrderId(orderNumber);
  return listOrders().find((o) => o.order_number === want) || null;
}

async function generateJobCard(input) {
  const body = input || {};
  const orderNumber = formatOrderId(body.order_number);
  if (!orderNumber) throw new Error("Choose an order number.");
  const order = findOrder(orderNumber);
  if (!order) throw new Error("Order " + orderNumber + " was not found.");
  const elig = jobCardEligibility(order.status);
  if (!elig.ok) throw new Error(elig.error);

  const cutting = body.cutting && typeof body.cutting === "object"
    ? {
      bars: Array.isArray(body.cutting.bars) ? body.cutting.bars : [],
      plates: Array.isArray(body.cutting.plates) ? body.cutting.plates : [],
      tubes: Array.isArray(body.cutting.tubes) ? body.cutting.tubes : [],
      wood: Array.isArray(body.cutting.wood) ? body.cutting.wood : []
    }
    : parsePastedCuttingList(body.cutting_text || body.pastedCuttingList || "");

  const dims = normalizeDimensions(body.dimensions, order.product);
  const dimensionCheck = assertNonStandardDimensions(order, body, dims);
  const created = todayIso();
  const found = catalog.lookupProduct(order.product);
  const record = {
    order_number: orderNumber,
    created_at: new Date().toISOString(),
    created_date: created,
    start_date: created,
    product: order.product || "",
    customer: order.client_name || "",
    design_type: order.type || "",
    colour: order.powder_coating || "",
    glass_type: order.doors || "N/A",
    variation: order.variation || "",
    description: order.detailed_description || order.product || "",
    province: order.province || "",
    dimensions: dims,
    dimensions_string: dimensionsLabel(dims, order.dimensions || "Standard"),
    image_url: (found && found.imageUrl) || body.image_url || "",
    cutting,
    regenerated: elig.mode === "regenerate",
    dimension_check: dimensionCheck
  };

  const dest = path.join(jobCardsDir(), safeOrderFile(orderNumber) + ".pdf");
  await writeJobCardPdf(record, dest);
  record.pdf_path = dest;
  const map = loadRecords();
  map[orderNumber] = record;
  saveRecords(map);

  const patch = {};
  if (elig.mode === "create") patch.status = AFTER_GENERATE_STATUS;
  if (!isStandardType(order.type)) patch.dimensions = dimensionsLabel(dims);
  let savedOrder = order;
  if (Object.keys(patch).length) {
    savedOrder = upsertOrder(Object.assign({}, order, patch));
  }

  return {
    record,
    order: savedOrder,
    status: savedOrder.status,
    pdf_url: pdfUrlFor(orderNumber)
  };
}

function readJobCardPdf(orderNumber) {
  const rec = getJobCard(orderNumber);
  if (!rec || !rec.pdf_path || !fs.existsSync(rec.pdf_path)) return null;
  return {
    buffer: fs.readFileSync(rec.pdf_path),
    filename: "JobCard_" + safeOrderFile(orderNumber) + ".pdf"
  };
}

module.exports = {
  FIRST_STATUSES,
  REGENERATE_STATUSES,
  AFTER_GENERATE_STATUS,
  todayIso,
  parsePastedCuttingList,
  cuttingCount,
  jobCardEligibility,
  applyOfficeOrderStatusLock,
  listEligibleOrders,
  listGeneratedJobCards,
  getJobCard,
  deleteAllJobCards,
  generateJobCard,
  readJobCardPdf,
  normalizeDimensions,
  emptyCutting,
  isStandardType,
  hasUsableDimensions,
  assertNonStandardDimensions,
  dimensionsLabel
};
