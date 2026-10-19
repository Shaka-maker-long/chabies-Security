"use strict";

/**
 * Isolated inventory prototype — data lives in inventory-prototype.json only.
 * Does not touch consumables, steel, glass, or live Orders rows.
 */

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const QRCode = require("qrcode");
const { dataDir } = require("./workbook-store");

const QR_PREFIX = "SDINV|MAT|";

function storePath() {
  return path.join(dataDir(), "inventory-prototype.json");
}

function nowIso() {
  return new Date().toISOString();
}

function id(prefix) {
  return prefix + "-" + crypto.randomBytes(4).toString("hex");
}

function emptyStore() {
  return {
    materials: [
      { id: "mat-a", code: "A", name: "Material A", unit: "pcs", onHand: 0, wip: 0 },
      { id: "mat-b", code: "B", name: "Material B", unit: "pcs", onHand: 0, wip: 0 }
    ],
    purchases: [],
    workers: [
      { id: "w1", name: "Thabo", activeOrderNumber: "S260301" },
      { id: "w2", name: "Lerato", activeOrderNumber: "S260302 A" }
    ],
    demoOrders: ["S260301", "S260302 A", "S260302 B", "S260310"],
    allocations: [],
    poSeq: 1000
  };
}

function loadStore() {
  try {
    const parsed = JSON.parse(fs.readFileSync(storePath(), "utf8"));
    const base = emptyStore();
    return {
      materials: Array.isArray(parsed.materials) && parsed.materials.length ? parsed.materials : base.materials,
      purchases: Array.isArray(parsed.purchases) ? parsed.purchases : [],
      workers: Array.isArray(parsed.workers) && parsed.workers.length ? parsed.workers : base.workers,
      demoOrders: Array.isArray(parsed.demoOrders) && parsed.demoOrders.length ? parsed.demoOrders : base.demoOrders,
      allocations: Array.isArray(parsed.allocations) ? parsed.allocations : [],
      poSeq: Number(parsed.poSeq) > 0 ? Number(parsed.poSeq) : 1000
    };
  } catch (e) {
    if (e && e.code !== "ENOENT") {
      console.error("[inventory-prototype] read failed", e.message || e);
    }
    return emptyStore();
  }
}

function saveStore(store) {
  const file = storePath();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(store, null, 2));
  fs.renameSync(tmp, file);
  return store;
}

function materialById(store, materialId) {
  return (store.materials || []).find((m) => m.id === materialId) || null;
}

function materialByCode(store, code) {
  const want = String(code || "").trim().toUpperCase();
  return (store.materials || []).find((m) => String(m.code || "").toUpperCase() === want) || null;
}

function qrPayloadFor(material) {
  return QR_PREFIX + String(material.code || "").toUpperCase();
}

function pdfDir() {
  const dir = path.join(dataDir(), "inventory-prototype-pdfs");
  fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function poPdfPath(purchaseId) {
  return path.join(pdfDir(), String(purchaseId) + "-po.pdf");
}

function labelsPdfPath(purchaseId) {
  return path.join(pdfDir(), String(purchaseId) + "-qr-labels.pdf");
}

function cmToPt(cm) {
  return (Number(cm) * 72) / 2.54;
}

async function qrPngBuffer(payload) {
  return QRCode.toBuffer(String(payload || ""), {
    type: "png",
    width: 480,
    margin: 1,
    errorCorrectionLevel: "M"
  });
}

async function qrDataUrl(payload) {
  return QRCode.toDataURL(String(payload || ""), {
    width: 280,
    margin: 1,
    errorCorrectionLevel: "M"
  });
}

function writePdfBuffer(doc) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    doc.on("data", (c) => chunks.push(c));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
    doc.end();
  });
}

async function buildPurchaseOrderPdf(purchase) {
  const PDFDocument = require("pdfkit");
  const doc = new PDFDocument({
    size: "A4",
    margin: 40,
    info: { Title: "Purchase Order " + (purchase.number || ""), Author: "Studio Delta" }
  });
  const margin = 40;
  const pageW = doc.page.width;
  const inner = pageW - margin * 2;
  const ink = "#1c1914";
  const muted = "#6b6358";
  const brass = "#8a6a2f";

  doc.fillColor(ink).font("Helvetica-Bold").fontSize(18).text("STUDIO DELTA", margin, margin);
  doc.fillColor(muted).font("Helvetica").fontSize(9).text("Inventory prototype · Purchase order", margin, margin + 22);
  doc.fillColor(ink).font("Helvetica-Bold").fontSize(16).text("PURCHASE ORDER", margin, margin, { width: inner, align: "right" });
  doc.fillColor(brass).font("Helvetica-Bold").fontSize(12).text(purchase.number || "", margin, margin + 22, { width: inner, align: "right" });

  doc.moveTo(margin, margin + 48).lineTo(margin + inner, margin + 48).strokeColor(brass).lineWidth(1.5).stroke();

  let y = margin + 62;
  doc.fillColor(muted).font("Helvetica-Bold").fontSize(8).text("SUPPLIER", margin, y);
  doc.fillColor(ink).font("Helvetica").fontSize(11).text(purchase.supplier || "—", margin, y + 12);
  doc.fillColor(muted).font("Helvetica-Bold").fontSize(8).text("DATE", margin + inner * 0.55, y);
  doc.fillColor(ink).font("Helvetica").fontSize(11).text(String(purchase.createdAt || "").slice(0, 10), margin + inner * 0.55, y + 12);

  y += 44;
  doc.fillColor(ink).font("Helvetica").fontSize(10)
    .text("Please supply the materials listed below.", margin, y);
  y += 22;

  const cols = [
    { h: "Code", w: 60 },
    { h: "Material", w: 220 },
    { h: "Unit", w: 50 },
    { h: "Qty", w: inner - 60 - 220 - 50 }
  ];
  doc.save().fillColor("#1c1914").rect(margin, y, inner, 22).fill().restore();
  let x = margin;
  cols.forEach((c) => {
    doc.fillColor("#fffdf8").font("Helvetica-Bold").fontSize(8).text(c.h, x + 4, y + 7, { width: c.w - 8 });
    x += c.w;
  });
  y += 22;

  (purchase.lines || []).forEach((line, idx) => {
    if (y > 760) {
      doc.addPage();
      y = margin;
    }
    if (idx % 2) doc.save().fillColor("#f3eee4").rect(margin, y, inner, 22).fill().restore();
    const cells = [line.code, line.name, line.unit || "pcs", String(line.qty)];
    x = margin;
    cells.forEach((cell, i) => {
      doc.fillColor(ink).font(i === 0 ? "Helvetica-Bold" : "Helvetica").fontSize(10)
        .text(cell, x + 4, y + 6, { width: cols[i].w - 8 });
      x += cols[i].w;
    });
    y += 22;
  });

  y += 28;
  doc.fillColor(muted).font("Helvetica").fontSize(8)
    .text("Prototype document — not the live consumables PO system.", margin, y, { width: inner });

  const buf = await writePdfBuffer(doc);
  fs.writeFileSync(poPdfPath(purchase.id), buf);
  return buf;
}

/** Cut-out sheet: each QR is exactly 5cm × 5cm for locker labels. */
async function buildQrLabelsPdf(purchase) {
  const PDFDocument = require("pdfkit");
  const size = cmToPt(5); // 5cm × 5cm
  const gap = cmToPt(0.6);
  const margin = 36;
  const doc = new PDFDocument({
    size: "A4",
    margin,
    info: { Title: "QR labels " + (purchase.number || ""), Author: "Studio Delta" }
  });
  const pageW = doc.page.width;
  const pageH = doc.page.height;
  const cols = Math.max(1, Math.floor((pageW - margin * 2 + gap) / (size + gap)));
  let col = 0;
  let x = margin;
  let y = margin + 28;

  doc.fillColor("#1c1914").font("Helvetica-Bold").fontSize(12)
    .text("Studio Delta · Material QR labels — cut on the dashed lines", margin, margin, { width: pageW - margin * 2 });
  doc.fillColor("#6b6358").font("Helvetica").fontSize(8)
    .text((purchase.number || "") + " · each square is 5 cm × 5 cm · stick on the locker", margin, margin + 14, {
      width: pageW - margin * 2
    });

  for (const line of purchase.lines || []) {
    if (y + size > pageH - margin) {
      doc.addPage();
      y = margin;
      col = 0;
      x = margin;
    }
    const qr = qrPayloadFor(line);
    const png = await qrPngBuffer(qr);

    // Cut guide
    doc.save()
      .lineWidth(0.6)
      .dash(3, { space: 2 })
      .strokeColor("#9a9080")
      .rect(x, y, size, size)
      .stroke()
      .undash()
      .restore();

    const pad = cmToPt(0.35);
    const imgSize = size - pad * 2 - 18;
    doc.image(png, x + (size - imgSize) / 2, y + pad, { width: imgSize, height: imgSize });
    doc.fillColor("#1c1914").font("Helvetica-Bold").fontSize(11)
      .text(String(line.code || ""), x, y + size - 16, { width: size, align: "center" });

    col += 1;
    if (col >= cols) {
      col = 0;
      x = margin;
      y += size + gap;
    } else {
      x += size + gap;
    }
  }

  const buf = await writePdfBuffer(doc);
  fs.writeFileSync(labelsPdfPath(purchase.id), buf);
  return buf;
}

async function labelRowsForPurchase(store, purchase) {
  const out = [];
  for (const line of purchase.lines || []) {
    const mat = materialById(store, line.materialId);
    const qr = qrPayloadFor(mat || line);
    out.push({
      materialId: line.materialId,
      code: line.code,
      name: line.name,
      unit: line.unit,
      qtyReceived: line.qty,
      qr,
      qrImage: await qrDataUrl(qr),
      poNumber: purchase.number,
      onHand: mat ? Number(mat.onHand) || 0 : 0
    });
  }
  return out;
}

function decoratePurchase(p) {
  const hasPoPdf = !!(p && p.id && fs.existsSync(poPdfPath(p.id)));
  const hasLabelsPdf = !!(p && p.id && fs.existsSync(labelsPdfPath(p.id)));
  return Object.assign({}, p, {
    hasPoPdf,
    hasLabelsPdf,
    poPdfUrl: hasPoPdf ? ("/api/inventory-prototype/purchases/" + encodeURIComponent(p.id) + "/po.pdf") : "",
    labelsPdfUrl: hasLabelsPdf ? ("/api/inventory-prototype/purchases/" + encodeURIComponent(p.id) + "/qr-labels.pdf") : ""
  });
}

function parseQrPayload(raw) {
  const s = String(raw || "").trim();
  if (!s) return null;
  const m = s.match(/^SDINV\|MAT\|([A-Za-z0-9_-]+)$/i);
  if (m) return m[1].toUpperCase();
  // Allow typing just A / B in the prototype collect box
  if (/^[A-Za-z0-9_-]+$/.test(s)) return s.toUpperCase();
  return null;
}

function snapshot() {
  const store = loadStore();
  return {
    prototype: true,
    note: "Sandbox only — does not affect live inventory or Railway production data.",
    dataFile: "inventory-prototype.json",
    materials: store.materials.map((m) => Object.assign({}, m, {
      qr: qrPayloadFor(m),
      available: Number(m.onHand) || 0
    })),
    purchases: (store.purchases || []).slice().sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || ""))).map(decoratePurchase),
    workers: store.workers,
    demoOrders: store.demoOrders,
    allocations: (store.allocations || []).slice().sort((a, b) => String(b.at || "").localeCompare(String(a.at || ""))).slice(0, 40),
    wipByOrder: wipByOrder(store)
  };
}

function wipByOrder(store) {
  const map = {};
  (store.allocations || []).forEach((a) => {
    const order = String(a.orderNumber || "").trim() || "(no order)";
    if (!map[order]) map[order] = { orderNumber: order, lines: {}, totalQty: 0 };
    const code = a.materialCode || a.materialId;
    map[order].lines[code] = (map[order].lines[code] || 0) + (Number(a.qty) || 0);
    map[order].totalQty += Number(a.qty) || 0;
  });
  return Object.keys(map).map((k) => {
    const row = map[k];
    return {
      orderNumber: row.orderNumber,
      totalQty: row.totalQty,
      lines: Object.keys(row.lines).map((code) => ({ code, qty: row.lines[code] }))
    };
  }).sort((a, b) => a.orderNumber.localeCompare(b.orderNumber));
}

async function createPurchase(body) {
  const store = loadStore();
  const linesIn = Array.isArray(body && body.lines) ? body.lines : [];
  if (!linesIn.length) throw new Error("Add at least one material line.");
  const lines = [];
  linesIn.forEach((line, i) => {
    const mat = materialById(store, line.materialId) || materialByCode(store, line.code);
    if (!mat) throw new Error("Unknown material on line " + (i + 1));
    const qty = Number(line.qty);
    if (!Number.isFinite(qty) || qty <= 0) throw new Error("Quantity must be more than 0 for " + mat.code);
    lines.push({
      materialId: mat.id,
      code: mat.code,
      name: mat.name,
      unit: mat.unit,
      qty: Math.round(qty * 1000) / 1000
    });
  });
  store.poSeq = (Number(store.poSeq) || 1000) + 1;
  const purchase = {
    id: id("po"),
    number: "PROTO-PO-" + store.poSeq,
    status: "ordered",
    supplier: String((body && body.supplier) || "Prototype supplier").trim() || "Prototype supplier",
    lines,
    createdAt: nowIso(),
    receivedAt: null
  };
  store.purchases.unshift(purchase);
  saveStore(store);
  await buildPurchaseOrderPdf(purchase);
  return decoratePurchase(purchase);
}

async function receivePurchase(purchaseId) {
  const store = loadStore();
  const purchase = (store.purchases || []).find((p) => p.id === purchaseId);
  if (!purchase) throw new Error("Purchase order not found");
  if (purchase.status === "received") throw new Error("This PO is already received");
  purchase.lines.forEach((line) => {
    const mat = materialById(store, line.materialId);
    if (!mat) throw new Error("Material missing: " + line.code);
    mat.onHand = Math.round(((Number(mat.onHand) || 0) + Number(line.qty)) * 1000) / 1000;
  });
  purchase.status = "received";
  purchase.receivedAt = nowIso();
  saveStore(store);
  if (!fs.existsSync(poPdfPath(purchase.id))) {
    await buildPurchaseOrderPdf(purchase);
  }
  await buildQrLabelsPdf(purchase);
  return {
    purchase: decoratePurchase(purchase),
    labels: await labelRowsForPurchase(store, purchase)
  };
}

async function labelsForPurchase(purchaseId) {
  const store = loadStore();
  const purchase = (store.purchases || []).find((p) => p.id === purchaseId);
  if (!purchase) throw new Error("Purchase order not found");
  if (purchase.status !== "received") {
    throw new Error("Receive the PO first, then print QR labels.");
  }
  return {
    purchase,
    labels: await labelRowsForPurchase(store, purchase)
  };
}

function setActiveJob(body) {
  const store = loadStore();
  const workerId = String((body && body.workerId) || "").trim();
  const orderNumber = String((body && body.orderNumber) || "").trim();
  const worker = (store.workers || []).find((w) => w.id === workerId);
  if (!worker) throw new Error("Pick a worker");
  if (!orderNumber) throw new Error("Active job number is required");
  worker.activeOrderNumber = orderNumber;
  if (store.demoOrders.indexOf(orderNumber) === -1) store.demoOrders.push(orderNumber);
  saveStore(store);
  return worker;
}

function allocate(body) {
  const store = loadStore();
  const workerId = String((body && body.workerId) || "").trim();
  const worker = (store.workers || []).find((w) => w.id === workerId);
  if (!worker) throw new Error("Pick who is collecting");
  const orderNumber = String(worker.activeOrderNumber || "").trim();
  if (!orderNumber) throw new Error("This person has no active job — set one first");

  const code = parseQrPayload(body && (body.qr || body.code));
  if (!code) throw new Error("Scan or enter a material QR (e.g. SDINV|MAT|A)");
  const mat = materialByCode(store, code);
  if (!mat) throw new Error("Unknown material code " + code);

  const qty = Number(body && body.qty);
  if (!Number.isFinite(qty) || qty <= 0) throw new Error("Quantity must be more than 0");
  const take = Math.round(qty * 1000) / 1000;
  const available = Number(mat.onHand) || 0;
  if (take > available) {
    throw new Error("Only " + available + " " + (mat.unit || "") + " of " + mat.code + " available");
  }

  mat.onHand = Math.round((available - take) * 1000) / 1000;
  mat.wip = Math.round(((Number(mat.wip) || 0) + take) * 1000) / 1000;

  const row = {
    id: id("alloc"),
    materialId: mat.id,
    materialCode: mat.code,
    materialName: mat.name,
    qty: take,
    unit: mat.unit,
    orderNumber,
    workerId: worker.id,
    workerName: worker.name,
    at: nowIso()
  };
  store.allocations.unshift(row);
  saveStore(store);
  return {
    allocation: row,
    material: Object.assign({}, mat, { qr: qrPayloadFor(mat), available: Number(mat.onHand) || 0 }),
    worker
  };
}

function resetPrototype() {
  const store = emptyStore();
  saveStore(store);
  return snapshot();
}

function lookupMaterial(qrOrCode) {
  const store = loadStore();
  const code = parseQrPayload(qrOrCode);
  if (!code) throw new Error("Could not read material from QR");
  const mat = materialByCode(store, code);
  if (!mat) throw new Error("Unknown material " + code);
  return Object.assign({}, mat, { qr: qrPayloadFor(mat), available: Number(mat.onHand) || 0 });
}

function mount(app) {
  // No office login — prototype sandbox for phone + desktop demos.
  app.get("/api/inventory-prototype", (_req, res) => {
    res.json({ ok: true, ...snapshot() });
  });

  app.post("/api/inventory-prototype/purchases", async (req, res) => {
    try {
      const purchase = await createPurchase(req.body || {});
      res.json({ ok: true, purchase, ...snapshot() });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/inventory-prototype/purchases/:id/receive", async (req, res) => {
    try {
      const out = await receivePurchase(req.params.id);
      res.json({ ok: true, ...out, ...snapshot() });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/inventory-prototype/purchases/:id/labels", async (req, res) => {
    try {
      res.json({ ok: true, ...(await labelsForPurchase(req.params.id)) });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  function sendPdf(res, filePath, filename) {
    if (!fs.existsSync(filePath)) {
      res.status(404).json({ ok: false, error: "PDF not found" });
      return;
    }
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", "inline; filename=\"" + filename + "\"");
    res.send(fs.readFileSync(filePath));
  }

  app.get("/api/inventory-prototype/purchases/:id/po.pdf", (req, res) => {
    const store = loadStore();
    const purchase = (store.purchases || []).find((p) => p.id === req.params.id);
    if (!purchase) {
      res.status(404).json({ ok: false, error: "Purchase order not found" });
      return;
    }
    sendPdf(res, poPdfPath(purchase.id), (purchase.number || "PO") + ".pdf");
  });

  app.get("/api/inventory-prototype/purchases/:id/qr-labels.pdf", async (req, res) => {
    try {
      const store = loadStore();
      const purchase = (store.purchases || []).find((p) => p.id === req.params.id);
      if (!purchase) {
        res.status(404).json({ ok: false, error: "Purchase order not found" });
        return;
      }
      if (purchase.status !== "received") {
        res.status(400).json({ ok: false, error: "Receive the PO first to generate QR label PDF." });
        return;
      }
      if (!fs.existsSync(labelsPdfPath(purchase.id))) {
        await buildQrLabelsPdf(purchase);
      }
      sendPdf(res, labelsPdfPath(purchase.id), (purchase.number || "PO") + "-QR-labels.pdf");
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/inventory-prototype/active-job", (req, res) => {
    try {
      const worker = setActiveJob(req.body || {});
      res.json({ ok: true, worker, ...snapshot() });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/inventory-prototype/lookup", (req, res) => {
    try {
      const material = lookupMaterial(req.query.qr || req.query.code);
      res.json({ ok: true, material });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/inventory-prototype/allocate", (req, res) => {
    try {
      const out = allocate(req.body || {});
      res.json({ ok: true, ...out, ...snapshot() });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/inventory-prototype/reset", (_req, res) => {
    res.json({ ok: true, ...resetPrototype() });
  });
}

module.exports = {
  QR_PREFIX,
  mount,
  snapshot,
  createPurchase,
  receivePurchase,
  allocate,
  setActiveJob,
  resetPrototype,
  parseQrPayload,
  qrPayloadFor,
  buildPurchaseOrderPdf,
  buildQrLabelsPdf,
  cmToPt
};
