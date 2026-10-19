"use strict";

/**
 * Isolated inventory prototype — data lives in inventory-prototype.json only.
 * Does not touch consumables, steel, glass, or live Orders rows.
 */

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
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
    purchases: (store.purchases || []).slice().sort((a, b) => String(b.createdAt || "").localeCompare(String(a.createdAt || ""))),
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

function createPurchase(body) {
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
  return purchase;
}

function receivePurchase(purchaseId) {
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
  return {
    purchase,
    labels: purchase.lines.map((line) => {
      const mat = materialById(store, line.materialId);
      return {
        materialId: line.materialId,
        code: line.code,
        name: line.name,
        unit: line.unit,
        qtyReceived: line.qty,
        qr: qrPayloadFor(mat || line),
        poNumber: purchase.number
      };
    })
  };
}

function labelsForPurchase(purchaseId) {
  const store = loadStore();
  const purchase = (store.purchases || []).find((p) => p.id === purchaseId);
  if (!purchase) throw new Error("Purchase order not found");
  if (purchase.status !== "received") {
    throw new Error("Receive the PO first, then print QR labels.");
  }
  return {
    purchase,
    labels: purchase.lines.map((line) => {
      const mat = materialById(store, line.materialId);
      return {
        materialId: line.materialId,
        code: line.code,
        name: line.name,
        unit: line.unit,
        qtyReceived: line.qty,
        qr: qrPayloadFor(mat || line),
        poNumber: purchase.number,
        onHand: mat ? Number(mat.onHand) || 0 : 0
      };
    })
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

  app.post("/api/inventory-prototype/purchases", (req, res) => {
    try {
      const purchase = createPurchase(req.body || {});
      res.json({ ok: true, purchase, ...snapshot() });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.post("/api/inventory-prototype/purchases/:id/receive", (req, res) => {
    try {
      const out = receivePurchase(req.params.id);
      res.json({ ok: true, ...out, ...snapshot() });
    } catch (e) {
      res.status(400).json({ ok: false, error: e.message || String(e) });
    }
  });

  app.get("/api/inventory-prototype/purchases/:id/labels", (req, res) => {
    try {
      res.json({ ok: true, ...labelsForPurchase(req.params.id) });
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
  qrPayloadFor
};
