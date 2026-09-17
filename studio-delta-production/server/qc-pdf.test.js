const fs = require("fs");
const os = require("os");
const path = require("path");
const assert = require("assert");

const dir = fs.mkdtempSync(path.join(os.tmpdir(), "sdp-qcpdf-"));
process.env.DATA_DIR = dir;
process.env.OFFICE_DB_PATH = path.join(dir, "studio-delta.json");
process.env.TZ = "Africa/Johannesburg";
delete process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
delete process.env.GOOGLE_APPLICATION_CREDENTIALS;

const { initWorkbook, getBook, persistWorkbook } = require("./workbook-store");
const qcPdf = require("./qc-pdf");
const sharp = require("sharp");

initWorkbook();

const PNG_1x1 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function pdfPageCount(buf) {
  const text = buf.toString("latin1");
  const pages = text.match(/\/Type\s*\/Page\b/g) || [];
  return pages.filter((row) => !/\/Pages\b/.test(row)).length;
}

function pdfText(buf) {
  const latin = buf.toString("latin1");
  const decoded = [];
  latin.replace(/<([0-9A-Fa-f]+)>/g, (_, hex) => {
    try { decoded.push(Buffer.from(hex, "hex").toString("latin1")); } catch (e) {}
    return "";
  });
  return latin + "\n" + decoded.join("");
}

assert.strictEqual(qcPdf.parseDriveFileId("https://drive.google.com/file/d/abc123xyz/view"), "abc123xyz");
assert.deepStrictEqual(
  qcPdf.driveIdsFromNotes("QC PDF: https://drive.google.com/file/d/abc123xyz\nQC PDF: /api/qc-pdfs/local/pdf"),
  ["abc123xyz"]
);
assert.deepStrictEqual(qcPdf.parseAnswersFromNotes(
  "Is the frame square?: Y\nAre the overall dimensions correct?: N\n\nQC PDF: https://example/x\nError adding to PDF Queue: Drive"
), [
  { q: "Is the frame square?", a: "Y" },
  { q: "Are the overall dimensions correct?", a: "N" }
]);

(async function main() {
  const jpeg = await sharp({
    create: { width: 120, height: 80, channels: 3, background: { r: 180, g: 40, b: 40 } }
  }).jpeg({ quality: 70 }).toBuffer();
  const jpegB64 = jpeg.toString("base64");
  const pngDataUrl = "data:image/png;base64," + PNG_1x1;
  const shot = { name: "shot.jpg", mime: "image/jpeg", data: jpegB64 };
  const finalFiles = [shot, shot, shot, shot, shot, shot, shot, null, null];

  assert.strictEqual(qcPdf.requiredPhotoCount("Final QC"), 7);
  assert.strictEqual(qcPdf.requiredPhotoCount("Pre-Powder Coating"), 6);

  const saved = await qcPdf.saveFromFinish({
    orderNum: "S260184",
    workerName: "Siya",
    processName: "Final QC",
    productName: "Violet Sideboard 3-Door",
    qcData: [
      { q: "Is the paint free of scratches?", a: "Y" },
      { q: "Is the product level?", a: "Y" }
    ],
    signatureUrl: pngDataUrl,
    filesData: finalFiles,
    rowToUpdate: 0,
    logId: "log-1"
  });
  assert.ok(saved.id);
  assert.ok(saved.url.indexOf("/api/qc-pdfs/") === 0);
  assert.ok(/Final QC/.test(saved.name), saved.name);
  assert.strictEqual(saved.photo_count, 7, JSON.stringify(saved));
  const listed = qcPdf.listReports();
  assert.strictEqual(listed.length, 1);
  assert.strictEqual(listed[0].order_number, "S260184");
  assert.strictEqual(listed[0].photo_count, 7);
  const file = qcPdf.readPdf(saved.id);
  assert.ok(file && file.buffer && file.buffer.slice(0, 4).toString() === "%PDF");
  const pages = pdfPageCount(file.buffer);
  assert.ok(pages >= 9, "checklist + 7 photos + signature last, got " + pages);
  const latin = file.buffer.toString("latin1");
  assert.ok(/DCTDecode/.test(latin), "JPEG photos must be embedded");
  const text = pdfText(file.buffer);
  assert.ok(text.indexOf("STUDIO DELTA") !== -1, "QC PDF must use the Studio Delta header");
  assert.ok(text.indexOf("FINAL QC") !== -1, text.slice(0, 400));
  assert.ok(text.indexOf("CHECKLIST") !== -1 || text.indexOf("Checklist") !== -1, "checklist table header");
  assert.ok(text.indexOf("SIGN-OFF") !== -1, "signature is a last-page sign-off");
  assert.ok(text.indexOf("S260184") !== -1);
  assert.ok(text.indexOf("Furniture") !== -1, "same tagline as purchase orders");
  const photoDir = path.join(dir, "qc-pdfs", saved.id);
  assert.ok(fs.existsSync(path.join(photoDir, "signature.jpg")));
  const jpgs = fs.readdirSync(photoDir).filter((name) => /\.jpg$/i.test(name) && name !== "signature.jpg");
  assert.strictEqual(jpgs.length, 7, jpgs.join(","));

  let missingThrew = "";
  try {
    await qcPdf.saveFromFinish({
      orderNum: "S260212",
      workerName: "Siya",
      processName: "Final QC",
      productName: "Violet Sideboard 3-Door",
      qcData: [{ q: "Is the frame square?", a: "Y" }],
      signatureUrl: pngDataUrl,
      filesData: [],
      rowToUpdate: 0,
      logId: "log-empty"
    });
  } catch (e) {
    missingThrew = String(e && e.message || e);
  }
  assert.ok(/photos/i.test(missingThrew), missingThrew);

  let stillThrew = "";
  try {
    await qcPdf.saveFromFinish({
      orderNum: "S260212",
      workerName: "Siya",
      processName: "Final QC",
      productName: "Violet Sideboard 3-Door",
      qcData: [{ q: "Is the frame square?", a: "Y" }],
      signatureUrl: pngDataUrl,
      filesData: [],
      rowToUpdate: 0,
      logId: "log-empty",
      allowMissingPhotos: true
    });
  } catch (e) {
    stillThrew = String(e && e.message || e);
  }
  assert.ok(/photos/i.test(stillThrew), "must not write a QC PDF with no images even for recovery");
  assert.ok(!qcPdf.listReports().some((row) => row.order_number === "S260212"), "no photo-less report is stored");

  const book = getBook();
  const logs = book.getSheetByName("Production_Log");
  logs.getRange(2, 1, 1, 9).setValues([[
    "log-old", "S260224", "Siya", "Quality Control", "Final QC", new Date(), new Date(),
    "Is the frame square?: Y\nDoes the unit match the drawing?: Y",
    ""
  ]]);
  logs.getRange(3, 1, 1, 9).setValues([[
    "log-drive", "S260214", "Siya", "Quality Control", "Final QC", new Date(), new Date(),
    "Is the frame square?: Y\n\nQC PDF: https://drive.google.com/file/d/abc123",
    ""
  ]]);
  persistWorkbook();
  const filled = await qcPdf.backfillFromLogs();
  assert.strictEqual(filled.saved, 0, JSON.stringify(filled));
  assert.ok(!qcPdf.listReports().some((row) => row.order_number === "S260224"), "do not invent a photo-less QC PDF from the log");
  const notes = String(logs.getRange(2, 8).getValue() || "");
  assert.ok(!/QC PDF:\s*\/api\/qc-pdfs\//.test(notes), notes);
  const driveNotes = String(logs.getRange(3, 8).getValue() || "");
  assert.ok(/drive\.google\.com/.test(driveNotes), driveNotes);
  const again = await qcPdf.backfillFromLogs();
  assert.strictEqual(again.saved, 0, JSON.stringify(again));

  const attached = await qcPdf.attachPhotos({
    logId: "log-blank-photos",
    orderNum: "S260300",
    processName: "Final QC",
    signatureUrl: pngDataUrl,
    filesData: [{ name: "front.jpg", mime: "image/jpeg", data: jpegB64 }]
  });
  assert.ok(attached.photo_count >= 1, JSON.stringify(attached));
  assert.ok(pdfPageCount(qcPdf.readPdf(attached.id).buffer) >= 3, "checklist + photo + signature last");
  assert.ok(pdfText(qcPdf.readPdf(attached.id).buffer).indexOf("SIGN-OFF") !== -1);

  console.log("qc-pdf.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
