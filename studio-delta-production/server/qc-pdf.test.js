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
    filesData: [
      { name: "front.jpg", mime: "image/jpeg", data: jpegB64 },
      { name: "level1.jpg", mime: "image/jpeg", data: "data:image/jpeg;base64," + jpegB64 },
      null,
      { name: "left.jpg", mime: "image/heic", data: jpegB64 }
    ],
    rowToUpdate: 0,
    logId: "log-1"
  });
  assert.ok(saved.id);
  assert.ok(saved.url.indexOf("/api/qc-pdfs/") === 0);
  assert.ok(/Final QC/.test(saved.name), saved.name);
  assert.strictEqual(saved.photo_count, 3, JSON.stringify(saved));
  const listed = qcPdf.listReports();
  assert.strictEqual(listed.length, 1);
  assert.strictEqual(listed[0].order_number, "S260184");
  assert.strictEqual(listed[0].photo_count, 3);
  const file = qcPdf.readPdf(saved.id);
  assert.ok(file && file.buffer && file.buffer.slice(0, 4).toString() === "%PDF");
  const pages = pdfPageCount(file.buffer);
  assert.ok(pages >= 5, "checklist + signature + 3 photos, got " + pages);
  const latin = file.buffer.toString("latin1");
  assert.ok(/DCTDecode/.test(latin), "JPEG photos must be embedded");
  const photoDir = path.join(dir, "qc-pdfs", saved.id);
  assert.ok(fs.existsSync(path.join(photoDir, "signature.jpg")));
  const jpgs = fs.readdirSync(photoDir).filter((name) => /\.jpg$/i.test(name) && name !== "signature.jpg");
  assert.strictEqual(jpgs.length, 3, jpgs.join(","));

  const noPhotos = await qcPdf.saveFromFinish({
    orderNum: "S260212",
    workerName: "Siya",
    processName: "Final QC",
    productName: "Violet Sideboard 3-Door",
    qcData: [{ q: "Is the frame square?", a: "Y" }],
    signatureUrl: "https://drive.google.com/file/d/abc123",
    filesData: [],
    rowToUpdate: 0,
    logId: "log-empty"
  });
  const emptyPdf = qcPdf.readPdf(noPhotos.id);
  assert.ok(emptyPdf && emptyPdf.buffer.slice(0, 4).toString() === "%PDF");
  assert.strictEqual(noPhotos.photo_count, 0);
  assert.ok(pdfPageCount(emptyPdf.buffer) <= 2, "Drive URL is not an image");
  const store = JSON.parse(fs.readFileSync(path.join(dir, "qc-pdfs.json"), "utf8"));
  const emptyRec = store.records.find((row) => row && row.id === noPhotos.id);
  assert.ok(emptyRec);
  const rebuilt = await qcPdf.rebuildRecordFromPhotos(emptyRec, [
    { name: "front.jpg", mime: "image/jpeg", data: jpegB64 }
  ], pngDataUrl);
  assert.ok(rebuilt);
  const rebuiltPdf = qcPdf.readPdf(noPhotos.id);
  assert.ok(pdfPageCount(rebuiltPdf.buffer) >= 3, "saved checklist PDF can be rebuilt with photos");
  assert.strictEqual(qcPdf.listReports().find((row) => row.id === noPhotos.id).photo_count, 1);

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
  assert.ok(filled.saved >= 2, JSON.stringify(filled));
  const notes = String(logs.getRange(2, 8).getValue() || "");
  assert.ok(/QC PDF:\s*\/api\/qc-pdfs\//.test(notes), notes);
  const driveNotes = String(logs.getRange(3, 8).getValue() || "");
  assert.ok(/QC PDF:\s*\/api\/qc-pdfs\//.test(driveNotes), driveNotes);
  const linked = qcPdf.listReports().find((row) => row.order_number === "S260214");
  assert.ok(linked && /abc123/.test(linked.drive_url || linked.url), JSON.stringify(linked));
  const again = await qcPdf.backfillFromLogs();
  assert.strictEqual(again.saved, 0, JSON.stringify(again));

  console.log("qc-pdf.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
