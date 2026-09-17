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

initWorkbook();

assert.deepStrictEqual(qcPdf.parseAnswersFromNotes(
  "Is the frame square?: Y\nAre the overall dimensions correct?: N\n\nQC PDF: https://example/x\nError adding to PDF Queue: Drive"
), [
  { q: "Is the frame square?", a: "Y" },
  { q: "Are the overall dimensions correct?", a: "N" }
]);

(async function main() {
  const saved = await qcPdf.saveFromFinish({
    orderNum: "S260212",
    workerName: "Siya",
    processName: "Final QC",
    productName: "Violet Sideboard 3-Door",
    qcData: [{ q: "Is the frame square?", a: "Y" }],
    signatureUrl: "",
    filesData: [],
    rowToUpdate: 0,
    logId: "log-1"
  });
  assert.ok(saved.id);
  assert.ok(saved.url.indexOf("/api/qc-pdfs/") === 0);
  assert.ok(/Final QC/.test(saved.name), saved.name);
  const listed = qcPdf.listReports();
  assert.strictEqual(listed.length, 1);
  assert.strictEqual(listed[0].order_number, "S260212");
  const file = qcPdf.readPdf(saved.id);
  assert.ok(file && file.buffer && file.buffer.slice(0, 4).toString() === "%PDF");

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
  const again = await qcPdf.backfillFromLogs();
  assert.strictEqual(again.saved, 0, JSON.stringify(again));

  console.log("qc-pdf.test.js ok");
})().catch((e) => {
  console.error(e && e.stack ? e.stack : e);
  process.exit(1);
});
