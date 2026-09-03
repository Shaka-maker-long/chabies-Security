const {getBook, persistWorkbook} = require('./server/workbook-store');

async function main() {
  const book = getBook();
  
  // Get or create ORDERS sheet
  let ordersSheet = book.getSheetByName('ORDERS');
  if (!ordersSheet) {
    ordersSheet = book.insertSheet('ORDERS');
    ordersSheet.getRange(1, 1, 1, 12).setValues([["QUOTE NUMBER", "ORDER NUMBER", "STATUS", "ASSIGNED OPERATOR", "TYPE", "CATEGORY", "PRODUCT", "VARIATION", "DOORS", "DETAILED DESCRIPTION", "DIMENSIONS", "POWDER COATING"]]);
  }
  
  // Add GATE06 order with Welding status
  const lastRow = ordersSheet.getLastRow();
  ordersSheet.getRange(lastRow + 1, 1, 1, 7).setValues([["", "GATE06", "Welding", "", "", "", "Garden Gate"]]);
  console.log("Added GATE06 order");
  
  // Get or create Production_Log sheet
  let logSheet = book.getSheetByName('Production_Log');
  if (!logSheet) {
    logSheet = book.insertSheet('Production_Log');
    logSheet.getRange(1, 1, 1, 13).setValues([["id", "Order", "Worker", "Role", "Task", "Start", "End", "QC", "Opened-Approved", "Pause Start", "Pause (cumulative mins)", "Duration (minutes)", "Meta"]]);
  }
  
  // Add an active log entry for GATE06
  const logRow = logSheet.getLastRow();
  const now = new Date();
  const logId = "active-" + Date.now();
  logSheet.getRange(logRow + 1, 1, 1, 8).setValues([[
    logId,
    "GATE06",
    "Admin",
    "Admin",
    "Welding",
    now.toISOString(),
    "", // No end time - still active
    ""
  ]]);
  console.log("Added active production log for GATE06");
  
  // Persist the workbook
  await persistWorkbook();
  console.log("Workbook persisted");
}

main().catch(e => {
  console.error(e && e.stack || e);
  process.exit(1);
});
