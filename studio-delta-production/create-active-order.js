const { callShopFunction } = require("./server/gas");
const { getBook } = require("./server/workbook-store");
const db = require("./server/db");

async function main() {
  // Create a new order with Assembly status
  db.upsertOrder({
    order_number: "GATE06",
    status: "Ready for Welding",
    product: "Garden Gate",
    client_name: "Test Client E"
  });

  console.log("Created GATE06 order");

  // Find the row
  const book = getBook();
  const ordersSheet = book.getSheetByName("ORDERS");
  const lastRow = ordersSheet.getLastRow();
  
  let gateRow = 0;
  if (lastRow >= 2) {
    for (let i = 2; i <= lastRow; i++) {
      const val = ordersSheet.getRange(i, 2).getValue();
      if (String(val).trim() === "GATE06") {
        gateRow = i;
        break;
      }
    }
  }

  if (gateRow > 0) {
    console.log("Starting Welding on GATE06 at row", gateRow);
    const result = await callShopFunction("startOrder", [gateRow, "Admin", "Welding"]);
    console.log("Started:", result);
    
    // Don't finish it - leave it running
    console.log("GATE06 should now be actively being worked on by Admin doing Welding");
  }
}

main().catch(e => {
  console.error(e && e.stack || e);
  process.exit(1);
});
