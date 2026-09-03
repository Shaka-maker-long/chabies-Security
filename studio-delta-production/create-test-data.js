const db = require("./server/db");
const { callShopFunction } = require("./server/gas");
const { getBook } = require("./server/workbook-store");

async function main() {
  const book = getBook();
  
  // Ensure Users sheet exists with Admin
  let usersSheet = book.getSheetByName("Users");
  if (!usersSheet) {
    usersSheet = book.insertSheet("Users");
    usersSheet.getRange(1, 1, 1, 4).setValues([["Name", "Role", "Password", "Tasks"]]);
  }
  
  // Add Admin user if not exists
  const lastUserRow = usersSheet.getLastRow();
  let adminExists = false;
  if (lastUserRow >= 2) {
    for (let i = 2; i <= lastUserRow; i++) {
      const name = usersSheet.getRange(i, 1).getValue();
      if (String(name).trim() === "Admin") {
        adminExists = true;
        // Update tasks to include Assembly
        usersSheet.getRange(i, 4).setValue("Welding, Profile Cutting, Assembly, Tagging");
        break;
      }
    }
  }
  
  if (!adminExists) {
    usersSheet.getRange(lastUserRow + 1, 1, 1, 4).setValues([["Admin", "Admin", "admin", "Welding, Profile Cutting, Assembly, Tagging"]]);
  }
  
  console.log("Ensured Admin user exists with Assembly task");

  // Create test orders with different statuses
  db.upsertOrder({
    order_number: "SCOL03",
    status: "Assembly",
    product: "Steel Column",
    client_name: "Test Client A"
  });

  db.upsertOrder({
    order_number: "GATE05",
    status: "Welding",
    product: "Garden Gate",
    client_name: "Test Client B"
  });

  db.upsertOrder({
    order_number: "NOT01",
    status: "Not Yet Started",
    product: "Balustrade",
    client_name: "Test Client C"
  });

  db.upsertOrder({
    order_number: "DELIV01",
    status: "Out for Delivery",
    product: "Fence Panels",
    client_name: "Test Client D"
  });

  console.log("Created 4 test orders");

  // Start work on SCOL03
  try {
    const ordersSheet = book.getSheetByName("ORDERS") || book.insertSheet("ORDERS");
    const lastRow = ordersSheet.getLastRow();
    
    // Find SCOL03 row
    let scol03Row = 0;
    if (lastRow >= 2) {
      const orderCol = 2; // B column for order number
      for (let i = 2; i <= lastRow; i++) {
        const val = ordersSheet.getRange(i, orderCol).getValue();
        if (String(val).trim() === "SCOL03") {
          scol03Row = i;
          break;
        }
      }
    }

    if (scol03Row > 0) {
      console.log("Starting work on SCOL03 at row", scol03Row, "with worker Admin");
      const startResult = await callShopFunction("startOrder", [scol03Row, "Admin", "Assembly"]);
      console.log("Started work on SCOL03:", startResult);
    } else {
      console.log("Could not find SCOL03 in orders sheet. Total rows:", lastRow);
    }
  } catch (e) {
    console.error("Error starting work:", e && e.message || e);
  }
}

main().catch(e => {
  console.error(e && e.stack || e);
  process.exit(1);
});
