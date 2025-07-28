/**
 * Gets unique department values from the Projects sheet for the dropdown filter.
 * @returns {Array<string>} Array of unique department names.
 */
function getUniqueDepartments() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Projects");
    if (!sheet) {
      console.error("Projects sheet not found");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return []; // No data or only headers
    }

    const headers = data[0];
    const departmentColIndex = headers.findIndex(header => 
      String(header || '').trim().toLowerCase() === 'department'
    );

    if (departmentColIndex === -1) {
      console.warn("Department column not found in Projects sheet");
      return [];
    }

    // Get unique department values (skip header row)
    const departments = new Set();
    for (let i = 1; i < data.length; i++) {
      const deptValue = String(data[i][departmentColIndex] || '').trim();
      if (deptValue && deptValue !== '') {
        departments.add(deptValue);
      }
    }

    // Convert Set to Array and sort alphabetically
    const uniqueDepartments = Array.from(departments).sort();
    console.log("Found unique departments:", uniqueDepartments);
    return uniqueDepartments;

  } catch (error) {
    console.error("Error getting unique departments:", error);
    return [];
  }
}

/**
 * Gets all data from a specified sheet.
 * @param {string} sheetName - The name of the sheet to get data from.
 * @returns {Object} Object containing headers and data arrays.
 */
function getSheetData(sheetName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      console.error(`Sheet "${sheetName}" not found`);
      return { headers: [], data: [] };
    }

    const data = sheet.getDataRange().getValues();
    if (data.length === 0) {
      return { headers: [], data: [] };
    }

    const headers = data[0];
    const rows = data.slice(1); // Skip header row

    console.log(`Retrieved ${rows.length} rows from sheet "${sheetName}"`);
    return {
      headers: headers,
      data: rows
    };

  } catch (error) {
    console.error(`Error getting sheet data for "${sheetName}":`, error);
    return { headers: [], data: [] };
  }
}