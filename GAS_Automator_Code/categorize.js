function standardizeEmailList(emailString) {
  if (!emailString) {
    return "";
  }

  let str = String(emailString).trim();

  // Defensive Check: If string includes '不需變更' (no change) or '不用改' (don't change), return original string.
  if (str.toLowerCase().includes('不需變更') || str.toLowerCase().includes('不用改')) {
    return str;
  }

  // 1. Unify all common delimiters (semicolon, comma+space, simple space) to a single comma
  let cleanedString = str.replace(/[\s;,\uFEFF\xA0]+/g, ',');

  // 2. Consolidate all consecutive commas (e.g., ,, or , , ,) to a single comma
  let finalString = cleanedString.replace(/,{2,}/g, ',');

  // 3. Remove leading and trailing commas
  finalString = finalString.replace(/^,|,$/g, '');

  return finalString;
}

// UTILITY: Normalizes strings for case-insensitive and spacing-insensitive comparisons.
function norm(s) {
    if (s === undefined || s === null) return "";
    return String(s).replace(/\r/g, "").replace(/\n/g, " ").trim().replace(/\s+/g, " ").toLowerCase();
}

// UTILITY: Finds column index by header name (exact, contains, reverse contains).
function findHeaderIndex(headers, target) {
    const t = norm(target);
    // 1. Exact match
    for (let i = 0; i < headers.length; i++) if (norm(headers[i]) === t) return i;
    // 2. Contains match
    for (let i = 0; i < headers.length; i++) if (norm(headers[i]).indexOf(t) !== -1) return i;
    // 3. Reverse Contains match
    for (let i = 0; i < headers.length; i++) if (t.indexOf(norm(headers[i])) !== -1) return i;
    return -1;
}

// MAIN FUNCTION: Classifies change requests and distributes data to target input sheets.
function ChangesSorting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('分類工具');

  const INPUT_SHEET_NAME = 'TMS_input';
  const TAKEOUT_SHEET_NAME = 'FO_input'; // Constant name for Takeout/FO sheet
  let takeoutSheet = ss.getSheetByName(TAKEOUT_SHEET_NAME);

  // Check and retrieve or create target sheets
  let inputSheet = ss.getSheetByName(INPUT_SHEET_NAME);
  if (!inputSheet) {
    inputSheet = ss.insertSheet(INPUT_SHEET_NAME);
    Logger.log(`Sheet ${INPUT_SHEET_NAME} does not exist, created automatically.`);
  }
  if (!takeoutSheet) {
    takeoutSheet = ss.insertSheet(TAKEOUT_SHEET_NAME);
    Logger.log(`Sheet ${TAKEOUT_SHEET_NAME} does not exist, created automatically.`);
  }

  // ---- 1. Read all data (reading both raw values and display values) ----
  const dataRange = sourceSheet.getDataRange();
  const fullData = dataRange.getValues();
  const fullDisplayData = dataRange.getDisplayValues();

  if (fullData.length < 3) {
    // Clear targets if source data is insufficient, then return
    clearSheetData(inputSheet);
    clearSheetData(takeoutSheet);
    SpreadsheetApp.getUi().alert(`Source data is insufficient, target pages have been cleared.`);
    return;
  }

  const dataRows = fullData.slice(2);
  const displayRows = fullDisplayData.slice(2);

  const takeoutData = [];
  const combinedData = [];

  dataRows.forEach((row, index) => {
    const displayRow = displayRows[index];

    // Column Indices (0-based)
    const F_raw = row[5]; // Branch ID (may contain comma separator)
    const G_raw = row[6]; // Original Restaurant Name (may contain comma separator)
    const H = row[7];      // Service Change Request (used for Takeout/FO determination)

    // Values after changes
    const L = row[11]; // New Company Header/Title
    const M = row[12]; // New Company Unified Business No.
    const O = row[14]; // New Company Representative
    const Q = row[16]; // New Company Registered Address
    const T_raw = row[19]; // New Restaurant Name (may contain comma separator)
    const V = row[21]; // New Billing Contact Name
    const W_PHONE = displayRow[22]; // Phone Number (display value keeps leading zero)
    const X = row[23]; // New Billing Contact Email
    const Y = row[24]; // New Contract Shipping Address
    // Bank Info
    const AA = row[26];
    const AB = row[27];
    const AC = row[28];
    const AD = row[29];
    const AE = row[30];
    const AF = row[31];
    // Request Serial Number
    const C = row[2];

    if (!F_raw) return;

    const standardizedEmail = standardizeEmailList(X);

    // *** CORE: Synchronized splitting of Column F (ID), G (Old Name), T (New Name) ***
    const ids = String(F_raw).split(',').map(id => id.trim()).filter(id => id);
    const oldNames = String(G_raw || '').split(',').map(name => name.trim()).filter(name => name);
    const newNames = String(T_raw || '').split(',').map(name => name.trim()).filter(name => name);

    // Iterate through all split IDs
    ids.forEach((singleID, idx) => {
        // Extract corresponding Old/New names; if Column G data is insufficient, use blank
        const oldName = (idx < oldNames.length) ? oldNames[idx] : '';
        const newName = (idx < newNames.length) ? newNames[idx] : '';

        // --- Core Determination: Any changes requiring the TMS update process ---
        const hasTaxChange = L || M;
        const hasNameChange = newNames.length > 0; // Check if there are any new names
        const hasBillingChange = V || W_PHONE || X || Y;

        if (hasTaxChange || hasNameChange || hasBillingChange) {

            // TMS Data Update Input Structure: [A:ID, B:Old Name, C:Header, D:Tax ID, ...]
            combinedData.push([
                singleID,                                // Col A: Branch ID
                oldName,                                 // Col B: Original Restaurant Name (Old Name)
                String(L || '').trim(),                  // Col C: New Company Header/Title
                String(M || '').trim(),                  // Col D: New Company Unified Business No.
                newName,                                 // Col E: New Restaurant Name (New Name)
                String(V || '').trim(),                  // Col F: New Billing Contact Name
                String(W_PHONE || '').trim(),            // Col G: New Billing Contact Phone
                standardizedEmail,                       // Col H: New Billing Contact Email
                String(Y || '').trim(),                  // Col I: New Contract Shipping Address
                String(C || '').trim()                   // Col J: Request Serial Number
            ]);
        }

        // Takeout/FO Change Determination (Kept separate)
        const hasTakeoutChange = H && H.toString().split(',').map(s => s.trim()).includes('外帶外送系統');
        if (hasTakeoutChange) {
            // Structure: [ID, Old Name, Header, Tax ID, New Name, Representative, Address, Bank Info..., Contact Name, Phone, Email, Address, Serial No.]
            takeoutData.push([
                singleID,
                oldName,
                L, M, newName, O, Q, AA, AB, AC, AD, AE, AF, V, W_PHONE, standardizedEmail, Y, C
            ]);
        }
    });
  });

  // --------------------------------------------------------------------------------
  // *** Clear old data (starting from Row 3) ***
  // --------------------------------------------------------------------------------
  clearSheetData(takeoutSheet);
  clearSheetData(inputSheet);

  // *** Write Takeout/FO changes (kept separate) ***
  if (takeoutData.length) {
    const startRow = 3;
    takeoutSheet.getRange(startRow, 1, takeoutData.length, takeoutData[0].length)
      .setValues(takeoutData);
  }

  // *** Critical Write: Centralized write for TMS Update Input (Overwrite Mode) ***
  if (combinedData.length) {
    const startRow = 3;
    inputSheet.getRange(startRow, 1, combinedData.length, combinedData[0].length).setValues(combinedData);
  }

  // Helper: Write Header (assuming header starts at Row 2)
  const setupHeaders = (sheet, startCol, headers) => {
    sheet.getRange(2, startCol, 1, headers.length).setValues([headers]);
  };

  // Set Header for TMS Input
  setupHeaders(inputSheet, 1, [
    "Branch ID",
    "Original Restaurant Name (Old)", // Col B
    "New Company Header/Title",
    "New Company Unified Business No.",
    "New Restaurant Name (New)",
    "New Billing Contact Name",
    "New Billing Contact Phone",
    "New Billing Contact Email",
    "New Contract Shipping Address"
  ]);

  // Set Header for Takeout/FO Changes
  setupHeaders(takeoutSheet, 1, [
    "Branch ID",
    "Original Restaurant Name (Old)", // Col B
    "New Company Header/Title",
    "New Company Unified Business No.",
    "New Restaurant Name (New)",
    "New Company Representative",
    "New Company Registered Address",
    "New Takeout/FO Bank Name",
    "New Takeout/FO Bank Code",
    "New Takeout/FO Branch Name",
    "New Takeout/FO Branch Code",
    "New Takeout/FO Account Holder",
    "New Takeout/FO Account No.",
    "New Billing Contact Name",
    "New Billing Contact Phone",
    "New Billing Contact Email",
    "New Contract Shipping Address"
  ]);
  // --------------------------------------------------------------------------------

  SpreadsheetApp.getUi().alert(`資料分類與分發完成！`);
}

// *** Added Helper Function: Clears sheet data (starting from Row 3) ***
function clearSheetData(sheet) {
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    const startRow = 3; // Start clearing from Row 3

    if (maxRows > startRow - 1) {
        // Get range from Row 3 to the last column
        sheet.getRange(startRow, 1, maxRows - startRow + 1, maxCols).clearContent();
    }
}