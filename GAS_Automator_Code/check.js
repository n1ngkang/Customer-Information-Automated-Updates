const PROPS = PropertiesService.getScriptProperties();

const CONFIG = {
  // File IDs (Read securely from GAS Properties)
  TOOL_FILE_ID: PROPS.getProperty('TOOL_FILE_ID'), // Master Tool File ID
  DATA_FILE_ID: PROPS.getProperty('DATA_FILE_ID'), // RD_Response Datasheet ID

  // Sheet Names
  TMS_INPUT_NAME: 'TMS_input',
  FO_INPUT_NAME: 'FO_input',
  MASTER_SHEET_NAME: '分類工具', // Classification Tool Sheet Name
  DATASHEET_NAME: 'RD_Response',

  // TMS_input Column Indices (1-based)
  TMS_RESTAURANT_ID_COL: 1,    // Column A: Restaurant ID
  TMS_SERIAL_COL: 10,          // Column J: Serial Number
  TMS_CHECKBOX_COL: 16,        // Column P: Checkbox
  TMS_D_COL: 4,                // Column D: Has Tax ID Change
  TMS_K_COL: 11,               // Column K: Corresponding Checkbox (Review Check)

  // FO_input Column Indices (1-based)
  FO_RESTAURANT_ID_COL: 1,     // Column A: Restaurant ID
  FO_SERIAL_COL: 18,           // Column R: Serial Number
  FO_CHECKBOX_COL: 22,         // Column V: Checkbox
  FO_CHECK_START_COL: 8,       // Column H: Start of Bank/Billing Change Check
  FO_CHECK_END_COL: 13,        // Column M: End of Bank/Billing Change Check

  // Master Sheet (Classification Tool) Column Indices (1-based)
  MASTER_RESTAURANT_ID_COL: 1, // Column A: Restaurant ID
  MASTER_SERIAL_COL: 3,        // Column C: Serial Number
  MASTER_SUMMARY_COL: 37,      // Column AK: Summary Checkbox
  MASTER_ALERT_COL: 38,        // Column AL: Pending Contract/Alert Check

  // RD_Response Column Indices (1-based)
  DATASHEET_SERIAL_COL: 3,     // Column C: Serial Number (for matching)
  DATASHEET_AK_COL: 37,        // Column AK: Completion Checkbox
  DATASHEET_AL_COL: 38,        // Column AL: Pending Contract Check

  // Keyword Settings
  SKIP_KEYWORDS: ['不需變更', ''] // 'No Change' keywords (kept for functional accuracy)
};

// ==========================================
// UTILITY FUNCTION: Check for Valid Content
// ==========================================
function isValidContent(value) {
  if (!value) return false;
  const strValue = String(value).trim();
  if (strValue === '') return false;

  // Check against 'No Change' keywords
  for (let keyword of CONFIG.SKIP_KEYWORDS) {
    if (keyword && strValue.includes(keyword)) {
      return false;
    }
  }
  return true;
}

// ==========================================
// Step 1: Collect All Source Data (TMS_input & FO_input)
// ==========================================
function collectAllSourceData() {
  // Use securely retrieved IDs
  const toolFile = SpreadsheetApp.openById(CONFIG.TOOL_FILE_ID);
  const tmsSheet = toolFile.getSheetByName(CONFIG.TMS_INPUT_NAME);
  const foSheet = toolFile.getSheetByName(CONFIG.FO_INPUT_NAME);

  if (!tmsSheet || !foSheet) {
    Logger.log('Error: Could not find source sheets (TMS_input or FO_input).');
    return null;
  }

  // Data structure: { serialNum: { restaurantId: { checkboxes: [], alerts: [], sources: [...] } } }
  const dataBySerial = {};

  // === Read TMS_input ===
  const tmsLastRow = tmsSheet.getLastRow();
  if (tmsLastRow >= 2) {
    const tmsRestaurantIds = tmsSheet.getRange(2, CONFIG.TMS_RESTAURANT_ID_COL, tmsLastRow - 1).getValues();
    const tmsSerialNums = tmsSheet.getRange(2, CONFIG.TMS_SERIAL_COL, tmsLastRow - 1).getValues();
    const tmsPCheckboxes = tmsSheet.getRange(2, CONFIG.TMS_CHECKBOX_COL, tmsLastRow - 1).getValues();
    const tmsDValues = tmsSheet.getRange(2, CONFIG.TMS_D_COL, tmsLastRow - 1).getValues();
    const tmsKCheckboxes = tmsSheet.getRange(2, CONFIG.TMS_K_COL, tmsLastRow - 1).getValues();

    for (let i = 0; i < tmsRestaurantIds.length; i++) {
      const restaurantId = tmsRestaurantIds[i][0];
      const serialNum = tmsSerialNums[i][0];
      const pChecked = tmsPCheckboxes[i][0] === true;
      const dValue = tmsDValues[i][0];
      const kChecked = tmsKCheckboxes[i][0] === true;

      if (!restaurantId || !serialNum) continue;

      // Check Anomaly: D-column has content but K-column is unchecked (Has Tax ID Change but not reviewed)
      const hasAlert = isValidContent(dValue) && !kChecked;

      // Initialize structure
      if (!dataBySerial[serialNum]) {
        dataBySerial[serialNum] = {};
      }
      if (!dataBySerial[serialNum][restaurantId]) {
        dataBySerial[serialNum][restaurantId] = {
          checkboxes: [],
          alerts: [],
          sources: []
        };
      }

      // Record data
      dataBySerial[serialNum][restaurantId].checkboxes.push(pChecked);
      dataBySerial[serialNum][restaurantId].alerts.push(hasAlert);
      dataBySerial[serialNum][restaurantId].sources.push('TMS_input');
    }

    Logger.log(`TMS_input: Read ${tmsRestaurantIds.length} records.`);
  }

  // === Read FO_input ===
  const foLastRow = foSheet.getLastRow();
  if (foLastRow >= 2) {
    const foRestaurantIds = foSheet.getRange(2, CONFIG.FO_RESTAURANT_ID_COL, foLastRow - 1).getValues();
    const foSerialNums = foSheet.getRange(2, CONFIG.FO_SERIAL_COL, foLastRow - 1).getValues();
    const foVCheckboxes = foSheet.getRange(2, CONFIG.FO_CHECKBOX_COL, foLastRow - 1).getValues();

    // Read Columns H-M
    const checkCols = CONFIG.FO_CHECK_END_COL - CONFIG.FO_CHECK_START_COL + 1;
    const foCheckValues = foSheet.getRange(2, CONFIG.FO_CHECK_START_COL, foLastRow - 1, checkCols).getValues();

    for (let i = 0; i < foRestaurantIds.length; i++) {
      const restaurantId = foRestaurantIds[i][0];
      const serialNum = foSerialNums[i][0];
      const vChecked = foVCheckboxes[i][0] === true;

      if (!restaurantId || !serialNum) continue;

      // Check if any column from H-M has valid content (Has Bank/Billing Change)
      const hasValidContent = foCheckValues[i].some(value => isValidContent(value));

      // Check Anomaly: H-M has content but V-column is unchecked
      const hasAlert = hasValidContent && !vChecked;

      // Initialize structure
      if (!dataBySerial[serialNum]) {
        dataBySerial[serialNum] = {};
      }
      if (!dataBySerial[serialNum][restaurantId]) {
        dataBySerial[serialNum][restaurantId] = {
          checkboxes: [],
          alerts: [],
          sources: []
        };
      }

      // Record data
      dataBySerial[serialNum][restaurantId].checkboxes.push(vChecked);
      dataBySerial[serialNum][restaurantId].alerts.push(hasAlert);
      dataBySerial[serialNum][restaurantId].sources.push('FO_input');
    }

    Logger.log(`FO_input: Read ${foRestaurantIds.length} records.`);
  }

  return dataBySerial;
}

// ==========================================
// Step 2: Calculate Summary Status per Serial Number
// ==========================================
function calculateSerialStatus(dataBySerial) {
  // Structure: { serialNum: { summaryCheckbox: bool, hasAlert: bool } }
  const serialStatus = {};

  Object.keys(dataBySerial).forEach(serialNum => {
    const restaurants = dataBySerial[serialNum];

    // Check if all sources for all restaurants under this serial number are checked
    let allChecked = true;
    let anyAlert = false;

    Object.keys(restaurants).forEach(restaurantId => {
      const data = restaurants[restaurantId];

      // All sources for this specific restaurant must be checked
      const restaurantAllChecked = data.checkboxes.every(checked => checked === true);
      if (!restaurantAllChecked) {
        allChecked = false;
      }

      // Mark if any anomaly exists
      const restaurantHasAlert = data.alerts.some(alert => alert === true);
      if (restaurantHasAlert) {
        anyAlert = true;
      }
    });

    serialStatus[serialNum] = {
      summaryCheckbox: allChecked,
      hasAlert: anyAlert
    };
  });

  Logger.log(`Calculation completed: Status for ${Object.keys(serialStatus).length} serial numbers.`);
  return serialStatus;
}

// ==========================================
// Step 3: Update Master Sheet (AK/AL Columns)
// ==========================================
function updateMasterSheet(serialStatus) {
  const toolFile = SpreadsheetApp.openById(CONFIG.TOOL_FILE_ID);
  const masterSheet = toolFile.getSheetByName(CONFIG.MASTER_SHEET_NAME);

  if (!masterSheet) {
    Logger.log('Error: Could not find Master Sheet (Classification Tool).');
    return;
  }

  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) return;

  // Read Master Sheet serial numbers (Column C)
  const masterSerials = masterSheet.getRange(2, CONFIG.MASTER_SERIAL_COL, lastRow - 1).getValues();

  // Collect cells needing updates
  const akUpdates = []; // Summary Checkbox updates (Column AK)
  const alUpdates = []; // Alert/Contract Check updates (Column AL)

  masterSerials.forEach((row, idx) => {
    const serialNum = row[0];
    const status = serialStatus[serialNum];
    const targetRow = idx + 2;

    if (status) {
      // Column AK: Summary Checkbox
      akUpdates.push({
        range: masterSheet.getRange(targetRow, CONFIG.MASTER_SUMMARY_COL),
        value: status.summaryCheckbox
      });

      // Column AL: Anomaly Alert
      alUpdates.push({
        range: masterSheet.getRange(targetRow, CONFIG.MASTER_ALERT_COL),
        value: status.hasAlert
      });
    } else {
      // No corresponding data found, clear/set to false
      akUpdates.push({
        range: masterSheet.getRange(targetRow, CONFIG.MASTER_SUMMARY_COL),
        value: false
      });
      alUpdates.push({
        range: masterSheet.getRange(targetRow, CONFIG.MASTER_ALERT_COL),
        value: false
      });
    }
  });

  // Batch update
  akUpdates.forEach(update => update.range.setValue(update.value));
  alUpdates.forEach(update => update.range.setValue(update.value));

  Logger.log(`Master Sheet: Updated ${akUpdates.length} AK cells and ${alUpdates.length} AL cells.`);
}

// ==========================================
// Step 4: Synchronize Status to RD_Response Datasheet
// ==========================================
function syncToDatasheet(serialStatus) {
  // Use securely retrieved ID
  const datasheet = SpreadsheetApp.openById(CONFIG.DATA_FILE_ID)
                                 .getSheetByName(CONFIG.DATASHEET_NAME);

  if (!datasheet) {
    Logger.log('Error: Could not find RD_Response datasheet.');
    return;
  }

  const lastRow = datasheet.getLastRow();
  if (lastRow < 2) return;

  // Read RD_Response serial numbers (Column C)
  const dataSerials = datasheet.getRange(2, CONFIG.DATASHEET_SERIAL_COL, lastRow - 1).getValues();

  // Collect cells needing updates
  const akUpdates = []; // Completion Checkbox updates (Column AK)
  const alUpdates = []; // Pending Contract Check updates (Column AL)

  dataSerials.forEach((row, idx) => {
    const serialNum = row[0];
    const status = serialStatus[serialNum];
    const targetRow = idx + 2;

    if (status) {
      // Column AK: Summary Checkbox
      if (status.summaryCheckbox) {
        akUpdates.push({
          range: datasheet.getRange(targetRow, CONFIG.DATASHEET_AK_COL),
          value: true
        });
      }

      // Column AL: Anomaly Alert
      if (status.hasAlert) {
        alUpdates.push({
          range: datasheet.getRange(targetRow, CONFIG.DATASHEET_AL_COL),
          value: true
        });
      }
    }
  });

  // Batch update
  akUpdates.forEach(update => update.range.setValue(update.value));
  alUpdates.forEach(update => update.range.setValue(update.value));

  Logger.log(`RD_Response: Updated ${akUpdates.length} AK cells and ${alUpdates.length} AL cells.`);
}

// ==========================================
// MAIN EXECUTION FUNCTION: Full Workflow
// ==========================================
function executeFullSync() {
  Logger.log('=== Starting Full Synchronization Workflow ===');

  // Step 1: Collect data from all source sheets
  Logger.log('Step 1: Collecting data from TMS_input and FO_input.');
  const dataBySerial = collectAllSourceData();

  if (!dataBySerial) {
    Logger.log('Could not read source data, workflow aborted.');
    return;
  }

  // Step 2: Calculate summary status for each serial number
  Logger.log('Step 2: Calculating summary status per serial number.');
  const serialStatus = calculateSerialStatus(dataBySerial);

  // Step 3: Update Master Sheet
  Logger.log('Step 3: Updating Master Sheet (AK/AL columns).');
  updateMasterSheet(serialStatus);

  // Step 4: Synchronize to RD_Response
  Logger.log('Step 4: Synchronizing status to RD_Response.');
  syncToDatasheet(serialStatus);

  Logger.log('=== Full Synchronization Workflow Completed ===');
}