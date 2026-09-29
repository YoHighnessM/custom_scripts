/**
 * Case Timestamp Sync — Google Apps Script Bridge
 * Version: 2.0.0
 * Description: Clean, modular, and robust backend bridge connecting Google Sheets
 *              with Tech24 Case Timestamp Sync Userscript.
 */

// =============================================================================
// 1. Configuration & Constants
// =============================================================================
const CONFIG = Object.freeze({
  SHEET_NAME: "Daily Activity Tracker",
  HEADER_ROW: 1,
  COLUMNS: Object.freeze({
    CASE_ID: 10,     // Column J
    REG_DATE: 15,    // Column O
    REG_TIME: 16,    // Column P
    CLOSED_DATE: 17, // Column Q
    CLOSED_TIME: 18, // Column R
  }),
  PROPERTIES: Object.freeze({
    SYNC_REQUESTED: "syncRequested",
    SYNC_STATUS: "syncStatus",
    LAST_RUN_SUMMARY: "lastRunSummary",
    LAST_RUN_TIME: "lastRunTime",
  })
});

// =============================================================================
// 2. Document Properties Manager
// =============================================================================
const PropertyManager = {
  get(key) {
    try {
      return PropertiesService.getDocumentProperties().getProperty(key);
    } catch (e) {
      Logger.log(`[PropertyManager.get Error] Key '${key}': ${e.message}`);
      return null;
    }
  },

  set(key, value) {
    try {
      PropertiesService.getDocumentProperties().setProperty(key, String(value));
    } catch (e) {
      Logger.log(`[PropertyManager.set Error] Key '${key}': ${e.message}`);
    }
  },

  setMany(propertiesObj) {
    try {
      PropertiesService.getDocumentProperties().setProperties(propertiesObj);
    } catch (e) {
      Logger.log(`[PropertyManager.setMany Error]: ${e.message}`);
    }
  }
};

// =============================================================================
// 3. Spreadsheet & Sheet Repository
// =============================================================================
const SheetRepository = {
  getSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error("No active spreadsheet found.");
    }
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      throw new Error(`Sheet "${CONFIG.SHEET_NAME}" not found.`);
    }
    return sheet;
  },

  getPendingRows() {
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) return [];

    const numRows = lastRow - CONFIG.HEADER_ROW;
    const maxCol = Math.max(...Object.values(CONFIG.COLUMNS));
    const data = sheet.getRange(CONFIG.HEADER_ROW + 1, 1, numRows, maxCol).getValues();

    const pending = [];
    data.forEach((rowValues, idx) => {
      const rowNum = CONFIG.HEADER_ROW + 1 + idx;
      const caseId = rowValues[CONFIG.COLUMNS.CASE_ID - 1];
      if (!caseId) return;

      const regDate = rowValues[CONFIG.COLUMNS.REG_DATE - 1];
      const closedDate = rowValues[CONFIG.COLUMNS.CLOSED_DATE - 1];

      const isRegEmpty = (regDate === "" || regDate === null || regDate === undefined);
      const isClosedEmpty = (closedDate === "" || closedDate === null || closedDate === undefined);

      // Queue if Start Date OR Closed Date is missing
      if (isRegEmpty || isClosedEmpty) {
        pending.push({
          row: rowNum,
          caseId: String(caseId).trim(),
        });
      }
    });

    return pending;
  },

  writeSingleResult(payload) {
    const row = Number(payload.row);
    if (!row || isNaN(row)) {
      throw new Error("Invalid or missing row parameter.");
    }

    const sheet = this.getSheet();
    const startCol = CONFIG.COLUMNS.REG_DATE;
    const range = sheet.getRange(row, startCol, 1, 4);
    const currentValues = range.getValues()[0];

    let modified = false;
    const isCellEmpty = (val) => (val === "" || val === null || val === undefined);

    // Only update empty cells to prevent overwriting existing valid timestamps
    if (isCellEmpty(currentValues[0]) && payload.regDate) {
      currentValues[0] = payload.regDate;
      modified = true;
    }
    if (isCellEmpty(currentValues[1]) && payload.regTime) {
      currentValues[1] = payload.regTime;
      modified = true;
    }
    if (isCellEmpty(currentValues[2]) && payload.closedDate) {
      currentValues[2] = payload.closedDate;
      modified = true;
    }
    if (isCellEmpty(currentValues[3]) && payload.closedTime) {
      currentValues[3] = payload.closedTime;
      modified = true;
    }

    if (modified) {
      range.setValues([currentValues]);
    }

    return { row, modified };
  }
};

// =============================================================================
// 4. Response Formatter
// =============================================================================
const ResponseHandler = {
  json(data) {
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  },

  error(message, details = null) {
    const payload = { error: message };
    if (details) payload.details = details;
    return this.json(payload);
  }
};

// =============================================================================
// 5. Menu UI & User Actions
// =============================================================================
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Actions")
      .addItem("Fill Timestamps", "requestSync")
      .addToUi();
  } catch (e) {
    Logger.log(`[onOpen Error]: ${e.message}`);
  }
}

function requestSync() {
  PropertyManager.setMany({
    [CONFIG.PROPERTIES.SYNC_REQUESTED]: "true",
    [CONFIG.PROPERTIES.SYNC_STATUS]: "pending",
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Sync requested! Switch to your Tech24 Dashboard tab. The sync will start automatically.",
    "Sync Ready",
    12
  );
}

// =============================================================================
// 6. HTTP Web App Controller
// =============================================================================
function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const action = params.action;

    switch (action) {
      case "ping":
        return ResponseHandler.json({ ok: true, ts: new Date().toISOString() });

      case "checkTrigger":
        return ResponseHandler.json({
          shouldRun: PropertyManager.get(CONFIG.PROPERTIES.SYNC_REQUESTED) === "true",
          status: PropertyManager.get(CONFIG.PROPERTIES.SYNC_STATUS) || "idle"
        });

      case "getPendingCases":
        PropertyManager.set(CONFIG.PROPERTIES.SYNC_STATUS, "running");
        return ResponseHandler.json({ rows: SheetRepository.getPendingRows() });

      case "submitResult":
        const result = SheetRepository.writeSingleResult({
          row: params.row,
          regDate: params.regDate || "",
          regTime: params.regTime || "",
          closedDate: params.closedDate || "",
          closedTime: params.closedTime || "",
        });
        return ResponseHandler.json({ ok: true, ...result });

      case "clearTrigger":
        PropertyManager.setMany({
          [CONFIG.PROPERTIES.SYNC_REQUESTED]: "false",
          [CONFIG.PROPERTIES.SYNC_STATUS]: "done",
          [CONFIG.PROPERTIES.LAST_RUN_SUMMARY]: params.summary || "{}",
          [CONFIG.PROPERTIES.LAST_RUN_TIME]: new Date().toISOString()
        });
        return ResponseHandler.json({ ok: true });

      default:
        return ResponseHandler.error("Unknown action requested.", { action });
    }

  } catch (err) {
    Logger.log(`[doGet Critical Error]: ${err.stack || err.message}`);
    return ResponseHandler.error(String(err.message || err));
  }
}
