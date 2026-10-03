/**
 * Case Timestamp Sync & Duration Calculation — Google Apps Script Bridge
 * Version: 2.1.0
 * Description: Modular backend bridge connecting Google Sheets
 *              with Tech24 Case Timestamp Sync Userscript & Duration Calculation.
 */

// =============================================================================
// 1. Configuration & Constants
// =============================================================================
const CONFIG = Object.freeze({
  SHEET_NAME: "Daily Activity Tracker",
  HEADER_ROW: 1,
  COLUMNS: Object.freeze({
    MACHINE_ID: 3,   // Column C
    CASE_ID: 10,     // Column J
    REG_DATE: 15,    // Column O
    REG_TIME: 16,    // Column P
    CLOSED_DATE: 17, // Column Q
    CLOSED_TIME: 18, // Column R
    DURATION: 19,    // Column S
  }),
  PROPERTIES: Object.freeze({
    SYNC_REQUESTED: "syncRequested",
    SYNC_STATUS: "syncStatus",
    LAST_RUN_SUMMARY: "lastRunSummary",
    LAST_RUN_TIME: "lastRunTime",
  })
});

const DURATION_CONFIG = Object.freeze({
  MASTER_DB_ID: "10MXO_TBs00rG1RutyA9sjrmygse8J-WbWrNL011vjRc",
  SHEET_NAME: "DB_Machines",
  MACHINE_COL_SHEET: 2, // Column B
  ZONE_COL_SHEET: 8,    // Column H
  START_ROW: 4,
  LIGHT_RED_COLOR: "#fce8e6",
  BUSINESS_START_HOUR: 8,  // 8:00 AM
  BUSINESS_END_HOUR: 17,   // 5:00 PM
  BUSINESS_DAYS: Object.freeze([1, 2, 3, 4, 5, 6]), // Mon-Sat, Sunday(0) excluded
  COLUMNS: Object.freeze({
    MACHINE_ID: 3,   // Column C
    REG_DATE: 15,    // Column O
    REG_TIME: 16,    // Column P
    CLOSED_DATE: 17, // Column Q
    CLOSED_TIME: 18, // Column R
    DURATION: 19,    // Column S
  }),
  ZONE_HOURS: Object.freeze({
    1: 7,
    2: 8,
    3: 14,
  }),
});

const BusinessCalendar = {
  isBusinessDay(date) {
    return DURATION_CONFIG.BUSINESS_DAYS.indexOf(date.getDay()) !== -1;
  },

  nextBusinessDayStart(date) {
    const d = new Date(date);
    d.setDate(d.getDate() + 1);
    d.setHours(DURATION_CONFIG.BUSINESS_START_HOUR, 0, 0, 0);
    while (!this.isBusinessDay(d)) {
      d.setDate(d.getDate() + 1);
    }
    return d;
  },

  clampToBusinessWindow(date) {
    const d = new Date(date);
    if (this.isBusinessDay(d)) {
      const startOfDay = new Date(d);
      startOfDay.setHours(DURATION_CONFIG.BUSINESS_START_HOUR, 0, 0, 0);
      const endOfDay = new Date(d);
      endOfDay.setHours(DURATION_CONFIG.BUSINESS_END_HOUR, 0, 0, 0);

      if (d < startOfDay) return startOfDay;
      if (d >= endOfDay) return this.nextBusinessDayStart(d);
      return d;
    }
    return this.nextBusinessDayStart(d);
  },

  addBusinessHours(start, hours) {
    let remainingMinutes = hours * 60;
    let cursor = this.clampToBusinessWindow(start);

    while (remainingMinutes > 0) {
      const endOfDay = new Date(cursor);
      endOfDay.setHours(DURATION_CONFIG.BUSINESS_END_HOUR, 0, 0, 0);

      const minutesLeftToday = (endOfDay.getTime() - cursor.getTime()) / 60000;

      if (minutesLeftToday >= remainingMinutes) {
        cursor = new Date(cursor.getTime() + remainingMinutes * 60000);
        remainingMinutes = 0;
      } else {
        remainingMinutes -= minutesLeftToday;
        cursor = this.nextBusinessDayStart(cursor);
      }
    }
    return cursor;
  },

  businessMinutesBetween(from, to) {
    let cursor = new Date(from);
    let total = 0;

    while (cursor < to) {
      if (!this.isBusinessDay(cursor)) {
        cursor = this.nextBusinessDayStart(cursor);
        continue;
      }
      const startOfDay = new Date(cursor);
      startOfDay.setHours(DURATION_CONFIG.BUSINESS_START_HOUR, 0, 0, 0);
      const endOfDay = new Date(cursor);
      endOfDay.setHours(DURATION_CONFIG.BUSINESS_END_HOUR, 0, 0, 0);

      if (cursor < startOfDay) {
        cursor = startOfDay;
        continue;
      }
      if (cursor >= endOfDay) {
        cursor = this.nextBusinessDayStart(cursor);
        continue;
      }

      const segmentEnd = to < endOfDay ? to : endOfDay;
      total += (segmentEnd.getTime() - cursor.getTime()) / 60000;
      cursor = segmentEnd >= endOfDay ? this.nextBusinessDayStart(segmentEnd) : segmentEnd;
    }
    return total;
  },

  calculateDurationMinutes(startDate, endDate) {
    if (startDate.getTime() <= endDate.getTime()) {
      return this.businessMinutesBetween(startDate, endDate);
    } else {
      return -this.businessMinutesBetween(endDate, startDate);
    }
  }
};

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
// 3. Duration & SLA Target Calculation Engine
// =============================================================================
let machineZoneCache_ = null;

function getMachineZoneMap() {
  if (machineZoneCache_) {
    return machineZoneCache_;
  }
  const map = {};
  try {
    const masterDb = SpreadsheetApp.openById(DURATION_CONFIG.MASTER_DB_ID);
    const sheet = masterDb.getSheetByName(DURATION_CONFIG.SHEET_NAME);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= DURATION_CONFIG.START_ROW) {
        const numRows = lastRow - DURATION_CONFIG.START_ROW + 1;
        const data = sheet
          .getRange(
            DURATION_CONFIG.START_ROW,
            1,
            numRows,
            DURATION_CONFIG.ZONE_COL_SHEET
          )
          .getValues();
        data.forEach((row) => {
          const mId = String(row[DURATION_CONFIG.MACHINE_COL_SHEET - 1] || "")
            .trim()
            .toLowerCase();
          const zoneVal = row[DURATION_CONFIG.ZONE_COL_SHEET - 1];
          if (mId) {
            const zNum = parseInt(zoneVal, 10);
            map[mId] = isNaN(zNum) ? 1 : zNum;
          }
        });
      }
    }
  } catch (e) {
    Logger.log(`[getMachineZoneMap Error]: ${e.message}`);
  }
  machineZoneCache_ = map;
  return map;
}

function getZoneForMachineId(machineId) {
  if (!machineId) return 1;
  const key = String(machineId).trim().toLowerCase();
  const map = getMachineZoneMap();
  return map[key] !== undefined ? map[key] : 1;
}

function parseDateTime(dateVal, timeVal) {
  if (dateVal === null || dateVal === undefined || dateVal === "") return null;

  let year, month, day;

  if (dateVal instanceof Date) {
    year = dateVal.getFullYear();
    month = dateVal.getMonth();
    day = dateVal.getDate();
  } else {
    const str = String(dateVal).trim();
    if (!str || str === "-") return null;

    let m = str.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
    if (m) {
      year = parseInt(m[1], 10);
      month = parseInt(m[2], 10) - 1;
      day = parseInt(m[3], 10);
    } else {
      const digits = str.match(/(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/);
      if (digits) {
        let p1 = parseInt(digits[1], 10);
        let p2 = parseInt(digits[2], 10);
        let p3 = parseInt(digits[3], 10);
        if (p3 < 100) p3 += 2000;

        if (p1 > 12) {
          day = p1;
          month = p2 - 1;
          year = p3;
        } else if (p2 > 12) {
          month = p1 - 1;
          day = p2;
          year = p3;
        } else {
          // Default: DD-MM-YYYY format
          day = p1;
          month = p2 - 1;
          year = p3;
        }
      } else {
        const d = new Date(str);
        if (isNaN(d.getTime())) return null;
        year = d.getFullYear();
        month = d.getMonth();
        day = d.getDate();
      }
    }
  }

  let hours = 0;
  let minutes = 0;

  if (timeVal !== null && timeVal !== undefined && timeVal !== "") {
    if (timeVal instanceof Date) {
      hours = timeVal.getHours();
      minutes = timeVal.getMinutes();
    } else {
      const tStr = String(timeVal).trim();
      const tm = tStr.match(/(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?/i);
      if (tm) {
        hours = parseInt(tm[1], 10);
        minutes = parseInt(tm[2], 10);
        const ampm = tm[3];
        if (ampm) {
          const isPM = ampm.toUpperCase() === "PM";
          if (isPM && hours !== 12) hours += 12;
          if (!isPM && hours === 12) hours = 0;
        }
      }
    }
  }

  return new Date(year, month, day, hours, minutes, 0);
}

function formatTargetDateTime(date) {
  const pad = (n) => String(n).padStart(2, "0");
  const day = pad(date.getDate());
  const month = pad(date.getMonth() + 1);
  const year = date.getFullYear();

  let hours = date.getHours();
  const ampm = hours >= 12 ? "PM" : "AM";
  hours = hours % 12;
  if (hours === 0) hours = 12;

  return `${day}-${month}-${year} ${pad(hours)}:${pad(date.getMinutes())} ${ampm}`;
}

function formatDurationMs(diffMs) {
  const totalMinutes = Math.round(diffMs / (1000 * 60));
  const sign = totalMinutes < 0 ? "-" : "";
  const abs = Math.abs(totalMinutes);
  const h = Math.floor(abs / 60);
  const m = abs % 60;
  return `${sign}${h}h ${m}m`;
}

function calculateRowDurationData(machineId, regDate, regTime, closedDate, closedTime) {
  const startDate = parseDateTime(regDate, regTime);
  const endDate = parseDateTime(closedDate, closedTime);

  if (!startDate || !endDate) {
    return null;
  }

  const zone = getZoneForMachineId(machineId);
  const targetHours = DURATION_CONFIG.ZONE_HOURS[zone] || 7;

  const bizDurationMinutes = BusinessCalendar.calculateDurationMinutes(startDate, endDate);
  const durationStr = formatDurationMs(bizDurationMinutes * 60 * 1000);

  const targetDate = BusinessCalendar.addBusinessHours(startDate, targetHours);
  const targetTimeStr = formatTargetDateTime(targetDate);

  const outputString = `${durationStr}  •  Z${zone}  •  Tgt: ${targetTimeStr}`;
  const isOverdue = endDate.getTime() > targetDate.getTime();

  return {
    outputString,
    isOverdue,
    zone,
    durationStr,
    targetTimeStr,
    startDate,
    endDate,
    targetDate,
  };
}

function updateDurationForRow(sheet, rowNum) {
  if (rowNum < 2) return;

  const range = sheet.getRange(rowNum, 1, 1, DURATION_CONFIG.COLUMNS.DURATION);
  const values = range.getValues()[0];

  const machineId = values[DURATION_CONFIG.COLUMNS.MACHINE_ID - 1];
  const regDate = values[DURATION_CONFIG.COLUMNS.REG_DATE - 1];
  const regTime = values[DURATION_CONFIG.COLUMNS.REG_TIME - 1];
  const closedDate = values[DURATION_CONFIG.COLUMNS.CLOSED_DATE - 1];
  const closedTime = values[DURATION_CONFIG.COLUMNS.CLOSED_TIME - 1];

  const resultCell = sheet.getRange(rowNum, DURATION_CONFIG.COLUMNS.DURATION);

  const isFilled = (val) => val !== null && val !== undefined && String(val).trim() !== "" && String(val).trim() !== "-";

  // Check if all 4 timestamp cells are filled
  if (!isFilled(regDate) || !isFilled(regTime) || !isFilled(closedDate) || !isFilled(closedTime)) {
    resultCell.clearContent();
    resultCell.setBackground(null);
    return;
  }

  const calcData = calculateRowDurationData(
    machineId,
    regDate,
    regTime,
    closedDate,
    closedTime
  );

  if (!calcData) {
    resultCell.clearContent();
    resultCell.setBackground(null);
    return;
  }

  resultCell.setValue(calcData.outputString);
  if (calcData.isOverdue) {
    resultCell.setBackground(DURATION_CONFIG.LIGHT_RED_COLOR);
  } else {
    resultCell.setBackground(null);
  }
}

function updateAllDurations(sheet) {
  const targetSheet = sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = targetSheet.getLastRow();
  if (lastRow < 2) return;

  for (let r = 2; r <= lastRow; r++) {
    updateDurationForRow(targetSheet, r);
  }
}

// =============================================================================
// 4. Spreadsheet & Sheet Repository
// =============================================================================
const SheetRepository = {
  getSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error("No active spreadsheet found.");
    }
    const sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME) || spreadsheet.getSheets()[0];
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
      // Trigger duration calculation on row write
      updateDurationForRow(sheet, row);
    }

    return { row, modified };
  }
};

// =============================================================================
// 5. Response Formatter
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
// 6. Menu UI & User Actions
// =============================================================================
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Actions")
      .addItem("Fill Timestamps", "requestSync")
      .addItem("Calculate Durations", "updateAllDurations")
      .addToUi();
  } catch (e) {
    Logger.log(`[onOpen Error]: ${e.message}`);
  }
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    if (row >= 2) {
      updateDurationForRow(sheet, row);
    }
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
// 7. HTTP Web App Controller
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
