/**
 * Tech24 Activity Tracker — Integrated Weekly Report & Case Timestamp Sync
 * Version: 2.7.0
 * Description: Fully integrated, modular Google Apps Script for Google Sheets.
 *              Combines automated Weekly Report Generation (Google Docs),
 *              Case Timestamp Synchronization Bridge for Tampermonkey Userscripts,
 *              and Automated Duration & SLA Target Calculation.
 */

// =============================================================================
// 1. Configuration & Constants
// =============================================================================
const TIMESTAMP_SYNC_CONFIG = Object.freeze({
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
  }),
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

const REPORT_CONFIG = Object.freeze({
  SHEET_NAME: "Activity Tracker",
  DOC_ID: "1kYbTodcPCiA5ngbQrNF6qwPGGxf5j_Fd1U8IqKO9bE8",
  TABLE_BACKGROUND: "#ffffff",
  TEXT_COLOR: "#000000",
  COVER: {
    title: "Weekly Report",
    district: "Bishoftu, Central Addis, Hawassa, South Addis & Wolayta",
    company: "Tech 24 Trading One Member PLC",
    overviewTitle: "Overview",
    overviewIntro:
      "This report provides a concise summary of key activities and accomplishments across all districts and banks for the week. It highlights:",
    overviewClosing:
      "This overview offers a clear snapshot of the week's performance, with detailed reports provided in the following pages.",
    overviewItems: [
      {
        label: "Case Management",
        value: "Registered, resolved, and pending cases.",
      },
      {
        label: "Support Activities",
        value:
          "On-site and remote assistance provided to ensure smooth operations.",
      },
      {
        label: "Spare Parts",
        value: "Replacements, Returned and Pending spare Data.",
      },
      {
        label: "Preventive Maintenance (PM)",
        value: "Progress and completed maintenance Tasks across locations.",
      },
      {
        label: "Operational Updates",
        value: "Weekly tasks, meeting highlights, and technician performance.",
      },
    ],
  },
  SIGN_OFF_NAMES: [
    "Dawit Michael (District Manager)",
    "Lelisa Dessalen (South Addis, Bishoftu, Wolayta Team Leader)",
    "Yohannes Mulu (Central Addis, Hawassa, Shashemene Team Leader)",
  ],
  TABLE_SOURCES: Object.freeze({
    DISTRICT_SUMMARY: {
      sheetName: "District Summary",
      tableName: "District Summary",
    },
    CASE_SUMMARY_BY_DISTRICT: {
      sheetName: "Case Summary",
      tableName: "Case Counts By District",
    },
    LAST_WEEK_CASE_AMOUNT: {
      sheetName: "Last Week Case Amount",
      tableName: "Last week case amount",
    },
    CASES_CLOSED_AFTER_REGISTRATION_DATE: {
      sheetName: "Post-Reg & Ongoing Cases",
      tableName: "Post Reg Cases",
    },
    ONGOING_CASES: {
      sheetName: "Post-Reg & Ongoing Cases",
      tableName: "Ongoing Cases",
    },
    BANK_SPECIFIC_CASE_COUNT: {
      sheetName: "Case Summary",
      tableName: "Case Counts By Bank",
    },
    REGISTRATION_TYPE_COUNTS: {
      sheetName: "Case Summary",
      tableName: "Reg Types",
    },
    PM_SUMMARY: {
      sheetName: "PM Summary",
      tableName: "PM amount by district",
    },
    BANK_SPECIFIC_PM_COUNT: {
      sheetName: "PM Summary",
      tableName: "PM amount by bank",
    },
    WEEKLY_TASKS: {
      sheetName: "Tasks & Challenges",
      tableName: "Tasks",
    },
    WEEKLY_MEETING_OVERVIEW: {
      sheetName: "Meetings",
      tableName: "Meetings",
    },
    PER_DIEM_COST: {
      sheetName: "Per Diem Cost",
      tableName: "Per Diem Cost",
    },
    CHANGED_SPARE_PARTS: {
      sheetName: "Changed Spare Parts",
      tableName: "Changed Spare Parts",
    },
    TECHNICIANS_WEEKLY_ACTIVITY: {
      sheetName: "Technicians Activity Summary",
      tableName: "Technicians activity amount",
    },
    CHALLENGES: {
      sheetName: "Tasks & Challenges",
      tableName: "Challenges",
    },
  }),
});

// Helper function to calculate current report period (Saturday to Friday)
function getWeeklyReportPeriod(refDate = new Date()) {
  const date = new Date(refDate);
  const day = date.getDay(); // 0 = Sun, 1 = Mon, ..., 5 = Fri, 6 = Sat

  const satOffset = day === 6 ? 0 : -(day + 1);
  const saturday = new Date(date);
  saturday.setDate(date.getDate() + satOffset);

  const friday = new Date(saturday);
  friday.setDate(saturday.getDate() + 6);

  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  const pad = (n) => String(n).padStart(2, "0");
  const startStr = `${monthNames[saturday.getMonth()]} ${pad(saturday.getDate())}`;
  const endStr = `${monthNames[friday.getMonth()]} ${pad(friday.getDate())}, ${friday.getFullYear()}`;

  return `${startStr} to ${endStr}`;
}

// =============================================================================
// 2. Menu Entry Points & Triggers
// =============================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Actions")
    .addItem("Fill Timestamps", "requestSync")
    .addItem("Calculate Durations", "updateAllDurations")
    .addSeparator()
    .addItem("Generate Report", "generateWeeklyReport")
    .addItem("Archive and Reset", "archiveAndReset")
    .addToUi();
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

function generateWeeklyReport() {
  const ui = SpreadsheetApp.getUi();

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const doc = DocumentApp.openById(REPORT_CONFIG.DOC_ID);
    const builder = new WeeklyReportBuilder(spreadsheet, doc);
    builder.build();
    ui.alert("Weekly report generated successfully.");
  } catch (error) {
    ui.alert(`Failed to generate report: ${error.message}`);
    throw error;
  }
}

function generateWeeklyReportFeb25() {
  generateWeeklyReport();
}

function archiveAndReset() {
  const ui = SpreadsheetApp.getUi();
  const confirmation = ui.alert(
    "Archive and Reset",
    "This will archive a copy of this spreadsheet to 'Weekly file backup' and reset configured data ranges. Continue?",
    ui.ButtonSet.YES_NO
  );

  if (confirmation !== ui.Button.YES) {
    return;
  }

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Determine period string and backup file name
    let periodStr = "";
    const dataSetSheet = spreadsheet.getSheetByName("Data Set");
    if (dataSetSheet) {
      const val = dataSetSheet.getRange("I1").getDisplayValue();
      if (val && val.trim() !== "") {
        periodStr = val.trim();
      }
    }
    if (!periodStr) {
      periodStr = getWeeklyReportPeriod();
    }

    const backupFileName = `Activity Tracker | ${periodStr}`;

    // 2. Archive copy to folder 'Weekly file backup'
    const backupFolderId = "1OGVZscnmDDAsjNAXixlHrrQo2sS-qBQU";
    const backupFolder = DriveApp.getFolderById(backupFolderId);
    const currentFile = DriveApp.getFileById(spreadsheet.getId());
    currentFile.makeCopy(backupFileName, backupFolder);

    // 3. Save values from sheet 'Case Summary' range C2:G7 before resetting
    let caseSummaryValues = null;
    const caseSummarySheet = spreadsheet.getSheetByName("Case Summary");
    if (caseSummarySheet) {
      caseSummaryValues = caseSummarySheet.getRange("C2:G7").getValues();
    }

    // 4. PerformReset/Clear of specified ranges
    const tableMetadata = SpreadsheetRepository.getTableMetadata(spreadsheet);

    // - sheet (Daily Activity Tracker) range (A2:Y200), clear values and remove background highlighting from S2:S200
    const dailySheet = spreadsheet.getSheetByName("Daily Activity Tracker");
    if (dailySheet) {
      SpreadsheetRepository.clearRangeValuesOnly(spreadsheet, "Daily Activity Tracker", "A2:Y200");
      dailySheet.getRange("S2:S200").setBackground(null);
    }

    // - sheet (No Case Tracker) range (A2:E100)
    SpreadsheetRepository.clearRangeValuesOnly(spreadsheet, "No Case Tracker", "A2:E100");

    // - sheet (Post-Reg & Ongoing Cases): table (Post Reg Cases) (G2:G) and table (Ongoing Cases)
    try {
      const postRegTableRange = SpreadsheetRepository.getTableRange(
        spreadsheet,
        tableMetadata,
        SpreadsheetRepository.getTableSource("CASES_CLOSED_AFTER_REGISTRATION_DATE")
      );
      if (postRegTableRange && postRegTableRange.getNumRows() > 1) {
        const colGRange = postRegTableRange.offset(1, 6, postRegTableRange.getNumRows() - 1, 1);
        SpreadsheetRepository.clearRangeObjectValuesOnly(colGRange);
      }
    } catch (e) {
      SpreadsheetRepository.clearRangeValuesOnly(spreadsheet, "Post-Reg & Ongoing Cases", "G2:G");
    }
    SpreadsheetRepository.clearTableBodyValuesOnly(spreadsheet, tableMetadata, "ONGOING_CASES");

    // - sheet (PM Summary): table (PM amount by district) (D2:F7) -> col indices 3, 4, 5
    SpreadsheetRepository.clearTableBodyValuesOnly(spreadsheet, tableMetadata, "PM_SUMMARY", [3, 4, 5]);

    // - sheet (Tasks & Challenges): table (Tasks) and table (Challenges)
    SpreadsheetRepository.clearTableBodyValuesOnly(spreadsheet, tableMetadata, "WEEKLY_TASKS");
    SpreadsheetRepository.clearTableBodyValuesOnly(spreadsheet, tableMetadata, "CHALLENGES");

    // - sheet (Meetings): table (Meetings)
    SpreadsheetRepository.clearTableBodyValuesOnly(spreadsheet, tableMetadata, "WEEKLY_MEETING_OVERVIEW");

    // - sheet (Per Diem Cost): table (Per Diem Cost)
    SpreadsheetRepository.clearTableBodyValuesOnly(spreadsheet, tableMetadata, "PER_DIEM_COST");

    // - sheet (Changed Spare Parts): range or table name (H2:H20)
    const changedSpareSheet = spreadsheet.getSheetByName("Changed Spare Parts");
    if (changedSpareSheet) {
      SpreadsheetRepository.clearRangeValuesOnly(spreadsheet, "Changed Spare Parts", "H2:H20");
    }

    // - sheet (Last Week Case Amount): range (B2:F7)
    SpreadsheetRepository.clearRangeValuesOnly(spreadsheet, "Last Week Case Amount", "B2:F7");

    // 5. Populate (Last Week Case Amount) range B2:F7 with saved values from (Case Summary) C2:G7
    const lastWeekSheet = spreadsheet.getSheetByName("Last Week Case Amount");
    if (lastWeekSheet && caseSummaryValues) {
      lastWeekSheet.getRange("B2:F7").setValues(caseSummaryValues);
    }

    ui.alert(`Archive created successfully as "${backupFileName}" and data ranges have been reset.`);
  } catch (error) {
    ui.alert(`Failed to archive and reset: ${error.message}`);
    throw error;
  }
}

function clearData() {
  archiveAndReset();
}

function clearDataFeb25() {
  archiveAndReset();
}

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
// 4. Case Timestamp Sync — Bridge Functions & Web App
// =============================================================================
function requestSync() {
  PropertyManager.setMany({
    [TIMESTAMP_SYNC_CONFIG.PROPERTIES.SYNC_REQUESTED]: "true",
    [TIMESTAMP_SYNC_CONFIG.PROPERTIES.SYNC_STATUS]: "pending",
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Sync requested! Switch to your Tech24 Dashboard tab. The sync will start automatically.",
    "Sync Ready",
    12
  );
}

function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const action = params.action;

    switch (action) {
      case "ping":
        return ResponseHandler.json({ ok: true, ts: new Date().toISOString() });

      case "checkTrigger":
        return ResponseHandler.json({
          shouldRun:
            PropertyManager.get(TIMESTAMP_SYNC_CONFIG.PROPERTIES.SYNC_REQUESTED) === "true",
          status:
            PropertyManager.get(TIMESTAMP_SYNC_CONFIG.PROPERTIES.SYNC_STATUS) || "idle",
        });

      case "getPendingCases":
        PropertyManager.set(TIMESTAMP_SYNC_CONFIG.PROPERTIES.SYNC_STATUS, "running");
        return ResponseHandler.json({
          rows: TimestampSyncRepository.getPendingRows(),
        });

      case "submitResult":
        const result = TimestampSyncRepository.writeSingleResult({
          row: params.row,
          regDate: params.regDate || "",
          regTime: params.regTime || "",
          closedDate: params.closedDate || "",
          closedTime: params.closedTime || "",
        });
        return ResponseHandler.json({ ok: true, ...result });

      case "clearTrigger":
        PropertyManager.setMany({
          [TIMESTAMP_SYNC_CONFIG.PROPERTIES.SYNC_REQUESTED]: "false",
          [TIMESTAMP_SYNC_CONFIG.PROPERTIES.SYNC_STATUS]: "done",
          [TIMESTAMP_SYNC_CONFIG.PROPERTIES.LAST_RUN_SUMMARY]: params.summary || "{}",
          [TIMESTAMP_SYNC_CONFIG.PROPERTIES.LAST_RUN_TIME]: new Date().toISOString(),
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
  },
};

const TimestampSyncRepository = {
  getSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error("No active spreadsheet found.");
    }
    const sheet =
      spreadsheet.getSheetByName(TIMESTAMP_SYNC_CONFIG.SHEET_NAME) ||
      spreadsheet.getSheetByName("Data Entry Form") ||
      spreadsheet.getSheets()[0];
    if (!sheet) {
      throw new Error(`Sheet "${TIMESTAMP_SYNC_CONFIG.SHEET_NAME}" not found.`);
    }
    return sheet;
  },

  getPendingRows() {
    const sheet = this.getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= TIMESTAMP_SYNC_CONFIG.HEADER_ROW) return [];

    const numRows = lastRow - TIMESTAMP_SYNC_CONFIG.HEADER_ROW;
    const maxCol = Math.max(...Object.values(TIMESTAMP_SYNC_CONFIG.COLUMNS));
    const data = sheet
      .getRange(TIMESTAMP_SYNC_CONFIG.HEADER_ROW + 1, 1, numRows, maxCol)
      .getValues();

    const pending = [];
    data.forEach((rowValues, idx) => {
      const rowNum = TIMESTAMP_SYNC_CONFIG.HEADER_ROW + 1 + idx;
      const caseId = rowValues[TIMESTAMP_SYNC_CONFIG.COLUMNS.CASE_ID - 1];
      if (!caseId) return;

      const regDate = rowValues[TIMESTAMP_SYNC_CONFIG.COLUMNS.REG_DATE - 1];
      const closedDate = rowValues[TIMESTAMP_SYNC_CONFIG.COLUMNS.CLOSED_DATE - 1];

      const isRegEmpty = regDate === "" || regDate === null || regDate === undefined;
      const isClosedEmpty = closedDate === "" || closedDate === null || closedDate === undefined;

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
    const startCol = TIMESTAMP_SYNC_CONFIG.COLUMNS.REG_DATE;
    const range = sheet.getRange(row, startCol, 1, 4);
    const currentValues = range.getValues()[0];

    let modified = false;
    const isCellEmpty = (val) => val === "" || val === null || val === undefined;

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
  },
};

const ResponseHandler = {
  json(data) {
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(
      ContentService.MimeType.JSON
    );
  },

  error(message, details = null) {
    const payload = { error: message };
    if (details) payload.details = details;
    return this.json(payload);
  },
};

// =============================================================================
// 5. Weekly Report Builder Engine
// =============================================================================
class WeeklyReportBuilder {
  constructor(spreadsheet, doc) {
    this.spreadsheet = spreadsheet;
    this.doc = doc;
    this.body = doc.getBody();
    this.tableCount = 0;
    this.tableMetadata = null;
    this.tableRangeCache = {};
  }

  build() {
    this.body.clear();

    this.addCoverPage();
    this.body.appendPageBreak();

    this.addDistrictSummaryTable();
    this.addCaseSummaryByDistrictTable();
    this.addCasesClosedAfterRegistrationDateTable();
    this.addOngoingCasesTable();
    this.addBankSpecificCaseCountTable();
    this.addRegistrationTypeCountsTable();
    this.addPmSummaryTable();
    this.addBankSpecificPmCountTable();
    this.addWeeklyTasksTable();
    this.addWeeklyMeetingOverviewTable();
    this.addPerDiemCostTable();
    this.addChangedSparePartsTable();
    this.addTechniciansWeeklyActivityTable();

    this.addChallengesSection();
    this.addNextWeeksPlanSection();

    this.addPmPlansTable();
    this.addSignOffBlock();

    this.doc.saveAndClose();
  }

  addCoverPage() {
    const cover = REPORT_CONFIG.COVER;

    let spacer;
    if (
      this.body.getNumChildren() > 0 &&
      this.body.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH &&
      this.isBlank(this.body.getChild(0).asParagraph().getText())
    ) {
      spacer = this.body.getChild(0).asParagraph();
      spacer.setText(" ");
    } else {
      spacer = this.body.insertParagraph(0, " ");
    }

    spacer.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    spacer.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    spacer.setSpacingBefore(0);
    spacer.editAsText().setBold(true);
    this.setElementBlackText(spacer);

    const title = this.body.insertParagraph(
      this.body.getChildIndex(spacer) + 1,
      cover.title
    );
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    title.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    title.setSpacingBefore(0);
    title.editAsText().setBold(true);
    this.setElementBlackText(title);

    this.body.appendParagraph("");

    const dataSetValues = this.getDisplayValuesSafely("Data Set", "I1:I1");
    const periodValue =
      (dataSetValues && dataSetValues.length > 0
        ? this.toText(dataSetValues[0][0])
        : "") || getWeeklyReportPeriod();
    this.appendLabeledParagraph("Period", periodValue);
    this.appendLabeledParagraph("District", cover.district);
    this.appendLabeledParagraph("Company", cover.company);

    this.body.appendParagraph("");

    const overviewTitle = this.body.appendParagraph(cover.overviewTitle);
    overviewTitle.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    overviewTitle.editAsText().setBold(true);
    this.setElementBlackText(overviewTitle);

    const overviewIntro = this.body.appendParagraph(cover.overviewIntro);
    overviewIntro.editAsText().setBold(false);
    this.setElementBlackText(overviewIntro);

    cover.overviewItems.forEach((item) => {
      this.appendLabeledListItem({
        label: item.label,
        value: item.value,
        glyphType: DocumentApp.GlyphType.BULLET,
        nestingLevel: 0,
      });
    });

    for (let i = 0; i < 15; i += 1) {
      this.body.appendParagraph("");
    }

    const closing = this.body.appendParagraph(cover.overviewClosing);
    closing.editAsText().setBold(true);
    this.setElementBlackText(closing);
  }

  addDistrictSummaryTable() {
    const rows = this.getTableDisplayValues("DISTRICT_SUMMARY");
    this.addTableSection("District Summary", rows, {
      boldFirstRow: true,
      boldFirstColumn: true,
      boldLastRow: true,
    });
  }

  addCaseSummaryByDistrictTable() {
    const rows = this.buildCaseSummaryByDistrictRows();
    this.addTableSection("Case Summary by District", rows, {
      boldFirstRow: true,
      boldFirstColumn: true,
    });
  }

  addCasesClosedAfterRegistrationDateTable() {
    const filteredRows = this.filterRowsByFilledColumn(
      this.getTableDisplayValues("CASES_CLOSED_AFTER_REGISTRATION_DATE"),
      1
    );
    this.addTableSection("Cases Closed After Reg. Date", filteredRows, {
      boldFirstRow: true,
      boldFirstColumn: true,
    });
  }

  addOngoingCasesTable() {
    const filteredRows = this.filterRowsByFilledColumn(
      this.getTableDisplayValues("ONGOING_CASES"),
      1
    );
    const rows = this.excludeColumns(filteredRows, [filteredRows[0].length - 1]);
    this.addTableSection("Ongoing Cases", rows, {
      boldFirstRow: true,
      boldFirstColumn: true,
    });
  }

  addBankSpecificCaseCountTable() {
    const rows = this.getTableDisplayValues("BANK_SPECIFIC_CASE_COUNT");
    this.addTableSection("Bank Specific Case Count", rows, {
      boldFirstRow: true,
      boldFirstColumn: true,
      boldLastRow: true,
    });
  }

  addRegistrationTypeCountsTable() {
    const rows = this.getTableDisplayValues("REGISTRATION_TYPE_COUNTS");
    this.addTableSection("Registration Type Counts", rows, {
      boldFirstRow: true,
      boldFirstColumn: true,
      boldLastRow: true,
    });
  }

  addPmSummaryTable() {
    const rawRows = this.getTableDisplayValues("PM_SUMMARY");
    const lastColIndex = rawRows.length > 0 ? rawRows[0].length - 1 : 5;
    const colsToInclude = [];
    for (let col = 0; col < lastColIndex; col += 1) {
      colsToInclude.push(col);
    }
    const rows = this.includeColumns(rawRows, colsToInclude);
    this.addTableSection("PM Summary", rows, {
      boldFirstRow: true,
      boldFirstColumn: true,
      boldLastRow: true,
    });
  }

  addBankSpecificPmCountTable() {
    const rows = this.getTableDisplayValues("BANK_SPECIFIC_PM_COUNT");
    this.addTableSection("Bank Specific PM Count", rows, {
      boldFirstRow: true,
      boldFirstColumn: true,
      boldLastRow: true,
    });
  }

  addWeeklyTasksTable() {
    const filteredRows = this.filterRowsByFilledColumn(
      this.getTableDisplayValues("WEEKLY_TASKS"),
      0
    );
    const lastColIndex = filteredRows.length > 0 ? filteredRows[0].length - 1 : 3;
    const rows = this.excludeColumns(filteredRows, [lastColIndex]);
    this.addTableSection("Weekly Tasks", rows, {
      boldFirstRow: true,
    });
  }

  addWeeklyMeetingOverviewTable() {
    const filteredRows = this.filterRowsByFilledColumn(
      this.getTableDisplayValues("WEEKLY_MEETING_OVERVIEW"),
      0
    );
    const lastColIndex = filteredRows.length > 0 ? filteredRows[0].length - 1 : 3;
    const rows = this.excludeColumns(filteredRows, [lastColIndex]);
    this.addTableSection("Weekly Meeting Overview", rows, {
      boldFirstRow: true,
    });
  }

  addPerDiemCostTable() {
    const filteredRows = this.filterRowsByAnyFilledCell(
      this.getTableDisplayValues("PER_DIEM_COST")
    );
    if (filteredRows.length <= 1) {
      this.addTableSection("Per Diem Cost", filteredRows, { boldFirstRow: true });
      return;
    }

    const rows = this.appendTableTotalsRow(filteredRows);
    this.addTableSection("Per Diem Cost", rows, {
      boldFirstRow: true,
      boldLastRow: true,
    });
  }

  addChangedSparePartsTable() {
    const rows = this.buildChangedSparePartsRows();
    this.addTableSection("Changed Spare Parts", rows, {
      boldFirstRow: true,
      boldFirstColumn: true,
    });
  }

  addTechniciansWeeklyActivityTable() {
    const rows = this.buildTechniciansWeeklyActivityRows();
    this.addTableSection("Technicians Weekly Activity", rows, {
      boldFirstRow: this.hasHeaderRow(rows),
    });
  }

  addChallengesSection() {
    const rows = this.getTableDisplayValues("CHALLENGES");

    const title = this.body.appendParagraph("Challenges");
    title.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    title.editAsText().setBold(false);
    this.setElementBlackText(title);

    if (!rows || rows.length === 0) {
      const noData = this.body.appendParagraph(this.getNoDataMessage("Challenges"));
      noData.editAsText().setBold(false);
      this.setElementBlackText(noData);
      return;
    }

    const header = rows[0];
    const dataRows = rows.slice(1);
    const hasData = this.renderBulletItemsFromRows(dataRows, header, {
      useNumberedListForFirst: true,
    });

    if (!hasData) {
      const noData = this.body.appendParagraph(this.getNoDataMessage("Challenges"));
      noData.editAsText().setBold(false);
      this.setElementBlackText(noData);
    }
  }

  addNextWeeksPlanSection() {
    const title = this.body.appendParagraph("Next Week's Plan");
    title.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    title.editAsText().setBold(false);
    this.setElementBlackText(title);

    this.addNextWeekBulletSubSection("Weekly Tasks", "WEEKLY_TASKS");
    this.addNextWeekBulletSubSection("Weekly Meetings", "WEEKLY_MEETING_OVERVIEW");
  }

  addNextWeekBulletSubSection(subTitle, tableSourceKey) {
    const rows = this.getTableDisplayValues(tableSourceKey);
    if (!rows || rows.length === 0) {
      return;
    }

    const subHeader = this.body.appendParagraph(subTitle);
    subHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    subHeader.editAsText().setBold(false);
    this.setElementBlackText(subHeader);

    const header = rows[0];
    const dataRows = rows.slice(1);
    this.renderBulletItemsFromRows(dataRows, header, {
      useNumberedListForFirst: false,
    });
  }

  renderBulletItemsFromRows(dataRows, headerRow, options = {}) {
    let count = 0;
    for (let i = 0; i < dataRows.length; i += 1) {
      const row = dataRows[i];
      const mainItem = this.toText(row[0]);
      if (this.isBlank(mainItem)) {
        continue;
      }

      count += 1;
      const listItem = this.body.appendListItem(mainItem);
      listItem.setNestingLevel(0);
      if (options.useNumberedListForFirst) {
        listItem.setGlyphType(DocumentApp.GlyphType.NUMBER);
        listItem.editAsText().setBold(true);
      } else {
        listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
        listItem.editAsText().setBold(false);
      }
      this.setElementBlackText(listItem);

      for (let col = 1; col < row.length; col += 1) {
        const label = this.toText(headerRow[col]) || `Detail ${col}`;
        const value = this.toText(row[col]);
        if (this.isBlank(value)) {
          continue;
        }

        this.appendLabeledListItem({
          label,
          value,
          glyphType: DocumentApp.GlyphType.HOLLOW_BULLET,
          nestingLevel: 1,
        });
      }
    }

    return count > 0;
  }

  addPmPlansTable() {
    const rawRows = this.getTableDisplayValues("PM_SUMMARY");
    const lastColIndex = rawRows.length > 0 ? rawRows[0].length - 1 : 5;
    const rows = this.includeColumns(rawRows, [0, 1, lastColIndex]);

    this.addTableSection("PM Plans", rows, {
      boldFirstRow: true,
      boldFirstColumn: true,
      boldLastRow: true,
    });
  }

  addSignOffBlock() {
    this.body.appendParagraph("");

    REPORT_CONFIG.SIGN_OFF_NAMES.forEach((name) => {
      const paragraph = this.body.appendParagraph(name);
      paragraph.editAsText().setBold(true);
      this.setElementBlackText(paragraph);
    });
  }

  addTableSection(title, rows, styleOptions) {
    const sectionTitle = this.body.appendParagraph(title);
    sectionTitle.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    sectionTitle.editAsText().setBold(false);
    this.setElementBlackText(sectionTitle);

    if (this.hasNoTableData(rows)) {
      const noData = this.body.appendParagraph(this.getNoDataMessage(title));
      noData.editAsText().setBold(false);
      this.setElementBlackText(noData);
      return;
    }

    const normalizedRows = rows.map((row) =>
      row.map((cell) => this.toText(cell))
    );
    const table = this.body.appendTable(normalizedRows);
    this.styleTable(table, styleOptions);
    this.addTableCaption(title, table);
  }

  addTableCaption(title, table) {
    this.tableCount += 1;
    const captionText = `Table ${this.tableCount}: ${title}`;
    const tableIndex = this.body.getChildIndex(table);
    const nextIndex = tableIndex + 1;

    let caption;
    if (nextIndex < this.body.getNumChildren()) {
      const nextElement = this.body.getChild(nextIndex);
      if (nextElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
        caption = nextElement.asParagraph();
        caption.setText(captionText);
      } else {
        caption = this.body.insertParagraph(nextIndex, captionText);
      }
    } else {
      caption = this.body.appendParagraph(captionText);
    }

    caption.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    caption.setSpacingBefore(10);
    caption.setSpacingAfter(0);
    caption.editAsText().setBold(false);
    this.setElementBlackText(caption);
  }

  styleTable(
    table,
    { boldFirstRow = false, boldFirstColumn = false, boldLastRow = false }
  ) {
    const rowCount = table.getNumRows();
    if (rowCount === 0) {
      return;
    }

    for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
      const row = table.getRow(rowIndex);
      const cellCount = row.getNumCells();

      for (let cellIndex = 0; cellIndex < cellCount; cellIndex += 1) {
        const cell = table.getCell(rowIndex, cellIndex);
        const text = cell.editAsText();
        const shouldBeBold =
          (boldFirstRow && rowIndex === 0) ||
          (boldFirstColumn && cellIndex === 0) ||
          (boldLastRow && rowIndex === rowCount - 1);

        cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        cell.setBackgroundColor(REPORT_CONFIG.TABLE_BACKGROUND);
        text.setBold(shouldBeBold);
        text.setForegroundColor(REPORT_CONFIG.TEXT_COLOR);
      }
    }
  }

  appendLabeledParagraph(label, value) {
    const paragraph = this.body.appendParagraph(`${label}: ${value}`);
    this.boldLabel(paragraph, label);
  }

  appendLabeledListItem({ label, value, glyphType, nestingLevel }) {
    const textValue = `${label}: ${value}`;
    const item = this.body.appendListItem(textValue);
    item.setNestingLevel(nestingLevel);
    item.setGlyphType(glyphType);
    this.boldLabel(item, label);
  }

  boldLabel(listItem, label) {
    const text = listItem.editAsText();
    text.setBold(false);

    const full = text.getText();
    const labelEnd = label.length;
    if (full.length >= labelEnd + 1) {
      text.setBold(0, labelEnd, true);
    }
    text.setForegroundColor(REPORT_CONFIG.TEXT_COLOR);
  }

  setElementBlackText(element) {
    const text = element.editAsText();
    text.setForegroundColor(REPORT_CONFIG.TEXT_COLOR);
  }

  appendTableTotalsRow(rows) {
    if (!rows || rows.length <= 1) {
      return rows;
    }

    const colCount = rows[0].length;
    const lastColIndex = colCount - 1;

    let total = 0;
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const val = this.toNumber(rows[rowIndex][lastColIndex]);
      total += val;
    }

    const totalRow = new Array(colCount).fill("");
    totalRow[0] = "Total";
    totalRow[lastColIndex] = this.formatAmountWithCommas(total);

    return rows.concat([totalRow]);
  }

  formatAmountWithCommas(value) {
    if (isNaN(value)) return "0";
    const isFloat = value % 1 !== 0;
    const parts = (isFloat ? value.toFixed(2) : String(Math.round(value))).split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return parts.join(".");
  }

  buildCaseSummaryByDistrictRows() {
    let thisWeekDisplay = this.getTableDisplayValues("CASE_SUMMARY_BY_DISTRICT");
    let thisWeekRaw = this.getTableRawValues("CASE_SUMMARY_BY_DISTRICT");

    let lastWeekRaw = null;
    try {
      lastWeekRaw = this.getTableRawValues("LAST_WEEK_CASE_AMOUNT");
    } catch (e) {
      lastWeekRaw = this.getRawValuesSafely("Last Week Case Amount", "A1:F7");
    }

    if (thisWeekDisplay.length > 1) {
      thisWeekDisplay = thisWeekDisplay.slice(0, -1);
      thisWeekRaw = thisWeekRaw.slice(0, -1);
    }
    if (lastWeekRaw && lastWeekRaw.length > 1) {
      lastWeekRaw = lastWeekRaw.slice(0, -1);
    }

    const rows = thisWeekDisplay.map((row) => row.slice());

    const lastWeekMap = {};
    if (lastWeekRaw && lastWeekRaw.length > 0) {
      for (let r = 1; r < lastWeekRaw.length; r += 1) {
        const districtKey =
          this.toText(lastWeekRaw[r][0]).toLowerCase() ||
          (lastWeekRaw[r][1] ? this.toText(lastWeekRaw[r][1]).toLowerCase() : "");
        if (districtKey) {
          lastWeekMap[districtKey] = lastWeekRaw[r];
        }
      }
    }

    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const districtKey =
        this.toText(thisWeekDisplay[rowIndex][1]).toLowerCase() ||
        this.toText(thisWeekDisplay[rowIndex][0]).toLowerCase();

      const matchingLwRow =
        lastWeekMap[districtKey] || (lastWeekRaw ? lastWeekRaw[rowIndex] : null);

      for (let colIndex = 2; colIndex < rows[rowIndex].length; colIndex += 1) {
        const currentRaw = thisWeekRaw[rowIndex][colIndex];
        const lwColIndex = colIndex - 1;
        const previousRaw =
          matchingLwRow && lwColIndex < matchingLwRow.length
            ? matchingLwRow[lwColIndex]
            : "";

        if (this.isBlank(currentRaw) && this.isBlank(previousRaw)) {
          rows[rowIndex][colIndex] = "";
          continue;
        }

        const currentValue = this.toNumber(currentRaw);
        const previousValue = this.toNumber(previousRaw);
        rows[rowIndex][colIndex] = this.formatComparisonCell(
          currentValue,
          previousValue
        );
      }
    }

    return rows;
  }

  buildChangedSparePartsRows() {
    const displayRows = this.getTableDisplayValues("CHANGED_SPARE_PARTS");
    const rawRows = this.getTableRawValues("CHANGED_SPARE_PARTS");
    if (displayRows.length === 0) {
      return [];
    }

    const result = [displayRows[0].slice()];
    for (let i = 1; i < displayRows.length; i += 1) {
      const row = displayRows[i].slice();
      const firstColumnValue = this.toText(row[0]);
      if (this.isBlank(firstColumnValue)) {
        continue;
      }

      const statusColIndex = row.length - 1;
      const rawStatus = rawRows[i][statusColIndex];
      row[statusColIndex] = this.isChecked(rawStatus) ? "Returned" : "Pending";
      result.push(row);
    }

    return result;
  }

  buildTechniciansWeeklyActivityRows() {
    const rows = this.getTableDisplayValues("TECHNICIANS_WEEKLY_ACTIVITY");
    if (rows.length === 0) {
      return [];
    }

    const hasHeader = this.hasHeaderRow(rows);
    const header = hasHeader ? [["No"].concat(rows[0])] : [["No"]];
    const startIndex = hasHeader ? 1 : 0;

    const dataRows = rows
      .slice(startIndex)
      .filter((row) => !this.isRowBlank(row))
      .sort((a, b) => {
        const totalA = this.toNumber(a[a.length - 1]);
        const totalB = this.toNumber(b[b.length - 1]);
        return totalB - totalA;
      })
      .map((row, index) => [String(index + 1)].concat(row));

    return header.concat(dataRows);
  }

  hasHeaderRow(rows) {
    if (!rows || rows.length === 0) {
      return false;
    }

    const firstRowLastCell = this.toText(
      rows[0][rows[0].length - 1]
    ).toLowerCase();
    return firstRowLastCell === "total activity";
  }

  filterRowsByFilledColumn(rows, columnIndex) {
    if (!rows || rows.length === 0) {
      return [];
    }

    const output = [rows[0]];
    for (let i = 1; i < rows.length; i += 1) {
      if (!this.isBlank(rows[i][columnIndex])) {
        output.push(rows[i]);
      }
    }

    return output;
  }

  filterRowsByAnyFilledCell(rows) {
    if (!rows || rows.length === 0) {
      return [];
    }

    const output = [rows[0]];
    for (let i = 1; i < rows.length; i += 1) {
      if (!this.isRowBlank(rows[i])) {
        output.push(rows[i]);
      }
    }

    return output;
  }

  excludeColumns(rows, excludedIndexes) {
    if (!rows || rows.length === 0) {
      return [];
    }

    const excludedSet = new Set(excludedIndexes);
    return rows.map((row) =>
      row.filter((_, columnIndex) => !excludedSet.has(columnIndex))
    );
  }

  includeColumns(rows, includedIndexes) {
    if (!rows || rows.length === 0) {
      return [];
    }

    return rows.map((row) =>
      includedIndexes.map((columnIndex) =>
        columnIndex < row.length ? row[columnIndex] : ""
      )
    );
  }

  hasNoTableData(rows) {
    if (!rows || rows.length === 0) {
      return true;
    }

    const nonBlankRows = rows.filter((row) => !this.isRowBlank(row));
    return nonBlankRows.length <= 1;
  }

  getNoDataMessage(title) {
    return `There is no data on ${title} Table`;
  }

  formatComparisonCell(currentValue, previousValue) {
    const percentChange =
      previousValue === 0
        ? currentValue === 0
          ? 0
          : 100
        : ((currentValue - previousValue) / Math.abs(previousValue)) * 100;

    const roundedPercent = Math.round(percentChange);
    const sign = roundedPercent > 0 ? "+" : "";
    return `${currentValue} (vs. ${previousValue} LW, ${sign}${roundedPercent}%)`;
  }

  getDisplayValuesSafely(sheetName, rangeA1) {
    const sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return null;
    }
    return sheet.getRange(rangeA1).getDisplayValues();
  }

  getRawValuesSafely(sheetName, rangeA1) {
    const sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return null;
    }
    return sheet.getRange(rangeA1).getValues();
  }

  getTableDisplayValues(sourceKey) {
    return this.getTableRange(sourceKey).getDisplayValues();
  }

  getTableRawValues(sourceKey) {
    return this.getTableRange(sourceKey).getValues();
  }

  getTableRange(sourceKey) {
    if (!this.tableRangeCache[sourceKey]) {
      this.tableRangeCache[sourceKey] = SpreadsheetRepository.getTableRange(
        this.spreadsheet,
        this.getTableMetadata(),
        SpreadsheetRepository.getTableSource(sourceKey)
      );
    }

    return this.tableRangeCache[sourceKey];
  }

  getTableMetadata() {
    if (!this.tableMetadata) {
      this.tableMetadata = SpreadsheetRepository.getTableMetadata(this.spreadsheet);
    }

    return this.tableMetadata;
  }

  getSheet(sheetName) {
    const sheet = this.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    return sheet;
  }

  isChecked(value) {
    if (value === true) {
      return true;
    }
    return this.toText(value).toLowerCase() === "true";
  }

  isBlank(value) {
    return this.toText(value) === "";
  }

  isRowBlank(row) {
    return row.every((cell) => this.isBlank(cell));
  }

  toText(value) {
    if (value === null || value === undefined) {
      return "";
    }
    return String(value).trim();
  }

  toNumber(value) {
    if (typeof value === "number") {
      return value;
    }

    const text = this.toText(value).replace(/,/g, "");
    if (text === "") {
      return 0;
    }

    const numeric = Number(text);
    return Number.isFinite(numeric) ? numeric : 0;
  }
}

// =============================================================================
// 6. Spreadsheet Repository & Metadata Helpers
// =============================================================================
const SpreadsheetRepository = {
  getSheetOrThrow(spreadsheet, sheetName) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`Sheet not found: ${sheetName}`);
    }
    return sheet;
  },

  getTableSource(sourceKey) {
    const source = REPORT_CONFIG.TABLE_SOURCES[sourceKey];
    if (!source) {
      throw new Error(`Table source not configured: ${sourceKey}`);
    }
    return source;
  },

  clearRangeValuesOnly(spreadsheet, sheetName, rangeA1) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return;
    const range = sheet.getRange(rangeA1);
    this.clearRangeObjectValuesOnly(range);
  },

  clearRangeObjectValuesOnly(range) {
    const formulas = range.getFormulas();
    const cleared = formulas.map((row) =>
      row.map((formula) => (formula === "" ? "" : formula))
    );

    range.setValues(cleared);
  },

  clearTableBodyValuesOnly(
    spreadsheet,
    tableMetadata,
    sourceKey,
    columnIndexes
  ) {
    const tableRange = this.getTableRange(
      spreadsheet,
      tableMetadata,
      this.getTableSource(sourceKey)
    );
    const bodyRange = this.getTableBodyRange(tableRange);
    if (!bodyRange) {
      return;
    }

    if (!columnIndexes || columnIndexes.length === 0) {
      this.clearRangeObjectValuesOnly(bodyRange);
      return;
    }

    columnIndexes.forEach((columnIndex) => {
      if (columnIndex < 0 || columnIndex >= tableRange.getNumColumns()) {
        return;
      }

      this.clearRangeObjectValuesOnly(
        bodyRange.offset(0, columnIndex, bodyRange.getNumRows(), 1)
      );
    });
  },

  uncheckTableBodyColumn(spreadsheet, tableMetadata, sourceKey, columnIndex) {
    const tableRange = this.getTableRange(
      spreadsheet,
      tableMetadata,
      this.getTableSource(sourceKey)
    );
    const bodyRange = this.getTableBodyRange(tableRange);
    if (
      !bodyRange ||
      columnIndex < 0 ||
      columnIndex >= tableRange.getNumColumns()
    ) {
      return;
    }

    bodyRange.offset(0, columnIndex, bodyRange.getNumRows(), 1).uncheck();
  },

  getTableBodyRange(tableRange) {
    const rowCount = tableRange.getNumRows();
    if (rowCount <= 1) {
      return null;
    }

    return tableRange.offset(1, 0, rowCount - 1, tableRange.getNumColumns());
  },

  getTableMetadata(spreadsheet) {
    const fields = "sheets(properties(title),tables(name,range))";

    if (typeof Sheets !== "undefined" && Sheets.Spreadsheets) {
      return Sheets.Spreadsheets.get(spreadsheet.getId(), { fields });
    }

    const url =
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheet.getId()}` +
      `?fields=${encodeURIComponent(fields)}`;
    const response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
      },
      muteHttpExceptions: true,
    });
    const responseCode = response.getResponseCode();
    if (responseCode < 200 || responseCode >= 300) {
      const errorText = response.getContentText();
      const apiError = this.parseApiError(errorText);
      if (
        apiError.reason === "SERVICE_DISABLED" ||
        /sheets api.*disabled|sheets api has not been used/i.test(
          apiError.message
        )
      ) {
        throw new Error(
          "Google Sheets API is disabled for this Apps Script project. Enable Google Sheets API in the Apps Script project, wait a few minutes, then rerun the report."
        );
      }

      throw new Error(
        `Failed to read Google Sheets table metadata: ${apiError.message}`
      );
    }

    return JSON.parse(response.getContentText());
  },

  getTableRange(spreadsheet, tableMetadata, source) {
    const sheet = this.getSheetOrThrow(spreadsheet, source.sheetName);
    const sheetResource = this.findSheetResource(tableMetadata, source.sheetName);
    const table = this.findTableResource(sheetResource, source.tableName);
    const gridRange = table.range || {};
    const startRowIndex = gridRange.startRowIndex || 0;
    const startColumnIndex = gridRange.startColumnIndex || 0;

    if (
      typeof gridRange.endRowIndex !== "number" ||
      typeof gridRange.endColumnIndex !== "number"
    ) {
      throw new Error(
        `Table range is incomplete: ${source.tableName} on ${source.sheetName}`
      );
    }

    return sheet.getRange(
      startRowIndex + 1,
      startColumnIndex + 1,
      gridRange.endRowIndex - startRowIndex,
      gridRange.endColumnIndex - startColumnIndex
    );
  },

  findSheetResource(tableMetadata, sheetName) {
    const sheets =
      tableMetadata && tableMetadata.sheets ? tableMetadata.sheets : [];
    const normalizedSheetName = this.normalizeTableLookupText(sheetName);
    const sheetResource = sheets.find((candidate) => {
      const title =
        candidate.properties && candidate.properties.title
          ? candidate.properties.title
          : "";
      return (
        title === sheetName ||
        this.normalizeTableLookupText(title) === normalizedSheetName
      );
    });

    if (!sheetResource) {
      throw new Error(`Sheet metadata not found: ${sheetName}`);
    }

    return sheetResource;
  },

  findTableResource(sheetResource, tableName) {
    const tables = sheetResource.tables || [];
    const normalizedTableName = this.normalizeTableLookupText(tableName);
    const table = tables.find(
      (candidate) =>
        candidate.name === tableName ||
        this.normalizeTableLookupText(candidate.name) === normalizedTableName
    );

    if (!table) {
      const sheetName =
        sheetResource.properties && sheetResource.properties.title
          ? sheetResource.properties.title
          : "";
      throw new Error(`Table not found: ${tableName} on ${sheetName}`);
    }

    return table;
  },

  parseApiError(errorText) {
    try {
      const parsed = JSON.parse(errorText);
      const error = parsed.error || {};
      const details = error.details || [];
      const errorInfo = details.find(
        (detail) => detail["@type"] === "type.googleapis.com/google.rpc.ErrorInfo"
      );

      return {
        message: error.message || errorText,
        reason: errorInfo && errorInfo.reason ? errorInfo.reason : "",
      };
    } catch (error) {
      return {
        message: errorText,
        reason: "",
      };
    }
  },

  normalizeTableLookupText(value) {
    return String(value || "")
      .replace(/_/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  },
};
