const REPORT_CONFIG = {
  SHEET_ID: "1fCC7NMoYakdcWbhAI9U-8WCrZ7hfIDaCfMMXInMVwbQ",
  DOC_ID: "1TbDEuQiFHeEEjy7Z_DfJ32c6we1s7Il3UOY3dS6n42s",
  TABLE_BACKGROUND: "#ffffff",
  TEXT_COLOR: "#000000",
  COVER: {
    title: "Weekly Report",
    period: "February 14 to February 20",
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
  TABLE_SOURCES: {
    DISTRICT_SUMMARY: {
      sheetName: "District Summary",
      tableName: "District Summary",
    },
    CASE_SUMMARY_BY_DISTRICT: {
      sheetName: "Case Summary",
      tableName: "Case Summary",
    },
    CASES_CLOSED_AFTER_REGISTRATION_DATE: {
      sheetName: "Case Summary",
      tableName: "Closed after Reg. Date",
    },
    ONGOING_CASES: {
      sheetName: "Case Summary",
      tableName: "Ongoing Cases",
    },
    BANK_SPECIFIC_CASE_COUNT: {
      sheetName: "Bank Specific Case",
      tableName: "Bank Specific Case Count",
    },
    REGISTRATION_TYPE_COUNTS: {
      sheetName: "Case Summary",
      tableName: "Case Registration Type Counts",
    },
    PM_SUMMARY: {
      sheetName: "PM Summary",
      tableName: "PM Summary",
    },
    BANK_SPECIFIC_PM_COUNT: {
      sheetName: "Bank Specific PM",
      tableName: "Bank Specific PM Count",
    },
    WEEKLY_TASKS: {
      sheetName: "Weekly Tasks",
      tableName: "Weekly Task",
    },
    WEEKLY_MEETING_OVERVIEW: {
      sheetName: "Weekly Meeting",
      tableName: "Weekly Meeting Overview",
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
      sheetName: "Technicians Weekly",
      tableName: "Technicians Weekly Activity",
    },
    CHALLENGES: {
      sheetName: "Challenges",
      tableName: "Challenges",
    },
  },
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Weekly Report")
    .addItem("Generate Report", "generateWeeklyReportFeb25")
    .addItem("Clear Data", "clearDataFeb25")
    .addToUi();
}

function generateWeeklyReportFeb25() {
  const ui = SpreadsheetApp.getUi();

  try {
    const spreadsheet = SpreadsheetApp.openById(REPORT_CONFIG.SHEET_ID);
    const doc = DocumentApp.openById(REPORT_CONFIG.DOC_ID);
    const builder = new WeeklyReportBuilder(spreadsheet, doc);
    builder.build();
    ui.alert("Weekly report generated successfully.");
  } catch (error) {
    ui.alert(`Failed to generate report: ${error.message}`);
    throw error;
  }
}

function clearDataFeb25() {
  const ui = SpreadsheetApp.getUi();
  const confirmation = ui.alert(
    "Clear Data",
    "This will clear configured data ranges while preserving formulas. Continue?",
    ui.ButtonSet.YES_NO,
  );

  if (confirmation !== ui.Button.YES) {
    return;
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(REPORT_CONFIG.SHEET_ID);
    const tableMetadata = getSpreadsheetTableMetadata_(spreadsheet);

    clearRangeValuesOnly_(spreadsheet, "Data Entry Form", "A3:P300");
    clearRangeValuesOnly_(spreadsheet, "Data Entry Form", "R3:V300");
    clearTableBodyValuesOnly_(
      spreadsheet,
      tableMetadata,
      "CASES_CLOSED_AFTER_REGISTRATION_DATE",
    );
    clearTableBodyValuesOnly_(spreadsheet, tableMetadata, "ONGOING_CASES");
    clearRangeValuesOnly_(spreadsheet, "SLA Penalities", "A2:A20");
    clearRangeValuesOnly_(spreadsheet, "SLA Penalities", "G2:G20");
    clearTableBodyValuesOnly_(spreadsheet, tableMetadata, "PM_SUMMARY", [
      3,
      4,
      5,
    ]);
    clearTableBodyValuesOnly_(spreadsheet, tableMetadata, "WEEKLY_TASKS");
    clearTableBodyValuesOnly_(
      spreadsheet,
      tableMetadata,
      "WEEKLY_MEETING_OVERVIEW",
    );
    clearTableBodyValuesOnly_(spreadsheet, tableMetadata, "PER_DIEM_COST");
    clearTableBodyValuesOnly_(spreadsheet, tableMetadata, "CHALLENGES");
    clearRangeValuesOnly_(spreadsheet, "Last Week Case Summary", "B2:F6");

    // Checkbox reset: checked rows become unchecked (TRUE -> FALSE).
    uncheckTableBodyColumn_(
      spreadsheet,
      tableMetadata,
      "CHANGED_SPARE_PARTS",
      8,
    );

    ui.alert("Configured data ranges were cleared successfully.");
  } catch (error) {
    ui.alert(`Failed to clear data: ${error.message}`);
    throw error;
  }
}

function clearRangeValuesOnly_(spreadsheet, sheetName, rangeA1) {
  const range = getSheetOrThrow_(spreadsheet, sheetName).getRange(rangeA1);
  clearRangeObjectValuesOnly_(range);
}

function clearRangeObjectValuesOnly_(range) {
  const formulas = range.getFormulas();
  const cleared = formulas.map((row) =>
    row.map((formula) => (formula === "" ? "" : formula)),
  );

  // Keep formulas and clear only plain values.
  range.setValues(cleared);
}

function clearTableBodyValuesOnly_(
  spreadsheet,
  tableMetadata,
  sourceKey,
  columnIndexes,
) {
  const tableRange = getTableRange_(
    spreadsheet,
    tableMetadata,
    getTableSource_(sourceKey),
  );
  const bodyRange = getTableBodyRange_(tableRange);
  if (!bodyRange) {
    return;
  }

  if (!columnIndexes || columnIndexes.length === 0) {
    clearRangeObjectValuesOnly_(bodyRange);
    return;
  }

  columnIndexes.forEach((columnIndex) => {
    if (columnIndex < 0 || columnIndex >= tableRange.getNumColumns()) {
      return;
    }

    clearRangeObjectValuesOnly_(
      bodyRange.offset(0, columnIndex, bodyRange.getNumRows(), 1),
    );
  });
}

function uncheckTableBodyColumn_(
  spreadsheet,
  tableMetadata,
  sourceKey,
  columnIndex,
) {
  const tableRange = getTableRange_(
    spreadsheet,
    tableMetadata,
    getTableSource_(sourceKey),
  );
  const bodyRange = getTableBodyRange_(tableRange);
  if (
    !bodyRange ||
    columnIndex < 0 ||
    columnIndex >= tableRange.getNumColumns()
  ) {
    return;
  }

  bodyRange.offset(0, columnIndex, bodyRange.getNumRows(), 1).uncheck();
}

function getTableBodyRange_(tableRange) {
  const rowCount = tableRange.getNumRows();
  if (rowCount <= 1) {
    return null;
  }

  return tableRange.offset(1, 0, rowCount - 1, tableRange.getNumColumns());
}

function getSheetOrThrow_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }
  return sheet;
}

function getTableSource_(sourceKey) {
  const source = REPORT_CONFIG.TABLE_SOURCES[sourceKey];
  if (!source) {
    throw new Error(`Table source not configured: ${sourceKey}`);
  }
  return source;
}

function getSpreadsheetTableMetadata_(spreadsheet) {
  const fields = "sheets(properties(title),tables(name,range))";

  if (typeof Sheets !== "undefined" && Sheets.Spreadsheets) {
    return Sheets.Spreadsheets.get(spreadsheet.getId(), {
      fields,
    });
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
    const apiError = parseApiError_(errorText);
    if (
      apiError.reason === "SERVICE_DISABLED" ||
      /sheets api.*disabled|sheets api has not been used/i.test(
        apiError.message,
      )
    ) {
      throw new Error(
        "Google Sheets API is disabled for this Apps Script project. Enable Google Sheets API in the Apps Script project, wait a few minutes, then rerun the report.",
      );
    }

    throw new Error(
      `Failed to read Google Sheets table metadata: ${apiError.message}`,
    );
  }

  return JSON.parse(response.getContentText());
}

function parseApiError_(errorText) {
  try {
    const parsed = JSON.parse(errorText);
    const error = parsed.error || {};
    const details = error.details || [];
    const errorInfo = details.find(
      (detail) => detail["@type"] === "type.googleapis.com/google.rpc.ErrorInfo",
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
}

function getTableRange_(spreadsheet, tableMetadata, source) {
  const sheet = getSheetOrThrow_(spreadsheet, source.sheetName);
  const sheetResource = findSheetResource_(tableMetadata, source.sheetName);
  const table = findTableResource_(sheetResource, source.tableName);
  const gridRange = table.range || {};
  const startRowIndex = gridRange.startRowIndex || 0;
  const startColumnIndex = gridRange.startColumnIndex || 0;

  if (
    typeof gridRange.endRowIndex !== "number" ||
    typeof gridRange.endColumnIndex !== "number"
  ) {
    throw new Error(
      `Table range is incomplete: ${source.tableName} on ${source.sheetName}`,
    );
  }

  return sheet.getRange(
    startRowIndex + 1,
    startColumnIndex + 1,
    gridRange.endRowIndex - startRowIndex,
    gridRange.endColumnIndex - startColumnIndex,
  );
}

function findSheetResource_(tableMetadata, sheetName) {
  const sheets =
    tableMetadata && tableMetadata.sheets ? tableMetadata.sheets : [];
  const normalizedSheetName = normalizeTableLookupText_(sheetName);
  const sheetResource = sheets.find((candidate) => {
    const title =
      candidate.properties && candidate.properties.title
        ? candidate.properties.title
        : "";
    return (
      title === sheetName ||
      normalizeTableLookupText_(title) === normalizedSheetName
    );
  });

  if (!sheetResource) {
    throw new Error(`Sheet metadata not found: ${sheetName}`);
  }

  return sheetResource;
}

function findTableResource_(sheetResource, tableName) {
  const tables = sheetResource.tables || [];
  const normalizedTableName = normalizeTableLookupText_(tableName);
  const table = tables.find(
    (candidate) =>
      candidate.name === tableName ||
      normalizeTableLookupText_(candidate.name) === normalizedTableName,
  );

  if (!table) {
    const sheetName =
      sheetResource.properties && sheetResource.properties.title
        ? sheetResource.properties.title
        : "";
    throw new Error(`Table not found: ${tableName} on ${sheetName}`);
  }

  return table;
}

function normalizeTableLookupText_(value) {
  return String(value || "")
    .replace(/_/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

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
      cover.title,
    );
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    title.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    title.setSpacingBefore(0);
    title.editAsText().setBold(true);
    this.setElementBlackText(title);

    this.body.appendParagraph("");

    const periodValue =
      this.toText(this.getDisplayValues("Data Set", "I1:I1")[0][0]) ||
      cover.period;
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
      1,
    );
    const rows = this.excludeColumns(filteredRows, [6]);
    this.addTableSection("Cases Closed After Reg. Date", rows, {
      boldFirstRow: true,
      boldFirstColumn: true,
    });
  }

  addOngoingCasesTable() {
    const filteredRows = this.filterRowsByFilledColumn(
      this.getTableDisplayValues("ONGOING_CASES"),
      1,
    );
    const rows = this.excludeColumns(filteredRows, [6]);
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
    const rows = this.includeColumns(
      this.getTableDisplayValues("PM_SUMMARY"),
      [0, 1, 2, 3, 4],
    );
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
      0,
    );
    const rows = this.excludeColumns(filteredRows, [3]);
    this.addTableSection("Weekly Tasks", rows, {
      boldFirstRow: true,
    });
  }

  addWeeklyMeetingOverviewTable() {
    const filteredRows = this.filterRowsByFilledColumn(
      this.getTableDisplayValues("WEEKLY_MEETING_OVERVIEW"),
      0,
    );
    const rows = this.excludeColumns(filteredRows, [3]);
    this.addTableSection("Weekly Meeting Overview", rows, {
      boldFirstRow: true,
    });
  }

  addPerDiemCostTable() {
    const rows = this.filterRowsByAnyFilledCell(
      this.getTableDisplayValues("PER_DIEM_COST"),
    );
    this.addTableSection("Per Diem Cost", rows, {
      boldFirstRow: true,
    });
  }

  addChangedSparePartsTable() {
    const rows = this.excludeColumns(this.buildChangedSparePartsRows(), [6]);
    this.addTableSection("Changed Spare Parts", rows, {
      boldFirstRow: true,
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
    const startIndex = this.isChallengesHeaderRow(rows[0]) ? 1 : 0;

    const title = this.body.appendParagraph("Challenges");
    title.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    title.editAsText().setBold(false);
    this.setElementBlackText(title);

    let hasData = false;
    for (let i = startIndex; i < rows.length; i += 1) {
      const challenge = this.toText(rows[i][0]);
      const description = this.toText(rows[i][1]);
      const actionTaken = this.toText(rows[i][2]);

      if (this.isBlank(challenge)) {
        continue;
      }
      hasData = true;

      const challengeItem = this.body.appendListItem(challenge);
      challengeItem.setNestingLevel(0);
      challengeItem.setGlyphType(DocumentApp.GlyphType.NUMBER);
      challengeItem.editAsText().setBold(true);
      this.setElementBlackText(challengeItem);

      this.appendLabeledListItem({
        label: "Description",
        value: description,
        glyphType: DocumentApp.GlyphType.HOLLOW_BULLET,
        nestingLevel: 1,
      });
      this.appendLabeledListItem({
        label: "Action Taken",
        value: actionTaken,
        glyphType: DocumentApp.GlyphType.HOLLOW_BULLET,
        nestingLevel: 1,
      });
    }

    if (!hasData) {
      const noData = this.body.appendParagraph(
        this.getNoDataMessage("Challenges"),
      );
      noData.editAsText().setBold(false);
      this.setElementBlackText(noData);
    }
  }

  addNextWeeksPlanSection() {
    const title = this.body.appendParagraph("Next Week's Plan");
    title.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    title.editAsText().setBold(false);
    this.setElementBlackText(title);

    this.addNextWeekSubSection({
      subTitle: "Weekly Tasks",
      tableSourceKey: "WEEKLY_TASKS",
      planColumnIndex: 3,
      defaultDetailOneLabel: "Task",
      defaultDetailTwoLabel: "Plan",
    });

    this.addNextWeekSubSection({
      subTitle: "Weekly Meetings",
      tableSourceKey: "WEEKLY_MEETING_OVERVIEW",
      planColumnIndex: 3,
      defaultDetailOneLabel: "Topic",
      defaultDetailTwoLabel: "Plan",
    });
  }

  addPmPlansTable() {
    const rows = this.includeColumns(this.getTableDisplayValues("PM_SUMMARY"), [
      0,
      1,
      5,
    ]);

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

  addNextWeekSubSection({
    subTitle,
    tableSourceKey,
    planColumnIndex,
    defaultDetailOneLabel,
    defaultDetailTwoLabel,
  }) {
    const rows = this.getTableDisplayValues(tableSourceKey);
    if (rows.length === 0) {
      return;
    }

    const subHeader = this.body.appendParagraph(subTitle);
    subHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
    subHeader.editAsText().setBold(false);
    this.setElementBlackText(subHeader);

    const header = rows[0];
    const detailTwoIndex =
      typeof planColumnIndex === "number" ? planColumnIndex : 2;
    const detailOneLabel = this.toText(header[1]) || defaultDetailOneLabel;
    const detailTwoLabel =
      this.toText(header[detailTwoIndex]) || defaultDetailTwoLabel;

    for (let i = 1; i < rows.length; i += 1) {
      const mainItem = this.toText(rows[i][0]);
      const detailOneValue = this.toText(rows[i][1]);
      const detailTwoValue = this.toText(rows[i][detailTwoIndex]);

      if (this.isBlank(mainItem)) {
        continue;
      }

      const districtItem = this.body.appendListItem(mainItem);
      districtItem.setNestingLevel(0);
      districtItem.setGlyphType(DocumentApp.GlyphType.BULLET);
      this.setElementBlackText(districtItem);

      this.appendLabeledListItem({
        label: detailOneLabel,
        value: detailOneValue,
        glyphType: DocumentApp.GlyphType.HOLLOW_BULLET,
        nestingLevel: 1,
      });
      this.appendLabeledListItem({
        label: detailTwoLabel,
        value: detailTwoValue,
        glyphType: DocumentApp.GlyphType.HOLLOW_BULLET,
        nestingLevel: 1,
      });
    }
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
      row.map((cell) => this.toText(cell)),
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
    { boldFirstRow = false, boldFirstColumn = false, boldLastRow = false },
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

  buildCaseSummaryByDistrictRows() {
    const thisWeekDisplay = this.getTableDisplayValues(
      "CASE_SUMMARY_BY_DISTRICT",
    );
    const thisWeekRaw = this.getTableRawValues("CASE_SUMMARY_BY_DISTRICT");
    const lastWeekRaw = this.getRawValues("Last Week Case Summary", "A1:F6");
    const rows = thisWeekDisplay.map((row) => row.slice());

    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      for (let colIndex = 2; colIndex < rows[rowIndex].length; colIndex += 1) {
        const currentRaw = thisWeekRaw[rowIndex][colIndex];
        const previousRaw =
          lastWeekRaw[rowIndex] && colIndex - 1 < lastWeekRaw[rowIndex].length
            ? lastWeekRaw[rowIndex][colIndex - 1]
            : "";

        if (this.isBlank(currentRaw) && this.isBlank(previousRaw)) {
          rows[rowIndex][colIndex] = "";
          continue;
        }

        const currentValue = this.toNumber(currentRaw);
        const previousValue = this.toNumber(previousRaw);
        rows[rowIndex][colIndex] = this.formatComparisonCell(
          currentValue,
          previousValue,
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

      const rawStatus = rawRows[i][8];
      row[8] = this.isChecked(rawStatus) ? "Returned" : "Pending";
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
    const header = hasHeader ? [["No"].concat(rows[0])] : [];
    const startIndex = hasHeader ? 1 : 0;

    const dataRows = rows
      .slice(startIndex)
      .filter((row) => !this.isRowBlank(row))
      .sort(
        (a, b) =>
          this.toNumber(b[b.length - 1]) - this.toNumber(a[a.length - 1]),
      )
      .map((row, index) => [String(index + 1)].concat(row));

    return header.concat(dataRows);
  }

  hasHeaderRow(rows) {
    if (!rows || rows.length === 0) {
      return false;
    }

    const firstRowLastCell = this.toText(
      rows[0][rows[0].length - 1],
    ).toLowerCase();
    return firstRowLastCell === "total activity";
  }

  isChallengesHeaderRow(row) {
    if (!row || row.length < 3) {
      return false;
    }

    return (
      this.toText(row[0]).toLowerCase() === "challenge" &&
      this.toText(row[1]).toLowerCase() === "description" &&
      this.toText(row[2]).toLowerCase() === "action taken"
    );
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
      row.filter((_, columnIndex) => !excludedSet.has(columnIndex)),
    );
  }

  includeColumns(rows, includedIndexes) {
    if (!rows || rows.length === 0) {
      return [];
    }

    return rows.map((row) =>
      includedIndexes.map((columnIndex) =>
        columnIndex < row.length ? row[columnIndex] : "",
      ),
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
    const arrow =
      currentValue > previousValue
        ? "↑"
        : currentValue < previousValue
          ? "↓"
          : "→";

    const percentChange =
      previousValue === 0
        ? currentValue === 0
          ? 0
          : 100
        : ((currentValue - previousValue) / Math.abs(previousValue)) * 100;

    const roundedPercent = Math.round(percentChange);
    const sign = roundedPercent > 0 ? "+" : "";
    return `${currentValue} (${arrow} ${sign}${roundedPercent}% from ${previousValue})`;
  }

  getDisplayValues(sheetName, rangeA1) {
    const sheet = this.getSheet(sheetName);
    return sheet.getRange(rangeA1).getDisplayValues();
  }

  getRawValues(sheetName, rangeA1) {
    const sheet = this.getSheet(sheetName);
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
      this.tableRangeCache[sourceKey] = getTableRange_(
        this.spreadsheet,
        this.getTableMetadata(),
        getTableSource_(sourceKey),
      );
    }

    return this.tableRangeCache[sourceKey];
  }

  getTableMetadata() {
    if (!this.tableMetadata) {
      this.tableMetadata = getSpreadsheetTableMetadata_(this.spreadsheet);
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

