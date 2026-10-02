/**
 * Weekly Report Google Docs Formatter
 * Version: 2.2.0
 * Description: Modular Google Apps Script for formatting Weekly Report Google Docs.
 *              Applies custom typography, weighted font styles, table alignment rules,
 *              column widths, and automatic totals rows with cell merging.
 */

// =============================================================================
// 1. Configuration & Constants
// =============================================================================
const DOC_STYLE_CONFIG = Object.freeze({
  HEADING1_STYLE: { fontFamily: "Lora", weight: 600, bold: false },
  HEADING2_STYLE: { fontFamily: "Lora", weight: 500, bold: false },
  HEADING3_STYLE: { fontFamily: "Lora", weight: 400, bold: false },
  TABLE_NAME_STYLE: { fontFamily: "Arial", weight: 400, bold: false },
  SIGN_OFF_NAMES: [
    "Dawit Michael (District Manager)",
    "Lelisa Dessalen (South Addis, Bishoftu, Wolayta Team Leader)",
    "Yohannes Mulu (Central Addis, Hawassa, Shashemene Team Leader)",
  ],
  SIGN_OFF_STYLE: {
    fontFamily: "Lora",
    nameWeight: 700,
    parenWeight: 500,
    nameBold: true,
    parenBold: false,
  },
  OVERVIEW_CLOSING_TEXT:
    "This overview offers a clear snapshot of the week's performance, with detailed reports provided in the following pages.",
  OVERVIEW_CLOSING_STYLE: { fontFamily: "Lora", weight: 600, bold: false },
  FIRST_COLUMN_WIDTH_POINTS: 28.8,
  FIRST_COLUMN_WIDTH_TABLE_TITLES: [
    "District Summary",
    "Case Summary by District",
    "Cases Closed After Reg. Date",
    "Ongoing Cases",
    "Bank Specific Case Count",
    "Registration Type Counts",
    "PM Summary",
    "Bank Specific PM Count",
    "Technicians Weekly Activity",
    "PM Plans",
    "PM Plan",
    "Per Diem Cost",
    "Changed Spare Parts",
  ],
  DYNAMIC_COLUMN_WIDTH_FROM_INDEX_2: {
    tableNames: new Set(["Bank Specific Case Count", "Bank Specific PM Count"]),
    startColumnIndex: 2,
    widthPoints: 54, // 0.75 in × 72 pt/in
  },
  TECHNICIANS_WEEKLY_ACTIVITY_COLUMN_WIDTHS: [
    { columnIndex: 2, widthPoints: 60.768 },
    { columnIndex: 3, widthPoints: 63 },
    { columnIndex: 4, widthPoints: 56.232 },
    { columnIndex: 5, widthPoints: 48.744 },
    { columnIndex: 6, widthPoints: 52.488 },
  ],
  CASES_CLOSED_AFTER_REG_DATE_COLUMN_WIDTHS: [
    { columnIndex: 1, widthPoints: 57.6 }, // 0.8 in × 72 pt/in
    { columnIndex: 3, widthPoints: 79.2 }, // 1.1 in × 72 pt/in
    { columnIndex: 4, widthPoints: 64.8 }, // 0.9 in × 72 pt/in
  ],
  TABLE_RULES: {
    "District Summary": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "Case Summary by District": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "Bank Specific Case Count": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "Registration Type Counts": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "PM Summary": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "Bank Specific PM Count": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "Technicians Weekly Activity": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "PM Plans": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "PM Plan": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "Cases Closed After Reg. Date": "FIRST_TWO_CENTER_REST_LEFT",
    "Ongoing Cases": "FIRST_TWO_CENTER_REST_LEFT",
    "Weekly Tasks": "ALL_LEFT",
    "Weekly Meeting": "ALL_LEFT",
    "Weekly Meeting Overview": "ALL_LEFT",
    "Meetings": "ALL_LEFT",
    "Changed Spare Parts": "SECOND_CENTER_REST_LEFT",
    "Per Diem Cost": "FIRST_AND_LAST_CENTER_REST_LEFT",
  },
  TOTAL_ROW_TABLE_TITLES: ["Per Diem Cost"],
});

// =============================================================================
// 2. Menu Entry Points & Installation Helpers
// =============================================================================
function onOpen(e) {
  try {
    const ui = DocumentApp.getUi();
    ui.createMenu("Report Formatting")
      .addItem("Format Report", "applyWeeklyReportDocFormatting")
      .addToUi();
  } catch (err) {
    Logger.log("[onOpen Error]: " + err.message);
  }
}

function onInstall(e) {
  onOpen(e);
}

function createCustomMenu() {
  onOpen(null);
}

function applyWeeklyReportDocFormatting() {
  const ui = DocumentApp.getUi();

  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const docId = doc.getId();

    TypographyStyler.applyStyles(docId);
    TableStyler.applyAlignmentAndTotals(body);

    doc.saveAndClose();
    ui.alert("Weekly report formatting applied successfully.");
  } catch (error) {
    ui.alert(`Formatting failed: ${error.message}`);
    throw error;
  }
}

// =============================================================================
// 3. Typography & Advanced Docs API Styler
// =============================================================================
const TypographyStyler = {
  applyStyles(docId) {
    if (typeof Docs === "undefined" || !Docs.Documents) {
      throw new Error(
        "Enable Advanced Google service 'Docs API' for weighted font styles."
      );
    }

    const documentData = Docs.Documents.get(docId);
    const requests = []
      .concat(this.buildTypographyRequests(documentData))
      .concat(this.buildFirstColumnWidthRequests(documentData));

    if (requests.length === 0) {
      return;
    }

    Docs.Documents.batchUpdate({ requests }, docId);
  },

  buildTypographyRequests(documentData) {
    const requests = [];
    const content =
      documentData.body && documentData.body.content
        ? documentData.body.content
        : [];
    const overviewClosing = StringUtil.normalizeWhitespace(
      DOC_STYLE_CONFIG.OVERVIEW_CLOSING_TEXT
    );
    const signOffNames = new Set(
      DOC_STYLE_CONFIG.SIGN_OFF_NAMES.map((name) =>
        StringUtil.normalizeWhitespace(name)
      )
    );

    for (let i = 0; i < content.length; i += 1) {
      const element = content[i];
      if (!element.paragraph) {
        continue;
      }

      const startIndex = element.startIndex;
      const endIndex = element.endIndex - 1;
      if (
        typeof startIndex !== "number" ||
        typeof endIndex !== "number" ||
        endIndex <= startIndex
      ) {
        continue;
      }

      const paragraph = element.paragraph;
      const paragraphText = this.getParagraphText(paragraph);
      const paragraphTextTrimmed = paragraphText.trim();
      const normalizedText = StringUtil.normalizeWhitespace(paragraphText);
      const namedStyleType = paragraph.paragraphStyle
        ? paragraph.paragraphStyle.namedStyleType
        : "";

      if (namedStyleType === "HEADING_1") {
        requests.push(
          this.createTextStyleRequest(
            startIndex,
            endIndex,
            DOC_STYLE_CONFIG.HEADING1_STYLE
          )
        );
      }

      if (namedStyleType === "HEADING_2") {
        requests.push(
          this.createTextStyleRequest(
            startIndex,
            endIndex,
            DOC_STYLE_CONFIG.HEADING2_STYLE
          )
        );
      }

      if (namedStyleType === "HEADING_3") {
        requests.push(
          this.createTextStyleRequest(
            startIndex,
            endIndex,
            DOC_STYLE_CONFIG.HEADING3_STYLE
          )
        );
      }

      if (/^Table\s+\d+\s*:/.test(paragraphTextTrimmed)) {
        requests.push(
          this.createTextStyleRequest(
            startIndex,
            endIndex,
            DOC_STYLE_CONFIG.TABLE_NAME_STYLE
          )
        );
      }

      if (normalizedText === overviewClosing) {
        requests.push(
          this.createTextStyleRequest(
            startIndex,
            endIndex,
            DOC_STYLE_CONFIG.OVERVIEW_CLOSING_STYLE
          )
        );
      }

      if (signOffNames.has(normalizedText)) {
        const signOffRequests = this.buildSignOffStyleRequests(
          startIndex,
          paragraphTextTrimmed,
          DOC_STYLE_CONFIG.SIGN_OFF_STYLE
        );
        requests.push.apply(requests, signOffRequests);
      }
    }

    return requests;
  },

  buildFirstColumnWidthRequests(documentData) {
    const requests = [];
    const content =
      documentData.body && documentData.body.content
        ? documentData.body.content
        : [];
    const targetTitles = new Set(
      DOC_STYLE_CONFIG.FIRST_COLUMN_WIDTH_TABLE_TITLES
    );
    let shouldCaptureNextTable = false;
    let capturedTableTitle = "";

    for (let i = 0; i < content.length; i += 1) {
      const element = content[i];

      if (element.paragraph) {
        const paragraph = element.paragraph;
        const namedStyleType = paragraph.paragraphStyle
          ? paragraph.paragraphStyle.namedStyleType
          : "";

        if (namedStyleType === "HEADING_2") {
          const title = StringUtil.normalizeWhitespace(
            this.getParagraphText(paragraph)
          );
          shouldCaptureNextTable = targetTitles.has(title);
          capturedTableTitle = title;
        }

        continue;
      }

      if (element.table) {
        if (shouldCaptureNextTable && typeof element.startIndex === "number") {
          requests.push(
            this.createTableColumnWidthRequest(
              element.startIndex,
              0,
              DOC_STYLE_CONFIG.FIRST_COLUMN_WIDTH_POINTS
            )
          );

          if (capturedTableTitle === "Technicians Weekly Activity") {
            DOC_STYLE_CONFIG.TECHNICIANS_WEEKLY_ACTIVITY_COLUMN_WIDTHS.forEach(
              (columnWidth) => {
                requests.push(
                  this.createTableColumnWidthRequest(
                    element.startIndex,
                    columnWidth.columnIndex,
                    columnWidth.widthPoints
                  )
                );
              }
            );
          }

          if (capturedTableTitle === "Cases Closed After Reg. Date") {
            DOC_STYLE_CONFIG.CASES_CLOSED_AFTER_REG_DATE_COLUMN_WIDTHS.forEach(
              (columnWidth) => {
                requests.push(
                  this.createTableColumnWidthRequest(
                    element.startIndex,
                    columnWidth.columnIndex,
                    columnWidth.widthPoints
                  )
                );
              }
            );
          }

          const dynCfg = DOC_STYLE_CONFIG.DYNAMIC_COLUMN_WIDTH_FROM_INDEX_2;
          if (dynCfg.tableNames.has(capturedTableTitle)) {
            const columnCount = this.getTableColumnCount(element.table);
            for (
              let col = dynCfg.startColumnIndex;
              col < columnCount;
              col += 1
            ) {
              requests.push(
                this.createTableColumnWidthRequest(
                  element.startIndex,
                  col,
                  dynCfg.widthPoints
                )
              );
            }
          }
        }
        shouldCaptureNextTable = false;
        capturedTableTitle = "";
      }
    }

    return requests;
  },

  getParagraphText(paragraph) {
    const elements = paragraph.elements || [];
    let text = "";

    for (let i = 0; i < elements.length; i += 1) {
      const paragraphElement = elements[i];
      if (paragraphElement.textRun && paragraphElement.textRun.content) {
        text += paragraphElement.textRun.content;
      }
    }

    return text.replace(/\n$/, "");
  },

  getTableColumnCount(table) {
    const rows = table.tableRows || [];
    if (rows.length === 0) {
      return 0;
    }
    return (rows[0].tableCells || []).length;
  },

  createTextStyleRequest(startIndex, endIndex, style) {
    return {
      updateTextStyle: {
        range: { startIndex, endIndex },
        textStyle: {
          weightedFontFamily: {
            fontFamily: style.fontFamily,
            weight: style.weight,
          },
          bold: style.bold,
        },
        fields: "weightedFontFamily,bold",
      },
    };
  },

  createTableColumnWidthRequest(tableStartIndex, columnIndex, widthPoints) {
    return {
      updateTableColumnProperties: {
        tableStartLocation: { index: tableStartIndex },
        columnIndices: [columnIndex],
        tableColumnProperties: {
          widthType: "FIXED_WIDTH",
          width: { magnitude: widthPoints, unit: "PT" },
        },
        fields: "width,widthType",
      },
    };
  },

  buildSignOffStyleRequests(startIndex, paragraphText, styleConfig) {
    const requests = [];
    const parenIndex = paragraphText.indexOf(" (");
    const endIndex = startIndex + paragraphText.length;

    if (parenIndex > 0) {
      const nameEnd = startIndex + parenIndex;
      requests.push(
        this.createWeightedFontRequest(startIndex, nameEnd, {
          fontFamily: styleConfig.fontFamily,
          weight: styleConfig.nameWeight,
          bold: styleConfig.nameBold,
          italic: false,
        })
      );
      requests.push(
        this.createWeightedFontRequest(nameEnd, endIndex, {
          fontFamily: styleConfig.fontFamily,
          weight: styleConfig.parenWeight,
          bold: styleConfig.parenBold,
          italic: false,
        })
      );
      return requests;
    }

    requests.push(
      this.createWeightedFontRequest(startIndex, endIndex, {
        fontFamily: styleConfig.fontFamily,
        weight: styleConfig.nameWeight,
        bold: styleConfig.nameBold,
        italic: false,
      })
    );
    return requests;
  },

  createWeightedFontRequest(startIndex, endIndex, style) {
    return {
      updateTextStyle: {
        range: { startIndex, endIndex },
        textStyle: {
          weightedFontFamily: {
            fontFamily: style.fontFamily,
            weight: style.weight,
          },
          bold: style.bold,
          italic: style.italic,
        },
        fields: "weightedFontFamily,bold,italic",
      },
    };
  },
};

// =============================================================================
// 4. Table Styler & Alignment Engine
// =============================================================================
const TableStyler = {
  applyAlignmentAndTotals(body) {
    const totalChildren = body.getNumChildren();
    const totalRowTitles = new Set(DOC_STYLE_CONFIG.TOTAL_ROW_TABLE_TITLES);

    for (let i = 0; i < totalChildren; i += 1) {
      const element = body.getChild(i);
      if (element.getType() !== DocumentApp.ElementType.PARAGRAPH) {
        continue;
      }

      const paragraph = element.asParagraph();
      if (paragraph.getHeading() !== DocumentApp.ParagraphHeading.HEADING2) {
        continue;
      }

      const tableTitle = paragraph.getText().trim();
      const ruleType = DOC_STYLE_CONFIG.TABLE_RULES[tableTitle];
      if (!ruleType) {
        continue;
      }

      const table = this.findNextTableAfter(body, i);
      if (!table) {
        continue;
      }

      this.styleTableByRule(table, ruleType);

      if (totalRowTitles.has(tableTitle)) {
        PerDiemCalculator.appendTotalsRow(table);
      }
    }
  },

  findNextTableAfter(body, headingIndex) {
    const totalChildren = body.getNumChildren();

    for (let i = headingIndex + 1; i < totalChildren; i += 1) {
      const element = body.getChild(i);
      const type = element.getType();

      if (type === DocumentApp.ElementType.TABLE) {
        return element.asTable();
      }

      if (type === DocumentApp.ElementType.PARAGRAPH) {
        const paragraph = element.asParagraph();
        if (
          paragraph.getHeading() === DocumentApp.ParagraphHeading.HEADING2 &&
          paragraph.getText().trim() !== ""
        ) {
          return null;
        }
      }
    }

    return null;
  },

  styleTableByRule(table, ruleType) {
    const rowCount = table.getNumRows();

    for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
      const row = table.getRow(rowIndex);
      const colCount = row.getNumCells();

      for (let colIndex = 0; colIndex < colCount; colIndex += 1) {
        const cell = table.getCell(rowIndex, colIndex);
        cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

        const horizontalAlignment = this.getHorizontalAlignmentForRule(
          ruleType,
          colIndex,
          colCount
        );
        this.setCellHorizontalAlignment(cell, horizontalAlignment);
      }
    }
  },

  getHorizontalAlignmentForRule(ruleType, colIndex, colCount) {
    switch (ruleType) {
      case "CENTER_ALL_EXCEPT_SECOND_LEFT":
        return colIndex === 1
          ? DocumentApp.HorizontalAlignment.LEFT
          : DocumentApp.HorizontalAlignment.CENTER;

      case "FIRST_TWO_CENTER_REST_LEFT":
        return colIndex <= 1
          ? DocumentApp.HorizontalAlignment.CENTER
          : DocumentApp.HorizontalAlignment.LEFT;

      case "ALL_LEFT":
        return DocumentApp.HorizontalAlignment.LEFT;

      case "SECOND_CENTER_REST_LEFT":
        return colIndex === 1
          ? DocumentApp.HorizontalAlignment.CENTER
          : DocumentApp.HorizontalAlignment.LEFT;

      case "FIRST_AND_LAST_CENTER_REST_LEFT":
        return colIndex === 0 || colIndex === colCount - 1
          ? DocumentApp.HorizontalAlignment.CENTER
          : DocumentApp.HorizontalAlignment.LEFT;

      default:
        return DocumentApp.HorizontalAlignment.LEFT;
    }
  },

  setCellHorizontalAlignment(cell, horizontalAlignment) {
    const childCount = cell.getNumChildren();

    for (let i = 0; i < childCount; i += 1) {
      const child = cell.getChild(i);
      const childType = child.getType();

      if (childType === DocumentApp.ElementType.PARAGRAPH) {
        child.asParagraph().setAlignment(horizontalAlignment);
      } else if (childType === DocumentApp.ElementType.LIST_ITEM) {
        child.asListItem().setAlignment(horizontalAlignment);
      }
    }
  },
};

// =============================================================================
// 5. Per Diem Totals Calculator
// =============================================================================
const PerDiemCalculator = {
  appendTotalsRow(table) {
    this.removeExistingTotalsRow(table);

    const rowCount = table.getNumRows();
    const colCount = rowCount > 0 ? table.getRow(0).getNumCells() : 0;
    if (colCount < 2) {
      return;
    }

    const amountColumnIndex = colCount - 1;
    let total = 0;
    for (let rowIndex = 1; rowIndex < rowCount; rowIndex += 1) {
      const value = this.parseAmount(
        table.getCell(rowIndex, amountColumnIndex).getText()
      );
      if (!isNaN(value)) {
        total += value;
      }
    }

    const totalRow = table.appendTableRow();

    for (let colIndex = 0; colIndex < colCount; colIndex += 1) {
      const cellText =
        colIndex === 0
          ? "Total"
          : colIndex === amountColumnIndex
          ? this.formatAmount(total)
          : "";
      const cell = totalRow.appendTableCell(cellText);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      cell.editAsText().setBold(
        colIndex === 0 || colIndex === amountColumnIndex
      );
    }

    // Merge cells 0 through 3 (first 4 columns) or 0 through colCount - 2
    const currentCells = totalRow.getNumCells();
    if (currentCells >= 5) {
      try {
        const mergedCell = totalRow.getCell(0).merge(totalRow.getCell(3));
        mergedCell.setText("Total");
        mergedCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        mergedCell.editAsText().setBold(true);
        TableStyler.setCellHorizontalAlignment(
          mergedCell,
          DocumentApp.HorizontalAlignment.CENTER
        );
      } catch (err) {
        Logger.log("[PerDiem Merge Error]: " + err.message);
      }
    } else if (currentCells > 2) {
      try {
        const endMergeIndex = currentCells - 2;
        const mergedCell = totalRow.getCell(0).merge(totalRow.getCell(endMergeIndex));
        mergedCell.setText("Total");
        mergedCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        mergedCell.editAsText().setBold(true);
        TableStyler.setCellHorizontalAlignment(
          mergedCell,
          DocumentApp.HorizontalAlignment.CENTER
        );
      } catch (err) {
        Logger.log("[PerDiem Merge Error]: " + err.message);
      }
    }

    // Amount cell (last cell)
    const finalCellCount = totalRow.getNumCells();
    if (finalCellCount > 0) {
      const amountCell = totalRow.getCell(finalCellCount - 1);
      amountCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      amountCell.editAsText().setBold(true);
      TableStyler.setCellHorizontalAlignment(
        amountCell,
        DocumentApp.HorizontalAlignment.CENTER
      );
    }
  },

  parseAmount(value) {
    const numericText = String(value || "").replace(/[^0-9.-]/g, "");
    if (numericText === "" || numericText === "-" || numericText === ".") {
      return NaN;
    }

    return parseFloat(numericText);
  },

  formatAmount(value) {
    const roundedValue = Math.round((value + Number.EPSILON) * 100) / 100;
    const parts = roundedValue.toFixed(2).split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return parts.join(".");
  },

  removeExistingTotalsRow(table) {
    const lastRowIndex = table.getNumRows() - 1;
    if (lastRowIndex <= 0) {
      return;
    }

    const lastRow = table.getRow(lastRowIndex);
    if (lastRow.getNumCells() === 0) {
      return;
    }

    const firstCellText = StringUtil.normalizeWhitespace(
      lastRow.getCell(0).getText()
    );
    if (firstCellText.toLowerCase() === "total") {
      lastRow.removeFromParent();
    }
  },
};

// =============================================================================
// 6. Utility Helpers
// =============================================================================
const StringUtil = {
  normalizeWhitespace(value) {
    return String(value || "")
      .replace(/\s+/g, " ")
      .trim();
  },
};
