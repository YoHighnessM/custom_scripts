const DOC_STYLE_CONFIG = {
  HEADING1_STYLE: {
    fontFamily: "Lora",
    weight: 600,
    bold: false,
  },
  HEADING2_STYLE: {
    fontFamily: "Lora",
    weight: 500,
    bold: false,
  },
  HEADING3_STYLE: {
    fontFamily: "Lora",
    weight: 400,
    bold: false,
  },
  TABLE_NAME_STYLE: {
    fontFamily: "Arial",
    weight: 400,
    bold: false,
  },
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
  OVERVIEW_CLOSING_STYLE: {
    fontFamily: "Lora",
    weight: 600,
    bold: false,
  },
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
    "Per Diem Cost",
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
  TABLE_RULES: {
    "District Summary": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "Case Summary by District": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "Bank Specific Case Count": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "Registration Type Counts": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "PM Summary": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "Bank Specific PM Count": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "Technicians Weekly Activity": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "PM Plans": "CENTER_ALL_EXCEPT_SECOND_LEFT",
    "Cases Closed After Reg. Date": "FIRST_TWO_CENTER_REST_LEFT",
    "Ongoing Cases": "FIRST_TWO_CENTER_REST_LEFT",
    "Weekly Tasks": "ALL_LEFT",
    "Weekly Meeting Overview": "ALL_LEFT",
    "Changed Spare Parts": "SECOND_CENTER_REST_LEFT",
    "Per Diem Cost": "FIRST_AND_LAST_CENTER_REST_LEFT",
  },
  TOTAL_ROW_TABLE_TITLES: ["Per Diem Cost"],
};

function onOpen() {
  DocumentApp.getUi()
    .createMenu("Report Formatting")
    .addItem("Format Report", "applyWeeklyReportDocFormatting")
    .addToUi();
}

function applyWeeklyReportDocFormatting() {
  const ui = DocumentApp.getUi();

  try {
    const doc = DocumentApp.getActiveDocument();
    const body = doc.getBody();
    const docId = doc.getId();

    applyTypographyStyles(docId);
    applyTableAlignmentRules(body);

    doc.saveAndClose();
    ui.alert("Weekly report formatting applied successfully.");
  } catch (error) {
    ui.alert(`Formatting failed: ${error.message}`);
    throw error;
  }
}

function applyTypographyStyles(docId) {
  if (typeof Docs === "undefined" || !Docs.Documents) {
    throw new Error(
      "Enable Advanced Google service 'Docs API' for weighted font styles.",
    );
  }

  const documentData = Docs.Documents.get(docId);
  const requests = []
    .concat(buildTypographyRequests(documentData))
    .concat(buildFirstColumnWidthRequests(documentData));
  if (requests.length === 0) {
    return;
  }

  Docs.Documents.batchUpdate({ requests }, docId);
}

function buildTypographyRequests(documentData) {
  const requests = [];
  const content =
    documentData.body && documentData.body.content
      ? documentData.body.content
      : [];
  const overviewClosing = normalizeWhitespace(
    DOC_STYLE_CONFIG.OVERVIEW_CLOSING_TEXT,
  );
  const signOffNames = new Set(
    DOC_STYLE_CONFIG.SIGN_OFF_NAMES.map((name) => normalizeWhitespace(name)),
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
    const paragraphText = getParagraphText(paragraph);
    const paragraphTextTrimmed = paragraphText.trim();
    const normalizedText = normalizeWhitespace(paragraphText);
    const namedStyleType = paragraph.paragraphStyle
      ? paragraph.paragraphStyle.namedStyleType
      : "";

    if (namedStyleType === "HEADING_1") {
      requests.push(
        createTextStyleRequest(
          startIndex,
          endIndex,
          DOC_STYLE_CONFIG.HEADING1_STYLE,
        ),
      );
    }

    if (namedStyleType === "HEADING_2") {
      requests.push(
        createTextStyleRequest(
          startIndex,
          endIndex,
          DOC_STYLE_CONFIG.HEADING2_STYLE,
        ),
      );
    }

    if (namedStyleType === "HEADING_3") {
      requests.push(
        createTextStyleRequest(
          startIndex,
          endIndex,
          DOC_STYLE_CONFIG.HEADING3_STYLE,
        ),
      );
    }

    if (/^Table\s+\d+\s*:/.test(paragraphTextTrimmed)) {
      requests.push(
        createTextStyleRequest(
          startIndex,
          endIndex,
          DOC_STYLE_CONFIG.TABLE_NAME_STYLE,
        ),
      );
    }

    if (normalizedText === overviewClosing) {
      requests.push(
        createTextStyleRequest(
          startIndex,
          endIndex,
          DOC_STYLE_CONFIG.OVERVIEW_CLOSING_STYLE,
        ),
      );
    }

    if (signOffNames.has(normalizedText)) {
      const signOffRequests = buildSignOffStyleRequests(
        startIndex,
        paragraphTextTrimmed,
        DOC_STYLE_CONFIG.SIGN_OFF_STYLE,
      );
      requests.push.apply(requests, signOffRequests);
    }
  }

  return requests;
}

function buildFirstColumnWidthRequests(documentData) {
  const requests = [];
  const content =
    documentData.body && documentData.body.content
      ? documentData.body.content
      : [];
  const targetTitles = new Set(
    DOC_STYLE_CONFIG.FIRST_COLUMN_WIDTH_TABLE_TITLES,
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
        const title = normalizeWhitespace(getParagraphText(paragraph));
        shouldCaptureNextTable = targetTitles.has(title);
        capturedTableTitle = title;
      }

      continue;
    }

    if (element.table) {
      if (shouldCaptureNextTable && typeof element.startIndex === "number") {
        requests.push(
          createFirstColumnWidthRequest(
            element.startIndex,
            DOC_STYLE_CONFIG.FIRST_COLUMN_WIDTH_POINTS,
          ),
        );

        if (capturedTableTitle === "Technicians Weekly Activity") {
          DOC_STYLE_CONFIG.TECHNICIANS_WEEKLY_ACTIVITY_COLUMN_WIDTHS.forEach(
            (columnWidth) => {
              requests.push(
                createTableColumnWidthRequest(
                  element.startIndex,
                  columnWidth.columnIndex,
                  columnWidth.widthPoints,
                ),
              );
            },
          );
        }

        const dynCfg = DOC_STYLE_CONFIG.DYNAMIC_COLUMN_WIDTH_FROM_INDEX_2;
        if (dynCfg.tableNames.has(capturedTableTitle)) {
          const columnCount = getTableColumnCount(element.table);
          for (let col = dynCfg.startColumnIndex; col < columnCount; col += 1) {
            requests.push(
              createTableColumnWidthRequest(
                element.startIndex,
                col,
                dynCfg.widthPoints,
              ),
            );
          }
        }
      }
      shouldCaptureNextTable = false;
      capturedTableTitle = "";
    }
  }

  return requests;
}

function getParagraphText(paragraph) {
  const elements = paragraph.elements || [];
  let text = "";

  for (let i = 0; i < elements.length; i += 1) {
    const paragraphElement = elements[i];
    if (paragraphElement.textRun && paragraphElement.textRun.content) {
      text += paragraphElement.textRun.content;
    }
  }

  return text.replace(/\n$/, "");
}

function createTextStyleRequest(startIndex, endIndex, style) {
  return {
    updateTextStyle: {
      range: {
        startIndex,
        endIndex,
      },
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
}

function getTableColumnCount(table) {
  const rows = table.tableRows || [];
  if (rows.length === 0) {
    return 0;
  }
  return (rows[0].tableCells || []).length;
}

function createFirstColumnWidthRequest(tableStartIndex, widthPoints) {
  return createTableColumnWidthRequest(tableStartIndex, 0, widthPoints);
}

function createTableColumnWidthRequest(
  tableStartIndex,
  columnIndex,
  widthPoints,
) {
  return {
    updateTableColumnProperties: {
      tableStartLocation: {
        index: tableStartIndex,
      },
      columnIndices: [columnIndex],
      tableColumnProperties: {
        widthType: "FIXED_WIDTH",
        width: {
          magnitude: widthPoints,
          unit: "PT",
        },
      },
      fields: "width,widthType",
    },
  };
}

function buildSignOffStyleRequests(startIndex, paragraphText, styleConfig) {
  const requests = [];
  const parenIndex = paragraphText.indexOf(" (");
  const endIndex = startIndex + paragraphText.length;

  if (parenIndex > 0) {
    const nameEnd = startIndex + parenIndex;
    requests.push(
      createWeightedFontRequest(startIndex, nameEnd, {
        fontFamily: styleConfig.fontFamily,
        weight: styleConfig.nameWeight,
        bold: styleConfig.nameBold,
        italic: false,
      }),
    );
    requests.push(
      createWeightedFontRequest(nameEnd, endIndex, {
        fontFamily: styleConfig.fontFamily,
        weight: styleConfig.parenWeight,
        bold: styleConfig.parenBold,
        italic: false,
      }),
    );
    return requests;
  }

  requests.push(
    createWeightedFontRequest(startIndex, endIndex, {
      fontFamily: styleConfig.fontFamily,
      weight: styleConfig.nameWeight,
      bold: styleConfig.nameBold,
      italic: false,
    }),
  );
  return requests;
}

function createWeightedFontRequest(startIndex, endIndex, style) {
  return {
    updateTextStyle: {
      range: {
        startIndex,
        endIndex,
      },
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
}

function applyTableAlignmentRules(body) {
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

    const table = findNextTableAfter(body, i);
    if (!table) {
      continue;
    }

    styleTableByRule(table, ruleType);

    if (totalRowTitles.has(tableTitle)) {
      appendPerDiemTotalsRow(table);
    }
  }
}

function findNextTableAfter(body, headingIndex) {
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
}

function styleTableByRule(table, ruleType) {
  const rowCount = table.getNumRows();

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const row = table.getRow(rowIndex);
    const colCount = row.getNumCells();

    for (let colIndex = 0; colIndex < colCount; colIndex += 1) {
      const cell = table.getCell(rowIndex, colIndex);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

      const horizontalAlignment = getHorizontalAlignmentForRule(
        ruleType,
        colIndex,
        colCount,
      );
      setCellHorizontalAlignment(cell, horizontalAlignment);
    }
  }
}

function getHorizontalAlignmentForRule(ruleType, colIndex, colCount) {
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
}

function setCellHorizontalAlignment(cell, horizontalAlignment) {
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
}

function appendPerDiemTotalsRow(table) {
  removeExistingPerDiemTotalsRow(table);

  const rowCount = table.getNumRows();
  const colCount = rowCount > 0 ? table.getRow(0).getNumCells() : 0;
  if (colCount < 2) {
    return;
  }

  const amountColumnIndex = colCount - 1;
  let total = 0;
  for (let rowIndex = 1; rowIndex < rowCount; rowIndex += 1) {
    const value = parseAmount(
      table.getCell(rowIndex, amountColumnIndex).getText(),
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
          ? formatAmount(total)
          : "";
    const cell = totalRow.appendTableCell(cellText);
    cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    cell.editAsText().setBold(
      colIndex === 0 || colIndex === amountColumnIndex,
    );
    setCellHorizontalAlignment(
      cell,
      getHorizontalAlignmentForRule(
        "FIRST_AND_LAST_CENTER_REST_LEFT",
        colIndex,
        colCount,
      ),
    );
  }
}

function parseAmount(value) {
  const numericText = String(value || "").replace(/[^0-9.-]/g, "");
  if (numericText === "" || numericText === "-" || numericText === ".") {
    return NaN;
  }

  return parseFloat(numericText);
}

function formatAmount(value) {
  const roundedValue = Math.round((value + Number.EPSILON) * 100) / 100;
  const parts = roundedValue.toFixed(2).split(".");
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return parts.join(".");
}

function removeExistingPerDiemTotalsRow(table) {
  const lastRowIndex = table.getNumRows() - 1;
  if (lastRowIndex <= 0) {
    return;
  }

  const lastRow = table.getRow(lastRowIndex);
  if (lastRow.getNumCells() === 0) {
    return;
  }

  const firstCellText = normalizeWhitespace(lastRow.getCell(0).getText());
  if (firstCellText.toLowerCase() === "total") {
    lastRow.removeFromParent();
  }
}

function normalizeWhitespace(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}
