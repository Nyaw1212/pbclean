//----------------------------------
// Constants
//----------------------------------

const SOURCE_SHEET = "Material List";
const TEMPLATE_SHEET = "template";
const PROJECTS_SHEET = "PROJECTS";

const SOURCE_START_ROW = 2;
const TEMPLATE_START_ROW = 6;

const PROJECT_JOB_CELL = "O1";
const TEMPLATE_JOB_CELL = "A2";
const TEMPLATE_ADDRESS_CELL = "A3";

function generateAllReports() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  //----------------------------------
  // Prevent Double Execution
  //----------------------------------

  const lock =
    LockService.getDocumentLock();

  if (!lock.tryLock(1000)) {
    ss.toast(
      "Another report is already running.",
      "Please Wait",
      5,
    );

    return;
  }

  try {
    //----------------------------------
    // Step 1: Refresh Material List
    //----------------------------------

    ss.toast(
      "Step 1 of 5: Updating Material List...",
      "Generating Report",
      10,
    );

    populateMaterialList();

    SpreadsheetApp.flush();

    //----------------------------------
    // Step 2: Clean Previous Report
    //----------------------------------

    ss.toast(
      "Step 2 of 5: Cleaning previous report...",
      "Generating Report",
      10,
    );

    cleanPreviousTable(false);

    SpreadsheetApp.flush();

    //----------------------------------
    // Step 3: Populate Project Info
    //----------------------------------

    ss.toast(
      "Step 3 of 5: Loading project information...",
      "Generating Report",
      10,
    );

    const projectFound =
      populateTemplateProjectInfo(false);

    if (!projectFound) {
      SpreadsheetApp
        .getUi()
        .alert(
          "The selected project was not found.",
        );

      return;
    }

    //----------------------------------
    // Step 4: Generate BOS Report
    //----------------------------------

    ss.toast(
      "Step 4 of 5: Generating BOS report...",
      "Generating Report",
      10,
    );

    generateBOSReport(false);

    SpreadsheetApp.flush();

    //----------------------------------
    // Step 5: Generate VAN Report
    //----------------------------------

    ss.toast(
      "Step 5 of 5: Generating VAN report...",
      "Generating Report",
      10,
    );

    generateVANReport(false);

    SpreadsheetApp.flush();

    //----------------------------------
    // Update Grand Total
    //----------------------------------

    updateGrandTotal(false);

    ss.toast(
      "Report completed successfully.",
      "Materials Report",
      5,
    );

    SpreadsheetApp
      .getUi()
      .alert(
        "Material List updated.\n\nBOS and VAN material reports generated successfully!",
      );
  } finally {
    lock.releaseLock();
  }
}

//----------------------------------
// Clean Previous Table
//----------------------------------

function cleanPreviousTable(
  showAlert = true,
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const template =
    ss.getSheetByName(
      TEMPLATE_SHEET,
    );

  if (!template) {
    throw new Error(
      `Sheet not found: ${TEMPLATE_SHEET}`,
    );
  }

  //----------------------------------
  // Clean Only Used Report Rows
  //----------------------------------

  const lastUsedRow =
    Math.max(
      template.getLastRow(),
      TEMPLATE_START_ROW,
    );

  const availableRows =
    lastUsedRow -
    TEMPLATE_START_ROW +
    1;

  const bosRange =
    template.getRange(
      TEMPLATE_START_ROW,
      1,
      availableRows,
      5,
    );

  const vanRange =
    template.getRange(
      TEMPLATE_START_ROW,
      7,
      availableRows,
      5,
    );

  //----------------------------------
  // Remove Merges
  //----------------------------------

  bosRange.breakApart();
  vanRange.breakApart();

  //----------------------------------
  // Clear Contents
  //----------------------------------

  bosRange.clearContent();
  vanRange.clearContent();

  //----------------------------------
  // Clear Formatting
  //----------------------------------

  bosRange
    .setBackground(null)
    .setFontColor("#000000")
    .setFontWeight("normal")
    .setFontStyle("normal")
    .setFontLine("none")
    .setShowHyperlink(false)
    .setHorizontalAlignment(null)
    .setVerticalAlignment(null)
    .setNumberFormat("General")
    .setBorder(
      false,
      false,
      false,
      false,
      false,
      false,
    );

  vanRange
    .setBackground(null)
    .setFontColor("#000000")
    .setFontWeight("normal")
    .setFontStyle("normal")
    .setFontLine("none")
    .setShowHyperlink(false)
    .setHorizontalAlignment(null)
    .setVerticalAlignment(null)
    .setNumberFormat("General")
    .setBorder(
      false,
      false,
      false,
      false,
      false,
      false,
    );

  //----------------------------------
  // Clear Grand Total
  //----------------------------------

  template
    .getRange("F1")
    .clearContent();

  if (showAlert) {
    SpreadsheetApp
      .getUi()
      .alert(
        "Previous report tables and formatting cleared successfully!",
      );
  }

  return true;
}

//----------------------------------
// Populate Template Project Info
//----------------------------------

function populateTemplateProjectInfo(
  showAlert = true,
) {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const projectsSheet =
    ss.getSheetByName(
      PROJECTS_SHEET,
    );

  const template =
    ss.getSheetByName(
      TEMPLATE_SHEET,
    );

  if (!projectsSheet) {
    throw new Error(
      `Sheet not found: ${PROJECTS_SHEET}`,
    );
  }

  if (!template) {
    throw new Error(
      `Sheet not found: ${TEMPLATE_SHEET}`,
    );
  }

  //----------------------------------
  // Read JobName from template!O1
  //----------------------------------

  const selectedJob =
    normalizeProjectText(
      template
        .getRange(
          PROJECT_JOB_CELL,
        )
        .getDisplayValue(),
    );

  if (!selectedJob) {
    template
      .getRange(
        `${TEMPLATE_JOB_CELL}:${TEMPLATE_ADDRESS_CELL}`,
      )
      .clearContent();

    if (showAlert) {
      SpreadsheetApp
        .getUi()
        .alert(
          `Select a JobName in ${TEMPLATE_SHEET}!${PROJECT_JOB_CELL}.`,
        );
    }

    return false;
  }

  //----------------------------------
  // PROJECTS:
  // A = ProjectID
  // B = JobName
  // C = Address
  //----------------------------------

  const lastRow =
    projectsSheet.getLastRow();

  if (lastRow < 2) {
    throw new Error(
      `${PROJECTS_SHEET} has no project records.`,
    );
  }

  const projectData =
    projectsSheet
      .getRange(
        2,
        1,
        lastRow - 1,
        3,
      )
      .getDisplayValues();

  //----------------------------------
  // Find JobName in Column B
  //----------------------------------

  const project =
    projectData.find(
      (row) =>
        normalizeProjectText(
          row[1],
        ) === selectedJob,
    );

  if (!project) {
    template
      .getRange(
        `${TEMPLATE_JOB_CELL}:${TEMPLATE_ADDRESS_CELL}`,
      )
      .clearContent();

    if (showAlert) {
      SpreadsheetApp
        .getUi()
        .alert(
          `JobName "${template
            .getRange(PROJECT_JOB_CELL)
            .getDisplayValue()}" was not found in ${PROJECTS_SHEET} column B.`,
        );
    }

    return false;
  }

  //----------------------------------
  // Populate template
  //----------------------------------

  const jobName =
    project[1];

  const address =
    project[2];

  template
    .getRange(
      TEMPLATE_JOB_CELL,
    )
    .setValue(
      jobName,
    );

  template
    .getRange(
      TEMPLATE_ADDRESS_CELL,
    )
    .setValue(
      address,
    );

  if (showAlert) {
    SpreadsheetApp
      .getUi()
      .alert(
        `Project information updated:\n${jobName}\n${address}`,
      );
  }

  return true;
}

//----------------------------------
// Normalize Project Text
//----------------------------------

function normalizeProjectText(
  value,
) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

//----------------------------------
// Get Source Data
//----------------------------------

function getMaterialSourceData() {
  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const source =
    ss.getSheetByName(SOURCE_SHEET);

  const template =
    ss.getSheetByName(TEMPLATE_SHEET);

  if (!source) {
    throw new Error(
      `Sheet not found: ${SOURCE_SHEET}`,
    );
  }

  if (!template) {
    throw new Error(
      `Sheet not found: ${TEMPLATE_SHEET}`,
    );
  }

  const lastRow =
    source.getLastRow();

  if (lastRow < SOURCE_START_ROW) {
    return {
      source,
      template,
      data: [],
    };
  }

  const values =
    source
      .getRange(
        SOURCE_START_ROW,
        1,
        lastRow -
          SOURCE_START_ROW +
          1,
        source.getLastColumn(),
      )
      .getValues();

  //----------------------------------
  // Preserve Original Source Row
  //----------------------------------

  const data =
    values.map(
      (row, index) => ({
        row,
        sourceRow:
          SOURCE_START_ROW +
          index,
      }),
    );

  return {
    source,
    template,
    data,
  };
}


//----------------------------------
// Ensure Sheet Has Enough Rows
//----------------------------------

function ensureSheetRows_(
  sheet,
  requiredLastRow,
) {
  const currentMaxRows =
    sheet.getMaxRows();

  if (
    requiredLastRow >
    currentMaxRows
  ) {
    sheet.insertRowsAfter(
      currentMaxRows,
      requiredLastRow -
        currentMaxRows,
    );
  }
}

//----------------------------------
// Generate BOS Report
//----------------------------------

function generateBOSReport(
  showAlert = true,
) {
  const {
    source,
    template,
    data,
  } = getMaterialSourceData();

  if (data.length === 0) {
    if (showAlert) {
      SpreadsheetApp
        .getUi()
        .alert(
          "No material data found.",
        );
    }

    return;
  }

  //----------------------------------
  // BOS Bottom Keywords
  //----------------------------------

  const bottomKeywords = [
    "BOX",
    "SC",
    "WWAY",
    "WIREWAY",
    "GUTTER",
    "PAD",
    "PADS",
  ];

  //----------------------------------
  // Filter BOS Materials
  //----------------------------------

  const bosMaterials =
    data.filter((item) => {
      const row =
        item.row;

      const used =
        Number(row[9]) || 0;

      const category =
        String(
          row[12] || "",
        )
          .trim()
          .toUpperCase();

      return (
        category === "BOS" &&
        used !== 0
      );
    });

  //----------------------------------
  // Sort BOS Materials
  //----------------------------------

  bosMaterials.sort(
    (itemA, itemB) => {
      const descriptionA =
        String(
          itemA.row[5] || "",
        ).toUpperCase();

      const descriptionB =
        String(
          itemB.row[5] || "",
        ).toUpperCase();

      const isBottomA =
        bottomKeywords.some(
          (keyword) =>
            descriptionA.includes(
              keyword,
            ),
        );

      const isBottomB =
        bottomKeywords.some(
          (keyword) =>
            descriptionB.includes(
              keyword,
            ),
        );

      if (
        isBottomA === isBottomB
      ) {
        return 0;
      }

      return isBottomA ? 1 : -1;
    },
  );

  //----------------------------------
  // Clear Previous BOS Report
  //----------------------------------

  const availableRows =
    template.getMaxRows() -
    TEMPLATE_START_ROW +
    1;

  template
    .getRange(
      TEMPLATE_START_ROW,
      1,
      availableRows,
      4,
    )
    .breakApart();

  template
    .getRange(
      TEMPLATE_START_ROW,
      1,
      availableRows,
      5,
    )
    .clearContent();

  let destRow =
    TEMPLATE_START_ROW;

  //----------------------------------
  // Process BOS Materials
  //----------------------------------

  bosMaterials.forEach(
    (item) => {

      //----------------------------------
      // Add Template Rows When Required
      //----------------------------------

      ensureSheetRows_(
        template,
        destRow + 1,
      );

      const row =
        item.row;

      const catalog =
        row[3];

      const description =
        row[5];

      const price =
        row[11] || row[6];

      const used =
        row[9];

      //----------------------------------
      // Apply BOS Row Formatting
      //----------------------------------

      const descriptionUpper =
        String(description || "")
          .trim()
          .toUpperCase();

      const greenKeywords = [
        "BOX",
        "SCWAY",
        "SC WWAY",
        "WWAY",
        "WIREWAY",
        "PAD",
        "PADS",
      ];

      const isGreenItem =
        greenKeywords.some(
          (keyword) =>
            descriptionUpper.includes(keyword),
        );

      const rowColor =
        isGreenItem
          ? "#d9ead3"
          : "#ead1dc";

      template
        .getRange(
          destRow,
          1,
          1,
          5,
        )
        .setBackground(rowColor)
        .setFontColor("#000000")
        .setFontWeight("normal")
        .setFontStyle("normal")
        .setVerticalAlignment("middle")
        .setBorder(
          true,
          true,
          true,
          true,
          true,
          true,
          "#000000",
          SpreadsheetApp.BorderStyle.SOLID,
        );

      //----------------------------------
      // C = Plain White
      //----------------------------------

      template
        .getRange(
          destRow,
          3,
        )
        .setBackground(
          "#ffffff",
        )
        .setFontColor(
          "#000000",
        )
        .setFontWeight(
          "normal",
        )
        .setFontStyle(
          "normal",
        )
        .setHorizontalAlignment(
          "center",
        )
        .setVerticalAlignment(
          "middle",
        )
        .setBorder(
          true,
          true,
          true,
          true,
          true,
          true,
        );

      //----------------------------------
      // Values
      //----------------------------------

      template
        .getRange(
          destRow,
          1,
        )
        .setValue(
          description,
        )
        .setHorizontalAlignment(
          "left",
        );

      template
        .getRange(
          destRow,
          2,
        )
        .setValue(
          catalog,
        )
        .setHorizontalAlignment(
          "left",
        );

      template
        .getRange(
          destRow,
          3,
        )
        .setValue(
          used,
        );

      template
        .getRange(
          destRow,
          4,
        )
        .setValue(
          price,
        )
        .setNumberFormat(
          "$#,##0.00",
        );

      //----------------------------------
      // Cost Formula
      //----------------------------------

      template
        .getRange(
          destRow,
          5,
        )
        .setFormula(
          `=C${destRow}*D${destRow}`,
        )
        .setNumberFormat(
          "$#,##0.00",
        )
        .setFontColor("#000000")
        .setFontLine("none")
        .setShowHyperlink(false);

      destRow++;
    },
  );

  //----------------------------------
  // BOS Total
  //----------------------------------

  if (
    destRow >
    TEMPLATE_START_ROW
  ) {

    ensureSheetRows_(
      template,
      destRow,
    );

    template
      .getRange(
        destRow - 1,
        1,
        1,
        5,
      )
      .copyTo(
        template.getRange(
          destRow,
          1,
        ),
        SpreadsheetApp
          .CopyPasteType
          .PASTE_FORMAT,
        false,
      );

    template
      .getRange(
        destRow,
        1,
        1,
        4,
      )
      .merge()
      .setValue(
        "Total",
      )
      .setHorizontalAlignment(
        "center",
      )
      .setVerticalAlignment(
        "middle",
      )
      .setFontWeight(
        "bold",
      )
      .setBackground(
        "#b7b7b7",
      );

    template
      .getRange(
        destRow,
        5,
      )
      .setFormula(
        `=SUM(E${TEMPLATE_START_ROW}:E${destRow - 1})`,
      )
      .setFontWeight(
        "bold",
      )
      .setBackground(
        "#b7b7b7",
      )
      .setNumberFormat(
        "$#,##0.00",
      )
      .setFontColor("#000000")
      .setFontLine("none")
      .setShowHyperlink(false);
  }

  if (showAlert) {
    updateGrandTotal(false);

    SpreadsheetApp
      .getUi()
      .alert(
        "BOS material report generated successfully!",
      );
  }
}

//----------------------------------
// Generate VAN Report
//----------------------------------

function generateVANReport(
  showAlert = true,
) {
  const {
    source,
    template,
    data,
  } = getMaterialSourceData();

  if (data.length === 0) {
    if (showAlert) {
      SpreadsheetApp
        .getUi()
        .alert(
          "No material data found.",
        );
    }

    return;
  }

  //----------------------------------
  // Filter VAN Materials
  //----------------------------------

  const vanMaterials =
    data.filter((item) => {
      const row =
        item.row;

      const used =
        Number(row[9]) || 0;

      const category =
        String(
          row[12] || "",
        )
          .trim()
          .toUpperCase();

      return (
        category === "VAN" &&
        used !== 0
      );
    });

  //----------------------------------
  // Clear Previous VAN Report
  //----------------------------------

  const availableRows =
    template.getMaxRows() -
    TEMPLATE_START_ROW +
    1;

  template
    .getRange(
      TEMPLATE_START_ROW,
      7,
      availableRows,
      4,
    )
    .breakApart();

  template
    .getRange(
      TEMPLATE_START_ROW,
      7,
      availableRows,
      5,
    )
    .clearContent();

  let destRow =
    TEMPLATE_START_ROW;

  //----------------------------------
  // Process VAN Materials
  //----------------------------------

  vanMaterials.forEach(
    (item) => {

      //----------------------------------
      // Add Template Rows When Required
      //----------------------------------

      ensureSheetRows_(
        template,
        destRow + 1,
      );

      const row =
        item.row;

      const catalog =
        row[3];

      const description =
        row[5];

      const price =
        row[11] || row[6];

      const used =
        row[9];

      //----------------------------------
      // Apply VAN Row Formatting
      //----------------------------------

      const descriptionUpper =
        String(description || "")
          .trim()
          .toUpperCase();

      const greenKeywords = [
        "BOX",
        "SCWAY",
        "SC WWAY",
        "WWAY",
        "WIREWAY",
        "PAD",
        "PADS",
      ];

      const isGreenItem =
        greenKeywords.some(
          (keyword) =>
            descriptionUpper.includes(keyword),
        );

      const rowColor =
        isGreenItem
          ? "#d9ead3"
          : "#ead1dc";

      template
        .getRange(
          destRow,
          7,
          1,
          5,
        )
        .setBackground(rowColor)
        .setFontColor("#000000")
        .setFontWeight("normal")
        .setFontStyle("normal")
        .setVerticalAlignment("middle")
        .setBorder(
          true,
          true,
          true,
          true,
          true,
          true,
          "#000000",
          SpreadsheetApp.BorderStyle.SOLID,
        );

      //----------------------------------
      // I = Plain White
      //----------------------------------

      template
        .getRange(
          destRow,
          9,
        )
        .setBackground(
          "#ffffff",
        )
        .setFontColor(
          "#000000",
        )
        .setFontWeight(
          "normal",
        )
        .setFontStyle(
          "normal",
        )
        .setHorizontalAlignment(
          "center",
        )
        .setVerticalAlignment(
          "middle",
        )
        .setBorder(
          true,
          true,
          true,
          true,
          true,
          true,
        );

      //----------------------------------
      // Values
      //----------------------------------

      template
        .getRange(
          destRow,
          7,
        )
        .setValue(
          description,
        )
        .setHorizontalAlignment(
          "left",
        );

      template
        .getRange(
          destRow,
          8,
        )
        .setValue(
          catalog,
        )
        .setHorizontalAlignment(
          "left",
        );

      template
        .getRange(
          destRow,
          9,
        )
        .setValue(
          used,
        );

      template
        .getRange(
          destRow,
          10,
        )
        .setValue(
          price,
        )
        .setNumberFormat(
          "$#,##0.00",
        );

      //----------------------------------
      // Cost Formula
      //----------------------------------

      template
        .getRange(
          destRow,
          11,
        )
        .setFormula(
          `=I${destRow}*J${destRow}`,
        )
        .setNumberFormat(
          "$#,##0.00",
        )
        .setFontColor("#000000")
        .setFontLine("none")
        .setShowHyperlink(false);

      destRow++;
    },
  );

  //----------------------------------
  // VAN Total
  //----------------------------------

  if (
    destRow >
    TEMPLATE_START_ROW
  ) {

    ensureSheetRows_(
      template,
      destRow,
    );

    template
      .getRange(
        destRow - 1,
        7,
        1,
        5,
      )
      .copyTo(
        template.getRange(
          destRow,
          7,
        ),
        SpreadsheetApp
          .CopyPasteType
          .PASTE_FORMAT,
        false,
      );

    template
      .getRange(
        destRow,
        7,
        1,
        4,
      )
      .merge()
      .setValue(
        "Total",
      )
      .setHorizontalAlignment(
        "center",
      )
      .setVerticalAlignment(
        "middle",
      )
      .setFontWeight(
        "bold",
      )
      .setBackground(
        "#b7b7b7",
      );

    template
      .getRange(
        destRow,
        11,
      )
      .setFormula(
        `=SUM(K${TEMPLATE_START_ROW}:K${destRow - 1})`,
      )
      .setFontWeight(
        "bold",
      )
      .setBackground(
        "#b7b7b7",
      )
      .setNumberFormat(
        "$#,##0.00",
      )
      .setFontColor("#000000")
      .setFontLine("none")
      .setShowHyperlink(false);
  }

  if (showAlert) {
    updateGrandTotal(false);

    SpreadsheetApp
      .getUi()
      .alert(
        "VAN material report generated successfully!",
      );
  }
}

//----------------------------------
// Update Grand Total
//----------------------------------

function updateGrandTotal(
  showAlert = true,
) {
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        TEMPLATE_SHEET,
      );

  if (!sheet) {
    throw new Error(
      `Sheet not found: ${TEMPLATE_SHEET}`,
    );
  }

  const lastRow =
    sheet.getLastRow();

  let bosTotalRow =
    null;

  let vanTotalRow =
    null;

  //----------------------------------
  // Find Total Rows
  //----------------------------------

  for (
    let row =
      TEMPLATE_START_ROW;
    row <= lastRow;
    row++
  ) {
    const bosLabel =
      String(
        sheet
          .getRange(row, 1)
          .getDisplayValue(),
      )
        .trim()
        .toUpperCase();

    const vanLabel =
      String(
        sheet
          .getRange(row, 7)
          .getDisplayValue(),
      )
        .trim()
        .toUpperCase();

    if (
      bosLabel === "TOTAL"
    ) {
      bosTotalRow =
        row;
    }

    if (
      vanLabel === "TOTAL"
    ) {
      vanTotalRow =
        row;
    }
  }

  //----------------------------------
  // Flush Report Formulas
  //----------------------------------

  SpreadsheetApp.flush();

  //----------------------------------
  // Read Totals
  //----------------------------------

  const bosTotal =
    bosTotalRow
      ? Number(
          sheet
            .getRange(
              bosTotalRow,
              5,
            )
            .getValue(),
        ) || 0
      : 0;

  const vanTotal =
    vanTotalRow
      ? Number(
          sheet
            .getRange(
              vanTotalRow,
              11,
            )
            .getValue(),
        ) || 0
      : 0;

  //----------------------------------
  // Grand Total
  //----------------------------------

  sheet
    .getRange("F1")
    .setValue(
      bosTotal +
        vanTotal,
    )
    .setNumberFormat(
      "$#,##0.00",
    );

  if (showAlert) {
    SpreadsheetApp
      .getUi()
      .alert(
        "Grand total updated successfully!",
      );
  }
}