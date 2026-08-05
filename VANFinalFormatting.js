//----------------------------------
// Install Fixed Report Menu
//----------------------------------

function installFixedReportMenu() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();

  const installed = triggers.some(
    trigger => trigger.getHandlerFunction() === "showFixedReportMenu_"
  );

  if (!installed) {
    ScriptApp
      .newTrigger("showFixedReportMenu_")
      .forSpreadsheet(spreadsheet)
      .onOpen()
      .create();
  }

  showFixedReportMenu_();

  spreadsheet.toast(
    installed
      ? "Fixed report menu is already installed."
      : "Fixed report menu installed successfully.",
    "Panel Builder",
    5
  );
}


//----------------------------------
// Fixed Report Menu
//----------------------------------

function showFixedReportMenu_() {
  SpreadsheetApp
    .getUi()
    .createMenu("Panel Builder Report")
    .addItem(
      "Generate Material List + Report",
      "generateAllReportsWithVanBackgrounds"
    )
    .addToUi();
}


//----------------------------------
// Generate Report + Final VAN Colors
//----------------------------------

function generateAllReportsWithVanBackgrounds() {
  generateAllReports();
  SpreadsheetApp.flush();
  applyVanBackgroundsFromMaterialList_();
}


//----------------------------------
// Apply VAN Backgrounds Last
//----------------------------------

function applyVanBackgroundsFromMaterialList_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const materialList = spreadsheet.getSheetByName("Material List");
  const template =
    spreadsheet.getSheetByName("template") ||
    spreadsheet.getSheetByName("Template");

  if (!materialList) {
    throw new Error('Sheet not found: "Material List".');
  }

  if (!template) {
    throw new Error('Sheet not found: "template" or "Template".');
  }

  const sourceStartRow = 2;
  const sourceLastRow = materialList.getLastRow();

  if (sourceLastRow < sourceStartRow) {
    return;
  }

  // Material List: D = Catalog, G = row background, M = category.
  const sourceRows = materialList
    .getRange(
      sourceStartRow,
      1,
      sourceLastRow - sourceStartRow + 1,
      Math.max(materialList.getLastColumn(), 13)
    )
    .getDisplayValues();

  const sourceBackgrounds = materialList
    .getRange(
      sourceStartRow,
      7,
      sourceLastRow - sourceStartRow + 1,
      1
    )
    .getBackgrounds();

  const backgroundByCatalog = new Map();

  sourceRows.forEach((row, index) => {
    const catalog = normalizeVanCatalog_(row[3]);
    const category = String(row[12] || "").trim().toUpperCase();

    if (catalog && category === "VAN") {
      backgroundByCatalog.set(catalog, sourceBackgrounds[index][0]);
    }
  });

  const reportStartRow = 6;
  const reportLastRow = template.getLastRow();

  if (reportLastRow < reportStartRow) {
    return;
  }

  // VAN report: H = Catalog. Apply source G background across G:J.
  const reportCatalogs = template
    .getRange(
      reportStartRow,
      8,
      reportLastRow - reportStartRow + 1,
      1
    )
    .getDisplayValues();

  reportCatalogs.forEach((row, index) => {
    const catalog = normalizeVanCatalog_(row[0]);
    const reportRow = reportStartRow + index;

    if (!catalog || !backgroundByCatalog.has(catalog)) {
      return;
    }

    const background = backgroundByCatalog.get(catalog);

    template
      .getRange(reportRow, 7, 1, 4) // G:J only
      .setBackground(background);
  });

  SpreadsheetApp.flush();
}


//----------------------------------
// Normalize VAN Catalog
//----------------------------------

function normalizeVanCatalog_(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");
}
