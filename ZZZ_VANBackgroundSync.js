//----------------------------------
// VAN Background Sync
//----------------------------------
// Keeps the left BOS table unchanged.
// After the normal VAN report is generated, this reads the
// background color from Material List column G and applies
// that same color across the matching VAN report row G:K.

const generateVANReportBase_ = generateVANReport;

generateVANReport = function(showAlert = true) {

  // Generate the existing VAN report first.
  generateVANReportBase_(false);

  // Then restore the VAN row colors from Material List.
  syncVANBackgroundsFromMaterialList_();

  if (showAlert) {
    updateGrandTotal(false);

    SpreadsheetApp
      .getUi()
      .alert(
        "VAN material report generated successfully!"
      );
  }
};


//----------------------------------
// Copy VAN Background Colors
//----------------------------------

function syncVANBackgroundsFromMaterialList_() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const sourceSheet =
    spreadsheet.getSheetByName("Material List");

  const templateSheet =
    spreadsheet.getSheetByName("template") ||
    spreadsheet.getSheetByName("Template");

  if (!sourceSheet || !templateSheet) {
    return;
  }

  const sourceLastRow =
    sourceSheet.getLastRow();

  if (sourceLastRow < 2) {
    return;
  }

  //----------------------------------
  // Material List columns used:
  // D = Catalog / SKU
  // F = Description
  // G = Row background color source
  //----------------------------------

  const sourceValues =
    sourceSheet
      .getRange(
        2,
        1,
        sourceLastRow - 1,
        Math.max(sourceSheet.getLastColumn(), 7)
      )
      .getDisplayValues();

  const sourceBackgrounds =
    sourceSheet
      .getRange(
        2,
        7,
        sourceLastRow - 1,
        1
      )
      .getBackgrounds();

  const sourceLookup =
    new Map();

  sourceValues.forEach((row, index) => {

    const catalog =
      normalizeVANFormatKey_(row[3]);

    const description =
      normalizeVANFormatKey_(row[5]);

    const combinedKey =
      `${catalog}|${description}`;

    const rowColor =
      sourceBackgrounds[index][0];

    if (!sourceLookup.has(combinedKey)) {
      sourceLookup.set(
        combinedKey,
        rowColor
      );
    }

    // Catalog-only fallback for cases where descriptions differ slightly.
    if (
      catalog &&
      !sourceLookup.has(`CATALOG:${catalog}`)
    ) {
      sourceLookup.set(
        `CATALOG:${catalog}`,
        rowColor
      );
    }
  });

  const templateLastRow =
    templateSheet.getLastRow();

  for (
    let rowNumber = 6;
    rowNumber <= templateLastRow;
    rowNumber++
  ) {

    const description =
      templateSheet
        .getRange(rowNumber, 7)
        .getDisplayValue();

    const catalog =
      templateSheet
        .getRange(rowNumber, 8)
        .getDisplayValue();

    const normalizedDescription =
      normalizeVANFormatKey_(description);

    const normalizedCatalog =
      normalizeVANFormatKey_(catalog);

    if (
      normalizedDescription === "TOTAL"
    ) {
      break;
    }

    if (
      !normalizedDescription &&
      !normalizedCatalog
    ) {
      continue;
    }

    const combinedKey =
      `${normalizedCatalog}|${normalizedDescription}`;

    const rowColor =
      sourceLookup.get(combinedKey) ||
      sourceLookup.get(
        `CATALOG:${normalizedCatalog}`
      );

    if (!rowColor) {
      continue;
    }

    // Apply the single Material List G background across VAN G:K.
    templateSheet
      .getRange(
        rowNumber,
        7,
        1,
        5
      )
      .setBackground(rowColor);
  }

  SpreadsheetApp.flush();
}


//----------------------------------
// Normalize VAN Match Key
//----------------------------------

function normalizeVANFormatKey_(value) {

  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}
