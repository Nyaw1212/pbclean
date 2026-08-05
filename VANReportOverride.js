//----------------------------------
// VAN Report Override
//----------------------------------
// Replaces the VAN renderer without running a second formatting pass.
// Background color comes directly from Material List column G.

generateVANReport = function(
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
  // Read all Material List G backgrounds once
  //----------------------------------

  const sourceBackgrounds =
    source
      .getRange(
        SOURCE_START_ROW,
        7,
        data.length,
        1,
      )
      .getBackgrounds();

  //----------------------------------
  // Filter VAN Materials
  //----------------------------------

  const vanMaterials =
    data.filter((item) => {
      const row = item.row;
      const used = Number(row[9]) || 0;
      const category =
        String(row[12] || "")
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

  let destRow = TEMPLATE_START_ROW;

  //----------------------------------
  // Process VAN Materials
  //----------------------------------

  vanMaterials.forEach((item) => {
    ensureSheetRows_(
      template,
      destRow + 1,
    );

    const row = item.row;
    const catalog = row[3];
    const description = row[5];
    const price = row[11] || row[6];
    const used = row[9];

    const sourceBackgroundIndex =
      item.sourceRow - SOURCE_START_ROW;

    const rowColor =
      sourceBackgrounds[
        sourceBackgroundIndex
      ][0] || "#ffffff";

    //----------------------------------
    // Apply source background to G,H,J,K
    // Keep I white
    //----------------------------------

    template
      .getRange(
        destRow,
        7,
        1,
        5,
      )
      .setBackgrounds([[
        rowColor,
        rowColor,
        "#ffffff",
        rowColor,
        rowColor,
      ]])
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
    // Values
    //----------------------------------

    template
      .getRange(destRow, 7)
      .setValue(description)
      .setHorizontalAlignment("left");

    template
      .getRange(destRow, 8)
      .setValue(catalog)
      .setHorizontalAlignment("left");

    template
      .getRange(destRow, 9)
      .setValue(used)
      .setHorizontalAlignment("center");

    template
      .getRange(destRow, 10)
      .setValue(price)
      .setNumberFormat("$#,##0.00");

    template
      .getRange(destRow, 11)
      .setFormula(
        `=I${destRow}*J${destRow}`,
      )
      .setNumberFormat("$#,##0.00")
      .setFontColor("#000000")
      .setFontLine("none")
      .setShowHyperlink(false);

    destRow++;
  });

  //----------------------------------
  // VAN Total
  //----------------------------------

  if (destRow > TEMPLATE_START_ROW) {
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
        template.getRange(destRow, 7),
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
      .setValue("Total")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontWeight("bold")
      .setBackground("#b7b7b7");

    template
      .getRange(destRow, 11)
      .setFormula(
        `=SUM(K${TEMPLATE_START_ROW}:K${destRow - 1})`,
      )
      .setFontWeight("bold")
      .setBackground("#b7b7b7")
      .setNumberFormat("$#,##0.00")
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
};
