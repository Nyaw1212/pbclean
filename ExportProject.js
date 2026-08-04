//----------------------------------
// Export Menu Installer
//----------------------------------

function installPanelBuilderExportMenu() {

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const existingTriggers =
    ScriptApp.getProjectTriggers();

  const alreadyInstalled =
    existingTriggers.some(
      trigger =>
        trigger.getHandlerFunction() ===
        "showPanelBuilderExportMenu_"
    );

  if (!alreadyInstalled) {

    ScriptApp
      .newTrigger(
        "showPanelBuilderExportMenu_"
      )
      .forSpreadsheet(
        spreadsheet
      )
      .onOpen()
      .create();

  }

  showPanelBuilderExportMenu_();

  spreadsheet.toast(
    alreadyInstalled
      ? "Export menu is already installed."
      : "Export menu installed successfully.",
    "Panel Builder",
    5
  );

}


//----------------------------------
// Export Menu
//----------------------------------

function showPanelBuilderExportMenu_() {

  SpreadsheetApp
    .getUi()
    .createMenu(
      "Panel Builder Export"
    )
    .addItem(
      "Export Template + Material List",
      "exportTemplateAndMaterialList"
    )
    .addToUi();

}


//----------------------------------
// Export Template + Material List
//----------------------------------

function exportTemplateAndMaterialList() {

  const sourceSpreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  const templateSheet =
    getSheetByNames_(
      sourceSpreadsheet,
      ["template", "Template"]
    );

  const materialListSheet =
    sourceSpreadsheet.getSheetByName(
      "Material List"
    );

  if (!templateSheet) {
    throw new Error(
      'Sheet not found: "template" or "Template".'
    );
  }

  if (!materialListSheet) {
    throw new Error(
      'Sheet not found: "Material List".'
    );
  }

  const jobName =
    String(
      templateSheet
        .getRange("A2")
        .getDisplayValue() ||
      templateSheet
        .getRange("O1")
        .getDisplayValue() ||
      "Panel Builder Export"
    )
      .trim()
      .replace(/[\\/:*?"<>|]/g, "-");

  const timestamp =
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HHmm"
    );

  const exportSpreadsheet =
    SpreadsheetApp.create(
      `${jobName} - Material Report - ${timestamp}`
    );

  const copiedTemplate =
    templateSheet.copyTo(
      exportSpreadsheet
    );

  copiedTemplate.setName(
    "Template"
  );

  const copiedMaterialList =
    materialListSheet.copyTo(
      exportSpreadsheet
    );

  copiedMaterialList.setName(
    "Material List"
  );

  const defaultSheet =
    exportSpreadsheet.getSheetByName(
      "Sheet1"
    );

  if (
    defaultSheet &&
    exportSpreadsheet.getSheets().length > 1
  ) {
    exportSpreadsheet.deleteSheet(
      defaultSheet
    );
  }

  exportSpreadsheet.setActiveSheet(
    copiedTemplate
  );

  exportSpreadsheet.moveActiveSheet(1);

  exportSpreadsheet.setActiveSheet(
    copiedMaterialList
  );

  exportSpreadsheet.moveActiveSheet(2);

  exportSpreadsheet.setActiveSheet(
    copiedTemplate
  );

  SpreadsheetApp.flush();

  const exportUrl =
    exportSpreadsheet.getUrl();

  sourceSpreadsheet.toast(
    "Template and Material List copied successfully.",
    "Export Complete",
    5
  );

  showExportCompleteDialog_(
    exportSpreadsheet.getName(),
    exportUrl
  );

  return {
    success: true,
    spreadsheetId:
      exportSpreadsheet.getId(),
    spreadsheetName:
      exportSpreadsheet.getName(),
    spreadsheetUrl:
      exportUrl,
  };

}


//----------------------------------
// Find Sheet Using Possible Names
//----------------------------------

function getSheetByNames_(
  spreadsheet,
  names
) {

  for (const name of names) {

    const sheet =
      spreadsheet.getSheetByName(name);

    if (sheet) {
      return sheet;
    }

  }

  return null;

}


//----------------------------------
// Export Completion Dialog
//----------------------------------

function showExportCompleteDialog_(
  spreadsheetName,
  spreadsheetUrl
) {

  const escapedName =
    String(spreadsheetName)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");

  const escapedUrl =
    String(spreadsheetUrl)
      .replace(/&/g, "&amp;")
      .replace(/"/g, "&quot;");

  const html =
    HtmlService
      .createHtmlOutput(
        `<div style="font-family:Arial,sans-serif;padding:18px;line-height:1.45">` +
        `<h3 style="margin:0 0 10px">Export complete</h3>` +
        `<p style="margin:0 0 16px">${escapedName}</p>` +
        `<a href="${escapedUrl}" target="_blank" ` +
        `style="display:inline-block;padding:10px 14px;background:#1769e0;color:#fff;text-decoration:none;border-radius:6px;font-weight:700">` +
        `Open new spreadsheet</a>` +
        `</div>`
      )
      .setWidth(420)
      .setHeight(190);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      "Panel Builder Export"
    );

}
