//----------------------------------
// Get Presets
//----------------------------------

function getPresets() {

  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName("PRESETS");

  if (!sheet) {
    throw new Error("PRESETS sheet not found.");
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = values.shift();

  return values.map(row => {

    const obj = {};

    headers.forEach((header, index) => {
      obj[String(header).trim()] = row[index];
    });

    return obj;

  });

}