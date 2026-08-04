//----------------------------------
// Get MML
//----------------------------------

function getMML() {

  const spreadsheet =
    getProjectSpreadsheet();

  const sheet =
    spreadsheet.getSheetByName("MML");

  if (!sheet) {
    throw new Error(
      'MML sheet not found in spreadsheet: ' +
      spreadsheet.getName()
    );
  }

  Logger.log(
    "MML Spreadsheet: " +
    spreadsheet.getName()
  );

  Logger.log(
    "MML Spreadsheet ID: " +
    spreadsheet.getId()
  );

  Logger.log(
    "MML Sheet: " +
    sheet.getName()
  );

  const values =
    sheet.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers =
    values.shift();

  return values.map(row => {

    const obj = {};

    headers.forEach(
      (header, index) => {

        const key =
          String(header).trim();

        let value =
          row[index];

        if (
          key === "CatalogNumber"
        ) {
          value =
            String(value || "")
              .trim();
        }

        if (
          key === "Placeable"
        ) {
          value =
            value === true ||
            String(value)
              .toUpperCase() ===
              "TRUE";
        }

        if (
          key === "Qty" ||
          key === "Quantity" ||
          key === "Price"
        ) {
          value =
            Number(value) || 0;
        }

        if (
          typeof value === "string"
        ) {
          value =
            value.trim();
        }

        obj[key] =
          value;
      }
    );

    return obj;
  });
}