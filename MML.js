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

  const priceSheet =
    spreadsheet.getSheetByName("PRICEUP");

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

  const priceMap =
    buildPriceUpMap_(priceSheet);

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

    const catalogNumber =
      normalizeSku_(obj.CatalogNumber);

    if (
      catalogNumber &&
      priceMap.has(catalogNumber)
    ) {
      obj.Price =
        priceMap.get(catalogNumber);
    }

    return obj;
  });
}

//----------------------------------
// Build PRICEUP Lookup
//----------------------------------

function buildPriceUpMap_(priceSheet) {

  const priceMap =
    new Map();

  if (!priceSheet) {
    Logger.log(
      "PRICEUP sheet not found. Using MML prices."
    );

    return priceMap;
  }

  const values =
    priceSheet.getDataRange().getValues();

  if (values.length < 2) {
    return priceMap;
  }

  const headers =
    values.shift().map(header =>
      String(header || "")
        .trim()
        .toUpperCase()
    );

  const skuIndex =
    headers.indexOf("SKU");

  const priceIndex =
    headers.indexOf("PRICE");

  const codeIndex =
    headers.indexOf("CODE");

  if (
    skuIndex === -1 ||
    priceIndex === -1
  ) {
    Logger.log(
      "PRICEUP is missing SKU or PRICE column. Using MML prices."
    );

    return priceMap;
  }

  values.forEach(row => {

    const sku =
      normalizeSku_(row[skuIndex]);

    if (!sku) {
      return;
    }

    const priceEach =
      normalizePriceEach_(
        row[priceIndex],
        codeIndex === -1
          ? "E"
          : row[codeIndex]
      );

    if (priceEach === null) {
      return;
    }

    // Keep the first exact SKU match,
    // matching normal spreadsheet lookup behavior.
    if (!priceMap.has(sku)) {
      priceMap.set(
        sku,
        priceEach
      );
    }
  });

  Logger.log(
    "PRICEUP matched SKUs loaded: " +
    priceMap.size
  );

  return priceMap;
}

//----------------------------------
// Normalize PRICEUP Price To Each
//----------------------------------

function normalizePriceEach_(price, code) {

  if (
    price === "" ||
    price === null ||
    price === undefined
  ) {
    return null;
  }

  const value =
    Number(price);

  if (!Number.isFinite(value)) {
    return null;
  }

  const uom =
    String(code || "E")
      .trim()
      .toUpperCase();

  switch (uom) {

    case "C":
      return value / 100;

    case "M":
      return value / 1000;

    case "E":
    default:
      return value;
  }
}

//----------------------------------
// Normalize SKU
//----------------------------------

function normalizeSku_(sku) {

  return String(sku || "")
    .trim()
    .toUpperCase();
}
