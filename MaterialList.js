//----------------------------------
// Custom Menu
//----------------------------------

function onOpen() {

  SpreadsheetApp
    .getUi()
    .createMenu("Material List")
    .addItem(
      "Generate Material List + Report",
      "generateAllReports"
    )
    .addSeparator()
    .addItem(
      "Populate Material List Only",
      "populateMaterialList"
    )
    .addItem(
      "Clear Material List Filters",
      "clearMaterialListFilters"
    )
    .addToUi();

}


//----------------------------------
// Configuration
//----------------------------------

const MATERIAL_LIST_CONFIG = {

  projectBOMSheet: "PROJECT_BOM",
  mmlSheet: "MML",
  materialListSheet: "Material List",
  templateSheet: "Template",

  jobNameCell: "O1",

  materialListHeaderRow: 1,
  materialListStartRow: 2,

  //----------------------------------
  // Material List Columns
  //----------------------------------

  materialListColumns: {

    packageQuantity: 3,      // C
    catalogNumber: 4,       // D
    packagePrice: 7,        // G
    quantity: 10,           // J
    unitPrice: 12,          // L
    inventoryCategory: 13,  // M

  },

  //----------------------------------
  // PROJECT_BOM Header Names
  //----------------------------------

  projectBOMHeaders: {

    jobName: [
      "JobName",
      "Job Name",
    ],

    catalogNumber: [
      "CatalogNumber",
      "Catalog Number",
      "CN",
    ],

    quantity: [
      "Quantity",
      "Qty",
      "QTY",
    ],

    inventoryCategory: [
      "InventoryCategory",
      "Inventory Category",
      "Category",
    ],

  },

};


//----------------------------------
// Populate Material List
//----------------------------------

function populateMaterialList() {

  const spreadsheet =
    SpreadsheetApp.getActive();

  const config =
    MATERIAL_LIST_CONFIG;

  const projectBOMSheet =
    spreadsheet.getSheetByName(
      config.projectBOMSheet
    );

  const mmlSheet =
    spreadsheet.getSheetByName(
      config.mmlSheet
    );

  const materialListSheet =
    spreadsheet.getSheetByName(
      config.materialListSheet
    );

  const templateSheet =
    spreadsheet.getSheetByName(
      config.templateSheet
    );

  //----------------------------------
  // Validate Sheets
  //----------------------------------

  if (!projectBOMSheet) {

    throw new Error(
      `${config.projectBOMSheet} sheet not found.`
    );

  }

  if (!mmlSheet) {

    throw new Error(
      `${config.mmlSheet} sheet not found.`
    );

  }

  if (!materialListSheet) {

    throw new Error(
      `${config.materialListSheet} sheet not found.`
    );

  }

  if (!templateSheet) {

    throw new Error(
      `${config.templateSheet} sheet not found.`
    );

  }

  //----------------------------------
  // Selected Job Name
  //----------------------------------

  const selectedJobName =
    String(
      templateSheet
        .getRange(config.jobNameCell)
        .getDisplayValue()
    ).trim();

  if (!selectedJobName) {

    throw new Error(
      `Select a JobName in ` +
      `${config.templateSheet}!${config.jobNameCell}.`
    );

  }

  //----------------------------------
  // Build MML InventoryCategory Lookup
  //----------------------------------

  const mmlData =
    mmlSheet
      .getDataRange()
      .getValues();

  if (mmlData.length < 1) {

    throw new Error(
      "MML contains no data."
    );

  }

  const mmlHeaders =
    mmlData[0];

  const mmlCatalogIndex =
    findHeaderIndex_(
      mmlHeaders,
      [
        "CatalogNumber",
        "Catalog Number",
        "CN"
      ]
    );

  const mmlInventoryIndex =
    findHeaderIndex_(
      mmlHeaders,
      [
        "InventoryCategory",
        "Inventory Category",
        "Category"
      ]
    );

  if (mmlCatalogIndex === -1) {

    throw new Error(
      "MML is missing the CatalogNumber header."
    );

  }

  if (mmlInventoryIndex === -1) {

    throw new Error(
      "MML is missing the InventoryCategory header."
    );

  }

  const mmlLookup =
    new Map();

  mmlData
    .slice(1)
    .forEach(row => {

      const catalogNumber =
        normalizeCatalogNumber_(
          row[mmlCatalogIndex]
        );

      if (!catalogNumber) {

        return;

      }

      const inventoryCategory =
        String(
          row[mmlInventoryIndex] || ""
        )
          .trim()
          .toUpperCase();

      if (
        inventoryCategory &&
        !mmlLookup.has(catalogNumber)
      ) {

        mmlLookup.set(
          catalogNumber,
          inventoryCategory
        );

      }

    });

  //----------------------------------
  // Read PROJECT_BOM
  //----------------------------------

  const projectData =
    projectBOMSheet
      .getDataRange()
      .getValues();

  if (projectData.length < 2) {

    throw new Error(
      "PROJECT_BOM contains no records."
    );

  }

  const projectHeaders =
    projectData[0];

  const indexes = {

    jobName: findHeaderIndex_(
      projectHeaders,
      config.projectBOMHeaders.jobName
    ),

    catalogNumber: findHeaderIndex_(
      projectHeaders,
      config.projectBOMHeaders.catalogNumber
    ),

    quantity: findHeaderIndex_(
      projectHeaders,
      config.projectBOMHeaders.quantity
    ),

    inventoryCategory: findHeaderIndex_(
      projectHeaders,
      config.projectBOMHeaders.inventoryCategory
    ),

  };

  //----------------------------------
  // Validate Required Headers
  //----------------------------------

  const missingHeaders =
    Object.keys(indexes)
      .filter(
        key => indexes[key] === -1
      );

  if (missingHeaders.length > 0) {

    throw new Error(
      "PROJECT_BOM is missing headers: " +
      missingHeaders.join(", ")
    );

  }

  //----------------------------------
  // Build PROJECT_BOM Lookup
  //----------------------------------

  const bomLookup =
    new Map();

  const normalizedSelectedJob =
    normalizeText_(
      selectedJobName
    );

  projectData
    .slice(1)
    .forEach(row => {

      const rowJobName =
        normalizeText_(
          row[indexes.jobName]
        );

      if (
        rowJobName !==
        normalizedSelectedJob
      ) {

        return;

      }

      const catalogNumber =
        normalizeCatalogNumber_(
          row[indexes.catalogNumber]
        );

      if (!catalogNumber) {

        return;

      }

      const quantity =
        toNumber_(
          row[indexes.quantity]
        );

      let inventoryCategory =
        String(
          row[indexes.inventoryCategory] || ""
        )
          .trim()
          .toUpperCase();

      //----------------------------------
      // Fallback to MML InventoryCategory
      //----------------------------------

      if (!inventoryCategory) {

        inventoryCategory =
          mmlLookup.get(
            catalogNumber
          ) || "";

      }

      //----------------------------------
      // Create Lookup Record
      //----------------------------------

      if (!bomLookup.has(catalogNumber)) {

        bomLookup.set(
          catalogNumber,
          {
            quantity: 0,
            inventoryCategory: "",
          }
        );

      }

      const record =
        bomLookup.get(catalogNumber);

      //----------------------------------
      // Add Duplicate Quantities
      //----------------------------------

      record.quantity += quantity;

      //----------------------------------
      // Preserve First Nonblank Category
      //----------------------------------

      if (
        !record.inventoryCategory &&
        inventoryCategory
      ) {

        record.inventoryCategory =
          inventoryCategory;

      }

    });

  //----------------------------------
  // Read Material List
  //----------------------------------

  const startRow =
    config.materialListStartRow;

  const lastRow =
    materialListSheet.getLastRow();

  if (lastRow < startRow) {

    spreadsheet.toast(
      "Material List contains no material rows.",
      "Material List",
      5
    );

    return;

  }

  const rowCount =
    lastRow - startRow + 1;

  const columns =
    config.materialListColumns;

  const catalogNumbers =
    materialListSheet
      .getRange(
        startRow,
        columns.catalogNumber,
        rowCount,
        1
      )
      .getValues();

  const packageQuantities =
    materialListSheet
      .getRange(
        startRow,
        columns.packageQuantity,
        rowCount,
        1
      )
      .getValues();

  const packagePrices =
    materialListSheet
      .getRange(
        startRow,
        columns.packagePrice,
        rowCount,
        1
      )
      .getValues();

  //----------------------------------
  // Build Output Values
  //----------------------------------

  const quantityOutput = [];
  const unitPriceOutput = [];
  const categoryOutput = [];

  for (
    let index = 0;
    index < rowCount;
    index++
  ) {

    const catalogNumber =
      normalizeCatalogNumber_(
        catalogNumbers[index][0]
      );

    const bomRecord =
      bomLookup.get(catalogNumber);

    //----------------------------------
    // Quantity → Column J
    //----------------------------------

    const quantity =
      bomRecord
        ? bomRecord.quantity
        : 0;

    quantityOutput.push([
      quantity > 0
        ? quantity
        : 0
    ]);

    //----------------------------------
    // Unit Price = G ÷ C → Column L
    //----------------------------------

    const packageQuantity =
      toNumber_(
        packageQuantities[index][0]
      );

    const packagePrice =
      toNumber_(
        packagePrices[index][0]
      );

    const unitPrice =
      packageQuantity > 0
        ? packagePrice / packageQuantity
        : "";

    unitPriceOutput.push([
      unitPrice
    ]);

    //----------------------------------
    // InventoryCategory → Column M
    //----------------------------------

    const inventoryCategory =
      bomRecord
        ? bomRecord.inventoryCategory
        : "";

    categoryOutput.push([
      inventoryCategory
    ]);

  }

  //----------------------------------
  // Write Values Only
  //----------------------------------
  // Background colors are preserved.
  //----------------------------------

  materialListSheet
    .getRange(
      startRow,
      columns.quantity,
      rowCount,
      1
    )
    .setValues(
      quantityOutput
    );

  materialListSheet
    .getRange(
      startRow,
      columns.unitPrice,
      rowCount,
      1
    )
    .setValues(
      unitPriceOutput
    )
    .setNumberFormat(
      "$#,##0.00"
    );

  materialListSheet
    .getRange(
      startRow,
      columns.inventoryCategory,
      rowCount,
      1
    )
    .setValues(
      categoryOutput
    );

  SpreadsheetApp.flush();

  //----------------------------------
  // Apply Filters
  //----------------------------------

  applyMaterialListFilters_(
    materialListSheet,
    lastRow
  );

  //----------------------------------
  // Completion Message
  //----------------------------------

  spreadsheet.toast(
    `${bomLookup.size} BOM catalog numbers loaded for ${selectedJobName}.`,
    "Material List Updated",
    5
  );

}


//----------------------------------
// Apply Material List Filters
//----------------------------------

function applyMaterialListFilters_(
  sheet,
  lastRow
) {

  const config =
    MATERIAL_LIST_CONFIG;

  const headerRow =
    config.materialListHeaderRow;

  const startRow =
    config.materialListStartRow;

  const quantityColumn =
    config
      .materialListColumns
      .quantity;

  const inventoryColumn =
    config
      .materialListColumns
      .inventoryCategory;

  const filterRowCount =
    lastRow - headerRow + 1;

  const lastColumn =
    Math.max(
      sheet.getLastColumn(),
      inventoryColumn
    );

  //----------------------------------
  // Get Existing Filter
  //----------------------------------

  let filter =
    sheet.getFilter();

  //----------------------------------
  // Recreate Filter if Range Changed
  //----------------------------------

  if (filter) {

    const filterRange =
      filter.getRange();

    const rangeMatches =
      filterRange.getRow() === headerRow &&
      filterRange.getColumn() === 1 &&
      filterRange.getNumRows() === filterRowCount &&
      filterRange.getNumColumns() === lastColumn;

    if (!rangeMatches) {

      filter.remove();

      filter = null;

    }

  }

  //----------------------------------
  // Create Filter
  //----------------------------------

  if (!filter) {

    filter =
      sheet
        .getRange(
          headerRow,
          1,
          filterRowCount,
          lastColumn
        )
        .createFilter();

  }

  //----------------------------------
  // Filter J: Quantity > 0
  //----------------------------------

  const quantityCriteria =
    SpreadsheetApp
      .newFilterCriteria()
      .whenNumberGreaterThan(0)
      .build();

  filter.setColumnFilterCriteria(
    quantityColumn,
    quantityCriteria
  );

  //----------------------------------
  // Read InventoryCategory Values
  //----------------------------------

  const inventoryValues =
    sheet
      .getRange(
        startRow,
        inventoryColumn,
        lastRow - startRow + 1,
        1
      )
      .getDisplayValues()
      .flat()
      .map(value =>
        String(value).trim()
      );

  //----------------------------------
  // Hide All Categories Except
  // VAN and BOS
  //----------------------------------

  const hiddenInventoryValues =
    [
      ...new Set(
        inventoryValues
      ),
    ]
      .filter(value => {

        const normalizedValue =
          value.toUpperCase();

        return (
          normalizedValue !== "VAN" &&
          normalizedValue !== "BOS"
        );

      });

  //----------------------------------
  // Apply Inventory Filter
  //----------------------------------

  if (
    hiddenInventoryValues.length > 0
  ) {

    const inventoryCriteria =
      SpreadsheetApp
        .newFilterCriteria()
        .setHiddenValues(
          hiddenInventoryValues
        )
        .build();

    filter.setColumnFilterCriteria(
      inventoryColumn,
      inventoryCriteria
    );

  } else {

    filter.removeColumnFilterCriteria(
      inventoryColumn
    );

  }

}


//----------------------------------
// Clear Material List Filters
//----------------------------------

function clearMaterialListFilters() {

  const spreadsheet =
    SpreadsheetApp.getActive();

  const sheet =
    spreadsheet.getSheetByName(
      MATERIAL_LIST_CONFIG.materialListSheet
    );

  if (!sheet) {

    throw new Error(
      `${MATERIAL_LIST_CONFIG.materialListSheet} sheet not found.`
    );

  }

  const filter =
    sheet.getFilter();

  if (!filter) {

    spreadsheet.toast(
      "No Material List filter is currently applied.",
      "Material List",
      4
    );

    return;

  }

  const quantityColumn =
    MATERIAL_LIST_CONFIG
      .materialListColumns
      .quantity;

  const inventoryColumn =
    MATERIAL_LIST_CONFIG
      .materialListColumns
      .inventoryCategory;

  filter.removeColumnFilterCriteria(
    quantityColumn
  );

  filter.removeColumnFilterCriteria(
    inventoryColumn
  );

  spreadsheet.toast(
    "Material List filters cleared.",
    "Material List",
    4
  );

}


//----------------------------------
// Find Header Index
//----------------------------------

function findHeaderIndex_(
  headers,
  aliases
) {

  const normalizedHeaders =
    headers.map(
      normalizeHeader_
    );

  for (
    const alias of aliases
  ) {

    const normalizedAlias =
      normalizeHeader_(
        alias
      );

    const index =
      normalizedHeaders.indexOf(
        normalizedAlias
      );

    if (index !== -1) {

      return index;

    }

  }

  return -1;

}


//----------------------------------
// Normalize Header
//----------------------------------

function normalizeHeader_(
  value
) {

  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(
      /[^a-z0-9]/g,
      ""
    );

}


//----------------------------------
// Normalize Text
//----------------------------------

function normalizeText_(
  value
) {

  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(
      /\s+/g,
      " "
    );

}


//----------------------------------
// Normalize Catalog Number
//----------------------------------

function normalizeCatalogNumber_(
  value
) {

  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(
      /\s+/g,
      ""
    );

}


//----------------------------------
// Convert Value to Number
//----------------------------------

function toNumber_(
  value
) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {

    return 0;

  }

  if (
    typeof value === "number"
  ) {

    return Number.isFinite(value)
      ? value
      : 0;

  }

  const cleaned =
    String(value)
      .replace(
        /[$,\s]/g,
        ""
      );

  const number =
    Number(cleaned);

  return Number.isFinite(number)
    ? number
    : 0;

}