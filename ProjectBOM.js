//----------------------------------
// Constants
//----------------------------------

const PROJECT_BOM_SHEET = "PROJECT_BOM";

//----------------------------------
// Normalize BOM Header
//----------------------------------

function normalizeProjectBOMHeader(value) {

  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");

}

//----------------------------------
// Header Map
//----------------------------------

function getProjectBOMHeaderMap(sheet) {

  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {

    throw new Error(
      "PROJECT_BOM does not contain any headers."
    );

  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0];

  const map = {};

  headers.forEach((header, index) => {

    const key =
      normalizeProjectBOMHeader(header);

    if (key) {

      map[key] = index;

    }

  });

  return {

    headers,

    map

  };

}

//----------------------------------
// Find BOM Column
//----------------------------------

function getProjectBOMColumn(
  map,
  names
) {

  for (let i = 0; i < names.length; i++) {

    const key =
      normalizeProjectBOMHeader(
        names[i]
      );

    if (map[key] !== undefined) {

      return map[key];

    }

  }

  return undefined;

}

//----------------------------------
// Set BOM Row Value
//----------------------------------

function setProjectBOMValue(
  row,
  map,
  names,
  value
) {

  const column =
    getProjectBOMColumn(
      map,
      names
    );

  if (column !== undefined) {

    row[column] = value;

  }

}

//----------------------------------
// Save Project BOM
//----------------------------------

function saveProjectBOM(

  projectId,

  jobName,

  address,

  materials

) {

  const spreadsheet =
    getProjectSpreadsheet();

  const sheet =
    spreadsheet.getSheetByName(
      PROJECT_BOM_SHEET
    );

  if (!sheet) {

    throw new Error(
      'Sheet "PROJECT_BOM" was not found.'
    );

  }

  const normalizedProjectId =
    String(projectId || "").trim();

  if (!normalizedProjectId) {

    throw new Error(
      "Project ID is required to save the BOM."
    );

  }

  //----------------------------------
  // Delete Existing Project BOM
  //----------------------------------

  deleteProjectBOM(
    normalizedProjectId
  );

  //----------------------------------
  // Validate Materials
  //----------------------------------

  const materialList =
    Array.isArray(materials)
      ? materials
      : [];

  if (!materialList.length) {

    return {

      success: true,

      projectId:
        normalizedProjectId,

      materialCount: 0

    };

  }

  //----------------------------------
  // Headers
  //----------------------------------

  const {
    headers,
    map
  } = getProjectBOMHeaderMap(sheet);

  //----------------------------------
  // Build Rows
  //----------------------------------

  const rows =
    materialList.map(material => {

      const row =
        new Array(
          headers.length
        ).fill("");

      setProjectBOMValue(

        row,

        map,

        [
          "ProjectID",
          "Project ID"
        ],

        normalizedProjectId

      );

      setProjectBOMValue(

        row,

        map,

        [
          "JobName",
          "Job Name"
        ],

        String(
          jobName || ""
        ).trim()

      );

      setProjectBOMValue(

        row,

        map,

        [
          "Address"
        ],

        String(
          address || ""
        ).trim()

      );

      setProjectBOMValue(

        row,

        map,

        [
          "CatalogNumber",
          "Catalog Number"
        ],

        String(

          material.CatalogNumber ||

          material.catalogNumber ||

          ""

        ).trim()

      );

      setProjectBOMValue(

        row,

        map,

        [
          "Description"
        ],

        String(

          material.Description ||

          material.description ||

          ""

        ).trim()

      );

      setProjectBOMValue(

        row,

        map,

        [
          "Type"
        ],

        String(

          material.Type ||

          material.type ||

          ""

        ).trim()

      );

      setProjectBOMValue(

        row,

        map,

        [
          "InventoryCategory",
          "Inventory Category"
        ],

        String(

          material.InventoryCategory ||

          material.inventoryCategory ||

          material.category ||

          ""

        ).trim()

      );

      setProjectBOMValue(

        row,

        map,

        [
          "Supplier"
        ],

        String(

          material.Supplier ||

          material.supplier ||

          ""

        ).trim()

      );

      setProjectBOMValue(

        row,

        map,

        [
          "VendorCode",
          "Vendor Code"
        ],

        String(

          material.VendorCode ||

          material.vendorCode ||

          ""

        ).trim()

      );

      setProjectBOMValue(

        row,

        map,

        [
          "Price"
        ],

        Number(

          material.Price ??

          material.price ??

          0

        ) || 0

      );

      setProjectBOMValue(

        row,

        map,

        [
          "Qty",
          "Quantity"
        ],

        Number(

          material.Qty ??

          material.qty ??

          material.Quantity ??

          0

        ) || 0

      );

      return row;

    });

  //----------------------------------
  // Append Rows
  //----------------------------------

  sheet
    .getRange(

      sheet.getLastRow() + 1,

      1,

      rows.length,

      headers.length

    )
    .setValues(rows);

  return {

    success: true,

    projectId:
      normalizedProjectId,

    materialCount:
      rows.length

  };

}

//----------------------------------
// Delete Existing Project BOM
//----------------------------------

function deleteProjectBOM(projectId) {

  const spreadsheet =
    getProjectSpreadsheet();

  const sheet =
    spreadsheet.getSheetByName(
      PROJECT_BOM_SHEET
    );

  if (!sheet) {

    return;

  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    return;

  }

  const { map } =
    getProjectBOMHeaderMap(sheet);

  const projectIdColumn =
    getProjectBOMColumn(

      map,

      [
        "ProjectID",
        "Project ID"
      ]

    );

  if (projectIdColumn === undefined) {

    throw new Error(
      "PROJECT_BOM is missing the ProjectID header."
    );

  }

  const values =
    sheet
      .getRange(

        2,

        1,

        lastRow - 1,

        sheet.getLastColumn()

      )
      .getValues();

  const normalizedProjectId =
    String(projectId || "").trim();

  //----------------------------------
  // Delete Bottom-Up
  //----------------------------------

  for (

    let index =
      values.length - 1;

    index >= 0;

    index--

  ) {

    const savedProjectId =
      String(

        values[index][projectIdColumn] ||

        ""

      ).trim();

    if (
      savedProjectId ===
      normalizedProjectId
    ) {

      sheet.deleteRow(
        index + 2
      );

    }

  }

}