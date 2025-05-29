function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Better Reports -> RAW/CUT/LIST')
      .addItem('Turn into PO!', 'entireAutomation')
      .addToUi();
}

function entireAutomation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // STEP 1: Rename current sheet to "RAW" and duplicate as "CUT"
  const rawSheet = ss.getActiveSheet();
  rawSheet.setName("RAW");

  const copiedSheet = rawSheet.copyTo(ss);
  copiedSheet.setName("CUT");

  const cutSheet = copiedSheet;
  ss.setActiveSheet(cutSheet);

  Logger.log("cutSheet name: " + cutSheet.getName());

  cutSheet.setFrozenRows(1);

// Remove "Bundle Variant Inventory quantity" column if it exists
headers = cutSheet.getRange(1, 1, 1, cutSheet.getLastColumn()).getValues()[0];
headerMap = headers.reduce((map, val, i) => {
  map[val] = i;
  return map;
}, {});

if ("Bundle Variant Inventory quantity" in headerMap) {
  const bundleColIndex = headerMap["Bundle Variant Inventory quantity"] + 1;
  cutSheet.deleteColumn(bundleColIndex);
  Logger.log('Removed "Bundle Variant Inventory quantity" column.');
}

// Rebuild headers and headerMap after deletion
headers = cutSheet.getRange(1, 1, 1, cutSheet.getLastColumn()).getValues()[0];
headerMap = headers.reduce((map, val, i) => {
  map[val] = i;
  return map;
}, {});


  // --- STEP 1B: Duplicate rows for "Quantity pending fulfillment" > 1 ---
  let qtyPendingIndex = headerMap["Quantity pending fulfillment"] + 1;
  const dataRange = cutSheet.getRange(2, 1, cutSheet.getLastRow() - 1, cutSheet.getLastColumn());
  const dataValues = dataRange.getValues();

  for (let i = dataValues.length - 1; i >= 0; i--) {
    const qtyVal = parseFloat(dataValues[i][qtyPendingIndex - 1]);
    if (!isNaN(qtyVal) && qtyVal > 1) {
      const duplicatesNeeded = Math.floor(qtyVal) - 1;
      const rowIndex = i + 2; // account for header

      // Insert duplicates
      for (let j = 0; j < duplicatesNeeded; j++) {
        cutSheet.insertRowsAfter(rowIndex + j, 1);
        cutSheet.getRange(rowIndex + j + 1, 1, 1, dataValues[i].length).setValues([dataValues[i]]);
      }

      // Strikethrough original cell
      const qtyCell = cutSheet.getRange(rowIndex, qtyPendingIndex);
      const originalText = qtyCell.getDisplayValue();
      const textStyle = SpreadsheetApp.newTextStyle().setStrikethrough(true).build();
      qtyCell.setValue(originalText);
      qtyCell.setTextStyle(textStyle);
    }
  }

  Logger.log('STEP 1B: Duplicated rows for quantities > 1 and applied strikethrough.');

  // --- PATCH STEP: Fill missing Blank Color from Size ---
  let blankColorColIndex = headerMap["Blank Color"] + 1;
  let sizeColIndexForPatch = headerMap["Size"] + 1;
  const lastRowPatch = cutSheet.getLastRow();

  for (let row = 2; row <= lastRowPatch; row++) {
    const blankColor = cutSheet.getRange(row, blankColorColIndex).getValue();
    if (blankColor === "" || blankColor === null) {
      const sizeValue = cutSheet.getRange(row, sizeColIndexForPatch).getValue();
      if (typeof sizeValue === "string" && sizeValue.includes("/")) {
        const color = sizeValue.split("/")[0].trim();
        cutSheet.getRange(row, blankColorColIndex).setValue(color);
      }
    }
  }

  Logger.log('PATCH: Filled missing "Blank Color" from "Size" column.');

  // STEP 2 onward (all your existing logic)
  filterRows(cutSheet, headerMap);
  Logger.log('Step 2: Deleted rows where Cancelled is not "FALSE" OR Refunds is not 0.');

  processColumnsDEFG(cutSheet, headerMap);
  Logger.log('Step 3: Cleaned up shipping and billing name fields.');

  cutSheet.deleteColumn(headerMap["Cancelled"] + 1);

  headers = cutSheet.getRange(1, 1, 1, cutSheet.getLastColumn()).getValues()[0];
  headerMap = headers.reduce((map, val, i) => {
    map[val] = i;
    return map;
  }, {});

  cutSheet.deleteColumn(headerMap["Refunds"] + 1);

  headers = cutSheet.getRange(1, 1, 1, cutSheet.getLastColumn()).getValues()[0];
  headerMap = headers.reduce((map, val, i) => {
    map[val] = i;
    return map;
  }, {});

  if ("Quantity pending fulfillment" in headerMap) {
   // cutSheet.deleteColumn(headerMap["Quantity pending fulfillment"] + 1);
    Logger.log('Step 5: Deleted "Quantity pending fulfillment" column.');
  }

  headers = cutSheet.getRange(1, 1, 1, cutSheet.getLastColumn()).getValues()[0];
  headerMap = headers.reduce((map, val, i) => {
    map[val] = i;
    return map;
  }, {});

  concatenateAndCleanColumns(cutSheet, headerMap);
  Logger.log('Step 6: Created Proper Name column and cleaned whitespace.');

  headers = cutSheet.getRange(1, 1, 1, cutSheet.getLastColumn()).getValues()[0];
  headerMap = headers.reduce((map, val, i) => {
    map[val] = i;
    return map;
  }, {});

  cleanCollectionColumn(cutSheet, headerMap);

  const headersRow = cutSheet.getRange(1, 1, 1, cutSheet.getLastColumn()).getValues()[0];
  for (let i = 0; i < headersRow.length; i++) {
    const rawHeader = headersRow[i].toString().trim().replace(/\s+/g, " ");
    if (rawHeader === "Billing First Name Billing Last Name Shipping First Name Shipping Last Name") {
      cutSheet.getRange(1, i + 1).setValue("Proper Name");
      Logger.log(`Step 6B.2: Renamed column ${i + 1} to "Proper Name".`);
      break;
    }
  }

  addItemSizeColumn(cutSheet);
  Logger.log('Step 10: Created Item + Size column and hid source columns.');

  headers = cutSheet.getRange(1, 1, 1, cutSheet.getLastColumn()).getValues()[0];
  headerMap = headers.reduce((map, val, i) => {
    map[val] = i;
    return map;
  }, {});

  const productCol = headerMap["Product"];
  const sizeCol = headerMap["Size"];

  if (productCol !== undefined) {
    cutSheet.hideColumn(cutSheet.getRange(1, productCol + 1));
  }
  if (sizeCol !== undefined) {
    cutSheet.hideColumn(cutSheet.getRange(1, sizeCol + 1));
  }


  Logger.log('Step 10.AB: Hid "Product" and "Size" columns.');

  resizeAllColumns(cutSheet);
  Logger.log('Step 9: Resized columns.');

  stepBYEBYE(cutSheet);
  removeEmptyRowsAndColumns();

  createListSheet(ss, cutSheet);
  Logger.log('Step 11: Created LIST sheet.');

  const listSheet = ss.getSheetByName("LIST");
  listSheet.setFrozenRows(1);

  resizeAllColumns(listSheet);
  Logger.log('Step 12: Resized LIST sheet.');

  const listHeaders = listSheet.getRange(1, 1, 1, listSheet.getLastColumn()).getValues()[0];
  const itemSizeCol = listHeaders.indexOf("Item + Size") + 1;
  if (itemSizeCol === 0) throw new Error('"Item + Size" column not found in LIST sheet.');

  listSheet.getRange(2, 1, listSheet.getLastRow() - 1, listSheet.getLastColumn())
    .sort({ column: itemSizeCol, ascending: true });
  Logger.log('Step 13: Sorted LIST by "Item + Size" A→Z.');

  addUniqueAndCountColumns(listSheet);
  Logger.log('Step 14: Added Unique and Count columns.');

  resizeAllColumns(listSheet);
  Logger.log('Step 15: Resized LIST sheet after unique count.');

  stepBYEBYE(listSheet);
  removeEmptyRowsAndColumns();
  Logger.log('Step 16: Byebye useless cells.');

  createPOSheet(ss);
  Logger.log('STEP 17: Created PO sheet.');
}

// ----------------------------- HELPER FUNCTIONS -----------------------------

function filterRows(sheet, headerMap) {
  const lastRow = sheet.getLastRow();
  const cancelledCol = headerMap["Cancelled"] + 1;
  const refundsCol = headerMap["Refunds"] + 1;
  const qtyPendingCol = headerMap["Quantity pending fulfillment"] !== undefined 
                        ? headerMap["Quantity pending fulfillment"] + 1 
                        : null;

  for (let row = lastRow; row >= 2; row--) {
    const cancelled = sheet.getRange(row, cancelledCol).getValue();
    const refunds = sheet.getRange(row, refundsCol).getValue();

    const cancelledNormalized = (cancelled === false || cancelled === "FALSE");
    const refundsNormalized = (refunds === 0 || refunds === "0");

    let qtyValid = true;
    if (qtyPendingCol !== null) {
      const qtyVal = sheet.getRange(row, qtyPendingCol).getValue();
      const qtyNum = parseFloat(qtyVal);
      qtyValid = !isNaN(qtyNum) && qtyNum >= 1;
    }

    if (!cancelledNormalized || !refundsNormalized || !qtyValid) {
      sheet.deleteRow(row);
    }
  }
}



function processColumnsDEFG(sheet, headerMap) {
  const cols = [
    headerMap["Billing First Name"] + 1,
    headerMap["Billing Last Name"] + 1,
    headerMap["Shipping First Name"] + 1,
    headerMap["Shipping Last Name"] + 1
  ];
  const lastRow = sheet.getLastRow();

  for (let row = 2; row <= lastRow; row++) {
    const values = cols.map(col => sheet.getRange(row, col).getValue());
    const seen = {};
    const updatedValues = [];

    for (let i = values.length - 1; i >= 0; i--) {
      const val = values[i];
      const valStr = val?.toString().trim().toLowerCase();
      if (!val || seen[valStr]) {
        updatedValues[i] = "";
      } else {
        seen[valStr] = true;
        updatedValues[i] = val.toString().trim();
      }
    }

    const capitalized = updatedValues.map(val => val ? val.charAt(0).toUpperCase() + val.slice(1).toLowerCase() : "");
    const nonEmptyCount = capitalized.filter(val => val !== "").length;
    if (nonEmptyCount >= 3) capitalized[1] += " &";

    cols.forEach((col, i) => sheet.getRange(row, col).setValue(capitalized[i]));
  }
}

function concatenateAndCleanColumns(sheet, headerMap) {
  const nameFields = [
    'Billing First Name',
    'Billing Last Name',
    'Shipping First Name',
    'Shipping Last Name'
  ];

  nameFields.forEach(field => {
    if (!(field in headerMap)) {
      throw new Error(`Missing expected name field: ${field}`);
    }
  });

  const lastRow = sheet.getLastRow();
  const shippingLastNameCol = headerMap['Shipping Last Name'] + 1;
  const insertAt = shippingLastNameCol + 1;

  sheet.insertColumnAfter(shippingLastNameCol);
  sheet.getRange(1, insertAt).setValue('Proper Name');

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const updatedHeaderMap = headers.reduce((map, val, i) => {
    map[val] = i;
    return map;
  }, {});

  const colIndices = nameFields.map(name => updatedHeaderMap[name] + 1);
  const nameData = colIndices.map(col => sheet.getRange(2, col, lastRow - 1).getValues());

  const merged = [];
  for (let i = 0; i < nameData[0].length; i++) {
    const parts = nameData.map(col => col[i][0] || '').join(' ');
    merged.push([parts.replace(/\s+/g, ' ').trim()]);
  }

  sheet.getRange(2, insertAt, merged.length, 1).setValues(merged);

  const deleteIndices = nameFields.map(name => updatedHeaderMap[name] + 1).sort((a, b) => b - a);
  deleteIndices.forEach(col => sheet.deleteColumn(col));
}


function addItemSizeColumn(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();
  const productColIndex = headers.indexOf("Product");
  const variantColIndex = headers.indexOf("Size");

  if (productColIndex === -1 || variantColIndex === -1) {
    throw new Error('Could not find "Product" or "Size" column.');
  }

  const insertCol = variantColIndex + 2;
  sheet.insertColumnAfter(variantColIndex + 1);
  sheet.getRange(1, insertCol).setValue("Item + Size");

  const productValues = sheet.getRange(2, productColIndex + 1, lastRow - 1).getValues();
  const variantValues = sheet.getRange(2, variantColIndex + 1, lastRow - 1).getValues();

  const combined = productValues.map((row, i) => {
    const product = row[0] ? row[0].toString().trim() : "";
    const variant = variantValues[i][0] ? variantValues[i][0].toString().trim() : "";
    return [`${product} - ${variant}`];
  });

  sheet.getRange(2, insertCol, combined.length, 1).setValues(combined);
}

function resizeAllColumns(sheet) {
  const lastCol = sheet.getLastColumn();
  for (let i = 1; i <= lastCol; i++) {
    sheet.autoResizeColumn(i);
  }
}

function stepBYEBYE(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // Hide fully empty columns (except header)
  for (let col = 0; col < lastCol; col++) {
    let isColEmpty = true;
    for (let row = 1; row < lastRow; row++) {
      if (data[row][col] !== "" && data[row][col] !== null) {
        isColEmpty = false;
        break;
      }
    }
    if (isColEmpty) {
      sheet.hideColumn(sheet.getRange(1, col + 1));
    }
  }

  // Hide fully empty rows (starting after header)
  for (let row = 1; row < lastRow; row++) {
    let isRowEmpty = true;
    for (let col = 0; col < lastCol; col++) {
      if (data[row][col] !== "" && data[row][col] !== null) {
        isRowEmpty = false;
        break;
      }
    }
    if (isRowEmpty) {
      sheet.hideRow(sheet.getRange(row + 1, 1));
    }
  }

  Logger.log("STEP BYEBYE: Hidden fully empty columns and rows.");
}

function removeEmptyRowsAndColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const maxRows = sheet.getMaxRows();
  const lastRow = sheet.getLastRow();
  const rowsToDelete = maxRows - lastRow;
  if (rowsToDelete > 0) {
    sheet.deleteRows(lastRow + 1, rowsToDelete);
    Logger.log(`Deleted ${rowsToDelete} trailing empty row(s).`);
  }

  const maxCols = sheet.getMaxColumns();
  const lastCol = sheet.getLastColumn();
  const colsToDelete = maxCols - lastCol;
  if (colsToDelete > 0) {
    sheet.deleteColumns(lastCol + 1, colsToDelete);
    Logger.log(`Deleted ${colsToDelete} trailing empty column(s).`);
  }
}


function createListSheet(ss, sourceSheet) {
  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const lastRow = sourceSheet.getLastRow();
  const required = ["Order #", "Proper Name", "SKU", "Item + Size"];
  const colIndices = {};

  required.forEach(header => {
    const index = headers.indexOf(header);
    if (index === -1) throw new Error(`Missing header in CUT sheet: ${header}`);
    colIndices[header] = index + 1;
  });

  const existing = ss.getSheetByName("LIST");
  if (existing) ss.deleteSheet(existing);

  const listSheet = ss.insertSheet("LIST");
  listSheet.appendRow(required);

  const data = [];
  for (let row = 2; row <= lastRow; row++) {
    const rowData = required.map(h => sourceSheet.getRange(row, colIndices[h]).getValue());
    data.push(rowData);
  }

  listSheet.getRange(2, 1, data.length, required.length).setValues(data);
}

function addUniqueAndCountColumns(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const itemSizeColIndex = headers.indexOf("Item + Size") + 1;
  if (itemSizeColIndex === 0) {
    throw new Error('"Item + Size" column not found in LIST sheet.');
  }

  const lastRow = sheet.getLastRow();
  const uniqueCol = itemSizeColIndex + 1;
  const countCol = itemSizeColIndex + 2;

  sheet.getRange(1, uniqueCol).setValue("Unique");
  sheet.getRange(1, countCol).setValue("Count");

  const itemColA1 = getA1Notation(itemSizeColIndex);
  const uniqueFormula = `=UNIQUE(${itemColA1}2:${itemColA1})`;
  sheet.getRange(2, uniqueCol).setFormula(uniqueFormula);

  SpreadsheetApp.flush();

  const uniqueValues = sheet.getRange(2, uniqueCol, sheet.getMaxRows() - 1).getValues();
  const uniqueCount = uniqueValues.findIndex(row => !row[0]);
  const fillCount = uniqueCount === -1 ? uniqueValues.length : uniqueCount;

  if (fillCount > 0) {
    const countFormula = `=IF(${getA1Notation(uniqueCol)}2="", "", COUNTIF(${itemColA1}:${itemColA1}, ${getA1Notation(uniqueCol)}2))`;
    const range = sheet.getRange(2, countCol);
    range.setFormula(countFormula);
    range.autoFill(sheet.getRange(2, countCol, fillCount), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  }

  resizeAllColumns(sheet);
}

function getA1Notation(colIndex) {
  let notation = "";
  while (colIndex > 0) {
    const remainder = (colIndex - 1) % 26;
    notation = String.fromCharCode(65 + remainder) + notation;
    colIndex = Math.floor((colIndex - 1) / 26);
  }
  return notation;
}



function createPOSheet(ss) {
  const cutSheet = ss.getSheetByName("CUT");
  if (!cutSheet) throw new Error('CUT sheet not found.');

  const headers = cutSheet.getRange(1, 1, 1, cutSheet.getLastColumn()).getValues()[0];
  const lastRow = cutSheet.getLastRow();

  const requiredCols = ["SKU", "Item + Size", "Vendor", "Style #", "Blank Type", "Blank Color"];
  const optionalCols = ["Front", "Back", "Left", "Right"];
  const countableCols = [...requiredCols.slice(2), "Complete Item Name"]; // Unique + Count fields

  const headerMap = headers.reduce((map, val, i) => {
    map[val] = i;
    return map;
  }, {});

  for (let col of requiredCols) {
    if (!(col in headerMap)) {
      throw new Error(`Missing required column in CUT sheet: ${col}`);
    }
  }

  const colIndices = [...requiredCols, ...optionalCols.filter(col => col in headerMap)].map(col => headerMap[col]);
  const data = cutSheet.getRange(2, 1, lastRow - 1, cutSheet.getLastColumn()).getValues();

  const processedData = data.map(row => {
    const values = colIndices.map(i => row[i] || "");
    return values;
  });

  const existingPO = ss.getSheetByName("PO");
  if (existingPO) ss.deleteSheet(existingPO);

  const poSheet = ss.insertSheet("PO");

  const poHeaders = [...requiredCols, ...optionalCols.filter(col => col in headerMap)];
  poSheet.appendRow([...poHeaders, "Complete Item Name"]);

  // Add values and build Complete Item Name
  const poDataWithName = processedData.map(row => {
    const [vendor, style, type, color] = [
      row[requiredCols.indexOf("Vendor")],
      row[requiredCols.indexOf("Style #")],
      row[requiredCols.indexOf("Blank Type")],
      row[requiredCols.indexOf("Blank Color")]
    ];
    const completeName = `${vendor} ${style} ${type} - ${color}`.trim();
    return [...row, completeName];
  });

  if (poDataWithName.length > 0) {
    poSheet.getRange(2, 1, poDataWithName.length, poHeaders.length + 1).setValues(poDataWithName);
  }

  const lastDataRow = poSheet.getLastRow();

  const fieldToIndexMap = poSheet.getRange(1, 1, 1, poSheet.getLastColumn()).getValues()[0]
    .reduce((map, name, i) => {
      map[name] = i + 1;
      return map;
    }, {});

  // Unique + Count columns
  const uniqueCountTargets = ["Item + Size", "SKU", ...countableCols];

  let currentCol = poSheet.getLastColumn() + 1;

  uniqueCountTargets.forEach(field => {
    poSheet.getRange(1, currentCol).setValue(`${field} Unique`);
    poSheet.getRange(1, currentCol + 1).setValue(`${field} Count`);

    const dataColA1 = getA1Notation(fieldToIndexMap[field]);

    poSheet.getRange(2, currentCol).setFormula(`=UNIQUE(${dataColA1}2:${dataColA1}${lastDataRow})`);
    SpreadsheetApp.flush();

    const uniqueValues = poSheet.getRange(2, currentCol, poSheet.getMaxRows() - 1).getValues();
    const uniqueCount = uniqueValues.findIndex(row => !row[0]);
    const fillCount = uniqueCount === -1 ? uniqueValues.length : uniqueCount;

    if (fillCount > 0) {
      const countFormula = `=IF(${getA1Notation(currentCol)}2="", "", COUNTIF(${dataColA1}:${dataColA1}, ${getA1Notation(currentCol)}2))`;
      poSheet.getRange(2, currentCol + 1).setFormula(countFormula);
      poSheet.getRange(2, currentCol + 1).autoFill(
        poSheet.getRange(2, currentCol + 1, fillCount),
        SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES
      );
    }

    currentCol += 2;
  });

  resizeAllColumns(poSheet);

    // STEP: Final polish

  // Freeze header row
  poSheet.setFrozenRows(1);

  // Hide all columns before "Complete Item Name Unique"
  const finalHeaders = poSheet.getRange(1, 1, 1, poSheet.getLastColumn()).getValues()[0];
  const hideUntilIndex = finalHeaders.indexOf("Complete Item Name Unique");
  if (hideUntilIndex !== -1) {
    for (let i = 1; i <= hideUntilIndex; i++) {
      poSheet.hideColumn(poSheet.getRange(1, i));
    }
  }

  // Resize all columns
  resizeAllColumns(poSheet);

  // Remove empty rows and columns
  stepBYEBYE(poSheet);
  removeEmptyRowsAndColumns();

}




function cleanCollectionColumn(sheet, headerMap) {
  if (!("Collection" in headerMap)) return;

  const colIndex = headerMap["Collection"] + 1;
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, colIndex, lastRow - 1).getValues();

  const cleaned = values.map(row => {
    let val = row[0];
    if (typeof val === "string" && val.startsWith("All,")) {
      val = val.replace(/^All,\s*/, ""); // Remove "All," and optional space
    }
    return [val];
  });

  sheet.getRange(2, colIndex, cleaned.length, 1).setValues(cleaned);
  Logger.log('Cleaned "Collection" column to remove "All," prefix.');
}
