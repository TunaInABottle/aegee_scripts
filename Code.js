/**
 * Adds a custom menu to the spreadsheet UI on open.
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts Menu')
    .addItem('Create EB Spreadsheets', 'openDialog')
    .addToUi();
}

/**
 * Opens a modal dialog for the user to input the folder URL, month, and year.
 */
function openDialog() {
  let htmlOutput = HtmlService.createHtmlOutputFromFile('Form')
    .setWidth(400)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Data Splitter');
}

/**
 * Quick-test entry point (bypasses the dialog).
 * Uses CONFIG.driveFolderId from the local Config.js file
 * (see Config.example.js; Config.js is kept out of git via .gitignore).
 */
function myCreateSpreadsheet() {
  prepareSplit(CONFIG.driveFolderId, "August", 2026);
}

/**
 * Returns all column indexes in a row where the value matches the target string.
 * @param {Array} row - Array of header values.
 * @param {string} target - Value to search for.
 * @returns {number[]}
 */
function findAllIndexes(row, target) {
  let indexes = [];
  for (let i = 0; i < row.length; i++) {
    if (row[i] === target) indexes.push(i);
  }
  return indexes;
}

/**
 * Builds a mapping from original columns to merged columns.
 * All EB columns are removed and replaced by a single "European Body Application" column.
 *
 * @param {string[]} headers - Original header row.
 * @param {number[]} allEbIndexes - Column indexes of all EB variants.
 * @returns {{ newHeaders: string[], colMap: number[] }}
 *   colMap[i] maps each new column index to the original column index,
 *   except the last entry which is -1 (the merged EB column).
 */
function buildMergedHeaders(headers, allEbIndexes) {
  let newHeaders = [];
  let colMap = [];
  for (let i = 0; i < headers.length; i++) {
    if (allEbIndexes.indexOf(i) === -1) {
      newHeaders.push(headers[i]);
      colMap.push(i);
    }
    // Insert "European Body Application" at position 4 (index 3)
    if (newHeaders.length === 3 && colMap.length === 3) {
      newHeaders.push("European Body Application");
      colMap.push(-1); // sentinel: merged EB value
    }
  }
  // Fallback: if fewer than 2 non-EB columns, append at the end
  if (colMap.indexOf(-1) === -1) {
    newHeaders.push("European Body Application");
    colMap.push(-1);
  }
  return { newHeaders: newHeaders, colMap: colMap };
}

/**
 * Transforms a row using the merged column mapping, inserting the EB value.
 *
 * @param {Array} originalRow - The source row.
 * @param {number[]} colMap - Column map from buildMergedHeaders.
 * @param {string} ebValue - The merged European Body value for this row.
 * @returns {Array}
 */
function mergeRow(originalRow, colMap, ebValue) {
  let newRow = [];
  for (let i = 0; i < colMap.length; i++) {
    newRow.push(colMap[i] === -1 ? ebValue : originalRow[colMap[i]]);
  }
  return newRow;
}

/**
 * Removes the first column (Timestamp) from headers and every row.
 * @param {string[]} headers
 * @param {Array[]} rows
 * @returns {{ headers: string[], rows: Array[] }}
 */
function dropFirstColumn(headers, rows) {
  let newHeaders = headers.slice(1);
  let newRows = [];
  for (let i = 0; i < rows.length; i++) {
    newRows.push(rows[i].slice(1));
  }
  return { headers: newHeaders, rows: newRows };
}

/**
 * Removes columns whose header is empty or contains "I agree" (privacy consent).
 * @param {string[]} headers
 * @param {Array[]} rows
 * @returns {{ headers: string[], rows: Array[] }}
 */
function dropPrivacyConsentColumns(headers, rows) {
  let keepIndexes = [];
  for (let i = 0; i < headers.length; i++) {
    let h = headers[i].toString().trim();
    if (h === "") continue;
    if (h.toLowerCase().indexOf("i agree") !== -1) continue;
    keepIndexes.push(i);
  }

  let newHeaders = [];
  for (let i = 0; i < keepIndexes.length; i++) {
    newHeaders.push(headers[keepIndexes[i]]);
  }

  let newRows = [];
  for (let r = 0; r < rows.length; r++) {
    let newRow = [];
    for (let i = 0; i < keepIndexes.length; i++) {
      newRow.push(rows[r][keepIndexes[i]]);
    }
    newRows.push(newRow);
  }

  return { headers: newHeaders, rows: newRows };
}

/**
 * Moves a file into a target folder by copying it there and trashing the original.
 * Works on both My Drive and shared drives without requiring the Drive API service.
 *
 * @param {string} fileId         - ID of the file to move.
 * @param {string} targetFolderId - ID of the destination folder.
 */
function moveFileToFolder(fileId, targetFolderId) {
  let file = DriveApp.getFileById(fileId);
  let targetFolder = DriveApp.getFolderById(targetFolderId);
  file.makeCopy(file.getName(), targetFolder);
  file.setTrashed(true);
}

/**
 * Deletes a folder and trashes all its contents.
 *
 * @param {string} folderId - ID of the folder to delete.
 */
function deleteFolder(folderId) {
  let folder = DriveApp.getFolderById(folderId);
  let files = folder.getFiles();
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
  folder.setTrashed(true);
}

/**
 * Reads the "Data to move" sheet, creates a child folder under the given
 * parent, copies all data into it, and returns the list of unique European
 * Body names found after merging all EB columns.
 *
 * Both header strings ("Which European Body..." and "To which European Body...")
 * can appear multiple times. All matching columns are merged: for each row,
 * at most one distinct value is expected across all EB columns. If multiple
 * different values appear in the same row, a conflict is reported.
 *
 * The child folder is named "{year}-{MM} (Delete 30 {delMonth} {delYear})".
 *
 * @param {string} folderId - Google Drive folder ID where the child folder will be created.
 * @param {string} month    - Full month name, e.g. "August".
 * @param {number|string} year - Four-digit year, e.g. 2026.
 *
 * @returns {{ bodyNames: string[], childFolderId: string } |
 *            { conflict: true, conflicts: Array<{row: number, values: string[]}> }}
 * @throws {Error} If the sheet "Data to move" is not found.
 * @throws {Error} If no European Body column is found in the headers.
 */
function prepareSplit(folderId, month, year) {
  const MONTHS = ["January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"];

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Data to move");
  if (!sheet) {
    throw new Error("Sheet 'Data to move' not found in this spreadsheet.");
  }

  let parentFolder = DriveApp.getFolderById(folderId);

  // Build child folder name: "2026-03 (Delete 30 June 2026)"
  let monthIndex = MONTHS.indexOf(month);
  let monthNum = monthIndex + 1;
  let monthStr = (monthNum < 10 ? "0" : "") + monthNum;
  let delMonthIndex = (monthIndex + 3) % 12;
  let delYear = (monthIndex + 3 >= 12) ? parseInt(year) + 1 : parseInt(year);
  let folderName = year + "-" + monthStr + " (Delete 30 " + MONTHS[delMonthIndex] + " " + delYear + ")";
  let childFolder = parentFolder.createFolder(folderName);

  let data = sheet.getDataRange().getValues();
  let headers = data[0];

  // Find ALL indexes for each EB header variant
  let col1Indexes = findAllIndexes(headers, "Which European Body do you wish to apply for?");
  let col2Indexes = findAllIndexes(headers, "To which European Body do you wish to apply for?");
  let allEbIndexes = col1Indexes.concat(col2Indexes);

  if (allEbIndexes.length === 0) {
    throw new Error("No European Body column found in the headers.");
  }

  // Merge all EB columns, checking for conflicts and missing values
  let ebBodies = {};
  let matchedRows = []; // rows that will be split
  let conflicts = [];
  let missingEbRows = [];

  for (let i = 1; i < data.length; i++) {
    let rowValues = [];
    for (let j = 0; j < allEbIndexes.length; j++) {
      let val = data[i][allEbIndexes[j]];
      if (val) rowValues.push(val.toString().trim());
    }

    if (rowValues.length === 0) {
      missingEbRows.push(i + 1);
      continue;
    }

    // Check if all found values are the same
    let uniqueValues = [];
    for (let v = 0; v < rowValues.length; v++) {
      if (uniqueValues.indexOf(rowValues[v]) === -1) {
        uniqueValues.push(rowValues[v]);
      }
    }

    if (uniqueValues.length > 1) {
      conflicts.push({ row: i + 1, values: rowValues });
      continue;
    }

    let ebBody = uniqueValues[0];
    if (!ebBodies[ebBody]) {
      ebBodies[ebBody] = [];
    }
    ebBodies[ebBody].push(data[i]);
    matchedRows.push({ original: data[i], ebValue: ebBody });
  }

  if (conflicts.length > 0 || missingEbRows.length > 0) {
    // Clean up the child folder we just created
    deleteFolder(childFolder.getId());

    return { conflict: true, conflicts: conflicts, missingEbRows: missingEbRows };
  }

  // Copy only the split-worthy data into the child folder (with merged EB column)
  let merged = buildMergedHeaders(headers, allEbIndexes);

  let fullHeaders = merged.newHeaders;
  let fullRows = [];
  for (let i = 0; i < matchedRows.length; i++) {
    fullRows.push(mergeRow(matchedRows[i].original, merged.colMap, matchedRows[i].ebValue));
  }

  // Drop the Timestamp column (first column)
  let noTimestamp = dropFirstColumn(fullHeaders, fullRows);
  let cleaned = dropPrivacyConsentColumns(noTimestamp.headers, noTimestamp.rows);

  let fullSs = SpreadsheetApp.create("Open Call " + month + " - All Data");
  let fullSheet = fullSs.getSheets()[0];

  fullSheet.appendRow(cleaned.headers);
  for (let i = 0; i < cleaned.rows.length; i++) {
    fullSheet.appendRow(cleaned.rows[i]);
  }

  let fullHeaderRange = fullSheet.getRange(1, 1, 1, cleaned.headers.length);
  fullHeaderRange.setFontWeight("bold");
  fullSheet.setFrozenRows(1);

  // Remove entirely empty columns
  let fullValues = fullSheet.getDataRange().getValues();
  let emptyCols = [];
  for (let col = 0; col < fullValues[0].length; col++) {
    let empty = true;
    for (let row = 1; row < fullValues.length; row++) {
      if (fullValues[row][col] !== "") {
        empty = false;
        break;
      }
    }
    if (empty) {
      emptyCols.push(col);
    }
  }
  for (let k = emptyCols.length - 1; k >= 0; k--) {
    fullSheet.deleteColumn(emptyCols[k] + 1);
  }

  moveFileToFolder(fullSs.getId(), childFolder.getId());

  return {
    bodyNames: Object.keys(ebBodies),
    childFolderId: childFolder.getId()
  };
}

/**
 * Creates a single spreadsheet for one European Body, reading from the
 * "Data to move" sheet in the active spreadsheet.
 *
 * Rows are matched by checking all EB columns (merged logic): a row
 * belongs to this body if any EB column matches and all non-empty values
 * are the same.
 *
 * @param {string} bodyName      - Exact European Body name to filter rows for.
 * @param {string} month         - Full month name, used in the spreadsheet title.
 * @param {string} childFolderId - Google Drive folder ID to place the result in.
 *
 * @returns {{ success: boolean, bodyName: string, error?: string }}
 */
function processBody(bodyName, month, childFolderId) {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Data to move");
    if (!sheet) {
      throw new Error("Sheet 'Data to move' not found in this spreadsheet.");
    }

    let data = sheet.getDataRange().getValues();
    let headers = data[0];

    let col1Indexes = findAllIndexes(headers, "Which European Body do you wish to apply for?");
    let col2Indexes = findAllIndexes(headers, "To which European Body do you wish to apply for?");
    let allEbIndexes = col1Indexes.concat(col2Indexes);
    let merged = buildMergedHeaders(headers, allEbIndexes);

    // Collect rows matching this body (merged column logic)
    let rawRows = [];
    for (let i = 1; i < data.length; i++) {
      for (let j = 0; j < allEbIndexes.length; j++) {
        let val = data[i][allEbIndexes[j]];
        if (val && val.toString().trim() === bodyName) {
          rawRows.push(mergeRow(data[i], merged.colMap, bodyName));
          break; // Don't add the same row twice
        }
      }
    }

    // Drop the Timestamp column (first column)
    let noTimestamp = dropFirstColumn(merged.newHeaders, rawRows);
    let cleaned = dropPrivacyConsentColumns(noTimestamp.headers, noTimestamp.rows);

    // Create spreadsheet
    let newSs = SpreadsheetApp.create("Open Call " + month + ": " + bodyName);
    Logger.log("Creating SS -- " + month + " Result: " + bodyName);
    let newSheet = newSs.getSheets()[0];

    newSheet.appendRow(cleaned.headers);
    for (let r = 0; r < cleaned.rows.length; r++) {
      newSheet.appendRow(cleaned.rows[r]);
    }

    // Bold headers, freeze first row
    let headerRange = newSheet.getRange(1, 1, 1, cleaned.headers.length);
    headerRange.setFontWeight("bold");
    newSheet.setFrozenRows(1);

    // Remove entirely empty columns
    let range = newSheet.getDataRange();
    let values = range.getValues();
    let emptyColumns = [];
    for (let col = 0; col < values[0].length; col++) {
      let empty = true;
      for (let row = 1; row < values.length; row++) {
        if (values[row][col] !== "") {
          empty = false;
          break;
        }
      }
      if (empty) {
        emptyColumns.push(col);
      }
    }
    for (let k = emptyColumns.length - 1; k >= 0; k--) {
      newSheet.deleteColumn(emptyColumns[k] + 1);
    }

    // Move file into the child folder
    moveFileToFolder(newSs.getId(), childFolderId);

    return { success: true, bodyName: bodyName };
  } catch (e) {
    Logger.log("Error processing '" + bodyName + "': " + e.message);
    return { success: false, bodyName: bodyName, error: e.message };
  }
}

/**
 * Deletes all data rows (keeps header) from the "Data to move" sheet.
 * Called after a fully successful split.
 */
function clearSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Data to move");
  if (!sheet) return;
  let lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
}
