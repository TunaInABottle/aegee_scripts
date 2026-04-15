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
 */
function myCreateSpreadsheet() {
  prepareSplit("YOUR_DRIVE_FOLDER_ID", "August", 2026);
}

/**
 * Reads the "Data to move" sheet, creates a child folder under the given
 * parent, and returns the list of unique European Body names found.
 *
 * The child folder is named "{year}-{MM} (Delete 30 {delMonth} {delYear})".
 * Does NOT create any spreadsheets — call processBody() for each name returned.
 *
 * @param {string} folderId - Google Drive folder ID where the child folder will be created.
 * @param {string} month    - Full month name, e.g. "August".
 * @param {number|string} year - Four-digit year, e.g. 2026.
 *
 * @returns {{ bodyNames: string[], childFolderId: string }}
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

  // Copy the full data into the child folder as a single spreadsheet
  let fullSs = SpreadsheetApp.create("Open Call " + month + " - All Data");
  let fullSheet = fullSs.getSheets()[0];

  for (let i = 0; i < data.length; i++) {
    fullSheet.appendRow(data[i]);
  }

  let fullHeaderRange = fullSheet.getRange(1, 1, 1, headers.length);
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

  let fullFile = DriveApp.getFileById(fullSs.getId());
  childFolder.addFile(fullFile);
  DriveApp.getRootFolder().removeFile(fullFile);

  let ebBodyIndex1 = headers.indexOf("Which European Body do you wish to apply for?");
  let ebBodyIndex2 = headers.indexOf("To which European Body do you wish to apply for?");

  let ebBodyIndexes = [];
  if (ebBodyIndex1 !== -1) {
    ebBodyIndexes.push(ebBodyIndex1);
  } else {
    Logger.log("idx1 not found");
  }
  if (ebBodyIndex2 !== -1 && ebBodyIndex1 !== ebBodyIndex2) {
    ebBodyIndexes.push(ebBodyIndex2);
  } else {
    Logger.log("idx2 not found");
  }

  if (ebBodyIndexes.length === 0) {
    throw new Error("Error: No European Body column found.");
  }

  let ebBodies = {};
  for (let j = 0; j < ebBodyIndexes.length; j++) {
    let ebBodyHeaderIdx = ebBodyIndexes[j];
    for (let i = 1; i < data.length; i++) {
      let ebBody = data[i][ebBodyHeaderIdx];
      if (ebBody && !ebBodies[ebBody]) {
        ebBodies[ebBody] = [];
      }
      if (ebBody) {
        ebBodies[ebBody].push(data[i]);
      }
    }
  }

  return {
    bodyNames: Object.keys(ebBodies),
    childFolderId: childFolder.getId()
  };
}

/**
 * Creates a single spreadsheet for one European Body, reading from the
 * "Data to move" sheet in the active spreadsheet.
 *
 * The new spreadsheet is moved into the specified child folder.
 *
 * @param {string} bodyName      - Exact European Body name to filter rows for.
 * @param {string} month         - Full month name, used in the spreadsheet title.
 * @param {string} childFolderId - Google Drive folder ID to place the result in.
 *
 * @returns {string} The name of the created spreadsheet.
 * @throws {Error} If the sheet "Data to move" is not found.
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

    // Find EB body columns
    let ebBodyIndex1 = headers.indexOf("Which European Body do you wish to apply for?");
    let ebBodyIndex2 = headers.indexOf("To which European Body do you wish to apply for?");
    let ebBodyIndexes = [];
    if (ebBodyIndex1 !== -1) ebBodyIndexes.push(ebBodyIndex1);
    if (ebBodyIndex2 !== -1 && ebBodyIndex1 !== ebBodyIndex2) ebBodyIndexes.push(ebBodyIndex2);

    if (ebBodyIndexes.length === 0) {
      throw new Error("Error: No European Body column found.");
    }

    // Collect rows matching this body
    let rows = [];
    for (let j = 0; j < ebBodyIndexes.length; j++) {
      let ebBodyHeaderIdx = ebBodyIndexes[j];
      for (let i = 1; i < data.length; i++) {
        if (data[i][ebBodyHeaderIdx] === bodyName) {
          rows.push(data[i]);
        }
      }
    }

    // Create spreadsheet
    let newSs = SpreadsheetApp.create("Open Call " + month + ": " + bodyName);
    Logger.log("Creating SS -- " + month + " Result: " + bodyName);
    let newSheet = newSs.getSheets()[0];

    newSheet.appendRow(headers);
    for (let r = 0; r < rows.length; r++) {
      newSheet.appendRow(rows[r]);
    }

    // Bold headers, freeze first row
    let headerRange = newSheet.getRange(1, 1, 1, headers.length);
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
    let newFile = DriveApp.getFileById(newSs.getId());
    DriveApp.getFolderById(childFolderId).addFile(newFile);
    DriveApp.getRootFolder().removeFile(newFile);

    return { success: true, bodyName: bodyName };
  } catch (e) {
    Logger.log("Error processing '" + bodyName + "': " + e.message);
    return { success: false, bodyName: bodyName, error: e.message };
  }
}
