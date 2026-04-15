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
    .setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Data Splitter');
}

/**
 * Creates spreadsheets split by European Body, reading from the
 * "Data to move" sheet in the active spreadsheet and placing results
 * in a new child folder under the given parent folder.
 *
 * @param {string} folderId  - Google Drive folder ID (or full URL) where
 *                             the child folder will be created.
 * @param {string} month     - Full month name, e.g. "August".
 * @param {number|string} year - Four-digit year, e.g. 2026.
 *
 * @returns {void} Creates one spreadsheet per European Body inside a new
 *                 child folder named "{year}-{MM} (Delete 30 {delMonth} {delYear})".
 *
 * @throws {Error} If the sheet "Data to move" is not found.
 * @throws {Error} If no European Body column is found in the headers.
 */
function createSpreadsheetsFromFolderAndSpreadsheet(folderId, month, year) {
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

  for (let ebBody in ebBodies) {
    let newSs = SpreadsheetApp.create("Open Call " + month + ": " + ebBody);
    Logger.log("Creating SS -- " + month + " Result: " + ebBody);
    let newSheet = newSs.getSheets()[0];

    newSheet.appendRow(headers);

    let rows = ebBodies[ebBody];
    for (let r = 0; r < rows.length; r++) {
      newSheet.appendRow(rows[r]);
    }

    let headerRange = newSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    newSheet.setFrozenRows(1);

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

    let newFile = DriveApp.getFileById(newSs.getId());
    childFolder.addFile(newFile);
    DriveApp.getRootFolder().removeFile(newFile);
  }
}
