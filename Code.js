function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts Menu')
    .addItem('Create EB Spreadsheets', 'openDialog')
    .addToUi();
}

function openDialog() {
  let htmlOutput = HtmlService.createHtmlOutputFromFile('Form')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Enter Folder and Spreadsheet IDs');
}

function myCreateSpreadsheet() {
  createSpreadsheetsFromFolderAndSpreadsheet("YOUR_DRIVE_FOLDER_ID", "YOUR_SPREADSHEET_ID", "August")
}

function createSpreadsheetsFromFolderAndSpreadsheet(folderId, spreadsheetId, month) {
  /**
   * month (str): eg July
   */
  let ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheets()[0];
  
  let parentFolder = DriveApp.getFolderById(folderId);
  
  let data = sheet.getDataRange().getValues();
  let headers = data[0];
  let ebBodyIndex1 = headers.indexOf("Which European Body do you wish to apply for?");
  let ebBodyIndex2 = headers.indexOf("To which European Body do you wish to apply for?");

  /**
  // Choose the first valid index, if any. If both are -1, it will stay -1.
  let ebBodyIndex = (ebBodyIndex1 !== -1) ? ebBodyIndex1 : ebBodyIndex2;

  if (ebBodyIndex === -1) {
    Logger.log("EB_Body column not found.");
    return;
  }
  */

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

  // Check if differingIndexes has any items
  if (ebBodyIndexes.length === 0) {
    throw new Error("Error: No European Body column found.");
  }

  // add each row to the matching european body of application
  let ebBodies = {};
  for (let j = 0; j < ebBodyIndexes.length; j++) {
    ebBodyHeaderIdx = ebBodyIndexes[j];
    for (let i = 1; i < data.length; i++) {
      let ebBody = data[i][ebBodyHeaderIdx];
      // add object ebBody if not existent
      if (ebBody && !ebBodies[ebBody]) {
        ebBodies[ebBody] = [];
      }
      if (ebBody) {
        ebBodies[ebBody].push(data[i]);
      }
    }
  }
  
  // iterate each EBbody
  for (let ebBody in ebBodies) {
    let newSs = SpreadsheetApp.create("Open Call " + month + ": " + ebBody);
    Logger.log("Creating SS -- " + month + " Result: " + ebBody);
    let newSheet = newSs.getSheets()[0];
    
    newSheet.appendRow(headers);
    
    let rows = ebBodies[ebBody];
    for (let j = 0; j < rows.length; j++) {
      newSheet.appendRow(rows[j]);
    }
    
    // set headers in new spreadsheet
    let headerRange = newSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    newSheet.setFrozenRows(1);
    
    let range = newSheet.getDataRange();
    let values = range.getValues();
    
    // Removes columns that have no answer
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
    parentFolder.addFile(newFile);
    DriveApp.getRootFolder().removeFile(newFile);
  }
}