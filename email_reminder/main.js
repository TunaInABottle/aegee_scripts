/**
 * reads emails from the spreadsheet and removes ones labeled as ignored,
 * it then proceeds to send an email of unanswered message if there is any
 */
function mainThreadFunc() {
  // read from spreadsheet
  threadsInSheet = getThreadsFromSheet(SHEET);

  // read from email
  unansweredThreads = getUnansweredThreads(100);

  // Remove EmailThreads whose ID is in the blacklist
  let filteredThreads = unansweredThreads.filter(thread =>
    !threadsInSheet["ignored"].some(blacklistThread => blacklistThread.id === thread.id)
  );

  // write to spreadsheet
  writeThreadsToSs(SHEET, filteredThreads)

  // send email
  if (filteredThreads.length > 0) {
    sendUnansweredThreadEmail(filteredThreads);
  }
}

/**
 * Updates spreadsheet unanswered emails
 */
function updateUnansweredEmailsInSheet() {
  // read from spreadsheet
  threadsInSheet = getThreadsFromSheet(SHEET);

  // read from email
  unansweredThreads = getUnansweredThreads(100);

  // Remove EmailThreads whose ID is in the blacklist
  let filteredThreads = unansweredThreads.filter(thread =>
    !threadsInSheet["ignored"].some(blacklistThread => blacklistThread.id === thread.id)
  );

  // write to spreadsheet
  writeThreadsToSs(SHEET, filteredThreads)
}

/**
 * Gets all unanswered threads from `observedMail()` which
 * are not flagged with label `SKIP_LAB`
 * 
 * @param {numeric} MaxQueryLen - how many entries should be fetched at most (default: 50)
 * 
 * @return {EmailThread[]} - list of threads matching this query
 */
function getUnansweredThreads(maxQueryLen = 50) {
  let threads = GmailApp.search(`in:inbox to:${observedMail()} -label:${SKIP_LAB}`, 0, maxQueryLen);
  let unansweredThreads = [];

  threads.forEach(thread => {
    let _thread = EmailThread.fromGoogleAPI(thread);

    if (_thread.last_message.recipient.includes(observedMail()) &&
    !_thread.last_message.sender.includes(observedMail())) {
      unansweredThreads.push(_thread);
    }
  });

  return unansweredThreads;
}

/**
 * Write threads in `SHEET`
 * 
 * @param {Sheet} sheet - in which sheet should the data be written
 * @param {EmailThread[]} threads -list of threads to be written in `SHEET`
 */
function writeThreadsToSs(sheet, threads) {

  // remove non-header rows
  const oldRange = sheet.getDataRange();
  if (oldRange.getNumRows() != 1) {
    oldRange.offset(1, 0, oldRange.getNumRows() - 1).clearContent();
  }

  let thread_ss = []
  threads.forEach(thread => {
    thread_ss.push(thread.toSpreadsheet());
  })

  let numRows = thread_ss.length;
  if (numRows == 0) {
    Logger.log("No emails today");
    return;
  };
  let numCols = thread_ss[0].length;
  let new_range = sheet.getRange(2, 1, numRows, numCols);
  new_range.setValues(thread_ss);
}

/**
 * Converts thread in a spreadsheet to objects
 * 
 * @param {Sheet} sheet - in which sheet should the data be read
 * @return {Object<String, EmailThread[]>} - a division among threads that have to be ignored
 *    and still relevant
 */
function getThreadsFromSheet(sheet) {
  let threadObjs = readObjsFromSpreadsheet(sheet);

  let threads = {};
  threads["ignored"] = [];
  threads["relevant"] = [];

  threadObjs.forEach((entry) => {
    // skip if there is no ID
    if (!entry["ID"]) {
      return;
    }
    let thread = EmailThread.fromGoogleThreadID(entry["ID"]);
    if (entry["Ignore"] === true) {
      thread.setIgnoreLabel();
      threads["ignored"].push(thread);
      //Logger.log(`thread ${thread.subject} got ignored`);
    } else {
      threads["relevant"].push(thread);
      //Logger.log(`thread ${thread.subject} got added to relevant`);
    }
  });
  return threads
}

/**
 * Convert `SHEET` range intervals to a list of generic objects
 * 
 * @param {Sheet} sheet - a spreadsheet sheet 
 * @return {Array<Object>} a list of generic objects were headers are the keys
 *  and the entries are the values
 */
function readObjsFromSpreadsheet(sheet) { 
  let dataRange = sheet.getDataRange();
  let objs = convertArrayToObjects(dataRange.getValues());

  return objs;
}


/**
 * Converts a 2D array of data values into an array of objects.
 * Each object represents a row in the data, with keys derived from the
 * first row (header row) and values from subsequent rows.
 *
 * @param {Array<Array<any>>} data - A 2D array of data values, typically obtained from a Google Sheets data range.
 * @return {Array<Object>} An array of objects, where each object represents a row in the data.
 *                          Keys are derived from the header row, and values are from corresponding rows.
 *
 * @example
 * var dataValues = [
 *   ["Name", "Age", "City"],
 *   ["John", 25, "New York"],
 *   ["Alice", 30, "London"],
 *   ["Bob", 28, "Tokyo"]
 * ];
 *
 * var result = convertDataToObjects(dataValues);
 * Logger.log(result);
 *
 * // Output:
 * // [
 * //   { "Name": "John", "Age": 25, "City": "New York" },
 * //   { "Name": "Alice", "Age": 30, "City": "London" },
 * //   { "Name": "Bob", "Age": 28, "City": "Tokyo" }
 * // ]
 */
function convertArrayToObjects(data) {
  // Get the headers (assuming they are in the first row)
  let headers = data[0];

  // Initialize an array to store objects
  let dataArray = [];

  // Iterate through rows starting from the second row (index 1)
  for (let i = 1; i < data.length; i++) {
    let rowData = data[i];
    let rowObject = {};

    // Iterate through columns and create key-value pairs
    for (let j = 0; j < headers.length; j++) {
      rowObject[headers[j]] = rowData[j];
    }

    // Push the object to the array
    dataArray.push(rowObject);
  }

  // Return the array of objects
  return dataArray;
}


/**
 * send email based on threads
 * 
 * @param {Array<EmailThread>} threads - a list of threads to be written in the email
 */
function sendUnansweredThreadEmail(threads) {

  let threadString = "<table><tr><th>Date</th><th>Subject</th><th>Last Sender</th></tr>";

  threads.forEach(thread => {
    const formattedDate = Utilities.formatDate(new Date(thread.last_date), "GMT", "yyyy/MM/dd (EEE)");

    threadString += `<tr><td>${formattedDate}</td><td>${thread.subject}</td><td>${thread.last_message.sender}</td></tr>`
  })
  threadString += "</table>";

  GmailApp.sendEmail(observedMail(),
    "[SCRIPT]Daily reminder unanswered emails", 
    ``,
    {"htmlBody":  `Hello Human,<br>These are the messages that still have pending answers from ${observedMail()}:<br><br>${threadString}<br>${footnote()}<br>${CLOSURE}<br><br>${scriptDisclaimer()}`,
     "from": observedMail(),
     "name": "Unanswered mail script"});

}