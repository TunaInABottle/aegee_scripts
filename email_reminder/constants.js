const NOW = new Date();
const MY_EMAIL = Session.getActiveUser().getEmail();

// Name of the label to skip
const SKIP_LAB = "SCRIPT_IGNORE"

// for new script
const REPLY_LABEL_PARENT = "NeedReply";

const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET = SS.getActiveSheet();


const NAME_SENDER = "Follow-up script"; //name that will appear as sender when a mail is sent through this script
//footer of the mail, warns the receiver that the message has been automatically generated

const CLOSURE = "<p>Hoping that I have been useful once again<br>Automatically yours,<br><i>your favourite Unanswered email script.</i></p>"

/**
 * Personal values (observed inbox, footer spreadsheet URL, contact note)
 * live in the local Config.js file (see Config.example.js), which is kept
 * out of version control via .gitignore.
 *
 * They are read lazily through the helpers below because Apps Script does
 * not guarantee the evaluation order of project files, so CONFIG might not
 * be available while this file is loaded.
 */
function getConfig_() {
  if (typeof CONFIG === "undefined") {
    throw new Error(
      "CONFIG is not defined: copy Config.example.js to Config.js, " +
      "fill in your values and push it with clasp."
    );
  }
  return CONFIG;
}

// Which inbox is checked for unanswered emails; the digest is sent here
function observedMail() {
  return getConfig_().observedMail;
}

//footer of the reminder mail, links the spreadsheet where receivers can
//mark threads that need no reply
function footnote() {
  return `<p>If an email does not need any reply, go to <a href="${getConfig_().footnoteSpreadsheetUrl}">The following spreadsheet</a> and check the corrispective checkbox.<br>If no thread needs reply, the email will not be sent, this is better for the environment!</p>`;
}

//disclaimer appended to every automatic mail
function scriptDisclaimer() {
  return `<font size="1">This message has been made with an automated script made by ${getConfig_().contactNote}, for any problem contact the author.`;
}