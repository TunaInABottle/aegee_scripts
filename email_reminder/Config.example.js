/**
 * Example configuration file.
 *
 * Copy this file to Config.js, fill in your own values, and keep Config.js
 * out of version control (it is listed in .gitignore). clasp still pushes
 * Config.js to your Apps Script project so the code can use it at runtime.
 */
const CONFIG = {
  // Inbox observed by the script; the daily digest of unanswered mail is sent here
  observedMail: "you@example.org",

  // Spreadsheet linked in the reminder footer, where receivers can mark
  // threads that need no reply
  footnoteSpreadsheetUrl: "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit?usp=sharing",

  // Contact note appended to every automatic email
  contactNote: "your.email{[at]}example.org",
};