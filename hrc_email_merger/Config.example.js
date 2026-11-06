/**
 * Example configuration file.
 *
 * Fill in your own values in a script named `Config` (Apps Script UI) or in a
 * local Config.js (clasp), and keep that out of version control (it is listed
 * in .gitignore).
 *
 * NOTE: this example file itself is excluded from clasp sync via .claspignore
 * and must NOT be added to the Apps Script project: it would declare CONFIG a
 * second time ("SyntaxError: Identifier 'CONFIG' has already been declared").
 */
const CONFIG = {
  // ID of the Google Drive folder holding the email templates. The title of
  // each Google Doc in the folder must exactly match the "Email subject"
  // value of the rows it should be used for (before the [square brackets]
  // are stripped for sending). Leave empty to use Gmail drafts instead.
  templateFolderId: "YOUR_TEMPLATE_FOLDER_ID",
};