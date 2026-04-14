// source: https://stackoverflow.com/questions/52981008/how-to-create-a-dropdown-prompt-in-google-app-script

function getFromMails() {
  return [Session.getEffectiveUser().getEmail(), ...GmailApp.getAliases()];
}

/**
 * Shows a modal dialog with a dropdown of available email aliases.
 * @param {string} callbackFunction - the server-side function to call with the selected alias
 */
function showEmailAliasDialog(callbackFunction) {
  const template = HtmlService.createTemplateFromFile('EmailAliasDialog');
  template.emailOptions = getFromMails();
  template.callbackFunction = callbackFunction;
  const html = template.evaluate()
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'From which address you want to send the emails?');
}

/** Menu entry: show dropdown then send emails */
function sendEmailsFromDialog() {
  showEmailAliasDialog('sendEmails');
}

/** Menu entry: show dropdown then send sample emails */
function sendSampleEmailsFromDialog() {
  showEmailAliasDialog('sendSampleEmails');
}
