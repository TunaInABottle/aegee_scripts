// source: https://stackoverflow.com/questions/52981008/how-to-create-a-dropdown-prompt-in-google-app-script



function fncOpenMyDialog() {
  //Open a dialog
  let htmlDlg = HtmlService.createHtmlOutputFromFile('dropdown_html')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(200)
      .setHeight(150);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlDlg, 'A Title Goes Here');
};


function test_dropdown() {
  Logger.log( fromMailDropdown() )
  Logger.log( getFromMails().join("\n") )
}


function getFromMails() {
  return [Session.getEffectiveUser().getEmail(), ...GmailApp.getAliases()];
}

function fromMailDropdown() {
  fromMail = getFromMails();
  fromMailString = [];
  fromMail.forEach((el) => fromMailString.push(`<option value="${(el)}">${(el)}</option>\n`));
  return fromMailString.join("")
}


////////////