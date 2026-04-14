// source: https://stackoverflow.com/questions/52981008/how-to-create-a-dropdown-prompt-in-google-app-script

/**
 * Returns the user's email and all aliases.
 * @return {string[]} array of email addresses
 */
function getFromMails() {
  return [Session.getEffectiveUser().getEmail(), ...GmailApp.getAliases()];
}

/**
 * Shows a modal dialog with a dropdown of available email aliases.
 * @param {string} mode - 'send' or 'sample'
 */
function showEmailAliasDialog(mode) {
  const template = HtmlService.createTemplateFromFile('EmailAliasDialog');
  template.emailOptions = getFromMails();
  template.mode = mode;
  const html = template.evaluate()
      .setWidth(400)
      .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Email Merger');
}

/** Menu entry: show dropdown then send emails */
function sendEmailsFromDialog() {
  showEmailAliasDialog('send');
}

/** Menu entry: show dropdown then send sample emails */
function sendSampleEmailsFromDialog() {
  showEmailAliasDialog('sample');
}

/**
 * Returns info about pending (unsent) emails in the active sheet.
 * @return {{total: number, indices: number[], recipients: string[]}}
 */
function getPendingEmailInfo() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getDisplayValues();
  const heads = data.shift();
  const recipientIdx = heads.indexOf(RECIPIENT_COL);
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));
  const indices = [];
  const recipients = [];
  obj.forEach(function(row, idx) {
    if (!row[EMAIL_SENT_COL].startsWith("Sent on ") && row[SUBJECT_COL] !== "" && row[RECIPIENT_COL] !== "") {
      indices.push(idx);
      recipients.push(row[RECIPIENT_COL]);
    }
  });
  return { total: indices.length, indices: indices, recipients: recipients };
}

/**
 * Sends a single email for the given row index.
 * @param {string} emailAlias - the alias to send from
 * @param {number} dataRowIndex - 0-based index into the data rows (after header)
 * @param {string} mode - 'send' or 'sample'
 * @param {string[]} seenSubjects - subjects already sent (for sample deduplication)
 * @return {{status: string, subject?: string, message?: string}}
 */
function sendSingleRow(emailAlias, dataRowIndex, mode, seenSubjects) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getDisplayValues();
  const heads = data.shift();
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  const row = data[dataRowIndex];
  const obj = heads.reduce((o, k, i) => (o[k] = row[i] || '', o), {});

  if (obj[EMAIL_SENT_COL].startsWith("Sent on ") || obj[SUBJECT_COL] === "" || obj[RECIPIENT_COL] === "") {
    return { status: 'skipped' };
  }

  try {
    const emailTemplate = getGmailTemplateFromDrafts_(obj[SUBJECT_COL]);
    const msgObj = fillInTemplateFromObject_(emailTemplate.message, obj);
    const rawSubject = msgObj.subject;

    if (mode === 'sample') {
      if (seenSubjects && seenSubjects.indexOf(rawSubject) !== -1) {
        return { status: 'skipped_duplicate' };
      }
      const finalSubject = rawSubject.replace(/\[[^\]]*\]/, "") + " (SAMPLE)";
      GmailApp.sendEmail(emailAlias, finalSubject, msgObj.text, {
        htmlBody: msgObj.html,
        from: emailAlias,
        attachments: emailTemplate.attachments,
        inlineImages: emailTemplate.inlineImages
      });
      return { status: 'sent', subject: rawSubject };
    } else {
      const finalSubject = rawSubject.replace(/\[[^\]]*\]/, "");
      GmailApp.sendEmail(obj[RECIPIENT_COL], finalSubject, msgObj.text, {
        htmlBody: msgObj.html,
        bcc: obj[BCC_COL],
        cc: obj[CC_COL],
        from: emailAlias,
        attachments: emailTemplate.attachments,
        inlineImages: emailTemplate.inlineImages
      });
      const date = new Date();
      const day = date.getDate();
      const month = date.toLocaleString('default', { month: 'long' });
      sheet.getRange(dataRowIndex + 2, emailSentColIdx + 1).setValue('Sent on ' + day + ' ' + month);
      return { status: 'sent' };
    }
  } catch(e) {
    if (mode !== 'sample') {
      sheet.getRange(dataRowIndex + 2, emailSentColIdx + 1).setValue(e.message);
    }
    return { status: 'error', message: e.message };
  }
}

/**
 * Get a Gmail draft message by matching the subject line.
 * @param {string} subject_line to search for draft message
 * @return {object} containing the subject, plain and html message body and attachments
 */
function getGmailTemplateFromDrafts_(subject_line) {
  try {
    const drafts = GmailApp.getDrafts();
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    const msg = draft.getMessage();
    const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true, includeAttachments: false});
    const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
    const htmlBody = msg.getBody();
    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});
    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];
    const inlineImagesObj = {};
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);
    return {
      message: {subject: subject_line, text: msg.getPlainBody(), html: htmlBody},
      attachments: attachments,
      inlineImages: inlineImagesObj
    };
  } catch(e) {
    throw new Error("Oops - can't find Gmail draft");
  }
}

/**
 * Filter draft objects with the matching subject line.
 * @param {string} subject_line to search for draft message
 * @return {function} filter function for GmailDraft objects
 */
function subjectFilter_(subject_line) {
  return function(element) {
    if (element.getMessage().getSubject() === subject_line) {
      return element;
    }
  };
}
