/**
 * HRC Email Merger — a Gmail mail merge driven by a Google Sheet.
 *
 * Adapted from Google's mail-merge sample by Martin Hawksey:
 *   https://developers.google.com/apps-script/samples/automations/mail-merge
 *   Copyright 2022 Martin Hawksey
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * Modified by TunaInABottle.
 *
 * References:
 *   - modal dialog with a dropdown fed from an array (used here)
 *     https://stackoverflow.com/questions/75636112/a-modal-dialog-box-with-some-text-a-drop-down-that-takes-values-from-an-array-a
 *   - dropdown-in-prompt/sidebar approaches (not used, kept for reference)
 *     https://stackoverflow.com/questions/52981008/how-to-create-a-dropdown-prompt-in-google-app-script
 *     https://stackoverflow.com/questions/32902865/how-to-create-a-drop-down-list-in-input-box
 */

/**
 * Change these to match the column names you are using for the email
 * recipient address, subject and email-sent columns.
 */
const RECIPIENT_COL  = "Recipient";
const SUBJECT_COL    = "Email subject";
const CC_COL         = "cc";
const BCC_COL        = "bcc";
const EMAIL_SENT_COL = "Email Sent";

/**
 * ID of the Google Drive folder that holds the email templates. The title
 * of each Google Doc in the folder must exactly match the subject of the
 * email (the "Email subject" column value, i.e. the subject before the
 * [square brackets] are stripped for sending).
 *
 * Configured in a `Config` script created from Config.example.js (the
 * example and Config.js are git- and clasp-ignored, so your folder ID never
 * lands in this repo). When unset/empty, the script falls back to looking
 * up a Gmail draft by subject instead.
 */
const TEMPLATE_FOLDER_ID = (typeof CONFIG !== "undefined" && CONFIG.templateFolderId)
    ? CONFIG.templateFolderId
    : "";

/**
 * Creates the "Mail Merge" menu for the user to run the scripts from.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Send Email samples', 'sendSampleEmailsFromDialog')
      .addItem('Send Emails', 'sendEmailsFromDialog')
      .addToUi();
}

/** Menu entry: show the alias dropdown then send emails. */
function sendEmailsFromDialog() {
  showEmailAliasDialog('send');
}

/** Menu entry: show the alias dropdown then send sample emails. */
function sendSampleEmailsFromDialog() {
  showEmailAliasDialog('sample');
}

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

/**
 * Returns info about pending (unsent) emails in the active sheet.
 * @return {{total: number, indices: number[], recipients: string[]}}
 */
function getPendingEmailInfo() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getDisplayValues();
  const heads = data.shift();
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
    const emailTemplate = getEmailTemplate_(obj[SUBJECT_COL]);
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
 * Returns the email template for a given subject line.
 *
 * When TEMPLATE_FOLDER_ID is set (via Config.js), the template is looked up
 * in that Drive folder (the document whose title matches the subject);
 * otherwise the Gmail draft with the matching subject is used (legacy
 * behaviour).
 *
 * @param {string} subject_line - raw subject from the sheet (before the
 *     [square brackets] are stripped for sending)
 * @return {object} containing the subject, plain and html message body and
 *     attachments and inline images
 */
function getEmailTemplate_(subject_line) {
  if (TEMPLATE_FOLDER_ID) {
    return getDriveTemplateFromFolder_(subject_line);
  }
  return getGmailTemplateFromDrafts_(subject_line);
}

/**
 * Get an email template from a Drive folder of Google Docs: the document
 * whose title equals the subject line provides the email structure.
 *
 * Google Docs are exported to HTML with their inline images embedded as
 * cid: references; text/html and text/plain files are used as-is.
 *
 * @param {string} subject_line to search for in the template folder
 * @return {object} containing the subject, plain and html message body,
 *     attachments and inline images
 */
function getDriveTemplateFromFolder_(subject_line) {
  let file;
  try {
    const folder = DriveApp.getFolderById(TEMPLATE_FOLDER_ID);
    const wanted = String(subject_line).trim();
    const files = folder.getFiles();
    while (files.hasNext()) {
      const candidate = files.next();
      if (String(candidate.getName()).trim() === wanted) {
        file = candidate;
        break;
      }
    }
  } catch (e) {
    throw new Error("Oops - can't read template folder (" + TEMPLATE_FOLDER_ID + "): " + e.message);
  }

  if (!file) {
    throw new Error("Oops - can't find a document titled \"" + subject_line + '" in the template folder');
  }

  const mimeType = file.getMimeType();
  let html;
  let text;
  const inlineImages = {};

  if (mimeType === MimeType.GOOGLE_DOCS) {
    const exported = getGoogleDocHtml_(file.getId());
    html = exported.html;
    text = htmlToPlainText_(html);
    Object.keys(exported.inlineImages).forEach(cid => inlineImages[cid] = exported.inlineImages[cid]);
  } else if (mimeType === MimeType.HTML) {
    html = file.getBlob().getDataAsString();
    text = htmlToPlainText_(html);
  } else if (mimeType === MimeType.PLAIN_TEXT) {
    text = file.getBlob().getDataAsString();
    html = '<div>' + escapeHtml_(text).replace(/\n/g, '<br>') + '</div>';
  } else {
    throw new Error('Oops - template "' + file.getName() + '" is a ' + mimeType +
        ' file; only Google Docs, HTML and plain text templates are supported');
  }

  return {
    message: {subject: subject_line, text: text, html: html},
    attachments: [],
    inlineImages: inlineImages
  };
}

/**
 * Export a Google Doc to HTML and turn its hosted inline images into
 * cid: references, so they render inside the email instead of pointing
 * at Google-hosted URLs.
 *
 * Uses the legacy Docs HTML-export endpoints, authenticated with the
 * script's own OAuth token (the Drive API no longer exports Docs to HTML).
 *
 * @param {string} fileId - ID of the Google Doc to export
 * @return {{html: string, inlineImages: object}} body HTML with cid:
 *     image references and a map of cid -> image Blob
 */
function getGoogleDocHtml_(fileId) {
  const exportUrls = [
    'https://docs.google.com/feeds/download/documents/export/Export' +
        '?id=' + encodeURIComponent(fileId) + '&exportFormat=html',
    'https://docs.google.com/document/d/' + encodeURIComponent(fileId) + '/export?format=html'
  ];

  let page = null;
  for (let i = 0; i < exportUrls.length && page === null; i++) {
    const res = UrlFetchApp.fetch(exportUrls[i], {
      headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200) {
      page = res.getContentText();
    }
  }
  if (page === null) {
    throw new Error('Oops - could not export document ' + fileId + ' to HTML');
  }

  // Keep only the body content of the exported page.
  const bodyMatch = page.match(/<body[^>]*>([\s\S]*)<\/body>/i);
  let html = bodyMatch ? bodyMatch[1] : page;

  // Replace the hosted image URLs with cid: references and collect the
  // corresponding blobs for GmailApp's inlineImages option.
  const inlineImages = {};
  let imageIndex = 0;
  html = html.replace(/<img[^>]*\ssrc="(https:\/\/[^"]+)"[^>]*>/gi, function(tag, src) {
    try {
      const cid = 'template_image_' + (++imageIndex);
      const blob = UrlFetchApp.fetch(src, {
        headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}
      }).getBlob();
      if (!/^image\//.test(blob.getContentType())) {
        blob.setContentType('image/png'); // hosted images often report no useful content type
      }
      blob.setName(cid);
      inlineImages[cid] = blob;
      return tag.replace(src, 'cid:' + cid);
    } catch (err) {
      return tag; // leave the hosted URL in place if the image cannot be fetched
    }
  });

  return {html: html, inlineImages: inlineImages};
}

/**
 * Minimal HTML-to-text conversion, used for the plain-text part of the
 * email when the template comes from a document.
 * @param {string} html to convert
 * @return {string} plain text
 */
function htmlToPlainText_(html) {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr|h[1-6]|blockquote)>/gi, '\n\n')
    .replace(/<li[^>]*>/gi, '  \u2022 ')
    .replace(/<\/li>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Escape HTML special characters.
 * @param {string} str to escape
 * @return {string} escaped string
 */
function escapeHtml_(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
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

/**
 * Fill template string with data object
 * @see https://stackoverflow.com/a/378000/1027723
 * @param {string} template string containing {{}} markers which are replaced with data
 * @param {object} data object used to replace {{}} markers
 * @return {object} message replaced with data
 */
function fillInTemplateFromObject_(template, data) {
  // We have two templates one for plain text and the html body
  // Stringifing the object means we can do a global replace
  let template_string = JSON.stringify(template);

  // Token replacement
  template_string = template_string.replace(/{{[^{}]+}}/g, key => {
    return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
  });
  return JSON.parse(template_string);
}

/**
 * Escape cell data to make JSON safe
 * @see https://stackoverflow.com/a/9204218/1027723
 * @param {string} str to escape JSON special characters from
 * @return {string} escaped string
 */
function escapeData_(str) {
  return str
    .replace(/[\\]/g, '\\\\')
    .replace(/[\\\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
}
