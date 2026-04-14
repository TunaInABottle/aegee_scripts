// To learn how to use this script, refer to the documentation:
// https://developers.google.com/apps-script/samples/automations/mail-merge

/*
This file is an adaptation of the file found in the above link made by
TunaInABottle

Copyright 2022 Martin Hawksey

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

/**
 * @OnlyCurrentDoc
*/

/**
 * Change these to match the column names you are using for email 
 * recipient addresses and email sent column.
*/
const RECIPIENT_COL  = "Recipient";
const SUBJECT_COL    = "Email subject";
const CC_COL  = "cc";
const BCC_COL  = "bcc";
const EMAIL_SENT_COL = "Email Sent";

/** 
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Send Email samples', 'sendSampleEmailsFromDialog')
      .addItem('Send Emails', 'sendEmailsFromDialog')
      .addToUi();
}

/**
 * Sends emails from sheet data.
 * @param {string} emailAlias (optional) TODO
 * @param {Sheet} sheet to read data from
*/
function sendEmails(emailAlias, sheet=SpreadsheetApp.getActiveSheet()) {

  // Ask from which alias send the communication 
  if (!emailAlias){
    emailAlias = Browser.inputBox("Mail Merge", 
                                      "From which address you want to send the e-mails from?" +
                                      "These are the emails you can use:  " +
                                      getFromMails().join("  OR  "),
                                      Browser.Buttons.OK_CANCEL);

    if (emailAlias === "cancel" || emailAlias == ""){ 
    // If no subject line, finishes up
    return;
    }
  }
  emailAlias = emailAlias.trim();

  // Gets the draft Gmail message to use as a template
  

  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetches displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  const heads = data.shift();  // Assumes row 1 contains our column headings

  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Creates an array to record sent emails
  const out = [];

  // Loops through all the rows of data
  obj.forEach(function(row, rowIdx){
    // Only sends emails if email_sent cell is blank and not hidden by a filter,
    // the subject and the recipient are not empty
    if (!row[EMAIL_SENT_COL].startsWith("Sent on ") && !row[SUBJECT_COL] == "" && !row[RECIPIENT_COL] == ""){
      try {
        const emailTemplate = getGmailTemplateFromDrafts_(row[SUBJECT_COL]);
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
        const finalSubject = msgObj.subject.replace(/\[[^\]]*\]/, "");

        // See https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
        // If you need to send emails with unicode/emoji characters change GmailApp for MailApp
        GmailApp.sendEmail(row[RECIPIENT_COL], finalSubject, msgObj.text, {
          htmlBody: msgObj.html,
          bcc: row[BCC_COL],
          cc: row[CC_COL],
          from: emailAlias,
          // name: 'name of the sender',
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        });
        // Edits cell to record email sent date
        const date = new Date()
        const day = date.getDate();
        const month = date.toLocaleString('default', { month: 'long' });
        out.push([`Sent on ${day} ${month}`]);
      } catch(e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });

  // writes the result of the operation in the spreadsheet
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);

  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */
  function getGmailTemplateFromDrafts_(subject_line){
    try {
      const drafts = GmailApp.getDrafts();
      // take first draft that matches the subject
      const draft = drafts.filter(subjectFilter_(subject_line))[0];
      // get the message object
      const msg = draft.getMessage();

      // Handles inline images and attachments so they can be included in the merge
      // Based on https://stackoverflow.com/a/65813881/1027723
      // Gets all attachments and inline image attachments
      const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
      const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
      const htmlBody = msg.getBody(); 

      // Creates an inline image object with the image name as key 
      // (can't rely on image index as array based on insert order)
      const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

      //Regexp searches for all img string positions with cid
      const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
      const matches = [...htmlBody.matchAll(imgexp)];

      //Initiates the allInlineImages object
      const inlineImagesObj = {};
      // built an inlineImagesObj from inline image matches
      matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

      return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody}, 
              attachments: attachments, inlineImages: inlineImagesObj };
    } catch(e) {
      throw new Error("Oops - can't find Gmail draft");
    }

    /**
     * Filter draft objects with the matching subject linemessage by matching the subject line.
     * @param {string} subject_line to search for draft message
     * @return {object} GmailDraft object
    */
    function subjectFilter_(subject_line){
      return function(element) {
        if (element.getMessage().getSubject() === subject_line) {
          return element;
        }
      }
    }
  }
}

/**
 * TODO
 * function to test if the emails have been properly formatted
 * It ignores fields such as CC and BCC and sends the email only
 * To the specified alias
 * One email per subject
 */
function sendSampleEmails(emailAlias, sheet=SpreadsheetApp.getActiveSheet()) {

  // Ask from which alias send the communication 
  if (!emailAlias){
    emailAlias = Browser.inputBox("Check Mail Sample", 
                                      "This script helps to debug if the email have properly been formatted. All emails are sent only to the specified alias --- From which address you want to send the e-mails from?" +
                                      "These are the emails you can use:  " +
                                      getFromMails().join("  OR  "),
                                      Browser.Buttons.OK_CANCEL);

    if (emailAlias === "cancel" || emailAlias == ""){ 
    // If no subject line, finishes up
    return;
    }
  }
  emailAlias = emailAlias.trim();

  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetches displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  const heads = data.shift();  // Assumes row 1 contains our column headings

  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Creates an array to record sent emails
  const out = [];
  const seenSubjects = []; // prevent to send multiple emails with same subject

  // Loops through all the rows of data
  obj.forEach(function(row, rowIdx){
    Logger.log(`email iteration ${rowIdx}`);
    // Only sends emails if email_sent cell is blank and not hidden by a filter,
    // the subject and the recipient are not empty
    if (!row[EMAIL_SENT_COL].startsWith("Sent on ") && !row[SUBJECT_COL] == "" && !row[RECIPIENT_COL] == ""){
      Logger.log("preparing to send email...")
      try {
        const emailTemplate = getGmailTemplateFromDrafts_(row[SUBJECT_COL]);
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
        const rawSubject   = msgObj.subject;
        const finalSubject = msgObj.subject.replace(/\[[^\]]*\]/, "") + " (SAMPLE)";
        let skipSending = false

        // ensure that the iteration only happens throughout unseen subjects
        if (!seenSubjects.includes(rawSubject)) {
            Logger.log("Unseen subject")
            seenSubjects.push(rawSubject);
        } else {
            skipSending = true;
        }

        if (!skipSending) {
          Logger.log("sending email...")
          GmailApp.sendEmail(emailAlias, finalSubject, msgObj.text, {
            htmlBody: msgObj.html,
            from: emailAlias,
            attachments: emailTemplate.attachments,
            inlineImages: emailTemplate.inlineImages
          });
        }
        out.push([""]);
      } catch(e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });

  // writes the result of the operation in the spreadsheet
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);

  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */
  function getGmailTemplateFromDrafts_(subject_line){
    try {
      const drafts = GmailApp.getDrafts();
      // take first draft that matches the subject
      const draft = drafts.filter(subjectFilter_(subject_line))[0];
      // get the message object
      const msg = draft.getMessage();

      // Handles inline images and attachments so they can be included in the merge
      // Based on https://stackoverflow.com/a/65813881/1027723
      // Gets all attachments and inline image attachments
      const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
      const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
      const htmlBody = msg.getBody(); 

      // Creates an inline image object with the image name as key 
      // (can't rely on image index as array based on insert order)
      const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

      //Regexp searches for all img string positions with cid
      const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
      const matches = [...htmlBody.matchAll(imgexp)];

      //Initiates the allInlineImages object
      const inlineImagesObj = {};
      // built an inlineImagesObj from inline image matches
      matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

      return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody}, 
              attachments: attachments, inlineImages: inlineImagesObj };
    } catch(e) {
      throw new Error("Oops - can't find Gmail draft");
    }

    /**
     * Filter draft objects with the matching subject linemessage by matching the subject line.
     * @param {string} subject_line to search for draft message
     * @return {object} GmailDraft object
    */
    function subjectFilter_(subject_line){
      return function(element) {
        if (element.getMessage().getSubject() === subject_line) {
          return element;
        }
      }
    }
  }
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
  return  JSON.parse(template_string);
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
    .replace(/[\"]/g, '\\\"')
    .replace(/[\/]/g, '\\/')
    .replace(/[\b]/g, '\\b')
    .replace(/[\f]/g, '\\f')
    .replace(/[\n]/g, '\\n')
    .replace(/[\r]/g, '\\r')
    .replace(/[\t]/g, '\\t');
};
