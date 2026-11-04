// Protype-based version of the code, currently unused
// @todo - test with multiple recipients

function prototype_testing() {
  let threads = GmailApp.search("is:starred", 0, 2);
  Logger.log(threads);

  custom_th = [];
  threads.forEach((thread) => {
    custom_th.push(EmailThread.fromGoogleAPI(thread))});

  Logger.log(custom_th);
}

////////////////////////////////
////////////////////////////////

/**
 * @TODO
 * used in get_last_unanswered_message
 */
class EmailNotInThreadError extends Error {
  constructor(message) {
    super(message);
    this.name = 'EmailNotInThreadError';
  }
}

////////////////////////////////
/**
 * Represents the Email user
 * 
 * @param {String} sender - email address of the user
 * 
 * @property {String} sender - email address of the user
 */
function EmailAccount(sender) {
  this.sender = sender;
}

/**
 * Sends a debug mail to `sender`
 * 
 * @param {String} emailBody - HTML content of the debug mail to be sent
 */
EmailAccount.prototype.sendDebugEmail = function (emailBody) {
    MailApp.sendEmail(this.sender, "[TEST] auto reply", "", {
    htmlBody: emailBody
  });
};

/**
 * @TODO change to get threads and return a list of emailthread
 * Make a query
 * 
 * @param {String} searchQuery - Gmail query
 * @param {numeric} max - how many result to output at best (default: 50)
 * @return {EmailThread[]} - list of threads matching this query
 */
EmailAccount.prototype.search = function(searchQuery, max = 50) {
  googleQueryResults = GmailApp.search(searchQuery, 0, max);

  threads = [];
  googleQueryResults.forEach(thread => {threads.push(EmailThread.fromGoogleAPI(thread))});

  return threads;
}

/////////////////
// EmailThread //
/////////////////

/**
 * Represents a thread of emails
 * 
 * @property {string} subject - subject of the thread
 * @property {EmailMessage[]} messages - emails contained in the thread sorted from the oldest
 * @property {EmailMessage} last_message - last email of the thread
 * @property {String?} last_recipient - last email address of whom received the mail
 * @property {Date} last_date - timestamp of the last message in the thread
 * @property {EmailLabel[]} labels - labels associated with this thread
 * @property {String?} id - ID of this thread
 */
function EmailThread(googleObj, id, subject, messages, labels) {
  this._googleObj = googleObj;
  this.id = id;
  this.subject = subject;
  this.messages = messages;
  this.last_message = this.messages[this.messages.length - 1];
  this.last_recipient = this.last_message.recipient;
  this.last_date = this.last_message.date;
  this.labels = labels;
}

/**
 * Converts a google thread object.
 * 
 * @param {gmailThread} gmailThread - a Google API representation
 * 
 * @return {EmailThread} - a non-Google thread 
 */
EmailThread.fromGoogleAPI = function (gmailThread) {
  let googleObj = gmailThread;
  let id = gmailThread.getId();
  let subject = gmailThread.getFirstMessageSubject()
  let messages = [];
  gmailThread.getMessages().forEach((message) => {messages.push(new EmailMessage(message))});

  let labels = [];
  gmailThread.getLabels().forEach((label) => {labels.push(new EmailLabel(label.getName()))});

  return new EmailThread(googleObj, id, subject, messages, labels);
}

/**
 * makes threads from Google ID
 * 
 * @param {String} threadId - a reference to a Google thread
 * 
 * @return {EmailThread} - a non-Google thread
 */
EmailThread.fromGoogleThreadID = function (threadId) {
  let googleObj = GmailApp.getThreadById(threadId);
  return new EmailThread.fromGoogleAPI(googleObj);
}

/**
 * Checks if the last recipients was not `MY_MAIL`
 * @TODO check aliases too
 * 
 * return {Boolean} - true if `MY_MAIL` was the recipient, false otherwise
 */
EmailThread.prototype.got_reply = function() {
  return this.last_recipient.indexOf(MY_EMAIL) == -1;
}

/**
 * Marks the conversation as replied through tags
 * 
 * @modify this._googleObj by removing `NEEDREPLY_LAB` children label
 * @modify this._googleObj by adding `GOTREPLY_LAB` children label
 */
EmailThread.prototype.remove_followup_label = function () {
  let need_reply_label = this.labels.filter(function(label) {
    return label.name.includes(NEEDREPLY_LAB.getName());
  })[0]
  // assume no multiples `NEEDREPLY_LAB` in same thread
  this._googleObj.removeLabel(need_reply_label._googleObj);
  this._googleObj.addLabel(GOTREPLY_LAB);
}

/**
 * @TODO
 */
EmailThread.prototype.setIgnoreLabel = function () {
  this._googleObj.addLabel(GmailApp.getUserLabelByName(SKIP_LAB));
}


/**
 * @TODO
 */
EmailThread.prototype.toSpreadsheet = function () {
  // sender email
  let match = this.last_message.sender.match(/<([^>]*)>/);
  let sender_mail = "";
  if (match) {
    sender_mail = match[1];
  } else {
    sender_mail = this.last_message.sender;
  }

  // sender name
  match = this.last_message.sender.match(/"([^"]*)"|([^<]+)/);
  let sender_name = "";
  if (match) {
    sender_name = match[1] || match[2];
  } else {
    sender_name = "unknown"
  }

  return [this.id, this.last_date, this.subject, sender_mail, sender_name]
}


/**
 * write object as string
 * 
 * @return {String} - thread formatted as string
 */
EmailThread.prototype.toString = function () {
  return `Thread: subject "${this.subject}"
  Last message sent to ${this.last_recipient} 
  (${this.messages.length} messages)\n`;
};

//////////////////
// EmailMessage //
//////////////////

/**
 * An email
 * @todo what happens with multiple recipients?
 * 
 * @property {String} sender - who sent the email
 * @property {String} recipient - who received the email
 * @property {String} cc - who is in CC
 * @property {String} bcc - who is in BCC
 * @property {String} subject - the subject of the email
 * @property {Date} date - when the email has been sent/received
 * @property {String} body - the content of the email
 * @property {String} opening - the first line of the body
 * 
 * assume - there is a row dedicated to the recipient (ex. Dear John,)
 */
function EmailMessage(gmailMessage) {
  this._googleObj = gmailMessage;
  this.sender = gmailMessage.getFrom();
  this.recipient = gmailMessage.getTo();
  this.cc = gmailMessage.getCc();
  this.bcc = gmailMessage.getBcc();
  this.subject = gmailMessage.getSubject();
  this.date = gmailMessage.getDate();
  this.body = gmailMessage.getBody();
  this.opening = this._get_opening();
};

/**
 * Checks if this has been sent by `sender_email`
 * 
 * @param {String} sender_email - the sender email to be checked
 * @returns {Boolean} - true if `sender_email` sent this email, false otherwise
 */
EmailMessage.prototype.is_sent_by = function(sender_email) {
  return ( this.sender.indexOf(sender_email) != -1);
}

/**
 * Sends a followup to this message
 * 
 * uses `NAME_SENDER` to decide the seen name of followup's sender
 * 
 * @param {String} upper_header - a line to prefix after the opening and before the old body 
 */
EmailMessage.prototype.send_followup = function(upper_header, sender_email) {
  let reply_body = `${this.opening}${upper_header}<br>${this.body}<hr>${scriptDisclaimer()}`;

  let email_details = {
        htmlBody: reply_body,
        cc: this.cc,
        bcc: this.bcc,
        name: NAME_SENDER
      }
  if (GmailApp.getAliases().includes(sender_email)) {
    email_details.from = sender_email;
  }

  try {
    this._googleObj.replyAll("", email_details)
    console.log(`Just sent a message to ${this.recipient}. Subject: ${this.subject}`);
  } catch(e) {
    Logger.log(e);
  }
}

/**
 * Take the first line from a mail
 * 
 * @todo verify long-term robustness with different emails
 * 
 * @return {String} first line of EmailMessage's body
 */
EmailMessage.prototype._get_opening = function() {
  return `${this.body.split("</div>")[0]}</div>`
}


////////////////
// EmailLabel //
////////////////

/**
 * Custom representation of a label
 * 
 * @param {String} labelName - a label object
 * 
 * @property {String} name - label name
 * @property {Array[EmailLabel]} children - the sublabels of this label
 */
function EmailLabel(labelName) {
  label = EmailLabel.setLabel(labelName);

  this._googleObj = label;
  this.name = label.getName();
  this.children = EmailLabel._children(label);

}

/**
 * safely set a label
 * 
 * @modify will create an email label `labelName` if does
 * not exists already
 * 
 * @param {String} labelName - the name of the label
 * @return {EmailLabel} - a label with name `labelName`
 */
EmailLabel.setLabel = function (labelName) {
  let label;
  let label_names = [];
  
  // convert GmailLabel array into a string array with label names
  GmailApp.getUserLabels().forEach((el) => { label_names.push(el.getName()) })

  if (label_names.includes(labelName)) {
    label = GmailApp.getUserLabelByName(labelName);
  } else {
    label = GmailApp.createLabel(labelName);
    Logger.log(`Label "${labelName}" has been created`)
  }
  return label
}

/**
 * Get the sub labels
 * 
 * @param {GmailLabel} parent - the parent of whom we want to search the children
 * @return {GmailLabel[]} - a list of `parent`'s children
 */
EmailLabel._children = function (parent) { 
  let name = parent.getName() + '/';
  let children = GmailApp.getUserLabels().filter(function(label) {
    return label.getName().slice(0, name.length) == name;
  });
  
  let ret_array = [];
  children.forEach((el) => { 
    let this_child = el.getName();
    ret_array.push( new EmailLabel(this_child)) 
  });

  return ret_array;
}

/**
 * write the children in a Gmail compatible query
 * 
 * @return {String} - a query-like formatting of the children of
 *  `this` 
 */
EmailLabel.prototype.childrenToQuery = function() {
  let query = "";
  let first = true;

  this.children.forEach((child) => {
    if (first) {
      query = "label:" + child.name.replace(/ |\//g, '-');
      first = false;
    } else {
      query = query + " OR label:" + child.name.replace(/ |\//g, '-') + " ";
    }
  })
  return query;
}

//////////////////////////////////////////

/**
 * @TODO
 * get all senders email of this user
 */

function getMySenderEmails() {
  return [MY_EMAIL].concat(GmailApp.getAliases());
}
