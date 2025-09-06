// To learn how to use this script, refer to the documentation:
// https://developers.google.com/apps-script/samples/automations/mail-merge

/*
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

// This version is authored in https://github.com/BigLep/gmail-mail-merge
 
/**
 * @OnlyCurrentDoc
*/
 
/**
 * Change these to match the column names you are using for email 
 * recipient addresses and email sent column.
*/
const RECIPIENT_COL  = "Recipient";
const EMAIL_SENT_COL = "Email Sent";
 
/** 
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
      .addItem('Send Emails (MailApp - best emoji support)', 'sendEmails')
      .addItem('Send Emails (GmailApp - worse emoji support)', 'sendEmailsWithGmail')
      .addItem('Draft Emails (Gmail App - worse emoji support)', 'draftEmails')
      .addToUi();
}
 
/**
 * Creates draft emails from sheet data instead of sending them.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
*/
function draftEmails(subjectLine, sheet=SpreadsheetApp.getActiveSheet()) {
  processEmails(subjectLine, sheet, true); // true = create drafts
}

/**
 * Sends emails from sheet data using MailApp (best emoji support).
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
*/
function sendEmails(subjectLine, sheet=SpreadsheetApp.getActiveSheet()) {
  processEmails(subjectLine, sheet, 'mailapp'); // 'mailapp' = send with MailApp
}

/**
 * Sends emails from sheet data using GmailApp (full Gmail features).
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
*/
function sendEmailsWithGmail(subjectLine, sheet=SpreadsheetApp.getActiveSheet()) {
  processEmails(subjectLine, sheet, 'gmailapp'); // 'gmailapp' = send with GmailApp
}

/**
 * Process emails from sheet data - create drafts, send with MailApp, or send with GmailApp.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
 * @param {string|boolean} mode 'mailapp', 'gmailapp', or true for drafts (backward compatibility)
*/
function processEmails(subjectLine, sheet=SpreadsheetApp.getActiveSheet(), mode='mailapp') {
  // Handle backward compatibility
  const createDrafts = mode === true;
  const useGmailApp = mode === 'gmailapp';
  const useMailApp = mode === 'mailapp' || mode === false;
  // option to skip browser prompt if you want to use this code in other projects
  if (!subjectLine){
    subjectLine = Browser.inputBox("Mail Merge", 
                                      "Type or copy/paste the subject line of the Gmail " +
                                      "draft message you would like to mail merge with:",
                                      Browser.Buttons.OK_CANCEL);
                                      
    if (subjectLine === "cancel" || subjectLine == ""){ 
    // If no subject line, finishes up
    return;
    }
  }
  
  // Gets the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  
  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  // Fetches displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const data = dataRange.getDisplayValues();

  // Assumes row 1 contains our column headings
  const heads = data.shift(); 
  
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
    // Only sends emails if email_sent cell is blank and not hidden by a filter
    if (row[EMAIL_SENT_COL] == ''){
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        // Parse recipients - support comma-separated email addresses
        const recipients = row[RECIPIENT_COL].split(',').map(email => email.trim()).filter(email => email);
        const primaryRecipient = recipients[0];
        const additionalRecipients = recipients.slice(1);

        if (createDrafts) {
          // Create draft using GmailApp with smart emoji conversion
          // Automatically converts problematic emojis to HTML entities to prevent corruption
          const draftMsg = prepareMessageForDraft_(msgObj);
          
          GmailApp.createDraft(primaryRecipient, draftMsg.subject, draftMsg.text, {
            htmlBody: draftMsg.html,
            cc: additionalRecipients.length > 0 ? additionalRecipients.join(',') : undefined,
            // bcc: 'a.bcc@email.com',
            // from: 'an.alias@email.com', 
            // name: 'name of the sender',
            // replyTo: 'a.reply@email.com',
            attachments: emailTemplate.attachments,
            inlineImages: emailTemplate.inlineImages
          });
          // Record draft created date
          out.push([`Draft created ${new Date()}`]);
        } else if (useGmailApp) {
          // Send email using GmailApp with smart emoji conversion for full Gmail features
          const sendMsg = prepareMessageForDraft_(msgObj); // Apply emoji conversion
          
          GmailApp.sendEmail(primaryRecipient, sendMsg.subject, sendMsg.text, {
            htmlBody: sendMsg.html,
            cc: additionalRecipients.length > 0 ? additionalRecipients.join(',') : undefined,
            // bcc: 'a.bcc@email.com',
            // from: 'an.alias@email.com', 
            // name: 'name of the sender',
            // replyTo: 'a.reply@email.com',
            // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
            attachments: emailTemplate.attachments,
            inlineImages: emailTemplate.inlineImages
          });
          // Record email sent date
          out.push([new Date()]);
        } else {
          // Send email using MailApp for reliable Unicode/emoji support (default)
          // See https://developers.google.com/apps-script/reference/mail/mail-app#sendEmail(String,String,String,Object)
          // MailApp handles emojis correctly where GmailApp can corrupt them during processing
          MailApp.sendEmail(primaryRecipient, msgObj.subject, msgObj.text, {
            htmlBody: msgObj.html,
            cc: additionalRecipients.length > 0 ? additionalRecipients.join(',') : undefined,
            // bcc: 'a.bcc@email.com',
            // name: 'name of the sender',
            // replyTo: 'a.reply@email.com',
            attachments: emailTemplate.attachments
            // Note: MailApp doesn't support inlineImages - they become regular attachments
          });
          // Record email sent date
          out.push([new Date()]);
        }
      } catch(e) {
        // modify cell to record error
        out.push([e.message]);
      }
    } else {
      out.push([row[EMAIL_SENT_COL]]);
    }
  });
  
  // Updates the sheet with new data
  sheet.getRange(2, emailSentColIdx+1, out.length).setValues(out);
  
  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */
  function getGmailTemplateFromDrafts_(subject_line){
    try {
      // get drafts
      const drafts = GmailApp.getDrafts();
      // filter the drafts that match subject line
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
  
  /**
   * Convert emojis to HTML entities to prevent corruption in Gmail drafts
   * @param {string} text containing emojis
   * @return {string} text with emojis converted to HTML entities
  */
  function convertEmojisToHtmlEntities_(text) {
    if (!text) return text;
    
    // Common emoji mappings to HTML entities
    const emojiMap = {
      // Your specific emojis
      '🥏': '&#127949;',        // frisbee
      '🐶': '&#128054;',        // dog
      '🤔': '&#129300;',        // thinking face
      '🙋': '&#128587;',        // raising hand
      
      // Common emojis that might get corrupted
      '😀': '&#128512;',        // grinning face
      '😁': '&#128513;',        // beaming face
      '😂': '&#128514;',        // face with tears of joy
      '😊': '&#128522;',        // smiling face with smiling eyes
      '👍': '&#128077;',        // thumbs up
      '👎': '&#128078;',        // thumbs down
      '❤️': '&#10764;&#65039;', // red heart with variant selector
      '🎉': '&#127881;',        // party popper
      '🔥': '&#128293;',        // fire
      '💯': '&#128175;',        // hundred points
      '🚀': '&#128640;',        // rocket
      '⭐': '&#11088;',         // star
      '✨': '&#10024;',         // sparkles
      
      // Keep simple emojis that work well (these are commented out - no conversion needed)
      // '☑️': '☑️',             // checkbox (works fine)
      // '✉️': '✉️',             // envelope (works fine) 
      // '❓': '❓',             // question mark (works fine)
    };
    
    let result = text;
    for (const [emoji, entity] of Object.entries(emojiMap)) {
      const regex = new RegExp(emoji, 'g');
      result = result.replace(regex, entity);
    }
    
    return result;
  }

  /**
   * Smart emoji handling for draft creation - converts problematic emojis to HTML entities
   * @param {object} msgObj message object with subject, text, html properties
   * @return {object} message object with emojis converted for draft compatibility
  */
  function prepareMessageForDraft_(msgObj) {
    return {
      subject: convertEmojisToHtmlEntities_(msgObj.subject),
      text: convertEmojisToHtmlEntities_(msgObj.text), 
      html: convertEmojisToHtmlEntities_(msgObj.html)
    };
  }

  /**
   * Fill template string with data object
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
}