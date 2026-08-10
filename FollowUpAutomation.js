/**
 * FollowUpAutomation.js
 * -----------------------------------------------------------------------
 * Part C: follow-up automation.
 *
 * SAFETY RULE: this file NEVER sends email automatically.
 * It only creates Gmail drafts after the original email was manually sent
 * and the configured delay has elapsed with no recipient reply detected.
 *
 * The setup function creates a daily time-based trigger. The daily checker:
 *   1. Finds Outreach_Drafts rows that were pushed to Gmail.
 *   2. Detects when the original Gmail draft is no longer in Drafts.
 *   3. Finds the corresponding sent thread by recipient + original subject.
 *   4. Records Date Sent in Outreach_Drafts.
 *   5. After Follow-Up Delay (default 5 days), checks for a recipient reply.
 *   6. Creates ONE follow-up Gmail draft if there is no reply.
 *
 * New columns are added to Outreach_Drafts automatically:
 *   Date Sent, Follow-Up Count, Last Follow-Up Date
 *
 * The setting is added to Settings automatically if missing:
 *   Follow-Up Delay (days) = 5
 */

const FOLLOW_UP_DELAY_SETTING = 'Follow-Up Delay (days)';
const DEFAULT_FOLLOW_UP_DELAY_DAYS = 5;
const FOLLOW_UP_SENT_STATUS = 'Follow-up Draft — Ready for Review';
const FOLLOW_UP_TRIGGER_HANDLER = 'runDailyFollowUpCheck';
const FOLLOW_UP_DATE_SENT_HEADER = 'Date Sent';
const FOLLOW_UP_COUNT_HEADER = 'Follow-Up Count';
const FOLLOW_UP_LAST_DATE_HEADER = 'Last Follow-Up Date';

/**
 * One-time setup. Run this once after pulling/pushing Part C.
 * It adds the setting/columns and installs exactly one daily trigger.
 */
function setupFollowUpAutomation() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Open the lead-generation spreadsheet before running setupFollowUpAutomation().');

  ensureFollowUpSetting_(ss);
  ensureFollowUpColumns_(ss);

  const triggers = ScriptApp.getProjectTriggers();
  const existing = triggers.filter(t => t.getHandlerFunction() === FOLLOW_UP_TRIGGER_HANDLER);
  for (let i = 1; i < existing.length; i++) {
    ScriptApp.deleteTrigger(existing[i]);
  }
  if (existing.length === 0) {
    ScriptApp.newTrigger(FOLLOW_UP_TRIGGER_HANDLER)
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
  }

  SpreadsheetApp.getUi().alert(
    'Follow-up automation is ready.\n\n' +
    'Delay: ' + getFollowUpDelayDays_() + ' day(s)\n' +
    'Runs: once daily\n\n' +
    'IMPORTANT: follow-ups are created as Gmail DRAFTS only. Nothing is auto-sent.'
  );
}

/** Daily trigger entry point. Safe to run manually for testing too. */
function runDailyFollowUpCheck() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('No active spreadsheet found.');

  const sheet = ss.getSheetByName(SHEET_DRAFTS);
  if (!sheet || sheet.getLastRow() < 2) return;

  ensureFollowUpColumns_(ss);

  const headers = getFollowUpHeaders_(sheet);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  const emailCol = headers.indexOf('Email');
  const subjectCol = headers.indexOf('Subject');
  const statusCol = headers.indexOf('Status');
  const gmailIdCol = headers.indexOf('Gmail Draft ID');
  const dateSentCol = headers.indexOf(FOLLOW_UP_DATE_SENT_HEADER);
  const countCol = headers.indexOf(FOLLOW_UP_COUNT_HEADER);
  const lastFollowCol = headers.indexOf(FOLLOW_UP_LAST_DATE_HEADER);
  const placeIdCol = headers.indexOf('Place ID');
  const nameCol = headers.indexOf('Business Name');

  if ([emailCol, subjectCol, statusCol, gmailIdCol, dateSentCol, countCol, lastFollowCol].some(c => c < 0)) {
    throw new Error('Outreach_Drafts is missing one or more required follow-up columns.');
  }

  const delayDays = getFollowUpDelayDays_();
  const now = new Date();
  const dateSentValues = data.map(row => [row[dateSentCol] || '']);
  const countValues = data.map(row => [row[countCol] || 0]);
  const lastFollowValues = data.map(row => [row[lastFollowCol] || '']);
  const statusValues = data.map(row => [row[statusCol] || '']);

  let sentDetected = 0;
  let followUpsCreated = 0;
  let repliesDetected = 0;

  data.forEach((row, i) => {
    const recipient = String(row[emailCol] || '').trim();
    const subject = String(row[subjectCol] || '').trim();
    const status = String(row[statusCol] || '').trim();
    const gmailDraftId = String(row[gmailIdCol] || '').trim();
    const dateSent = row[dateSentCol];
    const followUpCount = Number(row[countCol] || 0);

    if (!recipient || !subject) return;

    // Only process original Gmail drafts that were actually pushed.
    if (!gmailDraftId) return;
    if (status === FOLLOW_UP_SENT_STATUS) return;

    // If the Gmail draft still exists, it has not been manually sent yet.
    if (gmailDraftStillExists_(gmailDraftId)) return;

    // The draft disappeared from Drafts. Find the corresponding sent thread.
    const threadInfo = findOriginalSentThread_(recipient, subject);
    if (!threadInfo) return;

    // Record the actual sent date once detected.
    if (!dateSent) {
      dateSentValues[i] = [threadInfo.sentDate];
      sentDetected++;
    }

    const effectiveSentDate = dateSent ? new Date(dateSent) : threadInfo.sentDate;
    if (isNaN(effectiveSentDate.getTime())) return;

    // Do not follow up before the configured delay.
    const dueAt = new Date(effectiveSentDate.getTime() + delayDays * 24 * 60 * 60 * 1000);
    if (now < dueAt) return;

    // Check the thread for a genuine recipient reply.
    if (threadInfo.hasRecipientReply) {
      repliesDetected++;
      statusValues[i] = ['Replied'];
      return;
    }

    // Create at most one follow-up for now. A later follow-up cycle can be
    // added deliberately; this version does not create repeated nudges.
    if (followUpCount > 0 || row[lastFollowCol]) return;

    try {
      const followSubject = /^re:/i.test(subject) ? subject : 'Re: ' + subject;
      const followBody = buildFollowUpBody_(row[nameCol] || 'there');
      const draft = GmailApp.createDraft(recipient, followSubject, followBody);

      countValues[i] = [followUpCount + 1];
      lastFollowValues[i] = [now];
      statusValues[i] = [FOLLOW_UP_SENT_STATUS];
      followUpsCreated++;

      Logger.log('Created follow-up draft for ' + (row[nameCol] || recipient) + ': ' + draft.getId());
    } catch (err) {
      statusValues[i] = ['Follow-up Error: ' + err.message];
      Logger.log('Follow-up error for ' + recipient + ': ' + err.message);
    }
  });

  sheet.getRange(2, dateSentCol + 1, dateSentValues.length, 1).setValues(dateSentValues);
  sheet.getRange(2, countCol + 1, countValues.length, 1).setValues(countValues);
  sheet.getRange(2, lastFollowCol + 1, lastFollowValues.length, 1).setValues(lastFollowValues);
  sheet.getRange(2, statusCol + 1, statusValues.length, 1).setValues(statusValues);

  Logger.log(
    'Follow-up check complete. Sent detected: ' + sentDetected +
    ', replies detected: ' + repliesDetected +
    ', follow-up drafts created: ' + followUpsCreated
  );
}

/**
 * Returns true while a Gmail draft still exists. If the draft ID is no longer
 * valid, the original message may have been sent manually (or deleted).
 */
function gmailDraftStillExists_(draftId) {
  try {
    GmailApp.getDraft(draftId);
    return true;
  } catch (err) {
    return false;
  }
}

/**
 * Finds the original sent thread using recipient + subject. It deliberately
 * requires a sent message matching the original subject, reducing the chance
 * of treating an unrelated sent email as the original outreach.
 */
function findOriginalSentThread_(recipient, subject) {
  const safeEmail = recipient.replace(/[{}()\[\]"']/g, '');
  const safeSubject = subject.replace(/[{}()\[\]"']/g, '');
  const query = 'to:' + safeEmail + ' subject:"' + safeSubject + '" in:sent';

  let threads = [];
  try {
    threads = GmailApp.search(query, 0, 20);
  } catch (err) {
    Logger.log('Sent-thread search failed for ' + recipient + ': ' + err.message);
    return null;
  }

  if (!threads.length) return null;

  const userEmails = getKnownSenderEmails_();
  let best = null;

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      const from = extractEmailAddress_(message.getFrom());
      if (!userEmails.has(from.toLowerCase())) return;
      if (!message.getTo().toLowerCase().includes(recipient.toLowerCase())) return;
      if (message.getSubject().trim() !== subject.trim()) return;

      const sentDate = message.getDate();
      if (!best || sentDate < best.sentDate) {
        best = {
          thread: thread,
          sentDate: sentDate,
          hasRecipientReply: hasRecipientReply_(messages, userEmails, sentDate)
        };
      }
    });
  });

  return best;
}

function hasRecipientReply_(messages, userEmails, originalSentDate) {
  return messages.some(message => {
    const date = message.getDate();
    if (date <= originalSentDate) return false;
    const from = extractEmailAddress_(message.getFrom()).toLowerCase();
    return from && !userEmails.has(from);
  });
}

function getKnownSenderEmails_() {
  const set = new Set();
  const primary = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (primary) set.add(primary);
  try {
    GmailApp.getAliases().forEach(alias => set.add(String(alias).toLowerCase()));
  } catch (err) {
    Logger.log('Could not read Gmail aliases: ' + err.message);
  }
  return set;
}

function extractEmailAddress_(value) {
  const match = String(value || '').match(/<([^>]+)>/);
  return (match ? match[1] : String(value || '')).trim().toLowerCase();
}

function buildFollowUpBody_(businessName) {
  return 'Hi there,\n\n' +
    'Just floating this back up in case my earlier note about ' + (businessName || 'your business') + ' got buried — no pressure at all.\n\n' +
    'If it is worth a quick conversation, I\'d be happy to share a couple of ideas.\n\n' +
    'Best regards,\n\n' +
    'Harshika Gahlot\n' +
    'harshikagahlot01@gmail.com';
}

function ensureFollowUpSetting_(ss) {
  let sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) sheet = ss.insertSheet(SHEET_SETTINGS);

  const lastRow = sheet.getLastRow();
  const values = lastRow > 0 ? sheet.getRange(1, 1, lastRow, 2).getValues() : [];
  const found = values.some(row => String(row[0] || '').trim() === FOLLOW_UP_DELAY_SETTING);

  if (!found) {
    const row = Math.max(2, lastRow + 1);
    if (lastRow === 0) {
      sheet.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);
    }
    sheet.getRange(row, 1, 1, 2).setValues([[FOLLOW_UP_DELAY_SETTING, DEFAULT_FOLLOW_UP_DELAY_DAYS]]);
  }
}

function ensureFollowUpColumns_(ss) {
  const sheet = ss.getSheetByName(SHEET_DRAFTS);
  if (!sheet) return;

  const required = [FOLLOW_UP_DATE_SENT_HEADER, FOLLOW_UP_COUNT_HEADER, FOLLOW_UP_LAST_DATE_HEADER];
  const lastCol = sheet.getLastColumn();
  const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];

  required.forEach(header => {
    if (headers.indexOf(header) === -1) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(header).setFontWeight('bold');
    }
  });
}

function getFollowUpHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(v => String(v || '').trim());
}

function getFollowUpDelayDays_() {
  const settings = getSettings();
  const raw = Number(settings[FOLLOW_UP_DELAY_SETTING]);
  return isFinite(raw) && raw >= 1 ? raw : DEFAULT_FOLLOW_UP_DELAY_DAYS;
}
