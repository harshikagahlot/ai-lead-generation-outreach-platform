/**
 * PartDExistingDraftTestFix.js
 * Safe one-draft diagnostic for Part D v6.
 *
 * IMPORTANT:
 * - Does NOT change lead generation, qualification, or email filtering.
 * - Does NOT create a replacement draft when an old Gmail ID is invalid.
 * - Skips Gmail IDs that are no longer drafts (for example, sent/deleted items)
 *   instead of aborting the entire test.
 */

function testUpdateOneExistingPartDV6DraftSafe() {
  const ss = SpreadsheetApp.getActive();
  const drafts = ss.getSheetByName(SHEET_DRAFTS);
  if (!drafts || drafts.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Outreach_Drafts is empty.');
    return;
  }

  const qMap = buildQualifiedLeadMapForExistingDrafts_();
  const lastRow = drafts.getLastRow();
  const width = Math.max(DRAFT_HEADERS.length, drafts.getLastColumn());
  const data = drafts.getRange(2, 1, lastRow - 1, width).getValues();

  const subjectCol = DRAFT_HEADERS.indexOf('Subject') + 1;
  const bodyCol = DRAFT_HEADERS.indexOf('Email Draft') + 1;
  const gmailIdCol = DRAFT_HEADERS.indexOf('Gmail Draft ID') + 1;
  const placeCol = DRAFT_HEADERS.indexOf('Place ID') + 1;
  const statusCol = DRAFT_HEADERS.indexOf('Status') + 1;

  let checked = 0;
  let invalidDraftIds = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const placeId = String(row[placeCol - 1] || '').trim();
    const gmailDraftId = String(row[gmailIdCol - 1] || '').trim();
    const lead = qMap[placeId];

    if (!placeId || !gmailDraftId || !lead) continue;
    if (!lead.email || String(lead.recommendedChannel || '').toUpperCase() !== 'EMAIL') continue;
    if (!isValidOutreachEmail(lead.email)) continue;

    checked++;

    let draft;
    try {
      draft = GmailApp.getDraft(gmailDraftId);
    } catch (e) {
      invalidDraftIds++;
      Logger.log('Skipping non-draft Gmail ID on row ' + (i + 2) + ': ' + gmailDraftId + ' — ' + e.message);
      continue;
    }

    try {
      const email = buildPartDEmail_(lead);
      draft.update(lead.email, email.subject, email.body, {
        htmlBody: email.htmlBody,
        inlineImages: { vashaLogo: getVashaLogoBlob_() },
        name: 'Harshika'
      });

      drafts.getRange(i + 2, subjectCol).setValue(email.subject);
      drafts.getRange(i + 2, bodyCol).setValue(email.body);
      drafts.getRange(i + 2, statusCol).setValue('Part D v6 — Gmail draft updated');

      SpreadsheetApp.getUi().alert(
        'TEST SUCCESSFUL\n\n' +
        'Updated 1 existing Gmail draft:\n' + lead.name + '\n\n' +
        'Invalid/non-draft IDs skipped: ' + invalidDraftIds + '\n' +
        'The existing draft was updated in place. Nothing was sent.'
      );
      return;
    } catch (e) {
      SpreadsheetApp.getUi().alert('Draft update failed for ' + lead.name + ':\n\n' + e.message);
      return;
    }
  }

  SpreadsheetApp.getUi().alert(
    'No usable existing Gmail draft was found.\n\n' +
    'Checked matching rows: ' + checked + '\n' +
    'Invalid/non-draft Gmail IDs skipped: ' + invalidDraftIds + '\n\n' +
    'No new draft was created and nothing was sent.'
  );
}
