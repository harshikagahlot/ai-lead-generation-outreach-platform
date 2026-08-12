/**
 * PartDExistingDraftUpdater.js
 * -----------------------------------------------------------------------
 * Safely upgrades EXISTING Outreach_Drafts rows to Part D v6.
 *
 * IMPORTANT:
 * - Does NOT touch lead generation, qualification, email discovery,
 *   placeholder/spam filtering, or Qualified_Leads.
 * - Does NOT delete drafts.
 * - Uses the existing Gmail Draft ID to UPDATE the existing Gmail draft
 *   in place when that draft still exists.
 * - Uses buildPartDEmail_() as the single email-copy source of truth.
 */

function buildQualifiedLeadMapForExistingDrafts_() {
  const ss = SpreadsheetApp.getActive();
  const qualified = ss.getSheetByName(SHEET_QUALIFIED);
  const map = {};
  if (!qualified || qualified.getLastRow() < 2) return map;

  const rows = qualified.getRange(2, 1, qualified.getLastRow() - 1, QUALIFIED_HEADERS.length).getValues();
  rows.forEach(row => {
    const lead = mapQualifiedRowToLead(row);
    if (lead.placeId) map[String(lead.placeId).trim()] = lead;
  });
  return map;
}

/**
 * Safe one-draft test.
 * Updates ONLY the first existing Outreach_Drafts row that has a matching
 * Qualified_Leads record and Gmail Draft ID.
 */
function testUpdateOneExistingPartDV6Draft() {
  const ss = SpreadsheetApp.getActive();
  const drafts = ss.getSheetByName(SHEET_DRAFTS);
  if (!drafts || drafts.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Outreach_Drafts is empty.');
    return;
  }

  const qMap = buildQualifiedLeadMapForExistingDrafts_();
  const lastRow = drafts.getLastRow();
  const data = drafts.getRange(2, 1, lastRow - 1, Math.max(DRAFT_HEADERS.length, drafts.getLastColumn())).getValues();

  const emailCol = DRAFT_HEADERS.indexOf('Email') + 1;
  const subjectCol = DRAFT_HEADERS.indexOf('Subject') + 1;
  const bodyCol = DRAFT_HEADERS.indexOf('Email Draft') + 1;
  const gmailIdCol = DRAFT_HEADERS.indexOf('Gmail Draft ID') + 1;
  const placeCol = DRAFT_HEADERS.indexOf('Place ID') + 1;
  const statusCol = DRAFT_HEADERS.indexOf('Status') + 1;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const placeId = String(row[placeCol - 1] || '').trim();
    const gmailDraftId = String(row[gmailIdCol - 1] || '').trim();
    const lead = qMap[placeId];

    if (!placeId || !gmailDraftId || !lead) continue;
    if (!lead.email || String(lead.recommendedChannel || '').toUpperCase() !== 'EMAIL') continue;
    if (!isValidOutreachEmail(lead.email)) continue;

    const email = buildPartDEmail_(lead);
    const draft = GmailApp.getDraft(gmailDraftId);
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
      'The old draft was updated in place — no new draft was created and nothing was sent.'
    );
    return;
  }

  SpreadsheetApp.getUi().alert('No suitable existing Gmail draft was found to test.');
}

/**
 * Updates every existing Gmail draft that can be matched safely to a
 * Qualified_Leads record. Existing Gmail Draft IDs are preserved.
 */
function menuUpdateExistingPartDV6Drafts() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const drafts = ss.getSheetByName(SHEET_DRAFTS);
  if (!drafts || drafts.getLastRow() < 2) {
    ui.alert('Outreach_Drafts is empty.');
    return;
  }

  const confirm = ui.alert(
    'Update existing Gmail drafts?',
    'This will replace the contents of existing Gmail drafts with Part D v6.\n\n' +
    'It will NOT send, delete, or create drafts, and it will NOT change lead qualification/filtering.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const qMap = buildQualifiedLeadMapForExistingDrafts_();
  const lastRow = drafts.getLastRow();
  const width = Math.max(DRAFT_HEADERS.length, drafts.getLastColumn());
  const data = drafts.getRange(2, 1, lastRow - 1, width).getValues();

  const emailCol = DRAFT_HEADERS.indexOf('Email') + 1;
  const subjectCol = DRAFT_HEADERS.indexOf('Subject') + 1;
  const bodyCol = DRAFT_HEADERS.indexOf('Email Draft') + 1;
  const gmailIdCol = DRAFT_HEADERS.indexOf('Gmail Draft ID') + 1;
  const placeCol = DRAFT_HEADERS.indexOf('Place ID') + 1;
  const statusCol = DRAFT_HEADERS.indexOf('Status') + 1;

  let updated = 0;
  let skipped = 0;
  let errors = 0;

  data.forEach((row, i) => {
    const placeId = String(row[placeCol - 1] || '').trim();
    const gmailDraftId = String(row[gmailIdCol - 1] || '').trim();
    const lead = qMap[placeId];

    if (!placeId || !gmailDraftId || !lead || !lead.email || String(lead.recommendedChannel || '').toUpperCase() !== 'EMAIL' || !isValidOutreachEmail(lead.email)) {
      skipped++;
      return;
    }

    try {
      const email = buildPartDEmail_(lead);
      const draft = GmailApp.getDraft(gmailDraftId);
      draft.update(lead.email, email.subject, email.body, {
        htmlBody: email.htmlBody,
        inlineImages: { vashaLogo: getVashaLogoBlob_() },
        name: 'Harshika'
      });

      drafts.getRange(i + 2, subjectCol).setValue(email.subject);
      drafts.getRange(i + 2, bodyCol).setValue(email.body);
      drafts.getRange(i + 2, statusCol).setValue('Part D v6 — Gmail draft updated');
      updated++;
    } catch (e) {
      errors++;
      drafts.getRange(i + 2, statusCol).setValue('Update error: ' + e.message);
      Logger.log('Part D existing draft update error row ' + (i + 2) + ': ' + e.message);
    }
  });

  ui.alert(
    'Part D v6 existing-draft update complete!\n\n' +
    'Updated: ' + updated + '\n' +
    'Skipped: ' + skipped + '\n' +
    'Errors: ' + errors + '\n\n' +
    'No emails were sent and no Gmail drafts were deleted.'
  );
}
