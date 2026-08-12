/**
 * OutreachCore.js
 * -----------------------------------------------------------------------
 * Single source of truth for outreach draft orchestration.
 *
 * IMPORTANT:
 * - Lead generation, qualification, email validation and spam filtering are
 *   intentionally untouched.
 * - Email copy is generated ONLY by buildPartDEmail_() in
 *   EmailOutreachUpgrade.js.
 * - Gmail drafts are created only; nothing is sent automatically.
 */

function extractCityFromAddress(address) {
  if (!address) return '';
  let parts = String(address).split(',').map(s => s.trim()).filter(Boolean);
  const last = parts[parts.length - 1];
  if (last && /^(usa|united states|u\.s\.a\.?)$/i.test(last)) parts.pop();
  for (let i = parts.length - 1; i >= 0; i--) {
    if (/^[A-Z]{2}(\s*\d{5}(-\d{4})?)?$/i.test(parts[i])) return i > 0 ? parts[i - 1] : '';
  }
  return parts.length >= 2 ? parts[parts.length - 2] : (parts[0] || '');
}

function getExistingDraftPlaceIds(draftsSheet) {
  const ids = new Set();
  const lastRow = draftsSheet.getLastRow();
  if (lastRow < 2) return ids;
  const col = DRAFT_HEADERS.indexOf('Place ID') + 1;
  draftsSheet.getRange(2, col, lastRow - 1, 1).getValues().forEach(r => { if (r[0]) ids.add(String(r[0]).trim()); });
  return ids;
}

function mapQualifiedRowToLead(row) {
  const idx = name => QUALIFIED_HEADERS.indexOf(name);
  return {
    name: row[idx('Business Name')], industry: row[idx('Industry')], owner: row[idx('Owner')],
    email: row[idx('Email')], phone: row[idx('Phone')], website: row[idx('Website')],
    websiteStatus: row[idx('Website Status')], emailType: row[idx('Email Type')],
    recommendedChannel: row[idx('Recommended Channel')], readinessScore: row[idx('Readiness Score')],
    readinessNotes: row[idx('Readiness Notes')], rating: row[idx('Rating')], reviewCount: row[idx('Reviews')],
    address: row[idx('Address')], notes: row[idx('Notes')], placeId: row[idx('Place ID')]
  };
}

function findManualTrackingColumn(sheet, headerNameContains) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] && String(headers[i]).toLowerCase().indexOf(headerNameContains.toLowerCase()) !== -1) return i + 1;
  }
  return -1;
}

function ensureDraftsSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ensureSheetWithHeaders(ss, SHEET_DRAFTS, DRAFT_HEADERS);
  const draftCol = DRAFT_HEADERS.indexOf('Email Draft') + 1;
  sheet.setColumnWidth(draftCol, 500);
  sheet.getRange(1, draftCol, Math.max(sheet.getMaxRows(), 1000), 1).setWrap(true);
  const lastCol = sheet.getLastColumn();
  const existingHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  DRAFT_HEADERS.forEach((header, i) => {
    if ((existingHeaders[i] || '').toString().trim() !== header && i + 1 > lastCol) {
      sheet.getRange(1, i + 1).setValue(header).setFontWeight('bold');
    }
  });
  return sheet;
}

/**
 * Default draft action. It now uses Part D v6 directly — there is no legacy
 * buildEmailDraft() fallback anymore.
 */
function menuDraftOutreachEmails() {
  const ss = SpreadsheetApp.getActive();
  const qualifiedSheet = ss.getSheetByName(SHEET_QUALIFIED);
  if (!qualifiedSheet || qualifiedSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Qualified_Leads is empty — run a search first.');
    return;
  }

  const draftsSheet = ensureDraftsSheet();
  const existingPlaceIds = getExistingDraftPlaceIds(draftsSheet);
  const trackingCol = findManualTrackingColumn(qualifiedSheet, 'email update');
  const lastCol = Math.max(QUALIFIED_HEADERS.length, trackingCol > 0 ? trackingCol : 0);
  const qualifiedData = qualifiedSheet.getRange(2, 1, qualifiedSheet.getLastRow() - 1, lastCol).getValues();
  const newDraftRows = [];
  let skipped = 0, skippedSent = 0, skippedInvalid = 0;

  qualifiedData.forEach(row => {
    if (trackingCol > 0 && String(row[trackingCol - 1] || '').toLowerCase().trim() === 'sent') { skippedSent++; return; }
    const lead = mapQualifiedRowToLead(row);
    if (existingPlaceIds.has(String(lead.placeId || '').trim())) { skipped++; return; }
    if (!lead.email || String(lead.recommendedChannel || '').trim().toUpperCase() !== 'EMAIL' || !isValidOutreachEmail(lead.email)) {
      skippedInvalid++; return;
    }
    if ((Number(lead.readinessScore) || 0) < 50 || !hasConcreteObservation(lead)) { skippedInvalid++; return; }

    const email = buildPartDEmail_(lead);
    newDraftRows.push([
      lead.name, lead.industry, extractCityFromAddress(lead.address), lead.email,
      lead.phone, lead.websiteStatus, lead.recommendedChannel || '', email.subject, email.body,
      'Draft', lead.placeId, ''
    ]);
  });

  if (newDraftRows.length) {
    const startRow = draftsSheet.getLastRow() + 1;
    draftsSheet.getRange(startRow, 1, newDraftRows.length, DRAFT_HEADERS.length).setValues(newDraftRows);
  }

  SpreadsheetApp.getUi().alert(
    'Part D v6: drafted ' + newDraftRows.length + ' new email(s).\n' +
    'Skipped existing: ' + skipped + '\n' +
    'Skipped invalid/non-email/low-readiness: ' + skippedInvalid + '\n' +
    (trackingCol > 0 ? 'Skipped marked sent: ' + skippedSent : '')
  );
}

function menuRedraftAllOutreachEmails() {
  const draftsSheet = ensureDraftsSheet();
  clearSheetBody(draftsSheet);
  menuDraftOutreachEmails();
}

/**
 * Pushes the exact Part D v6 HTML email to Gmail, including the VASHA logo.
 * The qualified lead is looked up by Place ID so the full Part D data is
 * available when the Gmail draft is created.
 */
function menuPushDraftsToGmail() {
  const ss = SpreadsheetApp.getActive();
  ensureDraftsSheet();
  const draftsSheet = ss.getSheetByName(SHEET_DRAFTS);
  if (!draftsSheet || draftsSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('No drafts to push — the Outreach_Drafts sheet is empty.');
    return;
  }

  const qualifiedSheet = ss.getSheetByName(SHEET_QUALIFIED);
  const qMap = {};
  if (qualifiedSheet && qualifiedSheet.getLastRow() >= 2) {
    const qRows = qualifiedSheet.getRange(2, 1, qualifiedSheet.getLastRow() - 1, QUALIFIED_HEADERS.length).getValues();
    qRows.forEach(row => {
      const lead = mapQualifiedRowToLead(row);
      if (lead.placeId) qMap[String(lead.placeId).trim()] = lead;
    });
  }

  const lastRow = draftsSheet.getLastRow();
  const width = Math.max(DRAFT_HEADERS.length, draftsSheet.getLastColumn());
  const data = draftsSheet.getRange(2, 1, lastRow - 1, width).getValues();
  const emailCol = DRAFT_HEADERS.indexOf('Email') + 1;
  const subjectCol = DRAFT_HEADERS.indexOf('Subject') + 1;
  const bodyCol = DRAFT_HEADERS.indexOf('Email Draft') + 1;
  const statusCol = DRAFT_HEADERS.indexOf('Status') + 1;
  const gmailIdCol = DRAFT_HEADERS.indexOf('Gmail Draft ID') + 1;
  const channelCol = DRAFT_HEADERS.indexOf('Recommended Channel') + 1;
  const placeCol = DRAFT_HEADERS.indexOf('Place ID') + 1;

  const statusValues = data.map(row => [row[statusCol - 1] || '']);
  const gmailIdValues = data.map(row => [row[gmailIdCol - 1] || '']);
  let created = 0, skipped = 0, errors = 0;

  data.forEach((row, i) => {
    if (String(row[gmailIdCol - 1] || '').trim()) { skipped++; return; }
    const recipient = String(row[emailCol - 1] || '').trim();
    const placeId = String(row[placeCol - 1] || '').trim();
    const lead = qMap[placeId];
    if (!recipient || String(row[channelCol - 1] || '').toUpperCase() !== 'EMAIL' || !isValidOutreachEmail(recipient)) {
      statusValues[i] = ['Skipped: Invalid/non-email recipient']; skipped++; return;
    }

    try {
      const email = lead ? buildPartDEmail_(lead) : {
        subject: String(row[subjectCol - 1] || '').trim(),
        body: String(row[bodyCol - 1] || '').trim(),
        htmlBody: '<div style="font-family:Arial,sans-serif;white-space:pre-wrap;">' + escapeHtml_(String(row[bodyCol - 1] || '').trim()) + '</div>'
      };
      if (!email.subject || !email.body) { statusValues[i] = ['Skipped: Missing subject/body']; skipped++; return; }
      const draft = GmailApp.createDraft(recipient, email.subject, email.body, {
        htmlBody: email.htmlBody,
        inlineImages: { vashaLogo: getVashaLogoBlob_() },
        name: 'Harshika'
      });
      gmailIdValues[i] = [draft.getId()];
      statusValues[i] = ['Pushed to Gmail'];
      created++;
    } catch (e) {
      statusValues[i] = ['Error: ' + e.message];
      gmailIdValues[i] = [''];
      errors++;
      Logger.log('Gmail draft error row ' + (i + 2) + ': ' + e.message);
    }
  });

  draftsSheet.getRange(2, statusCol, statusValues.length, 1).setValues(statusValues);
  draftsSheet.getRange(2, gmailIdCol, gmailIdValues.length, 1).setValues(gmailIdValues);
  SpreadsheetApp.getUi().alert('Push to Gmail complete!\n\nCreated: ' + created + '\nSkipped: ' + skipped + '\nErrors: ' + errors + '\n\nNothing was sent automatically.');
}

function menuViewEmailOpens() {
  const ss = SpreadsheetApp.getActive();
  const opensSheet = ss.getSheetByName(SHEET_OPENS);
  const ui = SpreadsheetApp.getUi();
  if (!opensSheet || opensSheet.getLastRow() < 2) { ui.alert('No email open events recorded yet in "Email_Opens".'); return; }
  const opensData = opensSheet.getRange(2, 1, opensSheet.getLastRow() - 1, OPENS_HEADERS.length).getValues();
  const opensMap = {};
  opensData.forEach(row => {
    const timestamp = row[0], leadId = String(row[1] || '').trim(), bizName = String(row[2] || '').trim();
    if (!leadId) return;
    if (!opensMap[leadId]) opensMap[leadId] = { name: bizName || leadId, count: 0, firstOpened: timestamp, lastOpened: timestamp };
    opensMap[leadId].count++;
    opensMap[leadId].lastOpened = timestamp;
    if (bizName && (!opensMap[leadId].name || opensMap[leadId].name === leadId)) opensMap[leadId].name = bizName;
  });
  const leadIds = Object.keys(opensMap);
  if (!leadIds.length) { ui.alert('No valid lead opens recorded.'); return; }
  let message = 'EMAIL OPENS SUMMARY (' + leadIds.length + ' leads recorded)\n\n';
  leadIds.slice(0, 15).forEach((id, idx) => {
    const item = opensMap[id];
    const firstStr = item.firstOpened ? Utilities.formatDate(new Date(item.firstOpened), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : 'N/A';
    const lastStr = item.lastOpened ? Utilities.formatDate(new Date(item.lastOpened), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : 'N/A';
    message += (idx + 1) + '. ' + item.name + '\nOpened: Yes (' + item.count + ')\nFirst: ' + firstStr + ' | Last: ' + lastStr + '\n\n';
  });
  if (leadIds.length > 15) message += '... and ' + (leadIds.length - 15) + ' more. Check Email_Opens for full history.';
  ui.alert(message);
}

function getTrackingPixelHtml(leadId, settings) {
  if (!leadId) return '';
  const webAppUrl = settings && settings['Web App URL'] ? String(settings['Web App URL']).trim() : '';
  if (!webAppUrl) return '';
  return '<img src="' + webAppUrl + '?leadId=' + encodeURIComponent(leadId) + '" width="1" height="1" style="display:none !important;" alt="" />';
}
