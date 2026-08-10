/**
 * EmailOutreachUpgrade.js
 * Part D: compact, hypothesis-led outreach + VASHA HTML signature.
 *
 * This is intentionally a separate upgrade layer so Part A/B/C remain intact.
 * Run menuUpgradeOutreachEmails() after generating leads. It:
 *   1) matches Outreach_Drafts to Qualified_Leads by Place ID,
 *   2) creates a short, specific, multi-service email,
 *   3) records the observation + improvement hypothesis,
 *   4) creates/updates the Gmail draft with the VASHA logo inline.
 *
 * It NEVER sends email automatically.
 */

const PART_D_OBSERVATION_HEADER = 'Outreach Observation';
const PART_D_HYPOTHESIS_HEADER = 'Improvement Hypothesis';
const PART_D_WHY_HEADER = 'Why It May Matter';
const PART_D_EMAIL_VERSION_HEADER = 'Email Version';
const PART_D_EMAIL_VERSION = 'Part D v1';

function menuUpgradeOutreachEmails() {
  const ss = SpreadsheetApp.getActive();
  const qualified = ss.getSheetByName(SHEET_QUALIFIED);
  const drafts = ss.getSheetByName(SHEET_DRAFTS);
  if (!qualified || qualified.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Qualified_Leads is empty. Generate leads first.');
    return;
  }
  if (!drafts || drafts.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Outreach_Drafts is empty. Run "Draft Outreach Emails" first.');
    return;
  }

  ensurePartDColumns_(drafts);

  const qRows = qualified.getRange(2, 1, qualified.getLastRow() - 1, QUALIFIED_HEADERS.length).getValues();
  const qByPlace = {};
  qRows.forEach(row => {
    const lead = mapQualifiedRowToLead(row);
    if (lead.placeId) qByPlace[String(lead.placeId).trim()] = lead;
  });

  const headers = drafts.getRange(1, 1, 1, drafts.getLastColumn()).getValues()[0].map(String);
  const col = h => headers.indexOf(h);
  const data = drafts.getRange(2, 1, drafts.getLastRow() - 1, drafts.getLastColumn()).getValues();

  const obsValues = data.map(r => [r[col(PART_D_OBSERVATION_HEADER)] || '']);
  const hypValues = data.map(r => [r[col(PART_D_HYPOTHESIS_HEADER)] || '']);
  const whyValues = data.map(r => [r[col(PART_D_WHY_HEADER)] || '']);
  const versionValues = data.map(r => [r[col(PART_D_EMAIL_VERSION_HEADER)] || '']);

  const statusCol = col('Status');
  const gmailIdCol = col('Gmail Draft ID');
  const placeCol = col('Place ID');
  const subjectCol = col('Subject');
  const bodyCol = col('Email Draft');

  let upgraded = 0, created = 0, updated = 0, skipped = 0, errors = 0;

  data.forEach((row, i) => {
    const status = String(row[statusCol] || '').trim();
    const placeId = String(row[placeCol] || '').trim();
    const gmailId = String(row[gmailIdCol] || '').trim();
    const existingVersion = String(row[col(PART_D_EMAIL_VERSION_HEADER)] || '').trim();

    // Do not touch sent/replied/follow-up rows.
    if (!placeId || /sent|replied|follow-up/i.test(status)) { skipped++; return; }
    if (existingVersion === PART_D_EMAIL_VERSION && gmailId) { skipped++; return; }

    const lead = qByPlace[placeId];
    if (!lead) { skipped++; return; }
    if (!lead.email || String(lead.recommendedChannel || '').trim().toUpperCase() !== 'EMAIL') {
      skipped++;
      return;
    }
    if (!isValidOutreachEmail(lead.email)) { skipped++; return; }

    const readinessScore = Number(lead.readinessScore) || 0;
    if (readinessScore < 50 || !hasConcreteObservation(lead)) { skipped++; return; }

    try {
      const email = buildPartDEmail_(lead);

      obsValues[i] = [email.observation];
      hypValues[i] = [email.hypothesis];
      whyValues[i] = [email.why];
      versionValues[i] = [PART_D_EMAIL_VERSION];

      // Keep the sheet as the human-readable source of truth.
      drafts.getRange(i + 2, subjectCol + 1).setValue(email.subject);
      drafts.getRange(i + 2, bodyCol + 1).setValue(email.body);

      const options = {
        htmlBody: email.htmlBody,
        inlineImages: { vashaLogo: getVashaLogoBlob_() },
        name: 'Harshika'
      };

      if (gmailId) {
        try {
          GmailApp.getDraft(gmailId).update(lead.email, email.subject, email.body, options);
          updated++;
        } catch (updateErr) {
          // If the old ID is stale, create exactly one replacement draft.
          const draft = GmailApp.createDraft(lead.email, email.subject, email.body, options);
          drafts.getRange(i + 2, gmailIdCol + 1).setValue(draft.getId());
          created++;
        }
      } else {
        const draft = GmailApp.createDraft(lead.email, email.subject, email.body, options);
        drafts.getRange(i + 2, gmailIdCol + 1).setValue(draft.getId());
        created++;
      }

      drafts.getRange(i + 2, statusCol + 1).setValue('Part D — Gmail Draft Ready');
      upgraded++;
    } catch (err) {
      errors++;
      Logger.log('Part D error for row ' + (i + 2) + ': ' + err.message);
    }
  });

  drafts.getRange(2, col(PART_D_OBSERVATION_HEADER) + 1, obsValues.length, 1).setValues(obsValues);
  drafts.getRange(2, col(PART_D_HYPOTHESIS_HEADER) + 1, hypValues.length, 1).setValues(hypValues);
  drafts.getRange(2, col(PART_D_WHY_HEADER) + 1, whyValues.length, 1).setValues(whyValues);
  drafts.getRange(2, col(PART_D_EMAIL_VERSION_HEADER) + 1, versionValues.length, 1).setValues(versionValues);

  SpreadsheetApp.getUi().alert(
    'Part D complete.\n\n' +
    'Upgraded: ' + upgraded + '\n' +
    'New Gmail drafts: ' + created + '\n' +
    'Updated Gmail drafts: ' + updated + '\n' +
    'Skipped: ' + skipped + '\n' +
    'Errors: ' + errors + '\n\n' +
    'Nothing was sent automatically.'
  );
}

function ensurePartDColumns_(sheet) {
  [PART_D_OBSERVATION_HEADER, PART_D_HYPOTHESIS_HEADER, PART_D_WHY_HEADER, PART_D_EMAIL_VERSION_HEADER]
    .forEach(header => {
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
      if (headers.indexOf(header) === -1) {
        sheet.getRange(1, lastCol + 1).setValue(header).setFontWeight('bold');
      }
    });
}

function buildPartDEmail_(lead) {
  const business = lead.name || 'your business';
  const firstName = lead.owner ? String(lead.owner).split(/\s+/)[0] : 'there';
  const observation = getPartDObservation_(lead);
  const hypothesis = getPartDHypothesis_(lead);
  const why = getPartDWhy_(lead);

  const subject = 'A quick idea for ' + business;

  const body =
    'Hi ' + firstName + ',\n\n' +
    'I was looking into ' + business + ' and noticed ' + observation + '.\n\n' +
    'I had a thought that ' + hypothesis + '.\n\n' +
    'Not sure if this is already something you have in place, but if it is relevant, I can share the idea in a little more detail.\n\n' +
    'Best,\n\n' +
    'Harshika\n' +
    'VASHA Technologies\n' +
    'AI Automation • Custom Software • Business Systems';

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.6;color:#172033;">' +
    '<p>Hi ' + escapeHtml_(firstName) + ',</p>' +
    '<p>I was looking into ' + escapeHtml_(business) + ' and noticed ' + escapeHtml_(observation) + '.</p>' +
    '<p>I had a thought that ' + escapeHtml_(hypothesis) + '.</p>' +
    '<p>Not sure if this is already something you have in place, but if it is relevant, I can share the idea in a little more detail.</p>' +
    '<p>Best,<br><br><strong>Harshika</strong><br>VASHA Technologies<br>' +
    '<span style="color:#5b667a;">AI Automation &bull; Custom Software &bull; Business Systems</span></p>' +
    '<img src="cid:vashaLogo" alt="Vasha Technologies" width="360" style="display:block;margin-top:8px;max-width:100%;height:auto;">' +
    '</div>';

  return { subject: subject, body: body, htmlBody: htmlBody, observation: observation, hypothesis: hypothesis, why: why };
}

function getPartDObservation_(lead) {
  if (lead.websiteStatus === WEBSITE_STATUS.NO_WEBSITE) {
    return 'there does not appear to be a dedicated website listed for the business';
  }

  const obs = getStrongestObservation(lead.notes, lead.websiteStatus);
  return obs || 'there may be a small gap in the current online experience';
}

function getPartDHypothesis_(lead) {
  const i = (lead.industry || '').toLowerCase();

  if (/property|real estate|realt/.test(i)) {
    return 'a simple owner enquiry or property-intake flow could make it easier to capture and qualify new management opportunities';
  }
  if (/dent/.test(i)) {
    return 'a simpler appointment or patient-intake flow could reduce friction before someone contacts the practice';
  }
  if (/hvac|plumb|electric|roof/.test(i)) {
    return 'a short quote/request flow could make it easier to turn website visitors into qualified service enquiries';
  }
  if (/law|legal|attorney/.test(i)) {
    return 'a guided intake flow could make it easier to collect the right information before a consultation';
  }
  if (/salon|barber|spa|nail/.test(i)) {
    return 'a cleaner booking and follow-up flow could make it easier to convert visitors into appointments';
  }
  if (/auto|mechanic|repair shop/.test(i)) {
    return 'a simple service-request flow could make it easier to capture vehicle details and qualify enquiries';
  }
  if (/manufactur|factory|production/.test(i)) {
    return 'a lightweight internal workflow could reduce manual coordination around requests, approvals, or reporting';
  }
  if (/restaurant|caf|diner|bakery/.test(i)) {
    return 'a smoother enquiry, ordering, or repeat-customer flow could remove a little friction from the customer journey';
  }
  if (/gym|fitness|yoga|pilates|studio/.test(i)) {
    return 'a simple membership or class-enquiry flow could make it easier to turn interest into a conversation';
  }

  return 'a small digital workflow could remove friction from enquiries, follow-ups, or day-to-day operations';
}

function getPartDWhy_(lead) {
  const i = (lead.industry || '').toLowerCase();
  if (/property|real estate|realt/.test(i)) return 'It could help the team capture better-qualified property-management enquiries.';
  if (/dent|clinic|chiro|vet/.test(i)) return 'It could make the first step easier for new patients while giving the team cleaner information.';
  if (/hvac|plumb|electric|roof|auto|mechanic/.test(i)) return 'It could make incoming service requests easier to qualify and follow up.';
  if (/law|legal|attorney/.test(i)) return 'It could reduce back-and-forth before a consultation.';
  return 'The goal would be to make the next step easier for the customer and easier for the team.';
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
