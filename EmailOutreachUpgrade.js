/**
 * EmailOutreachUpgrade.js
 * Part D v6: concise, observation-led, business-specific cold outreach + VASHA signature.
 * NEVER sends email automatically.
 *
 * Copy principles:
 * - Lead with the real observation, not a generic compliment.
 * - Connect the observation to one concrete business-process hypothesis.
 * - Keep the non-assumption language, but do not repeat the same generic paragraph.
 * - Mention VASHA clearly without turning the email into a company pitch.
 * - Keep the email short enough for genuine cold outreach.
 *
 * IMPORTANT:
 * - This file does NOT change lead qualification or email filtering.
 * - Existing spam/non-working-email protections remain untouched elsewhere.
 */

const PART_D_OBSERVATION_HEADER = 'Outreach Observation';
const PART_D_HYPOTHESIS_HEADER = 'Improvement Hypothesis';
const PART_D_WHY_HEADER = 'Why It May Matter';
const PART_D_EMAIL_VERSION_HEADER = 'Email Version';
const PART_D_EMAIL_VERSION = 'Part D v6';

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

    if (!placeId || /sent|replied|follow-up/i.test(status)) { skipped++; return; }
    // v6 upgrades older Part D drafts; future runs skip already-upgraded v6 drafts.
    if (existingVersion === PART_D_EMAIL_VERSION && gmailId) { skipped++; return; }

    const lead = qByPlace[placeId];
    if (!lead || !lead.email || String(lead.recommendedChannel || '').trim().toUpperCase() !== 'EMAIL') { skipped++; return; }
    if (!isValidOutreachEmail(lead.email)) { skipped++; return; }

    const readinessScore = Number(lead.readinessScore) || 0;
    if (readinessScore < 50 || !hasConcreteObservation(lead)) { skipped++; return; }

    try {
      const email = buildPartDEmail_(lead);
      obsValues[i] = [email.observation];
      hypValues[i] = [email.hypothesis];
      whyValues[i] = [email.why];
      versionValues[i] = [PART_D_EMAIL_VERSION];

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
    'Part D v6 complete.\n\n' +
    'Upgraded: ' + upgraded + '\n' +
    'New Gmail drafts: ' + created + '\n' +
    'Updated Gmail drafts: ' + updated + '\n' +
    'Skipped: ' + skipped + '\n' +
    'Errors: ' + errors + '\n\n' +
    'Nothing was sent automatically.'
  );
}

function ensurePartDColumns_(sheet) {
  [PART_D_OBSERVATION_HEADER, PART_D_HYPOTHESIS_HEADER, PART_D_WHY_HEADER, PART_D_EMAIL_VERSION_HEADER].forEach(header => {
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    if (headers.indexOf(header) === -1) sheet.getRange(1, lastCol + 1).setValue(header).setFontWeight('bold');
  });
}

function buildPartDEmail_(lead) {
  const business = lead.name || 'your business';
  const firstName = lead.owner ? String(lead.owner).split(/\s+/)[0] : 'there';
  const observation = getPartDObservation_(lead);
  const hypothesis = getPartDHypothesis_(lead);
  const why = getPartDWhy_(lead);
  const capability = getPartDCapability_(lead);
  const structure = getPartDStructure_(lead, observation, hypothesis, why, capability);
  const subject = structure.subject || 'A quick idea for ' + business;

  const body =
    'Hi ' + firstName + ',\n\n' +
    structure.text +
    '\n\nBest,\n\nHarshika\nBusiness Development | VASHA Technologies\n' +
    'AI Automation • Custom Software • Business Systems\n' +
    '📧 harshikagahlot01@gmail.com';

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5;color:#172033;max-width:640px;">' +
      '<p style="margin:0 0 12px;">Hi ' + escapeHtml_(firstName) + ',</p>' +
      structure.html +
      '<p style="margin:16px 0 0;">Best,<br><br>' +
        '<strong>Harshika</strong><br>' +
        'Business Development | VASHA Technologies<br>' +
        '<span style="color:#5b667a;">AI Automation &bull; Custom Software &bull; Business Systems</span><br>' +
        '<span style="color:#5b667a;">&#128231; harshikagahlot01@gmail.com</span>' +
      '</p>' +
      '<img src="cid:vashaLogo" alt="VASHA Technologies" width="110" style="display:block;width:110px;max-width:110px;height:auto;margin-top:8px;">' +
    '</div>';

  return { subject, body, htmlBody, observation, hypothesis, why, capability };
}

function getPartDStructure_(lead, observation, hypothesis, why, capability) {
  const business = lead.name || 'your business';
  const i = String(lead.industry || '').toLowerCase();
  const context = String(lead.websiteStatus || '').toLowerCase();

  // New v6 structure: observation -> relevant implication -> one hypothesis -> VASHA -> low-pressure CTA.
  // This deliberately removes the repeated generic "a lot of businesses..." paragraph.
  if (/no website|without website|not found/.test(context)) {
    return makeStructure_(
      'A quick idea for ' + business,
      'I came across ' + business + ' and noticed there does not appear to be a dedicated website listed for the business.',
      'That can make the first step harder for someone discovering the business online, especially when they are deciding whether to enquire.',
      'I was wondering whether a simple website or enquiry flow could make it easier to understand what you offer and take the next step.',
      capability + ' If useful, I can send 2–3 concrete ideas based on what I noticed.'
    );
  }

  let subject = 'A quick idea for ' + business;
  if (/law|legal|attorney|lawyer|accounting|accountant|insurance|financial/.test(i)) subject = 'A workflow idea for ' + business;
  if (/manufactur|factory|industrial|wholesale|distribut|logistics|warehouse|supplier/.test(i)) subject = 'An operations idea for ' + business;
  if (/school|college|education|training|academy|tuition|university|coaching/.test(i)) subject = 'A digital idea for ' + business;

  const implication = getPartDImplication_(lead, observation);
  const idea = 'I was wondering whether ' + hypothesis + '.';
  const close = capability + ' I cannot tell from the outside if this is actually a priority, but if it is relevant, I can send 2–3 concrete ideas based on what I noticed.';

  return makeStructure_(
    subject,
    'I came across ' + business + ' and noticed ' + observation + '.',
    implication,
    idea + ' ' + why,
    close
  );
}

function getPartDImplication_(lead, observation) {
  const i = String(lead.industry || '').toLowerCase();
  if (/property|real estate|realt/.test(i)) {
    return 'That caught my attention because property-management businesses often have several enquiry and follow-up steps between an owner showing interest and becoming a real opportunity.';
  }
  if (/dent|clinic|chiro|vet|medical|health/.test(i)) {
    return 'That caught my attention because a prospective patient often needs a very clear next step before deciding to call, book, or submit an enquiry.';
  }
  if (/hvac|plumb|electric|roof|landscap|cleaning|pest|contractor/.test(i)) {
    return 'That caught my attention because service businesses can lose time when enquiries arrive without the details needed to qualify or follow up.';
  }
  if (/law|legal|attorney|lawyer/.test(i)) {
    return 'That caught my attention because legal enquiries often involve collecting enough context before the team can decide what happens next.';
  }
  if (/salon|barber|spa|nail|beauty|wellness/.test(i)) {
    return 'That caught my attention because the gap between someone browsing and actually booking is often only a few small steps.';
  }
  if (/auto|mechanic|repair|tire|car|collision/.test(i)) {
    return 'That caught my attention because service requests are easier to handle when the right vehicle and job details are captured early.';
  }
  if (/manufactur|factory|industrial|wholesale|distribut|logistics|warehouse|supplier/.test(i)) {
    return 'That caught my attention because operational requests can create a surprising amount of manual coordination when information moves between people or systems.';
  }
  if (/restaurant|cafe|caf|diner|bakery|catering|food/.test(i)) {
    return 'That caught my attention because small points of friction can affect whether a customer takes the next step.';
  }
  if (/gym|fitness|yoga|pilates|studio|sports/.test(i)) {
    return 'That caught my attention because turning online interest into an enquiry or booking usually depends on making the next step obvious.';
  }
  if (/school|college|education|academy|tuition|university|coaching/.test(i)) {
    return 'That caught my attention because prospective students and parents often need a clear path from an initial question to the next step.';
  }
  return 'That caught my attention because small gaps around enquiries, follow-ups, or handoffs can create unnecessary friction even when the core business is working well.';
}

function makeStructure_(subject, opening, implication, idea, close) {
  return {
    subject: subject,
    text: opening + '\n\n' + implication + '\n\n' + idea + '\n\n' + close,
    html:
      '<p style="margin:0 0 12px;">' + escapeHtml_(opening) + '</p>' +
      '<p style="margin:0 0 12px;">' + escapeHtml_(implication) + '</p>' +
      '<p style="margin:0 0 12px;">' + escapeHtml_(idea) + '</p>' +
      '<p style="margin:0;">' + escapeHtml_(close) + '</p>'
  };
}

function getPartDObservation_(lead) {
  if (lead.websiteStatus === WEBSITE_STATUS.NO_WEBSITE) {
    return 'there does not appear to be a dedicated website listed for the business';
  }
  const obs = getStrongestObservation(lead.notes, lead.websiteStatus);
  return obs || 'there may be a small gap in the current online experience';
}

function getPartDHypothesis_(lead) {
  const i = String(lead.industry || '').toLowerCase();
  if (/property|real estate|realt/.test(i)) return 'a simple owner enquiry or property-intake flow could make it easier to capture and qualify new management opportunities';
  if (/dent|clinic|chiro|vet|medical|health/.test(i)) return 'a simpler appointment or patient-intake flow could reduce friction before someone contacts the practice';
  if (/hvac|plumb|electric|roof|landscap|cleaning|pest|contractor/.test(i)) return 'a short quote or service-request flow could make it easier to turn website visitors into qualified enquiries';
  if (/law|legal|attorney|lawyer/.test(i)) return 'a guided intake flow could make it easier to collect the right information before a consultation';
  if (/salon|barber|spa|nail|beauty|wellness/.test(i)) return 'a cleaner booking and follow-up flow could make it easier to convert visitors into appointments';
  if (/auto|mechanic|repair|tire|car|collision/.test(i)) return 'a simple service-request flow could make it easier to capture vehicle details and qualify enquiries';
  if (/manufactur|factory|industrial|wholesale|distribut|logistics/.test(i)) return 'a lightweight internal workflow could reduce manual coordination around requests, approvals, or reporting';
  if (/restaurant|cafe|caf|diner|bakery|catering|food/.test(i)) return 'a smoother enquiry, ordering, or repeat-customer flow could remove friction from the customer journey';
  if (/gym|fitness|yoga|pilates|studio|sports/.test(i)) return 'a simple membership or class-enquiry flow could make it easier to turn interest into a conversation';
  if (/school|college|education|academy|tuition|university|coaching/.test(i)) return 'a clearer enquiry or application flow could make it easier for students and parents to take the next step';
  return 'a small digital workflow could remove friction from enquiries, follow-ups, or day-to-day operations';
}

function getPartDWhy_(lead) {
  const i = String(lead.industry || '').toLowerCase();
  if (/property|real estate|realt/.test(i)) return 'If relevant, that could make property-management enquiries easier to capture and qualify.';
  if (/dent|clinic|chiro|vet|medical|health/.test(i)) return 'If relevant, that could make the first step easier for new patients while giving the team cleaner information.';
  if (/hvac|plumb|electric|roof|auto|mechanic|repair|contractor/.test(i)) return 'If relevant, that could make incoming service requests easier to qualify and follow up.';
  if (/law|legal|attorney|lawyer/.test(i)) return 'If relevant, that could reduce some of the back-and-forth before a consultation.';
  if (/manufactur|factory|industrial|wholesale|distribut|logistics/.test(i)) return 'If relevant, that could reduce repetitive coordination and give the team a clearer workflow.';
  if (/restaurant|cafe|caf|diner|bakery|food/.test(i)) return 'If relevant, that could make the next customer action clearer and easier.';
  if (/gym|fitness|yoga|pilates|studio/.test(i)) return 'If relevant, that could make it easier to turn online interest into an enquiry.';
  if (/school|college|education|academy|tuition|university/.test(i)) return 'If relevant, that could make the enquiry process clearer while reducing repetitive questions for staff.';
  return 'If relevant, the goal would simply be to make the next step easier for the customer and easier for the team.';
}

function getPartDCapability_(lead) {
  return 'At VASHA Technologies, I am building practical solutions around AI automation, custom software, and business systems.';
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
