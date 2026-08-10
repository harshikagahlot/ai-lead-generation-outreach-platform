/**
 * EmailOutreachUpgrade.js
 * Part D: industry-aware, hypothesis-led outreach + VASHA HTML signature.
 * NEVER sends email automatically.
 */

const PART_D_OBSERVATION_HEADER = 'Outreach Observation';
const PART_D_HYPOTHESIS_HEADER = 'Improvement Hypothesis';
const PART_D_WHY_HEADER = 'Why It May Matter';
const PART_D_EMAIL_VERSION_HEADER = 'Email Version';
const PART_D_EMAIL_VERSION = 'Part D v3';

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

  const body = 'Hi ' + firstName + ',\n\n' + structure.text + '\n\nBest,\n\nHarshika\nVASHA Technologies\nAI Automation • Custom Software • Business Systems';

  const htmlBody = '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.6;color:#172033;">' +
    '<p>Hi ' + escapeHtml_(firstName) + ',</p>' + structure.html +
    '<p>Best,<br><br><strong>Harshika</strong><br>VASHA Technologies<br>' +
    '<span style="color:#5b667a;">AI Automation &bull; Custom Software &bull; Business Systems</span></p>' +
    '<img src="cid:vashaLogo" alt="Vasha Technologies" width="220" style="display:block;margin-top:8px;max-width:100%;height:auto;">' +
    '</div>';

  return { subject, body, htmlBody, observation, hypothesis, why, capability };
}

function getPartDStructure_(lead, observation, hypothesis, why, capability) {
  const i = String(lead.industry || '').toLowerCase();
  const context = String(lead.websiteStatus || '').toLowerCase();

  if (/no website|without website|not found/.test(context)) {
    return makeStructure_('A digital idea for ' + (lead.name || 'your business'),
      'I came across ' + (lead.name || 'your business') + ' and noticed there does not appear to be a dedicated website listed for the business.',
      'For a business like yours, a focused online presence could make it easier for new customers to understand what you offer and take the next step.',
      capability,
      'If you are already working on this, no worries. If not, I would be happy to share a simple approach.'
    );
  }

  if (/dent|clinic|chiro|vet|medical|health/.test(i)) {
    return makeStructure_('A quick idea for ' + (lead.name || 'your practice'),
      'I was looking into ' + (lead.name || 'your practice') + ' and noticed ' + observation + '.',
      'It made me wonder whether ' + hypothesis + '.', capability,
      why + ' If you already have something like this in place, no worries — I would still be happy to share the thought.'
    );
  }

  if (/property|real estate|realt|property management|apartment|leasing/.test(i)) {
    return makeStructure_('Idea for ' + (lead.name || 'your property business'),
      'I came across ' + (lead.name || 'your business') + ' and noticed ' + observation + '.',
      'One opportunity I was thinking about is whether ' + hypothesis + '.', capability,
      'That could help turn more owner interest into structured enquiries without adding much manual back-and-forth. If relevant, I can share the idea.'
    );
  }

  if (/gym|fitness|yoga|pilates|studio|sports|martial/.test(i)) {
    return makeStructure_('Quick idea for ' + (lead.name || 'your business'),
      'I was looking through ' + (lead.name || 'your business') + ' and noticed ' + observation + '.',
      'For a business that relies on local interest, I was thinking whether ' + hypothesis + '.', capability,
      'The goal would simply be to make the step from interest to enquiry a little easier. Happy to share the idea if useful.'
    );
  }

  if (/salon|barber|spa|nail|beauty|wellness/.test(i)) {
    return makeStructure_('A small idea for ' + (lead.name || 'your business'),
      'I came across ' + (lead.name || 'your business') + ' and noticed ' + observation + '.',
      'I was wondering whether ' + hypothesis + '.', capability,
      'It could make the booking journey a little easier for customers and the team. If that sounds relevant, I can send over the idea.'
    );
  }

  if (/hvac|plumb|electric|roof|landscap|cleaning|pest|contractor|construction|home service/.test(i)) {
    return makeStructure_('One idea for ' + (lead.name || 'your business'),
      'I was looking into ' + (lead.name || 'your business') + ' and noticed ' + observation + '.',
      'I had a thought that ' + hypothesis + '.', capability,
      why + ' If you would like, I can show what that could look like.'
    );
  }

  if (/auto|mechanic|repair|tire|car|collision|detailing/.test(i)) {
    return makeStructure_('Quick idea for ' + (lead.name || 'your business'),
      'I came across ' + (lead.name || 'your business') + ' and noticed ' + observation + '.',
      'I was thinking whether ' + hypothesis + '.', capability,
      'That could make it easier to capture the right details before the team follows up. If useful, I can share a simple version of the idea.'
    );
  }

  if (/law|legal|attorney|lawyer|accounting|accountant|insurance|financial/.test(i)) {
    return makeStructure_('A workflow idea for ' + (lead.name || 'your business'),
      'I was looking into ' + (lead.name || 'your business') + ' and noticed ' + observation + '.',
      'I wondered whether ' + hypothesis + '.', capability,
      why + ' I can share the idea without assuming it is something you need right now.'
    );
  }

  if (/restaurant|cafe|caf|diner|bakery|catering|food|bar|coffee/.test(i)) {
    return makeStructure_('Quick idea for ' + (lead.name || 'your business'),
      'I came across ' + (lead.name || 'your business') + ' and noticed ' + observation + '.',
      'I was thinking whether ' + hypothesis + '.', capability,
      'The aim would be to remove a little friction from the customer journey. Happy to share the thought if useful.'
    );
  }

  if (/manufactur|factory|industrial|wholesale|distribut|logistics|warehouse|supplier/.test(i)) {
    return makeStructure_('An operations idea for ' + (lead.name || 'your business'),
      'I was looking into ' + (lead.name || 'your business') + ' and noticed ' + observation + '.',
      'I wondered whether ' + hypothesis + '.', capability,
      why + ' If the process is currently handled manually, I would be happy to share a possible approach.'
    );
  }

  if (/school|college|education|training|academy|tuition|university|coaching/.test(i)) {
    return makeStructure_('A digital idea for ' + (lead.name || 'your business'),
      'I came across ' + (lead.name || 'your business') + ' and noticed ' + observation + '.',
      'I was thinking whether ' + hypothesis + '.', capability,
      'It could make the next step clearer for students, parents, or staff. If relevant, I can share the idea.'
    );
  }

  return makeStructure_('A quick idea for ' + (lead.name || 'your business'),
    'I was looking into ' + (lead.name || 'your business') + ' and noticed ' + observation + '.',
    'I had a thought that ' + hypothesis + '.', capability,
    why + ' If this is already handled another way, no worries — I can still share the idea if useful.'
  );
}

function makeStructure_(subject, opening, idea, capability, close) {
  return {
    subject: subject,
    text: opening + '\n\n' + idea + '\n\n' + capability + '\n\n' + close,
    html: '<p>' + escapeHtml_(opening) + '</p>' +
      '<p>' + escapeHtml_(idea) + '</p>' +
      '<p>' + escapeHtml_(capability) + '</p>' +
      '<p>' + escapeHtml_(close) + '</p>'
  };
}

function getPartDObservation_(lead) {
  if (lead.websiteStatus === WEBSITE_STATUS.NO_WEBSITE) return 'there does not appear to be a dedicated website listed for the business';
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
  if (/property|real estate|realt/.test(i)) return 'It could help the team capture better-qualified property-management enquiries.';
  if (/dent|clinic|chiro|vet|medical|health/.test(i)) return 'It could make the first step easier for new patients while giving the team cleaner information.';
  if (/hvac|plumb|electric|roof|auto|mechanic|repair|contractor/.test(i)) return 'It could make incoming service requests easier to qualify and follow up.';
  if (/law|legal|attorney|lawyer/.test(i)) return 'It could reduce back-and-forth before a consultation.';
  if (/manufactur|factory|industrial|wholesale|distribut|logistics/.test(i)) return 'It could reduce repetitive coordination and give the team a clearer workflow.';
  if (/restaurant|cafe|caf|diner|bakery|food/.test(i)) return 'It could make the next customer action clearer and easier.';
  if (/gym|fitness|yoga|pilates|studio/.test(i)) return 'It could make it easier to turn online interest into actual enquiries.';
  if (/school|college|education|academy|tuition|university/.test(i)) return 'It could make the enquiry process clearer while reducing repetitive questions for staff.';
  return 'The goal would be to make the next step easier for the customer and easier for the team.';
}

function getPartDCapability_(lead) {
  const i = String(lead.industry || '').toLowerCase();
  if (/property|real estate|realt/.test(i)) return 'I’m currently building VASHA Technologies around practical digital solutions for businesses, including websites, custom systems and automation — this is the kind of workflow we can help design and build.';
  if (/dent|clinic|chiro|vet|medical|health/.test(i)) return 'I’m currently building VASHA Technologies around practical digital solutions for businesses, including websites, custom software and automation — particularly where a smoother customer journey can make a difference.';
  if (/gym|fitness|yoga|pilates|studio|salon|barber|spa|nail|beauty/.test(i)) return 'I’m currently building VASHA Technologies around practical digital solutions for businesses, from customer-facing websites and booking flows to automation and internal systems.';
  if (/hvac|plumb|electric|roof|auto|mechanic|repair|contractor|cleaning|pest/.test(i)) return 'I’m currently building VASHA Technologies around practical digital solutions for businesses, including websites, enquiry systems and workflow automation.';
  if (/law|legal|attorney|lawyer|accounting|insurance|financial/.test(i)) return 'I’m currently building VASHA Technologies around practical digital solutions for businesses, including websites, custom systems and automation that can simplify information-heavy workflows.';
  if (/manufactur|factory|industrial|wholesale|distribut|logistics|warehouse/.test(i)) return 'I’m currently building VASHA Technologies around practical digital solutions for businesses, including internal systems, custom software and workflow automation.';
  if (/restaurant|cafe|caf|diner|bakery|food|catering/.test(i)) return 'I’m currently building VASHA Technologies around practical digital solutions for businesses, including customer-facing websites, enquiry flows and automation.';
  if (/school|college|education|academy|tuition|university|coaching/.test(i)) return 'I’m currently building VASHA Technologies around practical digital solutions for businesses, including websites, enquiry systems, custom software and automation.';
  return 'I’m currently building VASHA Technologies around practical digital solutions for businesses — from websites and custom software to automation and business systems.';
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/\"/g, '&quot;').replace(/'/g, '&#39;');
}
