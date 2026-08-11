/**
 * VashaOutreachDrafts.js
 * -----------------------------------------------------------------------
 * Corrected outreach-draft builder.
 *
 * IMPORTANT:
 * - Uses the existing Qualified_Leads sheet and existing email validation.
 * - Does NOT change lead generation, qualification, spam/placeholder filtering,
 *   or the existing Gmail architecture.
 * - This is the actual new-draft path; the old EmailDrafts.js template is not
 *   used by this function.
 * - VASHA is described honestly as an early-stage initiative being built,
 *   never as an established client-service company.
 */

function menuDraftVashaOutreachEmails() {
  const ss = SpreadsheetApp.getActive();
  const qualifiedSheet = ss.getSheetByName(SHEET_QUALIFIED);
  if (!qualifiedSheet || qualifiedSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Qualified_Leads is empty — generate leads first.');
    return;
  }

  ensureDraftsSheet();
  const draftsSheet = ss.getSheetByName(SHEET_DRAFTS);
  const existingPlaceIds = getExistingDraftPlaceIds(draftsSheet);
  const settings = getSettings();
  const senderName = settings[YOUR_NAME_SETTING] || 'Harshika Gahlot';

  const trackingCol = findManualTrackingColumn(qualifiedSheet, 'email update');
  const lastRow = qualifiedSheet.getLastRow();
  const lastCol = Math.max(QUALIFIED_HEADERS.length, trackingCol > 0 ? trackingCol : 0);
  const qualifiedData = qualifiedSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  let created = 0, skipped = 0, invalid = 0, sentMarked = 0;
  const newRows = [];

  qualifiedData.forEach(row => {
    if (trackingCol > 0) {
      const trackingValue = String(row[trackingCol - 1] || '').trim().toLowerCase();
      if (trackingValue === 'sent') { sentMarked++; return; }
    }

    const lead = mapQualifiedRowToLead(row);
    if (!lead.placeId || existingPlaceIds.has(lead.placeId)) { skipped++; return; }

    // Keep the existing email safety architecture. We do not loosen this.
    if (!lead.email || !isValidOutreachEmail(lead.email)) {
      invalid++;
      return;
    }

    const email = buildVashaOutreachEmail_(lead, senderName);
    newRows.push([
      lead.name,
      lead.industry,
      extractCityFromAddress(lead.address),
      lead.email,
      lead.phone,
      lead.websiteStatus,
      'Email',
      email.subject,
      email.body,
      'Draft',
      lead.placeId,
      ''
    ]);
    existingPlaceIds.add(lead.placeId);
    created++;
  });

  if (newRows.length) {
    const startRow = draftsSheet.getLastRow() + 1;
    draftsSheet.getRange(startRow, 1, newRows.length, DRAFT_HEADERS.length).setValues(newRows);
  }

  SpreadsheetApp.getUi().alert(
    'VASHA outreach drafts created.\n\n' +
    'New drafts: ' + created + '\n' +
    'Skipped (already drafted): ' + skipped + '\n' +
    'Skipped (invalid/placeholder email): ' + invalid + '\n' +
    (trackingCol > 0 ? 'Skipped (marked sent): ' + sentMarked + '\n' : '') +
    '\nNext: use "9. Push Drafts to Gmail".'
  );
}

function buildVashaOutreachEmail_(lead, senderName) {
  const business = lead.name || 'your business';
  const industry = lead.industry || 'your industry';
  const city = extractCityFromAddress(lead.address);
  const location = city ? ' in ' + city : '';
  const observation = getVashaObservation_(lead);
  const opportunity = getVashaOpportunity_(industry);
  const subject = 'A quick idea for ' + business;

  const body =
    'Hi there,\n\n' +
    'While looking into ' + business + location + ', I noticed ' + observation + '.\n\n' +
    'A lot of businesses do not necessarily have a marketing problem. Sometimes the bigger opportunity is the system behind the customer journey — enquiries, follow-ups, operations, or handoffs.\n\n' +
    'I cannot tell from the outside whether any of that is actually a priority for ' + business + ', so I do not want to assume there is a problem.\n\n' +
    'Depending on how you currently operate, ' + opportunity + '. If something like this is already handled well on your side, please ignore the thought.\n\n' +
    'I am currently building VASHA Technologies as an early-stage initiative around AI automation, custom software, and practical business systems. If the idea is relevant, I can send over 2–3 concrete ideas based on what I noticed — no pitch deck, just a short, useful message.\n\n' +
    'Best regards,\n\n' +
    senderName + '\n' +
    'Business Development | VASHA Technologies\n' +
    'AI Automation • Custom Software • Business Systems\n' +
    'harshikagahlot01@gmail.com';

  return { subject: subject, body: body };
}

function getVashaObservation_(lead) {
  const rating = Number(lead.rating) || 0;
  const reviews = Number(lead.reviewCount) || 0;

  if (lead.websiteStatus === WEBSITE_STATUS.NO_WEBSITE) {
    return 'there does not appear to be a dedicated website listed for the business';
  }

  const concrete = getStrongestObservation(lead.notes, lead.websiteStatus);
  if (concrete && concrete !== 'the site could use a bit of a refresh') {
    if (rating >= 4.5 && reviews >= 50) {
      return 'a ' + rating + '★ rating across ' + reviews + ' reviews, along with ' + concrete;
    }
    return concrete;
  }

  if (rating >= 4.5 && reviews >= 50) {
    return 'a ' + rating + '★ rating across ' + reviews + ' reviews — genuinely strong customer feedback';
  }
  if (rating >= 4.0 && reviews >= 15) {
    return 'consistently positive customer feedback (' + rating + '★ across ' + reviews + ' reviews)';
  }
  if (reviews > 0) {
    return 'real customer reviews on Google, which says a lot about the trust you have already built locally';
  }
  return 'the work you are doing in the ' + (lead.industry || 'local') + ' space';
}

function getVashaOpportunity_(industry) {
  const i = String(industry || '').toLowerCase();

  if (/dent|clinic|chiro|vet|medical|health/.test(i)) {
    return 'there may be useful opportunities around appointment enquiries, patient intake, reminders, or follow-ups that could make the next step easier';
  }
  if (/property|real estate|realt/.test(i)) {
    return 'there may be useful opportunities around owner enquiries, property intake, tenant communication, or maintenance workflows that could make requests easier to capture and qualify';
  }
  if (/hvac|plumb|electric|roof|landscap|clean|pest|contractor/.test(i)) {
    return 'there may be useful opportunities around quote requests, service scheduling, follow-ups, or job coordination that could reduce manual back-and-forth';
  }
  if (/law|legal|attorney|lawyer/.test(i)) {
    return 'there may be useful opportunities around client intake, document collection, appointment scheduling, or follow-ups that could reduce repetitive back-and-forth';
  }
  if (/manufactur|factory|industrial|wholesale|distribut|logistics|warehouse/.test(i)) {
    return 'there may be useful opportunities around internal requests, approvals, coordination, reporting, or information flow that could reduce repetitive manual work';
  }
  if (/salon|barber|spa|nail|beauty|wellness/.test(i)) {
    return 'there may be useful opportunities around booking, reminders, client follow-ups, or repeat-customer workflows that could make the customer journey smoother';
  }
  if (/auto|mechanic|repair|tire|collision/.test(i)) {
    return 'there may be useful opportunities around service requests, vehicle intake, quoting, scheduling, or follow-ups that could make enquiries easier to manage';
  }
  if (/school|college|education|academy|tuition|university|coaching|training/.test(i)) {
    return 'there may be useful opportunities around enquiries, applications, scheduling, or student/parent communication that could make the next step clearer';
  }
  if (/restaurant|cafe|caf|diner|bakery|catering|food/.test(i)) {
    return 'there may be useful opportunities around enquiries, ordering, reservations, or repeat-customer workflows that could remove friction from the customer journey';
  }
  if (/gym|fitness|yoga|pilates|studio|sports/.test(i)) {
    return 'there may be useful opportunities around membership enquiries, class bookings, reminders, or follow-ups that could make it easier to turn interest into action';
  }

  return 'there may be useful opportunities around enquiries, follow-ups, internal handoffs, or other day-to-day workflows that could remove friction';
}
