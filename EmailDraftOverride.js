// Email draft override: make the agreed VASHA outreach structure the default.
// This file intentionally leaves the rest of EmailDrafts.js untouched.
// The assignment below replaces the global builder used by menuDraftOutreachEmails().
function buildEmailDraftV2(lead, settings) {
  // Gate 1: Non-email channel — keep as an internal placeholder.
  if (lead.recommendedChannel && lead.recommendedChannel !== 'Email') {
    return {
      subject: 'Use ' + lead.recommendedChannel + ' Channel',
      body: 'Recommended channel for this lead is ' + lead.recommendedChannel + ' — not email.\n\nContact via: ' +
        (lead.recommendedChannel === 'Phone' ? (lead.phone || 'see phone column') : lead.recommendedChannel) +
        '\n\nReadiness Notes: ' + (lead.readinessNotes || '')
    };
  }

  // Gate 2: Low readiness score — needs more research.
  if ((Number(lead.readinessScore) || 0) < 50) {
    return {
      subject: 'Needs Review',
      body: 'This lead needs more research or a different outreach channel before contacting.\n\nReadiness Notes: ' + (lead.readinessNotes || '')
    };
  }

  const ownerLine = lead.owner ? lead.owner : 'there';
  const senderName = settings[YOUR_NAME_SETTING] || 'Harshika Gahlot';
  const businessName = lead.name || 'your business';
  const genuineDetail = buildGenuineDetail(lead);
  const isNoWebsite = lead.websiteStatus === WEBSITE_STATUS.NO_WEBSITE;
  const isGoodOrExcellent = lead.websiteStatus === WEBSITE_STATUS.GOOD || lead.websiteStatus === WEBSITE_STATUS.EXCELLENT;
  const observation = isNoWebsite
    ? "there isn't a dedicated website for the business yet"
    : getStrongestObservation(lead.notes, lead.websiteStatus);

  // Industry is used to suggest possibilities, never to claim a problem exists.
  const opportunity = getIndustryOpportunity(lead.industry);

  let openingLine;
  if (isNoWebsite) {
    openingLine =
      'While looking into ' + businessName + ', I noticed ' + genuineDetail + '. I also noticed that ' + observation + '.';
  } else if (isGoodOrExcellent) {
    openingLine =
      'While looking into ' + businessName + ', I noticed ' + genuineDetail + '. I also noticed ' + observation + '.';
  } else {
    openingLine =
      'While looking into ' + businessName + ', I noticed ' + genuineDetail + '. I also noticed ' + observation + '. It may be minor, but it caught my attention.';
  }

  const systemsLine =
    'A lot of businesses do not necessarily have a marketing problem. Sometimes the bigger opportunity is somewhere around the system behind the customer journey — how enquiries, follow-ups, operations, or internal handoffs work.';

  const nonAssumptionLine =
    'I cannot tell from the outside whether any of that is actually a priority for ' + businessName + ', so I do not want to assume there is a problem.';

  const opportunityLine =
    'Depending on how you currently operate, there may be useful opportunities around ' + opportunity + '.';

  const vashaLine =
    'I am currently building VASHA Technologies as an early-stage initiative focused on helping businesses explore and improve these kinds of workflows through AI automation, custom software, and practical business systems.';

  const ctaLine =
    'If it is relevant, I can send over 2–3 concrete ideas based on what I noticed — no pitch deck, just a short, useful message. If it is not relevant right now, that is completely fine too.';

  const subject = getRandomItem([
    'A quick idea for ' + businessName,
    'A thought on ' + businessName,
    'A quick thought for ' + businessName
  ]);

  const body =
    'Hi ' + ownerLine + ',\n\n' +
    openingLine + '\n\n' +
    systemsLine + ' ' + nonAssumptionLine + '\n\n' +
    opportunityLine + '\n\n' +
    vashaLine + '\n\n' +
    ctaLine + '\n\n' +
    'Best regards,\n' +
    senderName + '\n' +
    'Business Development | VASHA Technologies\n' +
    'AI Automation • Custom Software • Business Systems\n' +
    '📧 harshikagahlot01@gmail.com';

  return { subject: subject, body: body };
}

function getIndustryOpportunity(industry) {
  const i = (industry || '').toLowerCase();

  if (/dent/.test(i)) return 'appointment enquiries, reminders, patient intake, follow-ups, or reducing admin around recurring visits';
  if (/property|real estate|realt/.test(i)) return 'property enquiries, lead qualification, owner/tenant communication, maintenance workflows, or follow-ups';
  if (/manufactur|factory|production/.test(i)) return 'quoting, order visibility, inventory coordination, approvals, reporting, or internal workflow automation';
  if (/hvac|plumb|electric|roof/.test(i)) return 'service enquiries, scheduling, quote follow-ups, technician coordination, or customer updates';
  if (/law|legal|attorney/.test(i)) return 'client intake, appointment scheduling, document workflows, case updates, or follow-up';
  if (/tutor|educat|school|teach/.test(i)) return 'enquiries, student onboarding, scheduling, attendance, parent communication, or follow-up';
  if (/chiro|physical therap|vet|clinic/.test(i)) return 'appointments, intake, reminders, follow-ups, records, or patient communication';
  if (/landscap|lawn/.test(i)) return 'quoting, recurring-job scheduling, crew coordination, customer updates, or follow-ups';
  if (/florist|retail|gift|boutique/.test(i)) return 'enquiries, orders, inventory visibility, repeat-customer outreach, or customer follow-up';
  if (/salon|barber|spa|nail/.test(i)) return 'booking, reminders, no-show reduction, repeat visits, customer communication, or follow-up';
  if (/auto|mechanic|repair shop/.test(i)) return 'service enquiries, appointment scheduling, estimates, repair-status updates, or follow-ups';
  if (/clean/.test(i)) return 'quoting, recurring-job scheduling, staff coordination, customer updates, or follow-ups';
  if (/restaurant|caf|diner|bakery/.test(i)) return 'reservations, enquiries, online ordering, customer communication, repeat visits, or internal coordination';
  if (/gym|fitness|yoga|pilates|studio/.test(i)) return 'membership enquiries, class bookings, reminders, renewals, member communication, or follow-up';
  if (/marketing|advertis|agency/.test(i)) return 'lead intake, client onboarding, follow-up, reporting, internal handoffs, or workflow automation';
  if (/software|saas|technology|it/.test(i)) return 'lead qualification, onboarding, support workflows, reporting, internal handoffs, or repetitive process automation';

  return 'customer enquiries, follow-ups, scheduling, internal handoffs, reporting, or other repetitive workflows';
}

// Replace the original global builder for every execution.
buildEmailDraft = buildEmailDraftV2;
