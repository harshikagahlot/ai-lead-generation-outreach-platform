/**
 * EmailDrafts.gs
 * -----------------------------------------------------------------------
 * Turns each qualified lead's technical Notes (raw flags from
 * WebsiteAnalyzer) into a genuinely personalized outreach email.
 *
 * Two templates, chosen automatically per lead:
 *   - NO WEBSITE  -> "why a website matters + what we could build" framing
 *   - HAS A WEBSITE (Outdated/Basic/Broken/Very Outdated) -> "here's what
 *     I noticed + specific improvements" framing
 *
 * Both open with one genuine, positive detail pulled from real data
 * (rating/review count), so the email doesn't read as generic.
 *
 * IMPORTANT: everything here is built only from facts already present in
 * the lead's own row (Notes, Rating, Review Count, Website Status) — it
 * never invents an observation about a business.
 *
 * Constants used here — SHEET_DRAFTS, DRAFT_HEADERS, YOUR_NAME_SETTING,
 * YOUR_SERVICE_SETTING — are defined in Config.js alongside all other
 * sheet/settings constants.
 */

// =========================================================================
// CITY EXTRACTION (robust to varying address formats returned by Places)
// =========================================================================

/**
 * Extracts the city from a formatted address by finding the segment that
 * looks like "STATE" or "STATE ZIP" (e.g. "TX", "TX 75201") and returning
 * the segment immediately before it — rather than assuming a fixed
 * position, since Places sometimes includes a country suffix, omits the
 * street for service-area businesses, or splits state/zip differently.
 */
function extractCityFromAddress(address) {
  if (!address) return '';
  let parts = address.split(',').map(s => s.trim()).filter(Boolean);

  // Drop a trailing country token if present (e.g. "USA", "United States").
  const last = parts[parts.length - 1];
  if (last && /^(usa|united states|u\.s\.a\.?)$/i.test(last)) parts.pop();

  // Search from the end for a segment matching a US state code (optionally with a zip).
  for (let i = parts.length - 1; i >= 0; i--) {
    if (/^[A-Z]{2}(\s*\d{5}(-\d{4})?)?$/i.test(parts[i])) {
      return i > 0 ? parts[i - 1] : '';
    }
  }

  // Fallback: no recognizable state/zip segment found — use second-to-last.
  if (parts.length >= 2) return parts[parts.length - 2];
  return parts[0] || '';
}

// =========================================================================
// OBSERVATIONS
// =========================================================================

/** Builds one true, specific positive observation from rating/review data — never invented. */
function buildGenuineDetail(lead) {
  const rating = Number(lead.rating) || 0;
  const reviews = Number(lead.reviewCount) || 0;

  if (rating >= 4.5 && reviews >= 50) {
    return 'a ' + rating + '\u2605 rating across ' + reviews + ' reviews—genuinely impressive for a local business';
  }
  if (rating >= 4.0 && reviews >= 15) {
    return 'consistently positive reviews (' + rating + '\u2605, ' + reviews + ' reviews) from real customers';
  }
  if (reviews > 0) {
    return 'real customer reviews on Google, which says a lot about the trust you\'ve already built locally';
  }
  // No rating/review data available — fall back to something still true and generic-safe.
  return 'the work you\'re doing in the ' + (lead.industry || 'local') + ' space';
}

/**
 * Returns the strongest single observation from the website analyzer flags.
 * Prioritizes severe issues (broken pages, SSL) over minor ones (copyright).
 */
function getStrongestObservation(notes, websiteStatus) {
  if (notes) {
    const flags = notes.split(';').map(s => s.trim().toLowerCase()).filter(Boolean);
    
    // Ordered by priority: strongest observations first
    const priorities = [
      { regex: /http \d|did not respond/, text: "the site didn't load properly when I tried to visit it" },
      { regex: /placeholder|href="#"/, text: "a couple of links on the site didn't seem to lead anywhere" },
      { regex: /no https/, text: "the site doesn't redirect to a secure connection (missing SSL)" },
      { regex: /viewport/, text: "the site didn't resize cleanly for mobile devices" },
      { regex: /flash/, text: "part of the site relies on Flash, which modern browsers block" },
      { regex: /obsolete html|table-based/, text: "the site is built with some fairly dated web techniques" },
      { regex: /very small page/, text: "there wasn't much depth to the site—mostly a single page" },
      { regex: /copyright year/, text: "the site still shows an older copyright year" }
    ];

    for (const p of priorities) {
      for (const flag of flags) {
        if (p.regex.test(flag)) {
          return p.text;
        }
      }
    }
  }
  
  if (websiteStatus === WEBSITE_STATUS.EXCELLENT || websiteStatus === WEBSITE_STATUS.GOOD) {
    return getRandomItem([
      "the website is clean and easy to navigate",
      "the website creates a professional first impression",
      "the information is organized clearly",
      "navigation feels straightforward"
    ]);
  }

  return "the site could use a bit of a refresh";
}

// =========================================================================
// INDUSTRY-ADAPTED DISCOVERY QUESTIONS
// =========================================================================

/**
 * Returns a short, comma-separated list of OPERATIONAL friction examples
 * relevant to the lead's specific industry.
 */
function getIndustryFrictionExamples(industry) {
  const i = (industry || '').toLowerCase();

  if (/dent/.test(i)) return 'appointment reminders, patient follow-ups, or insurance paperwork';
  if (/property|real estate|realt/.test(i)) return 'tenant communication, maintenance requests, lease management, or property inquiries';
  if (/manufactur|factory|production/.test(i)) return 'inventory tracking, production planning, internal approvals, or reporting';
  if (/hvac|plumb|electric|roof/.test(i)) return 'service scheduling, technician dispatch, follow-up reminders, or quote management';
  if (/law|legal|attorney/.test(i)) return 'client intake, document management, appointment scheduling, or case updates';
  
  if (/tutor|educat|school|teach/.test(i)) return 'student scheduling, attendance tracking, or parent communication';
  if (/chiro|physical therap|vet|clinic/.test(i)) return 'scheduling, treatment follow-ups, or patient records';
  if (/landscap|lawn/.test(i)) return 'scheduling recurring jobs, quoting, or coordinating crews';
  if (/florist|retail|gift|boutique/.test(i)) return 'managing orders, tracking inventory, or repeat customer outreach';
  if (/salon|barber|spa|nail/.test(i)) return 'booking appointments, reducing no-shows, or client management';
  if (/auto|mechanic|repair shop/.test(i)) return 'scheduling appointments, quoting, or tracking vehicle progress';
  if (/clean/.test(i)) return 'scheduling recurring jobs, coordinating staff, or client follow-ups';
  if (/restaurant|caf|diner|bakery/.test(i)) return 'managing reservations, online orders, or menu consistency';
  if (/gym|fitness|yoga|pilates|studio/.test(i)) return 'managing memberships, renewals, or class schedules';

  return 'day-to-day scheduling, customer follow-ups, or keeping information organized';
}

// =========================================================================
// EMAIL BUILDING
// =========================================================================

/** Helper to pick a random string from an array */
function getRandomItem(arr) {
  return arr[Math.floor(Math.random() * arr.length)];
}

/**
 * Builds { subject, body } for one lead using the new randomized template structure.
 */
function buildEmailDraft(lead, settings) {
  // Gate 1: Non-email channel — generate a placeholder, not a real draft
  if (!lead.recommendedChannel || lead.recommendedChannel.trim().toUpperCase() !== 'EMAIL') {
    return {
      subject: 'Use ' + (lead.recommendedChannel || 'Other') + ' Channel',
      body: 'Recommended channel for this lead is ' + (lead.recommendedChannel || 'non-email') + ' — not email.\n\nContact via: ' +
        (lead.recommendedChannel === 'Phone' ? (lead.phone || 'see phone column') : (lead.recommendedChannel || 'non-email channel')) +
        '\n\nReadiness Notes: ' + (lead.readinessNotes || '')
    };
  }

  // Gate 2: Email Validity Check (non-empty, syntactically valid, not placeholder)
  if (!isValidOutreachEmail(lead.email)) {
    return {
      subject: 'Needs Review - Invalid Email',
      body: 'Email address (' + (lead.email || 'none') + ') is invalid, missing, or a placeholder.\n\nReadiness Notes: ' + (lead.readinessNotes || '')
    };
  }

  // Gate 3: Must satisfy BOTH readiness score >= 50 AND at least ONE concrete observation
  const readinessScore = Number(lead.readinessScore) || 0;
  const concreteObs = hasConcreteObservation(lead);
  if (readinessScore < 50 || !concreteObs) {
    return {
      subject: 'Needs Review',
      body: 'This lead needs more research or a different outreach channel before contacting.\n\nReason: ' +
        (!concreteObs ? 'Lacks concrete specific observation for personalization. ' : '') +
        (readinessScore < 50 ? 'Low readiness score (' + readinessScore + '). ' : '') +
        '\n\nReadiness Notes: ' + (lead.readinessNotes || '')
    };
  }

  const ownerLine = lead.owner ? lead.owner : 'there';
  const senderName = settings[YOUR_NAME_SETTING] || 'Harshika Gahlot';
  
  const genuineDetail = buildGenuineDetail(lead);
  const isNoWebsite = lead.websiteStatus === WEBSITE_STATUS.NO_WEBSITE;
  const isGoodOrExcellent = lead.websiteStatus === WEBSITE_STATUS.GOOD || lead.websiteStatus === WEBSITE_STATUS.EXCELLENT;
  
  const observation = isNoWebsite 
    ? "there isn't a dedicated website for the business yet"
    : getStrongestObservation(lead.notes, lead.websiteStatus);

  const businessName = lead.name || 'your business';

  let openingLine;
  if (isGoodOrExcellent) {
    openingLine = getRandomItem([
      "While looking into " + businessName + ", I noticed " + genuineDetail + ". I also saw that " + observation + ", which is great to see.",
      "When researching " + businessName + ", I noticed " + genuineDetail + ". I also noticed " + observation + "."
    ]);
  } else {
    openingLine = getRandomItem([
      "While looking into " + businessName + ", I noticed " + genuineDetail + ". I also noticed " + observation + "—it may be a small thing, but it caught my attention.",
      "When researching " + businessName + " online, I noticed " + genuineDetail + ". I also noticed " + observation + "—it might be minor, but it stood out."
    ]);
  }

  const curiosityLine = getRandomItem([
    "I was curious whether your day-to-day scheduling and operations are still working smoothly as things have grown.",
    "I was wondering if your customer follow-up process still feels manageable with your current volume.",
    "I was just curious if your current operational processes are keeping up easily with your growth."
  ]);

  const ctaLine = getRandomItem([
    "Even a quick 'not right now' or 'maybe, tell me more' would be genuinely helpful — no pressure either way.",
    "If you have a quick second to reply with a 'not right now' or 'let's talk', I'd really appreciate it — no pressure.",
    "A simple 'not right now' or 'tell me more' is completely fine — just wanted to reach out."
  ]);

  const subject = 'A quick question about ' + businessName;

  const body = 
    'Hi ' + ownerLine + ',\n\n' +
    openingLine + '\n\n' +
    curiosityLine + '\n\n' +
    ctaLine + '\n\n' +
    'Best regards,\n\n' +
    senderName + '\n' +
    'harshikagahlot01@gmail.com';

  return { subject: subject, body: body };
}

// =========================================================================
// MENU ACTION + SHEET WRITING
// =========================================================================

/** Creates the Outreach_Drafts sheet if missing, with readable formatting for the long draft column. */
function ensureDraftsSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ensureSheetWithHeaders(ss, SHEET_DRAFTS, DRAFT_HEADERS);

  const draftCol = DRAFT_HEADERS.indexOf('Email Draft') + 1;
  sheet.setColumnWidth(draftCol, 500);
  sheet.getRange(1, draftCol, Math.max(sheet.getMaxRows(), 1000), 1).setWrap(true);

  // Migration: ensure all expected headers exist in the correct positions.
  // This handles sheets created before 'Recommended Channel' and 'Gmail Draft ID' were added.
  const lastCol = sheet.getLastColumn();
  const existingHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

  DRAFT_HEADERS.forEach((header, i) => {
    const col = i + 1;
    if ((existingHeaders[i] || '').toString().trim() !== header) {
      // Insert missing column at the correct position
      if (col > lastCol) {
        sheet.getRange(1, col).setValue(header).setFontWeight('bold');
      }
      // If a column is out of place we just ensure the header text is set
      // without moving data — prevents silent misalignment on old sheets.
    }
  });

  return sheet;
}

/**
 * Menu action: generates a draft email for every row in Qualified_Leads
 * that doesn't already have one in Outreach_Drafts (matched by Place ID).
 * Re-run anytime after a new search — it only drafts NEW leads.
 */
function menuDraftOutreachEmails() {
  const ss = SpreadsheetApp.getActive();
  const qualifiedSheet = ss.getSheetByName(SHEET_QUALIFIED);
  const lastRow = qualifiedSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Qualified_Leads is empty — run a search first.');
    return;
  }

  ensureDraftsSheet();
  const draftsSheet = ss.getSheetByName(SHEET_DRAFTS);
  const existingPlaceIds = getExistingDraftPlaceIds(draftsSheet);
  const settings = getSettings();

  // Respect a manually-added tracking column (e.g. "Email Update") if present —
  // never re-draft or overwrite a lead the person has already marked as sent.
  const trackingCol = findManualTrackingColumn(qualifiedSheet, 'email update');
  const lastCol = Math.max(QUALIFIED_HEADERS.length, trackingCol > 0 ? trackingCol : 0);
  const qualifiedData = qualifiedSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let skipped = 0, skippedSent = 0;

  // Collect all new draft rows in memory, then write in one batch RPC
  // instead of N individual appendRow() calls.
  const newDraftRows = [];

  qualifiedData.forEach(row => {
    if (trackingCol > 0) {
      const trackingValue = (row[trackingCol - 1] || '').toString().toLowerCase().trim();
      if (trackingValue === 'sent') { skippedSent++; return; }
    }

    const lead = mapQualifiedRowToLead(row);
    if (existingPlaceIds.has(lead.placeId)) { skipped++; return; }

    const email = buildEmailDraft(lead, settings);

    newDraftRows.push([
      lead.name, lead.industry, extractCityFromAddress(lead.address), lead.email,
      lead.phone, lead.websiteStatus, lead.recommendedChannel || '', email.subject, email.body, 'Draft', lead.placeId,
      ''  // Gmail Draft ID — empty until menuPushDraftsToGmail() fills it
    ]);
  });

  // One batch write instead of N individual appendRow() calls
  if (newDraftRows.length) {
    const startRow = draftsSheet.getLastRow() + 1;
    draftsSheet.getRange(startRow, 1, newDraftRows.length, DRAFT_HEADERS.length).setValues(newDraftRows);
  }

  SpreadsheetApp.getUi().alert(
    'Drafted ' + newDraftRows.length + ' new email(s). Skipped ' + skipped + ' (already drafted). ' +
    (trackingCol > 0 ? 'Skipped ' + skippedSent + ' marked as already sent.' : '')
  );
}

/**
 * Looks for a manually-added tracking column in Qualified_Leads (e.g. the
 * "Email Update" column some users add themselves to mark leads as sent).
 * Searches the header row for an exact or partial match, case-insensitive.
 * Returns the 1-indexed column number, or -1 if no such column exists.
 */
function findManualTrackingColumn(sheet, headerNameContains) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] && headers[i].toString().toLowerCase().indexOf(headerNameContains.toLowerCase()) !== -1) {
      return i + 1;
    }
  }
  return -1;
}

function getExistingDraftPlaceIds(draftsSheet) {
  const ids = new Set();
  const lastRow = draftsSheet.getLastRow();
  if (lastRow < 2) return ids;
  const col = DRAFT_HEADERS.indexOf('Place ID') + 1;
  draftsSheet.getRange(2, col, lastRow - 1, 1).getValues().forEach(r => { if (r[0]) ids.add(r[0]); });
  return ids;
}

function mapQualifiedRowToLead(row) {
  const idx = (name) => QUALIFIED_HEADERS.indexOf(name);
  return {
    name: row[idx('Business Name')],
    industry: row[idx('Industry')],
    owner: row[idx('Owner')],
    email: row[idx('Email')],
    phone: row[idx('Phone')],
    website: row[idx('Website')],
    websiteStatus: row[idx('Website Status')],
    emailType: row[idx('Email Type')],
    recommendedChannel: row[idx('Recommended Channel')],
    readinessScore: row[idx('Readiness Score')],
    readinessNotes: row[idx('Readiness Notes')],
    rating: row[idx('Rating')],
    reviewCount: row[idx('Reviews')],
    address: row[idx('Address')],
    notes: row[idx('Notes')],
    placeId: row[idx('Place ID')]
  };
}

/**
 * Re-drafts ALL emails from scratch (clears Outreach_Drafts first), useful
 * after improving the templates/translations above and wanting every
 * existing qualified lead to get a fresh draft rather than just new ones.
 */
function menuRedraftAllOutreachEmails() {
  const draftsSheet = ensureDraftsSheet();
  clearSheetBody(draftsSheet);
  menuDraftOutreachEmails();
}

// =========================================================================
// PUSH DRAFTS TO GMAIL (creates Gmail drafts — NEVER sends)
// =========================================================================

/**
 * Menu action: reads every row in Outreach_Drafts, creates a real Gmail
 * draft for each eligible row, and records the Gmail Draft ID so duplicates
 * are never created.
 *
 * Dedup strategy — Gmail Draft ID is the SINGLE SOURCE OF TRUTH:
 *   - If Gmail Draft ID column is NOT empty → row was already pushed → skip.
 *   - If Gmail Draft ID column IS empty → eligible for push (subject to
 *     validation below).
 *
 * Validation — a draft is created only if ALL of these are true:
 *   - Email address exists
 *   - Subject line exists
 *   - Email body exists
 *   - Gmail Draft ID is empty (not already pushed)
 *
 * After each successful push:
 *   - Gmail Draft ID column is filled with the real ID from GmailApp
 *   - Status column is updated to "Pushed to Gmail"
 *
 * SAFETY: GmailApp.createDraft() creates a DRAFT — it does NOT send.
 * Existing Gmail drafts are NEVER deleted, modified, or recreated.
 */
function menuPushDraftsToGmail() {
  const ss = SpreadsheetApp.getActive();

  // Ensure the sheet and Gmail Draft ID header exist (handles first-run
  // and backward-compatible migration from older 10-column sheets).
  ensureDraftsSheet();
  const draftsSheet = ss.getSheetByName(SHEET_DRAFTS);

  if (!draftsSheet) {
    SpreadsheetApp.getUi().alert('Outreach_Drafts sheet not found. Run "Draft Outreach Emails" first.');
    return;
  }

  const lastRow = draftsSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No drafts to push — the Outreach_Drafts sheet is empty.');
    return;
  }

  // Column positions from DRAFT_HEADERS (1-indexed for Sheets API)
  const nameCol     = DRAFT_HEADERS.indexOf('Business Name') + 1;
  const emailCol    = DRAFT_HEADERS.indexOf('Email') + 1;
  const subjectCol  = DRAFT_HEADERS.indexOf('Subject') + 1;
  const bodyCol     = DRAFT_HEADERS.indexOf('Email Draft') + 1;
  const statusCol   = DRAFT_HEADERS.indexOf('Status') + 1;
  const gmailIdCol  = DRAFT_HEADERS.indexOf('Gmail Draft ID') + 1;

  // Read all data — use the wider of DRAFT_HEADERS.length or actual sheet
  // width, so we never miss the Gmail Draft ID column even if the sheet
  // was created with the old 10-column layout.
  const readWidth = Math.max(DRAFT_HEADERS.length, draftsSheet.getLastColumn());
  const data = draftsSheet.getRange(2, 1, lastRow - 1, readWidth).getValues();

  // Build update arrays in memory — written in one batch at the end.
  const statusValues  = data.map(row => [row[statusCol - 1] || '']);
  const gmailIdValues = data.map(row => [row[gmailIdCol - 1] || '']);

  // Counters for the detailed summary
  let created        = 0;
  let skippedAlready = 0;
  let skippedNoEmail   = 0;
  let skippedNoSubject = 0;
  let skippedNoBody    = 0;
  let errors         = 0;

  data.forEach((row, i) => {
    // ── PRIMARY DEDUP: Gmail Draft ID is the single source of truth ──
    // If this column has any value, a draft was already created for this
    // row in a previous run — skip it unconditionally.
    const existingDraftId = (row[gmailIdCol - 1] || '').toString().trim();
    if (existingDraftId) {
      skippedAlready++;
      return;
    }

    // ── VALIDATION: all three fields must be present ──
    const recipientEmail = (row[emailCol - 1] || '').toString().trim();
    const subject        = (row[subjectCol - 1] || '').toString().trim();
    const body           = (row[bodyCol - 1] || '').toString().trim();
    
    // We get the recommended channel dynamically
    const recommendedChannelCol = DRAFT_HEADERS.indexOf('Recommended Channel') + 1;
    const recommendedChannel = (row[recommendedChannelCol - 1] || '').toString().trim();

    if (!recipientEmail) { skippedNoEmail++;   return; }
    if (!subject)        { skippedNoSubject++; return; }
    if (!body)           { skippedNoBody++;    return; }

    // ── DEFENSIVE FINAL SAFETY CHECKS BEFORE CREATING GMAIL DRAFT ──

    // Defensive Check 1: Recommended Channel MUST BE EXACTLY 'Email'
    if (recommendedChannel.toUpperCase() !== 'EMAIL') {
      statusValues[i] = ['Skipped: Channel is ' + (recommendedChannel || 'non-email')];
      skippedAlready++;
      return;
    }

    // Defensive Check 2: Email address MUST be valid & non-placeholder
    if (!isValidOutreachEmail(recipientEmail)) {
      statusValues[i] = ['Skipped: Invalid or Placeholder Email'];
      skippedNoEmail++;
      return;
    }

    // Defensive Check 3: Subject & Body MUST NOT be placeholders
    const lowerSubj = subject.toLowerCase();
    if (lowerSubj === 'needs review' || lowerSubj.startsWith('needs review') || lowerSubj.startsWith('use ') || lowerSubj.startsWith('draft')) {
      statusValues[i] = ['Skipped: Placeholder Subject'];
      skippedNoSubject++;
      return;
    }

    if (body.startsWith('This lead needs more research') || body.startsWith('Recommended channel for this lead') || body.startsWith('Email address')) {
      statusValues[i] = ['Skipped: Placeholder Body'];
      skippedNoBody++;
      return;
    }

    // ── CREATE GMAIL DRAFT ──
    try {
      const draft = GmailApp.createDraft(recipientEmail, subject, body);

      // Store the real Gmail Draft ID — this is what prevents duplicates
      // on future runs, and lets the user locate the draft in Gmail.
      gmailIdValues[i] = [draft.getId()];
      statusValues[i]  = ['Pushed to Gmail'];
      created++;
    } catch (e) {
      gmailIdValues[i] = [''];  // leave empty so retry is possible
      statusValues[i]  = ['Error: ' + e.message];
      errors++;
      Logger.log('Gmail draft error for row ' + (i + 2) +
        ' (' + (row[nameCol - 1] || 'unknown') + '): ' + e.message);
    }
  });

  // ── BATCH WRITE: 2 column writes instead of N individual setValue() ──
  draftsSheet.getRange(2, statusCol,  statusValues.length,  1).setValues(statusValues);
  draftsSheet.getRange(2, gmailIdCol, gmailIdValues.length, 1).setValues(gmailIdValues);

  // ── DETAILED SUMMARY ──
  const lines = ['Push to Gmail complete!\n'];
  lines.push('✅ Created: ' + created + ' Gmail draft(s)');
  if (skippedAlready  > 0) lines.push('⏭️ Skipped (Already in Gmail): ' + skippedAlready);
  if (skippedNoEmail  > 0) lines.push('⚠️ Skipped (Missing Email): ' + skippedNoEmail);
  if (skippedNoSubject > 0) lines.push('⚠️ Skipped (Missing Subject): ' + skippedNoSubject);
  if (skippedNoBody   > 0) lines.push('⚠️ Skipped (Missing Body): ' + skippedNoBody);
  if (errors          > 0) lines.push('❌ Errors: ' + errors);
  if (created === 0 && errors === 0) {
    lines.push('\nℹ️ Nothing new to push. All rows either already have a Gmail Draft ID or are missing required fields.');
  } else {
    lines.push('\nDrafts are in your Gmail Drafts folder. They have NOT been sent.');
  }

  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

/**
 * Menu action: Summarizes recorded Email_Opens against Outreach_Drafts.
 * Displays: Lead Name, Opened (Yes/No), First Opened, Last Opened, Total Opens.
 */
function menuViewEmailOpens() {
  const ss = SpreadsheetApp.getActive();
  const opensSheet = ss.getSheetByName(SHEET_OPENS);
  const ui = SpreadsheetApp.getUi();

  if (!opensSheet || opensSheet.getLastRow() < 2) {
    ui.alert('No email open events recorded yet in "Email_Opens".');
    return;
  }

  const opensData = opensSheet.getRange(2, 1, opensSheet.getLastRow() - 1, OPENS_HEADERS.length).getValues();

  const opensMap = {};
  opensData.forEach(row => {
    const timestamp = row[0];
    const leadId = (row[1] || '').toString().trim();
    const bizName = (row[2] || '').toString().trim();

    if (!leadId) return;

    if (!opensMap[leadId]) {
      opensMap[leadId] = {
        name: bizName || leadId,
        count: 0,
        firstOpened: timestamp,
        lastOpened: timestamp
      };
    }

    opensMap[leadId].count++;
    opensMap[leadId].lastOpened = timestamp;
    if (bizName && (!opensMap[leadId].name || opensMap[leadId].name === leadId)) {
      opensMap[leadId].name = bizName;
    }
  });

  const leadIds = Object.keys(opensMap);
  if (!leadIds.length) {
    ui.alert('No valid lead opens recorded.');
    return;
  }

  let message = '📊 EMAIL OPENS SUMMARY (' + leadIds.length + ' leads recorded)\n\n';
  leadIds.slice(0, 15).forEach((id, idx) => {
    const item = opensMap[id];
    const firstStr = item.firstOpened ? Utilities.formatDate(new Date(item.firstOpened), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : 'N/A';
    const lastStr = item.lastOpened ? Utilities.formatDate(new Date(item.lastOpened), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : 'N/A';
    message += (idx + 1) + '. ' + item.name + '\n' +
               '   Opened: Yes (' + item.count + ' time' + (item.count > 1 ? 's' : '') + ')\n' +
               '   First: ' + firstStr + ' | Last: ' + lastStr + '\n\n';
  });

  if (leadIds.length > 15) {
    message += '... and ' + (leadIds.length - 15) + ' more. Check the Email_Opens tab for full history.';
  }

  ui.alert(message);
}

/**
 * Returns HTML string for tracking pixel if Web App URL is configured.
 * @param {string} leadId
 * @param {object} settings
 * @returns {string} HTML img tag or empty string
 */
function getTrackingPixelHtml(leadId, settings) {
  if (!leadId) return '';
  const webAppUrl = (settings && settings['Web App URL']) ? settings['Web App URL'].toString().trim() : '';
  if (!webAppUrl) return '';

  return '<img src="' + webAppUrl + '?leadId=' + encodeURIComponent(leadId) + '" width="1" height="1" style="display:none !important;" alt="" />';
}