/**
 * Code.gs
 * -----------------------------------------------------------------------
 * Main entry point for the Lead Generation system.
 *
 * Responsibilities:
 *   - Custom spreadsheet menu (onOpen)
 *   - User prompt flow for starting a new lead generation job
 *   - Lead orchestration: ties WebsiteAnalyzer + EmailFinder + LeadScoring
 *     into buildLeadFromPlace()
 *   - Google Places API (New) wrappers: searchPlacesText(), getPlaceDetails()
 *   - Menu actions 3–6: re-analyze websites, re-find emails, rebuild
 *     qualified/rejected lists, export CSV
 *
 * Helper functions used here but defined elsewhere:
 *   - clearSheetBody()        → Sheets.js
 *   - analyzeWebsite()        → WebsiteAnalyzer.js
 *   - findPublicEmail()       → EmailFinder.js
 *   - computeLeadScore()      → LeadScoring.js
 *   - evaluateQualification() → LeadScoring.js
 *   - startNewLeadJob()       → BatchProcessor.js
 *
 * PLACES API (New) docs:
 * https://developers.google.com/maps/documentation/places/web-service/text-search
 * https://developers.google.com/maps/documentation/places/web-service/place-details
 */

/** Adds the custom menu to the spreadsheet UI when the document is opened. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎯 Lead Gen Automation')
    .addItem('1. Initialize Workbook (run once)', 'initializeWorkbook')
    .addSeparator()
    .addItem('2. Generate Leads (start new job)', 'promptGenerateLeads')
    .addItem('2b. Continue Lead Generation Job', 'menuContinueLeadJob')
    .addItem('2c. Job Status', 'menuJobStatus')
    .addItem('2d. Cancel Current Job', 'menuCancelLeadJob')
    .addSeparator()
    .addItem('3. Analyze Websites (re-run on Raw_Data)', 'menuReanalyzeWebsites')
    .addItem('4. Find Emails (re-run on Raw_Data)', 'menuReFindEmails')
    .addItem('5. Filter Qualified Leads (rebuild from Raw_Data)', 'menuFilterQualifiedLeads')
    .addItem('6. Export Qualified Leads to CSV', 'menuExportCsv')
    .addItem('7. Draft Outreach Emails (new leads only)', 'menuDraftOutreachEmails')
    .addItem('8. Re-draft ALL Outreach Emails (fresh, overwrites)', 'menuRedraftAllOutreachEmails')
    .addItem('9. Push Drafts to Gmail (creates Gmail drafts only — never sends)', 'menuPushDraftsToGmail')
    .addToUi();
}

// =========================================================================
// 1. GENERATE LEADS — the main pipeline
// =========================================================================

/**
 * Prompts the user for search parameters (industry, city, state, max businesses)
 * and kicks off a new resumable lead generation job via startNewLeadJob().
 */
function promptGenerateLeads() {
  const ui = SpreadsheetApp.getUi();

  const industryResp = ui.prompt('Industry', 'e.g. "HVAC", "dentist", "hair salon"', ui.ButtonSet.OK_CANCEL);
  if (industryResp.getSelectedButton() !== ui.Button.OK) return;

  const cityResp = ui.prompt('City', 'e.g. "Salina"', ui.ButtonSet.OK_CANCEL);
  if (cityResp.getSelectedButton() !== ui.Button.OK) return;

  const stateResp = ui.prompt('State', 'e.g. "Kansas"', ui.ButtonSet.OK_CANCEL);
  if (stateResp.getSelectedButton() !== ui.Button.OK) return;

  const settings = getSettings();
  const defaultMax = settings['Default Max Businesses'] || 20;
  const maxResp = ui.prompt('Max businesses to check', 'Default: ' + defaultMax, ui.ButtonSet.OK_CANCEL);
  if (maxResp.getSelectedButton() !== ui.Button.OK) return;

  const industry = industryResp.getResponseText().trim();
  const city = cityResp.getResponseText().trim();
  const state = stateResp.getResponseText().trim();
  const maxBusinesses = parseInt(maxResp.getResponseText().trim(), 10) || defaultMax;

  if (!industry || !city || !state) return;

  startNewLeadJob(industry, city, state, maxBusinesses);
}

/**
 * Builds the full lead object for one Place: fetches website analysis,
 * discovers email, computes score, and assembles all fields.
 * Called once per lead by processBatch() in BatchProcessor.js.
 *
 * @param {Object} details - Place Details response from Google Places API (New)
 * @param {string} industry - the industry string from the user's search query
 * @returns {Object} lead object with all fields needed for Raw_Data row
 */
function buildLeadFromPlace(details, industry) {
  const websiteUri = details.websiteUri || '';
  const analysis = analyzeWebsite(websiteUri);
  const emailResult = findPublicEmail(websiteUri, analysis.html);
  const emailType = emailResult.type || '';
  const recommendedChannel = emailType === 'Named Person' ? 'Email' : (details.nationalPhoneNumber ? 'Phone' : 'None');

  const rawNotes = analysis.flags ? analysis.flags.join('; ') : '';
  const tempLead = {
    websiteStatus: analysis.status,
    rating: details.rating,
    reviewCount: details.userRatingCount,
    notes: rawNotes
  };

  const readiness = computeOutreachReadiness(tempLead, emailType, recommendedChannel);

  const score = computeLeadScore({
    websiteStatus: analysis.status,
    hasEmail: !!emailResult.email,
    rating: details.rating,
    reviewCount: details.userRatingCount,
    hasRecentReview: false, // Places (New) review timestamps could be wired in here later
    emailType: emailType,
    readinessScore: readiness.score
  });

  return {
    name: details.displayName ? details.displayName.text : '',
    industry: industry,
    address: details.formattedAddress || '',
    phone: details.nationalPhoneNumber || '',
    website: websiteUri,
    websiteStatus: analysis.status,
    email: emailResult.email,
    emailSourceUrl: emailResult.sourceUrl,
    emailType: emailType,
    recommendedChannel: recommendedChannel,
    readinessScore: readiness.score,
    readinessNotes: readiness.notes,
    score: score,
    rating: details.rating || '',
    reviewCount: details.userRatingCount || '',
    mapsUrl: details.googleMapsUri || '',
    lat: details.location ? details.location.latitude : '',
    lng: details.location ? details.location.longitude : '',
    placeId: details.id,
    businessStatus: details.businessStatus || '',
    owner: '', // Places API does not expose an owner name; left blank intentionally
    notes: analysis.flags ? analysis.flags.join('; ') : ''
  };
}

// =========================================================================
// PLACES API (New) WRAPPERS
// =========================================================================

/**
 * Searches the Google Places API (New) Text Search endpoint.
 * Handles pagination via nextPageToken to collect up to maxResults places.
 *
 * @param {string} industry - business type to search for (e.g. "HVAC")
 * @param {string} location - city + state string (e.g. "Salina, Kansas")
 * @param {number} maxResults - maximum number of places to return
 * @returns {Object[]} array of place objects with at least { id, displayName }
 */
function searchPlacesText(industry, location, maxResults) {
  const apiKey = getPlacesApiKey();
  const url = 'https://places.googleapis.com/v1/places:searchText';
  const results = [];
  let pageToken = null;

  do {
    const body = {
      textQuery: industry + ' in ' + location,
      maxResultCount: Math.min(20, maxResults - results.length)
    };
    if (pageToken) body.pageToken = pageToken;

    const response = fetchWithRetry(url, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'X-Goog-Api-Key': apiKey,
        'X-Goog-FieldMask': 'places.id,places.displayName,nextPageToken'
      },
      payload: JSON.stringify(body)
    });

    const json = JSON.parse(response.getContentText());
    if (json.error) throw new Error('Places API error: ' + json.error.message);

    (json.places || []).forEach(p => results.push(p));
    pageToken = json.nextPageToken || null;

    if (pageToken && results.length < maxResults) Utilities.sleep(2000); // token activation delay

  } while (pageToken && results.length < maxResults);

  return results.slice(0, maxResults);
}

/**
 * Fetches full details for a single place by its Place ID.
 * Requests only the fields needed for lead generation.
 *
 * @param {string} placeId - Google Places place ID
 * @returns {Object|null} place details object, or null if the request failed
 */
function getPlaceDetails(placeId) {
  const apiKey = getPlacesApiKey();
  const fields = [
    'id', 'displayName', 'formattedAddress', 'websiteUri',
    'nationalPhoneNumber', 'rating', 'userRatingCount', 'googleMapsUri',
    'businessStatus', 'location'
  ].join(',');

  const url = 'https://places.googleapis.com/v1/places/' + placeId;
  const response = fetchWithRetry(url, {
    method: 'get',
    headers: { 'X-Goog-Api-Key': apiKey, 'X-Goog-FieldMask': fields }
  });

  const json = JSON.parse(response.getContentText());
  if (json.error) {
    Logger.log('Place Details error for ' + placeId + ': ' + json.error.message);
    return null;
  }
  return json;
}

// =========================================================================
// MENU ACTIONS 3-6: operate on existing Raw_Data rows
// =========================================================================

/** Re-runs website analysis for every row in Raw_Data and updates status/score. */
function menuReanalyzeWebsites() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_RAW);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('Raw_Data is empty.'); return; }

  const websiteCol = RAW_HEADERS.indexOf('Website') + 1;
  const statusCol = RAW_HEADERS.indexOf('Website Status') + 1;
  const scoreCol  = RAW_HEADERS.indexOf('Lead Score') + 1;
  const emailCol  = RAW_HEADERS.indexOf('Email') + 1;
  const ratingCol = RAW_HEADERS.indexOf('Rating') + 1;
  const reviewCol = RAW_HEADERS.indexOf('Review Count') + 1;
  const notesCol  = RAW_HEADERS.indexOf('Notes') + 1;

  const data = sheet.getRange(2, 1, lastRow - 1, RAW_HEADERS.length).getValues();

  // Build column arrays in memory, then write each column in one batch RPC
  // instead of 3 × N individual setValue() calls.
  const statusValues = [];
  const scoreValues  = [];
  const notesValues  = [];

  data.forEach((row) => {
    const website = row[websiteCol - 1];
    const analysis = analyzeWebsite(website);
    statusValues.push([analysis.status]);
    notesValues.push([analysis.flags ? analysis.flags.join('; ') : '']);

    const hasEmail = !!row[emailCol - 1];
    const score = computeLeadScore({
      websiteStatus: analysis.status,
      hasEmail: hasEmail,
      rating: row[ratingCol - 1],
      reviewCount: row[reviewCol - 1]
    });
    scoreValues.push([score]);
  });

  // 3 batch writes instead of 3N individual setValue() calls
  sheet.getRange(2, statusCol, statusValues.length, 1).setValues(statusValues);
  sheet.getRange(2, scoreCol,  scoreValues.length,  1).setValues(scoreValues);
  sheet.getRange(2, notesCol,  notesValues.length,  1).setValues(notesValues);

  SpreadsheetApp.getUi().alert('Re-analyzed ' + data.length + ' website(s). Notes column refreshed too.');
}

/** Re-runs email discovery for every row in Raw_Data missing an email OR currently holding a known placeholder email. */
function menuReFindEmails() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_RAW);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('Raw_Data is empty.'); return; }

  const websiteCol  = RAW_HEADERS.indexOf('Website') + 1;
  const emailCol    = RAW_HEADERS.indexOf('Email') + 1;
  const emailSrcCol = RAW_HEADERS.indexOf('Email Source URL') + 1;

  const data = sheet.getRange(2, 1, lastRow - 1, RAW_HEADERS.length).getValues();
  let found = 0, cleared = 0;

  // Start with current values; only modify rows that need re-checking.
  // Write both columns back in one batch at the end instead of 2N individual setValue() calls.
  const emailValues    = data.map(row => [row[emailCol - 1]]);
  const emailSrcValues = data.map(row => [row[emailSrcCol - 1]]);

  data.forEach((row, i) => {
    const currentEmail = row[emailCol - 1];
    const needsRecheck = !currentEmail || isPlaceholderEmail(currentEmail);
    if (!needsRecheck) return; // already has a real, confirmed-good email

    const website = row[websiteCol - 1];
    if (!website) {
      if (currentEmail) { // had a placeholder but no site to re-check — just clear it
        emailValues[i]    = [''];
        emailSrcValues[i] = [''];
        cleared++;
      }
      return;
    }

    let homepageHtml = '';
    try {
      const homepageResponse = fetchWithRetry(normalizeUrl(website), { followRedirects: true, validateHttpsCertificates: false }, 1);
      if (homepageResponse.getResponseCode() < 400) homepageHtml = homepageResponse.getContentText();
    } catch (e) {
      // homepage fetch failed — findPublicEmail will still try the sub-pages below
    }

    const result = findPublicEmail(website, homepageHtml);
    if (result.email) {
      emailValues[i]    = [result.email];
      emailSrcValues[i] = [result.sourceUrl];
      found++;
    } else if (currentEmail) {
      // Had a placeholder, and no real email could be found to replace it — clear rather than keep the fake one.
      emailValues[i]    = [''];
      emailSrcValues[i] = [''];
      cleared++;
    }
  });

  // 2 batch writes instead of up to 2N individual setValue() calls
  sheet.getRange(2, emailCol,    emailValues.length,    1).setValues(emailValues);
  sheet.getRange(2, emailSrcCol, emailSrcValues.length, 1).setValues(emailSrcValues);

  SpreadsheetApp.getUi().alert('Found ' + found + ' real email(s). Cleared ' + cleared + ' placeholder email(s) with no real replacement found.');
}

/**
 * Rebuilds Qualified_Leads and Rejected_Leads from scratch based on current
 * Raw_Data. Uses batch setValues() instead of per-row appendRow() calls.
 */
function menuFilterQualifiedLeads() {
  const ss = SpreadsheetApp.getActive();
  const rawSheet = ss.getSheetByName(SHEET_RAW);
  const lastRow = rawSheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('Raw_Data is empty.'); return; }

  const qualifiedSheet = ss.getSheetByName(SHEET_QUALIFIED);
  const rejectedSheet  = ss.getSheetByName(SHEET_REJECTED);
  clearSheetBody(qualifiedSheet);
  clearSheetBody(rejectedSheet);

  const data = rawSheet.getRange(2, 1, lastRow - 1, RAW_HEADERS.length).getValues();

  // Collect all rows in memory, then write each sheet in a single batch RPC
  // instead of N individual appendRow() calls.
  const qualifiedRows = [];
  const rejectedRows  = [];

  data.forEach(row => {
    const lead = mapRawRowToLead(row);
    const qualification = evaluateQualification(!!lead.email, lead.websiteStatus);
    if (qualification.qualified) {
      qualifiedRows.push([
        lead.name, lead.industry, lead.owner || '', lead.email, lead.phone,
        lead.website || '', lead.websiteStatus, lead.emailType || '',
        lead.recommendedChannel || '', lead.readinessScore || '', lead.readinessNotes || '',
        lead.score, lead.rating || '',
        lead.reviewCount || '', lead.address, lead.mapsUrl || '', lead.notes || '',
        lead.placeId
      ]);
    } else {
      rejectedRows.push([lead.name, qualification.reason, lead.placeId || '', now()]);
    }
  });

  // 2 batch writes instead of N individual appendRow() calls
  if (qualifiedRows.length) {
    qualifiedSheet.getRange(2, 1, qualifiedRows.length, QUALIFIED_HEADERS.length).setValues(qualifiedRows);
  }
  if (rejectedRows.length) {
    rejectedSheet.getRange(2, 1, rejectedRows.length, REJECTED_HEADERS.length).setValues(rejectedRows);
  }

  SpreadsheetApp.getUi().alert('Rebuilt: ' + qualifiedRows.length + ' qualified, ' + rejectedRows.length + ' rejected.');
}

/**
 * Maps a raw row array (from Raw_Data sheet) into a lead object.
 * Uses RAW_HEADERS to find column positions dynamically.
 * @param {Array} row - one row of values from Raw_Data
 * @returns {Object} lead object compatible with appendQualifiedRow/appendRejectedRow
 */
function mapRawRowToLead(row) {
  const idx = (name) => RAW_HEADERS.indexOf(name);
  
  const lead = {
    name: row[idx('Business Name')],
    industry: row[idx('Industry')],
    address: row[idx('Address')],
    phone: row[idx('Phone')],
    website: row[idx('Website')],
    websiteStatus: row[idx('Website Status')],
    email: row[idx('Email')],
    rating: row[idx('Rating')],
    reviewCount: row[idx('Review Count')],
    mapsUrl: row[idx('Google Maps URL')],
    placeId: row[idx('Place ID')],
    owner: '',
    notes: row[idx('Notes')] || ''
  };

  lead.emailType = classifyEmail(lead.email);
  lead.recommendedChannel = lead.emailType === 'Named Person' ? 'Email' : (lead.phone ? 'Phone' : 'None');
  
  const readiness = computeOutreachReadiness(lead, lead.emailType, lead.recommendedChannel);
  lead.readinessScore = readiness.score;
  lead.readinessNotes = readiness.notes;

  lead.score = computeLeadScore({
    websiteStatus: lead.websiteStatus,
    hasEmail: !!lead.email,
    rating: lead.rating,
    reviewCount: lead.reviewCount,
    hasRecentReview: false,
    emailType: lead.emailType,
    readinessScore: lead.readinessScore
  });

  return lead;
}

/** Menu action: exports Qualified_Leads to a CSV file in Google Drive. */
function menuExportCsv() {
  const file = exportQualifiedLeadsToCsv();
  SpreadsheetApp.getUi().alert('Exported to Google Drive: ' + file.getName() + '\n\n' + file.getUrl());
}