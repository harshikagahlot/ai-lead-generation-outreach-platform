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
    .addSeparator()
    .addItem('10. View Email Opens Summary', 'menuViewEmailOpens')
    .addToUi();
}

// =========================================================================
// 1. GENERATE LEADS — the main pipeline
// =========================================================================

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
    hasRecentReview: false,
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
    owner: '',
    notes: analysis.flags ? analysis.flags.join('; ') : ''
  };
}

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
    if (pageToken && results.length < maxResults) Utilities.sleep(2000);
  } while (pageToken && results.length < maxResults);
  return results.slice(0, maxResults);
}

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
  const statusValues = [], scoreValues = [], notesValues = [];
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
  sheet.getRange(2, statusCol, statusValues.length, 1).setValues(statusValues);
  sheet.getRange(2, scoreCol, scoreValues.length, 1).setValues(scoreValues);
  sheet.getRange(2, notesCol, notesValues.length, 1).setValues(notesValues);
  SpreadsheetApp.getUi().alert('Re-analyzed ' + data.length + ' website(s). Notes column refreshed too.');
}

function menuReFindEmails() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_RAW);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert('Raw_Data is empty.'); return; }
  const websiteCol  = RAW_HEADERS.indexOf('Website') + 1;
  const emailCol    = RAW_HEADERS.indexOf('Email') + 1;
  const emailSrcCol = RAW_HEADERS.indexOf('Email Source URL') + 1;
  const data = sheet.getRange(2, 1, lastRow - 1, RAW_HEADERS.length).getValues();
  let found = 0, cleared = 0;
  const emailValues = data.map(row => [row[emailCol - 1]]);
  const emailSrcValues = data.map(row => [row[emailSrcCol - 1]]);
  data.forEach((row, i) => {
    const currentEmail = row[emailCol - 1];
    const needsRecheck = !currentEmail || isPlaceholderEmail(currentEmail);
    if (!needsRecheck) return;
    const website = row[websiteCol - 1];
    if (!website) {
      if (currentEmail) {
        emailValues[i] = [''];
        emailSrcValues[i] = [''];
        cleared++;
      }
      return;
    }
    let homepageHtml = '';
    try {
      const homepageResponse = fetchWithRetry(normalizeUrl(website), { followRedirects: true, validateHttpsCertificates: false }, 1);
      if (homepageResponse.getResponseCode() < 400) homepageHtml = homepageResponse.getContentText();
    } catch (e) {}
    const result = findPublicEmail(website, homepageHtml);
    if (result.email) {
      emailValues[i] = [result.email];
      emailSrcValues[i] = [result.sourceUrl];
      found++;
    } else if (currentEmail) {
      emailValues[i] = [''];
      emailSrcValues[i] = [''];
      cleared++;
    }
  });
  sheet.getRange(2, emailCol, emailValues.length, 1).setValues(emailValues);
  sheet.getRange(2, emailSrcCol, emailSrcValues.length, 1).setValues(emailSrcValues);
  SpreadsheetApp.getUi().alert('Found ' + found + ' real email(s). Cleared ' + cleared + ' placeholder email(s) with no real replacement found.');
}

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
  const qualifiedRows = [];
  const rejectedRows  = [];
  data.forEach(row => {
    const lead = mapRawRowToLead(row);
    const qualification = evaluateQualification(lead.email, lead.websiteStatus);
    if (qualification.qualified) {
      qualifiedRows.push([
        lead.name, lead.industry, lead.owner || '', lead.email, lead.phone,
        lead.website || '', lead.websiteStatus, lead.emailType || '',
        lead.recommendedChannel || '', lead.readinessScore || '', lead.readinessNotes || '',
        lead.score, lead.rating || '', lead.reviewCount || '', lead.address,
        lead.mapsUrl || '', lead.notes || '', lead.placeId
      ]);
    } else {
      rejectedRows.push([lead.name, qualification.reason, lead.placeId || '', now()]);
    }
  });
  if (qualifiedRows.length) qualifiedSheet.getRange(2, 1, qualifiedRows.length, QUALIFIED_HEADERS.length).setValues(qualifiedRows);
  if (rejectedRows.length) rejectedSheet.getRange(2, 1, rejectedRows.length, REJECTED_HEADERS.length).setValues(rejectedRows);
  SpreadsheetApp.getUi().alert('Rebuilt: ' + qualifiedRows.length + ' qualified, ' + rejectedRows.length + ' rejected.');
}

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

function menuExportCsv() {
  const file = exportQualifiedLeadsToCsv();
  SpreadsheetApp.getUi().alert('Exported to Google Drive: ' + file.getName() + '\n\n' + file.getUrl());
}

function doGet(e) {
  try {
    const leadId = (e && e.parameter && e.parameter.leadId) ? e.parameter.leadId.trim() : '';
    const userAgent = (e && e.parameter && e.parameter.userAgent) || '';
    if (leadId) logEmailOpen(leadId, userAgent);
  } catch (err) {
    Logger.log('doGet open tracking error: ' + err.message);
  }
  const transparentSvg = '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"/>';
  return ContentService.createTextOutput(transparentSvg).setMimeType(ContentService.Mime.XML);
}
