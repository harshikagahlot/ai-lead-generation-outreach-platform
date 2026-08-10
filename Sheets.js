/**
 * Sheets.gs
 * -----------------------------------------------------------------------
 * Everything related to creating, reading, and writing to the workbook's
 * sheets: Raw_Data, Qualified_Leads, Rejected_Leads, Logs, Outreach_Drafts,
 * and Settings.
 *
 * Also handles:
 *   - Workbook initialization (creating all sheets with headers)
 *   - Duplicate detection (Place ID + name/phone/address fingerprinting)
 *   - Row writers for all data sheets
 *   - CSV export of qualified leads to Google Drive
 */

/** One-time (or repeatable) setup: creates all sheets with headers if missing. */
function initializeWorkbook() {
  const ss = SpreadsheetApp.getActive();

  ensureSheetWithHeaders(ss, SHEET_RAW, RAW_HEADERS);
  ensureSheetWithHeaders(ss, SHEET_QUALIFIED, QUALIFIED_HEADERS);
  ensureSheetWithHeaders(ss, SHEET_REJECTED, REJECTED_HEADERS);
  ensureSheetWithHeaders(ss, SHEET_LOGS, LOG_HEADERS);
  ensureSheetWithHeaders(ss, SHEET_DRAFTS, DRAFT_HEADERS);
  ensureSheetWithHeaders(ss, SHEET_OPENS, OPENS_HEADERS);

  // Settings sheet is special: key/value pairs, seeded with defaults.
  let settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_SETTINGS);
    settingsSheet.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]).setFontWeight('bold');
    const rows = Object.keys(DEFAULT_SETTINGS).map(k => [k, DEFAULT_SETTINGS[k]]);
    settingsSheet.getRange(2, 1, rows.length, 2).setValues(rows);
    settingsSheet.setColumnWidths(1, 1, 240);
    settingsSheet.setColumnWidths(2, 1, 320);
  }

  SpreadsheetApp.getUi().alert(
    'Workbook initialized: Raw_Data, Qualified_Leads, Rejected_Leads, Logs, Settings.\n\n' +
    'Next: add your Google Places API Key on the Settings sheet, then run "Generate Leads."'
  );
}

/**
 * Creates a sheet with the given name and headers if it doesn't exist,
 * or returns the existing one. Idempotent — safe to call repeatedly.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - the active spreadsheet
 * @param {string} name - sheet tab name
 * @param {string[]} headers - column header values for row 1
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} the (possibly new) sheet
 */
function ensureSheetWithHeaders(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Returns a Set of existing Place IDs already present in Raw_Data (for dedup). */
function getExistingPlaceIds() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_RAW);
  const ids = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return ids;

  const placeIdCol = RAW_HEADERS.indexOf('Place ID') + 1;
  const values = sheet.getRange(2, placeIdCol, lastRow - 1, 1).getValues();
  values.forEach(r => { if (r[0]) ids.add(r[0]); });
  return ids;
}

/**
 * Returns existing (name, phone, address) fingerprints from Raw_Data for a
 * secondary duplicate check, in case a business appears under a different
 * Place ID (e.g. re-indexed listing) but is clearly the same business.
 */
function getExistingFingerprints() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_RAW);
  const fingerprints = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return fingerprints;

  const nameCol = RAW_HEADERS.indexOf('Business Name');
  const phoneCol = RAW_HEADERS.indexOf('Phone');
  const addressCol = RAW_HEADERS.indexOf('Address');
  const values = sheet.getRange(2, 1, lastRow - 1, RAW_HEADERS.length).getValues();

  values.forEach(row => {
    fingerprints.add(buildFingerprint(row[nameCol], row[phoneCol], row[addressCol]));
  });
  return fingerprints;
}

/**
 * Reads Raw_Data ONCE and returns both Place ID set and fingerprint set.
 * Used by processBatch() to avoid two separate sheet reads — the
 * fingerprint read already requires all columns, so extracting Place IDs
 * from the same data is free.
 *
 * @returns {{placeIds: Set<string>, fingerprints: Set<string>}}
 */
function getExistingDedupData() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_RAW);
  const placeIds = new Set();
  const fingerprints = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { placeIds: placeIds, fingerprints: fingerprints };

  const placeIdCol = RAW_HEADERS.indexOf('Place ID');
  const nameCol = RAW_HEADERS.indexOf('Business Name');
  const phoneCol = RAW_HEADERS.indexOf('Phone');
  const addressCol = RAW_HEADERS.indexOf('Address');
  const values = sheet.getRange(2, 1, lastRow - 1, RAW_HEADERS.length).getValues();

  values.forEach(row => {
    if (row[placeIdCol]) placeIds.add(row[placeIdCol]);
    fingerprints.add(buildFingerprint(row[nameCol], row[phoneCol], row[addressCol]));
  });
  return { placeIds: placeIds, fingerprints: fingerprints };
}

/**
 * Builds a dedup fingerprint from a business's name, phone, and address.
 * Uses normalized (lowercase, alphanumeric-only) values joined by pipes.
 * @param {string} name - business name
 * @param {string} phone - phone number
 * @param {string} address - street address
 * @returns {string} fingerprint string like "acmeplumbing|5551234567|123mainst"
 */
function buildFingerprint(name, phone, address) {
  return normalizeName(name) + '|' + normalizePhone(phone) + '|' + normalizeAddress(address);
}

/**
 * Appends one lead to the Raw_Data sheet.
 * @param {Object} lead - the lead object built by buildLeadFromPlace() in Code.js
 */
function appendRawRow(lead) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_RAW);
  sheet.appendRow([
    now(), lead.name, lead.industry, lead.address, lead.phone, lead.website || '',
    lead.websiteStatus, lead.email || '', lead.emailSourceUrl || '', lead.score,
    lead.rating || '', lead.reviewCount || '', lead.mapsUrl || '', lead.lat || '',
    lead.lng || '', lead.placeId, lead.businessStatus || '', lead.notes || ''
  ]);
}

/** Appends one row to Qualified_Leads. */
function appendQualifiedRow(lead) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_QUALIFIED);
  sheet.appendRow([
    lead.name, lead.industry, lead.owner || '', lead.email, lead.phone,
    lead.website || '', lead.websiteStatus, lead.emailType || '',
    lead.recommendedChannel || '', lead.readinessScore || '', lead.readinessNotes || '',
    lead.score, lead.rating || '',
    lead.reviewCount || '', lead.address, lead.mapsUrl || '', lead.notes || '',
    lead.placeId
  ]);
}

/** Appends one row to Rejected_Leads. */
function appendRejectedRow(name, reason, placeId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_REJECTED);
  sheet.appendRow([name, reason, placeId || '', now()]);
}

/** Appends one row to Logs. */
function appendLogRow(action, checked, qualified, rejected, errors, execSeconds, notes) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_LOGS);
  sheet.appendRow([now(), action, checked, qualified, rejected, errors, execSeconds, notes || '']);
}

/**
 * Clears all data rows from a sheet, preserving the header row (row 1).
 * Used by menuFilterQualifiedLeads() and menuRedraftAllOutreachEmails()
 * to rebuild sheet contents from scratch.
 */
function clearSheetBody(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
}


/**
 * Exports the Qualified_Leads sheet as a CSV file in Google Drive.
 * Uses proper CSV escaping (double-quote wrapping, quote doubling).
 * @returns {GoogleAppsScript.Drive.File} the newly created CSV file
 */
function exportQualifiedLeadsToCsv() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_QUALIFIED);
  const data = sheet.getDataRange().getValues();
  const csv = data.map(row =>
    row.map(cell => {
      const value = (cell === null || cell === undefined) ? '' : cell.toString();
      return '"' + value.replace(/"/g, '""') + '"';
    }).join(',')
  ).join('\n');

  const fileName = 'Qualified_Leads_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm') + '.csv';
  const file = DriveApp.createFile(fileName, csv, MimeType.CSV);
  return file;
}

/**
 * Logs an email open event into the Email_Opens sheet.
 * @param {string} leadId - Place ID or unique Lead ID
 * @param {string} userAgent - User agent string from HTTP request if available
 */
function logEmailOpen(leadId, userAgent) {
  if (!leadId) return;
  const ss = SpreadsheetApp.getActive();
  const opensSheet = ensureSheetWithHeaders(ss, SHEET_OPENS, OPENS_HEADERS);
  
  let businessName = '';
  try {
    const draftsSheet = ss.getSheetByName(SHEET_DRAFTS);
    if (draftsSheet && draftsSheet.getLastRow() > 1) {
      const pCol = DRAFT_HEADERS.indexOf('Place ID') + 1;
      const nCol = DRAFT_HEADERS.indexOf('Business Name') + 1;
      const data = draftsSheet.getRange(2, 1, draftsSheet.getLastRow() - 1, draftsSheet.getLastColumn()).getValues();
      for (let i = 0; i < data.length; i++) {
        if (data[i][pCol - 1] === leadId) {
          businessName = data[i][nCol - 1] || '';
          break;
        }
      }
    }
  } catch (e) {
    Logger.log('Could not resolve business name for open event: ' + e.message);
  }

  opensSheet.appendRow([now(), leadId, businessName, 'Opened', userAgent || '']);
}