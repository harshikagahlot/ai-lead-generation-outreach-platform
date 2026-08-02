/**
 * Config.gs
 * -----------------------------------------------------------------------
 * Central place for ALL project-wide constants and for reading
 * user-adjustable settings from the Settings sheet.
 *
 * Contents:
 *   - Sheet names (SHEET_RAW, SHEET_QUALIFIED, SHEET_REJECTED, etc.)
 *   - Settings-sheet key names (YOUR_NAME_SETTING, YOUR_SERVICE_SETTING)
 *   - Column header arrays for every sheet
 *   - Website classification status labels (WEBSITE_STATUS)
 *   - Qualifying website statuses for the lead filter
 *   - Lead scoring point values (SCORING_RULES)
 *   - Default settings for first-run initialization
 *   - getSettings() / getPlacesApiKey() helper functions
 *
 * All other files import constants from here via Apps Script's shared
 * global scope — this is the single source of truth.
 */

// ---- Sheet names ----
const SHEET_RAW = 'Raw_Data';
const SHEET_QUALIFIED = 'Qualified_Leads';
const SHEET_REJECTED = 'Rejected_Leads';
const SHEET_LOGS = 'Logs';
const SHEET_SETTINGS = 'Settings';
const SHEET_DRAFTS = 'Outreach_Drafts';

// ---- Settings-sheet key names (used by EmailDrafts.js) ----
const YOUR_NAME_SETTING = 'Your Name';
const YOUR_SERVICE_SETTING = 'Your Company / Service';

// ---- Column headers (one array per sheet, in column order) ----
const RAW_HEADERS = [
  'Timestamp', 'Business Name', 'Industry', 'Address', 'Phone', 'Website',
  'Website Status', 'Email', 'Email Source URL', 'Lead Score', 'Rating',
  'Review Count', 'Google Maps URL', 'Latitude', 'Longitude', 'Place ID',
  'Business Status', 'Notes'
];

const QUALIFIED_HEADERS = [
  'Business Name', 'Industry', 'Owner', 'Email', 'Phone', 'Website',
  'Website Status', 'Lead Score', 'Rating', 'Reviews', 'Address',
  'Google Maps URL', 'Notes', 'Place ID'
];

const REJECTED_HEADERS = ['Business Name', 'Reason Rejected', 'Place ID', 'Timestamp'];

const LOG_HEADERS = [
  'Timestamp', 'Action', 'Businesses Checked', 'Qualified', 'Rejected',
  'Errors', 'Execution Time (sec)', 'Notes'
];

const DRAFT_HEADERS = [
  'Business Name', 'Industry', 'City', 'Email', 'Phone', 'Website Status',
  'Subject', 'Email Draft', 'Status', 'Place ID', 'Gmail Draft ID'
];

// ---- Website classification thresholds (outdated-flag counts) ----
const WEBSITE_STATUS = {
  NO_WEBSITE: 'No Website',
  BROKEN: 'Broken',
  VERY_OUTDATED: 'Very Outdated',
  OUTDATED: 'Outdated',
  BASIC: 'Basic',
  GOOD: 'Good',
  EXCELLENT: 'Excellent'
};

// Statuses that make a lead eligible (per the mandatory filter rule)
const QUALIFYING_WEBSITE_STATUSES = [
  WEBSITE_STATUS.NO_WEBSITE,
  WEBSITE_STATUS.BROKEN,
  WEBSITE_STATUS.VERY_OUTDATED,
  WEBSITE_STATUS.OUTDATED
];

// ---- Lead scoring point values (tunable via SCORING_RULES below) ----
const SCORING_RULES = {
  NO_WEBSITE: 50,
  BROKEN: 45,
  VERY_OUTDATED: 40,
  OUTDATED: 30,
  BASIC: 20,
  HAS_EMAIL: 20,
  RATING_ABOVE_4_5: 10,
  REVIEWS_50_PLUS: 10,
  RECENT_REVIEW: 5
};

// ---- Default settings (used to seed the Settings sheet on first run) ----
const DEFAULT_SETTINGS = {
  'Google Places API Key': '',
  'Minimum Rating': 0,
  'Minimum Reviews': 0,
  'Default Max Businesses': 20,
  'Request Timeout Retries': 2,
  'Your Name': '',
  'Your Company / Service': '',
  'Batch Size (leads per execution)': 15,
  'Max Seconds Per Batch': 280
};

/**
 * Execution-scoped cache for settings. Populated on first call to
 * getSettings() and reused for all subsequent calls within the same
 * Apps Script execution. Automatically resets when the execution ends
 * (each trigger / menu action gets a fresh V8 context).
 */
var _settingsCache = null;

/**
 * Reads all key/value settings from the Settings sheet into an object.
 * Falls back to DEFAULT_SETTINGS for any missing key.
 *
 * Results are cached for the duration of the current execution to avoid
 * repeated Sheets RPCs — fetchWithRetry() alone can call this 10+ times
 * per lead, which previously meant thousands of redundant sheet reads
 * per batch.
 *
 * @returns {Object} key/value map of all settings (do not mutate)
 */
function getSettings() {
  if (_settingsCache) return _settingsCache;

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
  const settings = Object.assign({}, DEFAULT_SETTINGS);
  if (!sheet) { _settingsCache = settings; return settings; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { _settingsCache = settings; return settings; }

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  data.forEach(row => {
    if (row[0]) settings[row[0]] = row[1];
  });
  _settingsCache = settings;
  return settings;
}

/**
 * Convenience getter for just the API key, with a clear error if missing.
 * @returns {string} the Google Places API key from the Settings sheet
 * @throws {Error} if the key is blank or not set
 */
function getPlacesApiKey() {
  const key = getSettings()['Google Places API Key'];
  if (!key) {
    throw new Error('Missing "Google Places API Key" on the Settings sheet. Add it before running a search.');
  }
  return key;
}