/**
 * EmailOpenTracking.gs
 * -----------------------------------------------------------------------
 * Part B: email-open tracking support.
 *
 * IMPORTANT LIMITATION:
 * Google Apps Script ContentService does not expose an image/* MIME type.
 * This endpoint therefore returns a 1x1 SVG payload as XML after logging
 * the event. The logging endpoint itself works, but using it as a Gmail
 * tracking pixel is EXPERIMENTAL and is not guaranteed to be fetched by
 * Gmail's image proxy because the response is not served as image/svg+xml.
 * Do not treat open tracking as authoritative engagement data.
 */

const EMAIL_OPEN_SPREADSHEET_ID_PROPERTY = 'EMAIL_OPEN_TRACKING_SPREADSHEET_ID';

/**
 * One-time setup for the web-app endpoint.
 * Run this from the bound spreadsheet project while the target spreadsheet
 * is open. It stores the parent spreadsheet ID in Script Properties so the
 * web-app request does not depend on an "active" spreadsheet context.
 */
function initializeEmailOpenTracking() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'No active spreadsheet found. Open the lead-generation spreadsheet and run this function again.'
    );
  }

  PropertiesService.getScriptProperties().setProperty(
    EMAIL_OPEN_SPREADSHEET_ID_PROPERTY,
    ss.getId()
  );

  let sheet = ss.getSheetByName(SHEET_OPENS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_OPENS);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, OPENS_HEADERS.length)
      .setValues([OPENS_HEADERS]);
  }

  SpreadsheetApp.flush();

  Logger.log('Email open tracking initialized successfully.');
  Logger.log('Spreadsheet ID: ' + ss.getId());
  Logger.log('Sheet: ' + SHEET_OPENS);
}

/**
 * Returns the spreadsheet used by the tracking endpoint.
 * For normal spreadsheet/menu execution, the active spreadsheet is used.
 * For web-app execution, the stored spreadsheet ID is used because there
 * may be no active spreadsheet UI context.
 */
function getEmailOpenTrackingSpreadsheet() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  const id = PropertiesService.getScriptProperties().getProperty(
    EMAIL_OPEN_SPREADSHEET_ID_PROPERTY
  );
  if (!id) {
    throw new Error('Email open tracking is not initialized. Run initializeEmailOpenTracking() once from the bound spreadsheet.');
  }

  return SpreadsheetApp.openById(id);
}

/**
 * Web-app GET endpoint for an email-open event.
 *
 * Expected URL:
 *   WEB_APP_URL?leadId=PLACE_ID
 *
 * The endpoint logs the event first, then returns a 1x1 SVG payload.
 *
 * @param {Object} e Apps Script web-app event object.
 * @returns {GoogleAppsScript.Content.TextOutput} XML response.
 */
function doGet(e) {
  const leadId = e && e.parameter ? String(e.parameter.leadId || '').trim() : '';

  if (leadId) {
    try {
      logEmailOpenFromWebApp(leadId);
    } catch (err) {
      // Do not expose spreadsheet/database errors to the requester.
      console.error('Email open logging failed: ' + err.message);
    }
  }

  // ContentService supports XML but not image/* MIME types. This is a
  // transparent 1x1 SVG payload, but the XML MIME type means Gmail may not
  // treat it as an image. The endpoint is therefore intentionally marked
  // experimental until tested against the deployed Gmail image proxy.
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1" viewBox="0 0 1 1"><rect width="1" height="1" fill="none"/></svg>';
  return ContentService.createTextOutput(svg).setMimeType(ContentService.MimeType.XML);
}

/**
 * Web-app-safe logger. It does not rely on SpreadsheetApp.getActive(),
 * because a web-app request has no spreadsheet UI context.
 */
function logEmailOpenFromWebApp(leadId) {
  if (!leadId) return;

  const ss = getEmailOpenTrackingSpreadsheet();
  const opensSheet = ensureSheetWithHeaders(ss, SHEET_OPENS, OPENS_HEADERS);
  let businessName = '';

  try {
    const draftsSheet = ss.getSheetByName(SHEET_DRAFTS);
    if (draftsSheet && draftsSheet.getLastRow() > 1) {
      const pCol = DRAFT_HEADERS.indexOf('Place ID') + 1;
      const nCol = DRAFT_HEADERS.indexOf('Business Name') + 1;
      const data = draftsSheet
        .getRange(2, 1, draftsSheet.getLastRow() - 1, draftsSheet.getLastColumn())
        .getValues();

      for (let i = 0; i < data.length; i++) {
        if (String(data[i][pCol - 1] || '').trim() === leadId) {
          businessName = data[i][nCol - 1] || '';
          break;
        }
      }
    }
  } catch (err) {
    console.error('Could not resolve business name for open event: ' + err.message);
  }

  opensSheet.appendRow([now(), leadId, businessName, 'Opened', '']);
}

/**
 * Builds a per-lead summary of recorded open events and matches them to
 * Outreach_Drafts using Place ID.
 *
 * @returns {Object[]} summary rows.
 */
function getOpenTrackingSummary() {
  const ss = getEmailOpenTrackingSpreadsheet();
  const draftsSheet = ss.getSheetByName(SHEET_DRAFTS);
  const opensSheet = ss.getSheetByName(SHEET_OPENS);

  if (!draftsSheet) return [];

  const draftRows = draftsSheet.getLastRow() > 1
    ? draftsSheet.getRange(2, 1, draftsSheet.getLastRow() - 1, Math.max(DRAFT_HEADERS.length, draftsSheet.getLastColumn())).getValues()
    : [];

  const eventsByLead = {};

  if (opensSheet && opensSheet.getLastRow() > 1) {
    const opensRows = opensSheet.getRange(2, 1, opensSheet.getLastRow() - 1, OPENS_HEADERS.length).getValues();
    const tsCol = OPENS_HEADERS.indexOf('Timestamp');
    const leadCol = OPENS_HEADERS.indexOf('Lead ID');
    const eventCol = OPENS_HEADERS.indexOf('Event');

    opensRows.forEach(row => {
      const leadId = String(row[leadCol] || '').trim();
      if (!leadId || String(row[eventCol] || '').trim() !== 'Opened') return;

      const timestamp = row[tsCol] || '';
      if (!eventsByLead[leadId]) {
        eventsByLead[leadId] = {
          firstOpened: timestamp,
          lastOpened: timestamp,
          totalOpens: 1
        };
        return;
      }

      const item = eventsByLead[leadId];
      item.totalOpens++;
      if (!item.firstOpened || (timestamp && timestamp < item.firstOpened)) item.firstOpened = timestamp;
      if (!item.lastOpened || (timestamp && timestamp > item.lastOpened)) item.lastOpened = timestamp;
    });
  }

  const nameCol = DRAFT_HEADERS.indexOf('Business Name');
  const emailCol = DRAFT_HEADERS.indexOf('Email');
  const placeIdCol = DRAFT_HEADERS.indexOf('Place ID');
  const statusCol = DRAFT_HEADERS.indexOf('Status');

  return draftRows.map(row => {
    const leadId = String(row[placeIdCol] || '').trim();
    const event = eventsByLead[leadId] || { firstOpened: '', lastOpened: '', totalOpens: 0 };

    return {
      leadId: leadId,
      businessName: row[nameCol] || '',
      email: row[emailCol] || '',
      status: row[statusCol] || '',
      opened: event.totalOpens > 0 ? 'Yes' : 'No',
      firstOpened: event.firstOpened,
      lastOpened: event.lastOpened,
      totalOpens: event.totalOpens
    };
  });
}

/**
 * Menu action: displays a concise per-lead open summary.
 */
function menuViewEmailOpens() {
  const ui = SpreadsheetApp.getUi();
  const summary = getOpenTrackingSummary();

  if (!summary.length) {
    ui.alert('Email Opens', 'No Outreach_Drafts rows are available yet.', ui.ButtonSet.OK);
    return;
  }

  const openedLeads = summary.filter(r => r.opened === 'Yes').length;
  const totalRecordedOpens = summary.reduce((sum, r) => sum + Number(r.totalOpens || 0), 0);

  const lines = summary.slice(0, 50).map(r => {
    const first = r.firstOpened ? formatTrackingTimestamp(r.firstOpened) : '—';
    const last = r.lastOpened ? formatTrackingTimestamp(r.lastOpened) : '—';
    return [
      r.businessName || r.leadId || '(unknown lead)',
      r.opened,
      'opens=' + r.totalOpens,
      'first=' + first,
      'last=' + last
    ].join(' | ');
  });

  const omitted = summary.length > 50 ? '\n\nShowing first 50 of ' + summary.length + ' leads.' : '';
  const message =
    'Leads with recorded opens: ' + openedLeads + '/' + summary.length + '\n' +
    'Total recorded open events: ' + totalRecordedOpens + '\n\n' +
    lines.join('\n') + omitted +
    '\n\nNOTE: Opens are directional only. Gmail/email clients may block, cache, proxy, or prefetch images, so an event does not prove a person actually read the email.';

  ui.alert('Email Opens', message, ui.ButtonSet.OK);
}

/**
 * Lightweight Apps Script-side verification helper.
 * It exercises doGet() with a synthetic lead ID and verifies that the
 * endpoint returns XML text. It also confirms the logging path can write
 * to Email_Opens. Run only after initializeEmailOpenTracking().
 */
function testEmailOpenTracking() {
  const testLeadId = 'TEST_OPEN_TRACKING_' + new Date().getTime();
  const before = getEmailOpenTrackingSpreadsheet().getSheetByName(SHEET_OPENS).getLastRow();
  const response = doGet({ parameter: { leadId: testLeadId } });
  const after = getEmailOpenTrackingSpreadsheet().getSheetByName(SHEET_OPENS).getLastRow();

  if (after !== before + 1) {
    throw new Error('FAIL: doGet() did not append exactly one Email_Opens row.');
  }

  if (response.getMimeType() !== ContentService.MimeType.XML) {
    throw new Error('FAIL: tracking endpoint did not return XML MIME type.');
  }

  Logger.log('PASS: doGet() logged one event and returned XML.');
  return 'PASS';
}

/** Formats a timestamp for the spreadsheet UI without assuming a timezone. */
function formatTrackingTimestamp(value) {
  if (!value) return '—';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value);
}
