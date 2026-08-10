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
      // Apps Script does not expose the HTTP User-Agent through the doGet
      // event object, so leave that field blank rather than inventing it.
      logEmailOpen(leadId, '');
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
 * Builds a per-lead summary of recorded open events and matches them to
 * Outreach_Drafts using Place ID.
 *
 * @returns {Object[]} summary rows.
 */
function getOpenTrackingSummary() {
  const ss = SpreadsheetApp.getActive();
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
      if (!leadId) return;

      if (!eventsByLead[leadId]) {
        eventsByLead[leadId] = {
          firstOpened: row[tsCol] || '',
          lastOpened: row[tsCol] || '',
          totalOpens: 0
        };
      }

      if (String(row[eventCol] || '').trim() !== 'Opened') return;

      const timestamp = row[tsCol] || '';
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

/** Formats a timestamp for the spreadsheet UI without assuming a timezone. */
function formatTrackingTimestamp(value) {
  if (!value) return '—';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value);
}
