/**
 * AuthorizationSetup.js
 * Safe one-time authorization helper for the bound Apps Script project.
 * It deliberately performs harmless calls that require the scopes declared
 * in appsscript.json so Google can present the consent screen when needed.
 */
function authorizeLeadGenPermissions() {
  // Forces authorization for UrlFetchApp / script.external_request.
  const response = UrlFetchApp.fetch('https://www.google.com/generate_204', {
    muteHttpExceptions: true,
    followRedirects: true
  });

  // Forces authorization for ScriptApp / script.scriptapp.
  ScriptApp.getProjectTriggers();

  SpreadsheetApp.getUi().alert(
    'Authorization check complete.\n\n' +
    'External requests: OK (HTTP ' + response.getResponseCode() + ')\n' +
    'Script triggers: OK\n\n' +
    'You can now run Generate Leads from the Sheet menu.'
  );
}
