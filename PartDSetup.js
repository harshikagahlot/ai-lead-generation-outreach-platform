/**
 * PartDSetup.js
 * One-time setup for the Part D menu.
 */
const PART_D_MENU_NAME = '✨ Outreach Upgrade';

function setupPartD() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Open the lead-generation spreadsheet first.');

  const triggers = ScriptApp.getProjectTriggers();
  const existing = triggers.filter(t => t.getHandlerFunction() === 'addPartDMenu');
  if (!existing.length) {
    ScriptApp.newTrigger('addPartDMenu')
      .forSpreadsheet(ss)
      .onOpen()
      .create();
  }

  addPartDMenu();

  SpreadsheetApp.getUi().alert(
    'Part D is installed.\n\n' +
    'A new "' + PART_D_MENU_NAME + '" menu was added.\n' +
    'Use "Upgrade Outreach Emails" after generating/drafting leads.\n\n' +
    'It creates or updates Gmail DRAFTS only — it never sends.'
  );
}

function addPartDMenu() {
  SpreadsheetApp.getUi()
    .createMenu(PART_D_MENU_NAME)
    .addItem('1. Upgrade Outreach Emails', 'menuUpgradeOutreachEmails')
    .addItem('2. Test Part D', 'testPartD')
    .addToUi();
}
