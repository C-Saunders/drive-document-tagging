function getUi(): GoogleAppsScript.Base.Ui {
  try {
    return DocumentApp.getUi()
  } catch (_) {}

  try {
    return SpreadsheetApp.getUi()
  } catch (_) {}

  try {
    return SlidesApp.getUi()
  } catch (_) {}

  throw new Error('Add-on only supports Docs, Spreadsheets, and Slides')
}

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} _ The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(_: any): void {
  getUi().createAddonMenu()
    .addItem('Manage Tags', 'showSidebar')
    .addToUi()
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e: any): void {
  onOpen(e)
}

function showSidebar(): void {
  var ui = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('Document Tags')

  getUi().showSidebar(ui)
}

function include(filename: string): string {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent()
}

function showPicker() {
  var html = HtmlService.createHtmlOutputFromFile('file_picker')
    .setWidth(600)
    .setHeight(425)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  getUi().showModalDialog(html, 'Select Tag Document');
}

function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}
