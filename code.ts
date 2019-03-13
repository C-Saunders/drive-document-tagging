import 'google-apps-script'

type File = GoogleAppsScript.Drive.File

const config = {
  tagDocumentUrl: 'https://docs.google.com/document/d/1gbRR0wmTzFojUEXoRr6-91Mh0IDT3b1Fti_LWIphXj0/edit'
}

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e: any): void {
  DocumentApp.getUi().createAddonMenu()
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

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar(): void {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Document Tags');
  DocumentApp.getUi().showSidebar(ui);
}

function getTagDocumentUrl(): string {
  return config.tagDocumentUrl
}

type Tags = Array<string>

function getAvailableTags(): Tags {
  const tags = JSON.parse(DocumentApp.openByUrl(config.tagDocumentUrl).getBody().getText())

  // may be able to use the descriptions at some point as well
  return Object.keys(tags)
}

function getCurrentTags(): Tags {
  return getDocumentDescription(getActiveDocument())
}

function addTagToActiveDocument(tag: string): void {
  const activeDocument = getActiveDocument()
  const description = getDocumentDescription(activeDocument)

  if (!description.includes(tag)) {
    description.push(tag)
  }

  activeDocument.setDescription(description.join(', '))
}

function removeTagFromActiveDocument(tag: string): void {
  const activeDocument = getActiveDocument()
  const description = getDocumentDescription(activeDocument)

  activeDocument.setDescription(description.filter(item => item !== tag).join(', '))
}

function getActiveDocument(): File {
  return DriveApp.getFileById(DocumentApp.getActiveDocument().getId())
}

function getDocumentDescription(doc: File): Tags {
  const desc = doc.getDescription()
  return desc ? desc.split(/, ?/) : []
}
