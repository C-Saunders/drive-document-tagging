import 'google-apps-script'
import { getTagDocumentUrl } from './configuration'

type File = GoogleAppsScript.Drive.File

type Tags = Array<string>

function getAvailableTags(): Tags {
  const sheet = SpreadsheetApp.openByUrl(getTagDocumentUrl()).getSheets()[0]
  const values = sheet.getDataRange().getValues().slice(1) // exclude header

  return values.map(row => row[0].toString())
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
