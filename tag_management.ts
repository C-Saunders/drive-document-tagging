import 'google-apps-script'
import {
  getTagDocumentUrl,
} from './configuration'

type File = GoogleAppsScript.Drive.File

type Tags = Array<string>

function getAvailableTags(): Tags {
  const tags = JSON.parse(DocumentApp.openByUrl(getTagDocumentUrl()).getBody().getText())

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
