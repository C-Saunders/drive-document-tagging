import 'google-apps-script'
import { getTagDocumentUrl } from './configuration'

type File = GoogleAppsScript.Drive.File

type Tags = Array<string>

function getAvailableTags(): Tags {
  const tagDocumentUrl = getTagDocumentUrl()
  if (!tagDocumentUrl) {
    return []
  }

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

  // AppsScript does not support Array.includes
  if (description.indexOf(tag) === -1) {
    description.push(tag)
  }

  activeDocument.setDescription(description.join(', '))
}

function removeTagFromActiveDocument(tag: string): void {
  const activeDocument = getActiveDocument()
  const description = getDocumentDescription(activeDocument)

  activeDocument.setDescription(filterArray(description, item => item !== tag).join(', '))
}

function getActiveDocument(): File {
  const file = DocumentApp.getActiveDocument()
    || SpreadsheetApp.getActiveSpreadsheet()
    || SlidesApp.getActivePresentation()

  if (file !== null) {
    return DriveApp.getFileById(file.getId())
  }

  throw new Error('Add-on only supports Docs, Spreadsheets, and Slides')
}

function getDocumentDescription(doc: File): Tags {
  const desc = doc.getDescription()
  return desc ? desc.split(/, ?/) : []
}

// AppsScript does not support Array.filter
function filterArray<T>(arr: Array<T>, predicate: (item: T) => boolean): Array<T> {
  const result = []
  for (var i = 0; i < arr.length; i += 1) {
    const item = arr[i]
    if (predicate(item)) {
      result.push(item)
    }
  }

  return result
}
