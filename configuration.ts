const CONFIG_URL_KEY = 'TAG_DOC_JSON'
const scriptProperties = PropertiesService.getScriptProperties()

export {
  setTagDocumentUrl,
  getTagDocumentUrl,
}

function setTagDocumentUrl(): void {
  const ui = DocumentApp.getUi()
  const response = ui.prompt('URL of Drive document containing the tags JSON', ui.ButtonSet.OK_CANCEL)

  if (response.getSelectedButton() === ui.Button.OK) {
    scriptProperties.setProperty(CONFIG_URL_KEY, response.getResponseText())
  }
}

function getTagDocumentUrl(): string {
  return scriptProperties.getProperty(CONFIG_URL_KEY)
}
