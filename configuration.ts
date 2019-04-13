const CONFIG_URL_KEY = 'TAG_SPREADSHEET_URL'
const scriptProperties = PropertiesService.getScriptProperties()

export {
  setTagDocumentUrl,
  getTagDocumentUrl,
}

function setTagDocumentUrl(url: string): void {
  scriptProperties.setProperty(CONFIG_URL_KEY, url)
}

function getTagDocumentUrl(): string {
  return scriptProperties.getProperty(CONFIG_URL_KEY)
}
