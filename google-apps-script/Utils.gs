/** Общие вспомогательные функции. */

function getSpreadsheetParentFolder_() {
  const file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parents = file.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

function getAvailablePropertyStores_() {
  const stores = [];
  try { const p = PropertiesService.getDocumentProperties(); if (p) stores.push(p); } catch (error) {}
  try { stores.push(PropertiesService.getScriptProperties()); } catch (error) {}
  try { stores.push(PropertiesService.getUserProperties()); } catch (error) {}
  return stores;
}

function saveProperty_(key, value) {
  const stores = getAvailablePropertyStores_();
  for (let i = 0; i < stores.length; i++) {
    try { stores[i].setProperty(key, String(value)); return; } catch (error) {}
  }
}

function deleteProperty_(key) {
  getAvailablePropertyStores_().forEach(store => {
    try { store.deleteProperty(key); } catch (error) {}
  });
}

function showDriveLinkDialog_(title, name, url) {
  const html = HtmlService.createHtmlOutput(
    '<div style="font:14px Arial,sans-serif;padding:18px">' +
      '<p><b>' + escapeHtml_(name) + '</b></p>' +
      '<p><a href="' + url + '" target="_blank" style="display:inline-block;padding:10px 14px;background:#1a73e8;color:#fff;text-decoration:none;border-radius:6px">Открыть в Google Drive</a></p>' +
    '</div>'
  ).setWidth(420).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
