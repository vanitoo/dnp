/** Очистка сформированных PDF. */

function clearGeneratedPdfs() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Очистить сформированные PDF?',
    'Все PDF-файлы в папке квитанций и во всех её подпапках будут перемещены в корзину.\n\n' +
      'Шаблон, таблица и сами папки останутся на месте.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast('Поиск сформированных PDF…', 'ДНП', 5);
    const folder = getDnpPdfFolder_();
    const result = trashDnpPdfFilesRecursively_(folder);
    const message = result.failed
      ? 'Перемещено в корзину: ' + result.deleted + '. Не удалось удалить: ' + result.failed + '.'
      : 'Перемещено в корзину PDF-файлов: ' + result.deleted + '.';
    ss.toast(message, 'ДНП', 8);
    ui.alert('Очистка завершена', message + '\n\nПапка: ' + folder.getName(), ui.ButtonSet.OK);
  } catch (error) {
    const message = error && error.message ? error.message : String(error);
    ss.toast('Ошибка очистки PDF', 'ДНП', 8);
    ui.alert('Ошибка очистки PDF', message, ui.ButtonSet.OK);
    throw error;
  }
}

function getDnpPdfFolder_() {
  const folderId = getStoredPdfFolderId_();
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      folder.getName();
      return folder;
    } catch (error) {}
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Не удалось определить активную Google Таблицу.');
  const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
  const parents = spreadsheetFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const acceptedNames = ['Квитанции ДНП Комфорт', 'Receipts', 'Квитанции'];
  const folders = parent.getFolders();
  let fallback = null;

  while (folders.hasNext()) {
    const folder = folders.next();
    const name = folder.getName();
    if (acceptedNames.some(prefix => name.indexOf(prefix) === 0)) {
      savePdfFolderId_(folder.getId());
      return folder;
    }
    if (!fallback && /квитанц|receipt/i.test(name)) fallback = folder;
  }

  if (fallback) {
    savePdfFolderId_(fallback.getId());
    return fallback;
  }

  throw new Error('Папка сформированных квитанций не найдена. Выполните первичную настройку.');
}

function getStoredPdfFolderId_() {
  const stores = getAvailablePropertyStores_();
  for (let i = 0; i < stores.length; i++) {
    try {
      const value = stores[i].getProperty('PDF_FOLDER_ID');
      if (value) return value;
    } catch (error) {}
  }
  return '';
}

function savePdfFolderId_(folderId) {
  const stores = getAvailablePropertyStores_();
  for (let i = 0; i < stores.length; i++) {
    try {
      stores[i].setProperty('PDF_FOLDER_ID', folderId);
      return;
    } catch (error) {}
  }
}

function getAvailablePropertyStores_() {
  const stores = [];
  try { stores.push(PropertiesService.getDocumentProperties()); } catch (error) {}
  try { stores.push(PropertiesService.getScriptProperties()); } catch (error) {}
  try { stores.push(PropertiesService.getUserProperties()); } catch (error) {}
  return stores.filter(Boolean);
}

function trashDnpPdfFilesRecursively_(folder) {
  let deleted = 0;
  let failed = 0;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const isPdf = file.getMimeType() === MimeType.PDF || /\.pdf$/i.test(file.getName());
    if (!isPdf) continue;
    try { file.setTrashed(true); deleted++; } catch (error) { failed++; }
  }
  const childFolders = folder.getFolders();
  while (childFolders.hasNext()) {
    const result = trashDnpPdfFilesRecursively_(childFolders.next());
    deleted += result.deleted;
    failed += result.failed;
  }
  return { deleted: deleted, failed: failed };
}
