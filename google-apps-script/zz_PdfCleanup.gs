/**
 * Очистка сформированных PDF.
 *
 * Перемещает PDF из папки квитанций и всех вложенных папок в корзину.
 * Работает даже если DocumentProperties недоступны или PDF_FOLDER_ID ещё не записан.
 */

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
      ? 'Перемещено в корзину: ' + result.deleted +
        '. Не удалось удалить: ' + result.failed + '.'
      : 'Перемещено в корзину PDF-файлов: ' + result.deleted + '.';

    ss.toast(message, 'ДНП', 8);
    ui.alert(
      'Очистка завершена',
      message + '\n\nПапка: ' + folder.getName(),
      ui.ButtonSet.OK
    );
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
      const storedFolder = DriveApp.getFolderById(folderId);
      // Принудительное обращение, чтобы сразу проверить доступность папки.
      storedFolder.getName();
      return storedFolder;
    } catch (error) {
      // Сохранённый ID устарел или папка недоступна. Ищем заново.
    }
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Не удалось определить активную Google Таблицу.');
  }

  const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
  const parents = spreadsheetFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  const acceptedNames = [
    'Квитанции ДНП Комфорт',
    'Receipts',
    'Квитанции'
  ];

  const folders = parent.getFolders();
  let fallback = null;

  while (folders.hasNext()) {
    const folder = folders.next();
    const name = folder.getName();

    if (acceptedNames.some(prefix => name.indexOf(prefix) === 0)) {
      savePdfFolderId_(folder.getId());
      return folder;
    }

    if (!fallback && /квитанц|receipt/i.test(name)) {
      fallback = folder;
    }
  }

  if (fallback) {
    savePdfFolderId_(fallback.getId());
    return fallback;
  }

  throw new Error(
    'Папка сформированных квитанций не найдена рядом с таблицей.\n\n' +
    'Выполните «ДНП → Настройка → Первичная настройка», ' +
    'либо проверьте, что папка квитанций находится в той же папке Google Drive, что и таблица.'
  );
}

function getStoredPdfFolderId_() {
  const stores = getAvailablePropertyStores_();

  for (let i = 0; i < stores.length; i++) {
    try {
      const value = stores[i].getProperty('PDF_FOLDER_ID');
      if (value) return value;
    } catch (error) {
      // Переходим к следующему доступному хранилищу.
    }
  }

  return '';
}

function savePdfFolderId_(folderId) {
  const stores = getAvailablePropertyStores_();

  for (let i = 0; i < stores.length; i++) {
    try {
      stores[i].setProperty('PDF_FOLDER_ID', folderId);
      return;
    } catch (error) {
      // Переходим к следующему доступному хранилищу.
    }
  }
}

function getAvailablePropertyStores_() {
  const stores = [];

  try {
    const documentProperties = PropertiesService.getDocumentProperties();
    if (documentProperties) stores.push(documentProperties);
  } catch (error) {
    // DocumentProperties недоступны у автономного проекта.
  }

  try {
    stores.push(PropertiesService.getScriptProperties());
  } catch (error) {
    // Крайне редкий случай, но не ломаем очистку из-за хранилища настроек.
  }

  try {
    stores.push(PropertiesService.getUserProperties());
  } catch (error) {
    // То же самое.
  }

  return stores;
}

function trashDnpPdfFilesRecursively_(folder) {
  let deleted = 0;
  let failed = 0;

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const isPdf =
      file.getMimeType() === MimeType.PDF || /\.pdf$/i.test(file.getName());

    if (!isPdf) continue;

    try {
      file.setTrashed(true);
      deleted++;
    } catch (error) {
      failed++;
    }
  }

  const childFolders = folder.getFolders();
  while (childFolders.hasNext()) {
    const childResult = trashDnpPdfFilesRecursively_(childFolders.next());
    deleted += childResult.deleted;
    failed += childResult.failed;
  }

  return { deleted: deleted, failed: failed };
}
