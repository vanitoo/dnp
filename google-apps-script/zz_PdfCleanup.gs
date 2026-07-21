/**
 * Рабочая очистка сформированных PDF.
 *
 * Файл назван с префиксом zz_, чтобы его реализация clearGeneratedPdfs()
 * загружалась после временной заглушки из Code.gs. При переносе в Apps Script
 * рекомендуется удалить старую функцию-заглушку из конца Code.gs.
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

  const ss = SpreadsheetApp.getActive();
  const folder = getDnpPdfFolder_();

  ss.toast('Поиск сформированных PDF…', 'ДНП', 5);

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
}

function getDnpPdfFolder_() {
  const properties = PropertiesService.getDocumentProperties();
  const folderId = properties.getProperty('PDF_FOLDER_ID');

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (error) {
      // Папка могла быть удалена или перемещена вручную. Ищем её заново.
    }
  }

  const spreadsheetFile = DriveApp.getFileById(
    SpreadsheetApp.getActive().getId()
  );
  const parents = spreadsheetFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folders = parent.getFolders();

  while (folders.hasNext()) {
    const folder = folders.next();
    if (folder.getName().indexOf('Квитанции ДНП Комфорт') === 0) {
      properties.setProperty('PDF_FOLDER_ID', folder.getId());
      return folder;
    }
  }

  throw new Error(
    'Папка сформированных квитанций не найдена. ' +
    'Сначала выполните «ДНП → Настройка → Первичная настройка».'
  );
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
