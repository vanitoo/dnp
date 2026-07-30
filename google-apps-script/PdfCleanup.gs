/** Очистка сформированных PDF. */

function clearGeneratedPdfs() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Очистить сформированные PDF?',
    'Все PDF-файлы в папке квитанций и её подпапках будут перемещены в корзину. Папки и шаблон останутся.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActive();
  try {
    ss.toast('Поиск сформированных PDF…', 'ДНП', 5);
    const folder = getDnpPdfFolder_();
    const result = trashDnpPdfFilesRecursively_(folder);
    const message = result.failed
      ? 'Перемещено в корзину: ' + result.deleted + '. Не удалось удалить: ' + result.failed + '.'
      : 'Перемещено в корзину PDF-файлов: ' + result.deleted + '.';
    ss.toast(message, 'ДНП', 8);
    ui.alert('Очистка завершена', message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Ошибка очистки PDF', error.message || String(error), ui.ButtonSet.OK);
    throw error;
  }
}

function getDnpPdfFolder_() {
  const stores = getAvailablePropertyStores_();
  for (let i = 0; i < stores.length; i++) {
    try {
      const id = stores[i].getProperty('PDF_FOLDER_ID');
      if (id) {
        const folder = DriveApp.getFolderById(id);
        folder.getName();
        return folder;
      }
    } catch (error) {}
  }

  const parent = getSpreadsheetParentFolder_();
  const folders = parent.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (/^(Квитанции ДНП Комфорт|Квитанции|Receipts)/i.test(folder.getName())) {
      saveProperty_('PDF_FOLDER_ID', folder.getId());
      return folder;
    }
  }
  throw new Error('Папка квитанций не найдена. Выполните первичную настройку.');
}

function trashDnpPdfFilesRecursively_(folder) {
  let deleted = 0;
  let failed = 0;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== MimeType.PDF && !/\.pdf$/i.test(file.getName())) continue;
    try { file.setTrashed(true); deleted++; } catch (error) { failed++; }
  }
  const children = folder.getFolders();
  while (children.hasNext()) {
    const result = trashDnpPdfFilesRecursively_(children.next());
    deleted += result.deleted;
    failed += result.failed;
  }
  return { deleted, failed };
}
