/**
 * Формирование отдельных PDF-квитанций из листа выбранного года.
 *
 * Каждая квитанция занимает 7 строк. Номер участка берётся из столбца A.
 * PDF сохраняются в: Квитанции ДНП Комфорт/<год>/<MM Месяц>/
 */

function generatePdfsForMonth(year, month) {
  year = Number(year);
  month = Number(month);

  if (!Number.isInteger(year) || year < 2000 || year > 2100) {
    throw new Error('Некорректный год.');
  }
  if (!Number.isInteger(month) || month < 1 || month > 12) {
    throw new Error('Некорректный месяц.');
  }

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(String(year));
  if (!sheet) throw new Error('Лист «' + year + '» не найден.');

  const rootFolder = getDnpPdfFolder_();
  const yearFolder = getOrCreatePdfChildFolder_(rootFolder, String(year));
  const monthFolderName = String(month).padStart(2, '0') + ' ' + getRussianMonthName_(month);
  const monthFolder = getOrCreatePdfChildFolder_(yearFolder, monthFolderName);

  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  if (lastRow < 1) throw new Error('Лист «' + year + '» пуст.');

  const firstColumnValues = sheet.getRange(1, 1, lastRow, 1).getDisplayValues().flat();
  const receiptRows = [];

  firstColumnValues.forEach((value, index) => {
    const text = String(value || '').trim();
    if (!text) return;
    if (/участок|итого|тариф|электроэнерг|водоотвед/i.test(text)) return;

    // В рабочей таблице каждая квитанция начинается со строки с номером участка.
    receiptRows.push({ row: index + 1, plot: text });
  });

  if (!receiptRows.length) {
    throw new Error('Не найдены строки с номерами участков в столбце A.');
  }

  const spreadsheetId = ss.getId();
  const gid = sheet.getSheetId();
  const token = ScriptApp.getOAuthToken();
  let created = 0;
  let failed = 0;
  const errors = [];

  receiptRows.forEach((item, index) => {
    const startRow = item.row;
    const endRow = Math.min(startRow + 7, lastRow + 1); // r2 не включается
    const safePlot = sanitizePdfFileName_(item.plot);
    const fileName = 'Квитанция_участок_' + safePlot + '_' + year + '_' + String(month).padStart(2, '0') + '.pdf';

    ss.toast(
      'Формируется ' + (index + 1) + ' из ' + receiptRows.length + ': участок ' + item.plot,
      'ДНП',
      5
    );

    try {
      trashFilesByName_(monthFolder, fileName);

      const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export' +
        '?format=pdf' +
        '&gid=' + gid +
        '&size=A4' +
        '&portrait=true' +
        '&fitw=true' +
        '&sheetnames=false' +
        '&printtitle=false' +
        '&pagenumbers=false' +
        '&gridlines=false' +
        '&fzr=false' +
        '&top_margin=0.30' +
        '&bottom_margin=0.30' +
        '&left_margin=0.30' +
        '&right_margin=0.30' +
        '&r1=' + (startRow - 1) +
        '&r2=' + endRow +
        '&c1=0' +
        '&c2=' + lastColumn;

      const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      if (code !== 200) {
        throw new Error('Google вернул HTTP ' + code);
      }

      const blob = response.getBlob().setName(fileName);
      monthFolder.createFile(blob);
      created++;

      Utilities.sleep(250);
    } catch (error) {
      failed++;
      errors.push('Участок ' + item.plot + ': ' + error.message);
    }
  });

  const message = failed
    ? 'Создано PDF: ' + created + '. Ошибок: ' + failed + '. Первая ошибка: ' + errors[0]
    : 'Создано PDF: ' + created + '. Папка: ' + year + '/' + monthFolderName + '.';

  ss.toast(message, 'ДНП', 10);

  return {
    ok: failed === 0,
    created: created,
    failed: failed,
    folderId: monthFolder.getId(),
    folderUrl: monthFolder.getUrl(),
    message: message
  };
}

function getOrCreatePdfChildFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function getRussianMonthName_(month) {
  return [
    'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'
  ][month - 1];
}

function sanitizePdfFileName_(value) {
  return String(value)
    .trim()
    .replace(/[\\/:*?"<>|]+/g, '_')
    .replace(/\s+/g, '_');
}

function trashFilesByName_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    try {
      files.next().setTrashed(true);
    } catch (error) {
      // Старый файл не должен блокировать повторное формирование.
    }
  }
}
