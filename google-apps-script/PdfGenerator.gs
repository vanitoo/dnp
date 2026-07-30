/**
 * Формирование PDF-квитанций.
 *
 * Границы каждого участка определяются не фиксированным количеством строк,
 * а строками «Тариф» в столбце B. Блок начинается со строки «Тариф» и
 * заканчивается перед следующей строкой «Тариф».
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

  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  if (lastRow < 2) throw new Error('Лист «' + year + '» пуст.');

  const blocks = getReceiptBlocks_(sheet);
  if (!blocks.length) {
    throw new Error('Не найдены строки «Тариф» в столбце B.');
  }

  const rootFolder = getDnpPdfFolder_();
  const yearFolder = getOrCreatePdfChildFolder_(rootFolder, String(year));
  const monthFolderName = String(month).padStart(2, '0') + ' ' + getRussianMonthName_(month);
  const monthFolder = getOrCreatePdfChildFolder_(yearFolder, monthFolderName);

  const spreadsheetId = ss.getId();
  const gid = sheet.getSheetId();
  const token = ScriptApp.getOAuthToken();
  let created = 0;
  let failed = 0;
  const errors = [];

  blocks.forEach((block, index) => {
    const safePlot = sanitizePdfFileName_(block.plot);
    const fileName = 'Квитанция_участок_' + safePlot + '_' + year + '_' +
      String(month).padStart(2, '0') + '.pdf';

    ss.toast(
      'Формируется ' + (index + 1) + ' из ' + blocks.length + ': участок ' + block.plot,
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
        '&r1=' + (block.startRow - 1) +
        '&r2=' + block.endRow +
        '&c1=0' +
        '&c2=' + lastColumn;

      const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      if (code !== 200) throw new Error('Google вернул HTTP ' + code);

      monthFolder.createFile(response.getBlob().setName(fileName));
      created++;
      Utilities.sleep(250);
    } catch (error) {
      failed++;
      errors.push('Участок ' + block.plot + ': ' + error.message);
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

function getReceiptBlocks_(sheet) {
  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(1, 1, lastRow, 2).getDisplayValues();
  const starts = [];

  for (let index = 1; index < rows.length; index++) {
    const plot = String(rows[index][0] || '').trim();
    const label = normalizeDnpRowLabel_(rows[index][1]);
    if (/^тариф(?:ы)?$/.test(label)) {
      if (!plot) {
        throw new Error('В строке ' + (index + 1) + ' найден «Тариф», но в столбце A нет номера участка.');
      }
      starts.push({ startRow: index + 1, plot: plot });
    }
  }

  return starts.map((item, index) => ({
    plot: item.plot,
    startRow: item.startRow,
    endRow: index + 1 < starts.length ? starts[index + 1].startRow - 1 : lastRow
  }));
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
  return String(value).trim().replace(/[\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_');
}

function trashFilesByName_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    try { files.next().setTrashed(true); } catch (error) {}
  }
}
