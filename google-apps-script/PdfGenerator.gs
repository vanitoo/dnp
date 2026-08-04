/** Формирование PDF-квитанций строго через Google Docs-шаблон. */

function generatePdfsForMonth(year, month) {
  year = Number(year);
  month = Number(month);
  if (!Number.isInteger(year) || year < 2000 || year > 2100) throw new Error('Некорректный год.');
  if (!Number.isInteger(month) || month < 1 || month > 12) throw new Error('Некорректный месяц.');

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(String(year));
  if (!sheet) throw new Error('Лист «' + year + '» не найден.');

  const templateFile = getReceiptTemplateFile_();
  const rootFolder = getDnpPdfFolder_();
  const yearFolder = getOrCreatePdfChildFolder_(rootFolder, String(year));
  const monthFolderName = String(month).padStart(2, '0') + ' ' + getRussianMonthName_(month);
  const monthFolder = getOrCreatePdfChildFolder_(yearFolder, monthFolderName);
  const blocks = getReceiptBlocks_(sheet);
  const rates = getReceiptRates_();
  if (!blocks.length) throw new Error('Не найдены строки «Тариф» в столбце B.');

  let created = 0;
  let failed = 0;
  const errors = [];

  blocks.forEach((block, index) => {
    let tempFile = null;
    ss.toast('Формируется ' + (index + 1) + ' из ' + blocks.length + ': участок ' + block.plot, 'ДНП', 5);

    try {
      const receipt = buildReceiptData_(sheet, block, year, month, rates);
      const fileName = 'Квитанция_участок_' + sanitizePdfFileName_(block.plot) + '_' + year + '_' + String(month).padStart(2, '0') + '.pdf';
      trashFilesByName_(monthFolder, fileName);

      tempFile = templateFile.makeCopy('Временная квитанция ' + block.plot, monthFolder);
      const doc = DocumentApp.openById(tempFile.getId());
      fillReceiptTemplate_(doc, receipt);
      doc.saveAndClose();

      monthFolder.createFile(tempFile.getAs(MimeType.PDF).setName(fileName));
      tempFile.setTrashed(true);
      tempFile = null;
      created++;

      if (DNP_PDF_SLEEP_MS > 0) Utilities.sleep(DNP_PDF_SLEEP_MS);
    } catch (error) {
      failed++;
      errors.push('Участок ' + block.plot + ': ' + (error.message || error));
      if (tempFile) {
        try { tempFile.setTrashed(true); } catch (cleanupError) {}
      }
    }
  });

  const message = failed
    ? 'Создано PDF: ' + created + '. Ошибок: ' + failed + '. Первая ошибка: ' + errors[0]
    : 'Создано PDF: ' + created + '. Папка: ' + year + '/' + monthFolderName + '.';

  ss.toast(message, 'ДНП', 10);
  return {
    ok: failed === 0,
    created,
    failed,
    folderId: monthFolder.getId(),
    folderUrl: monthFolder.getUrl(),
    message,
  };
}

function getReceiptBlocks_(sheet) {
  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(1, 1, lastRow, 2).getDisplayValues();
  const starts = [];

  for (let index = 1; index < rows.length; index++) {
    const plot = String(rows[index][0] || '').trim();
    const label = normalizeReceiptLabel_(rows[index][1]);
    if (/^тариф(?:ы)?$/.test(label)) {
      if (!plot) {
        throw new Error('В строке ' + (index + 1) + ' найден «Тариф», но в столбце A нет номера участка.');
      }
      starts.push({ startRow: index + 1, plot });
    }
  }

  return starts.map((item, index) => ({
    plot: item.plot,
    startRow: item.startRow,
    endRow: index + 1 < starts.length ? starts[index + 1].startRow - 1 : lastRow,
  }));
}

function buildReceiptData_(sheet, block, year, month, rates) {
  const monthColumn = month + 2;
  const previousMonthColumn = monthColumn - 1;
  const rowCount = block.endRow - block.startRow + 1;
  const lastColumn = Math.max(sheet.getLastColumn(), monthColumn);
  const values = sheet.getRange(block.startRow, 1, rowCount, lastColumn).getDisplayValues();
  const byLabel = {};

  values.forEach(row => {
    const label = normalizeReceiptLabel_(row[1]);
    if (label) byLabel[label] = row;
  });

  const paymentRows = [];
  let calculatedTotal = 0;

  ['т1', 'т2', 'т3'].forEach((label, index) => {
    const row = byLabel[label];
    if (!row) return;

    const current = parseReceiptNumber_(row[monthColumn - 1]);
    const previous = month > 1 ? parseReceiptNumber_(row[previousMonthColumn - 1]) : null;
    const usage = current !== null && previous !== null ? current - previous : null;
    const rate = rates['T' + (index + 1)];
    const amount = usage !== null && rate !== null ? usage * rate : null;

    if (amount !== null) calculatedTotal += amount;
    paymentRows.push([
      'Электроэнергия Т' + (index + 1),
      formatReceiptValue_(current),
      formatReceiptValue_(previous),
      formatReceiptValue_(usage),
      formatReceiptMoney_(rate),
      formatReceiptMoney_(amount),
    ]);
  });

  const waterRow = findReceiptRow_(byLabel, [/^водоотвед/, /^вода$/]);
  if (waterRow) {
    const current = parseReceiptNumber_(waterRow[monthColumn - 1]);
    const previous = month > 1 ? parseReceiptNumber_(waterRow[previousMonthColumn - 1]) : null;
    const usage = current !== null && previous !== null ? current - previous : null;
    const amount = usage !== null && rates.WATER !== null ? usage * rates.WATER : null;

    if (amount !== null) calculatedTotal += amount;
    paymentRows.push([
      'Водоотведение',
      formatReceiptValue_(current),
      formatReceiptValue_(previous),
      formatReceiptValue_(usage),
      formatReceiptMoney_(rates.WATER),
      formatReceiptMoney_(amount),
    ]);
  }

  const ignoredLabels = /^(тариф(?:ы)?|т1|т2|т3|квтч|квт·ч|квт\/ч|водоотвед.*|вода|сумма.*|итого.*)$/;
  values.forEach(row => {
    const rawLabel = String(row[1] == null ? '' : row[1])
      .replace(/\u00a0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    const label = normalizeReceiptLabel_(rawLabel);
    if (!label || ignoredLabels.test(label)) return;

    const amount = parseReceiptNumber_(row[monthColumn - 1]);
    if (amount === null) return;

    calculatedTotal += amount;
    paymentRows.push([rawLabel, '—', '—', '—', '—', formatReceiptMoney_(amount)]);
  });

  const totalRow = findReceiptRow_(byLabel, [/^сумма/, /^итого/]);
  const storedTotal = totalRow ? parseReceiptNumber_(totalRow[monthColumn - 1]) : null;
  const total = storedTotal !== null ? storedTotal : calculatedTotal;

  return {
    plot: block.plot,
    year,
    month,
    monthName: getRussianMonthName_(month),
    total,
    paymentRows,
  };
}

function fillReceiptTemplate_(doc, receipt) {
  const body = doc.getBody();
  replaceReceiptText_(body, '{{PLOT}}', receipt.plot);
  replaceReceiptText_(body, '{{YEAR}}', receipt.year);
  replaceReceiptText_(body, '{{MONTH}}', String(receipt.month).padStart(2, '0'));
  replaceReceiptText_(body, '{{MONTH_NAME}}', receipt.monthName);
  replaceReceiptText_(body, '{{TOTAL}}', formatReceiptMoney_(receipt.total));
  insertPaymentTable_(body, receipt.paymentRows, receipt.total);
}

function insertPaymentTable_(body, paymentRows, total) {
  const found = body.findText('\\{\\{PAYMENT_TABLE\\}\\}');
  if (!found) throw new Error('В Google Docs-шаблоне не найден маркер {{PAYMENT_TABLE}}.');

  const paragraph = found.getElement().asText().getParent().asParagraph();
  const index = body.getChildIndex(paragraph);
  const rows = [[
    'Наименование платежа',
    'Текущее',
    'Предыдущее',
    'Объём',
    'Тариф',
    'Сумма к оплате',
  ]].concat(paymentRows);
  rows.push(['', '', '', '', 'ИТОГО К ОПЛАТЕ', formatReceiptMoney_(total) + ' руб.']);

  const table = body.insertTable(index, rows);
  paragraph.editAsText().setText('');

  table.getRow(0).editAsText().setBold(true);
  applyPaymentTableColumnWidths_(table);

  // Объединяем первые пять ячеек последней строки справа налево.
  // В итоге остаются две ячейки: большая с надписью и отдельная с суммой.
  const totalRow = table.getRow(table.getNumRows() - 1);
  for (let cellIndex = 4; cellIndex >= 1; cellIndex--) {
    totalRow.getCell(cellIndex).merge();
  }

  const labelCell = totalRow.getCell(0);
  const amountCell = totalRow.getCell(1);
  labelCell.setText('ИТОГО К ОПЛАТЕ');
  amountCell.setText(formatReceiptMoney_(total) + ' руб.');
  labelCell.editAsText().setBold(true);
  amountCell.editAsText().setBold(true);
  labelCell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
}

function getReceiptTemplateFile_() {
  const templateId = getStoredTemplateId_();
  if (!templateId) {
    throw new Error(
      'На листе «Настройки» не заполнен параметр templateDocId. ' +
      'Вставьте в столбец B ID старого Google Docs-шаблона.'
    );
  }

  try {
    const file = DriveApp.getFileById(templateId);
    if (file.getMimeType() !== MimeType.GOOGLE_DOCS) {
      throw new Error('Файл templateDocId не является Google Docs.');
    }
    return file;
  } catch (error) {
    throw new Error(
      'Не удалось открыть шаблон из строки templateDocId на листе «Настройки». ' +
      'Проверьте ID и доступ к документу.\n\n' + (error.message || error)
    );
  }
}

function getStoredTemplateId_() {
  const settingsValue = getTemplateIdFromSettings_();
  if (settingsValue) return settingsValue;

  const stores = getAvailablePropertyStores_();
  for (let i = 0; i < stores.length; i++) {
    try {
      const value = stores[i].getProperty('TEMPLATE_DOC_ID');
      if (value) return extractGoogleFileId_(value);
    } catch (error) {}
  }
  return '';
}

function getTemplateIdFromSettings_() {
  const settings = SpreadsheetApp.getActive().getSheetByName('Настройки');
  if (!settings || settings.getLastRow() < 1) return '';

  const values = settings.getRange(
    1,
    1,
    settings.getLastRow(),
    Math.max(2, Math.min(settings.getLastColumn(), 3))
  ).getDisplayValues();

  for (let row = 0; row < values.length; row++) {
    const key = String(values[row][0] || '').trim().toLowerCase();
    if (key === 'templatedocid' || key === 'template_doc_id' || key === 'id шаблона') {
      return extractGoogleFileId_(values[row][1]);
    }
  }
  return '';
}

function saveTemplateId_(templateId) {
  const cleanId = extractGoogleFileId_(templateId);
  const settings = SpreadsheetApp.getActive().getSheetByName('Настройки');

  if (settings) {
    const lastRow = Math.max(settings.getLastRow(), 1);
    const values = settings.getRange(
      1,
      1,
      lastRow,
      Math.max(2, Math.min(settings.getLastColumn(), 3))
    ).getDisplayValues();

    let targetRow = 0;
    for (let row = 0; row < values.length; row++) {
      const key = String(values[row][0] || '').trim().toLowerCase();
      if (key === 'templatedocid' || key === 'template_doc_id' || key === 'id шаблона') {
        targetRow = row + 1;
        break;
      }
    }

    if (!targetRow) {
      targetRow = settings.getLastRow() + 1;
      settings.getRange(targetRow, 1).setValue('templateDocId');
      if (settings.getMaxColumns() >= 3) {
        settings.getRange(targetRow, 3).setValue('ID Google Docs шаблона квитанции');
      }
    }
    settings.getRange(targetRow, 2).setValue(cleanId);
  }

  getAvailablePropertyStores_().forEach(store => {
    try { store.setProperty('TEMPLATE_DOC_ID', cleanId); } catch (error) {}
  });
}

function extractGoogleFileId_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';
  const match = text.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : text;
}

function getReceiptRates_() {
  const result = { T1: null, T2: null, T3: null, WATER: null };
  const settings = SpreadsheetApp.getActive().getSheetByName('Настройки');
  if (!settings || settings.getLastRow() < 2) return result;

  settings.getRange(2, 1, settings.getLastRow() - 1, 2)
    .getDisplayValues()
    .forEach(row => {
      const key = normalizeSettingKey_(row[0]);
      const value = parseReceiptNumber_(row[1]);

      if (/^(tariff|тариф)t?1$/.test(key) || key === 'тариф1') result.T1 = value;
      if (/^(tariff|тариф)t?2$/.test(key) || key === 'тариф2') result.T2 = value;
      if (/^(tariff|тариф)t?3$/.test(key) || key === 'тариф3') result.T3 = value;
      if (/водоотвед|sewage|wastewater|watertariff|tariffwater/.test(key)) result.WATER = value;
    });

  return result;
}

function normalizeSettingKey_(value) {
  return String(value == null ? '' : value)
    .replace(/\u00a0/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/[\s_\-.,:;()]+/g, '');
}

function findReceiptRow_(byLabel, patterns) {
  const labels = Object.keys(byLabel);
  for (let i = 0; i < labels.length; i++) {
    if (patterns.some(pattern => pattern.test(labels[i]))) return byLabel[labels[i]];
  }
  return null;
}

function replaceReceiptText_(body, marker, value) {
  body.replaceText(escapeReceiptRegex_(marker), String(value == null ? '' : value));
}

function escapeReceiptRegex_(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normalizeReceiptLabel_(value) {
  return String(value == null ? '' : value)
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/,?\s*₽$/i, '');
}

function parseReceiptNumber_(value) {
  const text = String(value == null ? '' : value)
    .replace(/\u00a0/g, '')
    .replace(/\s/g, '')
    .replace(',', '.')
    .replace(/[^0-9.\-]/g, '');
  if (!text || text === '-' || text === '.') return null;
  const number = Number(text);
  return Number.isFinite(number) ? number : null;
}

function formatReceiptValue_(value) {
  return value === null || value === undefined || !Number.isFinite(Number(value))
    ? ''
    : Number(value).toLocaleString('ru-RU', { maximumFractionDigits: 3 });
}

function formatReceiptMoney_(value) {
  return value === null || value === undefined || !Number.isFinite(Number(value))
    ? ''
    : Number(value).toLocaleString('ru-RU', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
}

function getOrCreatePdfChildFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function getRussianMonthName_(month) {
  return [
    'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь',
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
