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
      Utilities.sleep(150);
    } catch (error) {
      failed++;
      errors.push('Участок ' + block.plot + ': ' + (error.message || error));
      if (tempFile) try { tempFile.setTrashed(true); } catch (cleanupError) {}
    }
  });

  const message = failed
    ? 'Создано PDF: ' + created + '. Ошибок: ' + failed + '. Первая ошибка: ' + errors[0]
    : 'Создано PDF: ' + created + '. Папка: ' + year + '/' + monthFolderName + '.';
  ss.toast(message, 'ДНП', 10);
  return { ok: failed === 0, created, failed, folderId: monthFolder.getId(), folderUrl: monthFolder.getUrl(), message };
}

function getReceiptBlocks_(sheet) {
  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(1, 1, lastRow, 2).getDisplayValues();
  const starts = [];
  for (let index = 1; index < rows.length; index++) {
    const plot = String(rows[index][0] || '').trim();
    const label = normalizeReceiptLabel_(rows[index][1]);
    if (/^тариф(?:ы)?$/.test(label)) {
      if (!plot) throw new Error('В строке ' + (index + 1) + ' найден «Тариф», но в столбце A нет номера участка.');
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
      formatReceiptValue_(current), formatReceiptValue_(previous), formatReceiptValue_(usage),
      formatReceiptMoney_(rate), formatReceiptMoney_(amount),
    ]);
  });

  const waterRow = findReceiptRow_(byLabel, [/^водоотвед/, /^вода$/]);
  if (waterRow) {
    const current = parseReceiptNumber_(waterRow[monthColumn - 1]);
    const previous = month > 1 ? parseReceiptNumber_(waterRow[previousMonthColumn - 1]) : null;
    const usage = current !== null && previous !== null ? current - previous : null;
    const amount = usage !== null && rates.WATER !== null ? usage * rates.WATER : null;
    if (amount !== null) calculatedTotal += amount;
    paymentRows.push(['Водоотведение', formatReceiptValue_(current), formatReceiptValue_(previous), formatReceiptValue_(usage), formatReceiptMoney_(rates.WATER), formatReceiptMoney_(amount)]);
  }

  const targetRow = findReceiptRow_(byLabel, [/^целев/]);
  if (targetRow) {
    const amount = parseReceiptNumber_(targetRow[monthColumn - 1]);
    if (amount !== null) calculatedTotal += amount;
    paymentRows.push(['Целевые взносы', '—', '—', '—', '—', formatReceiptMoney_(amount)]);
  }

  const totalRow = findReceiptRow_(byLabel, [/^сумма/, /^итого/]);
  const storedTotal = totalRow ? parseReceiptNumber_(totalRow[monthColumn - 1]) : null;
  const total = storedTotal !== null ? storedTotal : calculatedTotal;
  return { plot: block.plot, year, month, monthName: getRussianMonthName_(month), total, paymentRows };
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
  const rows = [['Наименование платежа', 'Текущее', 'Предыдущее', 'Объём', 'Тариф', 'Сумма к оплате']].concat(paymentRows);
  rows.push(['ИТОГО К ОПЛАТЕ', '', '', '', '', formatReceiptMoney_(total) + ' руб.']);

  const table = body.insertTable(index, rows);
  body.removeChild(paragraph);
  table.getRow(0).editAsText().setBold(true);
  table.getRow(table.getNumRows() - 1).editAsText().setBold(true);
  applyPaymentTableColumnWidths_(table);
}

function getReceiptTemplateFile_() {
  const templateId = getStoredTemplateId_();
  if (templateId) {
    try {
      const file = DriveApp.getFileById(templateId);
      file.getName();
      return file;
    } catch (error) {}
  }
  const parent = getSpreadsheetParentFolder_();
  const files = parent.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (/^Шаблон квитанции ДНП Комфорт/.test(file.getName())) {
      saveTemplateId_(file.getId());
      return file;
    }
  }
  throw new Error('Шаблон квитанции не найден. Выполните «ДНП → Настройка → Создать шаблон квитанции».');
}

function getStoredTemplateId_() {
  const stores = getAvailablePropertyStores_();
  for (let i = 0; i < stores.length; i++) {
    try { const value = stores[i].getProperty('TEMPLATE_DOC_ID'); if (value) return value; } catch (error) {}
  }
  return '';
}

function saveTemplateId_(templateId) {
  const stores = getAvailablePropertyStores_();
  for (let i = 0; i < stores.length; i++) {
    try { stores[i].setProperty('TEMPLATE_DOC_ID', templateId); return; } catch (error) {}
  }
}

function getReceiptRates_() {
  const result = { T1: null, T2: null, T3: null, WATER: null };
  const settings = SpreadsheetApp.getActive().getSheetByName('Настройки');
  if (!settings || settings.getLastRow() < 2) return result;
  settings.getRange(2, 1, settings.getLastRow() - 1, 2).getDisplayValues().forEach(row => {
    const key = normalizeReceiptLabel_(row[0]);
    const value = parseReceiptNumber_(row[1]);
    if (/тариф\s*т1/.test(key)) result.T1 = value;
    if (/тариф\s*т2/.test(key)) result.T2 = value;
    if (/тариф\s*т3/.test(key)) result.T3 = value;
    if (/водоотвед/.test(key)) result.WATER = value;
  });
  return result;
}

function findReceiptRow_(byLabel, patterns) {
  const labels = Object.keys(byLabel);
  for (let i = 0; i < labels.length; i++) if (patterns.some(pattern => pattern.test(labels[i]))) return byLabel[labels[i]];
  return null;
}

function replaceReceiptText_(body, marker, value) { body.replaceText(escapeReceiptRegex_(marker), String(value == null ? '' : value)); }
function escapeReceiptRegex_(text) { return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function normalizeReceiptLabel_(value) { return String(value == null ? '' : value).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase().replace(/,?\s*₽$/i, ''); }
function parseReceiptNumber_(value) {
  const text = String(value == null ? '' : value).replace(/\u00a0/g, '').replace(/\s/g, '').replace(',', '.').replace(/[^0-9.\-]/g, '');
  if (!text || text === '-' || text === '.') return null;
  const number = Number(text);
  return Number.isFinite(number) ? number : null;
}
function formatReceiptValue_(value) { return value === null || value === undefined || !Number.isFinite(Number(value)) ? '' : Number(value).toLocaleString('ru-RU', { maximumFractionDigits: 3 }); }
function formatReceiptMoney_(value) { return value === null || value === undefined || !Number.isFinite(Number(value)) ? '' : Number(value).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function getOrCreatePdfChildFolder_(parent, name) { const folders = parent.getFoldersByName(name); return folders.hasNext() ? folders.next() : parent.createFolder(name); }
function getRussianMonthName_(month) { return ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'][month - 1]; }
function sanitizePdfFileName_(value) { return String(value).trim().replace(/[\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_'); }
function trashFilesByName_(folder, fileName) { const files = folder.getFilesByName(fileName); while (files.hasNext()) try { files.next().setTrashed(true); } catch (error) {} }
