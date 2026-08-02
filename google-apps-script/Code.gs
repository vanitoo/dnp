/** Ядро приложения DNP Receipts. */

const DNP_VERSION = '3.7.2';
const DNP_ADMIN_PASSWORD = '123456';
const DNP_SERVICE_SHEETS = {
  settings: 'Настройки',
  emails: 'Почты',
  journal: 'Журнал отправки',
};

function showAbout() {
  const ss = SpreadsheetApp.getActive();
  const years = getYearSheetNames_();
  SpreadsheetApp.getUi().alert(
    'DNP Receipts\n\n' +
    'Версия: ' + DNP_VERSION + '\n' +
    'Google Apps Script\n\n' +
    'Таблица: ' + ss.getName() + '\n' +
    'Листы годов: ' + (years.length ? years.join(', ') : 'не найдены')
  );
}

function getYearSheetNames_() {
  return SpreadsheetApp.getActive().getSheets()
    .map(sheet => sheet.getName().trim())
    .filter(name => /^\d{4}$/.test(name))
    .sort((a, b) => Number(a) - Number(b));
}

function countPlotsForYear(year) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(String(year));
  if (!sheet) throw new Error('Лист «' + year + '» не найден.');
  return getReceiptBlocks_(sheet).length;
}

function startPdfGenerationFromDialog(year, month) {
  const result = generatePdfsForMonthWithLog(Number(year), Number(month));
  return {
    ok: result && result.ok !== false,
    message: result && result.message ? result.message : 'Формирование завершено.',
  };
}
