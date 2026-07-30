/** Работа с листом почты. */

function syncEmailSheetPlots_() {
  const ss = SpreadsheetApp.getActive();
  const emailSheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails);
  const years = getYearSheetNames_();
  if (!emailSheet || !years.length) return;

  const source = ss.getSheetByName(years[years.length - 1]);
  const blocks = getReceiptBlocks_(source);
  const plots = blocks.map(block => String(block.plot).trim());
  const existing = emailSheet.getLastRow() > 1
    ? emailSheet.getRange(2, 1, emailSheet.getLastRow() - 1, 4).getValues()
    : [];
  const byPlot = new Map(existing.map(row => [String(row[0]).trim(), row]));
  const rows = plots.map(plot => byPlot.get(plot) || [plot, '', '', '']);

  if (emailSheet.getLastRow() > 1) {
    emailSheet.getRange(2, 1, emailSheet.getLastRow() - 1, 4).clearContent();
  }
  if (rows.length) emailSheet.getRange(2, 1, rows.length, 4).setValues(rows);
}

function fillTestEmails() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails);
  if (!sheet) { ensureServiceSheets_(); sheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails); }
  if (sheet.getLastRow() < 2) syncEmailSheetPlots_();

  const count = Math.max(sheet.getLastRow() - 1, 0);
  if (!count) { SpreadsheetApp.getUi().alert('На листе «Почты» нет участков.'); return; }

  const values = sheet.getRange(2, 1, count, 2).getValues();
  let filled = 0;
  values.forEach(row => {
    if (row[0] && !row[1]) {
      row[1] = String(row[0]).replace(/[^0-9A-Za-zА-Яа-я_-]+/g, '_') + '@mail.ru';
      filled++;
    }
  });
  sheet.getRange(2, 1, count, 2).setValues(values);
  ss.toast('Добавлено тестовых адресов: ' + filled, 'ДНП', 5);
}

function sendReceipts() {
  SpreadsheetApp.getUi().alert('Функция отправки квитанций будет подключена после проверки формирования PDF по новому шаблону.');
}
