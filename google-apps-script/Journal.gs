/** Журнал операций. */

function clearJournal() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert('Очистить журнал?', 'Будут удалены все строки, кроме заголовка.', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.getActive().getSheetByName(DNP_SERVICE_SHEETS.journal);
  if (!sheet) { ui.alert('Лист журнала не найден.'); return; }
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
  }
  SpreadsheetApp.getActive().toast('Журнал очищен', 'ДНП', 5);
}

function appendJournalRow_(operation, year, month, plot, email, status, errorText) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(DNP_SERVICE_SHEETS.journal);
  if (!sheet) { ensureServiceSheets_(); sheet = SpreadsheetApp.getActive().getSheetByName(DNP_SERVICE_SHEETS.journal); }
  sheet.appendRow([new Date(), operation || '', year || '', month || '', plot || '', email || '', status || '', errorText || '']);
}
