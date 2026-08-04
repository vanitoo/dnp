/** Минимальный журнал формирования PDF. */

const DNP_PDF_LOG_SHEET = 'Журнал PDF';

/**
 * Основная точка запуска из интерфейса.
 * Подробная диагностика отключена: генератор работает напрямую.
 * Если DNP_PDF_LOG_ENABLED=true, записывается только одна итоговая строка.
 */
function generatePdfsForMonthWithLog(year, month) {
  const runId = Utilities.getUuid().slice(0, 8);
  const startedAt = new Date();

  try {
    const result = generatePdfsForMonth(year, month);

    if (DNP_PDF_LOG_ENABLED) {
      appendPdfLog_(
        runId,
        'ИТОГ',
        year,
        month,
        '',
        result && result.ok !== false ? 'OK' : 'ERROR',
        (result && result.message ? result.message : 'Формирование завершено') +
          '; длительность=' + Math.round((new Date().getTime() - startedAt.getTime()) / 1000) + ' сек.'
      );
    }

    return result;
  } catch (error) {
    if (DNP_PDF_LOG_ENABLED) {
      appendPdfLog_(runId, 'ОШИБКА', year, month, '', 'ERROR', getDetailedErrorText_(error));
    }
    throw error;
  }
}

function ensurePdfLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(DNP_PDF_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DNP_PDF_LOG_SHEET);
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Дата', 'Запуск', 'Этап', 'Год', 'Месяц', 'Участок', 'Статус', 'Подробности'
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function appendPdfLog_(runId, stage, year, month, plot, status, details) {
  const sheet = ensurePdfLogSheet_();
  sheet.appendRow([
    new Date(), runId || '', stage || '', year || '', month || '', plot || '', status || '', details || ''
  ]);
}

function openPdfLogSheet() {
  const sheet = ensurePdfLogSheet_();
  sheet.showSheet();
  SpreadsheetApp.getActive().setActiveSheet(sheet);
}

function clearPdfLog() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert(
    'Очистить журнал PDF?',
    'Будут удалены все строки журнала PDF, кроме заголовка.',
    ui.ButtonSet.YES_NO
  ) !== ui.Button.YES) return;

  const sheet = ensurePdfLogSheet_();
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
  }
  SpreadsheetApp.getActive().toast('Журнал PDF очищен', 'ДНП', 5);
}

function getDetailedErrorText_(error) {
  if (!error) return 'Неизвестная ошибка';
  return error.message || String(error);
}
