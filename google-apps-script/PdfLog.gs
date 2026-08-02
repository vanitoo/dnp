/** Подробный журнал формирования PDF и диагностический генератор. */

const DNP_PDF_LOG_SHEET = 'Журнал PDF';

function generatePdfsForMonthWithLog(year, month) {
  year = Number(year);
  month = Number(month);
  const runId = Utilities.getUuid().slice(0, 8);
  const ss = SpreadsheetApp.getActive();

  ensurePdfLogSheet_();
  appendPdfLog_(runId, 'СТАРТ', year, month, '', 'INFO', 'Запуск формирования PDF');

  let templateFile;
  let monthFolder;
  try {
    if (!Number.isInteger(year) || year < 2000 || year > 2100) throw new Error('Некорректный год.');
    if (!Number.isInteger(month) || month < 1 || month > 12) throw new Error('Некорректный месяц.');

    const sheet = ss.getSheetByName(String(year));
    if (!sheet) throw new Error('Лист «' + year + '» не найден.');
    appendPdfLog_(runId, 'ЛИСТ', year, month, '', 'OK', 'Найден лист «' + year + '»');

    const templateId = getStoredTemplateId_();
    appendPdfLog_(runId, 'НАСТРОЙКИ', year, month, '', templateId ? 'OK' : 'ERROR', 'templateDocId=' + (templateId || '(пусто)'));

    templateFile = getReceiptTemplateFile_();
    appendPdfLog_(runId, 'ШАБЛОН', year, month, '', 'OK', templateFile.getName() + ' | ID=' + templateFile.getId());

    validateReceiptTemplateForLog_(templateFile, runId, year, month);

    const rootFolder = getDnpPdfFolder_();
    const yearFolder = getOrCreatePdfChildFolder_(rootFolder, String(year));
    const monthFolderName = String(month).padStart(2, '0') + ' ' + getRussianMonthName_(month);
    monthFolder = getOrCreatePdfChildFolder_(yearFolder, monthFolderName);
    appendPdfLog_(runId, 'ПАПКА', year, month, '', 'OK', monthFolder.getName() + ' | ID=' + monthFolder.getId());

    const blocks = getReceiptBlocks_(sheet);
    if (!blocks.length) throw new Error('Не найдены строки «Тариф» в столбце B.');
    appendPdfLog_(runId, 'УЧАСТКИ', year, month, '', 'OK', 'Найдено блоков: ' + blocks.length);

    const rates = getReceiptRates_();
    appendPdfLog_(runId, 'ТАРИФЫ', year, month, '', 'INFO', JSON.stringify(rates));

    let created = 0;
    let failed = 0;
    const errors = [];

    blocks.forEach((block, index) => {
      let tempFile = null;
      let stage = 'ДАННЫЕ';
      ss.toast('Формируется ' + (index + 1) + ' из ' + blocks.length + ': участок ' + block.plot, 'ДНП', 5);
      appendPdfLog_(runId, 'УЧАСТОК', year, month, block.plot, 'START', 'Строки ' + block.startRow + '–' + block.endRow);

      try {
        const receipt = buildReceiptData_(sheet, block, year, month, rates);
        appendPdfLog_(runId, 'ДАННЫЕ', year, month, block.plot, 'OK', 'Строк начислений: ' + receipt.paymentRows.length + '; итог=' + receipt.total);

        const fileName = 'Квитанция_участок_' + sanitizePdfFileName_(block.plot) + '_' + year + '_' + String(month).padStart(2, '0') + '.pdf';
        trashFilesByName_(monthFolder, fileName);

        stage = 'КОПИЯ ШАБЛОНА';
        tempFile = templateFile.makeCopy('Временная квитанция ' + block.plot, monthFolder);
        appendPdfLog_(runId, stage, year, month, block.plot, 'OK', 'ID=' + tempFile.getId());

        stage = 'ОТКРЫТИЕ DOC';
        const doc = DocumentApp.openById(tempFile.getId());
        appendPdfLog_(runId, stage, year, month, block.plot, 'OK', 'Документ открыт');

        stage = 'ЗАПОЛНЕНИЕ';
        fillReceiptTemplate_(doc, receipt);
        appendPdfLog_(runId, stage, year, month, block.plot, 'OK', 'Метки заменены, таблица вставлена');

        stage = 'СОХРАНЕНИЕ DOC';
        doc.saveAndClose();
        appendPdfLog_(runId, stage, year, month, block.plot, 'OK', 'Документ сохранён');

        stage = 'КОНВЕРТАЦИЯ PDF';
        monthFolder.createFile(tempFile.getAs(MimeType.PDF).setName(fileName));
        appendPdfLog_(runId, stage, year, month, block.plot, 'OK', fileName);

        tempFile.setTrashed(true);
        tempFile = null;
        created++;
        Utilities.sleep(150);
      } catch (error) {
        failed++;
        const errorText = getDetailedErrorText_(error);
        errors.push('Участок ' + block.plot + ', этап «' + stage + '»: ' + errorText);
        appendPdfLog_(runId, stage, year, month, block.plot, 'ERROR', errorText);
        if (tempFile) {
          try { tempFile.setTrashed(true); } catch (cleanupError) {
            appendPdfLog_(runId, 'ОЧИСТКА', year, month, block.plot, 'ERROR', getDetailedErrorText_(cleanupError));
          }
        }
      }
    });

    const message = failed
      ? 'Создано PDF: ' + created + '. Ошибок: ' + failed + '. Первая ошибка: ' + errors[0] + '. Подробности: лист «' + DNP_PDF_LOG_SHEET + '», запуск ' + runId + '.'
      : 'Создано PDF: ' + created + '. Подробности: лист «' + DNP_PDF_LOG_SHEET + '», запуск ' + runId + '.';

    appendPdfLog_(runId, 'ИТОГ', year, month, '', failed ? 'ERROR' : 'OK', 'Создано=' + created + '; ошибок=' + failed);
    ss.toast(message, 'ДНП', 10);
    return { ok: failed === 0, created, failed, folderId: monthFolder.getId(), folderUrl: monthFolder.getUrl(), runId, message };
  } catch (error) {
    const errorText = getDetailedErrorText_(error);
    appendPdfLog_(runId, 'КРИТИЧЕСКАЯ ОШИБКА', year, month, '', 'ERROR', errorText);
    throw new Error(errorText + '\n\nПодробности записаны на лист «' + DNP_PDF_LOG_SHEET + '», запуск ' + runId + '.');
  }
}

function validateReceiptTemplateForLog_(templateFile, runId, year, month) {
  let testCopy = null;
  try {
    testCopy = templateFile.makeCopy('Проверка шаблона ' + runId);
    const doc = DocumentApp.openById(testCopy.getId());
    const body = doc.getBody();
    const marker = body.findText('\\{\\{PAYMENT_TABLE\\}\\}');

    if (!marker) {
      throw new Error('Маркер {{PAYMENT_TABLE}} не найден в основном тексте документа. Проверьте, что он находится не в колонтитуле, рисунке или текстовом поле.');
    }

    const textElement = marker.getElement();
    const parent = textElement.getParent();
    const parentType = String(parent.getType());
    const grandParent = parent.getParent();
    const grandParentType = grandParent ? String(grandParent.getType()) : '(нет)';

    appendPdfLog_(runId, 'ПРОВЕРКА МАРКЕРА', year, month, '', 'OK', 'Родитель=' + parentType + '; выше=' + grandParentType + '; текст=' + textElement.asText().getText());

    if (parentType !== String(DocumentApp.ElementType.PARAGRAPH)) {
      throw new Error('Маркер найден, но его родитель — ' + parentType + ', а нужен отдельный абзац.');
    }

    if (grandParentType !== String(DocumentApp.ElementType.BODY_SECTION)) {
      throw new Error('Маркер находится внутри элемента ' + grandParentType + '. Перенесите его в обычный текст документа вне таблиц и колонтитулов.');
    }

    doc.saveAndClose();
  } finally {
    if (testCopy) {
      try { testCopy.setTrashed(true); } catch (error) {}
    }
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
  if (ui.alert('Очистить журнал PDF?', 'Будут удалены все строки журнала PDF, кроме заголовка.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const sheet = ensurePdfLogSheet_();
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
  SpreadsheetApp.getActive().toast('Журнал PDF очищен', 'ДНП', 5);
}

function getDetailedErrorText_(error) {
  if (!error) return 'Неизвестная ошибка';
  const message = error.message || String(error);
  const stack = error.stack ? String(error.stack).replace(/\n/g, ' | ') : '';
  return stack ? message + ' | ' + stack : message;
}
