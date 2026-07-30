/**
 * Исправления создания нового листа-года и чередования блоков.
 *
 * 1. При копировании сохраняются оформление, размеры, объединения,
 *    проверки данных и формулы.
 * 2. Очищаются только введённые вручную значения в строках данных:
 *    Т1, Т2, Т3, кВтч, Сумма.
 * 3. Строки «Тариф», номера месяцев и формулы сохраняются.
 * 4. Чередование цвета определяется строками «Тариф» в столбце B,
 *    поэтому код не зависит от фиксированного размера блока в 7 строк.
 */

function clearCopiedYearValues_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 3) return;

  const labels = sheet
    .getRange(1, 2, lastRow, 1)
    .getDisplayValues()
    .flat()
    .map(value => normalizeDnpRowLabel_(value));

  const dataRowPattern = /^(т1|т2|т3|квтч|квт·ч|квт\/ч|сумма)$/i;

  for (let row = 2; row <= lastRow; row++) {
    if (!dataRowPattern.test(labels[row - 1])) continue;

    const range = sheet.getRange(row, 3, 1, lastColumn - 2);
    const values = range.getValues()[0];
    const formulas = range.getFormulas()[0];

    let changed = false;
    for (let column = 0; column < values.length; column++) {
      // Формулы нового года сохраняем. Удаляем только введённые значения.
      if (!formulas[column] && values[column] !== '') {
        values[column] = '';
        changed = true;
      }
    }

    if (changed) range.setValues([values]);
  }
}

function normalizeDnpRowLabel_(value) {
  return String(value == null ? '' : value)
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function applySevenRowBandingToYearSheets_() {
  getYearSheetNames_().forEach(year => {
    const sheet = SpreadsheetApp.getActive().getSheetByName(year);
    if (sheet) applySevenRowBanding_(sheet);
  });
}

function applySevenRowBanding_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  if (lastRow < 2) return;

  const labels = sheet
    .getRange(1, 2, lastRow, 1)
    .getDisplayValues()
    .flat()
    .map(value => normalizeDnpRowLabel_(value));

  const tariffRows = [];
  for (let index = 1; index < labels.length; index++) {
    if (/^тариф(?:ы)?$/.test(labels[index])) {
      tariffRows.push(index + 1);
    }
  }

  const colors = ['#ffffff', '#f3f7fd'];

  // Основной режим: блок определяется от строки «Тариф» до строки
  // перед следующим «Тариф». Первая строка таблицы остаётся шапкой.
  if (tariffRows.length) {
    tariffRows.forEach((startRow, blockIndex) => {
      const endRow = blockIndex + 1 < tariffRows.length
        ? tariffRows[blockIndex + 1] - 1
        : lastRow;
      const rowCount = endRow - startRow + 1;

      if (rowCount > 0) {
        sheet
          .getRange(startRow, 1, rowCount, lastColumn)
          .setBackground(colors[blockIndex % colors.length]);
      }
    });
    return;
  }

  // Резервный режим для старых таблиц без меток «Тариф»:
  // чередуем блоки по 7 строк, начиная со строки 2.
  for (let row = 2, blockIndex = 0; row <= lastRow; row += 7, blockIndex++) {
    const rowCount = Math.min(7, lastRow - row + 1);
    sheet
      .getRange(row, 1, rowCount, lastColumn)
      .setBackground(colors[blockIndex % colors.length]);
  }
}
