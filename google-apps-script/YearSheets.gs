/** Создание нового года, очистка вводимых значений и окраска блоков. */

function clearCopiedYearValues_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 3) return;

  const labels = sheet.getRange(1, 2, lastRow, 1)
    .getDisplayValues().flat().map(normalizeDnpRowLabel_);

  const dataRowPattern = /^(т1|т2|т3|квтч|квт·ч|квт\/ч|водоотведение|целевые взносы|сумма|сумма,?\s*₽|сумма,?\s*р)$/i;

  for (let row = 2; row <= lastRow; row++) {
    if (!dataRowPattern.test(labels[row - 1])) continue;

    const range = sheet.getRange(row, 3, 1, lastColumn - 2);
    const values = range.getValues()[0];
    const formulas = range.getFormulas()[0];
    let changed = false;

    for (let column = 0; column < values.length; column++) {
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

  const labels = sheet.getRange(1, 2, lastRow, 1)
    .getDisplayValues().flat().map(normalizeDnpRowLabel_);
  const tariffRows = [];

  for (let index = 1; index < labels.length; index++) {
    if (/^тариф(?:ы)?$/.test(labels[index])) tariffRows.push(index + 1);
  }

  const colors = ['#ffffff', '#f3f7fd'];

  if (tariffRows.length) {
    tariffRows.forEach((startRow, blockIndex) => {
      const endRow = blockIndex + 1 < tariffRows.length
        ? tariffRows[blockIndex + 1] - 1
        : lastRow;
      const rowCount = endRow - startRow + 1;
      if (rowCount > 0) {
        sheet.getRange(startRow, 1, rowCount, lastColumn)
          .setBackground(colors[blockIndex % colors.length]);
      }
    });
    return;
  }

  for (let row = 2, blockIndex = 0; row <= lastRow; row += 7, blockIndex++) {
    const rowCount = Math.min(7, lastRow - row + 1);
    sheet.getRange(row, 1, rowCount, lastColumn)
      .setBackground(colors[blockIndex % colors.length]);
  }
}
