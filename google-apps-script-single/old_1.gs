function formatYearSheetUX() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // Проверяем, что открыт лист-года
  const year = Number(sheet.getName());

  if (!Number.isInteger(year) || year < 2000 || year > 2100) {
    SpreadsheetApp.getUi().alert(
      'Откройте лист года, например 2026.'
    );
    return;
  }

  const LABEL_COL = 2;       // B
  const FIRST_MONTH_COL = 3; // C
  const MONTH_COUNT = 12;    // C:N

  const lastRow = sheet.getLastRow();

  const labels = sheet
    .getRange(1, LABEL_COL, lastRow, 1)
    .getDisplayValues()
    .flat();

  let formatted = 0;

  labels.forEach(function(value, index) {

    const row = index + 1;

    const label = String(value || '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '');

    // Диапазон месяцев C:N
    const monthsRange = sheet.getRange(
      row,
      FIRST_MONTH_COL,
      1,
      MONTH_COUNT
    );

    // Название строки B
    const labelCell = sheet.getRange(row, LABEL_COL);

    // ==========================================
    // Т1 / Т2 / Т3
    // ручной ввод
    // ==========================================

    if (
      label === 'т1' ||
      label === 'т2' ||
      label === 'т3'
    ) {
      monthsRange
        .setBackground('#ffffff')
        .setFontWeight('normal');

      formatted++;
      return;
    }

    // ==========================================
    // кВтч
    // автоматический расчёт
    // ==========================================

    if (label === 'квтч') {
      monthsRange
        .setBackground('#eeeeee')
        .setFontWeight('bold');

      labelCell
        .setBackground('#eeeeee')
        .setFontWeight('bold');

      formatted++;
      return;
    }

    // ==========================================
    // Целевые взносы
    // ручной ввод
    // ==========================================

    if (
      label === 'целевыевзносы' ||
      label === 'целевойвзнос'
    ) {
      monthsRange
        .setBackground('#fff2cc')
        .setFontWeight('normal');

      labelCell
        .setBackground('#fff2cc');

      formatted++;
      return;
    }

    // ==========================================
    // Сумма
    // автоматический расчёт
    // ==========================================

    if (
      label === 'сумма,₽' ||
      label === 'сумма,р' ||
      label === 'сумма,руб.' ||
      label === 'сумма'
    ) {
      monthsRange
        .setBackground('#eeeeee')
        .setFontWeight('bold');

      labelCell
        .setBackground('#eeeeee')
        .setFontWeight('bold');

      formatted++;
    }
  });

  SpreadsheetApp.flush();

  ss.toast(
    'Оформление применено.\nОбработано строк: ' + formatted,
    'ДНП',
    6
  );
}
