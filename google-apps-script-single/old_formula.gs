function fillSumFormulasFromTariffs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const year = Number(sheet.getName());

  if (!Number.isInteger(year) || year < 2000 || year > 2100) {
    throw new Error('Откройте лист нужного года, например 2026.');
  }

  const TARIFF_SHEET_NAME = 'Тарифы';
  const LABEL_COL = 2; // B
  const JAN_COL = 3;   // C
  const DEC_COL = 14;  // N

  const tariffSheet = ss.getSheetByName(TARIFF_SHEET_NAME);

  if (!tariffSheet) {
    throw new Error('Не найден лист «Тарифы».');
  }

  // ----------------------------------------------------------
  // Читаем историю тарифов
  // A = год
  // B = месяц начала действия
  // C = T1
  // D = T2
  // E = T3
  // ----------------------------------------------------------

  const tariffLastRow = tariffSheet.getLastRow();

  if (tariffLastRow < 2) {
    throw new Error('На листе «Тарифы» нет тарифов.');
  }

  const tariffValues = tariffSheet
    .getRange(2, 1, tariffLastRow - 1, 5)
    .getValues();

  const tariffs = [];

  tariffValues.forEach(function(row, index) {
    const tariffYear = Number(row[0]);
    const tariffMonth = Number(row[1]);

    if (
      !Number.isInteger(tariffYear) ||
      !Number.isInteger(tariffMonth) ||
      tariffMonth < 1 ||
      tariffMonth > 12
    ) {
      return;
    }

    const tariffRow = index + 2;

    tariffs.push({
      year: tariffYear,
      month: tariffMonth,
      key: tariffYear * 100 + tariffMonth,

      t1: "'Тарифы'!$C$" + tariffRow,
      t2: "'Тарифы'!$D$" + tariffRow,
      t3: "'Тарифы'!$E$" + tariffRow
    });
  });

  tariffs.sort(function(a, b) {
    return a.key - b.key;
  });

  function getTariffForMonth_(targetYear, targetMonth) {
    const targetKey = targetYear * 100 + targetMonth;
    let result = null;

    for (let i = 0; i < tariffs.length; i++) {
      if (tariffs[i].key <= targetKey) {
        result = tariffs[i];
      } else {
        break;
      }
    }

    return result;
  }

  // ----------------------------------------------------------
  // Читаем текущий лист
  // ----------------------------------------------------------

  const lastRow = sheet.getLastRow();

  const labels = sheet
    .getRange(1, LABEL_COL, lastRow, 1)
    .getDisplayValues()
    .flat()
    .map(function(value) {
      return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '')
        .replace(/,?₽$/g, '');
    });

  const previousYear = year - 1;
  const previousSheet = ss.getSheetByName(String(previousYear));

  let blocksProcessed = 0;
  let formulasCreated = 0;
  let januarySkipped = 0;

  // ----------------------------------------------------------
  // Ищем блоки Т1 / Т2 / Т3
  // ----------------------------------------------------------

  for (let i = 0; i < labels.length; i++) {

    if (labels[i] !== 'т1') {
      continue;
    }

    if (
      labels[i + 1] !== 'т2' ||
      labels[i + 2] !== 'т3'
    ) {
      continue;
    }

    const t1Row = i + 1;
    const t2Row = t1Row + 1;
    const t3Row = t1Row + 2;

    // Ищем строку Сумма
    let sumRow = 0;

    for (
      let searchRow = t3Row + 1;
      searchRow <= Math.min(t3Row + 15, lastRow);
      searchRow++
    ) {

      const label = labels[searchRow - 1] || '';

      if (
        label.indexOf('сумма') === 0 ||
        label.indexOf('итого') === 0
      ) {
        sumRow = searchRow;
        break;
      }

      if (label === 'тариф') {
        break;
      }
    }

    if (!sumRow) {
      continue;
    }

    blocksProcessed++;

    // ========================================================
    // ФЕВРАЛЬ - ДЕКАБРЬ
    // ========================================================

    for (let month = 2; month <= 12; month++) {

      const tariff = getTariffForMonth_(year, month);

      if (!tariff) {
        continue;
      }

      const currentCol = month + 2;
      const previousCol = currentCol - 1;

      const currentT1 =
        sheet.getRange(t1Row, currentCol).getA1Notation();

      const currentT2 =
        sheet.getRange(t2Row, currentCol).getA1Notation();

      const currentT3 =
        sheet.getRange(t3Row, currentCol).getA1Notation();

      const previousT1 =
        sheet.getRange(t1Row, previousCol).getA1Notation();

      const previousT2 =
        sheet.getRange(t2Row, previousCol).getA1Notation();

      const previousT3 =
        sheet.getRange(t3Row, previousCol).getA1Notation();

      const currentRange =
        sheet
          .getRange(t1Row, currentCol, 3, 1)
          .getA1Notation();

      /*
       * Если текущий месяц полностью пустой — ничего не показываем.
       *
       * Если заполнен хотя бы Т1/Т2/Т3:
       *
       * (T1 текущий - T1 предыдущий) × тариф T1
       * +
       * (T2 текущий - T2 предыдущий) × тариф T2
       * +
       * (T3 текущий - T3 предыдущий) × тариф T3
       *
       * Пустые значения Google Sheets при арифметике считает нулями.
       */

      const formula =
        '=IF(' +
          'COUNT(' + currentRange + ')=0;' +
          '"";' +

          '(' + currentT1 + '-' + previousT1 + ')*' + tariff.t1 +
          '+' +
          '(' + currentT2 + '-' + previousT2 + ')*' + tariff.t2 +
          '+' +
          '(' + currentT3 + '-' + previousT3 + ')*' + tariff.t3 +
        ')';

      sheet
        .getRange(sumRow, currentCol)
        .setFormula(formula);

      formulasCreated++;
    }

    // ========================================================
    // ЯНВАРЬ
    // ========================================================

    const januaryTariff =
      getTariffForMonth_(year, 1);

    if (januaryTariff && previousSheet) {

      const currentT1 =
        sheet.getRange(t1Row, JAN_COL).getA1Notation();

      const currentT2 =
        sheet.getRange(t2Row, JAN_COL).getA1Notation();

      const currentT3 =
        sheet.getRange(t3Row, JAN_COL).getA1Notation();

      const currentRange =
        sheet
          .getRange(t1Row, JAN_COL, 3, 1)
          .getA1Notation();

      const previousT1 =
        "'" + previousYear + "'!" +
        previousSheet
          .getRange(t1Row, DEC_COL)
          .getA1Notation();

      const previousT2 =
        "'" + previousYear + "'!" +
        previousSheet
          .getRange(t2Row, DEC_COL)
          .getA1Notation();

      const previousT3 =
        "'" + previousYear + "'!" +
        previousSheet
          .getRange(t3Row, DEC_COL)
          .getA1Notation();

      const januaryFormula =
        '=IF(' +
          'COUNT(' + currentRange + ')=0;' +
          '"";' +

          '(' + currentT1 + '-' + previousT1 + ')*' + januaryTariff.t1 +
          '+' +
          '(' + currentT2 + '-' + previousT2 + ')*' + januaryTariff.t2 +
          '+' +
          '(' + currentT3 + '-' + previousT3 + ')*' + januaryTariff.t3 +
        ')';

      sheet
        .getRange(sumRow, JAN_COL)
        .setFormula(januaryFormula);

      formulasCreated++;

    } else {

      januarySkipped++;
    }
  }

  SpreadsheetApp.flush();

  let message =
    'Обработано участков: ' + blocksProcessed +
    '\nУстановлено формул: ' + formulasCreated;

  if (januarySkipped) {
    message +=
      '\nЯнварь пропущен: ' + januarySkipped;
  }

  ss.toast(
    message,
    'Формулы суммы установлены',
    10
  );
}