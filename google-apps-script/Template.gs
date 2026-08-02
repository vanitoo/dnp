/**
 * Создание и открытие Google Docs-шаблона квитанции.
 *
 * Ширины колонок динамической таблицы задаются в пунктах (pt).
 */

const DNP_PAYMENT_TABLE_WIDTHS = {
  name: 150,
  current: 66,
  previous: 72,
  usage: 58,
  rate: 58,
  amount: 82,
};

function createReceiptTemplate() {
  return createReceiptTemplateForCurrentFormat();
}

function createReceiptTemplateForCurrentFormat() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  if (!/^\d{4}$/.test(sheet.getName())) {
    ui.alert(
      'Сначала откройте лист нужного года, например 2026.\n\n' +
      'Список услуг для шаблона будет считан из текущего листа.'
    );
    return;
  }

  const services = getCurrentFormatServiceLabels_(sheet);
  if (!services.length) {
    ui.alert('В текущем листе не удалось определить строки услуг между «Тариф» и «Сумма».');
    return;
  }

  const answer = ui.alert(
    'Создать шаблон под текущий формат?',
    'Найдены строки:\n\n• ' + services.join('\n• ') +
      '\n\nБудет создан новый Google Документ и назначен активным шаблоном.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  const parent = getSpreadsheetParentFolder_();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm');
  const name = 'Шаблон квитанции ДНП Комфорт ' + timestamp;
  const doc = DocumentApp.create(name);
  const body = doc.getBody();

  body.clear();
  body.appendParagraph('ДНП «Дачный поселок «КОМФОРТ»')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .editAsText().setBold(true).setFontSize(12);

  body.appendParagraph('');
  body.appendParagraph('Наименование получателя платежа: ДНП «Дачный поселок «КОМФОРТ»');
  body.appendParagraph('ИНН {{INN}} / КПП {{KPP}}');
  body.appendParagraph('Р/с {{RS}} в {{BANK}}');
  body.appendParagraph('БИК {{BIK}}  Кор/с {{KS}}');
  body.appendParagraph('');
  body.appendParagraph('Наименование платежа: Ежемесячные начисления ДНП «Дачный поселок «Комфорт»')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');
  body.appendParagraph('Участок № {{PLOT}}    {{MONTH_NAME}} {{YEAR}} года')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Сумма оплаты: {{TOTAL}} руб.')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .editAsText().setBold(true);
  body.appendParagraph('');

  // Генератор сам вставит сюда строки Т1/Т2/Т3, водоотведения,
  // целевых взносов и других поддерживаемых услуг.
  body.appendParagraph('{{PAYMENT_TABLE}}');

  body.appendParagraph('');
  body.appendParagraph('Назначение платежа: участок № {{PLOT}}, {{MONTH_NAME}} {{YEAR}} года.');
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(parent);
  saveTemplateId_(file.getId());

  showDriveLinkDialog_(
    'Шаблон создан',
    file.getName() + '\nСтроки текущего формата: ' + services.join(', '),
    file.getUrl()
  );
}

function getCurrentFormatServiceLabels_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const labels = sheet.getRange(1, 2, lastRow, 1)
    .getDisplayValues()
    .flat()
    .map(value => String(value == null ? '' : value).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim());

  let tariffRow = -1;
  let endRow = labels.length;

  for (let index = 1; index < labels.length; index++) {
    if (/^тариф(?:ы)?$/i.test(labels[index])) {
      if (tariffRow === -1) tariffRow = index;
      else { endRow = index; break; }
    }
  }

  if (tariffRow === -1) return [];

  const result = [];
  for (let index = tariffRow + 1; index < endRow; index++) {
    const label = labels[index];
    if (!label) continue;
    if (/^(сумма|итого)/i.test(label)) break;
    if (!result.some(item => item.toLowerCase() === label.toLowerCase())) result.push(label);
  }
  return result;
}

function openReceiptTemplate() {
  const file = getReceiptTemplateFile_();
  showDriveLinkDialog_('Шаблон квитанции', file.getName(), file.getUrl());
}

function applyPaymentTableColumnWidths_(table) {
  const widths = [
    DNP_PAYMENT_TABLE_WIDTHS.name,
    DNP_PAYMENT_TABLE_WIDTHS.current,
    DNP_PAYMENT_TABLE_WIDTHS.previous,
    DNP_PAYMENT_TABLE_WIDTHS.usage,
    DNP_PAYMENT_TABLE_WIDTHS.rate,
    DNP_PAYMENT_TABLE_WIDTHS.amount,
  ];

  for (let rowIndex = 0; rowIndex < table.getNumRows(); rowIndex++) {
    const row = table.getRow(rowIndex);
    for (let columnIndex = 0; columnIndex < row.getNumCells() && columnIndex < widths.length; columnIndex++) {
      row.getCell(columnIndex).setWidth(widths[columnIndex]);
    }
  }
}
