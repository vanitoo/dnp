/**
 * Создание и открытие Google Docs-шаблона квитанции.
 *
 * Ширины колонок динамической таблицы задаются в пунктах (pt).
 * Их можно менять под свой макет: сумма должна примерно помещаться
 * в рабочую ширину страницы Google Docs.
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
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Создать шаблон квитанции?',
    'Будет создан новый Google Документ. Он станет активным шаблоном для последующего формирования PDF.',
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
  body.appendParagraph('Наименование платежа: Ежемесячный взнос на содержание ДНП «Дачный поселок «Комфорт»')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');
  body.appendParagraph('Участок № {{PLOT}}    {{MONTH_NAME}} {{YEAR}} года')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Сумма оплаты: {{TOTAL}} руб.')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .editAsText().setBold(true);
  body.appendParagraph('');

  // Маркер должен оставаться отдельным абзацем.
  body.appendParagraph('{{PAYMENT_TABLE}}');

  body.appendParagraph('');
  body.appendParagraph('Назначение платежа: участок № {{PLOT}}, {{MONTH_NAME}} {{YEAR}} года.');
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(parent);
  saveTemplateId_(file.getId());

  showDriveLinkDialog_('Шаблон создан', file.getName(), file.getUrl());
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
