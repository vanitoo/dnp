/** Работа с листом почты и последовательная отправка квитанций. */

function syncEmailSheetPlots_() {
  const ss = SpreadsheetApp.getActive();
  const emailSheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails);
  const years = getYearSheetNames_();
  if (!emailSheet || !years.length) return;

  ensureEmailSheetHeaders_(emailSheet);
  const source = ss.getSheetByName(years[years.length - 1]);
  const blocks = getReceiptBlocks_(source);
  const plots = blocks.map(block => String(block.plot).trim());
  const existing = emailSheet.getLastRow() > 1
    ? emailSheet.getRange(2, 1, emailSheet.getLastRow() - 1, 4).getValues()
    : [];
  const byPlot = new Map(existing.map(row => [String(row[0]).trim(), row]));
  const rows = plots.map(plot => byPlot.get(plot) || [plot, '', '', false]);

  if (emailSheet.getLastRow() > 1) {
    emailSheet.getRange(2, 1, emailSheet.getLastRow() - 1, 4).clearContent();
  }
  if (rows.length) emailSheet.getRange(2, 1, rows.length, 4).setValues(rows);
  if (rows.length) emailSheet.getRange(2, 4, rows.length, 1).insertCheckboxes();
}

function ensureEmailSheetHeaders_(sheet) {
  sheet.getRange('A1:D1').setValues([['Участок', 'Email', 'ФИО', 'Отправить']]);
  sheet.setFrozenRows(1);
  if (sheet.getMaxRows() > 1) sheet.getRange(2, 4, sheet.getMaxRows() - 1, 1).insertCheckboxes();
}

function fillTestEmails() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails);
  if (!sheet) { ensureServiceSheets_(); sheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails); }
  ensureEmailSheetHeaders_(sheet);
  if (sheet.getLastRow() < 2) syncEmailSheetPlots_();

  const count = Math.max(sheet.getLastRow() - 1, 0);
  if (!count) { SpreadsheetApp.getUi().alert('На листе «Почты» нет участков.'); return; }

  const values = sheet.getRange(2, 1, count, 4).getValues();
  let filled = 0;
  values.forEach(row => {
    if (row[0] && !row[1]) {
      row[1] = String(row[0]).replace(/[^0-9A-Za-zА-Яа-я_-]+/g, '_') + '@mail.ru';
      filled++;
    }
  });
  sheet.getRange(2, 1, count, 4).setValues(values);
  sheet.getRange(2, 4, count, 1).insertCheckboxes();
  ss.toast('Добавлено тестовых адресов: ' + filled, 'ДНП', 5);
}

function sendReceipts() {
  const years = getYearSheetNames_();
  if (!years.length) throw new Error('Не найдены листы с названиями годов.');

  const now = new Date();
  const currentYear = String(now.getFullYear());
  const currentMonth = now.getMonth() + 1;
  const defaultYear = years.includes(currentYear) ? currentYear : years[years.length - 1];
  const yearOptions = years.map(year =>
    '<option value="' + year + '"' + (year === defaultYear ? ' selected' : '') + '>' + year + '</option>'
  ).join('');

  const html = HtmlService.createHtmlOutput(`
<!doctype html><html><head><base target="_top"><style>
body{font:14px Arial,sans-serif;padding:18px;color:#202124}h2{margin:0 0 16px;font-size:18px}
label{display:block;margin:12px 0 6px;font-weight:600}select{width:100%;box-sizing:border-box;padding:9px;border:1px solid #dadce0;border-radius:6px}
.note{background:#f8f9fa;padding:10px;border-radius:6px;margin-top:14px;line-height:1.4}.buttons{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}
button{padding:9px 14px;border:0;border-radius:6px;cursor:pointer}.primary{background:#1a73e8;color:#fff}.secondary{background:#f1f3f4}#status{margin-top:12px;min-height:36px;color:#5f6368}
</style></head><body><h2>Отправка квитанций</h2>
<label>Год</label><select id="year">${yearOptions}</select>
<label>Месяц</label><select id="month">
<option value="1">Январь</option><option value="2">Февраль</option><option value="3">Март</option><option value="4">Апрель</option>
<option value="5">Май</option><option value="6">Июнь</option><option value="7">Июль</option><option value="8">Август</option>
<option value="9">Сентябрь</option><option value="10">Октябрь</option><option value="11">Ноябрь</option><option value="12">Декабрь</option></select>
<div class="note">Будут отправлены только строки листа «Почты», где заполнен Email и установлена галка «Отправить».</div>
<div id="status"></div><div class="buttons"><button class="secondary" onclick="google.script.host.close()">Отмена</button><button class="primary" id="send" onclick="run()">Отправить</button></div>
<script>
document.getElementById('month').value='${currentMonth}';
function run(){const b=document.getElementById('send'),s=document.getElementById('status');b.disabled=true;s.textContent='Идёт последовательная отправка…';google.script.run.withSuccessHandler(r=>{s.textContent=r.message;b.disabled=false}).withFailureHandler(e=>{s.textContent='Ошибка: '+e.message;b.disabled=false}).sendReceiptsForMonth(Number(document.getElementById('year').value),Number(document.getElementById('month').value))}
</script></body></html>`).setWidth(460).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'ДНП');
}

function sendReceiptsForMonth(year, month) {
  year = Number(year);
  month = Number(month);
  if (!Number.isInteger(year) || year < 2000 || year > 2100) throw new Error('Некорректный год.');
  if (!Number.isInteger(month) || month < 1 || month > 12) throw new Error('Некорректный месяц.');

  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails);
  if (!sheet) { ensureServiceSheets_(); sheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails); }
  ensureEmailSheetHeaders_(sheet);

  const root = getDnpPdfFolder_();
  const yearFolder = findChildFolderByNames_(root, [String(year)]);
  if (!yearFolder) throw new Error('Папка года «' + year + '» не найдена. Сначала сформируйте PDF.');
  const monthFolderName = String(month).padStart(2, '0') + ' ' + getRussianMonthName_(month);
  const monthFolder = findChildFolderByNames_(yearFolder, [monthFolderName]);
  if (!monthFolder) throw new Error('Папка «' + monthFolderName + '» не найдена. Сначала сформируйте PDF.');

  const count = Math.max(sheet.getLastRow() - 1, 0);
  if (!count) throw new Error('На листе «Почты» нет получателей.');
  const rows = sheet.getRange(2, 1, count, 4).getValues();
  const subjectTemplate = getMailSetting_('emailSubject') || 'Квитанция ДНП Комфорт: участок {{PLOT}}, {{MONTH_NAME}} {{YEAR}}';
  const bodyTemplate = getMailSetting_('emailBody') || 'Здравствуйте{{NAME_PART}}!\n\nНаправляем квитанцию по участку № {{PLOT}} за {{MONTH_NAME}} {{YEAR}} года.\n\nС уважением, ДНП «Комфорт».';

  let selected = 0;
  let sent = 0;
  let failed = 0;
  let skipped = 0;
  const errors = [];

  rows.forEach((row, index) => {
    const plot = String(row[0] == null ? '' : row[0]).trim();
    const email = String(row[1] == null ? '' : row[1]).trim();
    const name = String(row[2] == null ? '' : row[2]).trim();
    const shouldSend = isSendFlagEnabled_(row[3]);
    if (!shouldSend) return;
    selected++;

    if (!plot || !email) {
      skipped++;
      appendJournalRow_('EMAIL', year, month, plot, email, 'SKIP', 'Не заполнен участок или Email');
      return;
    }

    const fileName = 'Квитанция_участок_' + sanitizePdfFileName_(plot) + '_' + year + '_' + String(month).padStart(2, '0') + '.pdf';
    const files = monthFolder.getFilesByName(fileName);
    if (!files.hasNext()) {
      failed++;
      const text = 'Не найден PDF: ' + fileName;
      errors.push('Участок ' + plot + ': ' + text);
      appendJournalRow_('EMAIL', year, month, plot, email, 'ERROR', text);
      return;
    }

    try {
      const pdf = files.next();
      const replacements = {
        '{{PLOT}}': plot,
        '{{YEAR}}': String(year),
        '{{MONTH}}': String(month).padStart(2, '0'),
        '{{MONTH_NAME}}': getRussianMonthName_(month),
        '{{FIO}}': name,
        '{{NAME_PART}}': name ? ', ' + name : '',
      };
      const subject = replaceMailMarkers_(subjectTemplate, replacements);
      const body = replaceMailMarkers_(bodyTemplate, replacements);

      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body,
        attachments: [pdf.getBlob().setName(fileName)],
        name: 'ДНП Комфорт',
      });
      sent++;
      appendJournalRow_('EMAIL', year, month, plot, email, 'SENT', '');
      ss.toast('Отправлено ' + sent + ' из ' + selected + ': участок ' + plot, 'ДНП', 4);
      Utilities.sleep(100);
    } catch (error) {
      failed++;
      const text = error.message || String(error);
      errors.push('Участок ' + plot + ': ' + text);
      appendJournalRow_('EMAIL', year, month, plot, email, 'ERROR', text);
    }
  });

  if (!selected) throw new Error('Не выбрано ни одной строки. Установите галку в столбце «Отправить».');
  const message = 'Выбрано: ' + selected + '. Отправлено: ' + sent + '. Ошибок: ' + failed + '. Пропущено: ' + skipped + '.' +
    (errors.length ? ' Первая ошибка: ' + errors[0] : '');
  ss.toast(message, 'ДНП', 10);
  return { ok: failed === 0, selected, sent, failed, skipped, message };
}

function isSendFlagEnabled_(value) {
  if (value === true || value === 1) return true;
  const text = String(value == null ? '' : value).trim().toLowerCase();
  return ['да', 'yes', 'true', '1', 'отправить', 'x', '+'].includes(text);
}

function getMailSetting_(key) {
  const settings = SpreadsheetApp.getActive().getSheetByName(DNP_SERVICE_SHEETS.settings);
  if (!settings || settings.getLastRow() < 1) return '';
  const wanted = normalizeSettingKey_(key);
  const rows = settings.getRange(1, 1, settings.getLastRow(), Math.max(2, settings.getLastColumn())).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    if (normalizeSettingKey_(rows[i][0]) === wanted) return String(rows[i][1] || '').trim();
  }
  return '';
}

function replaceMailMarkers_(text, replacements) {
  let result = String(text == null ? '' : text);
  Object.keys(replacements).forEach(marker => {
    result = result.split(marker).join(replacements[marker]);
  });
  return result;
}
