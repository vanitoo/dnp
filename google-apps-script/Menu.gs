/** Меню и основные диалоги. */

function onOpen() {
  removeLastPdfSetting_();
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('ДНП')
    .addSubMenu(ui.createMenu('Квитанции')
      .addItem('Сформировать PDF', 'showPdfDialog')
      .addItem('Открыть папку месяца', 'openCurrentMonthFolder')
      .addItem('Очистить PDF', 'clearGeneratedPdfs'))
    .addSubMenu(ui.createMenu('Почта')
      .addItem('Отправить квитанции', 'sendReceipts'))
    .addSubMenu(ui.createMenu('Настройка')
      .addItem('Первичная настройка', 'showInitialSetupDialog')
      .addSeparator()
      .addItem('Создать шаблон квитанции', 'createReceiptTemplate')
      .addItem('Открыть шаблон квитанции', 'openReceiptTemplate')
      .addItem('Создать новый лист-год', 'showCreateYearDialog')
      .addSeparator()
      .addItem('Открыть настройки', 'openSettingsSheet')
      .addItem('Открыть почты', 'openEmailsSheet')
      .addItem('Заполнить тестовыми адресами', 'fillTestEmails')
      .addItem('Открыть журнал', 'openJournalSheet')
      .addItem('Скрыть служебные листы', 'hideServiceSheets')
      .addSeparator()
      .addItem('Очистить журнал', 'clearJournal'))
    .addSeparator()
    .addItem('О программе', 'showAbout')
    .addToUi();
}

function showPdfDialog() {
  const years = getYearSheetNames_();
  if (!years.length) {
    SpreadsheetApp.getUi().alert('Не найдено ни одного листа с названием года.');
    return;
  }

  const now = new Date();
  const currentYear = String(now.getFullYear());
  const currentMonth = now.getMonth() + 1;
  const defaultYear = years.includes(currentYear) ? currentYear : years[years.length - 1];
  const options = years.map(year =>
    '<option value="' + year + '"' + (year === defaultYear ? ' selected' : '') + '>' + year + '</option>'
  ).join('');

  const html = HtmlService.createHtmlOutput(`
<!doctype html><html><head><base target="_top"><style>
body{font:14px Arial,sans-serif;padding:18px;color:#202124}h2{margin:0 0 16px;font-size:18px}
label{display:block;margin:12px 0 6px;font-weight:600}select{width:100%;box-sizing:border-box;padding:9px;border:1px solid #dadce0;border-radius:6px}
.info{background:#f8f9fa;padding:10px;border-radius:6px;margin-top:14px}.buttons{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}
button{padding:9px 14px;border:0;border-radius:6px;cursor:pointer}.primary{background:#1a73e8;color:#fff}.secondary{background:#f1f3f4}#status{margin-top:12px;min-height:18px;color:#5f6368}
</style></head><body><h2>Формирование PDF</h2>
<label for="year">Год</label><select id="year" onchange="refreshCount()">${options}</select>
<label for="month">Месяц</label><select id="month">
<option value="1">Январь</option><option value="2">Февраль</option><option value="3">Март</option><option value="4">Апрель</option>
<option value="5">Май</option><option value="6">Июнь</option><option value="7">Июль</option><option value="8">Август</option>
<option value="9">Сентябрь</option><option value="10">Октябрь</option><option value="11">Ноябрь</option><option value="12">Декабрь</option></select>
<div class="info">Найдено участков: <b id="count">…</b></div><div id="status"></div>
<div class="buttons"><button class="secondary" onclick="google.script.host.close()">Отмена</button><button class="primary" id="submit" onclick="submitForm()">Сформировать</button></div>
<script>
document.getElementById('year').value='${defaultYear}';document.getElementById('month').value='${currentMonth}';
function refreshCount(){const year=document.getElementById('year').value;google.script.run.withSuccessHandler(c=>document.getElementById('count').textContent=c).withFailureHandler(e=>document.getElementById('status').textContent=e.message).countPlotsForYear(year)}
function submitForm(){const b=document.getElementById('submit'),s=document.getElementById('status');b.disabled=true;s.textContent='Формирование запущено…';google.script.run.withSuccessHandler(r=>{s.textContent=r&&r.message?r.message:'Готово';b.disabled=false}).withFailureHandler(e=>{s.textContent='Ошибка: '+e.message;b.disabled=false}).startPdfGenerationFromDialog(Number(document.getElementById('year').value),Number(document.getElementById('month').value))}
refreshCount();
</script></body></html>`).setWidth(440).setHeight(410);
  SpreadsheetApp.getUi().showModalDialog(html, 'ДНП');
}

function openCurrentMonthFolder() {
  const now = new Date();
  const year = String(now.getFullYear());
  const month = now.getMonth() + 1;
  const root = getDnpPdfFolder_();
  const yearFolder = findChildFolderByNames_(root, [year]);
  if (!yearFolder) throw new Error('Папка года «' + year + '» ещё не создана.');
  const monthName = String(month).padStart(2, '0') + ' ' + getRussianMonthName_(month);
  const monthFolder = findChildFolderByNames_(yearFolder, [monthName]);
  if (!monthFolder) throw new Error('Папка «' + monthName + '» не найдена.');
  showDriveLinkDialog_('Папка месяца', monthFolder.getName(), monthFolder.getUrl());
}

function findChildFolderByNames_(parent, names) {
  const normalized = names.map(name => String(name).trim().toLowerCase());
  const folders = parent.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (normalized.includes(folder.getName().trim().toLowerCase())) return folder;
  }
  return null;
}

function removeLastPdfSetting_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Настройки');
  if (sheet && sheet.getLastRow() > 0) {
    const values = sheet.getRange(1, 1, sheet.getLastRow(), Math.min(sheet.getLastColumn(), 2)).getDisplayValues();
    for (let row = values.length - 1; row >= 0; row--) {
      const key = String(values[row][0] || '').trim().toLowerCase();
      if (key === 'последний pdf' || key === 'последний файл pdf') sheet.deleteRow(row + 1);
    }
  }
}