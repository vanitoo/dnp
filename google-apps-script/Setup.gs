/** Первичная настройка и служебные листы. */

function showInitialSetupDialog() {
  const html = HtmlService.createHtmlOutput(`
<!doctype html><html><head><base target="_top"><style>
body{font:14px Arial,sans-serif;padding:18px;color:#202124}h2{margin:0 0 16px;font-size:18px}
label{display:block;margin:12px 0 6px;font-weight:600}input,select{width:100%;box-sizing:border-box;padding:9px;border:1px solid #dadce0;border-radius:6px}
.buttons{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}button{padding:9px 14px;border:0;border-radius:6px;cursor:pointer}
.primary{background:#1a73e8;color:#fff}.secondary{background:#f1f3f4}#status{margin-top:12px;color:#5f6368;min-height:18px}
</style></head><body><h2>Первичная настройка</h2>
<label for="password">Пароль администратора</label><input id="password" type="password" autocomplete="off" autofocus>
<label for="mode">Папка квитанций</label><select id="mode"><option value="reuse">Использовать существующую или создать</option><option value="recreate">Создать новую папку</option></select>
<div id="status"></div><div class="buttons"><button class="secondary" onclick="google.script.host.close()">Отмена</button><button class="primary" onclick="run()">Продолжить</button></div>
<script>
function run(){const s=document.getElementById('status');s.textContent='Выполняется настройка…';google.script.run.withSuccessHandler(r=>{s.textContent=r.message;setTimeout(()=>google.script.host.close(),1200)}).withFailureHandler(e=>s.textContent='Ошибка: '+e.message).runInitialSetup(document.getElementById('password').value,document.getElementById('mode').value)}
</script></body></html>`).setWidth(440).setHeight(330);
  SpreadsheetApp.getUi().showModalDialog(html, 'ДНП');
}

function runInitialSetup(password, mode) {
  if (String(password) !== DNP_ADMIN_PASSWORD) throw new Error('Неверный пароль.');
  if (!['reuse', 'recreate'].includes(mode)) throw new Error('Неизвестный режим настройки.');

  const ss = SpreadsheetApp.getActive();
  const parent = getSpreadsheetParentFolder_();
  const pdfFolder = getOrCreateSetupFolder_(parent, 'Квитанции ДНП Комфорт', mode === 'recreate');

  saveProperty_('PDF_FOLDER_ID', pdfFolder.getId());
  saveProperty_('APP_VERSION', DNP_VERSION);
  ensureServiceSheets_();
  applySevenRowBandingToYearSheets_();
  hideServiceSheets();

  ss.toast('Первичная настройка завершена', 'ДНП', 7);
  return {
    ok: true,
    message: 'Готово. Папка квитанций: ' + pdfFolder.getName() + '. Шаблон создаётся отдельно через меню «Создать шаблон квитанции».',
  };
}

function getOrCreateSetupFolder_(parent, name, forceCreate) {
  if (!forceCreate) {
    const found = parent.getFoldersByName(name);
    if (found.hasNext()) return found.next();
  }
  const suffix = forceCreate ? ' ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm-ss') : '';
  return parent.createFolder(name + suffix);
}

function ensureServiceSheets_() {
  const ss = SpreadsheetApp.getActive();
  if (!ss.getSheetByName(DNP_SERVICE_SHEETS.settings)) {
    const sheet = ss.insertSheet(DNP_SERVICE_SHEETS.settings);
    sheet.getRange('A1:B6').setValues([
      ['Параметр', 'Значение'], ['Версия', DNP_VERSION], ['Тариф Т1', 9.67],
      ['Тариф Т2', 3.51], ['Тариф Т3', 6.77], ['Тариф водоотведения', 53],
    ]);
    sheet.setFrozenRows(1);
  }
  if (!ss.getSheetByName(DNP_SERVICE_SHEETS.emails)) {
    const sheet = ss.insertSheet(DNP_SERVICE_SHEETS.emails);
    sheet.getRange('A1:D1').setValues([['Участок', 'Email', 'Статус', 'Комментарий']]);
    sheet.setFrozenRows(1);
    syncEmailSheetPlots_();
  }
  if (!ss.getSheetByName(DNP_SERVICE_SHEETS.journal)) {
    const sheet = ss.insertSheet(DNP_SERVICE_SHEETS.journal);
    sheet.getRange('A1:H1').setValues([['Дата','Операция','Год','Месяц','Участок','Email','Статус','Ошибка']]);
    sheet.setFrozenRows(1);
  }
}

function openSettingsSheet() { showAndActivateSheet_(DNP_SERVICE_SHEETS.settings); }
function openEmailsSheet() { showAndActivateSheet_(DNP_SERVICE_SHEETS.emails); }
function openJournalSheet() { showAndActivateSheet_(DNP_SERVICE_SHEETS.journal); }

function showAndActivateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(name);
  if (!sheet) { SpreadsheetApp.getUi().alert('Лист «' + name + '» не найден.'); return; }
  sheet.showSheet();
  ss.setActiveSheet(sheet);
}

function hideServiceSheets() {
  const ss = SpreadsheetApp.getActive();
  let hidden = 0;
  Object.values(DNP_SERVICE_SHEETS).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && !sheet.isSheetHidden()) { sheet.hideSheet(); hidden++; }
  });
  ss.toast('Скрыто служебных листов: ' + hidden, 'ДНП', 5);
}
