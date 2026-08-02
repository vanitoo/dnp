/**
 * DNP Receipts — единый файл Google Apps Script.
 * Файл создан автоматически из модулей каталога google-apps-script.
 * Для тестирования достаточно вставить только этот Code.gs.
 */


// ============================================================
// MODULE: Code.gs
// ============================================================

/** Ядро приложения DNP Receipts. */

const DNP_VERSION = '3.7.0';
const DNP_ADMIN_PASSWORD = '123456';
const DNP_SERVICE_SHEETS = {
  settings: 'Настройки',
  emails: 'Почты',
  journal: 'Журнал отправки',
};

function showAbout() {
  const ss = SpreadsheetApp.getActive();
  const years = getYearSheetNames_();
  SpreadsheetApp.getUi().alert(
    'DNP Receipts\n\n' +
    'Версия: ' + DNP_VERSION + '\n' +
    'Google Apps Script\n\n' +
    'Таблица: ' + ss.getName() + '\n' +
    'Листы годов: ' + (years.length ? years.join(', ') : 'не найдены')
  );
}

function getYearSheetNames_() {
  return SpreadsheetApp.getActive().getSheets()
    .map(sheet => sheet.getName().trim())
    .filter(name => /^\d{4}$/.test(name))
    .sort((a, b) => Number(a) - Number(b));
}

function countPlotsForYear(year) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(String(year));
  if (!sheet) throw new Error('Лист «' + year + '» не найден.');
  return getReceiptBlocks_(sheet).length;
}

function startPdfGenerationFromDialog(year, month) {
  const result = generatePdfsForMonth(Number(year), Number(month));
  return {
    ok: result && result.ok !== false,
    message: result && result.message ? result.message : 'Формирование завершено.',
  };
}


// ============================================================
// MODULE: Utils.gs
// ============================================================

/** Общие вспомогательные функции. */

function getSpreadsheetParentFolder_() {
  const file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parents = file.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

function getAvailablePropertyStores_() {
  const stores = [];
  try { const p = PropertiesService.getDocumentProperties(); if (p) stores.push(p); } catch (error) {}
  try { stores.push(PropertiesService.getScriptProperties()); } catch (error) {}
  try { stores.push(PropertiesService.getUserProperties()); } catch (error) {}
  return stores;
}

function saveProperty_(key, value) {
  const stores = getAvailablePropertyStores_();
  for (let i = 0; i < stores.length; i++) {
    try { stores[i].setProperty(key, String(value)); return; } catch (error) {}
  }
}

function deleteProperty_(key) {
  getAvailablePropertyStores_().forEach(store => {
    try { store.deleteProperty(key); } catch (error) {}
  });
}

function showDriveLinkDialog_(title, name, url) {
  const html = HtmlService.createHtmlOutput(
    '<div style="font:14px Arial,sans-serif;padding:18px">' +
      '<p><b>' + escapeHtml_(name) + '</b></p>' +
      '<p><a href="' + url + '" target="_blank" style="display:inline-block;padding:10px 14px;background:#1a73e8;color:#fff;text-decoration:none;border-radius:6px">Открыть в Google Drive</a></p>' +
    '</div>'
  ).setWidth(420).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}


// ============================================================
// MODULE: Menu.gs
// ============================================================

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
      .addItem('Добавить строчку услуги', 'showAddServiceRowDialog')
      .addItem('Создать шаблон под текущий формат', 'createReceiptTemplateForCurrentFormat')
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


// ============================================================
// MODULE: Setup.gs
// ============================================================

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


// ============================================================
// MODULE: YearSheets.gs
// ============================================================

/** Создание нового года, очистка вводимых значений и окраска блоков. */

function showCreateYearDialog() {
  const years = getYearSheetNames_();
  if (!years.length) {
    SpreadsheetApp.getUi().alert('Сначала создайте лист с названием года, например 2026.');
    return;
  }
  const suggestedYear = Math.max(...years.map(Number)) + 1;
  const options = years.slice().reverse().map(year => '<option value="' + year + '">' + year + '</option>').join('');
  const html = HtmlService.createHtmlOutput(`
<!doctype html><html><head><base target="_top"><style>
body{font:14px Arial,sans-serif;padding:18px;color:#202124}h2{margin:0 0 16px;font-size:18px}
label{display:block;margin:12px 0 6px;font-weight:600}input,select{width:100%;box-sizing:border-box;padding:9px;border:1px solid #dadce0;border-radius:6px}
.buttons{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}button{padding:9px 14px;border:0;border-radius:6px;cursor:pointer}.primary{background:#1a73e8;color:#fff}.secondary{background:#f1f3f4}#status{margin-top:12px;min-height:18px;color:#5f6368}
</style></head><body><h2>Создать новый лист-год</h2>
<label>Новый год</label><input id="newYear" type="number" min="2000" max="2100" value="${suggestedYear}">
<label>Копировать из листа</label><select id="sourceYear">${options}</select>
<div id="status"></div><div class="buttons"><button class="secondary" onclick="google.script.host.close()">Отмена</button><button class="primary" onclick="run()">Создать</button></div>
<script>
function run(){const s=document.getElementById('status');s.textContent='Создаётся лист…';google.script.run.withSuccessHandler(r=>{s.textContent=r.message;setTimeout(()=>google.script.host.close(),1200)}).withFailureHandler(e=>s.textContent='Ошибка: '+e.message).createNewYearSheet(Number(document.getElementById('newYear').value),document.getElementById('sourceYear').value,true)}
</script></body></html>`).setWidth(450).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'ДНП');
}

function createNewYearSheet(newYear, sourceYear, clearValues) {
  newYear = Number(newYear);
  sourceYear = String(sourceYear);
  if (!Number.isInteger(newYear) || newYear < 2000 || newYear > 2100) throw new Error('Некорректный новый год.');

  const ss = SpreadsheetApp.getActive();
  const newName = String(newYear);
  if (ss.getSheetByName(newName)) throw new Error('Лист «' + newName + '» уже существует.');
  const source = ss.getSheetByName(sourceYear);
  if (!source) throw new Error('Исходный лист «' + sourceYear + '» не найден.');

  const target = source.copyTo(ss).setName(newName);
  ss.setActiveSheet(target);
  if (clearValues) clearCopiedYearValues_(target);
  applySevenRowBanding_(target);
  ss.toast('Создан лист ' + newName + ' на основе ' + sourceYear, 'ДНП', 7);
  return { ok: true, message: 'Лист «' + newName + '» создан.' };
}

function clearCopiedYearValues_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 3) return;

  const labels = sheet.getRange(1, 2, lastRow, 1).getDisplayValues().flat().map(normalizeDnpRowLabel_);
  let insideBlock = false;

  for (let row = 2; row <= lastRow; row++) {
    const label = labels[row - 1];

    if (/^тариф(?:ы)?$/.test(label)) {
      insideBlock = true;
      continue; // Номера месяцев в строке «Тариф» сохраняем.
    }
    if (!insideBlock || !label) continue;

    // В каждом блоке очищаем все введённые значения услуг и итогов.
    // Это автоматически поддерживает новые строки: целевые взносы,
    // водоотведение и любые будущие услуги.
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
  return String(value == null ? '' : value).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
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

  const labels = sheet.getRange(1, 2, lastRow, 1).getDisplayValues().flat().map(normalizeDnpRowLabel_);
  const tariffRows = [];
  for (let index = 1; index < labels.length; index++) {
    if (/^тариф(?:ы)?$/.test(labels[index])) tariffRows.push(index + 1);
  }

  const colors = ['#ffffff', '#f3f7fd'];
  if (tariffRows.length) {
    tariffRows.forEach((startRow, blockIndex) => {
      const endRow = blockIndex + 1 < tariffRows.length ? tariffRows[blockIndex + 1] - 1 : lastRow;
      const rowCount = endRow - startRow + 1;
      if (rowCount > 0) sheet.getRange(startRow, 1, rowCount, lastColumn).setBackground(colors[blockIndex % colors.length]);
    });
    return;
  }

  for (let row = 2, blockIndex = 0; row <= lastRow; row += 7, blockIndex++) {
    const rowCount = Math.min(7, lastRow - row + 1);
    sheet.getRange(row, 1, rowCount, lastColumn).setBackground(colors[blockIndex % colors.length]);
  }
}


// ============================================================
// MODULE: Services.gs
// ============================================================

/** Добавление новой строки услуги во все блоки текущего листа-года. */

function showAddServiceRowDialog() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!/^\d{4}$/.test(sheet.getName())) {
    SpreadsheetApp.getUi().alert(
      'Сначала откройте лист нужного года, например 2026.\n\n' +
      'Строка услуги будет добавлена именно в текущий лист.'
    );
    return;
  }

  const html = HtmlService.createHtmlOutput(`
<!doctype html><html><head><base target="_top"><style>
body{font:14px Arial,sans-serif;padding:18px;color:#202124}h2{margin:0 0 14px;font-size:18px}
label{display:block;margin:10px 0 6px;font-weight:600}input{width:100%;box-sizing:border-box;padding:9px;border:1px solid #dadce0;border-radius:6px}
.note{margin-top:10px;color:#5f6368;line-height:1.4}.buttons{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}
button{padding:9px 14px;border:0;border-radius:6px;cursor:pointer}.primary{background:#1a73e8;color:#fff}.secondary{background:#f1f3f4}#status{margin-top:12px;min-height:18px;color:#5f6368}
</style></head><body>
<h2>Добавить строку услуги</h2>
<label for="serviceName">Название услуги</label>
<input id="serviceName" value="Целевые взносы" autofocus>
<div class="note">Скрипт проверит каждый блок от строки «Тариф» до следующей строки «Тариф». Если услуги нет, новая строка будет вставлена перед строкой «Сумма».</div>
<div id="status"></div>
<div class="buttons"><button class="secondary" onclick="google.script.host.close()">Отмена</button><button class="primary" onclick="run()">Добавить</button></div>
<script>
document.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();run()}if(e.key==='Escape')google.script.host.close()});
function run(){const name=document.getElementById('serviceName').value.trim(),status=document.getElementById('status');if(!name){status.textContent='Введите название услуги.';return}status.textContent='Проверка блоков…';google.script.run.withSuccessHandler(r=>{status.textContent=r.message;setTimeout(()=>google.script.host.close(),1400)}).withFailureHandler(e=>status.textContent='Ошибка: '+e.message).addServiceRowToCurrentSheet(name)}
</script></body></html>`).setWidth(480).setHeight(350);

  SpreadsheetApp.getUi().showModalDialog(html, 'ДНП');
}

function addServiceRowToCurrentSheet(serviceName) {
  serviceName = String(serviceName || '').replace(/\s+/g, ' ').trim();
  if (!serviceName) throw new Error('Название услуги не указано.');
  if (/^тариф(?:ы)?$/i.test(serviceName) || /^(сумма|итого)/i.test(serviceName)) {
    throw new Error('Название услуги не может быть «Тариф», «Сумма» или «Итого».');
  }

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  if (!/^\d{4}$/.test(sheet.getName())) {
    throw new Error('Текущий лист должен называться годом, например 2026.');
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), 2);
  if (lastRow < 2) throw new Error('На листе нет блоков участков.');

  const labels = sheet.getRange(1, 2, lastRow, 1).getDisplayValues().flat().map(normalizeDnpRowLabel_);
  const tariffRows = [];
  for (let index = 1; index < labels.length; index++) {
    if (/^тариф(?:ы)?$/.test(labels[index])) tariffRows.push(index + 1);
  }
  if (!tariffRows.length) throw new Error('В столбце B не найдены строки «Тариф».');

  const normalizedService = normalizeDnpRowLabel_(serviceName);
  const operations = [];
  let skipped = 0;

  for (let index = 0; index < tariffRows.length; index++) {
    const startRow = tariffRows[index];
    const endRow = index + 1 < tariffRows.length ? tariffRows[index + 1] - 1 : lastRow;
    const blockLabels = labels.slice(startRow - 1, endRow);

    if (blockLabels.includes(normalizedService)) {
      skipped++;
      continue;
    }

    let sumRow = -1;
    for (let row = endRow; row >= startRow; row--) {
      if (/^(сумма|итого)/.test(labels[row - 1])) {
        sumRow = row;
        break;
      }
    }
    if (sumRow === -1) {
      throw new Error('В блоке, начинающемся со строки ' + startRow + ', не найдена строка «Сумма».');
    }
    operations.push(sumRow);
  }

  // Вставляем снизу вверх, чтобы номера строк верхних блоков не сдвигались.
  operations.sort((a, b) => b - a).forEach(sumRow => {
    const sourceRow = Math.max(sumRow - 1, 1);
    const sourceHeight = sheet.getRowHeight(sourceRow);

    sheet.insertRowsBefore(sumRow, 1);
    sheet.getRange(sourceRow, 1, 1, lastColumn)
      .copyTo(sheet.getRange(sumRow, 1, 1, lastColumn), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sheet.getRange(sourceRow, 1, 1, lastColumn)
      .copyTo(sheet.getRange(sumRow, 1, 1, lastColumn), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    sheet.setRowHeight(sumRow, sourceHeight);
    sheet.getRange(sumRow, 1, 1, lastColumn).clearContent();
    sheet.getRange(sumRow, 2).setValue(serviceName);
  });

  applySevenRowBanding_(sheet);
  ss.toast('Добавлено строк: ' + operations.length + '. Пропущено: ' + skipped + '.', 'ДНП', 7);

  return {
    ok: true,
    added: operations.length,
    skipped: skipped,
    message: 'Готово. Добавлено: ' + operations.length + ', уже существовало: ' + skipped + '.'
  };
}


// ============================================================
// MODULE: Template.gs
// ============================================================

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


// ============================================================
// MODULE: PdfGenerator.gs
// ============================================================

/** Формирование PDF-квитанций строго через Google Docs-шаблон. */

function generatePdfsForMonth(year, month) {
  year = Number(year);
  month = Number(month);
  if (!Number.isInteger(year) || year < 2000 || year > 2100) throw new Error('Некорректный год.');
  if (!Number.isInteger(month) || month < 1 || month > 12) throw new Error('Некорректный месяц.');

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(String(year));
  if (!sheet) throw new Error('Лист «' + year + '» не найден.');

  const templateFile = getReceiptTemplateFile_();
  const rootFolder = getDnpPdfFolder_();
  const yearFolder = getOrCreatePdfChildFolder_(rootFolder, String(year));
  const monthFolderName = String(month).padStart(2, '0') + ' ' + getRussianMonthName_(month);
  const monthFolder = getOrCreatePdfChildFolder_(yearFolder, monthFolderName);
  const blocks = getReceiptBlocks_(sheet);
  const rates = getReceiptRates_();
  if (!blocks.length) throw new Error('Не найдены строки «Тариф» в столбце B.');

  let created = 0;
  let failed = 0;
  const errors = [];

  blocks.forEach((block, index) => {
    let tempFile = null;
    ss.toast('Формируется ' + (index + 1) + ' из ' + blocks.length + ': участок ' + block.plot, 'ДНП', 5);
    try {
      const receipt = buildReceiptData_(sheet, block, year, month, rates);
      const fileName = 'Квитанция_участок_' + sanitizePdfFileName_(block.plot) + '_' + year + '_' + String(month).padStart(2, '0') + '.pdf';
      trashFilesByName_(monthFolder, fileName);

      tempFile = templateFile.makeCopy('Временная квитанция ' + block.plot, monthFolder);
      const doc = DocumentApp.openById(tempFile.getId());
      fillReceiptTemplate_(doc, receipt);
      doc.saveAndClose();

      monthFolder.createFile(tempFile.getAs(MimeType.PDF).setName(fileName));
      tempFile.setTrashed(true);
      tempFile = null;
      created++;
      Utilities.sleep(150);
    } catch (error) {
      failed++;
      errors.push('Участок ' + block.plot + ': ' + (error.message || error));
      if (tempFile) try { tempFile.setTrashed(true); } catch (cleanupError) {}
    }
  });

  const message = failed
    ? 'Создано PDF: ' + created + '. Ошибок: ' + failed + '. Первая ошибка: ' + errors[0]
    : 'Создано PDF: ' + created + '. Папка: ' + year + '/' + monthFolderName + '.';
  ss.toast(message, 'ДНП', 10);
  return { ok: failed === 0, created, failed, folderId: monthFolder.getId(), folderUrl: monthFolder.getUrl(), message };
}

function getReceiptBlocks_(sheet) {
  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(1, 1, lastRow, 2).getDisplayValues();
  const starts = [];
  for (let index = 1; index < rows.length; index++) {
    const plot = String(rows[index][0] || '').trim();
    const label = normalizeReceiptLabel_(rows[index][1]);
    if (/^тариф(?:ы)?$/.test(label)) {
      if (!plot) throw new Error('В строке ' + (index + 1) + ' найден «Тариф», но в столбце A нет номера участка.');
      starts.push({ startRow: index + 1, plot });
    }
  }
  return starts.map((item, index) => ({
    plot: item.plot,
    startRow: item.startRow,
    endRow: index + 1 < starts.length ? starts[index + 1].startRow - 1 : lastRow,
  }));
}

function buildReceiptData_(sheet, block, year, month, rates) {
  const monthColumn = month + 2;
  const previousMonthColumn = monthColumn - 1;
  const rowCount = block.endRow - block.startRow + 1;
  const lastColumn = Math.max(sheet.getLastColumn(), monthColumn);
  const values = sheet.getRange(block.startRow, 1, rowCount, lastColumn).getDisplayValues();
  const byLabel = {};
  values.forEach(row => {
    const label = normalizeReceiptLabel_(row[1]);
    if (label) byLabel[label] = row;
  });

  const paymentRows = [];
  let calculatedTotal = 0;

  ['т1', 'т2', 'т3'].forEach((label, index) => {
    const row = byLabel[label];
    if (!row) return;
    const current = parseReceiptNumber_(row[monthColumn - 1]);
    const previous = month > 1 ? parseReceiptNumber_(row[previousMonthColumn - 1]) : null;
    const usage = current !== null && previous !== null ? current - previous : null;
    const rate = rates['T' + (index + 1)];
    const amount = usage !== null && rate !== null ? usage * rate : null;
    if (amount !== null) calculatedTotal += amount;
    paymentRows.push([
      'Электроэнергия Т' + (index + 1),
      formatReceiptValue_(current), formatReceiptValue_(previous), formatReceiptValue_(usage),
      formatReceiptMoney_(rate), formatReceiptMoney_(amount),
    ]);
  });

  const waterRow = findReceiptRow_(byLabel, [/^водоотвед/, /^вода$/]);
  if (waterRow) {
    const current = parseReceiptNumber_(waterRow[monthColumn - 1]);
    const previous = month > 1 ? parseReceiptNumber_(waterRow[previousMonthColumn - 1]) : null;
    const usage = current !== null && previous !== null ? current - previous : null;
    const amount = usage !== null && rates.WATER !== null ? usage * rates.WATER : null;
    if (amount !== null) calculatedTotal += amount;
    paymentRows.push(['Водоотведение', formatReceiptValue_(current), formatReceiptValue_(previous), formatReceiptValue_(usage), formatReceiptMoney_(rates.WATER), formatReceiptMoney_(amount)]);
  }

  // Все прочие строки между «Тариф» и «Сумма» считаем отдельными услугами.
  // В ячейке выбранного месяца хранится готовая сумма начисления.
  const ignoredLabels = /^(тариф(?:ы)?|т1|т2|т3|квтч|квт·ч|квт\/ч|водоотвед.*|вода|сумма.*|итого.*)$/;
  values.forEach(row => {
    const rawLabel = String(row[1] == null ? '' : row[1]).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
    const label = normalizeReceiptLabel_(rawLabel);
    if (!label || ignoredLabels.test(label)) return;

    const amount = parseReceiptNumber_(row[monthColumn - 1]);
    if (amount === null) return;

    calculatedTotal += amount;
    paymentRows.push([rawLabel, '—', '—', '—', '—', formatReceiptMoney_(amount)]);
  });

  const totalRow = findReceiptRow_(byLabel, [/^сумма/, /^итого/]);
  const storedTotal = totalRow ? parseReceiptNumber_(totalRow[monthColumn - 1]) : null;
  const total = storedTotal !== null ? storedTotal : calculatedTotal;
  return { plot: block.plot, year, month, monthName: getRussianMonthName_(month), total, paymentRows };
}

function fillReceiptTemplate_(doc, receipt) {
  const body = doc.getBody();
  replaceReceiptText_(body, '{{PLOT}}', receipt.plot);
  replaceReceiptText_(body, '{{YEAR}}', receipt.year);
  replaceReceiptText_(body, '{{MONTH}}', String(receipt.month).padStart(2, '0'));
  replaceReceiptText_(body, '{{MONTH_NAME}}', receipt.monthName);
  replaceReceiptText_(body, '{{TOTAL}}', formatReceiptMoney_(receipt.total));
  insertPaymentTable_(body, receipt.paymentRows, receipt.total);
}

function insertPaymentTable_(body, paymentRows, total) {
  const found = body.findText('\\{\\{PAYMENT_TABLE\\}\\}');
  if (!found) throw new Error('В Google Docs-шаблоне не найден маркер {{PAYMENT_TABLE}}.');
  const paragraph = found.getElement().asText().getParent().asParagraph();
  const index = body.getChildIndex(paragraph);
  const rows = [['Наименование платежа', 'Текущее', 'Предыдущее', 'Объём', 'Тариф', 'Сумма к оплате']].concat(paymentRows);
  rows.push(['ИТОГО К ОПЛАТЕ', '', '', '', '', formatReceiptMoney_(total) + ' руб.']);

  const table = body.insertTable(index, rows);
  body.removeChild(paragraph);
  table.getRow(0).editAsText().setBold(true);
  table.getRow(table.getNumRows() - 1).editAsText().setBold(true);
  applyPaymentTableColumnWidths_(table);
}

function getReceiptTemplateFile_() {
  const templateId = getStoredTemplateId_();
  if (templateId) {
    try {
      const file = DriveApp.getFileById(templateId);
      file.getName();
      return file;
    } catch (error) {}
  }
  const parent = getSpreadsheetParentFolder_();
  const files = parent.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (/^Шаблон квитанции ДНП Комфорт/.test(file.getName())) {
      saveTemplateId_(file.getId());
      return file;
    }
  }
  throw new Error('Шаблон квитанции не найден. Выполните «ДНП → Настройка → Создать шаблон под текущий формат».');
}

function getStoredTemplateId_() {
  const stores = getAvailablePropertyStores_();
  for (let i = 0; i < stores.length; i++) {
    try { const value = stores[i].getProperty('TEMPLATE_DOC_ID'); if (value) return value; } catch (error) {}
  }
  return '';
}

function saveTemplateId_(templateId) {
  const stores = getAvailablePropertyStores_();
  for (let i = 0; i < stores.length; i++) {
    try { stores[i].setProperty('TEMPLATE_DOC_ID', templateId); return; } catch (error) {}
  }
}

function getReceiptRates_() {
  const result = { T1: null, T2: null, T3: null, WATER: null };
  const settings = SpreadsheetApp.getActive().getSheetByName('Настройки');
  if (!settings || settings.getLastRow() < 2) return result;
  settings.getRange(2, 1, settings.getLastRow() - 1, 2).getDisplayValues().forEach(row => {
    const key = normalizeReceiptLabel_(row[0]);
    const value = parseReceiptNumber_(row[1]);
    if (/тариф\s*т1/.test(key)) result.T1 = value;
    if (/тариф\s*т2/.test(key)) result.T2 = value;
    if (/тариф\s*т3/.test(key)) result.T3 = value;
    if (/водоотвед/.test(key)) result.WATER = value;
  });
  return result;
}

function findReceiptRow_(byLabel, patterns) {
  const labels = Object.keys(byLabel);
  for (let i = 0; i < labels.length; i++) if (patterns.some(pattern => pattern.test(labels[i]))) return byLabel[labels[i]];
  return null;
}

function replaceReceiptText_(body, marker, value) { body.replaceText(escapeReceiptRegex_(marker), String(value == null ? '' : value)); }
function escapeReceiptRegex_(text) { return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function normalizeReceiptLabel_(value) { return String(value == null ? '' : value).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase().replace(/,?\s*₽$/i, ''); }
function parseReceiptNumber_(value) {
  const text = String(value == null ? '' : value).replace(/\u00a0/g, '').replace(/\s/g, '').replace(',', '.').replace(/[^0-9.\-]/g, '');
  if (!text || text === '-' || text === '.') return null;
  const number = Number(text);
  return Number.isFinite(number) ? number : null;
}
function formatReceiptValue_(value) { return value === null || value === undefined || !Number.isFinite(Number(value)) ? '' : Number(value).toLocaleString('ru-RU', { maximumFractionDigits: 3 }); }
function formatReceiptMoney_(value) { return value === null || value === undefined || !Number.isFinite(Number(value)) ? '' : Number(value).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
function getOrCreatePdfChildFolder_(parent, name) { const folders = parent.getFoldersByName(name); return folders.hasNext() ? folders.next() : parent.createFolder(name); }
function getRussianMonthName_(month) { return ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'][month - 1]; }
function sanitizePdfFileName_(value) { return String(value).trim().replace(/[\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_'); }
function trashFilesByName_(folder, fileName) { const files = folder.getFilesByName(fileName); while (files.hasNext()) try { files.next().setTrashed(true); } catch (error) {} }


// ============================================================
// MODULE: PdfCleanup.gs
// ============================================================

/** Очистка сформированных PDF. */

function clearGeneratedPdfs() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Очистить сформированные PDF?',
    'Все PDF-файлы в папке квитанций и её подпапках будут перемещены в корзину. Папки и шаблон останутся.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActive();
  try {
    ss.toast('Поиск сформированных PDF…', 'ДНП', 5);
    const folder = getDnpPdfFolder_();
    const result = trashDnpPdfFilesRecursively_(folder);
    const message = result.failed
      ? 'Перемещено в корзину: ' + result.deleted + '. Не удалось удалить: ' + result.failed + '.'
      : 'Перемещено в корзину PDF-файлов: ' + result.deleted + '.';
    ss.toast(message, 'ДНП', 8);
    ui.alert('Очистка завершена', message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Ошибка очистки PDF', error.message || String(error), ui.ButtonSet.OK);
    throw error;
  }
}

function getDnpPdfFolder_() {
  const stores = getAvailablePropertyStores_();
  for (let i = 0; i < stores.length; i++) {
    try {
      const id = stores[i].getProperty('PDF_FOLDER_ID');
      if (id) {
        const folder = DriveApp.getFolderById(id);
        folder.getName();
        return folder;
      }
    } catch (error) {}
  }

  const parent = getSpreadsheetParentFolder_();
  const folders = parent.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    if (/^(Квитанции ДНП Комфорт|Квитанции|Receipts)/i.test(folder.getName())) {
      saveProperty_('PDF_FOLDER_ID', folder.getId());
      return folder;
    }
  }
  throw new Error('Папка квитанций не найдена. Выполните первичную настройку.');
}

function trashDnpPdfFilesRecursively_(folder) {
  let deleted = 0;
  let failed = 0;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() !== MimeType.PDF && !/\.pdf$/i.test(file.getName())) continue;
    try { file.setTrashed(true); deleted++; } catch (error) { failed++; }
  }
  const children = folder.getFolders();
  while (children.hasNext()) {
    const result = trashDnpPdfFilesRecursively_(children.next());
    deleted += result.deleted;
    failed += result.failed;
  }
  return { deleted, failed };
}


// ============================================================
// MODULE: Mail.gs
// ============================================================

/** Работа с листом почты. */

function syncEmailSheetPlots_() {
  const ss = SpreadsheetApp.getActive();
  const emailSheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails);
  const years = getYearSheetNames_();
  if (!emailSheet || !years.length) return;

  const source = ss.getSheetByName(years[years.length - 1]);
  const blocks = getReceiptBlocks_(source);
  const plots = blocks.map(block => String(block.plot).trim());
  const existing = emailSheet.getLastRow() > 1
    ? emailSheet.getRange(2, 1, emailSheet.getLastRow() - 1, 4).getValues()
    : [];
  const byPlot = new Map(existing.map(row => [String(row[0]).trim(), row]));
  const rows = plots.map(plot => byPlot.get(plot) || [plot, '', '', '']);

  if (emailSheet.getLastRow() > 1) {
    emailSheet.getRange(2, 1, emailSheet.getLastRow() - 1, 4).clearContent();
  }
  if (rows.length) emailSheet.getRange(2, 1, rows.length, 4).setValues(rows);
}

function fillTestEmails() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails);
  if (!sheet) { ensureServiceSheets_(); sheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails); }
  if (sheet.getLastRow() < 2) syncEmailSheetPlots_();

  const count = Math.max(sheet.getLastRow() - 1, 0);
  if (!count) { SpreadsheetApp.getUi().alert('На листе «Почты» нет участков.'); return; }

  const values = sheet.getRange(2, 1, count, 2).getValues();
  let filled = 0;
  values.forEach(row => {
    if (row[0] && !row[1]) {
      row[1] = String(row[0]).replace(/[^0-9A-Za-zА-Яа-я_-]+/g, '_') + '@mail.ru';
      filled++;
    }
  });
  sheet.getRange(2, 1, count, 2).setValues(values);
  ss.toast('Добавлено тестовых адресов: ' + filled, 'ДНП', 5);
}

function sendReceipts() {
  SpreadsheetApp.getUi().alert('Функция отправки квитанций будет подключена после проверки формирования PDF по новому шаблону.');
}


// ============================================================
// MODULE: Journal.gs
// ============================================================

/** Журнал операций. */

function clearJournal() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert('Очистить журнал?', 'Будут удалены все строки, кроме заголовка.', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.getActive().getSheetByName(DNP_SERVICE_SHEETS.journal);
  if (!sheet) { ui.alert('Лист журнала не найден.'); return; }
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
  }
  SpreadsheetApp.getActive().toast('Журнал очищен', 'ДНП', 5);
}

function appendJournalRow_(operation, year, month, plot, email, status, errorText) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(DNP_SERVICE_SHEETS.journal);
  if (!sheet) { ensureServiceSheets_(); sheet = SpreadsheetApp.getActive().getSheetByName(DNP_SERVICE_SHEETS.journal); }
  sheet.appendRow([new Date(), operation || '', year || '', month || '', plot || '', email || '', status || '', errorText || '']);
}
