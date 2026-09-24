/**
 * DNP Receipts — единый файл Google Apps Script.
 * Для работы достаточно вставить только этот Code.gs.
 */

const DNP_VERSION = '3.12.0';
const DNP_ADMIN_PASSWORD = '123456';
const DNP_PDF_LOG_ENABLED = false;
const DNP_PDF_SLEEP_MS = 20;
const DNP_PASSWORD_CHECK_AFTER = '20261001';

const DNP_SERVICE_SHEETS = {
  settings: 'Настройки',
  emails: 'Почты',
  journal: 'Журнал отправки',
};

const DNP_TARIFF_SHEET_NAME = 'Тарифы';

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

function isOperationPasswordRequired_() {
//  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
//  return today > DNP_PASSWORD_CHECK_AFTER;
  return false;
}

function requireOperationPassword_(password) {
//  if (isOperationPasswordRequired_() && String(password == null ? '' : password).trim() !== DNP_VERSION) {
//    throw new Error('Неверный пароль. Укажите текущую версию программы.');
//  }
  return;
}

function startPdfGenerationFromDialog(year, month, password) {
  const result = generatePdfsForMonthWithLog(Number(year), Number(month), password);
  return {
    ok: result && result.ok !== false,
    created: Number(result && result.created || 0),
    failed: Number(result && result.failed || 0),
    message: result && result.message ? result.message : 'Формирование завершено.',
  };
}

const DNP_DRIVE_FOLDER_MIME = 'application/vnd.google-apps.folder';
const DNP_DRIVE_DOC_MIME = 'application/vnd.google-apps.document';
const DNP_DRIVE_PDF_MIME = 'application/pdf';
const DNP_ROOT_APP_PROPERTY = 'dnpReceiptsRoot';
const DNP_TEMPLATE_APP_PROPERTY = 'dnpReceiptsTemplate';

function dnpDriveId_(item) {
  if (!item) return '';
  if (typeof item === 'string') return item;
  return String(item.id || '');
}

function dnpDriveFolderUrl_(item) {
  const id = dnpDriveId_(item);
  return id ? 'https://drive.google.com/drive/folders/' + encodeURIComponent(id) : '';
}

function dnpDriveDocUrl_(item) {
  const id = dnpDriveId_(item);
  return id ? 'https://docs.google.com/document/d/' + encodeURIComponent(id) + '/edit' : '';
}

function dnpDriveEscapeQuery_(value) {
  return String(value == null ? '' : value)
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'");
}

function dnpDriveGet_(fileId) {
  const id = dnpDriveId_(fileId);
  if (!id) throw new Error('Не указан ID файла Google Drive.');
  return Drive.Files.get(id, {
    fields: 'id,name,mimeType,parents,trashed,appProperties'
  });
}

function dnpDriveList_(query) {
  const files = [];
  let pageToken = null;
  do {
    const options = {
      q: query,
      pageSize: 1000,
      fields: 'nextPageToken,files(id,name,mimeType,parents,trashed,appProperties)'
    };
    if (pageToken) options.pageToken = pageToken;
    const response = Drive.Files.list(options);
    if (response.files && response.files.length) files.push(...response.files);
    pageToken = response.nextPageToken || null;
  } while (pageToken);
  return files;
}

function dnpDriveChildren_(parent, mimeType) {
  const parentId = dnpDriveId_(parent);
  if (!parentId) return [];
  let query = "'" + dnpDriveEscapeQuery_(parentId) + "' in parents and trashed = false";
  if (mimeType) query += " and mimeType = '" + dnpDriveEscapeQuery_(mimeType) + "'";
  return dnpDriveList_(query);
}

function dnpDriveFindByName_(parent, name, mimeType) {
  const parentId = dnpDriveId_(parent);
  let query = "'" + dnpDriveEscapeQuery_(parentId) + "' in parents" +
    " and name = '" + dnpDriveEscapeQuery_(name) + "'" +
    ' and trashed = false';
  if (mimeType) query += " and mimeType = '" + dnpDriveEscapeQuery_(mimeType) + "'";
  return dnpDriveList_(query);
}

function dnpDriveCreateFolder_(name, parent, appProperties) {
  const parentId = dnpDriveId_(parent);
  const resource = {
    name: String(name),
    mimeType: DNP_DRIVE_FOLDER_MIME,
    appProperties: appProperties || {}
  };
  if (parentId) resource.parents = [parentId];
  const created = Drive.Files.create(resource);
  return {
    id: created.id,
    name: String(name),
    mimeType: DNP_DRIVE_FOLDER_MIME,
    parents: parentId ? [parentId] : [],
    appProperties: resource.appProperties
  };
}

function dnpDriveCreateGoogleDoc_(name, parent, appProperties) {
  const parentId = dnpDriveId_(parent);
  const resource = {
    name: String(name),
    mimeType: DNP_DRIVE_DOC_MIME,
    appProperties: appProperties || {}
  };
  if (parentId) resource.parents = [parentId];
  const created = Drive.Files.create(resource);
  return {
    id: created.id,
    name: String(name),
    mimeType: DNP_DRIVE_DOC_MIME,
    parents: parentId ? [parentId] : [],
    appProperties: resource.appProperties
  };
}

function dnpDriveCopy_(source, name, parent) {
  const sourceId = dnpDriveId_(source);
  const parentId = dnpDriveId_(parent);
  return Drive.Files.copy({
    name: String(name),
    parents: [parentId]
  }, sourceId, {
    fields: 'id,name,mimeType,parents,trashed,appProperties'
  });
}

function dnpDriveTrash_(file) {
  const id = dnpDriveId_(file);
  if (!id) return;
  Drive.Files.update({ trashed: true }, id, null, { fields: 'id,trashed' });
}

function dnpDriveCreateBlobFile_(parent, name, blob, mimeType) {
  const parentId = dnpDriveId_(parent);
  const data = blob.copyBlob().setName(String(name));
  if (mimeType) data.setContentType(mimeType);
  return Drive.Files.create({
    name: String(name),
    mimeType: mimeType || data.getContentType(),
    parents: [parentId]
  }, data, {
    fields: 'id,name,mimeType,parents'
  });
}

function dnpDriveFetchBlob_(url, fileName) {
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Google Drive вернул HTTP ' + code + ': ' + response.getContentText().slice(0, 500));
  }
  const blob = response.getBlob();
  if (fileName) blob.setName(fileName);
  return blob;
}

function dnpDriveExportPdf_(doc, fileName) {
  const id = dnpDriveId_(doc);
  const url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) +
    '/export?mimeType=' + encodeURIComponent(DNP_DRIVE_PDF_MIME);
  return dnpDriveFetchBlob_(url, fileName).setContentType(DNP_DRIVE_PDF_MIME);
}

function dnpDriveDownloadBlob_(file, fileName) {
  const id = dnpDriveId_(file);
  const url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) + '?alt=media';
  return dnpDriveFetchBlob_(url, fileName);
}

function getSettingsSheet_() {
  return SpreadsheetApp.getActive().getSheetByName(DNP_SERVICE_SHEETS.settings);
}

function getSettingValue_(key) {
  const sheet = getSettingsSheet_();
  if (!sheet || sheet.getLastRow() < 1) return '';

  const wanted = normalizeSettingKey_(key);
  const values = sheet
    .getRange(1, 1, sheet.getLastRow(), Math.max(2, Math.min(sheet.getLastColumn(), 3)))
    .getDisplayValues();

  for (let row = 0; row < values.length; row++) {
    if (normalizeSettingKey_(values[row][0]) === wanted) {
      return String(values[row][1] == null ? '' : values[row][1]).trim();
    }
  }
  return '';
}

function setSettingValue_(key, value, comment) {
  ensureServiceSheets_();
  const sheet = getSettingsSheet_();
  const wanted = normalizeSettingKey_(key);
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = sheet
    .getRange(1, 1, lastRow, Math.max(2, Math.min(sheet.getLastColumn(), 3)))
    .getDisplayValues();

  let targetRow = 0;
  for (let row = 0; row < values.length; row++) {
    if (normalizeSettingKey_(values[row][0]) === wanted) {
      targetRow = row + 1;
      break;
    }
  }

  if (!targetRow) {
    targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1).setValue(key);
  }

  sheet.getRange(targetRow, 2).setValue(String(value == null ? '' : value));
  if (comment && sheet.getMaxColumns() >= 3) {
    sheet.getRange(targetRow, 3).setValue(comment);
  }
  return targetRow;
}

function deleteSetting_(key) {
  const sheet = getSettingsSheet_();
  if (!sheet || sheet.getLastRow() < 1) return false;

  const wanted = normalizeSettingKey_(key);
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getDisplayValues();
  for (let row = values.length - 1; row >= 0; row--) {
    if (normalizeSettingKey_(values[row][0]) === wanted) {
      sheet.deleteRow(row + 1);
      return true;
    }
  }
  return false;
}

function removeObsoleteSettings_() {
  // v3.12: структура листов фиксирована кодом, отдельный архив не используется.
  [
    'archiveFolderId',
    'dataSheetName',
    'emailSheetName',
    'logSheetName',
    'qrImageFileId'
  ].forEach(deleteSetting_);

  // Старые тарифы удаляем из «Настройки» только когда на листе «Тарифы»
  // уже есть хотя бы одна строка данных. Так обновление не уничтожит
  // единственную копию тарифов в старой таблице.
  const tariffSheet = SpreadsheetApp.getActive().getSheetByName(DNP_TARIFF_SHEET_NAME);
  if (tariffSheet && tariffSheet.getLastRow() >= 2) {
    [
      't1Rate',
      't2Rate',
      't3Rate',
      'waterRate',
      'Тариф Т1',
      'Тариф Т2',
      'Тариф Т3',
      'Тариф водоотведения'
    ].forEach(deleteSetting_);
  }
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

function onOpen() {
  removeLastPdfSetting_();
  removeObsoleteSettings_();
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('ДНП')
    .addSubMenu(ui.createMenu('Квитанции')
      .addItem('Сформировать PDF', 'showPdfDialog')
      .addItem('Открыть папку месяца', 'openCurrentMonthFolder')
      .addItem('Открыть журнал PDF', 'openPdfLogSheet')
      .addItem('Очистить PDF', 'clearGeneratedPdfs'))
    .addSubMenu(ui.createMenu('Почта')
      .addItem('Отправить квитанции', 'sendReceipts'))
    .addSubMenu(ui.createMenu('Настройка')
      .addItem('Первичная настройка', 'showInitialSetupDialog')
      .addSeparator()
      .addItem('Проставить формулы кВтч', 'fillKwhFormulas')
      .addItem('Проставить формулы суммы', 'fillSumFormulasFromTariffs')
      .addItem('Оформить лист', 'formatYearSheetUX')
      .addItem('Добавить строчку услуги', 'showAddServiceRowDialog')
      .addItem('Создать шаблон под текущий формат', 'createReceiptTemplateForCurrentFormat')
      .addItem('Открыть шаблон квитанции', 'openReceiptTemplate')
      .addItem('Создать новый лист-год', 'showCreateYearDialog')
      .addSeparator()
      .addItem('Открыть настройки', 'openSettingsSheet')
      .addItem('Открыть почты', 'openEmailsSheet')
      .addItem('Заполнить тестовыми адресами', 'fillTestEmails')
      .addItem('Открыть журнал', 'openJournalSheet')
      .addItem('Открыть журнал PDF', 'openPdfLogSheet')
      .addItem('Скрыть служебные листы', 'hideServiceSheets')
      .addSeparator()
      .addItem('Очистить журнал PDF', 'clearPdfLog')
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
  const passwordField = isOperationPasswordRequired_()
    ? '<label for="password">Пароль</label><input id="password" type="password" autocomplete="off">'
    : '';

  const html = HtmlService.createHtmlOutput(`
<!doctype html><html><head><base target="_top"><style>
body{font:14px Arial,sans-serif;padding:18px;color:#202124}h2{margin:0 0 16px;font-size:18px}
label{display:block;margin:12px 0 6px;font-weight:600}select,input{width:100%;box-sizing:border-box;padding:9px;border:1px solid #dadce0;border-radius:6px}
.info{background:#f8f9fa;padding:10px;border-radius:6px;margin-top:14px;line-height:1.5}.buttons{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}
button{padding:9px 14px;border:0;border-radius:6px;cursor:pointer}.primary{background:#1a73e8;color:#fff}.secondary{background:#f1f3f4}#status{margin-top:12px;min-height:18px;color:#5f6368}
</style></head><body><h2>Формирование PDF</h2>
<label for="year">Год</label><select id="year" onchange="refreshCount()">${options}</select>
<label for="month">Месяц</label><select id="month">
<option value="1">Январь</option><option value="2">Февраль</option><option value="3">Март</option><option value="4">Апрель</option>
<option value="5">Май</option><option value="6">Июнь</option><option value="7">Июль</option><option value="8">Август</option>
<option value="9">Сентябрь</option><option value="10">Октябрь</option><option value="11">Ноябрь</option><option value="12">Декабрь</option></select>
${passwordField}
<div class="info">Найдено участков: <b id="count">…</b></div><div id="status"></div>
<div class="buttons"><button class="secondary" onclick="google.script.host.close()">Отмена</button><button class="primary" id="submit" onclick="submitForm()">Сформировать</button></div>
<script>
document.getElementById('year').value='${defaultYear}';document.getElementById('month').value='${currentMonth}';
function refreshCount(){const year=document.getElementById('year').value;google.script.run.withSuccessHandler(c=>document.getElementById('count').textContent=c).withFailureHandler(e=>document.getElementById('status').textContent=e.message).countPlotsForYear(year)}
function submitForm(){const b=document.getElementById('submit'),s=document.getElementById('status'),p=document.getElementById('password');b.disabled=true;s.textContent='Формирование запущено…';google.script.run.withSuccessHandler(r=>{document.body.innerHTML='<h2>Формирование завершено</h2><div class="info">Создано PDF: <b>'+Number(r.created||0)+'</b><br>Ошибок: <b>'+Number(r.failed||0)+'</b></div><div class="buttons"><button class="primary" onclick="google.script.host.close()">Закрыть</button></div>'}).withFailureHandler(e=>{s.textContent='Ошибка: '+e.message;b.disabled=false}).startPdfGenerationFromDialog(Number(document.getElementById('year').value),Number(document.getElementById('month').value),p?p.value:'')}
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
  showDriveLinkDialog_('Папка месяца', monthFolder.name, dnpDriveFolderUrl_(monthFolder));
}

function findChildFolderByNames_(parent, names) {
  const normalized = names.map(name => String(name).trim().toLowerCase());
  const folders = dnpDriveChildren_(parent, DNP_DRIVE_FOLDER_MIME);
  for (let i = 0; i < folders.length; i++) {
    if (normalized.includes(String(folders[i].name || '').trim().toLowerCase())) return folders[i];
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

function showInitialSetupDialog() {
  const html = HtmlService.createHtmlOutput(`
<!doctype html><html><head><base target="_top"><style>
body{font:14px Arial,sans-serif;padding:18px;color:#202124}h2{margin:0 0 16px;font-size:18px}
label{display:block;margin:12px 0 6px;font-weight:600}input,select{width:100%;box-sizing:border-box;padding:9px;border:1px solid #dadce0;border-radius:6px}
.buttons{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}button{padding:9px 14px;border:0;border-radius:6px;cursor:pointer}
.primary{background:#1a73e8;color:#fff}.secondary{background:#f1f3f4}#status{margin-top:12px;color:#5f6368;min-height:18px}
</style></head><body><h2>Первичная настройка</h2>
<label for="password">Пароль администратора</label><input id="password" type="password" autocomplete="off" autofocus>
<label for="mode">Папка квитанций</label><select id="mode"><option value="reuse">Использовать существующую</option><option value="recreate">Создать новую папку</option></select>
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
  ensureServiceSheets_();
  removeObsoleteSettings_();

  const folderName = getSettingValue_('mainFolderName') || 'Квитанции ДНП Комфорт';
  const pdfFolder = getOrCreateSetupFolder_(folderName, mode === 'recreate');

  setSettingValue_(
    'outputFolderId',
    pdfFolder.id,
    'Корневая папка квитанций. Она же архив: внутри создаются папки по годам и месяцам.'
  );

  applySevenRowBandingToYearSheets_();
  hideServiceSheets();

  ss.toast('Первичная настройка завершена', 'ДНП', 7);
  return { ok: true, message: 'Готово. Папка квитанций: ' + pdfFolder.name + '.' };
}

function getOrCreateSetupFolder_(name, forceCreate) {
  const configuredId = extractGoogleFileId_(getSettingValue_('outputFolderId'));

  if (!forceCreate && configuredId) {
    try {
      const existing = dnpDriveGet_(configuredId);
      if (existing.trashed) throw new Error('папка находится в корзине');
      if (existing.mimeType !== DNP_DRIVE_FOLDER_MIME) throw new Error('ID указывает не на папку');
      return existing;
    } catch (error) {
      throw new Error(
        'Папка из Настройки → outputFolderId недоступна: ' + (error.message || error) +
        '. ID не изменён. Если хотите создать новую папку, выберите режим «Создать новую папку».'
      );
    }
  }

  if (!forceCreate && !configuredId) {
    const created = dnpDriveCreateFolder_(name, '', { [DNP_ROOT_APP_PROPERTY]: '1' });
    return created;
  }

  const suffix = configuredId
    ? ' ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm-ss')
    : '';
  return dnpDriveCreateFolder_(name + suffix, '', { [DNP_ROOT_APP_PROPERTY]: '1' });
}

function ensureServiceSheets_() {
  const ss = SpreadsheetApp.getActive();

  if (!ss.getSheetByName(DNP_SERVICE_SHEETS.settings)) {
    const sheet = ss.insertSheet(DNP_SERVICE_SHEETS.settings);
    const rows = [
      ['Параметр', 'Значение', 'Комментарий'],
      ['sendMode', 'TEST', 'TEST — реальные получатели не получают письма; LIVE — отправка реальным адресатам.'],
      ['testEmail', '', 'Если указан, в режиме TEST все выбранные квитанции отправляются только на этот адрес.'],
      ['emailSubject', 'Квитанция ДНП Комфорт: участок {{PLOT}}, {{MONTH_NAME}} {{YEAR}}', 'Тема письма.'],
      ['mainFolderName', 'Квитанции ДНП Комфорт', 'Имя корневой папки PDF.'],
      ['templateDocName', 'Шаблон квитанции ДНП Комфорт', 'Базовое имя Google Docs шаблона.'],
      ['templateDocId', '', 'ID Google Docs шаблона квитанции.'],
      ['outputFolderId', '', 'ID корневой папки квитанций; она же архив.'],
      ['emailBody', 'Здравствуйте{{NAME_PART}}!\n\nНаправляем квитанцию по участку № {{PLOT}} за {{MONTH_NAME}} {{YEAR}} года.\n\nС уважением, ДНП «Комфорт».', 'Текст письма.']
    ];
    sheet.getRange(1, 1, rows.length, 3).setValues(rows);
    sheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName(DNP_TARIFF_SHEET_NAME)) {
    const sheet = ss.insertSheet(DNP_TARIFF_SHEET_NAME);
    sheet.getRange('A1:F1').setValues([['Год', 'Месяц', 'T1', 'T2', 'T3', 'Водоотведение']]);
    sheet.setFrozenRows(1);
  }

  if (!ss.getSheetByName(DNP_SERVICE_SHEETS.emails)) {
    const sheet = ss.insertSheet(DNP_SERVICE_SHEETS.emails);
    sheet.getRange('A1:D1').setValues([['Участок', 'Email', 'ФИО', 'Отправить']]);
    sheet.setFrozenRows(1);
    syncEmailSheetPlots_();
  } else {
    ensureEmailSheetHeaders_(ss.getSheetByName(DNP_SERVICE_SHEETS.emails));
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
    if (/^тариф(?:ы)?$/.test(label)) { insideBlock = true; continue; }
    if (!insideBlock || !label) continue;

    const range = sheet.getRange(row, 3, 1, lastColumn - 2);
    const values = range.getValues()[0];
    const formulas = range.getFormulas()[0];
    let changed = false;
    for (let column = 0; column < values.length; column++) {
      if (!formulas[column] && values[column] !== '') { values[column] = ''; changed = true; }
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
      if (rowCount <= 0) return;
      sheet.getRange(startRow, 1, rowCount, lastColumn).setBackground(colors[blockIndex % colors.length]);
      sheet.getRange(startRow, 1, 1, lastColumn).setBorder(
        true, null, null, null, null, null,
        '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK
      );
    });
    return;
  }

  for (let row = 2, blockIndex = 0; row <= lastRow; row += 7, blockIndex++) {
    const rowCount = Math.min(7, lastRow - row + 1);
    sheet.getRange(row, 1, rowCount, lastColumn).setBackground(colors[blockIndex % colors.length]);
    sheet.getRange(row, 1, 1, lastColumn).setBorder(
      true, null, null, null, null, null,
      '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK
    );
  }
}

function showAddServiceRowDialog() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!/^\d{4}$/.test(sheet.getName())) {
    SpreadsheetApp.getUi().alert('Сначала откройте лист нужного года, например 2026.');
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
<div class="note">Если услуги нет, новая строка будет вставлена перед строкой «Сумма» во всех блоках текущего листа.</div>
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
  if (!/^\d{4}$/.test(sheet.getName())) throw new Error('Текущий лист должен называться годом.');

  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), 2);
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
    if (blockLabels.includes(normalizedService)) { skipped++; continue; }

    let sumRow = -1;
    for (let row = endRow; row >= startRow; row--) {
      if (/^(сумма|итого)/.test(labels[row - 1])) { sumRow = row; break; }
    }
    if (sumRow === -1) throw new Error('В блоке со строки ' + startRow + ' не найдена строка «Сумма».');
    operations.push(sumRow);
  }

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
  return { ok: true, added: operations.length, skipped, message: 'Готово. Добавлено: ' + operations.length + ', уже существовало: ' + skipped + '.' };
}

const DNP_PAYMENT_TABLE_WIDTHS = {
  name: 150,
  current: 66,
  previous: 72,
  usage: 58,
  rate: 58,
  amount: 82,
};


/**
 * Google Docs API helpers.
 * Важно: здесь используется только Advanced Docs API. Он умеет работать
 * с документами по scope drive.file, поэтому скрипту не нужен общий
 * https://www.googleapis.com/auth/documents.
 */
function dnpDocsGet_(documentId) {
  return Docs.Documents.get(String(documentId));
}

function dnpDocsBatchUpdate_(documentId, requests) {
  if (!requests || !requests.length) return;
  Docs.Documents.batchUpdate({ requests: requests }, String(documentId));
}

function dnpDocsWriteDefaultTemplate_(documentId) {
  const title = 'ДНП «Дачный поселок «КОМФОРТ»';
  const plotLine = 'Участок № {{PLOT}}    {{MONTH_NAME}} {{YEAR}} года';
  const totalLine = 'Сумма оплаты: {{TOTAL}} руб.';
  const marker = '{{PAYMENT_TABLE}}';
  const text = title + '\n\n' + plotLine + '\n' + totalLine + '\n\n' + marker + '\n';

  const titleStart = 1;
  const titleEnd = titleStart + title.length;
  const plotStart = titleEnd + 2;
  const plotEnd = plotStart + plotLine.length;
  const totalStart = plotEnd + 1;
  const totalEnd = totalStart + totalLine.length;

  dnpDocsBatchUpdate_(documentId, [
    { insertText: { location: { index: 1 }, text: text } },
    {
      updateTextStyle: {
        range: { startIndex: titleStart, endIndex: titleEnd },
        textStyle: { bold: true },
        fields: 'bold'
      }
    },
    {
      updateParagraphStyle: {
        range: { startIndex: titleStart, endIndex: titleEnd + 1 },
        paragraphStyle: { alignment: 'CENTER' },
        fields: 'alignment'
      }
    },
    {
      updateParagraphStyle: {
        range: { startIndex: plotStart, endIndex: plotEnd + 1 },
        paragraphStyle: { alignment: 'CENTER' },
        fields: 'alignment'
      }
    },
    {
      updateTextStyle: {
        range: { startIndex: totalStart, endIndex: totalEnd },
        textStyle: { bold: true },
        fields: 'bold'
      }
    },
    {
      updateParagraphStyle: {
        range: { startIndex: totalStart, endIndex: totalEnd + 1 },
        paragraphStyle: { alignment: 'CENTER' },
        fields: 'alignment'
      }
    }
  ]);
}

function dnpDocsReplaceAllText_(documentId, replacements) {
  const requests = Object.keys(replacements).map(marker => ({
    replaceAllText: {
      containsText: { text: marker, matchCase: true },
      replaceText: String(replacements[marker] == null ? '' : replacements[marker])
    }
  }));
  dnpDocsBatchUpdate_(documentId, requests);
}

function dnpDocsFindTextRange_(documentId, needle) {
  const doc = dnpDocsGet_(documentId);
  const content = doc && doc.body && doc.body.content ? doc.body.content : [];

  for (let i = 0; i < content.length; i++) {
    const structural = content[i];
    const paragraph = structural.paragraph;
    if (!paragraph || !paragraph.elements) continue;

    let combined = '';
    const pieces = [];
    paragraph.elements.forEach(element => {
      const run = element.textRun;
      if (!run || run.content == null) return;
      const value = String(run.content);
      pieces.push({
        textStart: combined.length,
        textEnd: combined.length + value.length,
        docStart: Number(element.startIndex),
        docEnd: Number(element.endIndex),
        text: value,
      });
      combined += value;
    });

    const foundAt = combined.indexOf(String(needle));
    if (foundAt < 0) continue;

    const foundEnd = foundAt + String(needle).length;
    let startIndex = null;
    let endIndex = null;

    for (let p = 0; p < pieces.length; p++) {
      const piece = pieces[p];
      if (startIndex == null && foundAt >= piece.textStart && foundAt <= piece.textEnd) {
        startIndex = piece.docStart + (foundAt - piece.textStart);
      }
      if (foundEnd >= piece.textStart && foundEnd <= piece.textEnd) {
        endIndex = piece.docStart + (foundEnd - piece.textStart);
        break;
      }
    }

    if (startIndex != null && endIndex != null) {
      return { startIndex: startIndex, endIndex: endIndex };
    }
  }

  return null;
}

function dnpDocsFindNearestTable_(documentId, nearIndex) {
  const doc = dnpDocsGet_(documentId);
  const content = doc && doc.body && doc.body.content ? doc.body.content : [];
  let best = null;
  let bestDistance = Number.MAX_SAFE_INTEGER;

  content.forEach(element => {
    if (!element.table) return;
    const startIndex = Number(element.startIndex || 0);
    const distance = Math.abs(startIndex - Number(nearIndex || 0));
    if (distance < bestDistance) {
      best = element;
      bestDistance = distance;
    }
  });

  return best;
}

function dnpDocsCellInsertIndex_(cell) {
  if (!cell || !cell.content || !cell.content.length) return null;
  for (let i = 0; i < cell.content.length; i++) {
    const structural = cell.content[i];
    if (structural.paragraph && structural.startIndex != null) {
      return Number(structural.startIndex);
    }
  }
  return null;
}

function dnpDocsCellTextRange_(cell) {
  if (!cell || !cell.content || !cell.content.length) return null;
  let start = null;
  let end = null;

  cell.content.forEach(structural => {
    if (!structural.paragraph) return;
    if (start == null && structural.startIndex != null) start = Number(structural.startIndex);
    if (structural.endIndex != null) end = Number(structural.endIndex);
  });

  if (start == null || end == null) return null;
  // Последний символ абзаца — перевод строки; его не нужно делать жирным.
  end = Math.max(start, end - 1);
  return end > start ? { startIndex: start, endIndex: end } : null;
}

function dnpDocsInsertPaymentTable_(documentId, paymentRows, total) {
  const marker = '{{PAYMENT_TABLE}}';
  const markerRange = dnpDocsFindTextRange_(documentId, marker);
  if (!markerRange) {
    throw new Error('В Google Docs-шаблоне не найден маркер {{PAYMENT_TABLE}}.');
  }

  // Сначала удаляем маркер, оставляя сам абзац на месте.
  dnpDocsBatchUpdate_(documentId, [{
    deleteContentRange: {
      range: {
        startIndex: markerRange.startIndex,
        endIndex: markerRange.endIndex
      }
    }
  }]);

  const rows = [[
    'Наименование платежа',
    'Текущее',
    'Предыдущее',
    'Объём',
    'Тариф',
    'Сумма к оплате',
  ]].concat(paymentRows || []);

  rows.push([
    'ИТОГО К ОПЛАТЕ',
    '', '', '', '',
    formatReceiptMoney_(total) + ' руб.'
  ]);

  // Docs API вставляет перед таблицей перевод строки; сама таблица обычно
  // начинается с markerRange.startIndex + 1. Точный индекс затем читаем обратно.
  dnpDocsBatchUpdate_(documentId, [{
    insertTable: {
      rows: rows.length,
      columns: 6,
      location: { index: markerRange.startIndex }
    }
  }]);

  let tableElement = dnpDocsFindNearestTable_(documentId, markerRange.startIndex + 1);
  if (!tableElement || !tableElement.table || !tableElement.table.tableRows) {
    throw new Error('Google Docs API не вернул созданную таблицу.');
  }

  // Заполняем ячейки с конца документа к началу. Так вставка текста не сдвигает
  // индексы ещё не обработанных ячеек.
  const insertRequests = [];
  const tableRows = tableElement.table.tableRows;
  for (let r = 0; r < rows.length && r < tableRows.length; r++) {
    const cells = tableRows[r].tableCells || [];
    for (let c = 0; c < rows[r].length && c < cells.length; c++) {
      const value = String(rows[r][c] == null ? '' : rows[r][c]);
      if (!value) continue;
      const index = dnpDocsCellInsertIndex_(cells[c]);
      if (index == null) continue;
      insertRequests.push({
        index: index,
        request: { insertText: { location: { index: index }, text: value } }
      });
    }
  }

  insertRequests.sort((a, b) => b.index - a.index);
  dnpDocsBatchUpdate_(documentId, insertRequests.map(item => item.request));

  // Повторно читаем структуру уже после вставки текста и применяем оформление.
  tableElement = dnpDocsFindNearestTable_(documentId, markerRange.startIndex + 1);
  const table = tableElement.table;
  const tableStartIndex = Number(tableElement.startIndex);
  const lastRowIndex = Math.max(0, rows.length - 1);
  const styleRequests = [];

  const widths = [
    DNP_PAYMENT_TABLE_WIDTHS.name,
    DNP_PAYMENT_TABLE_WIDTHS.current,
    DNP_PAYMENT_TABLE_WIDTHS.previous,
    DNP_PAYMENT_TABLE_WIDTHS.usage,
    DNP_PAYMENT_TABLE_WIDTHS.rate,
    DNP_PAYMENT_TABLE_WIDTHS.amount,
  ];

  widths.forEach((width, columnIndex) => {
    styleRequests.push({
      updateTableColumnProperties: {
        tableStartLocation: { index: tableStartIndex },
        columnIndices: [columnIndex],
        tableColumnProperties: {
          width: { magnitude: Number(width), unit: 'PT' }
        },
        fields: 'width'
      }
    });
  });

  // Жирный заголовок.
  const headerCells = table.tableRows && table.tableRows[0]
    ? (table.tableRows[0].tableCells || [])
    : [];
  headerCells.forEach(cell => {
    const range = dnpDocsCellTextRange_(cell);
    if (!range) return;
    styleRequests.push({
      updateTextStyle: {
        range: range,
        textStyle: { bold: true },
        fields: 'bold'
      }
    });
  });

  // Жирный итог и выравнивание подписи справа.
  const totalRow = table.tableRows && table.tableRows[lastRowIndex]
    ? table.tableRows[lastRowIndex]
    : null;
  if (totalRow && totalRow.tableCells && totalRow.tableCells.length >= 6) {
    const labelRange = dnpDocsCellTextRange_(totalRow.tableCells[0]);
    const amountRange = dnpDocsCellTextRange_(totalRow.tableCells[5]);

    [labelRange, amountRange].forEach(range => {
      if (!range) return;
      styleRequests.push({
        updateTextStyle: {
          range: range,
          textStyle: { bold: true },
          fields: 'bold'
        }
      });
    });

    if (labelRange) {
      styleRequests.push({
        updateParagraphStyle: {
          range: labelRange,
          paragraphStyle: { alignment: 'END' },
          fields: 'alignment'
        }
      });
    }

    // Объединяем первые пять ячеек последней строки, как в старой версии.
    styleRequests.push({
      mergeTableCells: {
        tableRange: {
          tableCellLocation: {
            tableStartLocation: { index: tableStartIndex },
            rowIndex: lastRowIndex,
            columnIndex: 0
          },
          rowSpan: 1,
          columnSpan: 5
        }
      }
    });
  }

  dnpDocsBatchUpdate_(documentId, styleRequests);
}

function createReceiptTemplate() { return createReceiptTemplateForCurrentFormat(); }

function createReceiptTemplateForCurrentFormat() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!/^\d{4}$/.test(sheet.getName())) {
    ui.alert('Сначала откройте лист нужного года, например 2026.');
    return;
  }

  const services = getCurrentFormatServiceLabels_(sheet);
  if (!services.length) { ui.alert('Не удалось определить строки услуг.'); return; }

  const answer = ui.alert(
    'Создать шаблон под текущий формат?',
    'Найдены строки:\n\n• ' + services.join('\n• ') + '\n\nБудет создан новый Google Документ внутри папки ДНП.',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  const root = getDnpPdfFolder_();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm');
  const templateBaseName = getSettingValue_('templateDocName') || 'Шаблон квитанции ДНП Комфорт';
  const name = templateBaseName + ' ' + timestamp;
  const appProperties = {};
  appProperties[DNP_TEMPLATE_APP_PROPERTY] = '1';
  const file = dnpDriveCreateGoogleDoc_(name, root, appProperties);

  dnpDocsWriteDefaultTemplate_(file.id);

  saveTemplateId_(file.id);
  showDriveLinkDialog_('Шаблон создан', name, dnpDriveDocUrl_(file));
}

function getCurrentFormatServiceLabels_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const labels = sheet.getRange(1, 2, lastRow, 1).getDisplayValues().flat()
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
  showDriveLinkDialog_('Шаблон квитанции', file.name, dnpDriveDocUrl_(file));
}





function generatePdfsForMonth(year, month, password) {
  requireOperationPassword_(password);
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
  const rates = getReceiptRates_(year, month);
  const previousDecemberReadings = month === 1 ? getPreviousDecemberReadings_(year) : null;
  if (!blocks.length) throw new Error('Не найдены строки «Тариф» в столбце B.');

  let created = 0;
  let failed = 0;
  const errors = [];

  blocks.forEach((block, index) => {
    let tempFileId = '';
    ss.toast('Формируется ' + (index + 1) + ' из ' + blocks.length + ': участок ' + block.plot, 'ДНП', 5);
    try {
      const receipt = buildReceiptData_(sheet, block, year, month, rates, previousDecemberReadings);
      const fileName = 'Квитанция_участок_' + sanitizePdfFileName_(block.plot) + '_' + year + '_' + String(month).padStart(2, '0') + '.pdf';
      trashFilesByName_(monthFolder, fileName);

      const tempFile = dnpDriveCopy_(
        templateFile,
        'Временная квитанция ' + block.plot,
        monthFolder
      );
      tempFileId = tempFile.id;

      fillReceiptTemplate_(tempFileId, receipt);

      const pdfBlob = dnpDriveExportPdf_(tempFileId, fileName);
      dnpDriveCreateBlobFile_(monthFolder, fileName, pdfBlob, DNP_DRIVE_PDF_MIME);

      dnpDriveTrash_(tempFileId);
      tempFileId = '';
      created++;
      if (DNP_PDF_SLEEP_MS > 0) Utilities.sleep(DNP_PDF_SLEEP_MS);
    } catch (error) {
      failed++;
      errors.push('Участок ' + block.plot + ': ' + (error.message || error));
      if (tempFileId) {
        try { dnpDriveTrash_(tempFileId); } catch (cleanupError) {}
      }
    }
  });

  const message = failed
    ? 'Создано PDF: ' + created + '. Ошибок: ' + failed + '. Первая ошибка: ' + errors[0]
    : 'Создано PDF: ' + created + '. Папка: ' + year + '/' + monthFolderName + '.';
  ss.toast(message, 'ДНП', 10);
  return {
    ok: failed === 0,
    created,
    failed,
    folderId: monthFolder.id,
    folderUrl: dnpDriveFolderUrl_(monthFolder),
    message
  };
}

function getReceiptBlocks_(sheet) {
  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(1, 1, lastRow, 2).getDisplayValues();
  const starts = [];
  for (let index = 1; index < rows.length; index++) {
    const plot = String(rows[index][0] || '').trim();
    const label = normalizeReceiptLabel_(rows[index][1]);
    if (/^тариф(?:ы)?$/.test(label)) {
      if (!plot) throw new Error('В строке ' + (index + 1) + ' нет номера участка.');
      starts.push({ startRow: index + 1, plot });
    }
  }
  return starts.map((item, index) => ({
    plot: item.plot,
    startRow: item.startRow,
    endRow: index + 1 < starts.length ? starts[index + 1].startRow - 1 : lastRow,
  }));
}

function getPreviousDecemberReadings_(year) {
  const previousSheet = SpreadsheetApp.getActive().getSheetByName(String(Number(year) - 1));
  if (!previousSheet) return {};

  const decemberColumn = 14; // A, B, затем январь–декабрь в C:N.
  const lastColumn = Math.max(previousSheet.getLastColumn(), decemberColumn);
  const readingsByPlot = {};

  getReceiptBlocks_(previousSheet).forEach(block => {
    const rowCount = block.endRow - block.startRow + 1;
    const values = previousSheet.getRange(block.startRow, 1, rowCount, lastColumn).getDisplayValues();
    const byLabel = {};
    values.forEach(row => {
      const label = normalizeReceiptLabel_(row[1]);
      if (label) byLabel[label] = parseReceiptNumber_(row[decemberColumn - 1]);
    });
    readingsByPlot[String(block.plot).trim()] = byLabel;
  });

  return readingsByPlot;
}

function buildReceiptData_(sheet, block, year, month, rates, previousDecemberReadings) {
  const monthColumn = month + 2;
  const previousMonthColumn = monthColumn - 1;
  const rowCount = block.endRow - block.startRow + 1;
  const lastColumn = Math.max(sheet.getLastColumn(), monthColumn);
  const values = sheet.getRange(block.startRow, 1, rowCount, lastColumn).getDisplayValues();
  const byLabel = {};
  values.forEach(row => { const label = normalizeReceiptLabel_(row[1]); if (label) byLabel[label] = row; });
  const previousByLabel = month === 1
    ? ((previousDecemberReadings || {})[String(block.plot).trim()] || {})
    : null;

  const paymentRows = [];
  let calculatedTotal = 0;
  ['т1', 'т2', 'т3'].forEach((label, index) => {
    const row = byLabel[label];
    if (!row) return;
    const current = parseReceiptNumber_(row[monthColumn - 1]);
    const previous = month > 1
      ? parseReceiptNumber_(row[previousMonthColumn - 1])
      : (previousByLabel[label] === undefined ? null : previousByLabel[label]);
    const usage = current !== null && previous !== null ? current - previous : null;
    const rate = rates['T' + (index + 1)];
    if (usage !== null && rate === null) {
      throw new Error('На листе «Тарифы» не заполнен тариф T' + (index + 1) + ' для выбранного месяца.');
    }
    const amount = usage !== null && rate !== null ? usage * rate : null;
    if (amount !== null) calculatedTotal += amount;
    paymentRows.push(['Электроэнергия Т' + (index + 1), formatReceiptValue_(current), formatReceiptValue_(previous), formatReceiptValue_(usage), formatReceiptMoney_(rate), formatReceiptMoney_(amount)]);
  });

  const waterRow = findReceiptRow_(byLabel, [/^водоотвед/, /^вода$/]);
  if (waterRow) {
    const current = parseReceiptNumber_(waterRow[monthColumn - 1]);
    const previousWater = month === 1
      ? findReceiptRow_(previousByLabel, [/^водоотвед/, /^вода$/])
      : null;
    const previous = month > 1
      ? parseReceiptNumber_(waterRow[previousMonthColumn - 1])
      : (previousWater === undefined ? null : previousWater);
    const usage = current !== null && previous !== null ? current - previous : null;
    if (usage !== null && rates.WATER === null) {
      throw new Error('На листе «Тарифы» не заполнен тариф водоотведения для выбранного месяца.');
    }
    const amount = usage !== null && rates.WATER !== null ? usage * rates.WATER : null;
    if (amount !== null) calculatedTotal += amount;
    paymentRows.push(['Водоотведение', formatReceiptValue_(current), formatReceiptValue_(previous), formatReceiptValue_(usage), formatReceiptMoney_(rates.WATER), formatReceiptMoney_(amount)]);
  }

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

  // Итог PDF всегда складывается из рассчитанных сумм строк таблицы:
  // объём × тариф для ресурсов плюс фиксированные платежи и взносы.
  const total = calculatedTotal;
  return { plot: block.plot, year, month, monthName: getRussianMonthName_(month), total, paymentRows };
}

function fillReceiptTemplate_(documentId, receipt) {
  const replacements = {
    '{{PLOT}}': String(receipt.plot == null ? '' : receipt.plot),
    '{{YEAR}}': String(receipt.year == null ? '' : receipt.year),
    '{{MONTH}}': String(receipt.month).padStart(2, '0'),
    '{{MONTH_NAME}}': String(receipt.monthName == null ? '' : receipt.monthName),
    '{{TOTAL}}': formatReceiptMoney_(receipt.total),
  };

  dnpDocsReplaceAllText_(documentId, replacements);
  dnpDocsInsertPaymentTable_(documentId, receipt.paymentRows || [], receipt.total);
}



function getReceiptTemplateFile_() {
  const templateId = getStoredTemplateId_();
  if (!templateId) throw new Error('На листе «Настройки» не заполнен параметр templateDocId.');
  try {
    const file = dnpDriveGet_(templateId);
    if (file.trashed) throw new Error('Файл находится в корзине.');
    if (file.mimeType !== DNP_DRIVE_DOC_MIME) throw new Error('Файл templateDocId не является Google Docs.');
    return file;
  } catch (error) {
    throw new Error('Не удалось открыть шаблон из строки templateDocId. ' + (error.message || error));
  }
}

function getStoredTemplateId_() {
  return extractGoogleFileId_(getSettingValue_('templateDocId'));
}

function saveTemplateId_(templateId) {
  const cleanId = extractGoogleFileId_(templateId);
  setSettingValue_(
    'templateDocId',
    cleanId,
    'ID Google Docs шаблона квитанции.'
  );
}

function extractGoogleFileId_(value) {
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';
  const match = text.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : text;
}

function getReceiptRates_(year, month) {
  year = Number(year);
  month = Number(month);
  if (!Number.isInteger(year) || year < 2000 || year > 2100) throw new Error('Некорректный год для тарифа.');
  if (!Number.isInteger(month) || month < 1 || month > 12) throw new Error('Некорректный месяц для тарифа.');

  const sheet = SpreadsheetApp.getActive().getSheetByName(DNP_TARIFF_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error('Лист «' + DNP_TARIFF_SHEET_NAME + '» отсутствует или не содержит тарифов.');
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const targetKey = year * 100 + month;
  let selected = null;

  rows.forEach((row, index) => {
    const tariffYear = Number(row[0]);
    const tariffMonth = Number(row[1]);
    if (!Number.isInteger(tariffYear) || !Number.isInteger(tariffMonth) || tariffMonth < 1 || tariffMonth > 12) return;

    const key = tariffYear * 100 + tariffMonth;
    if (key > targetKey || (selected && key <= selected.key)) return;

    selected = {
      key,
      row: index + 2,
      T1: parseReceiptNumber_(row[2]),
      T2: parseReceiptNumber_(row[3]),
      T3: parseReceiptNumber_(row[4]),
      WATER: parseReceiptNumber_(row[5])
    };
  });

  if (!selected) {
    throw new Error('На листе «' + DNP_TARIFF_SHEET_NAME + '» не найден тариф, действующий на ' + String(month).padStart(2, '0') + '.' + year + '.');
  }

  ['T1', 'T2', 'T3', 'WATER'].forEach(key => {
    if (selected[key] === null) {
      const label = key === 'WATER' ? 'Водоотведение' : key;
      throw new Error('На листе «' + DNP_TARIFF_SHEET_NAME + '» в строке ' + selected.row + ' не заполнен тариф «' + label + '».');
    }
  });

  return { T1: selected.T1, T2: selected.T2, T3: selected.T3, WATER: selected.WATER };
}

function normalizeSettingKey_(value) {
  return String(value == null ? '' : value).replace(/\u00a0/g, ' ').trim().toLowerCase().replace(/[\s_\-.,:;()]+/g, '');
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

function formatReceiptValue_(value) {
  return value === null || value === undefined || !Number.isFinite(Number(value)) ? '' : Number(value).toLocaleString('ru-RU', { maximumFractionDigits: 3 });
}

function formatReceiptMoney_(value) {
  return value === null || value === undefined || !Number.isFinite(Number(value)) ? '' : Number(value).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function getOrCreatePdfChildFolder_(parent, name) {
  const found = dnpDriveFindByName_(parent, name, DNP_DRIVE_FOLDER_MIME);
  return found.length ? found[0] : dnpDriveCreateFolder_(name, parent, {});
}
function getRussianMonthName_(month) { return ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'][month - 1]; }
function sanitizePdfFileName_(value) { return String(value).trim().replace(/[\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_'); }
function trashFilesByName_(folder, fileName) {
  const files = dnpDriveFindByName_(folder, fileName, null);
  files.forEach(file => {
    try { dnpDriveTrash_(file); } catch (error) {}
  });
}

const DNP_PDF_LOG_SHEET = 'Журнал PDF';

function generatePdfsForMonthWithLog(year, month, password) {
  const runId = Utilities.getUuid().slice(0, 8);
  const startedAt = new Date();
  try {
    const result = generatePdfsForMonth(year, month, password);
    if (DNP_PDF_LOG_ENABLED) appendPdfLog_(runId, 'ИТОГ', year, month, '', result && result.ok !== false ? 'OK' : 'ERROR', (result.message || '') + '; длительность=' + Math.round((new Date().getTime() - startedAt.getTime()) / 1000) + ' сек.');
    return result;
  } catch (error) {
    if (DNP_PDF_LOG_ENABLED) appendPdfLog_(runId, 'ОШИБКА', year, month, '', 'ERROR', getDetailedErrorText_(error));
    throw error;
  }
}

function ensurePdfLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(DNP_PDF_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DNP_PDF_LOG_SHEET);
    sheet.getRange(1, 1, 1, 8).setValues([['Дата','Запуск','Этап','Год','Месяц','Участок','Статус','Подробности']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function appendPdfLog_(runId, stage, year, month, plot, status, details) { ensurePdfLogSheet_().appendRow([new Date(),runId||'',stage||'',year||'',month||'',plot||'',status||'',details||'']); }
function openPdfLogSheet() { const sheet = ensurePdfLogSheet_(); sheet.showSheet(); SpreadsheetApp.getActive().setActiveSheet(sheet); }
function clearPdfLog() { const ui = SpreadsheetApp.getUi(); if (ui.alert('Очистить журнал PDF?','Будут удалены все строки, кроме заголовка.',ui.ButtonSet.YES_NO)!==ui.Button.YES)return; const sheet=ensurePdfLogSheet_(); if(sheet.getLastRow()>1)sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getMaxColumns()).clearContent(); }
function getDetailedErrorText_(error) { return error ? (error.message || String(error)) : 'Неизвестная ошибка'; }

function clearGeneratedPdfs() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Очистить сформированные PDF?','Все PDF-файлы будут перемещены в корзину.',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;
  const folder = getDnpPdfFolder_();
  const result = trashDnpPdfFilesRecursively_(folder);
  ui.alert('Очистка завершена','Перемещено в корзину PDF-файлов: '+result.deleted+'.',ui.ButtonSet.OK);
}

function getDnpPdfFolder_() {
  const id = extractGoogleFileId_(getSettingValue_('outputFolderId'));
  if (!id) {
    throw new Error(
      'На листе «Настройки» не заполнен outputFolderId. ' +
      'Укажите ID папки «Квитанции ДНП Комфорт» или выполните «ДНП → Настройка → Первичная настройка».'
    );
  }

  try {
    const folder = dnpDriveGet_(id);
    if (folder.trashed) throw new Error('папка находится в корзине');
    if (folder.mimeType !== DNP_DRIVE_FOLDER_MIME) throw new Error('outputFolderId указывает не на папку');
    return folder;
  } catch (error) {
    throw new Error(
      'Не удалось открыть папку из Настройки → outputFolderId (' + id + '). ' +
      (error.message || error)
    );
  }
}

function trashDnpPdfFilesRecursively_(folder) {
  let deleted = 0;
  let failed = 0;
  const items = dnpDriveChildren_(folder, null);

  items.forEach(item => {
    if (item.mimeType === DNP_DRIVE_FOLDER_MIME) {
      const result = trashDnpPdfFilesRecursively_(item);
      deleted += result.deleted;
      failed += result.failed;
      return;
    }

    if (item.mimeType !== DNP_DRIVE_PDF_MIME && !/\.pdf$/i.test(item.name || '')) return;
    try {
      dnpDriveTrash_(item);
      deleted++;
    } catch (error) {
      failed++;
    }
  });

  return { deleted, failed };
}

function syncEmailSheetPlots_() {
  const ss = SpreadsheetApp.getActive();
  const emailSheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails);
  const years = getYearSheetNames_();
  if (!emailSheet || !years.length) return;
  ensureEmailSheetHeaders_(emailSheet);
  const source = ss.getSheetByName(years[years.length - 1]);
  const plots = getReceiptBlocks_(source).map(block => String(block.plot).trim());
  const existing = emailSheet.getLastRow() > 1 ? emailSheet.getRange(2,1,emailSheet.getLastRow()-1,4).getValues() : [];
  const byPlot = new Map(existing.map(row => [String(row[0]).trim(), row]));
  const rows = plots.map(plot => byPlot.get(plot) || [plot,'','',false]);
  if (emailSheet.getLastRow() > 1) emailSheet.getRange(2,1,emailSheet.getLastRow()-1,4).clearContent();
  if (rows.length) emailSheet.getRange(2,1,rows.length,4).setValues(rows);
  if (rows.length) emailSheet.getRange(2,4,rows.length,1).insertCheckboxes();
}

function ensureEmailSheetHeaders_(sheet) { sheet.getRange('A1:D1').setValues([['Участок','Email','ФИО','Отправить']]); sheet.setFrozenRows(1); if(sheet.getMaxRows()>1)sheet.getRange(2,4,sheet.getMaxRows()-1,1).insertCheckboxes(); }

function fillTestEmails() {
  const ss=SpreadsheetApp.getActive(); let sheet=ss.getSheetByName(DNP_SERVICE_SHEETS.emails); if(!sheet){ensureServiceSheets_();sheet=ss.getSheetByName(DNP_SERVICE_SHEETS.emails);} ensureEmailSheetHeaders_(sheet); if(sheet.getLastRow()<2)syncEmailSheetPlots_();
  const count=Math.max(sheet.getLastRow()-1,0); if(!count){SpreadsheetApp.getUi().alert('На листе «Почты» нет участков.');return;}
  const values=sheet.getRange(2,1,count,4).getValues(); let filled=0; values.forEach(row=>{if(row[0]&&!row[1]){row[1]=String(row[0]).replace(/[^0-9A-Za-zА-Яа-я_-]+/g,'_')+'@mail.ru';filled++;}}); sheet.getRange(2,1,count,4).setValues(values); sheet.getRange(2,4,count,1).insertCheckboxes(); ss.toast('Добавлено тестовых адресов: '+filled,'ДНП',5);
}

function sendReceipts() {
  const years = getYearSheetNames_();
  if (!years.length) throw new Error('Не найдены листы с названиями годов.');

  const now = new Date();
  const currentYear = String(now.getFullYear());
  const currentMonth = now.getMonth() + 1;
  const defaultYear = years.includes(currentYear) ? currentYear : years[years.length - 1];
  const yearOptions = years.map(year => '<option value="' + year + '"' + (year === defaultYear ? ' selected' : '') + '>' + year + '</option>').join('');
  const passwordField = isOperationPasswordRequired_() ? '<label for="password">Пароль</label><input id="password" type="password" autocomplete="off">' : '';

  const sendMode = getSendMode_();
  const testEmail = getSettingValue_('testEmail');
  let modeNote = '';
  if (sendMode === 'LIVE') {
    modeNote = '<b>LIVE:</b> письма уйдут реальным получателям, отмеченным на листе «Почты».';
  } else if (testEmail) {
    modeNote = '<b>TEST:</b> реальные получатели ничего не получат. Все выбранные квитанции будут отправлены только на <b>' + escapeHtml_(testEmail) + '</b>.';
  } else {
    modeNote = '<b>TEST:</b> testEmail не указан, поэтому письма отправляться не будут. Запуск только проверит наличие PDF и сформирует журнал.';
  }

  const html = HtmlService.createHtmlOutput(`
<!doctype html><html><head><base target="_top"><style>
body{font:14px Arial,sans-serif;padding:18px;color:#202124}h2{margin:0 0 16px;font-size:18px}label{display:block;margin:12px 0 6px;font-weight:600}select,input{width:100%;box-sizing:border-box;padding:9px;border:1px solid #dadce0;border-radius:6px}.note{background:#f8f9fa;padding:10px;border-radius:6px;margin-top:14px;line-height:1.4}.buttons{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}button{padding:9px 14px;border:0;border-radius:6px;cursor:pointer}.primary{background:#1a73e8;color:#fff}.secondary{background:#f1f3f4}#status{margin-top:12px;min-height:36px;color:#5f6368}
</style></head><body><h2>Отправка квитанций</h2><label>Год</label><select id="year">${yearOptions}</select><label>Месяц</label><select id="month"><option value="1">Январь</option><option value="2">Февраль</option><option value="3">Март</option><option value="4">Апрель</option><option value="5">Май</option><option value="6">Июнь</option><option value="7">Июль</option><option value="8">Август</option><option value="9">Сентябрь</option><option value="10">Октябрь</option><option value="11">Ноябрь</option><option value="12">Декабрь</option></select>${passwordField}<div class="note">${modeNote}</div><div id="status"></div><div class="buttons"><button class="secondary" onclick="google.script.host.close()">Отмена</button><button class="primary" id="send" onclick="run()">${sendMode === 'LIVE' ? 'Отправить' : 'Запустить TEST'}</button></div>
<script>document.getElementById('month').value='${currentMonth}';function run(){const b=document.getElementById('send'),s=document.getElementById('status'),p=document.getElementById('password');b.disabled=true;s.textContent='Идёт обработка…';google.script.run.withSuccessHandler(r=>{document.body.innerHTML='<h2>Обработка завершена</h2><div class="note">Режим: <b>'+String(r.mode||'')+'</b><br>Отправлено писем: <b>'+Number(r.sent||0)+'</b><br>TEST без отправки: <b>'+Number(r.dryRun||0)+'</b><br>Ошибок: <b>'+Number(r.failed||0)+'</b><br>Пропущено: <b>'+Number(r.skipped||0)+'</b></div><div class="buttons"><button class="primary" onclick="google.script.host.close()">Закрыть</button></div>'}).withFailureHandler(e=>{s.textContent='Ошибка: '+e.message;b.disabled=false}).sendReceiptsForMonth(Number(document.getElementById('year').value),Number(document.getElementById('month').value),p?p.value:'')}</script></body></html>`).setWidth(480).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(html, 'ДНП');
}

function getSendMode_() {
  const mode = String(getSettingValue_('sendMode') || 'TEST').trim().toUpperCase();
  if (!['TEST', 'LIVE'].includes(mode)) {
    throw new Error('Некорректный sendMode: «' + mode + '». Допустимо только TEST или LIVE.');
  }
  return mode;
}

function sendReceiptsForMonth(year, month, password) {
  requireOperationPassword_(password);
  year = Number(year);
  month = Number(month);
  if (!Number.isInteger(year) || year < 2000 || year > 2100) throw new Error('Некорректный год.');
  if (!Number.isInteger(month) || month < 1 || month > 12) throw new Error('Некорректный месяц.');

  const sendMode = getSendMode_();
  const testEmail = String(getSettingValue_('testEmail') || '').trim();
  if (sendMode === 'TEST' && testEmail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(testEmail)) {
    throw new Error('В Настройки → testEmail указан некорректный адрес: ' + testEmail);
  }

  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails);
  if (!sheet) {
    ensureServiceSheets_();
    sheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails);
  }
  ensureEmailSheetHeaders_(sheet);

  const root = getDnpPdfFolder_();
  const yearFolder = findChildFolderByNames_(root, [String(year)]);
  if (!yearFolder) throw new Error('Папка года не найдена.');

  const monthFolderName = String(month).padStart(2, '0') + ' ' + getRussianMonthName_(month);
  const monthFolder = findChildFolderByNames_(yearFolder, [monthFolderName]);
  if (!monthFolder) throw new Error('Папка месяца не найдена.');

  const count = Math.max(sheet.getLastRow() - 1, 0);
  if (!count) throw new Error('На листе «Почты» нет получателей.');

  const rows = sheet.getRange(2, 1, count, 4).getValues();
  const subjectTemplate = getMailSetting_('emailSubject') ||
    'Квитанция ДНП Комфорт: участок {{PLOT}}, {{MONTH_NAME}} {{YEAR}}';
  const bodyTemplate = getMailSetting_('emailBody') ||
    'Здравствуйте{{NAME_PART}}!\n\nНаправляем квитанцию по участку № {{PLOT}} за {{MONTH_NAME}} {{YEAR}} года.\n\nС уважением, ДНП «Комфорт».';

  let selected = 0;
  let sent = 0;
  let dryRun = 0;
  let failed = 0;
  let skipped = 0;
  const errors = [];

  rows.forEach(row => {
    const plot = String(row[0] == null ? '' : row[0]).trim();
    const email = String(row[1] == null ? '' : row[1]).trim();
    const name = String(row[2] == null ? '' : row[2]).trim();
    if (!isSendFlagEnabled_(row[3])) return;

    selected++;
    if (!plot || !email) {
      skipped++;
      appendJournalRow_('EMAIL', year, month, plot, email, 'SKIP', 'Не заполнен участок или Email');
      return;
    }

    const fileName = 'Квитанция_участок_' + sanitizePdfFileName_(plot) + '_' + year + '_' + String(month).padStart(2, '0') + '.pdf';
    const files = dnpDriveFindByName_(monthFolder, fileName, DNP_DRIVE_PDF_MIME);
    if (!files.length) {
      failed++;
      const text = 'Не найден PDF: ' + fileName;
      errors.push('Участок ' + plot + ': ' + text);
      appendJournalRow_('EMAIL', year, month, plot, email, 'ERROR', text);
      return;
    }

    try {
      const monthNumber = String(month).padStart(2, '0');
      const monthName = getRussianMonthName_(month);
      const replacements = {
        '{{PLOT}}': plot,
        '{{YEAR}}': String(year),
        '{{MONTH}}': monthNumber,
        '{{MONTH_NAME}}': monthName,
        '{{FIO}}': name,
        '{{NAME_PART}}': name ? ', ' + name : '',
        '{{P}}': plot,
        '{{Pi}}': plot,
        '{{Y}}': String(year),
        '{{M}}': monthName,
        '{{MN}}': monthName,
        '{{MM}}': monthNumber,
        '{{F}}': name
      };

      let subject = replaceMailMarkers_(subjectTemplate, replacements);
      let body = replaceMailMarkers_(bodyTemplate, replacements);

      if (sendMode === 'TEST' && !testEmail) {
        dryRun++;
        appendJournalRow_('EMAIL', year, month, plot, email, 'TEST', 'Письмо не отправлено: testEmail не указан. PDF найден: ' + fileName);
        return;
      }

      const actualRecipient = sendMode === 'TEST' ? testEmail : email;
      if (sendMode === 'TEST') {
        subject = '[TEST → ' + email + '] ' + subject;
        body = 'ТЕСТОВАЯ ОТПРАВКА. Реальный получатель: ' + email + '\n\n' + body;
      }

      const pdfBlob = dnpDriveDownloadBlob_(files[0], fileName);
      MailApp.sendEmail({
        to: actualRecipient,
        subject,
        body,
        attachments: [pdfBlob],
        name: 'ДНП Комфорт'
      });

      sent++;
      appendJournalRow_(
        'EMAIL',
        year,
        month,
        plot,
        email,
        sendMode === 'TEST' ? 'TEST_SENT' : 'SENT',
        sendMode === 'TEST' ? 'Отправлено только на ' + actualRecipient : ''
      );
      Utilities.sleep(100);
    } catch (error) {
      failed++;
      const text = error.message || String(error);
      errors.push('Участок ' + plot + ': ' + text);
      appendJournalRow_('EMAIL', year, month, plot, email, 'ERROR', text);
    }
  });

  if (!selected) throw new Error('Не выбрано ни одной строки.');

  const message =
    'Режим: ' + sendMode +
    '. Выбрано: ' + selected +
    '. Отправлено: ' + sent +
    '. TEST без отправки: ' + dryRun +
    '. Ошибок: ' + failed +
    '. Пропущено: ' + skipped + '.' +
    (errors.length ? ' Первая ошибка: ' + errors[0] : '');

  ss.toast(message, 'ДНП', 10);
  return { ok: failed === 0, mode: sendMode, selected, sent, dryRun, failed, skipped, message };
}

function isSendFlagEnabled_(value){if(value===true||value===1)return true;return['да','yes','true','1','отправить','x','+'].includes(String(value==null?'':value).trim().toLowerCase());}
function getMailSetting_(key){return getSettingValue_(key);}
function replaceMailMarkers_(text,replacements){let result=String(text==null?'':text);Object.keys(replacements).forEach(marker=>{result=result.split(marker).join(replacements[marker]);});return result;}

function clearJournal(){const ui=SpreadsheetApp.getUi();if(ui.alert('Очистить журнал?','Будут удалены все строки, кроме заголовка.',ui.ButtonSet.YES_NO)!==ui.Button.YES)return;const sheet=SpreadsheetApp.getActive().getSheetByName(DNP_SERVICE_SHEETS.journal);if(!sheet){ui.alert('Лист журнала не найден.');return;}if(sheet.getLastRow()>1)sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getMaxColumns()).clearContent();}
function appendJournalRow_(operation,year,month,plot,email,status,errorText){let sheet=SpreadsheetApp.getActive().getSheetByName(DNP_SERVICE_SHEETS.journal);if(!sheet){ensureServiceSheets_();sheet=SpreadsheetApp.getActive().getSheetByName(DNP_SERVICE_SHEETS.journal);}sheet.appendRow([new Date(),operation||'',year||'',month||'',plot||'',email||'',status||'',errorText||'']);}



function fillKwhFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const year = Number(sheet.getName());

  if (!Number.isInteger(year) || year < 2000 || year > 2100) {
    SpreadsheetApp.getUi().alert(
      'Откройте лист года, например 2026.'
    );
    return;
  }

  const LABEL_COL = 2; // B
  const JAN_COL = 3;   // C
  const DEC_COL = 14;  // N

  const lastRow = sheet.getLastRow();

  const labels = sheet
    .getRange(1, LABEL_COL, lastRow, 1)
    .getDisplayValues()
    .flat()
    .map(value =>
      String(value || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '')
    );

  const previousYear = year - 1;
  const previousSheet = ss.getSheetByName(String(previousYear));

  let blocksFound = 0;
  let januaryCreated = 0;
  let januarySkipped = 0;

  for (let i = 0; i < labels.length; i++) {

    // Ищем строку кВтч
    if (labels[i] !== 'квтч') {
      continue;
    }

    const kwhRow = i + 1;

    // Структура блока:
    //
    // Т1
    // Т2
    // Т3
    // кВтч

    const t1Row = kwhRow - 3;
    const t2Row = kwhRow - 2;
    const t3Row = kwhRow - 1;

    // Проверяем, что действительно нашли нужный блок
    if (
      t1Row < 1 ||
      labels[t1Row - 1] !== 'т1' ||
      labels[t2Row - 1] !== 'т2' ||
      labels[t3Row - 1] !== 'т3'
    ) {
      continue;
    }

    blocksFound++;

    // ==========================================
    // ФЕВРАЛЬ - ДЕКАБРЬ
    // ==========================================

    for (let col = JAN_COL + 1; col <= DEC_COL; col++) {

      const currentT1 = sheet.getRange(t1Row, col).getA1Notation();
      const currentT2 = sheet.getRange(t2Row, col).getA1Notation();
      const currentT3 = sheet.getRange(t3Row, col).getA1Notation();

      const previousT1 = sheet.getRange(t1Row, col - 1).getA1Notation();
      const previousT2 = sheet.getRange(t2Row, col - 1).getA1Notation();
      const previousT3 = sheet.getRange(t3Row, col - 1).getA1Notation();

      const formula =
        '=(' +
        currentT1 + '+' +
        currentT2 + '+' +
        currentT3 +
        ')-(' +
        previousT1 + '+' +
        previousT2 + '+' +
        previousT3 +
        ')';

      sheet
        .getRange(kwhRow, col)
        .setFormula(formula);
    }

    // ==========================================
    // ЯНВАРЬ
    // ==========================================

    if (previousSheet) {

      const currentT1 =
        sheet.getRange(t1Row, JAN_COL).getA1Notation();

      const currentT2 =
        sheet.getRange(t2Row, JAN_COL).getA1Notation();

      const currentT3 =
        sheet.getRange(t3Row, JAN_COL).getA1Notation();


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
        '=(' +
        currentT1 + '+' +
        currentT2 + '+' +
        currentT3 +
        ')-(' +
        previousT1 + '+' +
        previousT2 + '+' +
        previousT3 +
        ')';

      sheet
        .getRange(kwhRow, JAN_COL)
        .setFormula(januaryFormula);

      januaryCreated++;

    } else {

      // Если предыдущего года нет —
      // январь оставляем пустым.
      sheet
        .getRange(kwhRow, JAN_COL)
        .clearContent();

      januarySkipped++;
    }
  }

  SpreadsheetApp.flush();

  let message =
    'Обработано участков: ' + blocksFound;

  if (januaryCreated > 0) {
    message +=
      '\nЯнварь связан с декабрём ' +
      previousYear +
      ': ' +
      januaryCreated;
  }

  if (januarySkipped > 0) {
    message +=
      '\nЛист ' +
      previousYear +
      ' не найден — январь пропущен.';
  }

  ss.toast(
    message,
    'Формулы кВтч установлены',
    8
  );
}




// ============================================================================
// Дополнительные функции листа — ранее были в formula.gs и 1.gs.
// Теперь входят в единый Code.gs.
// ============================================================================

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
  if (!tariffSheet) throw new Error('Не найден лист «Тарифы».');

  const tariffLastRow = tariffSheet.getLastRow();
  if (tariffLastRow < 2) throw new Error('На листе «Тарифы» нет тарифов.');

  const tariffValues = tariffSheet.getRange(2, 1, tariffLastRow - 1, 5).getValues();
  const tariffs = [];

  tariffValues.forEach(function(row, index) {
    const tariffYear = Number(row[0]);
    const tariffMonth = Number(row[1]);
    if (!Number.isInteger(tariffYear) || !Number.isInteger(tariffMonth) || tariffMonth < 1 || tariffMonth > 12) return;
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

  tariffs.sort(function(a, b) { return a.key - b.key; });

  function getTariffForMonth_(targetYear, targetMonth) {
    const targetKey = targetYear * 100 + targetMonth;
    let result = null;
    for (let i = 0; i < tariffs.length; i++) {
      if (tariffs[i].key <= targetKey) result = tariffs[i];
      else break;
    }
    return result;
  }

  const lastRow = sheet.getLastRow();
  const labels = sheet.getRange(1, LABEL_COL, lastRow, 1).getDisplayValues().flat().map(function(value) {
    return String(value || '').trim().toLowerCase().replace(/\s+/g, '').replace(/,?₽$/g, '');
  });

  const previousYear = year - 1;
  const previousSheet = ss.getSheetByName(String(previousYear));
  let blocksProcessed = 0;
  let formulasCreated = 0;
  let januarySkipped = 0;

  for (let i = 0; i < labels.length; i++) {
    if (labels[i] !== 'т1') continue;
    if (labels[i + 1] !== 'т2' || labels[i + 2] !== 'т3') continue;

    const t1Row = i + 1;
    const t2Row = t1Row + 1;
    const t3Row = t1Row + 2;
    let sumRow = 0;

    for (let searchRow = t3Row + 1; searchRow <= Math.min(t3Row + 15, lastRow); searchRow++) {
      const label = labels[searchRow - 1] || '';
      if (label.indexOf('сумма') === 0 || label.indexOf('итого') === 0) {
        sumRow = searchRow;
        break;
      }
      if (label === 'тариф') break;
    }
    if (!sumRow) continue;

    blocksProcessed++;

    for (let month = 2; month <= 12; month++) {
      const tariff = getTariffForMonth_(year, month);
      if (!tariff) continue;

      const currentCol = month + 2;
      const previousCol = currentCol - 1;
      const currentT1 = sheet.getRange(t1Row, currentCol).getA1Notation();
      const currentT2 = sheet.getRange(t2Row, currentCol).getA1Notation();
      const currentT3 = sheet.getRange(t3Row, currentCol).getA1Notation();
      const previousT1 = sheet.getRange(t1Row, previousCol).getA1Notation();
      const previousT2 = sheet.getRange(t2Row, previousCol).getA1Notation();
      const previousT3 = sheet.getRange(t3Row, previousCol).getA1Notation();
      const currentRange = sheet.getRange(t1Row, currentCol, 3, 1).getA1Notation();

      const formula =
        '=IF(' +
          'COUNT(' + currentRange + ')=0;' +
          '"";' +
          '(' + currentT1 + '-' + previousT1 + ')*' + tariff.t1 + '+' +
          '(' + currentT2 + '-' + previousT2 + ')*' + tariff.t2 + '+' +
          '(' + currentT3 + '-' + previousT3 + ')*' + tariff.t3 +
        ')';

      sheet.getRange(sumRow, currentCol).setFormula(formula);
      formulasCreated++;
    }

    const januaryTariff = getTariffForMonth_(year, 1);
    if (januaryTariff && previousSheet) {
      const currentT1 = sheet.getRange(t1Row, JAN_COL).getA1Notation();
      const currentT2 = sheet.getRange(t2Row, JAN_COL).getA1Notation();
      const currentT3 = sheet.getRange(t3Row, JAN_COL).getA1Notation();
      const currentRange = sheet.getRange(t1Row, JAN_COL, 3, 1).getA1Notation();
      const previousT1 = "'" + previousYear + "'!" + previousSheet.getRange(t1Row, DEC_COL).getA1Notation();
      const previousT2 = "'" + previousYear + "'!" + previousSheet.getRange(t2Row, DEC_COL).getA1Notation();
      const previousT3 = "'" + previousYear + "'!" + previousSheet.getRange(t3Row, DEC_COL).getA1Notation();

      const januaryFormula =
        '=IF(' +
          'COUNT(' + currentRange + ')=0;' +
          '"";' +
          '(' + currentT1 + '-' + previousT1 + ')*' + januaryTariff.t1 + '+' +
          '(' + currentT2 + '-' + previousT2 + ')*' + januaryTariff.t2 + '+' +
          '(' + currentT3 + '-' + previousT3 + ')*' + januaryTariff.t3 +
        ')';

      sheet.getRange(sumRow, JAN_COL).setFormula(januaryFormula);
      formulasCreated++;
    } else {
      januarySkipped++;
    }
  }

  SpreadsheetApp.flush();
  let message = 'Обработано участков: ' + blocksProcessed + '\nУстановлено формул: ' + formulasCreated;
  if (januarySkipped) message += '\nЯнварь пропущен: ' + januarySkipped;
  ss.toast(message, 'Формулы суммы установлены', 10);
}

function formatYearSheetUX() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const year = Number(sheet.getName());

  if (!Number.isInteger(year) || year < 2000 || year > 2100) {
    SpreadsheetApp.getUi().alert('Откройте лист года, например 2026.');
    return;
  }

  const LABEL_COL = 2;
  const FIRST_MONTH_COL = 3;
  const MONTH_COUNT = 12;
  const lastRow = sheet.getLastRow();
  const labels = sheet.getRange(1, LABEL_COL, lastRow, 1).getDisplayValues().flat();
  let formatted = 0;

  labels.forEach(function(value, index) {
    const row = index + 1;
    const label = String(value || '').trim().toLowerCase().replace(/\s+/g, '');
    const monthsRange = sheet.getRange(row, FIRST_MONTH_COL, 1, MONTH_COUNT);
    const labelCell = sheet.getRange(row, LABEL_COL);

    if (label === 'т1' || label === 'т2' || label === 'т3') {
      monthsRange.setBackground('#ffffff').setFontWeight('normal');
      formatted++;
      return;
    }

    if (label === 'квтч') {
      monthsRange.setBackground('#eeeeee').setFontWeight('bold');
      labelCell.setBackground('#eeeeee').setFontWeight('bold');
      formatted++;
      return;
    }

    if (label === 'целевыевзносы' || label === 'целевойвзнос') {
      monthsRange.setBackground('#fff2cc').setFontWeight('normal');
      labelCell.setBackground('#fff2cc');
      formatted++;
      return;
    }

    if (label === 'сумма,₽' || label === 'сумма,р' || label === 'сумма,руб.' || label === 'сумма') {
      monthsRange.setBackground('#eeeeee').setFontWeight('bold');
      labelCell.setBackground('#eeeeee').setFontWeight('bold');
      formatted++;
    }
  });

  SpreadsheetApp.flush();
  ss.toast('Оформление применено.\nОбработано строк: ' + formatted, 'ДНП', 6);
}
