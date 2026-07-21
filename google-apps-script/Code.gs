/**
 * DNP Receipts — интерфейс v3.5.0
 *
 * Этот файл заменяет старое меню/диалоги.
 * Внешние HTML-файлы НЕ нужны: окна создаются прямо из Code.gs.
 *
 * ВАЖНО:
 * Для запуска фактической генерации PDF ниже вызывается функция:
 *   generatePdfsForMonth(year, month)
 * Если в вашем старом коде функция называется иначе — поменяйте одну строку
 * в startPdfGenerationFromDialog().
 */

const DNP_VERSION = '3.5.0';
const DNP_ADMIN_PASSWORD = '123456';

const DNP_SERVICE_SHEETS = {
  settings: 'Настройки',
  emails: 'Почты',
  journal: 'Журнал отправки',
};

// При создании нового года:
// столбец A сохраняется, значения без формул в B:последний столбец очищаются.
const DNP_YEAR_KEEP_COLUMNS = 1;

// ---------- МЕНЮ ----------

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('ДНП')
    .addSubMenu(
      ui.createMenu('Квитанции')
        .addItem('Сформировать PDF', 'showPdfDialog')
    )
    .addSubMenu(
      ui.createMenu('Почта')
        .addItem('Открыть почты', 'openEmailsSheet')
        .addItem('Заполнить тестовыми адресами', 'fillTestEmails')
        .addSeparator()
        .addItem('Отправить квитанции', 'sendReceipts')
    )
    .addSubMenu(
      ui.createMenu('Настройка')
        .addItem('Первичная настройка', 'showInitialSetupDialog')
        .addItem('Создать новый лист-год', 'showCreateYearDialog')
        .addSeparator()
        .addItem('Открыть настройки', 'openSettingsSheet')
        .addItem('Открыть журнал', 'openJournalSheet')
        .addItem('Скрыть служебные листы', 'hideServiceSheets')
        .addSeparator()
        .addItem('Очистить сформированные PDF', 'clearGeneratedPdfs')
        .addItem('Очистить журнал', 'clearJournal')
    )
    .addSeparator()
    .addItem('О программе', 'showAbout')
    .addToUi();
}

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

// ---------- СЛУЖЕБНЫЕ ЛИСТЫ ----------

function openSettingsSheet() {
  showAndActivateSheet_(DNP_SERVICE_SHEETS.settings);
}

function openEmailsSheet() {
  showAndActivateSheet_(DNP_SERVICE_SHEETS.emails);
}

function openJournalSheet() {
  showAndActivateSheet_(DNP_SERVICE_SHEETS.journal);
}

function showAndActivateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(name);

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Лист «' + name + '» не найден.');
    return;
  }

  sheet.showSheet();
  ss.setActiveSheet(sheet);
}

function hideServiceSheets() {
  const ss = SpreadsheetApp.getActive();
  let hidden = 0;

  Object.values(DNP_SERVICE_SHEETS).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && !sheet.isSheetHidden()) {
      sheet.hideSheet();
      hidden++;
    }
  });

  ss.toast('Скрыто служебных листов: ' + hidden, 'ДНП', 5);
}

// ---------- ПЕРВИЧНАЯ НАСТРОЙКА ----------

function showInitialSetupDialog() {
  const html = HtmlService.createHtmlOutput(`
<!doctype html>
<html>
<head>
  <base target="_top">
  <style>
    body { font: 14px Arial, sans-serif; padding: 18px; color: #202124; }
    h2 { margin: 0 0 16px; font-size: 18px; }
    label { display:block; margin: 12px 0 6px; font-weight: 600; }
    input, select { width:100%; box-sizing:border-box; padding:9px; border:1px solid #dadce0; border-radius:6px; }
    .buttons { display:flex; justify-content:flex-end; gap:8px; margin-top:18px; }
    button { padding:9px 14px; border:0; border-radius:6px; cursor:pointer; }
    .primary { background:#1a73e8; color:white; }
    .secondary { background:#f1f3f4; }
    #status { margin-top:12px; color:#5f6368; min-height:18px; }
  </style>
</head>
<body>
  <h2>Первичная настройка</h2>

  <label for="password">Пароль администратора</label>
  <input id="password" type="password" autocomplete="off" autofocus>

  <label for="mode">Что делать с существующей структурой</label>
  <select id="mode">
    <option value="reuse">Использовать найденные папки и шаблон</option>
    <option value="recreate">Пересоздать папки и шаблон</option>
  </select>

  <div id="status"></div>

  <div class="buttons">
    <button class="secondary" onclick="google.script.host.close()">Отмена</button>
    <button class="primary" onclick="submitSetup()">Продолжить</button>
  </div>

<script>
  document.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.preventDefault();
      submitSetup();
    }
    if (e.key === 'Escape') google.script.host.close();
  });

  function submitSetup() {
    const password = document.getElementById('password').value;
    const mode = document.getElementById('mode').value;
    const status = document.getElementById('status');

    status.textContent = 'Выполняется настройка…';

    google.script.run
      .withSuccessHandler(result => {
        status.textContent = result.message || 'Готово';
        setTimeout(() => google.script.host.close(), 1200);
      })
      .withFailureHandler(error => {
        status.textContent = 'Ошибка: ' + error.message;
      })
      .runInitialSetup(password, mode);
  }
</script>
</body>
</html>`).setWidth(440).setHeight(330);

  SpreadsheetApp.getUi().showModalDialog(html, 'ДНП');
}

function runInitialSetup(password, mode) {
  if (String(password) !== DNP_ADMIN_PASSWORD) {
    throw new Error('Неверный пароль.');
  }

  if (!['reuse', 'recreate'].includes(mode)) {
    throw new Error('Неизвестный режим настройки.');
  }

  const ss = SpreadsheetApp.getActive();
  const parent = getSpreadsheetParentFolder_();

  if (mode === 'recreate') {
    resetStoredSetupIds_();
  }

  const pdfFolder = getOrCreateChildFolder_(
    parent,
    'Квитанции ДНП Комфорт',
    mode === 'recreate'
  );

  const templateFile = getOrCreateTemplateFile_(
    parent,
    'Шаблон квитанции ДНП Комфорт',
    mode === 'recreate'
  );

  ensureServiceSheets_();
  applySevenRowBandingToYearSheets_();
  hideServiceSheets();

  const props = PropertiesService.getDocumentProperties();
  props.setProperties({
    PDF_FOLDER_ID: pdfFolder.getId(),
    TEMPLATE_DOC_ID: templateFile.getId(),
    APP_VERSION: DNP_VERSION,
  });

  ss.toast('Первичная настройка завершена', 'ДНП', 7);

  return {
    ok: true,
    message:
      'Готово. Папка: ' + pdfFolder.getName() +
      '. Шаблон: ' + templateFile.getName() + '.'
  };
}

function getSpreadsheetParentFolder_() {
  const file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parents = file.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

function getOrCreateChildFolder_(parent, name, forceCreate) {
  if (!forceCreate) {
    const found = parent.getFoldersByName(name);
    if (found.hasNext()) return found.next();
  }

  const suffix = forceCreate
    ? ' ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm-ss')
    : '';

  return parent.createFolder(name + suffix);
}

function getOrCreateTemplateFile_(parent, name, forceCreate) {
  if (!forceCreate) {
    const files = parent.getFilesByName(name);
    if (files.hasNext()) return files.next();
  }

  const docName = forceCreate
    ? name + ' ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH-mm-ss')
    : name;

  const doc = DocumentApp.create(docName);
  const body = doc.getBody();

  body.appendParagraph('ДНП «Дачный поселок «КОМФОРТ»')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph('Участок № {{PLOT}}');
  body.appendParagraph('{{MONTH_NAME}} {{YEAR}} год');
  body.appendParagraph('Сумма оплаты: {{TOTAL_TEXT}}');

  const table = body.appendTable([
    ['Наименование', 'Текущее', 'Предыдущее', 'Расход', 'Тариф', 'Сумма'],
    ['Т1', '{{T1_CURRENT}}', '{{T1_PREV}}', '{{T1_USE}}', '{{T1_RATE}}', '{{T1_SUM}}'],
    ['Т2', '{{T2_CURRENT}}', '{{T2_PREV}}', '{{T2_USE}}', '{{T2_RATE}}', '{{T2_SUM}}'],
    ['Т3', '{{T3_CURRENT}}', '{{T3_PREV}}', '{{T3_USE}}', '{{T3_RATE}}', '{{T3_SUM}}'],
    ['ИТОГО', '', '', '', '', '{{TOTAL}}'],
  ]);

  table.getRow(0).editAsText().setBold(true);
  doc.saveAndClose();

  const file = DriveApp.getFileById(doc.getId());
  file.moveTo(parent);
  return file;
}

function resetStoredSetupIds_() {
  PropertiesService.getDocumentProperties()
    .deleteProperty('PDF_FOLDER_ID')
    .deleteProperty('TEMPLATE_DOC_ID');
}

function ensureServiceSheets_() {
  const ss = SpreadsheetApp.getActive();

  if (!ss.getSheetByName(DNP_SERVICE_SHEETS.settings)) {
    const sheet = ss.insertSheet(DNP_SERVICE_SHEETS.settings);
    sheet.getRange('A1:B6').setValues([
      ['Параметр', 'Значение'],
      ['Версия', DNP_VERSION],
      ['Тариф Т1', 9.67],
      ['Тариф Т2', 3.51],
      ['Тариф Т3', 6.77],
      ['Тариф водоотведения', 53],
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
    sheet.getRange('A1:H1').setValues([[
      'Дата', 'Операция', 'Год', 'Месяц',
      'Участок', 'Email', 'Статус', 'Ошибка'
    ]]);
    sheet.setFrozenRows(1);
  }
}

// ---------- ФОРМИРОВАНИЕ PDF ----------

function showPdfDialog() {
  const years = getYearSheetNames_();
  if (!years.length) {
    SpreadsheetApp.getUi().alert(
      'Не найдено ни одного листа с названием года, например 2025, 2026 или 2027.'
    );
    return;
  }

  const options = years
    .map(y => `<option value="${y}">${y}</option>`)
    .join('');

  const currentMonth = new Date().getMonth() + 1;

  const html = HtmlService.createHtmlOutput(`
<!doctype html>
<html>
<head>
  <base target="_top">
  <style>
    body { font:14px Arial,sans-serif; padding:18px; color:#202124; }
    h2 { margin:0 0 16px; font-size:18px; }
    label { display:block; margin:12px 0 6px; font-weight:600; }
    select { width:100%; box-sizing:border-box; padding:9px; border:1px solid #dadce0; border-radius:6px; }
    .info { background:#f8f9fa; padding:10px; border-radius:6px; margin-top:14px; }
    .buttons { display:flex; justify-content:flex-end; gap:8px; margin-top:18px; }
    button { padding:9px 14px; border:0; border-radius:6px; cursor:pointer; }
    .primary { background:#1a73e8; color:#fff; }
    .secondary { background:#f1f3f4; }
    #status { margin-top:12px; min-height:18px; color:#5f6368; }
  </style>
</head>
<body>
  <h2>Формирование PDF</h2>

  <label for="year">Год</label>
  <select id="year" onchange="refreshCount()">${options}</select>

  <label for="month">Месяц</label>
  <select id="month">
    <option value="1">Январь</option>
    <option value="2">Февраль</option>
    <option value="3">Март</option>
    <option value="4">Апрель</option>
    <option value="5">Май</option>
    <option value="6">Июнь</option>
    <option value="7">Июль</option>
    <option value="8">Август</option>
    <option value="9">Сентябрь</option>
    <option value="10">Октябрь</option>
    <option value="11">Ноябрь</option>
    <option value="12">Декабрь</option>
  </select>

  <div class="info">Найдено участков: <b id="count">…</b></div>
  <div id="status"></div>

  <div class="buttons">
    <button class="secondary" onclick="google.script.host.close()">Отмена</button>
    <button class="primary" id="submit" onclick="submitForm()">Сформировать</button>
  </div>

<script>
  document.getElementById('month').value = '${currentMonth}';

  document.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.preventDefault();
      submitForm();
    }
    if (e.key === 'Escape') google.script.host.close();
  });

  function refreshCount() {
    const year = document.getElementById('year').value;
    document.getElementById('count').textContent = '…';

    google.script.run
      .withSuccessHandler(count => {
        document.getElementById('count').textContent = count;
      })
      .withFailureHandler(error => {
        document.getElementById('count').textContent = 'ошибка';
        document.getElementById('status').textContent = error.message;
      })
      .countPlotsForYear(year);
  }

  function submitForm() {
    const year = Number(document.getElementById('year').value);
    const month = Number(document.getElementById('month').value);
    const button = document.getElementById('submit');
    const status = document.getElementById('status');

    button.disabled = true;
    status.textContent = 'Формирование запущено. Прогресс показывается внизу таблицы.';

    google.script.run
      .withSuccessHandler(result => {
        status.textContent = result && result.message ? result.message : 'Готово';
        button.disabled = false;
      })
      .withFailureHandler(error => {
        status.textContent = 'Ошибка: ' + error.message;
        button.disabled = false;
      })
      .startPdfGenerationFromDialog(year, month);
  }

  refreshCount();
</script>
</body>
</html>`).setWidth(440).setHeight(410);

  SpreadsheetApp.getUi().showModalDialog(html, 'ДНП');
}

function getYearSheetNames_() {
  return SpreadsheetApp.getActive()
    .getSheets()
    .map(sheet => sheet.getName().trim())
    .filter(name => /^\d{4}$/.test(name))
    .sort((a, b) => Number(a) - Number(b));
}

function countPlotsForYear(year) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(String(year));
  if (!sheet) throw new Error('Лист «' + year + '» не найден.');

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 0;

  const values = sheet.getRange(1, 1, lastRow, 1).getDisplayValues().flat();

  return values.filter(value => {
    const text = String(value).trim();
    return text &&
      !/участок|итого|тариф|электроэнерг|водоотвед/i.test(text);
  }).length;
}

function startPdfGenerationFromDialog(year, month) {
  if (!Number.isInteger(year) || year < 2000 || year > 2100) {
    throw new Error('Некорректный год.');
  }
  if (!Number.isInteger(month) || month < 1 || month > 12) {
    throw new Error('Некорректный месяц.');
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName(String(year));
  if (!sheet) throw new Error('Лист «' + year + '» не найден.');

  SpreadsheetApp.getActive().toast(
    'Запуск формирования PDF за ' + month + '.' + year,
    'ДНП',
    5
  );

  if (typeof generatePdfsForMonth !== 'function') {
    throw new Error(
      'Не найдена функция generatePdfsForMonth(year, month). ' +
      'Переименуйте здесь вызов под функцию из вашего старого кода.'
    );
  }

  const result = generatePdfsForMonth(year, month);

  return {
    ok: true,
    message: result && result.message
      ? result.message
      : 'Формирование завершено.'
  };
}

// ---------- СОЗДАНИЕ НОВОГО ГОДА ----------

function showCreateYearDialog() {
  const years = getYearSheetNames_();
  if (!years.length) {
    SpreadsheetApp.getUi().alert(
      'Сначала создайте хотя бы один лист с названием года, например 2026.'
    );
    return;
  }

  const latestYear = Math.max(...years.map(Number));
  const suggestedYear = latestYear + 1;
  const sourceOptions = years
    .slice()
    .reverse()
    .map(y => `<option value="${y}">${y}</option>`)
    .join('');

  const html = HtmlService.createHtmlOutput(`
<!doctype html>
<html>
<head>
  <base target="_top">
  <style>
    body { font:14px Arial,sans-serif; padding:18px; color:#202124; }
    h2 { margin:0 0 16px; font-size:18px; }
    label { display:block; margin:12px 0 6px; font-weight:600; }
    input,select { width:100%; box-sizing:border-box; padding:9px; border:1px solid #dadce0; border-radius:6px; }
    .check { display:flex; align-items:center; gap:8px; margin-top:14px; }
    .check input { width:auto; }
    .buttons { display:flex; justify-content:flex-end; gap:8px; margin-top:18px; }
    button { padding:9px 14px; border:0; border-radius:6px; cursor:pointer; }
    .primary { background:#1a73e8; color:#fff; }
    .secondary { background:#f1f3f4; }
    #status { margin-top:12px; min-height:18px; color:#5f6368; }
  </style>
</head>
<body>
  <h2>Создать новый лист-год</h2>

  <label for="newYear">Новый год</label>
  <input id="newYear" type="number" min="2000" max="2100" value="${suggestedYear}" autofocus>

  <label for="sourceYear">Копировать из листа</label>
  <select id="sourceYear">${sourceOptions}</select>

  <label class="check">
    <input id="clearValues" type="checkbox" checked>
    Очистить данные, сохранив столбец A, формулы и оформление
  </label>

  <div id="status"></div>

  <div class="buttons">
    <button class="secondary" onclick="google.script.host.close()">Отмена</button>
    <button class="primary" onclick="createYear()">Создать</button>
  </div>

<script>
  document.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      e.preventDefault();
      createYear();
    }
    if (e.key === 'Escape') google.script.host.close();
  });

  function createYear() {
    const newYear = Number(document.getElementById('newYear').value);
    const sourceYear = document.getElementById('sourceYear').value;
    const clearValues = document.getElementById('clearValues').checked;
    const status = document.getElementById('status');

    status.textContent = 'Создаётся лист…';

    google.script.run
      .withSuccessHandler(result => {
        status.textContent = result.message;
        setTimeout(() => google.script.host.close(), 1200);
      })
      .withFailureHandler(error => {
        status.textContent = 'Ошибка: ' + error.message;
      })
      .createNewYearSheet(newYear, sourceYear, clearValues);
  }
</script>
</body>
</html>`).setWidth(450).setHeight(390);

  SpreadsheetApp.getUi().showModalDialog(html, 'ДНП');
}

function createNewYearSheet(newYear, sourceYear, clearValues) {
  newYear = Number(newYear);
  sourceYear = String(sourceYear);

  if (!Number.isInteger(newYear) || newYear < 2000 || newYear > 2100) {
    throw new Error('Некорректный новый год.');
  }

  const ss = SpreadsheetApp.getActive();
  const newName = String(newYear);

  if (ss.getSheetByName(newName)) {
    throw new Error('Лист «' + newName + '» уже существует.');
  }

  const source = ss.getSheetByName(sourceYear);
  if (!source) {
    throw new Error('Исходный лист «' + sourceYear + '» не найден.');
  }

  const target = source.copyTo(ss).setName(newName);
  ss.setActiveSheet(target);

  if (clearValues) {
    clearCopiedYearValues_(target);
  }

  applySevenRowBanding_(target);

  ss.toast(
    'Создан лист ' + newName + ' на основе ' + sourceYear,
    'ДНП',
    7
  );

  return {
    ok: true,
    message: 'Лист «' + newName + '» создан.'
  };
}

function clearCopiedYearValues_(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  const formulas = range.getFormulas();

  if (!values.length || !values[0].length) return;

  for (let r = 0; r < values.length; r++) {
    for (let c = DNP_YEAR_KEEP_COLUMNS; c < values[r].length; c++) {
      if (!formulas[r][c]) {
        values[r][c] = '';
      }
    }
  }

  range.setValues(values);
}

// ---------- АВТОРАЗМЕТКА ПО 7 СТРОК ----------

function applySevenRowBandingToYearSheets_() {
  getYearSheetNames_().forEach(year => {
    const sheet = SpreadsheetApp.getActive().getSheetByName(year);
    if (sheet) applySevenRowBanding_(sheet);
  });
}

function applySevenRowBanding_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  if (lastRow < 1) return;

  const startRow = 1;
  const colors = ['#ffffff', '#f3f7fd'];

  for (let row = startRow; row <= lastRow; row += 7) {
    const count = Math.min(7, lastRow - row + 1);
    const bandIndex = Math.floor((row - startRow) / 7) % 2;
    sheet.getRange(row, 1, count, lastColumn)
      .setBackground(colors[bandIndex]);
  }
}

// ---------- ПОЧТА / ЖУРНАЛ / ОЧИСТКА ----------

function syncEmailSheetPlots_() {
  const ss = SpreadsheetApp.getActive();
  const emailSheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails);
  const years = getYearSheetNames_();
  if (!emailSheet || !years.length) return;

  const source = ss.getSheetByName(years[years.length - 1]);
  const plots = source
    .getRange(1, 1, source.getLastRow(), 1)
    .getDisplayValues()
    .flat()
    .map(v => String(v).trim())
    .filter(v => v && !/участок|итого|тариф|электроэнерг|водоотвед/i.test(v));

  const existing = emailSheet.getLastRow() > 1
    ? emailSheet.getRange(2, 1, emailSheet.getLastRow() - 1, 4).getValues()
    : [];

  const existingByPlot = new Map(
    existing.map(row => [String(row[0]).trim(), row])
  );

  const rows = plots.map(plot => {
    const old = existingByPlot.get(plot);
    return old || [plot, '', '', ''];
  });

  if (emailSheet.getLastRow() > 1) {
    emailSheet.getRange(2, 1, emailSheet.getLastRow() - 1, 4).clearContent();
  }

  if (rows.length) {
    emailSheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }
}

function fillTestEmails() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails);

  if (!sheet) {
    ensureServiceSheets_();
    sheet = ss.getSheetByName(DNP_SERVICE_SHEETS.emails);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    syncEmailSheetPlots_();
  }

  const count = Math.max(sheet.getLastRow() - 1, 0);
  if (!count) {
    SpreadsheetApp.getUi().alert('На листе «Почты» нет участков.');
    return;
  }

  const values = sheet.getRange(2, 1, count, 2).getValues();
  let filled = 0;

  values.forEach(row => {
    if (row[0] && !row[1]) {
      const safe = String(row[0]).replace(/[^0-9A-Za-zА-Яа-я_-]+/g, '_');
      row[1] = safe + '@mail.ru';
      filled++;
    }
  });

  sheet.getRange(2, 1, count, 2).setValues(values);
  ss.toast('Добавлено тестовых адресов: ' + filled, 'ДНП', 5);
}

function clearJournal() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Очистить журнал?',
    'Будут удалены все строки журнала, кроме заголовка.',
    ui.ButtonSet.YES_NO
  );

  if (answer !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.getActive()
    .getSheetByName(DNP_SERVICE_SHEETS.journal);

  if (!sheet) {
    ui.alert('Лист журнала не найден.');
    return;
  }

  if (sheet.getLastRow() > 1) {
    sheet.getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      sheet.getMaxColumns()
    ).clearContent();
  }

  SpreadsheetApp.getActive().toast('Журнал очищен', 'ДНП', 5);
}

function sendReceipts() {
  SpreadsheetApp.getUi().alert(
    'Подключите здесь вашу существующую функцию последовательной отправки.'
  );
}

function clearGeneratedPdfs() {
  SpreadsheetApp.getUi().alert(
    'Подключите здесь вашу существующую функцию очистки сформированных PDF.'
  );
}
