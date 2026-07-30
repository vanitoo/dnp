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
  const dataRowPattern = /^(т1|т2|т3|квтч|квт·ч|квт\/ч|водоотведение|целевые взносы|сумма|сумма,?\s*₽|сумма,?\s*р)$/i;

  for (let row = 2; row <= lastRow; row++) {
    if (!dataRowPattern.test(labels[row - 1])) continue;
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
      if (rowCount > 0) sheet.getRange(startRow, 1, rowCount, lastColumn).setBackground(colors[blockIndex % colors.length]);
    });
    return;
  }

  for (let row = 2, blockIndex = 0; row <= lastRow; row += 7, blockIndex++) {
    const rowCount = Math.min(7, lastRow - row + 1);
    sheet.getRange(row, 1, rowCount, lastColumn).setBackground(colors[blockIndex % colors.length]);
  }
}
