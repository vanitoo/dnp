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
