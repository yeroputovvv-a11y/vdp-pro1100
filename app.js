const pdfInput = document.getElementById('pdfFile');
const xlsxInput = document.getElementById('xlsxFile');
const prepare = document.getElementById('prepare');
const status = document.getElementById('status');

let pdf = null;
let parsed = null;

function rawText(value) {
  if (value === null || value === undefined) return '';
  return String(value);
}

function hasAnyContent(value) {
  return rawText(value).length > 0; // even one space counts for the print run
}

function hasMeaningfulValue(value) {
  return rawText(value).trim().length > 0;
}

function rowUsed(row, fieldCount) {
  for (let c = 0; c < fieldCount; c += 1) {
    if (hasAnyContent(row[c])) return true;
  }
  return false;
}

function rowHasVariableData(row, fieldCount) {
  for (let c = 0; c < fieldCount; c += 1) {
    if (hasMeaningfulValue(row[c])) return true;
  }
  return false;
}

function escapeHtml(value) {
  return rawText(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function refreshButton() {
  prepare.disabled = !(pdf && parsed && parsed.ok);
}

function resetAnalysis() {
  document.getElementById('summary').hidden = true;
  document.getElementById('fields').hidden = true;
  document.getElementById('preview').hidden = true;
  document.getElementById('fieldList').innerHTML = '';
  document.getElementById('previewTable').innerHTML = '';
}

function renderAnalysis(parsedData) {
  const { headers, reference, records, usedRows, variableRows, reserveRows, variableFields } = parsedData;

  document.getElementById('totalPrint').textContent = usedRows.length.toLocaleString('ru-RU');
  document.getElementById('variablePrint').textContent = variableRows.length.toLocaleString('ru-RU');
  document.getElementById('reservePrint').textContent = reserveRows.length.toLocaleString('ru-RU');
  document.getElementById('fieldCount').textContent = variableFields.length.toLocaleString('ru-RU');
  document.getElementById('summary').hidden = false;

  const fieldRows = variableFields.map((field) => `
    <div class="field-row">
      <div class="field-name">${escapeHtml(field.name)}</div>
      <div class="field-reference">${field.reference === '' ? '<em>пустой эталон</em>' : escapeHtml(field.reference)}</div>
      <div class="field-meta">${field.usedInData} строк${field.usedInData === 1 ? 'а' : field.usedInData < 5 ? 'и' : ''} с данными</div>
    </div>
  `).join('');
  document.getElementById('fieldList').innerHTML = fieldRows || '<div class="empty">Переменные колонки не найдены.</div>';
  document.getElementById('fields').hidden = false;

  const sample = records.slice(0, 8);
  const tableHead = ['№', ...headers.slice(1)].map(h => `<th>${escapeHtml(h || '')}</th>`).join('');
  const tableBody = sample.map((record) => {
    const cells = record.displayValues.map(v => `<td>${escapeHtml(v)}</td>`).join('');
    return `<tr><td class="row-no">${record.index}</td>${cells}</tr>`;
  }).join('');
  document.getElementById('previewTable').innerHTML = `
    <div class="table-wrap"><table><thead><tr>${tableHead}</tr></thead><tbody>${tableBody}</tbody></table></div>
  `;
  document.getElementById('preview').hidden = false;
}

pdfInput.addEventListener('change', (event) => {
  pdf = event.target.files[0] || null;
  document.getElementById('pdfName').textContent = pdf ? pdf.name : 'Файл не выбран';
  status.textContent = pdf ? 'PDF выбран. Загрузите Excel с данными.' : 'Выберите PDF и Excel.';
  refreshButton();
});

xlsxInput.addEventListener('change', async (event) => {
  const file = event.target.files[0] || null;
  document.getElementById('xlsxName').textContent = file ? file.name : 'Файл не выбран';
  parsed = null;
  resetAnalysis();
  if (!file) {
    status.textContent = 'Выберите PDF и Excel.';
    refreshButton();
    return;
  }

  try {
    const data = await file.arrayBuffer();
    const book = XLSX.read(data, { type: 'array', cellDates: false, raw: true });
    if (!book.SheetNames.length) throw new Error('В книге нет листов.');

    const sheet = book.Sheets[book.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
    if (matrix.length < 2) throw new Error('Нужны как минимум две строки: заголовки и эталонные значения.');

    const width = Math.max(
      matrix[0]?.length || 0,
      matrix[1]?.length || 0,
      ...matrix.slice(2).map(row => row.length)
    );
    if (width < 2) throw new Error('Нужно минимум одно переменное поле кроме колонки №.');

    const headers = Array.from({ length: width }, (_, i) => rawText(matrix[0]?.[i] ?? ''));
    const reference = Array.from({ length: width }, (_, i) => rawText(matrix[1]?.[i] ?? ''));

    if (!headers.slice(1).some(Boolean)) throw new Error('В строке 1 нет названий переменных колонок.');

    const records = [];
    const usedRows = [];
    const variableRows = [];
    const reserveRows = [];

    for (let sourceRow = 2; sourceRow < matrix.length; sourceRow += 1) {
      const row = Array.from({ length: width }, (_, i) => rawText(matrix[sourceRow]?.[i] ?? ''));
      if (!rowUsed(row, width - 0)) continue;

      const variable = rowHasVariableData(row.slice(1), width - 1);
      const record = {
        index: usedRows.length + 1,
        sourceRow: sourceRow + 1,
        variable,
        values: row,
        displayValues: row.slice(1),
      };
      records.push(record);
      usedRows.push(record);
      if (variable) variableRows.push(record);
      else reserveRows.push(record);
    }

    const variableFields = [];
    for (let c = 1; c < width; c += 1) {
      const name = headers[c];
      if (!name) continue;
      const usedInData = records.filter(r => hasMeaningfulValue(r.values[c])).length;
      variableFields.push({
        index: c,
        name,
        reference: reference[c],
        usedInData,
      });
    }

    parsed = {
      ok: true,
      sheetName: book.SheetNames[0],
      headers,
      reference,
      records,
      usedRows,
      variableRows,
      reserveRows,
      variableFields,
    };

    renderAnalysis(parsed);
    status.textContent = `Excel проверен: ${usedRows.length} пронумерованных экземпляров, из них ${variableRows.length} с переменными данными и ${reserveRows.length} резервных.`;
  } catch (error) {
    status.textContent = `Ошибка чтения Excel: ${error.message}`;
  }

  refreshButton();
});

prepare.addEventListener('click', () => {
  if (!parsed || !pdf) return;
  status.textContent = 'Структура данных принята. Следующий этап — полный анализ PDF и автоматический поиск всех совпадений с эталонами.';
});

refreshButton();
