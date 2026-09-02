const pdfInput = document.getElementById('pdfFile');
const xlsxInput = document.getElementById('xlsxFile');
const prepare = document.getElementById('prepare');
const status = document.getElementById('status');
const templateBtn = document.getElementById('downloadTemplate');

let pdf = null;
let parsed = null;

function rawText(value) {
  if (value === null || value === undefined) return '';
  return String(value);
}

function hasAnyContent(value) {
  // One or more spaces count as an intentionally numbered/reserved row.
  return rawText(value).length > 0;
}

function hasMeaningfulValue(value) {
  return rawText(value).trim().length > 0;
}

function rowUsed(row) {
  return row.some(hasAnyContent);
}

function rowHasVariableData(row) {
  return row.some(hasMeaningfulValue);
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
  const { headers, records, usedRows, variableRows, reserveRows, variableFields } = parsedData;

  document.getElementById('totalPrint').textContent = usedRows.length.toLocaleString('ru-RU');
  document.getElementById('variablePrint').textContent = variableRows.length.toLocaleString('ru-RU');
  document.getElementById('reservePrint').textContent = reserveRows.length.toLocaleString('ru-RU');
  document.getElementById('fieldCount').textContent = variableFields.length.toLocaleString('ru-RU');
  document.getElementById('summary').hidden = false;

  document.getElementById('fieldList').innerHTML = variableFields.map((field) => `
    <div class="field-row">
      <div class="field-name">${escapeHtml(field.name)}</div>
      <div class="field-reference">${field.reference === '' ? '<em>пустой эталон</em>' : escapeHtml(field.reference)}</div>
      <div class="field-meta">${field.usedInData} экземпляров с заполнением</div>
    </div>
  `).join('') || '<div class="empty">Переменные колонки не найдены.</div>';
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

async function downloadExcelTemplate() {
  if (!window.ExcelJS) {
    status.textContent = 'Не удалось загрузить модуль создания Excel. Обновите страницу и повторите попытку.';
    return;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'VDP PRO 1100';
  workbook.created = new Date();
  const sheet = workbook.addWorksheet('Данные', { views: [{ state: 'frozen', ySplit: 2 }] });

  const headers = ['№', 'Номер изделия', 'Фамилия', 'Имя', 'Отчество', 'Звание'];
  const reference = ['ЭТАЛОН', '529260704929', 'Иванов', 'Иван', 'Иванович', 'майор'];
  sheet.addRow(headers);
  sheet.addRow(reference);

  for (let r = 3; r <= 1002; r += 1) {
    sheet.getCell(r, 1).value = { formula: `IF(COUNTA(B${r}:F${r})=0,"",MAX($A$2:A${r-1})+1)` };
  }

  // Clear instructional sample data from the editable rows: only the reference row is prefilled.
  for (let r = 3; r <= 1002; r += 1) {
    for (let c = 2; c <= 6; c += 1) sheet.getCell(r, c).value = '';
  }

  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };
  sheet.getRow(1).alignment = { horizontal: 'center', vertical: 'middle' };

  sheet.getRow(2).font = { bold: true, italic: true, color: { argb: 'FF7C3AED' } };
  sheet.getRow(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
  sheet.getRow(2).alignment = { horizontal: 'center', vertical: 'middle' };

  sheet.getColumn(1).width = 8;
  sheet.getColumn(2).width = 20;
  sheet.getColumn(3).width = 20;
  sheet.getColumn(4).width = 18;
  sheet.getColumn(5).width = 22;
  sheet.getColumn(6).width = 18;

  sheet.getRow(2).height = 24;
  for (let r = 3; r <= 1002; r += 1) sheet.getRow(r).height = 20;
  for (let r = 1; r <= 1002; r += 1) {
    sheet.getCell(r, 1).alignment = { horizontal: 'center', vertical: 'middle' };
  }
  sheet.getColumn(1).numFmt = '0';

  const info = workbook.addWorksheet('Инструкция');
  info.addRow(['ШАБЛОН ДАННЫХ VDP PRO 1100']);
  info.mergeCells('A1:C1');
  info.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 14 };
  info.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };
  info.getRow(1).alignment = { horizontal: 'center' };
  [
    ['Строка 1', 'Названия колонок.'],
    ['Строка 2', 'Эталонные значения. Они используются для поиска переменных полей в PDF.'],
    ['Строки 3+', 'Одна строка = один печатаемый экземпляр.'],
    ['Пробел', 'Один или несколько пробелов в ячейке тоже делают строку частью общего тиража; такой экземпляр может быть резервным.'],
    ['Пустая ячейка', 'Конкретное поле в этом экземпляре не заменяется.'],
    ['Сводка', 'Веб-программа считает общий тираж, экземпляры с переменными данными и резерв.'],
  ].forEach(row => info.addRow(row));
  info.getColumn(1).width = 20;
  info.getColumn(2).width = 85;
  info.getColumn(3).width = 4;
  info.getColumn(2).alignment = { wrapText: true, vertical: 'top' };
  info.getColumn(1).font = { bold: true };

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'VDP_PRO1100_template.xlsx';
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  status.textContent = 'Шаблон Excel скачан. Заполните строку «ЭТАЛОН» и данные ниже неё, затем загрузите файл обратно.';
}

templateBtn.addEventListener('click', downloadExcelTemplate);

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
      if (!rowUsed(row.slice(1))) continue; // column № itself is generated by the program

      const variable = rowHasVariableData(row.slice(1));
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
      variableFields.push({ index: c, name, reference: reference[c], usedInData });
    }

    parsed = { ok: true, headers, reference, records, usedRows, variableRows, reserveRows, variableFields };
    renderAnalysis(parsed);
    status.textContent = `Excel проверен: ${usedRows.length} экземпляров в тираже, ${variableRows.length} с переменными данными, ${reserveRows.length} резервных.`;
  } catch (error) {
    status.textContent = `Ошибка чтения Excel: ${error.message}`;
  }

  refreshButton();
});

prepare.addEventListener('click', () => {
  if (!parsed || !pdf) return;
  status.textContent = 'Данные приняты. Следующий этап — полный анализ PDF, поиск всех эталонных совпадений и проверка количества мест замены.';
});

refreshButton();
