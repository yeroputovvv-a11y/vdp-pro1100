const pdfInput = document.getElementById('pdfFile');
const xlsxInput = document.getElementById('xlsxFile');
const prepare = document.getElementById('prepare');
const status = document.getElementById('status');
const templateBtn = document.getElementById('downloadTemplate');
const fieldCountInput = document.getElementById('fieldCountInput');

let pdf = null;
let parsed = null;

pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

const S = v => v == null ? '' : String(v);
const T = v => S(v).trim();
const E = v => S(v)
  .replaceAll('&', '&amp;')
  .replaceAll('<', '&lt;')
  .replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;')
  .replaceAll("'", '&#039;');

function reset() {
  for (const id of ['summary', 'fields', 'preview']) document.getElementById(id).hidden = true;
  document.getElementById('fieldList').innerHTML = '';
  document.getElementById('previewTable').innerHTML = '';
}

function refresh() {
  prepare.disabled = !(pdf && parsed?.ok);
}

function parseExcel(book) {
  const sh = book.Sheets[book.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(sh, { header: 1, defval: '', raw: true });
  if (matrix.length < 2) throw Error('Нужны строки 1 и 2.');

  const headers = (matrix[0] || []).map(S);
  const refs = (matrix[1] || []).map(S);
  const numberCol = headers.findIndex(x => ['№', 'номер', 'номер строки'].includes(T(x).toLowerCase()));
  const nc = numberCol >= 0 ? numberCol : 0;
  const width = Math.max(headers.length, refs.length, ...matrix.slice(2).map(x => x.length));

  const fields = [];
  for (let c = nc + 1; c < width; c++) {
    if (T(headers[c])) fields.push({ col: c, name: headers[c], ref: refs[c] || '', used: 0 });
  }
  if (!fields.length) throw Error('После колонки «№» нет названных переменных.');

  const used = [];
  const variable = [];
  const reserve = [];

  for (let rr = 2; rr < matrix.length; rr++) {
    const row = Array.from({ length: width }, (_, i) => S(matrix[rr]?.[i]));
    const vals = fields.map(f => row[f.col]);
    if (!vals.some(v => S(v).length)) continue;

    const hasMeaningful = vals.some(v => T(v).length);
    const rec = { no: used.length + 1, values: vals };
    used.push(rec);
    (hasMeaningful ? variable : reserve).push(rec);
    fields.forEach((f, i) => { if (T(vals[i])) f.used += 1; });
  }

  return { ok: true, headers, refs, fields, used, variable, reserve };
}

function renderExcel(p) {
  document.getElementById('totalPrint').textContent = p.used.length.toLocaleString('ru-RU');
  document.getElementById('variablePrint').textContent = p.variable.length.toLocaleString('ru-RU');
  document.getElementById('reservePrint').textContent = p.reserve.length.toLocaleString('ru-RU');
  document.getElementById('fieldCount').textContent = p.fields.length.toLocaleString('ru-RU');
  document.getElementById('summary').hidden = false;

  const sample = p.used.slice(0, 8);
  const tableHead = `<th>№</th>${p.fields.map(f => `<th>${E(f.name)}</th>`).join('')}`;
  const tableBody = sample.map(x => `<tr><td class="row-no">${x.no}</td>${x.values.map(v => `<td>${E(v)}</td>`).join('')}</tr>`).join('');
  document.getElementById('previewTable').innerHTML = `<div class="table-wrap"><table><thead><tr>${tableHead}</tr></thead><tbody>${tableBody}</tbody></table></div>`;
  document.getElementById('preview').hidden = false;
}

function makeLines(items) {
  const groups = [];
  for (const i of items.filter(x => x.str)) {
    const z = {
      t: i.str,
      x: +(i.transform?.[4] || 0),
      y: +(i.transform?.[5] || 0),
      w: +(i.width || 0),
      h: +(i.height || Math.abs(i.transform?.[3] || 0))
    };
    let g = groups.find(a => Math.abs(a.y - z.y) <= 2.5);
    if (!g) { g = { y: z.y, a: [] }; groups.push(g); }
    g.a.push(z);
  }

  return groups.sort((a, b) => b.y - a.y).map(g => {
    g.a.sort((a, b) => a.x - b.x);
    let s = '';
    for (let i = 0; i < g.a.length; i++) {
      if (i) {
        const prev = g.a[i - 1];
        const cur = g.a[i];
        const gap = cur.x - (prev.x + prev.w);
        if (gap > Math.max(1, cur.h * 0.15) && !(/[\s(\[«]$/.test(s)) && !(/^[,.;:!?%)\]}]/.test(cur.t))) s += ' ';
      }
      s += g.a[i].t;
    }
    return { t: T(s) };
  });
}

function occurrences(text, needle) {
  let count = 0;
  let pos = 0;
  while (true) {
    const found = text.indexOf(needle, pos);
    if (found < 0) break;
    count += 1;
    pos = found + Math.max(1, needle.length);
  }
  return count;
}

async function analyze(file, fields) {
  const doc = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const pages = [];

  for (let p = 1; p <= doc.numPages; p++) {
    const content = await (await doc.getPage(p)).getTextContent();
    pages.push({ p, lines: makeLines(content.items) });
  }

  const duplicateRefs = new Map();
  fields.forEach(f => {
    const ref = T(f.ref);
    if (!ref) return;
    const arr = duplicateRefs.get(ref) || [];
    arr.push(f.name);
    duplicateRefs.set(ref, arr);
  });

  return fields.map(f => {
    const ref = T(f.ref);
    if (!ref) return { ...f, count: 0, loc: [], state: 'empty' };

    let count = 0;
    const loc = [];
    for (const page of pages) {
      for (const line of page.lines) {
        const c = occurrences(line.t, ref);
        if (c) {
          count += c;
          loc.push({ p: page.p, c });
        }
      }
    }

    const duplicate = (duplicateRefs.get(ref) || []).length > 1;
    return { ...f, count, loc, state: duplicate ? 'duplicate' : count ? 'found' : 'not-found' };
  });
}

async function run() {
  prepare.disabled = true;
  status.textContent = 'Анализирую весь PDF…';

  try {
    const analysis = await analyze(pdf, parsed.fields);
    document.getElementById('fieldList').innerHTML = analysis.map(f => {
      const icon = f.state === 'found' ? '✓' : f.state === 'not-found' ? '⚠' : f.state === 'duplicate' ? '⛔' : '—';
      const msg = f.state === 'found'
        ? `Найдено совпадений: ${f.count}`
        : f.state === 'not-found'
          ? 'Совпадений не найдено'
          : f.state === 'duplicate'
            ? 'Эталон совпадает с другим полем'
            : 'Эталон не задан';
      const loc = f.loc.map(x => `<span class="location">стр. ${x.p}${x.c > 1 ? ` ×${x.c}` : ''}</span>`).join('');
      return `<div class="field-row ${f.state === 'not-found' ? 'field-warning' : ''} ${f.state === 'duplicate' ? 'field-error' : ''}">
        <div class="field-check">${icon}</div>
        <div class="field-name">${E(f.name)}</div>
        <div class="field-reference">Эталон: ${E(f.ref || '—')}</div>
        <div class="field-meta"><b>${msg}</b><div class="locations">${loc}</div></div>
      </div>`;
    }).join('');

    document.getElementById('fields').hidden = false;
    const total = analysis.reduce((sum, f) => sum + f.count, 0);
    const problems = analysis.filter(f => ['not-found', 'duplicate'].includes(f.state)).length;
    status.textContent = problems
      ? `PDF проверен: ${total} совпадений. Требуют внимания: ${problems}.`
      : `PDF проверен: ${total} совпадений. Все эталонные значения найдены.`;
  } catch (error) {
    status.textContent = `Ошибка анализа PDF: ${error.message}`;
  } finally {
    refresh();
  }
}

function columnLetter(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function styleTemplateSheet(sheet, headerEndCol, refEndCol) {
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };
  const refFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };

  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = headerFill;
  sheet.getRow(1).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  sheet.getRow(1).height = 24;

  sheet.getRow(2).font = { bold: true, italic: true, color: { argb: 'FF7C3AED' } };
  sheet.getRow(2).fill = refFill;
  sheet.getRow(2).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  sheet.getRow(2).height = 24;

  for (let c = 1; c <= headerEndCol; c++) {
    const cell = sheet.getCell(1, c);
    cell.border = { bottom: { style: 'thin', color: { argb: 'FFD0D5DD' } } };
    sheet.getCell(2, c).border = { bottom: { style: 'thin', color: { argb: 'FFD0D5DD' } } };
    sheet.getColumn(c).width = c === 1 ? 4 : c === 2 ? 24 : 20;
  }

  for (let c = 3; c <= headerEndCol; c++) sheet.getColumn(c).width = c === 3 ? 4 : 20;
  sheet.getColumn(1).width = 24;
  sheet.getColumn(2).width = 18;

  if (refEndCol >= 4) sheet.getColumn(4).width = 8;
  sheet.freezePanes.freezeRows(2);
}

async function downloadExcelTemplate(event) {
  event?.preventDefault();

  if (!window.ExcelJS) {
    status.textContent = 'Модуль ExcelJS ещё не загрузился. Обновите страницу и повторите попытку.';
    return;
  }

  const requested = Math.min(100, Math.max(1, Number.parseInt(fieldCountInput?.value || '5', 10) || 5));
  if (fieldCountInput) fieldCountInput.value = String(requested);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'VDP PRO 1100';
  wb.created = new Date();

  const sheet = wb.addWorksheet('Данные');
  // A:B — статистика, C — разделитель, D — №, E+ — переменные поля.
  const numberCol = 4;
  const firstVarCol = 5;
  const lastVarCol = firstVarCol + requested - 1;
  const helperCol = lastVarCol + 1;
  const helperLetter = columnLetter(helperCol);
  const firstVarLetter = columnLetter(firstVarCol);
  const lastVarLetter = columnLetter(lastVarCol);

  sheet.getCell('A1').value = 'Показатель';
  sheet.getCell('B1').value = 'Значение';
  sheet.getCell('A2').value = 'Общий тираж';
  sheet.getCell('A3').value = 'С переменными данными';
  sheet.getCell('A4').value = 'Резерв без данных';

  sheet.getCell(numberCol, 1).value = '';
  sheet.getCell(numberCol, 2).value = '';
  sheet.getCell(1, numberCol).value = '№';

  const defaultNames = ['Номер изделия', 'Фамилия', 'Имя', 'Отчество', 'Звание'];
  const defaultRefs = ['529260704929', 'Иванов', 'Иван', 'Иванович', 'майор'];
  for (let i = 0; i < requested; i++) {
    const col = firstVarCol + i;
    sheet.getCell(1, col).value = defaultNames[i] || `Поле ${i + 1}`;
    sheet.getCell(2, col).value = defaultRefs[i] || '';
  }
  sheet.getCell(2, numberCol).value = 'ЭТАЛОН';

  for (let r = 3; r <= 1002; r++) {
    sheet.getCell(r, numberCol).value = {
      formula: `IF(COUNTA(${firstVarLetter}${r}:${lastVarLetter}${r})>0,ROW()-2,"")`
    };
    sheet.getCell(r, helperCol).value = {
      formula: `IF(SUMPRODUCT(--(LEN(TRIM(${firstVarLetter}${r}:${lastVarLetter}${r}&""))>0))>0,1,0)`
    };
    for (let c = firstVarCol; c <= lastVarCol; c++) sheet.getCell(r, c).value = '';
  }

  sheet.getCell('B2').value = { formula: '=COUNT(D3:D1002)' };
  sheet.getCell('B3').value = { formula: `=SUM(${helperLetter}3:${helperLetter}1002)` };
  sheet.getCell('B4').value = { formula: '=B2-B3' };

  styleTemplateSheet(sheet, lastVarCol, lastVarCol);
  sheet.getColumn(helperCol).hidden = true;
  sheet.getColumn(helperCol).width = 3;
  sheet.getColumn(3).width = 3;
  sheet.getColumn(numberCol).width = 8;
  sheet.getColumn(numberCol).alignment = { horizontal: 'center', vertical: 'middle' };

  for (const a of ['A1', 'B1']) {
    sheet.getCell(a).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getCell(a).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };
    sheet.getCell(a).alignment = { horizontal: 'center', vertical: 'middle' };
  }
  for (const a of ['A2', 'A3', 'A4']) sheet.getCell(a).font = { bold: true };
  for (const a of ['B2', 'B3', 'B4']) sheet.getCell(a).font = { bold: true, size: 12 };
  for (let r = 2; r <= 4; r++) {
    sheet.getCell(`A${r}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
    sheet.getCell(`B${r}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
  }

  const info = wb.addWorksheet('Инструкция');
  info.getCell('A1').value = 'ШАБЛОН VDP PRO 1100';
  info.mergeCells('A1:B1');
  info.getCell('A1').font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 14 };
  info.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2937' } };
  info.getCell('A1').alignment = { horizontal: 'center' };
  const instructions = [
    ['Строка 1', 'Названия переменных колонок. Колонка «№» обязательна.'],
    ['Строка 2', 'ЭТАЛОННЫЕ значения. По ним программа ищет совпадения во всём PDF.'],
    ['Строки 3+', 'Любое значение или даже пробел в переменной ячейке нумерует строку и включает её в общий тираж.'],
    ['Пустая строка', 'Если во всех переменных ячейках действительно ничего нет, строка не нумеруется и в тираж не входит.'],
    ['Пустая ячейка', 'Конкретное поле в этом экземпляре не заменяется.'],
    ['Только пробелы', 'Строка получает номер и считается резервом без содержательных переменных данных.'],
    ['Столбцы', `В этом шаблоне создано ${requested} переменных полей. Количество можно выбрать перед скачиванием нового шаблона.`],
    ['Сводка', 'Слева автоматически считаются общий тираж, экземпляры с данными и резерв.'],
  ];
  instructions.forEach((row, i) => {
    info.getCell(i + 2, 1).value = row[0];
    info.getCell(i + 2, 2).value = row[1];
  });
  info.getColumn(1).width = 24;
  info.getColumn(2).width = 100;
  info.getColumn(1).font = { bold: true };
  info.getColumn(2).alignment = { wrapText: true, vertical: 'top' };

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'VDP_PRO1100_template.xlsx';
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  status.textContent = `Шаблон Excel скачан: ${requested} переменных пол${requested === 1 ? 'е' : requested < 5 ? 'я' : 'ей'}.`;
}

templateBtn.addEventListener('click', downloadExcelTemplate);

pdfInput.addEventListener('change', e => {
  pdf = e.target.files[0] || null;
  document.getElementById('pdfName').textContent = pdf ? pdf.name : 'Файл не выбран';
  status.textContent = pdf ? 'PDF выбран. Загрузите Excel с данными.' : 'Выберите PDF и Excel.';
  reset();
  refresh();
});

xlsxInput.addEventListener('change', async e => {
  const f = e.target.files[0] || null;
  document.getElementById('xlsxName').textContent = f ? f.name : 'Файл не выбран';
  parsed = null;
  reset();
  if (!f) { refresh(); return; }

  try {
    parsed = parseExcel(XLSX.read(await f.arrayBuffer(), { type: 'array', raw: true }));
    renderExcel(parsed);
    status.textContent = `Excel проверен: тираж ${parsed.used.length}, с переменными данными ${parsed.variable.length}, резерв ${parsed.reserve.length}.`;
  } catch (err) {
    status.textContent = `Ошибка чтения Excel: ${err.message}`;
  }
  refresh();
});

prepare.addEventListener('click', run);
refresh();
