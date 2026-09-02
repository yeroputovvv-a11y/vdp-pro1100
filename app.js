const pdfInput = document.getElementById('pdfFile');
const xlsxInput = document.getElementById('xlsxFile');
const prepare = document.getElementById('prepare');
const status = document.getElementById('status');
const templateBtn = document.getElementById('downloadTemplate');

let pdf = null;
let parsed = null;
let pdfAnalysis = null;

pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

function rawText(value) { return value === null || value === undefined ? '' : String(value); }
function hasAnyContent(value) { return rawText(value).length > 0; }
function hasMeaningfulValue(value) { return rawText(value).trim().length > 0; }
function rowUsed(row) { return row.some(hasAnyContent); }
function rowHasVariableData(row) { return row.some(hasMeaningfulValue); }
function escapeHtml(value) { return rawText(value).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#039;'); }
function normalizeForSearch(text) { return rawText(text).replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim(); }
function countOccurrences(text, needle) { if (!needle) return 0; let count=0, pos=0; while(true){ const found=text.indexOf(needle,pos); if(found<0) break; count++; pos=found+Math.max(needle.length,1); } return count; }
function refreshButton() { prepare.disabled = !(pdf && parsed && parsed.ok); }

function resetAnalysis() {
  document.getElementById('summary').hidden=true;
  document.getElementById('fields').hidden=true;
  document.getElementById('preview').hidden=true;
  document.getElementById('fieldList').innerHTML='';
  document.getElementById('previewTable').innerHTML='';
}

function groupTextItems(items) {
  const usable=items.filter(i=>typeof i.str==='string' && i.str.length).map((i,index)=>({
    index,text:i.str,x:Number(i.transform?.[4]||0),y:Number(i.transform?.[5]||0),width:Number(i.width||0),height:Number(i.height||Math.abs(i.transform?.[3]||0)),fontName:i.fontName||''
  }));
  const lines=[]; const yTolerance=2.5;
  for(const item of usable){
    let line=lines.find(candidate=>Math.abs(candidate.y-item.y)<=yTolerance);
    if(!line){line={y:item.y,items:[]};lines.push(line);} line.items.push(item);
  }
  lines.sort((a,b)=>b.y-a.y);
  return lines.map(line=>{
    line.items.sort((a,b)=>a.x-b.x);
    let text='';
    for(let i=0;i<line.items.length;i++){
      const item=line.items[i];
      if(text){
        const prev=line.items[i-1];
        const gap=item.x-(prev.x+prev.width);
        if(gap>Math.max(1,item.height*0.15) && !/[\s(\[«]$/.test(text) && !/^[,.;:!?%)\]}]/.test(item.text)) text+=' ';
      }
      text+=item.text;
    }
    return { text:normalizeForSearch(text), x:Math.min(...line.items.map(i=>i.x)), y:line.y, height:Math.max(...line.items.map(i=>i.height)), itemIndexes:line.items.map(i=>i.index) };
  });
}

async function analyzePdf(file, fields) {
  const data=new Uint8Array(await file.arrayBuffer());
  const doc=await pdfjsLib.getDocument({data}).promise;
  const pages=[]; let textItemsTotal=0;
  for(let pageNumber=1;pageNumber<=doc.numPages;pageNumber++){
    const page=await doc.getPage(pageNumber);
    const content=await page.getTextContent();
    textItemsTotal+=content.items.length;
    pages.push({pageNumber,lines:groupTextItems(content.items)});
  }

  const normalizedRefs=new Map();
  for(const field of fields){
    const ref=normalizeForSearch(field.reference);
    if(!ref) continue;
    const arr=normalizedRefs.get(ref)||[]; arr.push(field.name); normalizedRefs.set(ref,arr);
  }

  const results=fields.map(field=>{
    const reference=normalizeForSearch(field.reference);
    if(!reference) return {...field,count:0,locations:[],state:'empty-reference'};
    const duplicateNames=normalizedRefs.get(reference)||[];
    let count=0; const locations=[];
    for(const page of pages){
      for(const line of page.lines){
        const c=countOccurrences(line.text,reference);
        if(!c) continue;
        count+=c;
        locations.push({page:page.pageNumber,x:Math.round(line.x*100)/100,y:Math.round(line.y*100)/100,count:c,excerpt:line.text});
      }
    }
    const duplicate=duplicateNames.length>1;
    return {...field,count,locations,state:duplicate?'duplicate':count?'found':'not-found',duplicate,duplicateNames};
  });
  return {numPages:doc.numPages,textItemsTotal,fields:results};
}

function renderPdfResults(analysis){
  const list=document.getElementById('fieldList');
  list.innerHTML=analysis.fields.map((field,index)=>{
    const checked=field.count>0 && field.state!=='duplicate' ? 'checked' : '';
    let stateText='';
    if(field.state==='empty-reference') stateText='Эталон пустой — поиск не выполнялся';
    else if(field.state==='duplicate') stateText=`Эталон совпадает с полями: ${field.duplicateNames.map(escapeHtml).join(', ')}`;
    else if(field.count===0) stateText='Совпадений не найдено';
    else stateText=`Найдено совпадений: ${field.count}`;
    const locations=field.locations.length ? field.locations.map(loc=>`<span class="location">стр. ${loc.page}${loc.count>1?` ×${loc.count}`:''}</span>`).join(' ') : '';
    return `<div class="field-row ${field.state==='not-found'?'field-warning':''} ${field.state==='duplicate'?'field-error':''}">
      <div class="field-check"><label><input type="checkbox" data-field-index="${index}" ${checked}> <b>Заменять найденные</b></label></div>
      <div class="field-name">${escapeHtml(field.name)}</div>
      <div class="field-reference"><span class="muted-small">Эталон:</span> ${escapeHtml(field.reference||'—')}</div>
      <div class="field-meta"><b>${escapeHtml(stateText)}</b>${locations?`<div class="locations">${locations}</div>`:''}</div>
    </div>`;
  }).join('') || '<div class="empty">Нет переменных полей для анализа.</div>';
  document.getElementById('pdfNote').textContent=`PDF полностью проанализирован: ${analysis.numPages} стр., ${analysis.textItemsTotal} текстовых объектов. Для каждого эталона показано число совпадений; замена пока не выполняется.`;
  document.getElementById('fields').hidden=false;
}

async function runFullAnalysis(){
  status.textContent='Анализирую весь PDF…'; prepare.disabled=true;
  try{
    pdfAnalysis=await analyzePdf(pdf,parsed.variableFields);
    renderPdfResults(pdfAnalysis);
    const totalFound=pdfAnalysis.fields.reduce((sum,f)=>sum+f.count,0);
    const problems=pdfAnalysis.fields.filter(f=>['not-found','duplicate'].includes(f.state)).length;
    status.textContent=problems?`Анализ завершён: найдено ${totalFound} совпадений. Полей для проверки: ${problems}.`:`Анализ завершён: найдено ${totalFound} совпадений. Все эталоны найдены.`;
  }catch(error){pdfAnalysis=null;status.textContent=`Ошибка анализа PDF: ${error.message}`;}
  finally{refreshButton();}
}

function renderExcelData(parsedData){
  const {headers,records,usedRows,variableRows,reserveRows,variableFields}=parsedData;
  document.getElementById('totalPrint').textContent=usedRows.length.toLocaleString('ru-RU');
  document.getElementById('variablePrint').textContent=variableRows.length.toLocaleString('ru-RU');
  document.getElementById('reservePrint').textContent=reserveRows.length.toLocaleString('ru-RU');
  document.getElementById('fieldCount').textContent=variableFields.length.toLocaleString('ru-RU');
  document.getElementById('summary').hidden=false;
  const sample=records.slice(0,8);
  const tableHead=['№',...headers.slice(1)].map(h=>`<th>${escapeHtml(h||'')}</th>`).join('');
  const tableBody=sample.map(record=>`<tr><td class="row-no">${record.index}</td>${record.displayValues.map(v=>`<td>${escapeHtml(v)}</td>`).join('')}</tr>`).join('');
  document.getElementById('previewTable').innerHTML=`<div class="table-wrap"><table><thead><tr>${tableHead}</tr></thead><tbody>${tableBody}</tbody></table></div>`;
  document.getElementById('preview').hidden=false;
}

async function downloadExcelTemplate(){
  if(!window.ExcelJS){status.textContent='Не удалось загрузить модуль создания Excel. Обновите страницу.';return;}
  const workbook=new ExcelJS.Workbook(); workbook.creator='VDP PRO 1100'; workbook.created=new Date();
  const sheet=workbook.addWorksheet('Данные',{views:[{state:'frozen',ySplit:2,xSplit:1}]});
  sheet.addRow(['№','Номер изделия','Фамилия','Имя','Отчество','Звание']);
  sheet.addRow(['ЭТАЛОН','529260704929','Иванов','Иван','Иванович','майор']);
  for(let r=3;r<=1002;r++){sheet.getCell(r,1).value={formula:`IF(COUNTA(B${r}:F${r})>0,MAX($A$2:A${r-1})+1,"")`};for(let c=2;c<=6;c++)sheet.getCell(r,c).value='';}
  sheet.getCell('H1').value='Показатель'; sheet.getCell('I1').value='Значение';
  sheet.getCell('H2').value='Общий тираж'; sheet.getCell('I2').value={formula:'=COUNT(A3:A1002)'};
  sheet.getCell('H3').value='С переменными данными'; sheet.getCell('I3').value={formula:'=SUMPRODUCT(--(MMULT(--(LEN(TRIM(B3:F1002&""))>0),TRANSPOSE(COLUMN(B3:F3)^0))>0))'};
  sheet.getCell('H4').value='Резерв без данных'; sheet.getCell('I4').value={formula:'=I2-I3'};
  const headerFill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1F2937'}}; const refFill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF2CC'}};
  sheet.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}}; sheet.getRow(1).fill=headerFill; sheet.getRow(1).alignment={horizontal:'center',vertical:'middle',wrapText:true};
  sheet.getRow(2).font={bold:true,italic:true,color:{argb:'FF7C3AED'}}; sheet.getRow(2).fill=refFill; sheet.getRow(2).alignment={horizontal:'center',vertical:'middle',wrapText:true};
  sheet.getCell('H1').font={bold:true,color:{argb:'FFFFFFFF'}}; sheet.getCell('I1').font={bold:true,color:{argb:'FFFFFFFF'}}; sheet.getCell('H1').fill=headerFill; sheet.getCell('I1').fill=headerFill;
  ['H2','H3','H4'].forEach(a=>sheet.getCell(a).font={bold:true}); ['I2','I3','I4'].forEach(a=>sheet.getCell(a).font={bold:true,size:12});
  [8,20,20,18,22,18,2,24,18].forEach((w,i)=>sheet.getColumn(i+1).width=w);
  for(let r=1;r<=1002;r++){sheet.getCell(r,1).alignment={horizontal:'center',vertical:'middle'};sheet.getCell(r,1).numFmt='0';}
  const info=workbook.addWorksheet('Инструкция'); info.mergeCells('A1:C1'); info.getCell('A1').value='ШАБЛОН ДАННЫХ VDP PRO 1100'; info.getCell('A1').font={bold:true,color:{argb:'FFFFFFFF'},size:14}; info.getCell('A1').fill=headerFill; info.getCell('A1').alignment={horizontal:'center'};
  [['Строка 1','Названия колонок.'],['Строка 2','ЭТАЛОННЫЕ значения. По ним программа ищет совпадения по всему PDF.'],['Строки 3+','Одна пронумерованная строка = один печатаемый экземпляр.'],['Пустая ячейка','Данное поле в конкретном экземпляре не заменяется.'],['Пробел','Строка участвует в тираже и получает номер; без содержательных значений это резерв.']].forEach(row=>info.addRow(row));
  info.getColumn(1).width=20; info.getColumn(2).width=90; info.getColumn(2).alignment={wrapText:true,vertical:'top'}; info.getColumn(1).font={bold:true};
  const buffer=await workbook.xlsx.writeBuffer(); const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}); const url=URL.createObjectURL(blob); const link=document.createElement('a'); link.href=url; link.download='VDP_PRO1100_template.xlsx'; document.body.appendChild(link); link.click(); link.remove(); setTimeout(()=>URL.revokeObjectURL(url),1000); status.textContent='Шаблон Excel скачан.';
}

templateBtn.addEventListener('click',downloadExcelTemplate);
pdfInput.addEventListener('change',event=>{pdf=event.target.files[0]||null;document.getElementById('pdfName').textContent=pdf?pdf.name:'Файл не выбран';pdfAnalysis=null;status.textContent=pdf?'PDF выбран. Загрузите Excel с данными.':'Выберите PDF и Excel.';refreshButton();});

xlsxInput.addEventListener('change',async event=>{
  const file=event.target.files[0]||null; document.getElementById('xlsxName').textContent=file?file.name:'Файл не выбран'; parsed=null; pdfAnalysis=null; resetAnalysis();
  if(!file){status.textContent='Выберите PDF и Excel.';refreshButton();return;}
  try{
    const data=await file.arrayBuffer(); const book=XLSX.read(data,{type:'array',cellDates:false,raw:true}); if(!book.SheetNames.length)throw new Error('В книге нет листов.');
    const sheet=book.Sheets[book.SheetNames[0]]; const matrix=XLSX.utils.sheet_to_json(sheet,{header:1,defval:'',raw:true}); if(matrix.length<2)throw new Error('Нужны как минимум две строки: названия колонок и эталонные значения.');
    const width=Math.max(matrix[0]?.length||0,matrix[1]?.length||0,...matrix.slice(2).map(row=>row.length)); if(width<2)throw new Error('Нужно минимум одно переменное поле кроме колонки №.');
    const headers=Array.from({length:width},(_,i)=>rawText(matrix[0]?.[i]??'')); const reference=Array.from({length:width},(_,i)=>rawText(matrix[1]?.[i]??'')); if(!headers.slice(1).some(Boolean))throw new Error('В строке 1 нет названий переменных колонок.');
    const records=[],usedRows=[],variableRows=[],reserveRows=[];
    for(let sourceRow=2;sourceRow<matrix.length;sourceRow++){
      const row=Array.from({length:width},(_,i)=>rawText(matrix[sourceRow]?.[i]??'')); if(!rowUsed(row.slice(1)))continue;
      const variable=rowHasVariableData(row.slice(1)); const record={index:usedRows.length+1,sourceRow:sourceRow+1,variable,values:row,displayValues:row.slice(1)}; records.push(record);usedRows.push(record);if(variable)variableRows.push(record);else reserveRows.push(record);
    }
    const variableFields=[]; for(let c=1;c<width;c++){const name=headers[c];if(!name)continue;const usedInData=records.filter(r=>hasMeaningfulValue(r.values[c])).length;variableFields.push({index:c,name,reference:reference[c],usedInData});}
    parsed={ok:true,headers,reference,records,usedRows,variableRows,reserveRows,variableFields}; renderExcelData(parsed); status.textContent=`Excel проверен: ${usedRows.length} экземпляров в тираже, ${variableRows.length} с переменными данными, ${reserveRows.length} резервных. Нажмите «ПРОВЕРИТЬ PDF И ДАННЫЕ».`;
  }catch(error){status.textContent=`Ошибка чтения Excel: ${error.message}`;}
  refreshButton();
});

prepare.addEventListener('click',runFullAnalysis);
refreshButton();
