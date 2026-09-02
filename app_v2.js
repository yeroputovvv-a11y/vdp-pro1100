const pdfInput=document.getElementById('pdfFile');
const xlsxInput=document.getElementById('xlsxFile');
const prepare=document.getElementById('prepare');
const status=document.getElementById('status');
const templateBtn=document.getElementById('downloadTemplate');
const fieldCountInput=document.getElementById('fieldCountInput');
let pdf=null;
let parsed=null;
pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
const S=v=>v==null?'':String(v);
const T=v=>S(v).trim();
const E=v=>S(v).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#039;');
function reset(){for(const id of ['summary','fields','preview'])document.getElementById(id).hidden=true;document.getElementById('fieldList').innerHTML='';document.getElementById('previewTable').innerHTML='';}
function refresh(){prepare.disabled=!(pdf&&parsed?.ok);}
function parseExcel(book){
  const sh=book.Sheets[book.SheetNames[0]],m=XLSX.utils.sheet_to_json(sh,{header:1,defval:'',raw:true});
  if(m.length<2)throw Error('Нужны строки 1 и 2.');
  const h=(m[0]||[]).map(S),r=(m[1]||[]).map(S);
  let nc=h.findIndex(x=>['№','номер','номер строки'].includes(T(x).toLowerCase()));if(nc<0)nc=0;
  const width=Math.max(h.length,r.length,...m.slice(2).map(x=>x.length));
  const fields=[];for(let c=nc+1;c<width;c++)if(T(h[c]))fields.push({col:c,name:h[c],ref:r[c]||'',used:0});
  if(!fields.length)throw Error('После колонки «№» нет названных переменных.');
  const used=[],variable=[],reserve=[];
  for(let rr=2;rr<m.length;rr++){
    const row=Array.from({length:width},(_,i)=>S(m[rr]?.[i])),vals=fields.map(f=>row[f.col]);
    if(!vals.some(v=>S(v).length))continue;
    const has=vals.some(v=>T(v).length),rec={no:used.length+1,values:vals};used.push(rec);(has?variable:reserve).push(rec);
  }
  return{ok:true,fields,used,variable,reserve};
}
function renderExcel(p){
  document.getElementById('totalPrint').textContent=p.used.length.toLocaleString('ru-RU');
  document.getElementById('variablePrint').textContent=p.variable.length.toLocaleString('ru-RU');
  document.getElementById('reservePrint').textContent=p.reserve.length.toLocaleString('ru-RU');
  document.getElementById('fieldCount').textContent=p.fields.length.toLocaleString('ru-RU');
  document.getElementById('summary').hidden=false;
  const head='<th>№</th>'+p.fields.map(f=>`<th>${E(f.name)}</th>`).join('');
  const body=p.used.slice(0,8).map(x=>`<tr><td class="row-no">${x.no}</td>${x.values.map(v=>`<td>${E(v)}</td>`).join('')}</tr>`).join('');
  document.getElementById('previewTable').innerHTML=`<div class="table-wrap"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div>`;
  document.getElementById('preview').hidden=false;
}
function makeLines(items){
  const groups=[];
  for(const i of items.filter(x=>x.str)){
    const z={t:i.str,x:+(i.transform?.[4]||0),y:+(i.transform?.[5]||0),w:+(i.width||0),h:+(i.height||Math.abs(i.transform?.[3]||0))};
    let g=groups.find(a=>Math.abs(a.y-z.y)<=2.5);if(!g){g={y:z.y,a:[]};groups.push(g);}g.a.push(z);
  }
  return groups.sort((a,b)=>b.y-a.y).map(g=>{g.a.sort((a,b)=>a.x-b.x);let s='';for(let i=0;i<g.a.length;i++){if(i){const p=g.a[i-1],q=g.a[i],gap=q.x-(p.x+p.w);if(gap>Math.max(1,q.h*.15)&&!/[\s(\[«]$/.test(s)&&!/^[,.;:!?%)\]}]/.test(q.t))s+=' ';}s+=g.a[i].t;}return{t:T(s)};});
}
function occ(a,b){let n=0,p=0;if(!b)return 0;while((p=a.indexOf(b,p))>=0){n++;p+=Math.max(1,b.length);}return n;}
async function analyze(file,fields){
  const doc=await pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise,pages=[];
  for(let p=1;p<=doc.numPages;p++){const c=await(await doc.getPage(p)).getTextContent();pages.push({p,lines:makeLines(c.items)});}
  const same=new Map();fields.forEach(f=>{const ref=T(f.ref);if(ref){const a=same.get(ref)||[];a.push(f.name);same.set(ref,a);}});
  return fields.map(f=>{const ref=T(f.ref);if(!ref)return{...f,count:0,loc:[],state:'empty'};let n=0,loc=[];for(const pg of pages)for(const line of pg.lines){const c=occ(line.t,ref);if(c){n+=c;loc.push({p:pg.p,c});}}const duplicate=(same.get(ref)||[]).length>1;return{...f,count:n,loc,state:duplicate?'duplicate':n?'found':'not-found'};});
}
async function run(){
  prepare.disabled=true;status.textContent='Анализирую весь PDF…';
  try{const a=await analyze(pdf,parsed.fields);document.getElementById('fieldList').innerHTML=a.map(f=>{const icon=f.state==='found'?'✓':f.state==='not-found'?'⚠':f.state==='duplicate'?'⛔':'—';const msg=f.state==='found'?`Найдено совпадений: ${f.count}`:f.state==='not-found'?'Совпадений не найдено':f.state==='duplicate'?'Эталон совпадает с другим полем':'Эталон не задан';const loc=f.loc.map(x=>`<span class="location">стр. ${x.p}${x.c>1?' ×'+x.c:''}</span>`).join('');return`<div class="field-row ${f.state==='not-found'?'field-warning':''} ${f.state==='duplicate'?'field-error':''}"><div class="field-check">${icon}</div><div class="field-name">${E(f.name)}</div><div class="field-reference">Эталон: ${E(f.ref||'—')}</div><div class="field-meta"><b>${msg}</b><div class="locations">${loc}</div></div></div>`;}).join('');document.getElementById('fields').hidden=false;const total=a.reduce((q,f)=>q+f.count,0),problems=a.filter(f=>['not-found','duplicate'].includes(f.state)).length;status.textContent=problems?`PDF проверен: ${total} совпадений. Требуют внимания: ${problems}.`:`PDF проверен: ${total} совпадений. Все эталонные значения найдены.`;}catch(e){status.textContent='Ошибка анализа PDF: '+e.message;}finally{refresh();}
}
function columnLetter(n){return XLSX.utils.encode_col(n-1);}
function buildTemplateSheet(fieldCount,maxRows){
  const names=[];for(let i=1;i<=fieldCount;i++)names.push(i===1?'Номер изделия':i===2?'Фамилия':i===3?'Имя':i===4?'Отчество':i===5?'Звание':`Поле ${i}`);
  const refs=[];for(let i=1;i<=fieldCount;i++)refs.push(i===1?'529260704929':i===2?'Иванов':i===3?'Иван':i===4?'Иванович':i===5?'майор':'');
  const variableStartCol=5,variableEndCol=variableStartCol+fieldCount-1,helperCol=variableEndCol+1;
  const varStart=columnLetter(variableStartCol),varEnd=columnLetter(variableEndCol),helper=columnLetter(helperCol);
  const rows=[['Показатель','Значение','','№',...names,'__data_flag'],['Общий тираж','','','ЭТАЛОН',...refs,''] ,['С переменными данными','','','','','','','',''],['Резерв без данных','','','','','','','','']];
  for(let r=5;r<=maxRows+4;r++){
    const anyFormula=`IF(COUNTA(${varStart}${r}:${varEnd}${r})>0,ROW()-4,"")`;
    const parts=names.map((_,i)=>{const col=columnLetter(variableStartCol+i);return`LEN(TRIM(${col}${r}&""))>0`;});
    const dataFormula=`IF(OR(${parts.join(',')}),1,0)`;
    rows.push(['','','',`=${anyFormula}`,...Array(fieldCount).fill(''),`=${dataFormula}`]);
  }
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws['B5']={f:`COUNT(D5:D${maxRows+4})`};
  ws['B6']={f:`SUM(${helper}5:${helper}${maxRows+4})`};
  ws['B7']={f:'B5-B6'};
  ws['!cols']=[{wch:24},{wch:16},{wch:3},{wch:7}];for(let i=0;i<fieldCount;i++)ws['!cols'].push({wch:20});ws['!cols'].push({wch:3,hidden:true});
  const hdr={fill:{patternType:'solid',fgColor:{rgb:'1F2937'}},font:{bold:true,color:{rgb:'FFFFFF'}},alignment:{horizontal:'center',vertical:'center',wrap_text:true}};
  const ref={fill:{patternType:'solid',fgColor:{rgb:'FFF2CC'}},font:{bold:true,italic:true,color:{rgb:'7C3AED'}},alignment:{horizontal:'center',vertical:'center',wrap_text:true}};
  ['A1','B1','D1'].concat(names.map((_,i)=>XLSX.utils.encode_cell({r:0,c:variableStartCol-1+i}))).forEach(a=>{if(ws[a])ws[a].s=hdr;});
  for(let c=4;c<=variableEndCol;c++){const a=XLSX.utils.encode_cell({r:1,c:c-1});if(ws[a])ws[a].s=ref;}
  for(let r=5;r<=maxRows+4;r++){ws[XLSX.utils.encode_cell({r:r-1,c:3})].s={font:{bold:true}};}
  ws['A5'].s={font:{bold:true}};ws['A6'].s={font:{bold:true}};ws['A7'].s={font:{bold:true}};
  ws['!freeze']={xSplit:4,ySplit:2};
  return ws;
}
function downloadTemplate(){
  try{
    if(!window.XLSX){status.textContent='Модуль Excel не загрузился. Обновите страницу (Ctrl+F5).';return;}
    let count=Number.parseInt(fieldCountInput?.value||'5',10);if(!Number.isFinite(count)||count<1)count=5;if(count>100)count=100;
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,buildTemplateSheet(count,1000),'Данные');
    const info=XLSX.utils.aoa_to_sheet([['ШАБЛОН ДАННЫХ VDP PRO 1100',''],['Строка 1','Названия колонок. № обязательна; переменные поля находятся справа от неё.'],['Строка 2','ЭТАЛОННЫЕ значения для поиска по всему PDF.'],['Строки 5+','Любое значение или пробел в переменной ячейке включает строку в тираж и даёт номер.'],['Пустая строка','Не нумеруется и не входит в тираж.'],['Только пробелы','Нумеруются, но считаются резервом без содержательных переменных данных.']]);
    info['!cols']=[{wch:28},{wch:100}];XLSX.utils.book_append_sheet(wb,info,'Инструкция');
    XLSX.writeFile(wb,'VDP_PRO1100_template.xlsx');status.textContent='Шаблон Excel сформирован и скачивается.';
  }catch(e){console.error(e);status.textContent='Ошибка создания шаблона Excel: '+e.message;}
}
templateBtn.addEventListener('click',downloadTemplate);
pdfInput.addEventListener('change',e=>{pdf=e.target.files[0]||null;document.getElementById('pdfName').textContent=pdf?pdf.name:'Файл не выбран';status.textContent=pdf?'PDF выбран. Загрузите Excel с данными.':'Выберите PDF и Excel.';reset();refresh();});
xlsxInput.addEventListener('change',async e=>{const f=e.target.files[0]||null;document.getElementById('xlsxName').textContent=f?f.name:'Файл не выбран';parsed=null;reset();if(!f){refresh();return;}try{parsed=parseExcel(XLSX.read(await f.arrayBuffer(),{type:'array',raw:true}));renderExcel(parsed);status.textContent=`Excel проверен: тираж ${parsed.used.length}, с переменными данными ${parsed.variable.length}, резерв ${parsed.reserve.length}.`;}catch(err){status.textContent='Ошибка чтения Excel: '+err.message;}refresh();});
prepare.addEventListener('click',run);refresh();
