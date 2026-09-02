const pdfInput=document.getElementById('pdfFile');
const xlsxInput=document.getElementById('xlsxFile');
const prepare=document.getElementById('prepare');
const status=document.getElementById('status');
const download=document.getElementById('download');
let pdf=null, workbook=null;

function refresh(){prepare.disabled=!(pdf&&workbook)}
pdfInput.addEventListener('change',e=>{pdf=e.target.files[0]||null;document.getElementById('pdfName').textContent=pdf?pdf.name:'Файл не выбран';refresh()});
xlsxInput.addEventListener('change',async e=>{const f=e.target.files[0]||null;document.getElementById('xlsxName').textContent=f?f.name:'Файл не выбран';workbook=null;if(!f){refresh();return}try{const data=await f.arrayBuffer();workbook=XLSX.read(data,{type:'array',cellDates:false});const sheet=workbook.Sheets[workbook.SheetNames[0]];const rows=XLSX.utils.sheet_to_json(sheet,{header:1,raw:false});const count=Math.max(0,rows.length-1);status.textContent=`Excel прочитан: ${count} номеров. Нажмите «Подготовить файл».`; }catch(err){status.textContent='Ошибка чтения Excel: '+err.message}refresh()});
prepare.addEventListener('click',()=>{status.textContent='Интерфейс готов. Следующим этапом подключим точную обработку PDF без изменения шрифта и координат.';download.hidden=true});
