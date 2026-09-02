(function(){
  function makeTemplate(){
    if(!window.XLSX){ alert('Модуль Excel не загрузился. Нажмите Ctrl+F5 и повторите.'); return; }
    const wb = XLSX.utils.book_new();
    const vars = ['Номер изделия','Фамилия','Имя','Отчество','Звание'];
    const rows = [
      ['Показатель','Значение','','№',...vars,'__data_flag'],
      ['Общий тираж','', '', 'ЭТАЛОН','529260704929','Иванов','Иван','Иванович','майор',''],
      ['С переменными данными','', '', '','','','','','',''],
      ['Резерв без данных','', '', '','','','','','','']
    ];
    for(let r=5;r<=1004;r++){
      rows.push(['','','',`=IF(COUNTA(E${r}:I${r})>0,MAX($D$2:D${r-1})+1,"")`,'','','','','',`=IF(OR(LEN(TRIM(E${r}&""))>0,LEN(TRIM(F${r}&""))>0,LEN(TRIM(G${r}&""))>0,LEN(TRIM(H${r}&""))>0,LEN(TRIM(I${r}&""))>0),1,0)`]);
    }
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['B2']={f:'COUNT(D5:D1004)'};
    ws['B3']={f:'SUM(J5:J1004)'};
    ws['B4']={f:'B2-B3'};
    ws['!cols']=[{wch:24},{wch:16},{wch:3},{wch:7},{wch:20},{wch:20},{wch:18},{wch:22},{wch:18},{wch:3}];
    const header={fill:{patternType:'solid',fgColor:{rgb:'1F2937'}},font:{bold:true,color:{rgb:'FFFFFF'}},alignment:{horizontal:'center',vertical:'center',wrap_text:true}};
    const ref={fill:{patternType:'solid',fgColor:{rgb:'FFF2CC'}},font:{bold:true,italic:true,color:{rgb:'7C3AED'}},alignment:{horizontal:'center',vertical:'center',wrap_text:true}};
    ['A1','B1','D1','E1','F1','G1','H1','I1','J1'].forEach(a=>{if(ws[a]) ws[a].s=header;});
    ['D2','E2','F2','G2','H2','I2'].forEach(a=>{if(ws[a]) ws[a].s=ref;});
    ws['!freeze']={xSplit:4,ySplit:2};
    XLSX.utils.book_append_sheet(wb,ws,'Данные');
    const info=[
      ['ШАБЛОН ДАННЫХ VDP PRO 1100',''],
      ['Строка 1','Названия колонок. № обязательна; переменные поля идут справа от неё.'],
      ['Строка 2','ЭТАЛОННЫЕ значения для поиска по всему PDF.'],
      ['Строки 5+','Любое значение или пробел в переменной ячейке включает строку в тираж и даёт номер.'],
      ['Полностью пустая строка','Не нумеруется и не входит в тираж.'],
      ['Только пробелы','Нумеруются, но считаются резервом без содержательных переменных данных.']
    ];
    const wi=XLSX.utils.aoa_to_sheet(info); wi['!cols']=[{wch:26},{wch:100}]; XLSX.utils.book_append_sheet(wb,wi,'Инструкция');
    XLSX.writeFile(wb,'VDP_PRO1100_template.xlsx');
    const status=document.getElementById('status'); if(status) status.textContent='Шаблон Excel сформирован и скачивается.';
  }
  const old=document.getElementById('downloadTemplate');
  if(!old) return;
  const fresh=old.cloneNode(true);
  old.replaceWith(fresh);
  fresh.addEventListener('click',makeTemplate);
})();
