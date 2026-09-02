(function(){
  const oldButton=document.getElementById('downloadTemplate');
  if(!oldButton) return;
  const fresh=oldButton.cloneNode(true);
  oldButton.replaceWith(fresh);

  function colLetter(n){
    let s='';
    while(n>0){
      const r=(n-1)%26;
      s=String.fromCharCode(65+r)+s;
      n=Math.floor((n-1)/26);
    }
    return s;
  }

  function makeTemplate(fieldCount,maxRows){
    const names=[];
    const refs=[];
    for(let i=1;i<=fieldCount;i++){
      names.push(i===1?'Номер изделия':i===2?'Фамилия':i===3?'Имя':i===4?'Отчество':i===5?'Звание':`Поле ${i}`);
      refs.push(i===1?'529260704929':i===2?'Иванов':i===3?'Иван':i===4?'Иванович':i===5?'майор':'');
    }

    const firstVar=5; // E
    const lastVar=firstVar+fieldCount-1;
    const helperCol=lastVar+1;
    const firstVarLetter=colLetter(firstVar);
    const lastVarLetter=colLetter(lastVar);
    const helperLetter=colLetter(helperCol);
    const rowStart=5;
    const rowEnd=maxRows+4;

    const rows=[];
    rows.push(['Показатель','Значение','','№',...names,'__data_flag']);
    rows.push(['Общий тираж','','','ЭТАЛОН',...refs,'']);
    rows.push(['С переменными данными','','','','','','','','']);
    rows.push(['Резерв без данных','','','','','','','','']);

    for(let r=rowStart;r<=rowEnd;r++){
      const numberFormula=`IF(COUNTA(${firstVarLetter}${r}:${lastVarLetter}${r})>0,COUNT($D$${rowStart}:D${r-1})+1,"")`;
      const meaningfulParts=[];
      for(let c=firstVar;c<=lastVar;c++){
        const letter=colLetter(c);
        meaningfulParts.push(`LEN(TRIM(${letter}${r}&""))>0`);
      }
      const helperFormula=`IF(OR(${meaningfulParts.join(',')}),1,0)`;
      const row=new Array(helperCol).fill('');
      row[3]={f:numberFormula};
      row[helperCol-1]={f:helperFormula};
      rows.push(row);
    }

    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws['B2']={f:`COUNT(D${rowStart}:D${rowEnd})`};
    ws['B3']={f:`SUM(${helperLetter}${rowStart}:${helperLetter}${rowEnd})`};
    ws['B4']={f:'B2-B3'};

    const header={fill:{patternType:'solid',fgColor:{rgb:'1F2937'}},font:{bold:true,color:{rgb:'FFFFFF'}},alignment:{horizontal:'center',vertical:'center',wrap_text:true}};
    const ref={fill:{patternType:'solid',fgColor:{rgb:'FFF2CC'}},font:{bold:true,italic:true,color:{rgb:'7C3AED'}},alignment:{horizontal:'center',vertical:'center',wrap_text:true}};
    const headerCells=['A1','B1','D1'];
    for(let i=0;i<fieldCount;i++) headerCells.push(XLSX.utils.encode_cell({r:0,c:firstVar-1+i}));
    for(const a of headerCells) if(ws[a]) ws[a].s=header;
    const refCells=['D2'];
    for(let i=0;i<fieldCount;i++) refCells.push(XLSX.utils.encode_cell({r:1,c:firstVar-1+i}));
    for(const a of refCells) if(ws[a]) ws[a].s=ref;
    for(const a of ['A2','A3','A4']) if(ws[a]) ws[a].s={font:{bold:true}};
    for(let r=rowStart;r<=rowEnd;r++){
      const a=XLSX.utils.encode_cell({r:r-1,c:3});
      if(ws[a]) ws[a].s={font:{bold:true},alignment:{horizontal:'center'}};
    }

    ws['!cols']=[{wch:24},{wch:16},{wch:3},{wch:7}];
    for(let i=0;i<fieldCount;i++) ws['!cols'].push({wch:20});
    ws['!cols'].push({wch:3,hidden:true});
    ws['!freeze']={xSplit:4,ySplit:2};
    return ws;
  }

  function download(){
    try{
      if(!window.XLSX){
        alert('Модуль Excel ещё не загрузился. Нажмите Ctrl+F5 и повторите.');
        return;
      }
      const input=document.getElementById('fieldCountInput');
      let fieldCount=Number.parseInt(input?.value||'5',10);
      if(!Number.isFinite(fieldCount)||fieldCount<1) fieldCount=5;
      if(fieldCount>100) fieldCount=100;
      if(input) input.value=String(fieldCount);

      const wb=XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb,makeTemplate(fieldCount,1000),'Данные');
      const info=XLSX.utils.aoa_to_sheet([
        ['ШАБЛОН ДАННЫХ VDP PRO 1100',''],
        ['Строка 1','Названия колонок. Колонка № обязательна; переменные поля идут справа от неё.'],
        ['Строка 2','ЭТАЛОННЫЕ значения для поиска по всему PDF.'],
        ['Строки 5+','Любое значение или пробел в переменной ячейке включает строку в тираж и даёт номер.'],
        ['Пустая строка','Не нумеруется и не входит в тираж.'],
        ['Только пробелы','Нумеруются, но считаются резервом без содержательных переменных данных.'],
        ['Переменные поля','Количество полей можно менять для каждого заказа.']
      ]);
      info['!cols']=[{wch:28},{wch:100}];
      XLSX.utils.book_append_sheet(wb,info,'Инструкция');
      XLSX.writeFile(wb,'VDP_PRO1100_template.xlsx');
      const status=document.getElementById('status');
      if(status) status.textContent='Шаблон Excel сформирован и скачивается.';
    }catch(error){
      console.error(error);
      const status=document.getElementById('status');
      if(status) status.textContent='Ошибка создания шаблона Excel: '+error.message;
    }
  }

  fresh.addEventListener('click',download);
})();
