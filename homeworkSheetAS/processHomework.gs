function getColumnIndex(i){
  let char = String.fromCharCode(65 + i%26);
  let times = Math.floor(i/26)+1;
  return(new Array(times).fill(char).join(''));
}
const formula = {
  '剩餘天數': i => `=IFERROR(DATEDIF(TODAY(), $E${i}, "D"), "")`, 
  '顯示狀態': i => `=IFS(AND(NOT($F${i} = ""), $H${i} = TRUE), TRUE, NOT($A${i} = ""), FALSE, TRUE, "")`, 
  '啟用狀態': () => 'TRUE'
};

function homeworkUpdate() {
  /* load sheet */
  let spreadsheet = SpreadsheetApp.openById('1T1piWWBZ5fCFXxh92T587GFt_Fab2fYujdlZ8L7hHcE');
  const sheets = {
    allHomework: '所有作業', 
    newHomework: '新增作業', 
    updateHomework: '更新作業', 
    variables: '變數存放'
  };
  for(let key in sheets){
    sheets[key] = spreadsheet.getSheetByName(sheets[key]);
  }
  let variables = sheets.variables.getDataRange().getValues();
  function getVariable(name){
    return(variables[1][variables[0].indexOf(name)]);
  }
  function setVariable(name, value){
    let column = variables[0].indexOf(name);
    variables[1][column] = value;
    sheets.variables.getRange(`${getColumnIndex(column)}2`).setValues([[value]]);
  }

  /* new homework */
  (() => {
    let allHomeworkData = sheets.allHomework.getDataRange().getValues();
    let targetTitle = allHomeworkData[0];
    let data = sheets.newHomework.getDataRange().getValues();
    let mode = sheets.newHomework.getRange('F1').getValues()[0][0];
    if(mode == '保存'){
      let maxIdNow = getVariable('最大編號');
      for(let i = 1; i < data.length; i++){
        maxIdNow++;
        var lastRowIndex = sheets.allHomework.getLastRow();
        let rowObject = {};
        data[i].map((v, j) => {
          rowObject[data[0][j]] = v;
        });
        let newRow = targetTitle.map(n => 
        n == '項目編號' ? 
        maxIdNow : 
        rowObject[n] !== undefined ? 
        rowObject[n] : 
        formula[n] !== undefined ? 
        formula[n](lastRowIndex+1) : 
        '');
        console.log(newRow);
        sheets.allHomework.insertRowsAfter(lastRowIndex, 1);
        sheets.allHomework.appendRow(newRow);
      }
      sheets.newHomework.getRange('2:1000').clear({
        formatOnly: true,
        contentsOnly: true
      });
      setVariable('最大編號', maxIdNow);
    }
    else if(mode == '清空'){
      sheets.newHomework.getRange('2:1000').clear({
        formatOnly: true,
        contentsOnly: true
      });
    }
    sheets.newHomework.getRange('F1').setValues([['~ 動作 ~']]);
  })();
  
  /* update homework */
  (() => {
    let data = sheets.updateHomework.getDataRange().getValues();
    let mode = sheets.updateHomework.getRange('G1').getValues()[0][0];
    console.log(mode);
    if(mode == '載入'){
      let allHomeworkData = sheets.allHomework.getDataRange().getValues();
      let targetTitle = data[0];
      let idList = allHomeworkData.map(n => n[0]);
      for(let i = 1; i < data.length; i++){
        let rowObject = {};
        data[i].map((v, j) => {
          if(v !== '') rowObject[data[0][j]] = v;
        });
        let rowIndex = idList.indexOf(rowObject['更新編號']);
        oriRow = allHomeworkData[rowIndex];
        let oriRowObject = {};
        oriRow.map((v, j) => {
          if(v !== '') oriRowObject[allHomeworkData[0][j]] = v;
        });
        let newRow = targetTitle.map((n, i) => 
        rowObject[n] !== undefined ? 
        rowObject[n] : 
        oriRowObject[n] !== undefined ? 
        oriRowObject[n] : 
        '');
        console.log(newRow);
        sheets.updateHomework.getRange(`A${i+1}:G${i+1}`).setValues([newRow]);
      }
    }
    else if(mode == '更新'){
      let allHomeworkData = sheets.allHomework.getDataRange().getValues();
      let targetTitle = allHomeworkData[0];
      let idList = allHomeworkData.map(n => n[0]);
      for(let i = 1; i < data.length; i++){
        let rowObject = {};
        data[i].map((v, j) => {
          if(v !== '') rowObject[data[0][j]] = v;
        });
        let rowIndex = idList.indexOf(rowObject['更新編號']);
        oriRow = allHomeworkData[rowIndex];
        let newRow = targetTitle.map((n, i) => 
        rowObject[n] !== undefined ? 
        rowObject[n] : 
        formula[n] !== undefined ? 
        formula[n](rowIndex+1) : 
        oriRow[i] !== undefined ? 
        oriRow[i] : 
        '');
        console.log(newRow);
        sheets.allHomework.getRange(`A${rowIndex+1}:I${rowIndex+1}`).setValues([newRow]);
      }
      sheets.updateHomework.getRange('2:1000').clear({
        formatOnly: true,
        contentsOnly: true
      });
    }
    else if(mode == '清空'){
      sheets.updateHomework.getRange('2:1000').clear({
        formatOnly: true,
        contentsOnly: true
      });
    }
    else if(['啟用', '禁用'].indexOf(mode) > -1){
      let allHomeworkData = sheets.allHomework.getDataRange().getValues();
      let targetTitle = allHomeworkData[0];
      let column = getColumnIndex(targetTitle.indexOf('啟用狀態'));
      let idList = allHomeworkData.map(n => n[0]);
      for(let i = 1; i < data.length; i++){
        let rowObject = {};
        data[i].map((v, j) => {
          if(v !== '') rowObject[data[0][j]] = v;
        });
        let rowIndex = idList.indexOf(rowObject['更新編號']);
        console.log(rowIndex);
        sheets.allHomework.getRange(`${column}${rowIndex+1}`).setValues([[mode == '啟用' ? 'TRUE' : 'FALSE']]);
      }
      sheets.updateHomework.getRange('2:1000').clear({
        formatOnly: true,
        contentsOnly: true
      });
    }
    sheets.updateHomework.getRange('G1').setValues([['~ 動作 ~']]);
  })();
}
