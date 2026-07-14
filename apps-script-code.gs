const SHEET_NAME = 'KV';

function getSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet){
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['key','value','updated_at']);
  }
  return sheet;
}

function findRow(sheet, key){
  const data = sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(data[i][0] === key) return i+1;
  }
  return -1;
}

function doGet(e){
  const key = e.parameter.key;
  const sheet = getSheet();
  if(!key){
    return ContentService.createTextOutput(JSON.stringify({error:'missing key'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const row = findRow(sheet, key);
  if(row === -1){
    return ContentService.createTextOutput(JSON.stringify({value:null}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const value = sheet.getRange(row,2).getValue();
  return ContentService.createTextOutput(JSON.stringify({value: value}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  const body = JSON.parse(e.postData.contents);
  const key = body.key;
  const value = body.value;
  const sheet = getSheet();
  const row = findRow(sheet, key);
  const now = new Date().toISOString();
  if(row === -1){
    sheet.appendRow([key, value, now]);
  } else {
    sheet.getRange(row,2).setValue(value);
    sheet.getRange(row,3).setValue(now);
  }
  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}
