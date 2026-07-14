const SHEET_NAME = 'KV';

// Az admin belépéshez tartozó jelszót NEM ebben a fájlban tároljuk (mert ez a
// fájl publikus GitHub-repóban van), hanem az Apps Script projekt Script
// Properties beállításánál: Project Settings (fogaskerék ikon) → Script
// Properties → "Add script property" → kulcs: ADMIN_PASSWORD, érték: a jelszó.
const OTP_EMAIL = 'vialtumwines@gmail.com'; // ide megy az admin-belépési egyszer használatos kód
const OTP_TTL_MS = 5 * 60 * 1000; // a kód 5 percig érvényes
const OTP_MAX_ATTEMPTS = 5;

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

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e){
  const key = e.parameter.key;
  const sheet = getSheet();
  if(!key){
    return jsonOut({error:'missing key'});
  }
  const row = findRow(sheet, key);
  if(row === -1){
    return jsonOut({value:null});
  }
  const value = sheet.getRange(row,2).getValue();
  return jsonOut({value: value});
}

function doPost(e){
  const body = JSON.parse(e.postData.contents);

  if(body.action === 'admin-login-request'){
    return handleAdminLoginRequest(body.password);
  }
  if(body.action === 'admin-login-verify'){
    return handleAdminLoginVerify(body.otp);
  }

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
  return jsonOut({ok:true});
}

function handleAdminLoginRequest(password){
  const props = PropertiesService.getScriptProperties();
  const realPassword = props.getProperty('ADMIN_PASSWORD');
  if(!realPassword){
    return jsonOut({ok:false, error:'not-configured'});
  }
  if(password !== realPassword){
    return jsonOut({ok:false, error:'wrong-password'});
  }

  const otp = String(Math.floor(100000 + Math.random() * 900000));
  props.setProperty('OTP_CODE', otp);
  props.setProperty('OTP_EXPIRES', String(Date.now() + OTP_TTL_MS));
  props.setProperty('OTP_ATTEMPTS', '0');

  MailApp.sendEmail(
    OTP_EMAIL,
    'Vialtum Pince — admin belépési kód',
    'Admin belépési kód: ' + otp + '\n\nA kód 5 percig érvényes.\n\nHa nem te próbáltál belépni, hagyd figyelmen kívül ezt az emailt.'
  );

  return jsonOut({ok:true});
}

function handleAdminLoginVerify(otp){
  const props = PropertiesService.getScriptProperties();
  const storedOtp = props.getProperty('OTP_CODE');
  const expires = Number(props.getProperty('OTP_EXPIRES') || '0');
  const attempts = Number(props.getProperty('OTP_ATTEMPTS') || '0');

  if(!storedOtp || Date.now() > expires){
    return jsonOut({ok:false, error:'expired'});
  }
  if(attempts >= OTP_MAX_ATTEMPTS){
    return jsonOut({ok:false, error:'too-many-attempts'});
  }
  if(String(otp) !== storedOtp){
    props.setProperty('OTP_ATTEMPTS', String(attempts + 1));
    return jsonOut({ok:false, error:'wrong-otp'});
  }

  props.deleteProperty('OTP_CODE');
  props.deleteProperty('OTP_EXPIRES');
  props.deleteProperty('OTP_ATTEMPTS');
  return jsonOut({ok:true});
}
