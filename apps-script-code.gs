const SHEET_NAME = 'KV';

// Az admin belépéshez tartozó jelszót NEM ebben a fájlban tároljuk (mert ez a
// fájl publikus GitHub-repóban van), hanem az Apps Script projekt Script
// Properties beállításánál: Project Settings (fogaskerék ikon) → Script
// Properties → "Add script property" → kulcs: ADMIN_PASSWORD, érték: a jelszó.

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

function getKvValue(sheet, key){
  const row = findRow(sheet, key);
  if(row === -1) return null;
  return sheet.getRange(row,2).getValue();
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

  if(body.action === 'admin-login'){
    return handleAdminLogin(body.password);
  }
  if(body.action === 'notify-new-booking'){
    return handleNotifyNewBooking(body.bookingId);
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

function handleAdminLogin(password){
  const props = PropertiesService.getScriptProperties();
  const realPassword = props.getProperty('ADMIN_PASSWORD');
  if(!realPassword){
    return jsonOut({ok:false, error:'not-configured'});
  }
  if(password !== realPassword){
    return jsonOut({ok:false, error:'wrong-password'});
  }
  return jsonOut({ok:true});
}

function handleNotifyNewBooking(bookingId){
  try{
    const sheet = getSheet();

    const settingsRaw = getKvValue(sheet, 'kostolo-ertesites-beallitasok');
    const settings = settingsRaw ? JSON.parse(settingsRaw) : null;
    if(!settings || !settings.enabled || !settings.email){
      return jsonOut({ok:true, skipped:true});
    }

    const bookingsRaw = getKvValue(sheet, 'kostolo-foglalasok');
    const bookings = bookingsRaw ? JSON.parse(bookingsRaw) : [];
    const booking = bookings.find(function(b){ return b.id === bookingId; });
    if(!booking){
      return jsonOut({ok:false, error:'booking-not-found'});
    }

    const subject = 'Új foglalás – ' + booking.eventTitle + ' (' + booking.eventDate + ')';
    const lines = [
      'Új foglalás érkezett a weboldalról:',
      '',
      'Kóstoló/csomag: ' + booking.eventTitle,
      'Dátum: ' + booking.eventDate + (booking.idopont ? ', ' + booking.idopont : ''),
      'Név: ' + booking.nev,
      'Telefon: ' + booking.telefon,
      'Email: ' + booking.email,
      'Létszám: ' + booking.letszam + ' fő'
    ];
    if(booking.addons && booking.addons.length){
      lines.push('Extrák: ' + booking.addons.map(function(a){ return a.label; }).join(', '));
    }

    MailApp.sendEmail(settings.email, subject, lines.join('\n'));
    return jsonOut({ok:true});
  }catch(e){
    return jsonOut({ok:false, error:'send-failed'});
  }
}
