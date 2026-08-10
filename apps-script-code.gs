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
  const historyKey = e.parameter.history;
  const sheet = getSheet();

  if(historyKey){
    return jsonOut({entries: getHistoryFor(historyKey)});
  }

  if(!key){
    return jsonOut({error:'missing key'});
  }
  const row = findRow(sheet, key);
  if(row === -1){
    return jsonOut({value:null, updated_at:null});
  }
  const value = sheet.getRange(row,2).getValue();
  const updatedAt = sheet.getRange(row,3).getValue();
  return jsonOut({value: value, updated_at: updatedAt ? String(updatedAt) : null});
}

// A History munkalapról adott kulcshoz tartozó bejegyzéseket adja vissza
// (legújabb elöl), hogy egy véletlen felülírás után kézzel visszakereshető
// legyen egy korábbi állapot ?history=<kulcs> paraméterrel.
function getHistoryFor(key){
  try{
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hist = ss.getSheetByName('History');
    if(!hist) return [];
    const data = hist.getDataRange().getValues();
    const entries = [];
    for(let i=1;i<data.length;i++){
      if(data[i][1] === key){
        entries.push({timestamp: data[i][0], old_value: data[i][2], new_value: data[i][3]});
      }
    }
    entries.reverse();
    return entries;
  }catch(err){
    return [];
  }
}

// Ütközés-védelem: a kliens minden GET-nél megjegyzi a kapott updated_at-ot, és
// visszaküldi expected_updated_at-ként a következő mentésnél. Ha időközben más
// session már mentett (tehát a Sheet-ben lévő updated_at eltér az elvárttól),
// a mentést elutasítjuk ahelyett, hogy csendben felülírnánk a köztes változást
// (ez okozta korábban, hogy egy régebbi admin-fül mentése eltüntetett eseményeket).
//
// FONTOS: ha a kulcshoz már van mentett adat, az expected_updated_at KÖTELEZŐ —
// ha hiányzik (pl. mert egy még régebbi, e védelem előtti kódot futtató,
// napok/hetek óta nyitva hagyott admin-fül küldte a kérést), a mentést EZ IS
// ütközésként utasítja el, nem csak akkor, ha ténylegesen eltér az időbélyeg.
// E nélkül a szigorítás nélkül pont egy ilyen "néma" régi fül tudta kétszer is
// csendben felülírni/eltüntetni a friss adatokat (2026-07-28, 2026-07-29).
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

  if(row !== -1){
    const currentUpdatedAt = sheet.getRange(row,3).getValue();
    const currentUpdatedAtStr = currentUpdatedAt ? String(currentUpdatedAt) : '';
    const expected = body.expected_updated_at ? String(body.expected_updated_at) : '';
    if(expected !== currentUpdatedAtStr){
      return jsonOut({
        ok:false,
        error:'conflict',
        current_value: sheet.getRange(row,2).getValue(),
        current_updated_at: currentUpdatedAt ? String(currentUpdatedAt) : null
      });
    }
  }

  if(row === -1){
    sheet.appendRow([key, value, now]);
  } else {
    const oldValue = sheet.getRange(row,2).getValue();
    logHistory(key, oldValue, value, now);
    sheet.getRange(row,2).setValue(value);
    sheet.getRange(row,3).setValue(now);
  }
  return jsonOut({ok:true, updated_at: now});
}

// Minden felülírás előtt elmenti a régi értéket egy külön "History" munkalapra,
// hogy egy jövőbeli véletlen/ütközéses felülírás esetén vissza lehessen keresni
// és kézzel visszaállítani a korábbi állapotot. Ha a naplózás elhasal, a mentést
// nem akasztja meg.
function logHistory(key, oldValue, newValue, timestamp){
  try{
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hist = ss.getSheetByName('History');
    if(!hist){
      hist = ss.insertSheet('History');
      hist.appendRow(['timestamp','key','old_value','new_value']);
    }
    hist.appendRow([timestamp, key, oldValue, newValue]);
  }catch(err){
    // szándékosan elnyelve, a mentés a napló nélkül is menjen át
  }
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
