// Terikat pada Google Sheet RSVP Luqman & Izzati.
// Deploy > Manage deployments > Edit > New version > Deploy.
// Kekalkan Execute as: Me dan Who has access: Anyone.
// RSVP disimpan dalam tab RSVP; ucapan tanpa RSVP dalam tab Ucapan.
const TAB_NAME = 'RSVP';
const HEADERS = ['id', 'name', 'attendance', 'guests', 'wish', 'timestamp'];
const NOTE_HEADERS = ['id', 'name', 'wish', 'timestamp'];

function getSheet_(tabName, headers) {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = book.getSheetByName(tabName);
  if (!sheet) sheet = book.insertSheet(tabName);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  return sheet;
}
function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function literal_(value, max) {
  const text = String(value || '').trim().slice(0, max);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}
function upsert_(sheet, id, row) {
  const values = sheet.getDataRange().getValues();
  const existing = values.findIndex((r, i) => i > 0 && String(r[0]) === id);
  if (existing > 0) sheet.getRange(existing + 1, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return json_({ok:false, error:'Sila cuba lagi'});
  try {
    const d = JSON.parse(e.postData.contents);
    const name = String(d.name || '').trim().slice(0, 100);
    if (name.length < 2) return json_({ok:false, error:'Nama tidak lengkap'});
    const id = String(d.id || Utilities.getUuid()).slice(0, 100);
    const timestamp = new Date().toISOString();
    if (d.type === 'wish') {
      const wish = String(d.wish || '').trim().slice(0, 800);
      if (!wish) return json_({ok:false, error:'Ucapan tidak lengkap'});
      if (!/^wish-[\w-]+$/.test(id)) return json_({ok:false, error:'ID ucapan tidak sah'});
      const sheet = getSheet_('Ucapan', NOTE_HEADERS);
      upsert_(sheet, id, [literal_(id,100), literal_(name,100), literal_(wish,800), timestamp]);
      return json_({ok:true, type:'wish', id:id});
    }
    if (!['Hadir','Tidak Hadir'].includes(d.attendance)) {
      return json_({ok:false, error:'Sila pilih kehadiran'});
    }
    const guests = d.attendance === 'Hadir' ? Math.max(1, Math.min(10, Number(d.guests) || 1)) : 0;
    const sheet = getSheet_(TAB_NAME, HEADERS);
    upsert_(sheet, id, [literal_(id,100), literal_(name,100), d.attendance,
      guests, literal_(d.wish,800), timestamp]);
    return json_({ok:true, type:'rsvp', id:id});
  } catch (error) {
    return json_({ok:false, error:'Permintaan tidak sah'});
  } finally { lock.releaseLock(); }
}
function doGet() {
  const rsvp = getSheet_(TAB_NAME, HEADERS).getDataRange().getValues().slice(1);
  const notes = getSheet_('Ucapan', NOTE_HEADERS).getDataRange().getValues().slice(1);
  const iso = value => value instanceof Date ? value.toISOString() : String(value);
  const wishes = rsvp.filter(r => String(r[4]).trim()).map(r => ({
    id:String(r[0]), name:String(r[1]), wish:String(r[4]), timestamp:iso(r[5])
  })).concat(notes.filter(r => String(r[2]).trim()).map(r => ({
    id:String(r[0]), name:String(r[1]), wish:String(r[2]), timestamp:iso(r[3])
  }))).sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp)).slice(0,300);
  return json_({version:2, capabilities:['wish'], wishes:wishes});
}
