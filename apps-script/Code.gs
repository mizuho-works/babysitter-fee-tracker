/**
 * ベビーシッター/ファミサポ料金シミュレーター用の依頼状況スプレッドシート連携API。
 * スプレッドシートの「拡張機能 > Apps Script」に貼り付けてデプロイしてください。
 */

const SHEET_NAME = "依頼状況";
const HEADERS = ["id","year","month","day","dow","label","assignment","assignmentName","hours","rate","confirmed"];

// Columns that must stay plain text (1-indexed): id, dow, label, assignment, assignmentName.
// Without this, Sheets auto-converts values like "9/7" into a date, corrupting them.
const TEXT_COLUMNS = [1, 5, 6, 7, 8];

function getSheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet){
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if(sheet.getLastRow() === 0){
    sheet.appendRow(HEADERS);
  }
  ensureTextFormat_(sheet);
  return sheet;
}

function ensureTextFormat_(sheet){
  const rows = Math.max(sheet.getMaxRows(), 1000);
  TEXT_COLUMNS.forEach(col=>{
    sheet.getRange(1, col, rows, 1).setNumberFormat("@");
  });
}

function jsonOut_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function readAllEntries_(){
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  if(data.length < 2) return [];
  const headers = data[0];
  const idCol = headers.indexOf("id");
  return data.slice(1)
    .filter(row => row[idCol] !== "" && row[idCol] !== "id")
    .map(row=>{
      const obj = {};
      headers.forEach((h,i)=>{ obj[h] = row[i]; });
      obj.confirmed = obj.confirmed === true || obj.confirmed === "TRUE" || obj.confirmed === "true";
      return obj;
    });
}

function findRowIndexById_(sheet, id){
  const data = sheet.getDataRange().getValues();
  const idCol = HEADERS.indexOf("id");
  for(let i=1;i<data.length;i++){
    if(String(data[i][idCol]) === String(id)) return i+1;
  }
  return -1;
}

function doGet(e){
  try{
    return jsonOut_({ ok:true, entries: readAllEntries_() });
  }catch(err){
    return jsonOut_({ ok:false, error: String(err) });
  }
}

function doPost(e){
  // Serialize writes: overlapping executions used to interleave clear/append and duplicate rows.
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(30000);
  }catch(lockErr){
    return jsonOut_({ ok:false, error: "busy" });
  }
  try{
    return handlePost_(e);
  }finally{
    lock.releaseLock();
  }
}

function handlePost_(e){
  try{
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if(action === "syncCalendar" || action === "calendarStatus"){
      const keyError = calendarKeyError_(body.key);
      if(keyError) return jsonOut_({ ok:false, error: keyError });
      if(action === "calendarStatus"){
        checkCalendarAccess();
        return jsonOut_({ ok:true });
      }
      return jsonOut_(Object.assign({ ok:true }, syncCalendar_(body)));
    }

    const sheet = getSheet_();

    if(action === "upsert"){
      const entry = body.entry;
      const rowValues = HEADERS.map(h => entry[h] !== undefined ? entry[h] : "");
      const rowIndex = findRowIndexById_(sheet, entry.id);
      if(rowIndex === -1){
        sheet.appendRow(rowValues);
      } else {
        sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([rowValues]);
      }
      return jsonOut_({ ok:true });
    }

    if(action === "delete"){
      const rowIndex = findRowIndexById_(sheet, body.id);
      if(rowIndex !== -1) sheet.deleteRow(rowIndex);
      return jsonOut_({ ok:true });
    }

    if(action === "replaceAll"){
      const seen = {};
      const entries = (body.entries || []).filter(entry=>{
        if(seen[entry.id]) return false;
        seen[entry.id] = true;
        return true;
      });
      sheet.clearContents();
      const rows = [HEADERS].concat(entries.map(entry=> HEADERS.map(h => entry[h] !== undefined ? entry[h] : "")));
      ensureTextFormat_(sheet);
      sheet.getRange(1, 1, rows.length, HEADERS.length).setValues(rows);
      return jsonOut_({ ok:true });
    }

    return jsonOut_({ ok:false, error: "unknown action: " + action });
  }catch(err){
    return jsonOut_({ ok:false, error: String(err) });
  }
}

/**
 * Pushes confirmed entries into a dedicated Google Calendar (which TimeTree can then import).
 * Idempotent: every event carries a tag with the entry id, so sending the same entry again
 * updates (or leaves alone) the existing event instead of creating a second one, and events
 * whose entry is no longer confirmed / no longer exists are removed. Only events created by
 * this script (those with the tag) inside the given range are ever modified or deleted.
 */
const CAL_NAME = "ファミサポ予定(タイムツリー連携用)";
const CAL_TAG = "famisup";
const CAL_KEY_PROPERTY = "CALENDAR_SYNC_KEY";

function calendarKeyError_(key){
  const expected = PropertiesService.getScriptProperties().getProperty(CAL_KEY_PROPERTY);
  if(!expected) return "calendar_key_not_configured";
  if(typeof key !== "string" || !key || key.length > 128) return "calendar_key_invalid";
  const digest = value=>Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  const actualHash = digest(key);
  const expectedHash = digest(expected);
  let difference = 0;
  for(let i = 0; i < expectedHash.length; i++){
    difference |= actualHash[i] ^ expectedHash[i];
  }
  return difference ? "calendar_key_invalid" : "";
}

function checkCalendarAccess(){
  CalendarApp.getDefaultCalendar().getId();
}

function getBridgeCalendar_(){
  const found = CalendarApp.getCalendarsByName(CAL_NAME);
  if(found.length) return found[0];
  return CalendarApp.createCalendar(CAL_NAME, { timeZone: Session.getScriptTimeZone() });
}

function sameDay_(a, b){
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function syncCalendar_(body){
  const items = body.entries || [];
  const rs = body.rangeStart, re = body.rangeEnd;
  const cal = getBridgeCalendar_();
  const start = new Date(rs.year, rs.month - 1, rs.day);
  const end = new Date(re.year, re.month - 1, re.day + 1);

  const existing = {};
  cal.getEvents(start, end).forEach(ev=>{
    const id = ev.getTag(CAL_TAG);
    if(!id) return;
    (existing[id] = existing[id] || []).push(ev);
  });

  const wanted = {};
  let created = 0, updated = 0, unchanged = 0, removed = 0;

  items.forEach(it=>{
    if(wanted[it.id]) return;
    wanted[it.id] = true;
    const date = new Date(it.year, it.month - 1, it.day);
    const evs = existing[it.id] || [];
    if(evs.length === 0){
      const ev = cal.createAllDayEvent(it.title, date, { description: it.description || "" });
      ev.setTag(CAL_TAG, it.id);
      created++;
      return;
    }
    const ev = evs[0];
    evs.slice(1).forEach(extra=>{ extra.deleteEvent(); removed++; });
    let changed = false;
    if(!ev.isAllDayEvent() || !sameDay_(ev.getAllDayStartDate(), date)){ ev.setAllDayDate(date); changed = true; }
    if(ev.getTitle() !== it.title){ ev.setTitle(it.title); changed = true; }
    if((ev.getDescription() || "") !== (it.description || "")){ ev.setDescription(it.description || ""); changed = true; }
    if(changed) updated++; else unchanged++;
  });

  Object.keys(existing).forEach(id=>{
    if(wanted[id]) return;
    existing[id].forEach(ev=>{ ev.deleteEvent(); removed++; });
  });

  return { created: created, updated: updated, unchanged: unchanged, removed: removed, calendar: cal.getName() };
}
