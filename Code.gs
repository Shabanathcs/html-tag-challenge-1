/**
 * HTML Tag Challenge – score collector
 * Paste this into Extensions > Apps Script of a Google Sheet, then
 * Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */
const HEADERS = ['Timestamp', 'Name', 'Class', 'Round 1: Skeleton', 'Round 2: Sort Tags', 'Round 3: Steps', 'Total Points', 'Badge', 'Time Taken (sec)'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const row = [new Date(), clean(d.name), clean(d.cls), num(d.r1), num(d.r2), num(d.r3), num(d.total), clean(d.badge), num(d.seconds)];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    [sheetFor(ss, 'All Scores'), sheetFor(ss, clean(d.cls) || 'No Class')].forEach(sh => sh.appendRow(row));
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('HTML Tag Challenge score collector is running.');
}

function sheetFor(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#2D1B69').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}
function clean(v) { return String(v == null ? '' : v).replace(/^[=+\-@]/, "'").slice(0, 60); }
function num(v) { const n = Number(v); return isFinite(n) ? n : 0; }
