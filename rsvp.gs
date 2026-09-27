// Google Apps Script backend for the "Eu vou!" RSVP button.
// Setup: Google Sheet -> Extensions -> Apps Script -> paste this -> Save.
//   Project Settings (gear) -> Script properties -> add INVITE_CODE = the code you put after # in the shared link.
//   Deploy -> New deployment -> type "Web app", Execute as: Me, Who has access: Anyone
//   -> copy the /exec URL into RSVP_URL in index.html.
// The invite code lives only in Script properties, never in this public file.

const MAX_ROWS = 150;          // hard cap so spam can't flood the sheet
const MAX_PER_VISITOR = 3;     // same browser can log at most this many taps

const FIELDS = ['receivedAt', 'name', 'guests', 'visitorId', 'tapNumber', 'tappedAt', 'localTime', 'timezone',
  'secondsOnPage', 'language', 'languages', 'userAgent', 'platform', 'mobile', 'screen', 'viewport',
  'pixelRatio', 'touch', 'darkMode', 'connection', 'referrer', 'page', 'musicPlaying'];

function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return reply('bad'); }

  const code = PropertiesService.getScriptProperties().getProperty('INVITE_CODE');
  if (!code || data.code !== code) return reply('denied');
  if (!data.name || String(data.name).length > 60) return reply('bad');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(FIELDS);
    const rows = sheet.getLastRow() - 1;
    if (rows >= MAX_ROWS) return reply('full');

    if (rows > 0) {
      const col = FIELDS.indexOf('visitorId') + 1;
      const ids = sheet.getRange(2, col, rows, 1).getValues().flat();
      if (ids.filter(id => id === data.visitorId).length >= MAX_PER_VISITOR) return reply('dup');
    }

    data.receivedAt = new Date();
    // prefix a quote so text starting with = + - @ is never run as a formula
    sheet.appendRow(FIELDS.map(f => {
      const v = data[f] ?? '';
      return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
    }));
    return reply('ok');
  } finally {
    lock.releaseLock();
  }
}

function reply(s) { return ContentService.createTextOutput(s); }
