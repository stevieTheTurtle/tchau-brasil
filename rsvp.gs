// Google Apps Script backend for the "Eu vou!" RSVP button.
// Setup: Google Sheet -> Extensions -> Apps Script -> paste this -> Deploy -> New deployment
//   type "Web app", Execute as: Me, Who has access: Anyone -> copy the /exec URL into RSVP_URL in index.html.

const FIELDS = ['receivedAt', 'name', 'guests', 'visitorId', 'tapNumber', 'tappedAt', 'localTime', 'timezone',
  'secondsOnPage', 'language', 'languages', 'userAgent', 'platform', 'mobile', 'screen', 'viewport',
  'pixelRatio', 'touch', 'darkMode', 'connection', 'referrer', 'page', 'musicPlaying'];

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) sheet.appendRow(FIELDS);
  const data = JSON.parse(e.postData.contents);
  data.receivedAt = new Date();
  sheet.appendRow(FIELDS.map(f => data[f] ?? ''));
  return ContentService.createTextOutput('ok');
}
