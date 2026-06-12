/**
 * Brigade Eternia — Lead Capture Web App
 *
 * Setup:
 * 1. Open the "Brigade Eternia" Google Sheet.
 * 2. Extensions > Apps Script — paste this whole file in as Code.gs.
 * 3. Run `setupSheet` once (top toolbar > Run, select setupSheet) and
 *    approve the permission prompts. This creates the formatted header row.
 * 4. Deploy > New deployment > type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the deployment URL and paste it into SHEET_ENDPOINT in index.html
 *    (replacing the existing script.google.com/.../exec URL).
 */

const SHEET_NAME = 'Sheet1';
const HEADERS = ['Timestamp', 'Name', 'Phone', 'Email', 'Form Source'];

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
    || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  ensureHeaders(sheet);

  const data = (e && e.parameter) || {};
  sheet.appendRow([
    new Date(),
    data.name  || '',
    data.phone || '',
    data.email || '',
    data.source || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    formatHeaders(sheet);
  }
}

function formatHeaders(sheet) {
  const range = sheet.getRange(1, 1, 1, HEADERS.length);
  range.setFontWeight('bold')
       .setFontColor('#ffffff')
       .setBackground('#274e13')
       .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);
}

/** Run once manually from the Apps Script editor to initialize the sheet. */
function setupSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
    || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureHeaders(sheet);
}
