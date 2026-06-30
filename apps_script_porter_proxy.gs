/**
 * Smart Porter Google Sheets proxy for GitHub Pages.
 *
 * Deploy:
 * 1. Open https://script.google.com and create a new project.
 * 2. Paste this file into Code.gs.
 * 3. Deploy > New deployment > Web app.
 * 4. Execute as: Me.
 * 5. Who has access: Anyone.
 * 6. Copy the Web app URL into APPS_SCRIPT_PROXY_URL in porter_dashboard.html.
 */

const SPREADSHEET_ID = '1NqVI8vdWm4x4Svf-nvVqkk0s1hYPg-qHvkUwjPalMqs';
const DEFAULT_GID = '800055064';

// Optional: set a token such as 'my-secret' and call ?token=my-secret.
// Keep empty to allow the dashboard to read without a token.
const ACCESS_TOKEN = '';

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    validateAccess_(params);

    const gid = String(params.gid || DEFAULT_GID);
    const format = String(params.format || 'csv').toLowerCase();
    const sheet = getSheetByGid_(gid);
    const values = sheet.getDataRange().getDisplayValues();

    if (format === 'json') {
      return jsonOutput_({
        ok: true,
        spreadsheetId: SPREADSHEET_ID,
        gid: gid,
        sheetName: sheet.getName(),
        updatedAt: new Date().toISOString(),
        values: values
      });
    }

    return csvOutput_(values);
  } catch (error) {
    return jsonOutput_({
      ok: false,
      error: error && error.message ? error.message : String(error)
    });
  }
}

function validateAccess_(params) {
  if (!ACCESS_TOKEN) return;

  if (params.token !== ACCESS_TOKEN) {
    throw new Error('Unauthorized');
  }
}

function getSheetByGid_(gid) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = spreadsheet.getSheets();
  const sheet = sheets.find(item => String(item.getSheetId()) === String(gid));

  if (!sheet) {
    throw new Error(`Sheet gid ${gid} was not found`);
  }

  return sheet;
}

function csvOutput_(values) {
  return ContentService
    .createTextOutput(toCsv_(values))
    .setMimeType(ContentService.MimeType.CSV);
}

function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function toCsv_(values) {
  return values.map(row => row.map(escapeCsvCell_).join(',')).join('\r\n');
}

function escapeCsvCell_(value) {
  const text = value === null || value === undefined ? '' : String(value);

  if (/[",\r\n]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }

  return text;
}
