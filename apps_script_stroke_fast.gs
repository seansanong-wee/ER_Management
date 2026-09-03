/**
 * Stroke Fast Track Google Sheets API for stroke_fast.html.
 *
 * Deploy:
 * 1. Open the Stroke Fast Track Google Sheet.
 * 2. Extensions > Apps Script.
 * 3. Paste this file into Code.gs.
 * 4. If the script is not bound to that spreadsheet, set SPREADSHEET_ID.
 * 5. Deploy > New deployment > Web app.
 * 6. Execute as: Me, Who has access: Anyone.
 * 7. Copy the Web app URL into GAS_API_URL in stroke_fast.html.
 */

const SPREADSHEET_ID = '';
const DEFAULT_SHEET_NAME = '';
const ACCESS_TOKEN = '';

function doGet(e) {
  const params = (e && e.parameter) || {};

  try {
    validateAccess_(params);

    const sheet = getSheet_(params.sheet);
    const values = sheet.getDataRange().getDisplayValues();

    return output_(params, {
      ok: true,
      sheetName: sheet.getName(),
      updatedAt: new Date().toISOString(),
      values: values,
      records: valuesToRecords_(values)
    });
  } catch (error) {
    return output_(params, {
      ok: false,
      error: getErrorMessage_(error)
    });
  }
}

function validateAccess_(params) {
  if (!ACCESS_TOKEN) return;

  if (params.token !== ACCESS_TOKEN) {
    throw new Error('Unauthorized');
  }
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('Set SPREADSHEET_ID because this script is not bound to a spreadsheet');
  }

  return spreadsheet;
}

function getSheet_(sheetNameFromRequest) {
  const spreadsheet = getSpreadsheet_();
  const sheetName = String(sheetNameFromRequest || DEFAULT_SHEET_NAME || '').trim();

  if (sheetName) {
    const namedSheet = spreadsheet.getSheetByName(sheetName);
    if (!namedSheet) throw new Error('Sheet not found: ' + sheetName);
    return namedSheet;
  }

  const sheets = spreadsheet.getSheets();
  if (!sheets.length) throw new Error('Spreadsheet has no sheets');

  return sheets[0];
}

function valuesToRecords_(values) {
  if (!values || values.length < 2) return [];

  const headers = values[0].map(function(header) {
    return String(header || '').trim();
  });

  return values.slice(1).map(function(row, rowIndex) {
    const record = { __rowNumber: rowIndex + 2 };

    headers.forEach(function(header, columnIndex) {
      if (header) record[header] = row[columnIndex] || '';
    });

    return record;
  });
}

function output_(params, payload) {
  const callback = String(params.callback || params.jsonp || '').trim();
  const json = JSON.stringify(payload);

  if (callback) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callback)) {
      throw new Error('Invalid JSONP callback name');
    }

    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getErrorMessage_(error) {
  return error && error.message ? error.message : String(error);
}
