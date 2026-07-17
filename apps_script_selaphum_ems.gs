/**
 * Selaphum EMS Google Sheets API.
 *
 * Deploy:
 * 1. Open the existing Google Apps Script project used by selaphum_ems.html.
 * 2. Replace Code.gs with this file, or paste these functions into that project.
 * 3. If the script is not bound to the EMS spreadsheet, set SPREADSHEET_ID below.
 * 4. Deploy > Manage deployments > Edit > New version > Deploy.
 * 5. Keep the same Web app URL in selaphum_ems.html.
 */

const SPREADSHEET_ID = '';

const SHEET_NAME_MAP = {
  personnel: 'ข้อมูลบุคลากร',
  units: 'ข้อมูลหน่วยกู้ชีพ',
  vehicles: 'ยานพาหนะ'
};

const SHEET_HEADER_SIGNATURES = {
  personnel: ['เลขประจำตัวประชาชน', 'ชื่อ', 'สกุล', 'เบอร์โทรศัพท์', 'คุณวุฒิ', 'สังกัด'],
  units: ['ชื่อหน่วยกู้ชีพ', 'ชื่อหน่วย', 'ระดับ', 'เบอร์โทร', 'จำนวนคน', 'ผู้ดูแล', 'ผู้บริหาร'],
  vehicles: ['ทะเบียนรถ', 'รถคันที่', 'จังหวัดทะเบียนรถ', 'สังกัดหน่วย', 'ประเภทพาหนะ', 'สติ๊กเกอร์', 'พรบ']
};

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const sheet = getSheet_(params.sheet);
    return jsonOutput_(readRows_(sheet));
  } catch (error) {
    return jsonOutput_({
      ok: false,
      error: getErrorMessage_(error)
    });
  }
}

function doPost(e) {
  try {
    const payload = parsePayload_(e);
    const action = String(payload.action || '').toLowerCase();
    const sheet = getSheet_(payload.sheet);

    if (action === 'create' || action === 'add') {
      return jsonOutput_(createRow_(sheet, payload.data || {}));
    }

    if (action === 'update' || action === 'edit') {
      return jsonOutput_(updateRow_(sheet, payload));
    }

    if (action === 'delete' || action === 'remove') {
      return jsonOutput_(deleteRow_(sheet, payload));
    }

    throw new Error('Unsupported action: ' + action);
  } catch (error) {
    return jsonOutput_({
      ok: false,
      error: getErrorMessage_(error)
    });
  }
}

function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Missing POST body');
  }

  return JSON.parse(e.postData.contents);
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

function getSheet_(sheetKey) {
  const normalizedKey = String(sheetKey || '').trim();
  const sheetName = SHEET_NAME_MAP[normalizedKey] || normalizedKey;

  if (!sheetName) {
    throw new Error('Missing sheet name');
  }

  const spreadsheet = getSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    return sheet;
  }

  const detectedSheet = findSheetByHeaders_(spreadsheet, normalizedKey);
  if (detectedSheet) {
    return detectedSheet;
  }

  const availableSheets = spreadsheet.getSheets().map(function(item) {
    return item.getName();
  }).join(', ');
  throw new Error('Sheet not found: ' + sheetName + '. Available sheets: ' + availableSheets);
}

function findSheetByHeaders_(spreadsheet, sheetKey) {
  const signatures = SHEET_HEADER_SIGNATURES[sheetKey];
  if (!signatures) return null;

  let bestSheet = null;
  let bestScore = 0;

  spreadsheet.getSheets().forEach(function(sheet) {
    const lastColumn = sheet.getLastColumn();
    if (lastColumn < 1 || sheet.getLastRow() < 1) return;

    const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
      .map(function(header) {
        return String(header || '').trim();
      });
    const score = signatures.reduce(function(total, signature) {
      const hasSignature = headers.some(function(header) {
        return header.indexOf(signature) !== -1;
      });
      return total + (hasSignature ? 1 : 0);
    }, 0);

    if (score > bestScore) {
      bestScore = score;
      bestSheet = sheet;
    }
  });

  return bestScore >= 2 ? bestSheet : null;
}

function readRows_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length === 0) return [];

  const headers = values[0].map(function(header) {
    return String(header || '').trim();
  });

  return values.slice(1).map(function(row, rowIndex) {
    const item = { __rowNumber: rowIndex + 2 };
    headers.forEach(function(header, columnIndex) {
      if (header) item[header] = row[columnIndex] || '';
    });
    return item;
  });
}

function createRow_(sheet, data) {
  const cleanData = cleanData_(data);
  const headers = ensureHeaders_(sheet, cleanData);
  const rowValues = headers.map(function(header) {
    return cleanData[header] !== undefined ? cleanData[header] : '';
  });

  sheet.appendRow(rowValues);
  const rowNumber = sheet.getLastRow();

  return {
    ok: true,
    action: 'create',
    rowNumber: rowNumber,
    row: rowToObject_(headers, rowValues, rowNumber)
  };
}

function updateRow_(sheet, payload) {
  const cleanData = cleanData_(payload.data || {});
  const headers = ensureHeaders_(sheet, cleanData);
  const rowNumber = getTargetRowNumber_(sheet, payload);
  const existingValues = sheet.getRange(rowNumber, 1, 1, headers.length).getDisplayValues()[0];
  const rowValues = headers.map(function(header, columnIndex) {
    return cleanData[header] !== undefined ? cleanData[header] : existingValues[columnIndex] || '';
  });

  sheet.getRange(rowNumber, 1, 1, headers.length).setValues([rowValues]);

  return {
    ok: true,
    action: 'update',
    rowNumber: rowNumber,
    row: rowToObject_(headers, rowValues, rowNumber)
  };
}

function deleteRow_(sheet, payload) {
  const rowNumber = getTargetRowNumber_(sheet, payload);
  sheet.deleteRow(rowNumber);

  return {
    ok: true,
    action: 'delete',
    rowNumber: rowNumber
  };
}

function getTargetRowNumber_(sheet, payload) {
  const rowNumber = Number(payload.rowNumber || payload.rowIndex);

  if (Number.isFinite(rowNumber) && rowNumber >= 2 && rowNumber <= sheet.getLastRow()) {
    return rowNumber;
  }

  throw new Error('Invalid row number');
}

function ensureHeaders_(sheet, data) {
  const keys = Object.keys(data || {}).filter(function(key) {
    return key && !isInternalField_(key);
  });
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  let headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .map(function(header) {
      return String(header || '').trim();
    });

  if (sheet.getLastRow() === 0 || headers.every(function(header) { return !header; })) {
    headers = keys;
    if (headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    return headers;
  }

  const missingHeaders = keys.filter(function(key) {
    return headers.indexOf(key) === -1;
  });

  if (missingHeaders.length) {
    sheet.getRange(1, headers.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    headers = headers.concat(missingHeaders);
  }

  return headers;
}

function cleanData_(data) {
  const cleanData = {};
  Object.keys(data || {}).forEach(function(key) {
    if (!isInternalField_(key)) {
      cleanData[key] = data[key];
    }
  });
  return cleanData;
}

function rowToObject_(headers, rowValues, rowNumber) {
  const item = { __rowNumber: rowNumber };
  headers.forEach(function(header, index) {
    if (header) item[header] = rowValues[index] || '';
  });
  return item;
}

function isInternalField_(field) {
  return String(field || '').indexOf('__') === 0;
}

function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getErrorMessage_(error) {
  return error && error.message ? error.message : String(error);
}
