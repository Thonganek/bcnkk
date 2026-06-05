const SPREADSHEET_ID = '';
const SHEET_NAME = 'ข้อมูลนักเรียนทุน';
const ACCESS_USERNAME = 'admin';
const ACCESS_PASSWORD = '12345';
const CANONICAL_HEADERS = ['ลำดับ', 'ชื่อ-สกุล', 'สถาบันการศึกษา', 'ชั้นปี', 'เกรดเฉลี่ย'];

const FIELD_ALIASES = {
  no: ['ลำดับ', 'เลขที่', 'No', 'no', 'Number', 'number'],
  name: ['ชื่อ-สกุล', 'ชื่อ สกุล', 'ชื่อ', 'ชื่อนักศึกษา', 'ชื่อนักเรียน', 'ชื่อผู้รับทุน', 'Name', 'name'],
  institution: ['สถาบันการศึกษา', 'สถานศึกษา', 'วิทยาลัย', 'มหาวิทยาลัย', 'Institution', 'institution', 'School', 'school'],
  year: ['ชั้นปี', 'ปี', 'ระดับชั้น', 'Year', 'year', 'Class', 'class'],
  gpa: ['เกรดเฉลี่ย', 'GPA', 'gpa', 'เกรด', 'GPAX', 'gpax']
};

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const callback = params.callback;

  try {
    if (params.action === 'saveRows') {
      return handleSave_(params, callback);
    }

    return handleRead_(params, callback);
  } catch (error) {
    return outputJson_({
      success: false,
      error: error && error.message ? error.message : String(error)
    }, callback);
  }
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents)
      : {};

    if (body.action === 'saveRows') {
      return handleSave_(body, body.callback);
    }

    return outputJson_({ success: false, error: 'ไม่พบ action ที่รองรับ' }, body.callback);
  } catch (error) {
    return outputJson_({
      success: false,
      error: error && error.message ? error.message : String(error)
    });
  }
}

function handleRead_(params, callback) {
  const isAuthenticated = isAuthenticated_(params);
  const sheet = getDataSheet_();
  const rows = readStudents_(sheet);
  const summary = buildSummary_(rows);
  const payload = {
    success: true,
    authenticated: isAuthenticated,
    sheetName: sheet.getName(),
    headers: CANONICAL_HEADERS,
    summary: summary
  };

  if (isAuthenticated) {
    payload.data = rows;
  }

  return outputJson_(payload, callback);
}

function handleSave_(params, callback) {
  if (!isAuthenticated_(params)) {
    return outputJson_({ success: false, error: 'ไม่มีสิทธิ์บันทึกข้อมูล' }, callback);
  }

  const rows = typeof params.rows === 'string'
    ? JSON.parse(params.rows || '[]')
    : (params.rows || []);

  if (!Array.isArray(rows) || rows.length === 0) {
    return outputJson_({ success: false, error: 'ไม่พบข้อมูลสำหรับบันทึก' }, callback);
  }

  const sheet = getDataSheet_();
  const result = upsertRows_(sheet, rows);
  const normalizedRows = readStudents_(sheet);

  return outputJson_({
    success: true,
    authenticated: true,
    saved: result.saved,
    updated: result.updated,
    appended: result.appended,
    skipped: result.skipped,
    sheetName: sheet.getName(),
    headers: CANONICAL_HEADERS,
    summary: buildSummary_(normalizedRows),
    data: normalizedRows
  }, callback);
}

function getDataSheet_() {
  const spreadsheet = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error('ไม่พบไฟล์ Google Sheets กรุณาใส่ SPREADSHEET_ID หรือผูก Apps Script กับไฟล์ชีต');
  }

  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.getSheets()[0];

  if (!sheet) {
    throw new Error('ไม่พบชีตในไฟล์ Google Sheets');
  }

  ensureHeaders_(sheet);
  return sheet;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, CANONICAL_HEADERS.length).setValues([CANONICAL_HEADERS]);
    return;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), CANONICAL_HEADERS.length))
    .getDisplayValues()[0];
  const hasAnyHeader = currentHeaders.some(function(header) {
    return String(header || '').trim() !== '';
  });

  if (!hasAnyHeader) {
    sheet.getRange(1, 1, 1, CANONICAL_HEADERS.length).setValues([CANONICAL_HEADERS]);
    return;
  }

  sheet.getRange(1, 1, 1, CANONICAL_HEADERS.length).setValues([CANONICAL_HEADERS]);
}

function readStudents_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.length
    ? values[0].map(function(header, index) {
        return String(header || '').trim() || 'column_' + (index + 1);
      })
    : CANONICAL_HEADERS;

  return values.slice(1)
    .filter(function(row) {
      return row.some(function(cell) {
        return String(cell || '').trim() !== '';
      });
    })
    .map(function(row) {
      const item = {};
      headers.forEach(function(header, index) {
        item[header] = row[index] || '';
      });
      return item;
    })
    .map(function(row, index) {
      return normalizeStudent_(row, index);
    })
    .filter(function(row) {
      return row.name !== '';
    });
}

function upsertRows_(sheet, rows) {
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.length ? values[0] : CANONICAL_HEADERS;
  const columnMap = { no: 0, name: 1, institution: 2, year: 3, gpa: 4 };
  const existingByNo = {};
  let maxNo = 0;
  let updated = 0;
  let appended = 0;
  let skipped = 0;

  for (let i = 1; i < values.length; i++) {
    const no = String(values[i][columnMap.no] || '').trim();
    const noNumber = parseNumber_(no);
    if (no) existingByNo[no] = i + 1;
    if (!isNaN(noNumber) && noNumber > maxNo) maxNo = noNumber;
  }

  rows.forEach(function(inputRow) {
    const student = normalizeInputStudent_(inputRow);
    if (!student.name) {
      skipped++;
      return;
    }

    let no = student.no ? String(student.no).trim() : '';
    if (!no) {
      maxNo++;
      no = String(maxNo);
      student.no = no;
    }

    const rowValues = [];
    rowValues[columnMap.no] = student.no;
    rowValues[columnMap.name] = student.name;
    rowValues[columnMap.institution] = student.institution;
    rowValues[columnMap.year] = student.year;
    rowValues[columnMap.gpa] = student.gpa;

    if (existingByNo[no]) {
      sheet.getRange(existingByNo[no], 1, 1, CANONICAL_HEADERS.length).setValues([rowValues]);
      updated++;
    } else {
      sheet.appendRow(rowValues);
      existingByNo[no] = sheet.getLastRow();
      appended++;
    }
  });

  return {
    saved: updated + appended,
    updated: updated,
    appended: appended,
    skipped: skipped
  };
}

function buildColumnMap_(headers) {
  const map = {
    no: findColumn_(headers, FIELD_ALIASES.no),
    name: findColumn_(headers, FIELD_ALIASES.name),
    institution: findColumn_(headers, FIELD_ALIASES.institution),
    year: findColumn_(headers, FIELD_ALIASES.year),
    gpa: findColumn_(headers, FIELD_ALIASES.gpa)
  };

  if (map.no < 0 || map.name < 0 || map.institution < 0 || map.year < 0 || map.gpa < 0) {
    sheetHeadersReset_();
    return { no: 0, name: 1, institution: 2, year: 3, gpa: 4 };
  }

  return map;
}

function sheetHeadersReset_() {
  // Kept as a named no-op to avoid changing existing sheet layouts unexpectedly.
}

function findColumn_(headers, aliases) {
  const normalizedAliases = aliases.map(normalizeKey_);
  for (let i = 0; i < headers.length; i++) {
    if (normalizedAliases.indexOf(normalizeKey_(headers[i])) > -1) {
      return i;
    }
  }
  return -1;
}

function normalizeStudent_(row, index) {
  const year = parseNumber_(readField_(row, FIELD_ALIASES.year));
  const gpa = parseNumber_(readField_(row, FIELD_ALIASES.gpa));

  return {
    no: readField_(row, FIELD_ALIASES.no) || String(index + 1),
    name: readField_(row, FIELD_ALIASES.name),
    institution: readField_(row, FIELD_ALIASES.institution) || '-',
    year: isNaN(year) || year === 0 ? '' : String(year),
    gpa: isNaN(gpa) ? '' : gpa.toFixed(2)
  };
}

function normalizeInputStudent_(row) {
  const year = parseNumber_(readField_(row, FIELD_ALIASES.year));
  const gpa = parseNumber_(readField_(row, FIELD_ALIASES.gpa));

  return {
    no: readField_(row, FIELD_ALIASES.no),
    name: readField_(row, FIELD_ALIASES.name),
    institution: readField_(row, FIELD_ALIASES.institution) || '-',
    year: isNaN(year) || year === 0 ? '' : String(year),
    gpa: isNaN(gpa) ? '' : gpa.toFixed(2)
  };
}

function buildSummary_(rows) {
  const byYearMap = {};
  let totalGpa = 0;
  let validGpaCount = 0;

  rows.forEach(function(row) {
    const yearNumber = parseNumber_(row.year);
    const yearKey = isNaN(yearNumber) || yearNumber === 0 ? 'ไม่ระบุ' : 'ปี ' + yearNumber;
    const gpa = parseNumber_(row.gpa);

    if (!byYearMap[yearKey]) {
      byYearMap[yearKey] = {
        year: yearKey,
        yearNumber: isNaN(yearNumber) ? 0 : yearNumber,
        count: 0,
        gpaTotal: 0,
        gpaCount: 0,
        maxGpa: null,
        minGpa: null
      };
    }

    byYearMap[yearKey].count++;

    if (!isNaN(gpa) && gpa > 0) {
      byYearMap[yearKey].gpaTotal += gpa;
      byYearMap[yearKey].gpaCount++;
      byYearMap[yearKey].maxGpa = byYearMap[yearKey].maxGpa === null ? gpa : Math.max(byYearMap[yearKey].maxGpa, gpa);
      byYearMap[yearKey].minGpa = byYearMap[yearKey].minGpa === null ? gpa : Math.min(byYearMap[yearKey].minGpa, gpa);
      totalGpa += gpa;
      validGpaCount++;
    }
  });

  const byYear = [1, 2, 3, 4]
    .map(function(yearNumber) {
      const item = byYearMap['ปี ' + yearNumber];

      if (!item) {
        return {
          year: 'ปี ' + yearNumber,
          yearNumber: yearNumber,
          count: 0,
          avgGpa: 0,
          maxGpa: 0,
          minGpa: 0,
          gpaCount: 0
        };
      }

      return {
        year: item.year,
        yearNumber: item.yearNumber,
        count: item.count,
        avgGpa: item.gpaCount ? Number((item.gpaTotal / item.gpaCount).toFixed(2)) : 0,
        maxGpa: item.maxGpa === null ? 0 : Number(item.maxGpa.toFixed(2)),
        minGpa: item.minGpa === null ? 0 : Number(item.minGpa.toFixed(2)),
        gpaCount: item.gpaCount
      };
    });

  return {
    totalStudents: rows.length,
    avgGpa: validGpaCount ? Number((totalGpa / validGpaCount).toFixed(2)) : 0,
    validGpaCount: validGpaCount,
    byYear: byYear
  };
}

function readField_(row, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    if (Object.prototype.hasOwnProperty.call(row, aliases[i])) {
      return String(row[aliases[i]] || '').trim();
    }
  }

  const normalizedAliases = aliases.map(normalizeKey_);
  const keys = Object.keys(row);
  for (let j = 0; j < keys.length; j++) {
    if (normalizedAliases.indexOf(normalizeKey_(keys[j])) > -1) {
      return String(row[keys[j]] || '').trim();
    }
  }

  return '';
}

function normalizeKey_(value) {
  return String(value || '').replace(/\s+/g, '').toLowerCase();
}

function parseNumber_(value) {
  const cleaned = String(value || '').replace(',', '.').replace(/[^\d.-]/g, '');
  return cleaned ? Number(cleaned) : NaN;
}

function isAuthenticated_(params) {
  return params.username === ACCESS_USERNAME && params.password === ACCESS_PASSWORD;
}

function outputJson_(payload, callback) {
  const json = JSON.stringify(payload);
  const body = callback ? callback + '(' + json + ');' : json;
  const mimeType = callback
    ? ContentService.MimeType.JAVASCRIPT
    : ContentService.MimeType.JSON;

  return ContentService.createTextOutput(body).setMimeType(mimeType);
}
