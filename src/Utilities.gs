const Spreadsheet = SpreadsheetApp;

function getSpreadsheet() {
  return Spreadsheet.getActive();
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function ensureSheet(name, headers) {
  const sheet = getSheet(name);
  if (headers && headers.length) {
    const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const matches = headers.every(function (header, idx) {
      return firstRow[idx] === header;
    });
    if (!matches) {
      sheet.clear({ contentsOnly: true });
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getHeaderMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.reduce(function (acc, header, index) {
    if (header) {
      acc[header] = index + 1;
    }
    return acc;
  }, {});
}

function requireColumn(map, columnName) {
  if (!map[columnName]) {
    throw new Error('Missing column: ' + columnName);
  }
  return map[columnName];
}

function findRowByValue(sheet, columnIndex, value) {
  if (!value) {
    return -1;
  }
  const data = sheet.getRange(2, columnIndex, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(value)) {
      return i + 2;
    }
  }
  return -1;
}

function readSetting(key) {
  const sheet = ensureSheet(CONFIG.sheets.settings, CONFIG.headers.settings);
  const map = getHeaderMap(sheet);
  const keyCol = requireColumn(map, 'キー');
  const valueCol = requireColumn(map, '値');
  const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][keyCol - 1] === key) {
      return data[i][valueCol - 1];
    }
  }
  return null;
}

function writeSetting(key, value, description) {
  const sheet = ensureSheet(CONFIG.sheets.settings, CONFIG.headers.settings);
  const map = getHeaderMap(sheet);
  const keyCol = requireColumn(map, 'キー');
  const valueCol = requireColumn(map, '値');
  const row = findRowByValue(sheet, keyCol, key);
  const values = [key, value, description || ''];
  if (row === -1) {
    sheet.appendRow(values);
  } else {
    sheet.getRange(row, keyCol, 1, CONFIG.headers.settings.length).setValues([values]);
  }
}

function incrementCounter(key, initialValue) {
  const currentRaw = readSetting(key);
  const current = currentRaw ? Number(currentRaw) : Number(initialValue || 0);
  const next = current + 1;
  writeSetting(key, next, 'Auto increment counter');
  return next;
}

function formatId(prefix, value, digits) {
  const padded = Utilities.formatString('%0' + digits + 'd', value);
  return prefix + padded;
}

function parseNumber(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }
  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

function parseDate(value) {
  if (!value) {
    return null;
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value;
  }
  const date = new Date(value);
  return isNaN(date.getTime()) ? null : date;
}

function formatDate(value) {
  const date = parseDate(value);
  if (!date) {
    return '';
  }
  return Utilities.formatDate(date, CONFIG.timezone, 'yyyy-MM-dd');
}

function today() {
  return Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd');
}

function truthy(value) {
  return value === true || value === 'TRUE' || value === 'true' || value === 1 || value === '1';
}

function boolToCheck(value) {
  return value ? true : false;
}

function recordAudit(eventName, detail) {
  const sheet = ensureSheet(CONFIG.sheets.audit, CONFIG.headers.audit);
  sheet.appendRow([
    Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd HH:mm:ss'),
    Session.getActiveUser().getEmail(),
    eventName,
    detail || ''
  ]);
}

function parseDelimited(input) {
  if (!input) {
    return [];
  }
  const trimmed = input.trim();
  if (!trimmed) {
    return [];
  }
  const separator = trimmed.indexOf('\t') !== -1 ? '\t' : ',';
  return trimmed
    .split('\n')
    .filter(Boolean)
    .map(function (line) {
      return line.split(separator).map(function (col) {
        return col.trim();
      });
    });
}

function getListValues(category) {
  const sheet = ensureSheet(CONFIG.sheets.lists, CONFIG.headers.lists);
  const map = getHeaderMap(sheet);
  const categoryCol = requireColumn(map, 'カテゴリ');
  const valueCol = requireColumn(map, '値');
  const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), sheet.getLastColumn()).getValues();
  return data
    .filter(function (row) {
      return row[categoryCol - 1] === category;
    })
    .map(function (row) {
      return row[valueCol - 1];
    })
    .filter(Boolean);
}

function isStatusAllowed(status) {
  return getListValues('注文ステータス').indexOf(status) !== -1;
}

function normaliseRate(rate) {
  if (typeof rate === 'number') {
    return rate;
  }
  if (typeof rate === 'string') {
    const trimmed = rate.trim().replace('%', '');
    const num = Number(trimmed);
    return isNaN(num) ? 0 : num / 100;
  }
  return 0;
}

function coalesce(value, fallback) {
  return value === null || value === undefined || value === '' ? fallback : value;
}

function getAggregateStatuses() {
  const raw = readSetting('売上集計対象ステータス');
  const base = raw || CONFIG.defaultConstants['売上集計対象ステータス'];
  return String(base)
    .split(',')
    .map(function (item) {
      return item.trim();
    })
    .filter(Boolean);
}

function addValidation(sheet, rangeA1, list) {
  if (!list || !list.length) {
    return;
  }
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
  sheet.getRange(rangeA1).setDataValidation(rule);
}