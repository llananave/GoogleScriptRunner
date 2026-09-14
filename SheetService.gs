/**
 * ============================================================================
 *  SheetService.gs
 *  Creates / opens the backing Google Sheet that stores all reconciliation
 *  data, and provides low-level helpers for reading & writing rows.
 *
 *  SCHEMA NOTE (column customization):
 *  Every BANK / BO sheet has a small set of FIXED "classified" columns that
 *  the app relies on (Row Id, Date, Reference, Amount, Status, plus internal
 *  bookkeeping like Match Id / Match Channel / Match Type / Resolution Note /
 *  Imported At / Source Hash). Everything else is an "extra" column: any
 *  additional source column the user chose to retain at import time, kept
 *  under its ORIGINAL header text, appended after the fixed columns. Extra
 *  columns are purely informational - they display in Exceptions/Export but
 *  are never used for matching.
 * ============================================================================
 */

var CONFIG_SHEET_NAME = 'Config';
var LOG_SHEET_NAME = 'Reconciliation_Log';

// "Type" and "Match Target" are appended at the END of the existing columns
// (not inserted in the middle) so pre-existing Config rows/sheets never
// shift and stay valid.
var CONFIG_HEADERS = ['Channel Key', 'Display Name', 'Tolerance Days', 'Amount Tolerance', 'Active', 'Created At', 'Type', 'Match Target'];
var CONFIG_COLS = { ChannelKey: 1, DisplayName: 2, ToleranceDays: 3, AmountTolerance: 4, Active: 5, CreatedAt: 6, Type: 7, MatchTarget: 8 };

var CHANNEL_TYPE_BANK = 'BANK';
var CHANNEL_TYPE_BO = 'BO';
var CHANNEL_TYPE_BOTH = 'BOTH';

var LOG_HEADERS = ['Timestamp', 'Channel', 'Total Bank', 'Total Back-Office', 'Matched', 'Unmatched Bank', 'Unmatched Back-Office', 'Match Rate', 'Run By'];

// ---- Fixed / classified columns -------------------------------------------
// Date, Reference, Amount and Status are the "default classified fields"
// every BANK/BO sheet always has. Row Id / Match Id / Match Channel /
// Match Type / Resolution Note / Imported At / Source Hash are internal
// bookkeeping columns the app needs regardless of what the user chooses to
// retain from their source file.
var BANK_HEADERS = ['Row Id', 'Date', 'Reference', 'Amount', 'Status', 'Match Id', 'Match Channel', 'Match Type', 'Resolution Note', 'Imported At', 'Source Hash'];
var BANK_COLS = { RowId: 1, Date: 2, Reference: 3, Amount: 4, Status: 5, MatchId: 6, MatchChannel: 7, MatchType: 8, ResolutionNote: 9, ImportedAt: 10, SourceHash: 11 };

var BO_HEADERS = ['Row Id', 'Date', 'Reference', 'Amount', 'Status', 'Match Id', 'Match Channel', 'Match Type', 'Resolution Note', 'Imported At', 'Source Hash'];
var BO_COLS = { RowId: 1, Date: 2, Reference: 3, Amount: 4, Status: 5, MatchId: 6, MatchChannel: 7, MatchType: 8, ResolutionNote: 9, ImportedAt: 10, SourceHash: 11 };

// Legacy (pre-column-customization) layouts, kept only so old sheets can be
// migrated in place the first time they're touched - see
// migrateLegacyChannelSheet() below. Not used anywhere else.
var LEGACY_BANK_HEADERS = ['Row Id', 'Date', 'Reference', 'Description', 'Deposit', 'Withdrawal', 'Amount', 'Status', 'Match Id', 'Match Type', 'Resolution Note', 'Imported At', 'Source Hash'];
var LEGACY_BO_HEADERS = ['Row Id', 'Date', 'Patron / Account Id', 'Txn Type', 'Reference', 'Secondary Reference', 'Amount', 'Status', 'Match Id', 'Match Type', 'Resolution Note', 'Imported At', 'Source Hash'];

var STATUS_MATCHED = 'Matched';
var STATUS_UNMATCHED = 'Unmatched';
var STATUS_IGNORED = 'Ignored';

var SIDE_BANK = 'BANK';
var SIDE_BO = 'BO';

/**
 * Returns (creating if necessary) the spreadsheet that stores all app data.
 * The spreadsheet id is cached in Script Properties so every user of the
 * deployed web app shares the same data store.
 *
 * IMPORTANT: if the stored id can no longer be opened (wrong account,
 * permission revoked, sheet trashed, property lost after a redeploy/copy of
 * the project, etc.) this does NOT silently create a brand-new empty
 * spreadsheet and orphan your existing data. Instead it records what
 * happened (surfaced on the Settings > Setup Diagnostics panel) and throws,
 * so you get a clear error instead of an app that looks "empty" for no
 * apparent reason. Use reconnectDataSpreadsheet() to point the app back at
 * the correct spreadsheet if this happens.
 */
function getDataSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('DATA_SS_ID');

  if (ssId) {
    try {
      return SpreadsheetApp.openById(ssId);
    } catch (err) {
      props.setProperty('DATA_SS_ID_OPEN_ERROR', (err && err.message) ? err.message : String(err));
      props.setProperty('DATA_SS_ID_OPEN_ERROR_AT', new Date().toISOString());
      throw new Error(
        'Could not open the connected data spreadsheet (id: ' + ssId + '). It may have been trashed, ' +
        'moved, or the account running this script no longer has access to it. Your data has NOT been ' +
        'deleted or replaced. Go to Settings > Setup Diagnostics and use "Reconnect to an existing ' +
        'spreadsheet" with the correct spreadsheet URL, or restore access to the original one. ' +
        '(Underlying error: ' + ((err && err.message) ? err.message : String(err)) + ')'
      );
    }
  }

  // No id stored yet at all - this is a genuinely first-ever run, safe to create.
  var ss = SpreadsheetApp.create(DATA_SPREADSHEET_NAME);
  props.setProperty('DATA_SS_ID', ss.getId());
  props.deleteProperty('DATA_SS_ID_OPEN_ERROR');
  props.deleteProperty('DATA_SS_ID_OPEN_ERROR_AT');
  var defaultSheet = ss.getSheets()[0];
  defaultSheet.setName('Welcome');
  defaultSheet.getRange('A1').setValue(
    'This spreadsheet is the data store for the "' + APP_TITLE + '" web app. ' +
    'Use the web app UI to manage everything here - manual edits are possible but not required.'
  );
  return ss;
}

/**
 * Points the app at a different, already-existing data spreadsheet (by id or
 * full URL). Use this to recover from a lost/incorrect DATA_SS_ID script
 * property without losing data that already lives in that spreadsheet -
 * nothing here modifies the target spreadsheet's contents.
 */
function reconnectDataSpreadsheet(idOrUrl) {
  var input = String(idOrUrl || '').trim();
  if (!input) throw new Error('Please provide a spreadsheet URL or id.');

  var id = input;
  var match = input.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match) id = match[1];

  var ss;
  try {
    ss = SpreadsheetApp.openById(id);
  } catch (err) {
    throw new Error('Could not open that spreadsheet - double-check the URL/id and that this account has access to it.');
  }

  var props = PropertiesService.getScriptProperties();
  props.setProperty('DATA_SS_ID', ss.getId());
  props.deleteProperty('DATA_SS_ID_OPEN_ERROR');
  props.deleteProperty('DATA_SS_ID_OPEN_ERROR_AT');

  ensureConfigSheet(ss);
  ensureLogSheet(ss);

  return {
    spreadsheetUrl: ss.getUrl(),
    spreadsheetName: ss.getName(),
    configRowCount: Math.max(0, ensureConfigSheet(ss).getLastRow() - 1)
  };
}

function ensureConfigSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
    sheet.appendRow(CONFIG_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, CONFIG_HEADERS.length).setFontWeight('bold').setBackground('#1a3c6e').setFontColor('#ffffff');
    sheet.autoResizeColumns(1, CONFIG_HEADERS.length);
  } else {
    // Migration: older Config sheets predate the "Type" / "Match Target"
    // columns. Add any missing header in place, without touching any
    // existing data - existing rows simply read back as blank (Type treated
    // as "BOTH", Match Target treated as "not yet configured").
    [CONFIG_COLS.Type, CONFIG_COLS.MatchTarget].forEach(function (colIndex) {
      var expectedHeader = CONFIG_HEADERS[colIndex - 1];
      var lastCol = sheet.getLastColumn();
      if (lastCol < colIndex || String(sheet.getRange(1, colIndex).getValue()).trim() !== expectedHeader) {
        sheet.getRange(1, colIndex).setValue(expectedHeader).setFontWeight('bold').setBackground('#1a3c6e').setFontColor('#ffffff');
      }
    });
  }
  return sheet;
}

function ensureLogSheet(ss) {
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(LOG_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setFontWeight('bold').setBackground('#1a3c6e').setFontColor('#ffffff');
  }
  return sheet;
}

/**
 * Turns a free-form channel display name into a safe, unique-ish sheet name
 * fragment (letters, numbers, spaces and a few symbols only).
 */
function sanitizeChannelKey(name) {
  return String(name).trim().replace(/[^a-zA-Z0-9 _.-]/g, '').substring(0, 60);
}

function bankSheetName(channelKey) {
  return ('BANK - ' + channelKey).substring(0, 100);
}

function boSheetName(channelKey) {
  return ('BO - ' + channelKey).substring(0, 100);
}

/**
 * If `sheet` still has the pre-column-customization layout (fixed
 * Description/Deposit/Withdrawal or Patron/TxnType/SecondaryReference
 * columns), rewrites it in place to the new fixed schema (Row Id, Date,
 * Reference, Amount, Status, ...bookkeeping...) and moves the old classified
 * columns to the end as "extra" columns under their original headers, so no
 * historical data is lost. Safe to call repeatedly - it's a no-op once a
 * sheet is already on the new layout (whether freshly created or already
 * migrated).
 */
function migrateLegacyChannelSheet(sheet, side) {
  if (!sheet || sheet.getLastRow() < 1) return;
  var legacyHeaders = side === SIDE_BANK ? LEGACY_BANK_HEADERS : LEGACY_BO_HEADERS;
  var lastCol = sheet.getLastColumn();
  if (lastCol < legacyHeaders.length) return; // too narrow to be the legacy layout

  var currentHeaders = sheet.getRange(1, 1, 1, legacyHeaders.length).getValues()[0].map(function (h) { return String(h).trim(); });
  var isLegacy = legacyHeaders.every(function (h, i) { return currentHeaders[i] === h; });
  if (!isLegacy) return; // already new layout (or something custom) - leave it alone

  var lastRow = sheet.getLastRow();
  var totalCols = sheet.getLastColumn();
  var allValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, totalCols).getValues() : [];

  // Any columns beyond the legacy fixed set were already "extra" columns
  // added by a previous run of the new import code - preserve them verbatim.
  var preexistingExtraHeaders = [];
  if (totalCols > legacyHeaders.length) {
    preexistingExtraHeaders = sheet.getRange(1, legacyHeaders.length + 1, 1, totalCols - legacyHeaders.length).getValues()[0];
  }

  var newExtraHeaders, buildRow;
  if (side === SIDE_BANK) {
    // legacy: RowId,Date,Reference,Description,Deposit,Withdrawal,Amount,Status,MatchId,MatchType,ResolutionNote,ImportedAt,SourceHash
    newExtraHeaders = ['Description', 'Deposit', 'Withdrawal'].concat(preexistingExtraHeaders);
    buildRow = function (r) {
      var extra = [r[3], r[4], r[5]].concat(r.slice(13));
      return [r[0], r[1], r[2], r[6], r[7], r[8], '', r[9], r[10], r[11], r[12]].concat(extra);
    };
  } else {
    // legacy: RowId,Date,PatronId,TxnType,Reference,SecondaryReference,Amount,Status,MatchId,MatchType,ResolutionNote,ImportedAt,SourceHash
    newExtraHeaders = ['Patron / Account Id', 'Txn Type', 'Secondary Reference'].concat(preexistingExtraHeaders);
    buildRow = function (r) {
      var extra = [r[2], r[3], r[5]].concat(r.slice(13));
      return [r[0], r[1], r[4], r[6], r[7], r[8], '', r[9], r[10], r[11], r[12]].concat(extra);
    };
  }

  var fixedHeaders = side === SIDE_BANK ? BANK_HEADERS : BO_HEADERS;
  var newHeaderRow = fixedHeaders.concat(newExtraHeaders);
  var newDataRows = allValues.map(buildRow);

  sheet.clear();
  sheet.getRange(1, 1, 1, newHeaderRow.length).setValues([newHeaderRow]);
  sheet.setFrozenRows(1);
  var bg = side === SIDE_BANK ? '#0b5394' : '#38761d';
  sheet.getRange(1, 1, 1, fixedHeaders.length).setFontWeight('bold').setBackground(bg).setFontColor('#ffffff');
  if (newExtraHeaders.length > 0) {
    sheet.getRange(1, fixedHeaders.length + 1, 1, newExtraHeaders.length).setFontWeight('bold').setBackground('#666666').setFontColor('#ffffff');
  }
  if (newDataRows.length > 0) {
    sheet.getRange(2, 1, newDataRows.length, newHeaderRow.length).setValues(newDataRows);
  }
}

/**
 * Creates the BANK and/or BO sheet(s) for a channel, based on its Type -
 * only the side(s) that type actually needs are created.
 *   - forceType, if given, is used directly (e.g. while creating a brand new
 *     channel, before its Config row can be looked up).
 *   - otherwise the channel's stored Type is looked up via getChannel();
 *     an unknown/missing channel defaults to BOTH for backward compatibility
 *     with every channel created before the Type column existed.
 * Returns { bankSheet, boSheet } - a side's key is omitted if that side
 * isn't part of the channel's type. Existing sheets are migrated in place if
 * they still use the pre-column-customization layout (see
 * migrateLegacyChannelSheet).
 */
function ensureChannelSheets(channelKey, forceType) {
  var ss = getDataSpreadsheet();
  var type = forceType || ((getChannel(channelKey) || {}).type) || CHANNEL_TYPE_BOTH;
  type = String(type).toUpperCase();

  var wantBank = (type === CHANNEL_TYPE_BANK || type === CHANNEL_TYPE_BOTH);
  var wantBo = (type === CHANNEL_TYPE_BO || type === CHANNEL_TYPE_BOTH);
  var result = {};

  if (wantBank) {
    var bankSheet = ss.getSheetByName(bankSheetName(channelKey));
    if (!bankSheet) {
      bankSheet = ss.insertSheet(bankSheetName(channelKey));
      bankSheet.appendRow(BANK_HEADERS);
      bankSheet.setFrozenRows(1);
      bankSheet.getRange(1, 1, 1, BANK_HEADERS.length).setFontWeight('bold').setBackground('#0b5394').setFontColor('#ffffff');
    } else {
      migrateLegacyChannelSheet(bankSheet, SIDE_BANK);
    }
    result.bankSheet = bankSheet;
  }
  if (wantBo) {
    var boSheet = ss.getSheetByName(boSheetName(channelKey));
    if (!boSheet) {
      boSheet = ss.insertSheet(boSheetName(channelKey));
      boSheet.appendRow(BO_HEADERS);
      boSheet.setFrozenRows(1);
      boSheet.getRange(1, 1, 1, BO_HEADERS.length).setFontWeight('bold').setBackground('#38761d').setFontColor('#ffffff');
    } else {
      migrateLegacyChannelSheet(boSheet, SIDE_BO);
    }
    result.boSheet = boSheet;
  }
  return result;
}

function getSheetForSide(channelKey, side) {
  var channel = getChannel(channelKey);
  var type = channel ? channel.type : CHANNEL_TYPE_BOTH;
  var sideNeeded = side === SIDE_BANK ? CHANNEL_TYPE_BANK : CHANNEL_TYPE_BO;
  if (type !== CHANNEL_TYPE_BOTH && type !== sideNeeded) {
    var label = channel ? channel.displayName : channelKey;
    throw new Error(
      '"' + label + '" is set up as ' + (type === CHANNEL_TYPE_BANK ? 'Bank-only' : 'Back-Office-only') +
      ' in Settings, so it has no ' + (sideNeeded === CHANNEL_TYPE_BANK ? 'Bank' : 'Back-Office') + ' sheet.'
    );
  }
  var ss = getDataSpreadsheet();
  var name = side === SIDE_BANK ? bankSheetName(channelKey) : boSheetName(channelKey);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    ensureChannelSheets(channelKey, type);
    sheet = ss.getSheetByName(name);
  } else {
    migrateLegacyChannelSheet(sheet, side);
  }
  return sheet;
}

/**
 * Reads an entire sheet (minus header row) into an array of plain objects
 * keyed by the given COLS map. Includes the 1-based sheet row number as
 * "_row" so callers can write back to the exact row later.
 */
function readSheetAsObjects(sheet, colsMap) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var keys = Object.keys(colsMap);
  var results = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (row.join('') === '') continue; // skip fully blank rows
    var obj = { _row: i + 2 };
    for (var k = 0; k < keys.length; k++) {
      var key = keys[k];
      var colIndex = colsMap[key] - 1;
      obj[key] = row[colIndex];
    }
    results.push(obj);
  }
  return results;
}

/**
 * Returns the "extra" (non-classified) column headers of a channel sheet -
 * whatever the user chose to retain from their source file, in sheet order -
 * as [{ header, colIndex (1-based) }, ...]. Columns are considered "extra"
 * if they sit beyond the fixed classified/bookkeeping columns.
 */
function getExtraColumnDefs(sheet, fixedColsMap) {
  var fixedCount = Object.keys(fixedColsMap).length;
  var lastCol = sheet.getLastColumn();
  if (lastCol <= fixedCount) return [];
  var headers = sheet.getRange(1, fixedCount + 1, 1, lastCol - fixedCount).getValues()[0];
  var defs = [];
  for (var i = 0; i < headers.length; i++) {
    var header = String(headers[i] || '').trim();
    if (!header) continue;
    defs.push({ header: header, colIndex: fixedCount + i + 1 });
  }
  return defs;
}

/**
 * Ensures every header in `headers` exists as an extra column on `sheet`
 * (appending new columns at the end, matched case-insensitively/trimmed
 * against what's already there so re-importing with the same selection
 * reuses the same column instead of duplicating it). Returns a map of
 * header -> 1-based column index for every requested header.
 */
function ensureExtraColumns(sheet, fixedColsMap, headers) {
  var existing = getExtraColumnDefs(sheet, fixedColsMap);
  var map = {};
  existing.forEach(function (d) { map[d.header.toLowerCase()] = d.colIndex; });

  var toAppend = [];
  headers.forEach(function (h) {
    var header = String(h || '').trim();
    if (!header) return;
    var key = header.toLowerCase();
    if (map.hasOwnProperty(key)) return;
    toAppend.push(header);
  });

  if (toAppend.length > 0) {
    var startCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, startCol, 1, toAppend.length).setValues([toAppend])
      .setFontWeight('bold').setBackground('#666666').setFontColor('#ffffff');
    toAppend.forEach(function (header, i) { map[header.toLowerCase()] = startCol + i; });
  }

  var result = {};
  headers.forEach(function (h) {
    var header = String(h || '').trim();
    if (!header) return;
    result[header] = map[header.toLowerCase()];
  });
  return result;
}

/**
 * Like readSheetAsObjects, but also captures every "extra" column into an
 * `extra` map keyed by header name (see getExtraColumnDefs). Use this
 * wherever a row needs to be displayed or exported, so user-retained columns
 * come along; use plain readSheetAsObjects for matching/internal logic that
 * only ever needs the fixed classified fields.
 */
function readSheetRowsWithExtras(sheet, colsMap) {
  var extraDefs = getExtraColumnDefs(sheet, colsMap);
  var rows = readSheetAsObjects(sheet, colsMap);
  if (extraDefs.length === 0) {
    rows.forEach(function (r) { r.extra = {}; });
    return rows;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return rows;
  var minExtraCol = extraDefs.reduce(function (min, d) { return Math.min(min, d.colIndex); }, extraDefs[0].colIndex);
  var maxExtraCol = extraDefs.reduce(function (max, d) { return Math.max(max, d.colIndex); }, extraDefs[0].colIndex);
  var extraValues = sheet.getRange(2, minExtraCol, lastRow - 1, maxExtraCol - minExtraCol + 1).getValues();

  rows.forEach(function (r) {
    var extra = {};
    var valueRow = extraValues[r._row - 2]; // r._row is 1-based sheet row; extraValues is 0-based from row 2
    extraDefs.forEach(function (d) {
      extra[d.header] = valueRow ? valueRow[d.colIndex - minExtraCol] : '';
    });
    r.extra = extra;
  });
  return rows;
}

function colLetter(colIndex1Based) {
  var s = '';
  var n = colIndex1Based;
  while (n > 0) {
    var mod = (n - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    n = Math.floor((n - mod) / 26);
  }
  return s;
}
