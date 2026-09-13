/**
 * ============================================================================
 *  SheetService.gs
 *  Creates / opens the backing Google Sheet that stores all reconciliation
 *  data, and provides low-level helpers for reading & writing rows.
 * ============================================================================
 */

var CONFIG_SHEET_NAME = 'Config';
var LOG_SHEET_NAME = 'Reconciliation_Log';

// "Type" is appended at the END of the existing columns (not inserted in the
// middle) so pre-existing Config rows/sheets never shift and stay valid.
var CONFIG_HEADERS = ['Channel Key', 'Display Name', 'Tolerance Days', 'Amount Tolerance', 'Active', 'Created At', 'Type'];
var CONFIG_COLS = { ChannelKey: 1, DisplayName: 2, ToleranceDays: 3, AmountTolerance: 4, Active: 5, CreatedAt: 6, Type: 7 };

var CHANNEL_TYPE_BANK = 'BANK';
var CHANNEL_TYPE_BO = 'BO';
var CHANNEL_TYPE_BOTH = 'BOTH';

var LOG_HEADERS = ['Timestamp', 'Channel', 'Total Bank', 'Total Back-Office', 'Matched', 'Unmatched Bank', 'Unmatched Back-Office', 'Match Rate', 'Run By'];

var BANK_HEADERS = ['Row Id', 'Date', 'Reference', 'Description', 'Deposit', 'Withdrawal', 'Amount', 'Status', 'Match Id', 'Match Type', 'Resolution Note', 'Imported At', 'Source Hash'];
var BANK_COLS = { RowId: 1, Date: 2, Reference: 3, Description: 4, Deposit: 5, Withdrawal: 6, Amount: 7, Status: 8, MatchId: 9, MatchType: 10, ResolutionNote: 11, ImportedAt: 12, SourceHash: 13 };

var BO_HEADERS = ['Row Id', 'Date', 'Patron / Account Id', 'Txn Type', 'Reference', 'Secondary Reference', 'Amount', 'Status', 'Match Id', 'Match Type', 'Resolution Note', 'Imported At', 'Source Hash'];
var BO_COLS = { RowId: 1, Date: 2, PatronId: 3, TxnType: 4, Reference: 5, SecondaryReference: 6, Amount: 7, Status: 8, MatchId: 9, MatchType: 10, ResolutionNote: 11, ImportedAt: 12, SourceHash: 13 };

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
    // Migration: older Config sheets predate the "Type" column. Add the
    // header in place if it's missing, without touching any existing data -
    // existing rows simply read back as blank Type (treated as "BOTH").
    var lastCol = sheet.getLastColumn();
    var typeColIndex = CONFIG_COLS.Type;
    if (lastCol < typeColIndex || String(sheet.getRange(1, typeColIndex).getValue()).trim() !== 'Type') {
      sheet.getRange(1, typeColIndex).setValue('Type').setFontWeight('bold').setBackground('#1a3c6e').setFontColor('#ffffff');
    }
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
 * Creates the BANK and/or BO sheet(s) for a channel, based on its Type -
 * only the side(s) that type actually needs are created.
 *   - forceType, if given, is used directly (e.g. while creating a brand new
 *     channel, before its Config row can be looked up).
 *   - otherwise the channel's stored Type is looked up via getChannel();
 *     an unknown/missing channel defaults to BOTH for backward compatibility
 *     with every channel created before the Type column existed.
 * Returns { bankSheet, boSheet } - a side's key is omitted if that side
 * isn't part of the channel's type.
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
