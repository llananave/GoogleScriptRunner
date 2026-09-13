/**
 * ============================================================================
 *  SheetService.gs
 *  Creates / opens the backing Google Sheet that stores all reconciliation
 *  data, and provides low-level helpers for reading & writing rows.
 * ============================================================================
 */

var CONFIG_SHEET_NAME = 'Config';
var LOG_SHEET_NAME = 'Reconciliation_Log';

var CONFIG_HEADERS = ['Channel Key', 'Display Name', 'Tolerance Days', 'Amount Tolerance', 'Active', 'Created At'];
var CONFIG_COLS = { ChannelKey: 1, DisplayName: 2, ToleranceDays: 3, AmountTolerance: 4, Active: 5, CreatedAt: 6 };

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
 */
function getDataSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('DATA_SS_ID');
  var ss = null;

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (err) {
      ss = null; // stored id is stale / inaccessible - fall through and recreate
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create(DATA_SPREADSHEET_NAME);
    props.setProperty('DATA_SS_ID', ss.getId());
    var defaultSheet = ss.getSheets()[0];
    defaultSheet.setName('Welcome');
    defaultSheet.getRange('A1').setValue(
      'This spreadsheet is the data store for the "' + APP_TITLE + '" web app. ' +
      'Use the web app UI to manage everything here - manual edits are possible but not required.'
    );
  }

  return ss;
}

function ensureConfigSheet(ss) {
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
    sheet.appendRow(CONFIG_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, CONFIG_HEADERS.length).setFontWeight('bold').setBackground('#1a3c6e').setFontColor('#ffffff');
    sheet.autoResizeColumns(1, CONFIG_HEADERS.length);
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
 * Creates the BANK / BO sheets for a channel if they don't already exist.
 */
function ensureChannelSheets(channelKey) {
  var ss = getDataSpreadsheet();
  var bankSheet = ss.getSheetByName(bankSheetName(channelKey));
  if (!bankSheet) {
    bankSheet = ss.insertSheet(bankSheetName(channelKey));
    bankSheet.appendRow(BANK_HEADERS);
    bankSheet.setFrozenRows(1);
    bankSheet.getRange(1, 1, 1, BANK_HEADERS.length).setFontWeight('bold').setBackground('#0b5394').setFontColor('#ffffff');
  }
  var boSheet = ss.getSheetByName(boSheetName(channelKey));
  if (!boSheet) {
    boSheet = ss.insertSheet(boSheetName(channelKey));
    boSheet.appendRow(BO_HEADERS);
    boSheet.setFrozenRows(1);
    boSheet.getRange(1, 1, 1, BO_HEADERS.length).setFontWeight('bold').setBackground('#38761d').setFontColor('#ffffff');
  }
  return { bankSheet: bankSheet, boSheet: boSheet };
}

function getSheetForSide(channelKey, side) {
  var ss = getDataSpreadsheet();
  var name = side === SIDE_BANK ? bankSheetName(channelKey) : boSheetName(channelKey);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    ensureChannelSheets(channelKey);
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
