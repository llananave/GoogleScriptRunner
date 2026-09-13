/**
 * ============================================================================
 *  Importer.gs
 *  Turns pasted CSV text (any column layout) into the standardized BANK / BO
 *  row format, using a column mapping supplied by the user in the UI.
 *
 *  Different banks and payment platforms export wildly different column
 *  sets (see the original inspiration workbook: UB Online, GCash, Maya and
 *  DragonPay exports all differ) so rather than hard-coding a layout per
 *  channel, the user maps columns once per channel and it is reusable.
 * ============================================================================
 */

/**
 * Parses CSV text and returns the header row + a small preview of data rows
 * so the client can render a column-mapping form.
 */
function previewCsv(csvText) {
  var rows = Utilities.parseCsv(csvText);
  if (!rows || rows.length === 0) throw new Error('No data found in the pasted text.');
  var headers = rows[0].map(function (h) { return String(h).trim(); });
  var sample = rows.slice(1, 6);
  return {
    headers: headers,
    sampleRows: sample,
    totalDataRows: rows.length - 1
  };
}

function computeSourceHash(parts) {
  var raw = parts.join('|');
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8);
  return digest.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function parseAmount(value) {
  if (value === null || value === undefined || value === '') return 0;
  var n = parseFloat(String(value).replace(/,/g, '').replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function parseDateValue(value) {
  if (value instanceof Date) return value;
  var d = new Date(value);
  if (!isNaN(d.getTime())) return d;
  return new Date(); // fall back rather than blow up the whole import
}

function existingHashSet(sheet, colsMap) {
  var lastRow = sheet.getLastRow();
  var set = {};
  if (lastRow < 2) return set;
  var col = colsMap.SourceHash;
  var values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0]) set[values[i][0]] = true;
  }
  return set;
}

/**
 * Imports pasted CSV text for a channel + side (BANK or BO) using the given
 * column mapping. Skips rows already imported (matched by content hash).
 *
 * mapping (BANK side):
 *   { dateCol, refCol, descCol, amountMode: 'single'|'split',
 *     amountCol, depositCol, withdrawalCol, signConvention: 'positiveDeposit'|'positiveWithdrawal' }
 *
 * mapping (BO side):
 *   { dateCol, patronCol, typeCol, refCol, secondaryRefCol, amountCol }
 */
function importData(channelKey, side, csvText, mapping) {
  var channel = getChannel(channelKey);
  if (!channel) throw new Error('Unknown channel: ' + channelKey);

  var rows = Utilities.parseCsv(csvText);
  if (!rows || rows.length < 2) throw new Error('No data rows found in the pasted text.');
  var headers = rows[0].map(function (h) { return String(h).trim(); });
  var dataRows = rows.slice(1);

  function idx(colName) {
    if (!colName) return -1;
    return headers.indexOf(colName);
  }

  var sheet = getSheetForSide(channelKey, side);
  var colsMap = side === SIDE_BANK ? BANK_COLS : BO_COLS;
  var seenHashes = existingHashSet(sheet, colsMap);

  var newRows = [];
  var skipped = 0;
  var now = new Date();

  if (side === SIDE_BANK) {
    var dateIdx = idx(mapping.dateCol);
    var refIdx = idx(mapping.refCol);
    var descIdx = idx(mapping.descCol);
    var amountIdx = idx(mapping.amountCol);
    var depositIdx = idx(mapping.depositCol);
    var withdrawalIdx = idx(mapping.withdrawalCol);

    dataRows.forEach(function (r) {
      if (r.join('') === '') return;
      var date = dateIdx >= 0 ? parseDateValue(r[dateIdx]) : now;
      var ref = refIdx >= 0 ? String(r[refIdx]).trim() : '';
      var desc = descIdx >= 0 ? String(r[descIdx]).trim() : '';
      var deposit = 0, withdrawal = 0;

      if (mapping.amountMode === 'split') {
        deposit = depositIdx >= 0 ? parseAmount(r[depositIdx]) : 0;
        withdrawal = withdrawalIdx >= 0 ? parseAmount(r[withdrawalIdx]) : 0;
      } else {
        var amt = amountIdx >= 0 ? parseAmount(r[amountIdx]) : 0;
        if (amt >= 0) { deposit = amt; } else { withdrawal = Math.abs(amt); }
      }

      var amount = deposit - withdrawal;
      var hash = computeSourceHash([channelKey, side, date.getTime(), ref, amount.toFixed(2), desc]);
      if (seenHashes[hash]) { skipped++; return; }
      seenHashes[hash] = true;

      var rowId = Utilities.getUuid();
      newRows.push([rowId, date, ref, desc, deposit, withdrawal, amount, STATUS_UNMATCHED, '', '', '', now, hash]);
    });
  } else {
    var dateIdx2 = idx(mapping.dateCol);
    var patronIdx = idx(mapping.patronCol);
    var typeIdx = idx(mapping.typeCol);
    var refIdx2 = idx(mapping.refCol);
    var secRefIdx = idx(mapping.secondaryRefCol);
    var amountIdx2 = idx(mapping.amountCol);

    dataRows.forEach(function (r) {
      if (r.join('') === '') return;
      var date = dateIdx2 >= 0 ? parseDateValue(r[dateIdx2]) : now;
      var patron = patronIdx >= 0 ? String(r[patronIdx]).trim() : '';
      var type = typeIdx >= 0 ? String(r[typeIdx]).trim() : '';
      var ref = refIdx2 >= 0 ? String(r[refIdx2]).trim() : '';
      var secRef = secRefIdx >= 0 ? String(r[secRefIdx]).trim() : '';
      var amount = amountIdx2 >= 0 ? parseAmount(r[amountIdx2]) : 0;

      var hash = computeSourceHash([channelKey, side, date.getTime(), ref, amount.toFixed(2), patron]);
      if (seenHashes[hash]) { skipped++; return; }
      seenHashes[hash] = true;

      var rowId = Utilities.getUuid();
      newRows.push([rowId, date, patron, type, ref, secRef, amount, STATUS_UNMATCHED, '', '', '', now, hash]);
    });
  }

  if (newRows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
  }

  return {
    imported: newRows.length,
    skippedDuplicates: skipped,
    totalRowsInFile: dataRows.length
  };
}