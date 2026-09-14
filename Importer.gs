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
 *
 *  Only Date, Reference and Amount are "classified" - required, and used by
 *  the matching engine. Everything else in the source file is opt-in: the
 *  user picks which additional columns to retain via `mapping.extraCols`,
 *  and each one is kept under its ORIGINAL header text as an extra column on
 *  the destination sheet (informational only, never used for matching).
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
 * mapping (both sides):
 *   { dateCol, refCol,
 *     amountMode: 'single'|'split', amountCol, depositCol, withdrawalCol,
 *     extraCols: [sourceHeaderName, ...] }
 *
 * `amountMode`/`depositCol`/`withdrawalCol` only apply when computing the
 * classified Amount for BANK rows (positive = deposit, negative =
 * withdrawal); for BO rows amountCol is used directly. `extraCols` is the
 * list of additional source columns (by header name, exactly as they appear
 * in the uploaded file) the user chose to retain - kept under their original
 * header text, informational only.
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

  // Extra (opt-in, non-classified) columns the user chose to retain, kept
  // under their original header text.
  var extraCols = (mapping.extraCols || []).filter(function (h) { return headers.indexOf(h) !== -1; });
  var extraColMap = ensureExtraColumns(sheet, colsMap, extraCols); // header -> 1-based sheet column
  var fixedCount = Object.keys(colsMap).length;
  var totalCols = Math.max(sheet.getLastColumn(), fixedCount);

  var newRows = [];
  var skipped = 0;
  var now = new Date();

  var dateIdx = idx(mapping.dateCol);
  var refIdx = idx(mapping.refCol);
  var amountIdx = idx(mapping.amountCol);
  var depositIdx = idx(mapping.depositCol);
  var withdrawalIdx = idx(mapping.withdrawalCol);

  dataRows.forEach(function (r) {
    if (r.join('') === '') return;
    var date = dateIdx >= 0 ? parseDateValue(r[dateIdx]) : now;
    var ref = refIdx >= 0 ? String(r[refIdx]).trim() : '';
    var amount;

    if (side === SIDE_BANK && mapping.amountMode === 'split') {
      var deposit = depositIdx >= 0 ? parseAmount(r[depositIdx]) : 0;
      var withdrawal = withdrawalIdx >= 0 ? parseAmount(r[withdrawalIdx]) : 0;
      amount = deposit - withdrawal;
    } else {
      amount = amountIdx >= 0 ? parseAmount(r[amountIdx]) : 0;
    }

    var hash = computeSourceHash([channelKey, side, date.getTime(), ref, amount.toFixed(2)]);
    if (seenHashes[hash]) { skipped++; return; }
    seenHashes[hash] = true;

    var rowId = Utilities.getUuid();
    // Fixed columns first (Row Id, Date, Reference, Amount, Status, Match Id,
    // Match Channel, Match Type, Resolution Note, Imported At, Source Hash),
    // then extra columns in sheet-column order.
    var row = new Array(totalCols).fill('');
    row[colsMap.RowId - 1] = rowId;
    row[colsMap.Date - 1] = date;
    row[colsMap.Reference - 1] = ref;
    row[colsMap.Amount - 1] = amount;
    row[colsMap.Status - 1] = STATUS_UNMATCHED;
    row[colsMap.ImportedAt - 1] = now;
    row[colsMap.SourceHash - 1] = hash;

    extraCols.forEach(function (header) {
      var colIndex = extraColMap[header];
      if (!colIndex) return;
      var srcIdx = headers.indexOf(header);
      if (srcIdx === -1) return;
      row[colIndex - 1] = r[srcIdx];
    });

    newRows.push(row);
  });

  if (newRows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, totalCols).setValues(newRows);
  }

  return {
    imported: newRows.length,
    skippedDuplicates: skipped,
    totalRowsInFile: dataRows.length,
    retainedColumns: extraCols
  };
}
