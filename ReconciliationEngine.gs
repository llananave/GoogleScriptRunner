/**
 * ============================================================================
 *  ReconciliationEngine.gs
 *  Matches BANK rows against BO (back-office) rows for a channel.
 *
 *  Pass 1 - Reference match: bank Reference == BO Reference or Secondary
 *           Reference (case-insensitive, trimmed) AND amounts agree within
 *           tolerance. Highest-confidence match.
 *  Pass 2 - Amount + Date match: for everything still unmatched, pair rows
 *           whose amounts agree within tolerance and whose dates fall within
 *           the channel's tolerance-day window, picking the closest date.
 *  Anything left over is surfaced as an exception for manual review.
 * ============================================================================
 */

function msPerDay() { return 24 * 60 * 60 * 1000; }

function amountsMatch(a, b, tolerance) {
  return Math.abs(a - b) <= (tolerance + 0.0000001);
}

function normalizeRef(ref) {
  return String(ref || '').trim().toLowerCase();
}

function dateOf(row) {
  var d = row.Date;
  return d instanceof Date ? d : new Date(d);
}

/**
 * Runs reconciliation for one channel and writes results back to the sheets.
 * Returns a summary object.
 */
function runReconciliation(channelKey) {
  var channel = getChannel(channelKey);
  if (!channel) throw new Error('Unknown channel: ' + channelKey);

  var sheets = ensureChannelSheets(channelKey);
  var bankSheet = sheets.bankSheet;
  var boSheet = sheets.boSheet;

  var bankRows = readSheetAsObjects(bankSheet, BANK_COLS).filter(function (r) { return r.Status === STATUS_UNMATCHED || !r.Status; });
  var boRows = readSheetAsObjects(boSheet, BO_COLS).filter(function (r) { return r.Status === STATUS_UNMATCHED || !r.Status; });

  var totalBankBefore = readSheetAsObjects(bankSheet, BANK_COLS).length;
  var totalBoBefore = readSheetAsObjects(boSheet, BO_COLS).length;

  var boAvailable = boRows.slice(); // mutable working list
  var matchedCount = 0;

  var bankWrites = []; // {row, values: {Status, MatchId, MatchType}}
  var boWrites = [];

  var tolerance = Number(channel.amountTolerance) || 0.01;
  var toleranceDays = Number(channel.toleranceDays) || 1;

  // ---- Pass 1: Reference match -------------------------------------------
  var refIndex = {}; // normalizedRef -> [boRow,...]
  boAvailable.forEach(function (bo) {
    [bo.Reference, bo.SecondaryReference].forEach(function (ref) {
      var norm = normalizeRef(ref);
      if (!norm) return;
      if (!refIndex[norm]) refIndex[norm] = [];
      refIndex[norm].push(bo);
    });
  });

  var boUsed = {}; // _row -> true

  bankRows.forEach(function (bank) {
    var norm = normalizeRef(bank.Reference);
    if (!norm || !refIndex[norm]) return;
    var candidates = refIndex[norm].filter(function (bo) { return !boUsed[bo._row]; });
    if (candidates.length === 0) return;

    // Prefer a candidate whose amount also matches; else take the first.
    var best = null;
    for (var i = 0; i < candidates.length; i++) {
      if (amountsMatch(candidates[i].Amount, bank.Amount, tolerance)) { best = candidates[i]; break; }
    }
    if (!best) return; // reference matched but amount disagrees - leave for review, don't force a bad match

    var matchId = Utilities.getUuid();
    bankWrites.push({ row: bank._row, Status: STATUS_MATCHED, MatchId: matchId, MatchType: 'Reference Match' });
    boWrites.push({ row: best._row, Status: STATUS_MATCHED, MatchId: matchId, MatchType: 'Reference Match' });
    boUsed[best._row] = true;
    bank._matched = true;
    matchedCount++;
  });

  // ---- Pass 2: Amount + Date match ---------------------------------------
  var stillOpenBank = bankRows.filter(function (b) { return !b._matched; });
  var stillOpenBo = boAvailable.filter(function (bo) { return !boUsed[bo._row]; });

  stillOpenBank.forEach(function (bank) {
    var bankDate = dateOf(bank);
    var bestMatch = null;
    var bestDiff = Infinity;

    for (var i = 0; i < stillOpenBo.length; i++) {
      var bo = stillOpenBo[i];
      if (boUsed[bo._row]) continue;
      if (!amountsMatch(bo.Amount, bank.Amount, tolerance)) continue;
      var diffDays = Math.abs(dateOf(bo).getTime() - bankDate.getTime()) / msPerDay();
      if (diffDays <= toleranceDays && diffDays < bestDiff) {
        bestDiff = diffDays;
        bestMatch = bo;
      }
    }

    if (bestMatch) {
      var matchId = Utilities.getUuid();
      bankWrites.push({ row: bank._row, Status: STATUS_MATCHED, MatchId: matchId, MatchType: 'Amount + Date Match' });
      boWrites.push({ row: bestMatch._row, Status: STATUS_MATCHED, MatchId: matchId, MatchType: 'Amount + Date Match' });
      boUsed[bestMatch._row] = true;
      bank._matched = true;
      matchedCount++;
    }
  });

  // ---- Persist writes ------------------------------------------------------
  applyStatusWrites(bankSheet, BANK_COLS, bankWrites);
  applyStatusWrites(boSheet, BO_COLS, boWrites);

  var unmatchedBank = bankRows.filter(function (b) { return !b._matched; }).length;
  var unmatchedBo = stillOpenBo.filter(function (bo) { return !boUsed[bo._row]; }).length;
  var matchRate = (totalBankBefore + totalBoBefore) === 0 ? 0 :
    (matchedCount * 2) / (totalBankBefore + totalBoBefore);

  var ss = getDataSpreadsheet();
  var logSheet = ensureLogSheet(ss);
  logSheet.appendRow([
    new Date(), channel.displayName, totalBankBefore, totalBoBefore,
    matchedCount, unmatchedBank, unmatchedBo, matchRate,
    Session.getActiveUser().getEmail() || 'unknown'
  ]);

  return {
    channel: channel.displayName,
    totalBank: totalBankBefore,
    totalBo: totalBoBefore,
    newlyMatched: matchedCount,
    unmatchedBank: unmatchedBank,
    unmatchedBo: unmatchedBo,
    matchRate: matchRate
  };
}

function applyStatusWrites(sheet, colsMap, writes) {
  writes.forEach(function (w) {
    sheet.getRange(w.row, colsMap.Status).setValue(w.Status);
    sheet.getRange(w.row, colsMap.MatchId).setValue(w.MatchId);
    sheet.getRange(w.row, colsMap.MatchType).setValue(w.MatchType);
  });
}

/**
 * Returns unmatched (and ignored, if includeIgnored) rows for both sides of
 * a channel, for the Exceptions tab.
 */
function getExceptions(channelKey, includeIgnored) {
  var sheets = ensureChannelSheets(channelKey);
  var bankRows = readSheetAsObjects(sheets.bankSheet, BANK_COLS);
  var boRows = readSheetAsObjects(sheets.boSheet, BO_COLS);

  function filterFn(r) {
    if (r.Status === STATUS_UNMATCHED || !r.Status) return true;
    if (includeIgnored && r.Status === STATUS_IGNORED) return true;
    return false;
  }

  return {
    bank: bankRows.filter(filterFn).map(formatBankRowForClient),
    bo: boRows.filter(filterFn).map(formatBoRowForClient)
  };
}

function formatBankRowForClient(r) {
  return {
    rowId: r.RowId,
    date: r.Date instanceof Date ? r.Date.toISOString() : String(r.Date),
    reference: r.Reference,
    description: r.Description,
    deposit: r.Deposit,
    withdrawal: r.Withdrawal,
    amount: r.Amount,
    status: r.Status,
    matchType: r.MatchType,
    resolutionNote: r.ResolutionNote,
    _row: r._row
  };
}

function formatBoRowForClient(r) {
  return {
    rowId: r.RowId,
    date: r.Date instanceof Date ? r.Date.toISOString() : String(r.Date),
    patronId: r.PatronId,
    txnType: r.TxnType,
    reference: r.Reference,
    secondaryReference: r.SecondaryReference,
    amount: r.Amount,
    status: r.Status,
    matchType: r.MatchType,
    resolutionNote: r.ResolutionNote,
    _row: r._row
  };
}

function findRowById(sheet, colsMap, rowId) {
  var rows = readSheetAsObjects(sheet, colsMap);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].RowId === rowId) return rows[i];
  }
  return null;
}

/**
 * Manually pairs one bank row with one BO row (user-confirmed match).
 */
function manualMatch(channelKey, bankRowId, boRowId) {
  var sheets = ensureChannelSheets(channelKey);
  var bankRow = findRowById(sheets.bankSheet, BANK_COLS, bankRowId);
  var boRow = findRowById(sheets.boSheet, BO_COLS, boRowId);
  if (!bankRow) throw new Error('Bank row not found.');
  if (!boRow) throw new Error('Back-office row not found.');
  if (bankRow.Status === STATUS_MATCHED) throw new Error('That bank row is already matched.');
  if (boRow.Status === STATUS_MATCHED) throw new Error('That back-office row is already matched.');

  var matchId = Utilities.getUuid();
  applyStatusWrites(sheets.bankSheet, BANK_COLS, [{ row: bankRow._row, Status: STATUS_MATCHED, MatchId: matchId, MatchType: 'Manual Match' }]);
  applyStatusWrites(sheets.boSheet, BO_COLS, [{ row: boRow._row, Status: STATUS_MATCHED, MatchId: matchId, MatchType: 'Manual Match' }]);
  return { ok: true, matchId: matchId };
}

/**
 * Marks a single exception row as Ignored (e.g. known timing difference,
 * bank fee, or a transaction that will never have a counterpart), with a
 * note explaining why.
 */
function ignoreException(channelKey, side, rowId, note) {
  var sheet = getSheetForSide(channelKey, side);
  var colsMap = side === SIDE_BANK ? BANK_COLS : BO_COLS;
  var row = findRowById(sheet, colsMap, rowId);
  if (!row) throw new Error('Row not found.');
  sheet.getRange(row._row, colsMap.Status).setValue(STATUS_IGNORED);
  sheet.getRange(row._row, colsMap.ResolutionNote).setValue(note || '');
  return { ok: true };
}

/**
 * Reopens a previously matched or ignored row back to Unmatched, breaking
 * its match pair if it had one.
 */
function reopenRow(channelKey, side, rowId) {
  var sheet = getSheetForSide(channelKey, side);
  var colsMap = side === SIDE_BANK ? BANK_COLS : BO_COLS;
  var row = findRowById(sheet, colsMap, rowId);
  if (!row) throw new Error('Row not found.');

  var matchId = row.MatchId;
  if (matchId) {
    // Also reopen the partner row on the other side.
    var otherSide = side === SIDE_BANK ? SIDE_BO : SIDE_BANK;
    var otherSheet = getSheetForSide(channelKey, otherSide);
    var otherColsMap = otherSide === SIDE_BANK ? BANK_COLS : BO_COLS;
    var otherRows = readSheetAsObjects(otherSheet, otherColsMap);
    otherRows.forEach(function (r) {
      if (r.MatchId === matchId) {
        otherSheet.getRange(r._row, otherColsMap.Status).setValue(STATUS_UNMATCHED);
        otherSheet.getRange(r._row, otherColsMap.MatchId).setValue('');
        otherSheet.getRange(r._row, otherColsMap.MatchType).setValue('');
      }
    });
  }

  sheet.getRange(row._row, colsMap.Status).setValue(STATUS_UNMATCHED);
  sheet.getRange(row._row, colsMap.MatchId).setValue('');
  sheet.getRange(row._row, colsMap.MatchType).setValue('');
  sheet.getRange(row._row, colsMap.ResolutionNote).setValue('');
  return { ok: true };
}
