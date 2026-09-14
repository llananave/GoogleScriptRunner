/**
 * ============================================================================
 *  ReconciliationEngine.gs
 *  Matches BANK rows against BO (back-office) rows.
 *
 *  A channel's Bank side reconciles against whichever Back-Office-capable
 *  channel its "Match Target" (Config.gs: resolveMatchTarget) points to -
 *  which may be itself (the classic paired "Both" behavior) or a different
 *  channel entirely, letting several Bank channels share one Back-Office
 *  channel or a Bank channel be repointed without re-importing anything.
 *
 *  Pass 1 - Reference match: bank Reference == BO Reference (case-insensitive,
 *           trimmed) AND amounts agree within tolerance. Highest-confidence
 *           match.
 *  Pass 2 - Amount + Date match: for everything still unmatched, pair rows
 *           whose amounts agree within tolerance and whose dates fall within
 *           the channel's tolerance-day window, picking the closest date.
 *  Anything left over is surfaced as an exception for manual review.
 *
 *  Only the classified fields (Date, Reference, Amount) are ever used for
 *  matching. Any "extra" columns the user chose to retain at import time are
 *  informational only.
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
 * Looks up the Bank-capable channel + Back-Office-capable channel that
 * should be reconciled together, starting from either side's channel key.
 * Throws a descriptive error if the source channel has no Bank side or no
 * match target configured.
 */
function resolveReconciliationPair(bankChannelKey) {
  var channel = getChannel(bankChannelKey);
  if (!channel) throw new Error('Unknown channel: ' + bankChannelKey);
  if (channel.type === CHANNEL_TYPE_BO) {
    throw new Error(
      '"' + channel.displayName + '" is a Back-Office-only channel. Run reconciliation from the Bank ' +
      'channel that is set (in Settings) to match against it instead.'
    );
  }
  var boChannel = resolveMatchTarget(channel);
  if (!boChannel) {
    throw new Error(
      '"' + channel.displayName + '" has no Back-Office match target configured yet. Go to Settings and ' +
      'choose which Back-Office channel it should reconcile against.'
    );
  }
  return { bankChannel: channel, boChannel: boChannel };
}

/**
 * Runs reconciliation for one Bank-capable channel against its configured
 * Back-Office match target, and writes results back to the sheets. Returns
 * a summary object.
 */
function runReconciliation(bankChannelKey) {
  var pair = resolveReconciliationPair(bankChannelKey);
  var bankChannel = pair.bankChannel;
  var boChannel = pair.boChannel;

  var bankSheet = ensureChannelSheets(bankChannel.key).bankSheet;
  var boSheet = ensureChannelSheets(boChannel.key).boSheet;

  var bankRows = readSheetAsObjects(bankSheet, BANK_COLS).filter(function (r) { return r.Status === STATUS_UNMATCHED || !r.Status; });
  var boRows = readSheetAsObjects(boSheet, BO_COLS).filter(function (r) { return r.Status === STATUS_UNMATCHED || !r.Status; });

  var totalBankBefore = readSheetAsObjects(bankSheet, BANK_COLS).length;
  var totalBoBefore = readSheetAsObjects(boSheet, BO_COLS).length;

  var boAvailable = boRows.slice(); // mutable working list
  var matchedCount = 0;

  var bankWrites = []; // {row, values: {Status, MatchId, MatchType, MatchChannel}}
  var boWrites = [];

  var tolerance = Number(bankChannel.amountTolerance) || 0.01;
  var toleranceDays = Number(bankChannel.toleranceDays) || 1;

  // ---- Pass 1: Reference match -------------------------------------------
  var refIndex = {}; // normalizedRef -> [boRow,...]
  boAvailable.forEach(function (bo) {
    var norm = normalizeRef(bo.Reference);
    if (!norm) return;
    if (!refIndex[norm]) refIndex[norm] = [];
    refIndex[norm].push(bo);
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
    bankWrites.push({ row: bank._row, Status: STATUS_MATCHED, MatchId: matchId, MatchType: 'Reference Match', MatchChannel: boChannel.key });
    boWrites.push({ row: best._row, Status: STATUS_MATCHED, MatchId: matchId, MatchType: 'Reference Match', MatchChannel: bankChannel.key });
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
      bankWrites.push({ row: bank._row, Status: STATUS_MATCHED, MatchId: matchId, MatchType: 'Amount + Date Match', MatchChannel: boChannel.key });
      boWrites.push({ row: bestMatch._row, Status: STATUS_MATCHED, MatchId: matchId, MatchType: 'Amount + Date Match', MatchChannel: bankChannel.key });
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

  var channelLabel = bankChannel.key === boChannel.key ? bankChannel.displayName : (bankChannel.displayName + ' \u2192 ' + boChannel.displayName);

  var ss = getDataSpreadsheet();
  var logSheet = ensureLogSheet(ss);
  logSheet.appendRow([
    new Date(), channelLabel, totalBankBefore, totalBoBefore,
    matchedCount, unmatchedBank, unmatchedBo, matchRate,
    Session.getActiveUser().getEmail() || 'unknown'
  ]);

  return {
    channel: channelLabel,
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
    if (w.hasOwnProperty('MatchChannel')) sheet.getRange(w.row, colsMap.MatchChannel).setValue(w.MatchChannel);
  });
}

/**
 * Returns unmatched (and ignored, if includeIgnored) rows for both sides of
 * a channel, for the Exceptions tab. `channelKey` can be any channel - a
 * Bank/Both channel shows its own Bank rows plus its resolved match target's
 * BO rows; a Back-Office-only channel shows just its own BO rows (which may
 * have been matched against several different Bank channels).
 */
function getExceptions(channelKey, includeIgnored) {
  var channel = getChannel(channelKey);
  if (!channel) throw new Error('Unknown channel: ' + channelKey);

  function filterFn(r) {
    if (r.Status === STATUS_UNMATCHED || !r.Status) return true;
    if (includeIgnored && r.Status === STATUS_IGNORED) return true;
    return false;
  }

  var bankRows = [], boRows = [], bankExtraHeaders = [], boExtraHeaders = [], boChannel = null;

  if (channel.type === CHANNEL_TYPE_BANK || channel.type === CHANNEL_TYPE_BOTH) {
    var bankSheet = ensureChannelSheets(channel.key).bankSheet;
    bankExtraHeaders = getExtraColumnDefs(bankSheet, BANK_COLS).map(function (d) { return d.header; });
    bankRows = readSheetRowsWithExtras(bankSheet, BANK_COLS).filter(filterFn).map(formatRowForClient);
    boChannel = resolveMatchTarget(channel);
  }

  if (channel.type === CHANNEL_TYPE_BO) {
    boChannel = channel;
  }

  if (boChannel) {
    var boSheet = ensureChannelSheets(boChannel.key).boSheet;
    boExtraHeaders = getExtraColumnDefs(boSheet, BO_COLS).map(function (d) { return d.header; });
    boRows = readSheetRowsWithExtras(boSheet, BO_COLS).filter(filterFn).map(formatRowForClient);
  }

  return {
    bank: bankRows,
    bo: boRows,
    channelType: channel.type,
    bankExtraHeaders: bankExtraHeaders,
    boExtraHeaders: boExtraHeaders,
    boChannelName: boChannel ? boChannel.displayName : null
  };
}

/**
 * Formats one row (BANK or BO, same fixed schema) for the client, including
 * an `extra` map of any additional columns the user chose to retain (see
 * readSheetRowsWithExtras).
 */
function formatRowForClient(r) {
  return {
    rowId: r.RowId,
    date: r.Date instanceof Date ? r.Date.toISOString() : String(r.Date),
    reference: r.Reference,
    amount: r.Amount,
    status: r.Status,
    matchType: r.MatchType,
    resolutionNote: r.ResolutionNote,
    extra: r.extra || {},
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
 * `channelKey` is the Bank-capable channel; the BO row is looked up on its
 * resolved match target.
 */
function manualMatch(channelKey, bankRowId, boRowId) {
  var pair = resolveReconciliationPair(channelKey);
  var bankSheet = ensureChannelSheets(pair.bankChannel.key).bankSheet;
  var boSheet = ensureChannelSheets(pair.boChannel.key).boSheet;

  var bankRow = findRowById(bankSheet, BANK_COLS, bankRowId);
  var boRow = findRowById(boSheet, BO_COLS, boRowId);
  if (!bankRow) throw new Error('Bank row not found.');
  if (!boRow) throw new Error('Back-office row not found.');
  if (bankRow.Status === STATUS_MATCHED) throw new Error('That bank row is already matched.');
  if (boRow.Status === STATUS_MATCHED) throw new Error('That back-office row is already matched.');

  var matchId = Utilities.getUuid();
  applyStatusWrites(bankSheet, BANK_COLS, [{ row: bankRow._row, Status: STATUS_MATCHED, MatchId: matchId, MatchType: 'Manual Match', MatchChannel: pair.boChannel.key }]);
  applyStatusWrites(boSheet, BO_COLS, [{ row: boRow._row, Status: STATUS_MATCHED, MatchId: matchId, MatchType: 'Manual Match', MatchChannel: pair.bankChannel.key }]);
  return { ok: true, matchId: matchId };
}

/**
 * Resolves the sheet for one side of a channel as seen from the Exceptions
 * tab: BANK always means this channel's own Bank sheet; BO means this
 * channel's own BO sheet if it's BO-capable itself, otherwise its resolved
 * match target's BO sheet.
 */
function getSheetForChannelSide(channelKey, side) {
  var channel = getChannel(channelKey);
  if (!channel) throw new Error('Unknown channel: ' + channelKey);
  if (side === SIDE_BANK) {
    if (channel.type === CHANNEL_TYPE_BO) {
      throw new Error('"' + channel.displayName + '" is Back-Office-only, so it has no Bank sheet.');
    }
    return ensureChannelSheets(channel.key).bankSheet;
  }
  if (channel.type === CHANNEL_TYPE_BO || channel.type === CHANNEL_TYPE_BOTH) {
    return ensureChannelSheets(channel.key).boSheet;
  }
  var boChannel = resolveMatchTarget(channel);
  if (!boChannel) throw new Error('"' + channel.displayName + '" has no Back-Office match target configured.');
  return ensureChannelSheets(boChannel.key).boSheet;
}

/**
 * Marks a single exception row as Ignored (e.g. known timing difference,
 * bank fee, or a transaction that will never have a counterpart), with a
 * note explaining why.
 */
function ignoreException(channelKey, side, rowId, note) {
  var sheet = getSheetForChannelSide(channelKey, side);
  var colsMap = side === SIDE_BANK ? BANK_COLS : BO_COLS;
  var row = findRowById(sheet, colsMap, rowId);
  if (!row) throw new Error('Row not found.');
  sheet.getRange(row._row, colsMap.Status).setValue(STATUS_IGNORED);
  sheet.getRange(row._row, colsMap.ResolutionNote).setValue(note || '');
  return { ok: true };
}

/**
 * Finds a row by Match Id across every Bank sheet in the data spreadsheet.
 * Needed because a Back-Office channel may be the match target of several
 * different Bank channels, so a BO row's counterpart isn't necessarily in
 * "the same-named" Bank sheet - its Match Channel column (recorded at match
 * time) tells us which Bank channel to look in.
 */
function findBankRowByMatchId(matchChannelKey, matchId) {
  if (!matchChannelKey) return null;
  var channel = getChannel(matchChannelKey);
  if (!channel || (channel.type !== CHANNEL_TYPE_BANK && channel.type !== CHANNEL_TYPE_BOTH)) return null;
  var sheet = ensureChannelSheets(channel.key).bankSheet;
  var rows = readSheetAsObjects(sheet, BANK_COLS);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].MatchId === matchId) return { sheet: sheet, row: rows[i] };
  }
  return null;
}

function findBoRowByMatchId(matchChannelKey, matchId) {
  if (!matchChannelKey) return null;
  var channel = getChannel(matchChannelKey);
  if (!channel || (channel.type !== CHANNEL_TYPE_BO && channel.type !== CHANNEL_TYPE_BOTH)) return null;
  var sheet = ensureChannelSheets(channel.key).boSheet;
  var rows = readSheetAsObjects(sheet, BO_COLS);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].MatchId === matchId) return { sheet: sheet, row: rows[i] };
  }
  return null;
}

/**
 * Reopens a previously matched or ignored row back to Unmatched, breaking
 * its match pair if it had one (the counterpart is found via its recorded
 * Match Channel, since a BO row's counterpart Bank row may live on a
 * different channel's sheet than the one being viewed).
 */
function reopenRow(channelKey, side, rowId) {
  var sheet = getSheetForChannelSide(channelKey, side);
  var colsMap = side === SIDE_BANK ? BANK_COLS : BO_COLS;
  var row = findRowById(sheet, colsMap, rowId);
  if (!row) throw new Error('Row not found.');

  var matchId = row.MatchId;
  if (matchId) {
    var counterpart = side === SIDE_BANK
      ? findBoRowByMatchId(row.MatchChannel, matchId)
      : findBankRowByMatchId(row.MatchChannel, matchId);
    if (counterpart) {
      var otherColsMap = side === SIDE_BANK ? BO_COLS : BANK_COLS;
      counterpart.sheet.getRange(counterpart.row._row, otherColsMap.Status).setValue(STATUS_UNMATCHED);
      counterpart.sheet.getRange(counterpart.row._row, otherColsMap.MatchId).setValue('');
      counterpart.sheet.getRange(counterpart.row._row, otherColsMap.MatchType).setValue('');
      counterpart.sheet.getRange(counterpart.row._row, otherColsMap.MatchChannel).setValue('');
    }
  }

  sheet.getRange(row._row, colsMap.Status).setValue(STATUS_UNMATCHED);
  sheet.getRange(row._row, colsMap.MatchId).setValue('');
  sheet.getRange(row._row, colsMap.MatchType).setValue('');
  sheet.getRange(row._row, colsMap.MatchChannel).setValue('');
  sheet.getRange(row._row, colsMap.ResolutionNote).setValue('');
  return { ok: true };
}
