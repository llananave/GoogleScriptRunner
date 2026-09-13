/**
 * ============================================================================
 *  Dashboard.gs
 *  Aggregates per-channel and overall totals for the Dashboard tab.
 * ============================================================================
 */

function getDashboardData() {
  var channels = getChannels();
  var ss = getDataSpreadsheet();

  var overall = { totalBank: 0, totalBo: 0, matched: 0, unmatchedBank: 0, unmatchedBo: 0, ignored: 0 };
  var perChannel = [];

  channels.forEach(function (ch) {
    var bankSheet = ss.getSheetByName(bankSheetName(ch.key));
    var boSheet = ss.getSheetByName(boSheetName(ch.key));
    var bankRows = bankSheet ? readSheetAsObjects(bankSheet, BANK_COLS) : [];
    var boRows = boSheet ? readSheetAsObjects(boSheet, BO_COLS) : [];

    var counts = { totalBank: bankRows.length, totalBo: boRows.length, matched: 0, unmatchedBank: 0, unmatchedBo: 0, ignored: 0 };

    bankRows.forEach(function (r) {
      if (r.Status === STATUS_MATCHED) counts.matched++;
      else if (r.Status === STATUS_IGNORED) counts.ignored++;
      else counts.unmatchedBank++;
    });
    boRows.forEach(function (r) {
      if (r.Status === STATUS_IGNORED) counts.ignored++;
      else if (r.Status !== STATUS_MATCHED) counts.unmatchedBo++;
    });

    var totalRows = counts.totalBank + counts.totalBo;
    var matchRate = totalRows === 0 ? 0 : (counts.matched * 2) / totalRows;

    overall.totalBank += counts.totalBank;
    overall.totalBo += counts.totalBo;
    overall.matched += counts.matched;
    overall.unmatchedBank += counts.unmatchedBank;
    overall.unmatchedBo += counts.unmatchedBo;
    overall.ignored += counts.ignored;

    perChannel.push({
      key: ch.key,
      displayName: ch.displayName,
      active: ch.active,
      totalBank: counts.totalBank,
      totalBo: counts.totalBo,
      matched: counts.matched,
      unmatchedBank: counts.unmatchedBank,
      unmatchedBo: counts.unmatchedBo,
      ignored: counts.ignored,
      matchRate: matchRate,
      lastRun: getLastRunTimestamp(ch.displayName)
    });
  });

  var overallTotalRows = overall.totalBank + overall.totalBo;
  overall.matchRate = overallTotalRows === 0 ? 0 : (overall.matched * 2) / overallTotalRows;

  return { overall: overall, channels: perChannel, spreadsheetUrl: ss.getUrl() };
}

function getLastRunTimestamp(channelDisplayName) {
  var ss = getDataSpreadsheet();
  var logSheet = ensureLogSheet(ss);
  var lastRow = logSheet.getLastRow();
  if (lastRow < 2) return null;
  var values = logSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var latest = null;
  for (var i = 0; i < values.length; i++) {
    if (values[i][1] === channelDisplayName) {
      var ts = values[i][0];
      if (!latest || ts > latest) latest = ts;
    }
  }
  return latest instanceof Date ? latest.toISOString() : latest;
}