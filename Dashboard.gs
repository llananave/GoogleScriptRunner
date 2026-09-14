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
    var boTarget = (ch.type === CHANNEL_TYPE_BANK || ch.type === CHANNEL_TYPE_BOTH) ? resolveMatchTarget(ch) : null;

    var counts = { totalBank: bankRows.length, totalBo: boRows.length, matched: 0, unmatchedBank: 0, unmatchedBo: 0, ignored: 0 };
    var matchedBank = 0, matchedBo = 0;

    bankRows.forEach(function (r) {
      if (r.Status === STATUS_MATCHED) matchedBank++;
      else if (r.Status === STATUS_IGNORED) counts.ignored++;
      else counts.unmatchedBank++;
    });
    boRows.forEach(function (r) {
      if (r.Status === STATUS_MATCHED) matchedBo++;
      else if (r.Status === STATUS_IGNORED) counts.ignored++;
      else counts.unmatchedBo++;
    });

    // Every match pair is counted exactly once, from its owning Bank
    // channel's side, so summing this into `overall` never double-counts -
    // even though the paired Back-Office row may live on a different
    // channel's sheet (a shared Back-Office target). `matchedBank` is what
    // feeds the site-wide total.
    counts.matched = matchedBank;

    // For the per-channel DISPLAY row, a Back-Office-only channel has no
    // Bank rows of its own to count from, so show its own matched count
    // instead (this is display-only and intentionally NOT added to
    // `overall.matched`, to avoid double-counting pairs already counted via
    // the Bank channel(s) that match into it).
    var displayMatched = ch.type === CHANNEL_TYPE_BO ? matchedBo : matchedBank;
    var displayMatchRate;
    if (ch.type === CHANNEL_TYPE_BANK) {
      displayMatchRate = counts.totalBank === 0 ? 0 : matchedBank / counts.totalBank;
    } else if (ch.type === CHANNEL_TYPE_BO) {
      displayMatchRate = counts.totalBo === 0 ? 0 : matchedBo / counts.totalBo;
    } else {
      var totalRows = counts.totalBank + counts.totalBo;
      displayMatchRate = totalRows === 0 ? 0 : (matchedBank + matchedBo) / totalRows;
    }

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
      type: ch.type,
      matchesTo: boTarget ? boTarget.displayName : null,
      totalBank: counts.totalBank,
      totalBo: counts.totalBo,
      matched: displayMatched,
      unmatchedBank: counts.unmatchedBank,
      unmatchedBo: counts.unmatchedBo,
      ignored: counts.ignored,
      matchRate: displayMatchRate,
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