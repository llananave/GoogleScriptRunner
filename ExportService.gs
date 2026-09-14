/**
 * ============================================================================
 *  ExportService.gs
 *  Produces downloadable CSV exports and a saved report tab per channel.
 * ============================================================================
 */

function csvEscape(value) {
  if (value === null || value === undefined) return '';
  var s = value instanceof Date ? value.toISOString() : String(value);
  if (/[",\n]/.test(s)) {
    s = '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function toCsv(headers, rows) {
  var lines = [headers.map(csvEscape).join(',')];
  rows.forEach(function (r) {
    lines.push(r.map(csvEscape).join(','));
  });
  return lines.join('\n');
}

/**
 * Resolves the Bank sheet (if any) and Back-Office sheet (via the channel's
 * match target, if any) that make up one channel's data, for export/report
 * purposes. Never throws - a side is simply omitted if not applicable.
 */
function resolveChannelSides(channel) {
  var bankSheet = null, boSheet = null, boChannel = null;
  if (channel.type === CHANNEL_TYPE_BANK || channel.type === CHANNEL_TYPE_BOTH) {
    bankSheet = ensureChannelSheets(channel.key).bankSheet;
    boChannel = resolveMatchTarget(channel);
  }
  if (channel.type === CHANNEL_TYPE_BO) {
    boChannel = channel;
  }
  if (boChannel) {
    boSheet = ensureChannelSheets(boChannel.key).boSheet;
  }
  return { bankSheet: bankSheet, boSheet: boSheet, boChannel: boChannel };
}

/**
 * Builds a CSV string of every row (matched + unmatched + ignored) for a
 * channel, for the "Export Full Data" button. Returned to the client, which
 * triggers a browser download - no Drive permissions required. Any "extra"
 * columns the user retained at import time are included, with their
 * original header text, one column per distinct extra header seen on either
 * side.
 */
function exportChannelCsv(channelKey) {
  var channel = getChannel(channelKey);
  if (!channel) throw new Error('Unknown channel: ' + channelKey);
  var sides = resolveChannelSides(channel);

  var bankRows = sides.bankSheet ? readSheetRowsWithExtras(sides.bankSheet, BANK_COLS) : [];
  var boRows = sides.boSheet ? readSheetRowsWithExtras(sides.boSheet, BO_COLS) : [];

  var extraHeaders = [];
  var seen = {};
  bankRows.concat(boRows).forEach(function (r) {
    Object.keys(r.extra || {}).forEach(function (h) {
      if (!seen[h]) { seen[h] = true; extraHeaders.push(h); }
    });
  });

  var headers = ['Side', 'Date', 'Reference', 'Amount', 'Status', 'Match Type', 'Match Id', 'Resolution Note'].concat(extraHeaders);
  var rows = [];

  function extraValues(r) {
    return extraHeaders.map(function (h) { return (r.extra && r.extra[h] !== undefined) ? r.extra[h] : ''; });
  }

  bankRows.forEach(function (r) {
    rows.push(['BANK', r.Date, r.Reference, r.Amount, r.Status, r.MatchType, r.MatchId, r.ResolutionNote].concat(extraValues(r)));
  });
  boRows.forEach(function (r) {
    rows.push(['BACK-OFFICE', r.Date, r.Reference, r.Amount, r.Status, r.MatchType, r.MatchId, r.ResolutionNote].concat(extraValues(r)));
  });

  return {
    filename: 'Reconciliation - ' + channel.displayName + ' - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd') + '.csv',
    csv: toCsv(headers, rows)
  };
}

/**
 * Writes a dated snapshot report as a new tab in the data spreadsheet and
 * returns its URL so the client can open it directly.
 */
function writeChannelReportSheet(channelKey) {
  var channel = getChannel(channelKey);
  if (!channel) throw new Error('Unknown channel: ' + channelKey);
  var ss = getDataSpreadsheet();
  var sides = resolveChannelSides(channel);

  var bankRows = sides.bankSheet ? readSheetRowsWithExtras(sides.bankSheet, BANK_COLS) : [];
  var boRows = sides.boSheet ? readSheetRowsWithExtras(sides.boSheet, BO_COLS) : [];
  var matched = bankRows.filter(function (r) { return r.Status === STATUS_MATCHED; }).length;
  var unmatchedBank = bankRows.filter(function (r) { return r.Status === STATUS_UNMATCHED || !r.Status; });
  var unmatchedBo = boRows.filter(function (r) { return r.Status === STATUS_UNMATCHED || !r.Status; });

  var bankExtraHeaders = sides.bankSheet ? getExtraColumnDefs(sides.bankSheet, BANK_COLS).map(function (d) { return d.header; }) : [];
  var boExtraHeaders = sides.boSheet ? getExtraColumnDefs(sides.boSheet, BO_COLS).map(function (d) { return d.header; }) : [];

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd HHmm');
  var reportName = ('Report - ' + channel.displayName + ' - ' + stamp).substring(0, 100);
  var report = ss.insertSheet(reportName);

  report.appendRow(['Reconciliation Report']);
  report.appendRow(['Channel', channel.displayName]);
  if (sides.boChannel) report.appendRow(['Matches Against', sides.boChannel.displayName]);
  report.appendRow(['Generated At', new Date()]);
  report.appendRow([]);
  report.appendRow(['Total Bank Rows', bankRows.length]);
  report.appendRow(['Total Back-Office Rows', boRows.length]);
  report.appendRow(['Matched Pairs', matched]);
  report.appendRow(['Unmatched Bank Rows', unmatchedBank.length]);
  report.appendRow(['Unmatched Back-Office Rows', unmatchedBo.length]);
  report.appendRow([]);

  var bankHeaderRow = ['Row Id', 'Date', 'Reference', 'Amount', 'Status', 'Match Id', 'Match Type', 'Resolution Note'].concat(bankExtraHeaders);
  report.appendRow(['-- Unmatched Bank Rows --']);
  report.appendRow(bankHeaderRow);
  unmatchedBank.forEach(function (r) {
    var extra = bankExtraHeaders.map(function (h) { return r.extra[h]; });
    report.appendRow([r.RowId, r.Date, r.Reference, r.Amount, r.Status, r.MatchId, r.MatchType, r.ResolutionNote].concat(extra));
  });
  report.appendRow([]);

  var boHeaderRow = ['Row Id', 'Date', 'Reference', 'Amount', 'Status', 'Match Id', 'Match Type', 'Resolution Note'].concat(boExtraHeaders);
  report.appendRow(['-- Unmatched Back-Office Rows --']);
  report.appendRow(boHeaderRow);
  unmatchedBo.forEach(function (r) {
    var extra = boExtraHeaders.map(function (h) { return r.extra[h]; });
    report.appendRow([r.RowId, r.Date, r.Reference, r.Amount, r.Status, r.MatchId, r.MatchType, r.ResolutionNote].concat(extra));
  });

  report.getRange('A1').setFontWeight('bold').setFontSize(14);
  report.autoResizeColumns(1, Math.max(bankHeaderRow.length, boHeaderRow.length));

  return { url: ss.getUrl() + '#gid=' + report.getSheetId(), sheetName: reportName };
}
