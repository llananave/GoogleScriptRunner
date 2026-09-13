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
 * Builds a CSV string of every row (matched + unmatched + ignored) for a
 * channel, for the "Export Full Data" button. Returned to the client, which
 * triggers a browser download - no Drive permissions required.
 */
function exportChannelCsv(channelKey) {
  var channel = getChannel(channelKey);
  if (!channel) throw new Error('Unknown channel: ' + channelKey);
  var sheets = ensureChannelSheets(channelKey);

  var bankRows = sheets.bankSheet ? readSheetAsObjects(sheets.bankSheet, BANK_COLS) : [];
  var boRows = sheets.boSheet ? readSheetAsObjects(sheets.boSheet, BO_COLS) : [];

  var headers = ['Side', 'Date', 'Reference', 'Secondary Ref / Patron', 'Description / Type', 'Amount', 'Status', 'Match Type', 'Match Id', 'Resolution Note'];
  var rows = [];

  bankRows.forEach(function (r) {
    rows.push(['BANK', r.Date, r.Reference, '', r.Description, r.Amount, r.Status, r.MatchType, r.MatchId, r.ResolutionNote]);
  });
  boRows.forEach(function (r) {
    rows.push(['BACK-OFFICE', r.Date, r.Reference, r.PatronId, r.TxnType, r.Amount, r.Status, r.MatchType, r.MatchId, r.ResolutionNote]);
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
  var sheets = ensureChannelSheets(channelKey);

  var bankRows = sheets.bankSheet ? readSheetAsObjects(sheets.bankSheet, BANK_COLS) : [];
  var boRows = sheets.boSheet ? readSheetAsObjects(sheets.boSheet, BO_COLS) : [];
  var matched = bankRows.filter(function (r) { return r.Status === STATUS_MATCHED; }).length;
  var unmatchedBank = bankRows.filter(function (r) { return r.Status === STATUS_UNMATCHED || !r.Status; }).length;
  var unmatchedBo = boRows.filter(function (r) { return r.Status === STATUS_UNMATCHED || !r.Status; }).length;

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd HHmm');
  var reportName = ('Report - ' + channel.displayName + ' - ' + stamp).substring(0, 100);
  var report = ss.insertSheet(reportName);

  report.appendRow(['Reconciliation Report']);
  report.appendRow(['Channel', channel.displayName]);
  report.appendRow(['Generated At', new Date()]);
  report.appendRow([]);
  report.appendRow(['Total Bank Rows', bankRows.length]);
  report.appendRow(['Total Back-Office Rows', boRows.length]);
  report.appendRow(['Matched Pairs', matched]);
  report.appendRow(['Unmatched Bank Rows', unmatchedBank]);
  report.appendRow(['Unmatched Back-Office Rows', unmatchedBo]);
  report.appendRow([]);
  report.appendRow(['-- Unmatched Bank Rows --']);
  report.appendRow(BANK_HEADERS);
  bankRows.filter(function (r) { return r.Status === STATUS_UNMATCHED || !r.Status; }).forEach(function (r) {
    report.appendRow([r.RowId, r.Date, r.Reference, r.Description, r.Deposit, r.Withdrawal, r.Amount, r.Status, r.MatchId, r.MatchType, r.ResolutionNote, r.ImportedAt, r.SourceHash]);
  });
  report.appendRow([]);
  report.appendRow(['-- Unmatched Back-Office Rows --']);
  report.appendRow(BO_HEADERS);
  boRows.filter(function (r) { return r.Status === STATUS_UNMATCHED || !r.Status; }).forEach(function (r) {
    report.appendRow([r.RowId, r.Date, r.PatronId, r.TxnType, r.Reference, r.SecondaryReference, r.Amount, r.Status, r.MatchId, r.MatchType, r.ResolutionNote, r.ImportedAt, r.SourceHash]);
  });

  report.getRange('A1').setFontWeight('bold').setFontSize(14);
  report.getRange('A11').setFontWeight('bold');
  report.autoResizeColumns(1, 13);

  return { url: ss.getUrl() + '#gid=' + report.getSheetId(), sheetName: reportName };
}
