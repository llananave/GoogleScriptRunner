/**
 * ============================================================================
 *  BANK & BACK-OFFICE RECONCILIATION APP
 *  Main entry point (Code.gs)
 * ----------------------------------------------------------------------------
 *  A standalone Apps Script web application that reconciles raw bank
 *  statement exports against back-office / platform transaction exports
 *  across multiple payment channels (e.g. UB Online, GCash, Maya, DragonPay).
 *
 *  Inspired by a manual multi-tab spreadsheet reconciliation workbook,
 *  rebuilt as a data-driven web app so any number of channels can be
 *  configured without touching code.
 * ============================================================================
 */

// The human-facing name of the app + the backing data spreadsheet it creates.
var APP_TITLE = 'Bank & Back-Office Reconciliation';
var DATA_SPREADSHEET_NAME = 'Recon App - Data Store';

/**
 * Serves the web app UI.
 */
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  template.appTitle = APP_TITLE;
  return template
    .evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://www.gstatic.com/images/icons/material/system/2x/account_balance_googblue_48dp.png');
}

/**
 * Allows Index.html to pull in CSS.html / JS.html as <?!= include('X'); ?>.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Adds a menu when this project is opened as a container-bound script
 * (harmless / unused when deployed purely as a standalone web app).
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Reconciliation App')
      .addItem('Open Web App', 'showAppUrlDialog')
      .addItem('Initialize / Repair Data Store', 'initializeApp')
      .addToUi();
  } catch (err) {
    // Not running in a Sheets context - safe to ignore.
  }
}

function showAppUrlDialog() {
  var url = ScriptApp.getService().getUrl();
  var ui = SpreadsheetApp.getUi();
  ui.alert('Reconciliation Web App', url || 'Deploy this project as a Web App first (Deploy > New deployment).', ui.ButtonSet.OK);
}

/**
 * Idempotent bootstrap - safe to call any number of times.
 * Called automatically by the client when the app first loads.
 */
function initializeApp() {
  var ss = getDataSpreadsheet();
  ensureConfigSheet(ss);
  ensureLogSheet(ss);
  seedDefaultChannelsIfEmpty();
  return {
    ok: true,
    spreadsheetUrl: ss.getUrl(),
    spreadsheetName: ss.getName()
  };
}

/**
 * Diagnostic snapshot for the Settings tab, so setup problems (no data
 * spreadsheet yet, zero channels seeded, wrong executing user, etc.) can be
 * seen directly in the UI instead of requiring the Apps Script Executions
 * log.
 */
function getSetupDiagnostics() {
  var props = PropertiesService.getScriptProperties();
  var storedId = props.getProperty('DATA_SS_ID');
  var info = { storedSpreadsheetId: storedId || null, ok: false };

  try {
    var ss = getDataSpreadsheet();
    info.spreadsheetUrl = ss.getUrl();
    info.spreadsheetName = ss.getName();
    info.sheetNames = ss.getSheets().map(function (s) { return s.getName(); });

    var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    info.configSheetExists = !!configSheet;
    info.configRowCount = configSheet ? Math.max(0, configSheet.getLastRow() - 1) : 0;

    var activeUser = '';
    var effectiveUser = '';
    try { activeUser = Session.getActiveUser().getEmail(); } catch (e1) {}
    try { effectiveUser = Session.getEffectiveUser().getEmail(); } catch (e2) {}
    info.activeUser = activeUser || '(not visible - normal for some deployment types)';
    info.effectiveUser = effectiveUser || '(not visible)';
    info.ok = true;
  } catch (err) {
    info.error = (err && err.message) ? err.message : String(err);
  }

  return info;
}
