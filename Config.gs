/**
 * ============================================================================
 *  Config.gs
 *  Manages the list of reconciliation channels (banks / e-wallets / payment
 *  processors), each with its own matching tolerances.
 * ============================================================================
 */

var DEFAULT_CHANNELS = [
  'UB Online',
  'GCash',
  'Maya SO 4.0',
  'Maya FMAX',
  'Maya SO Old',
  'DragonPay 4.0',
  'DragonPay FMAX'
];

function seedDefaultChannelsIfEmpty() {
  var channels = getChannels();
  if (channels.length > 0) return;
  DEFAULT_CHANNELS.forEach(function (name) {
    try {
      addChannel(name, 1, 0.01);
    } catch (err) {
      // Don't let one bad row (e.g. a name collision from a partial earlier
      // run) stop the rest of the defaults from being seeded.
      Logger.log('seedDefaultChannelsIfEmpty: failed to add "' + name + '": ' + err);
    }
  });
}

/**
 * Public: returns every configured channel with its settings and (cheap)
 * live counts are NOT included here - see Dashboard.gs for aggregated stats.
 */
function getChannels() {
  var ss = getDataSpreadsheet();
  var sheet = ensureConfigSheet(ss);
  return readSheetAsObjects(sheet, CONFIG_COLS).map(function (row) {
    return {
      key: row.ChannelKey,
      displayName: row.DisplayName,
      toleranceDays: Number(row.ToleranceDays) || 1,
      amountTolerance: Number(row.AmountTolerance) || 0.01,
      active: row.Active === true || row.Active === 'TRUE' || row.Active === 'true',
      createdAt: row.CreatedAt,
      _row: row._row
    };
  });
}

function getChannel(channelKey) {
  var channels = getChannels();
  for (var i = 0; i < channels.length; i++) {
    if (channels[i].key === channelKey) return channels[i];
  }
  return null;
}

function channelKeyExists(key) {
  return getChannel(key) !== null;
}

/**
 * Adds a new channel. Returns the created channel object.
 */
function addChannel(displayName, toleranceDays, amountTolerance) {
  displayName = String(displayName || '').trim();
  if (!displayName) throw new Error('Channel name is required.');

  var key = sanitizeChannelKey(displayName);
  if (!key) throw new Error('Channel name must contain at least one letter or number.');
  if (channelKeyExists(key)) throw new Error('A channel named "' + displayName + '" already exists.');

  var ss = getDataSpreadsheet();
  var sheet = ensureConfigSheet(ss);
  sheet.appendRow([
    key,
    displayName,
    toleranceDays || 1,
    amountTolerance || 0.01,
    true,
    new Date()
  ]);

  ensureChannelSheets(key);
  return getChannel(key);
}

function updateChannel(channelKey, updates) {
  var ss = getDataSpreadsheet();
  var sheet = ensureConfigSheet(ss);
  var rows = readSheetAsObjects(sheet, CONFIG_COLS);
  var target = null;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].ChannelKey === channelKey) { target = rows[i]; break; }
  }
  if (!target) throw new Error('Channel not found: ' + channelKey);

  if (updates.hasOwnProperty('toleranceDays')) {
    sheet.getRange(target._row, CONFIG_COLS.ToleranceDays).setValue(Number(updates.toleranceDays));
  }
  if (updates.hasOwnProperty('amountTolerance')) {
    sheet.getRange(target._row, CONFIG_COLS.AmountTolerance).setValue(Number(updates.amountTolerance));
  }
  if (updates.hasOwnProperty('active')) {
    sheet.getRange(target._row, CONFIG_COLS.Active).setValue(Boolean(updates.active));
  }
  if (updates.hasOwnProperty('displayName') && updates.displayName) {
    sheet.getRange(target._row, CONFIG_COLS.DisplayName).setValue(String(updates.displayName));
  }
  return getChannel(channelKey);
}

/**
 * Deletes a channel's config row. Optionally also deletes its BANK/BO sheets
 * (destructive - client must confirm with the user before calling this).
 */
function deleteChannel(channelKey, alsoDeleteData) {
  var ss = getDataSpreadsheet();
  var sheet = ensureConfigSheet(ss);
  var rows = readSheetAsObjects(sheet, CONFIG_COLS);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].ChannelKey === channelKey) {
      sheet.deleteRow(rows[i]._row);
      break;
    }
  }
  if (alsoDeleteData) {
    var bankSheet = ss.getSheetByName(bankSheetName(channelKey));
    var boSheet = ss.getSheetByName(boSheetName(channelKey));
    if (bankSheet) ss.deleteSheet(bankSheet);
    if (boSheet) ss.deleteSheet(boSheet);
  }
  return true;
}

/**
 * Scans the data spreadsheet for "BANK - X" / "BO - X" sheet pairs that
 * don't have a matching row in the Config tab - e.g. because the Config
 * row was removed via "Delete channel" (which intentionally leaves the
 * underlying data sheets alone) but the sheets themselves were never
 * deleted - and recreates a Config row for each one so the channel shows
 * back up in the app. This NEVER creates, renames, or deletes any BANK/BO
 * sheet; it only ever appends missing Config rows.
 */
function findOrphanedChannelSheets() {
  var ss = getDataSpreadsheet();
  var existingKeys = {};
  getChannels().forEach(function (c) { existingKeys[c.key] = true; });

  var bankPrefix = 'BANK - ';
  var boPrefix = 'BO - ';
  var pairs = {};

  ss.getSheets().forEach(function (s) {
    var name = s.getName();
    if (name.indexOf(bankPrefix) === 0) {
      var key = name.substring(bankPrefix.length);
      pairs[key] = pairs[key] || {};
      pairs[key].bank = true;
    } else if (name.indexOf(boPrefix) === 0) {
      var key2 = name.substring(boPrefix.length);
      pairs[key2] = pairs[key2] || {};
      pairs[key2].bo = true;
    }
  });

  var orphans = [];
  Object.keys(pairs).forEach(function (key) {
    if (existingKeys[key]) return; // already has a Config row - not orphaned
    orphans.push({
      key: key,
      hasBank: !!pairs[key].bank,
      hasBo: !!pairs[key].bo,
      complete: !!(pairs[key].bank && pairs[key].bo)
    });
  });

  return orphans;
}

/**
 * Recreates Config rows for orphaned channel sheet pairs found by
 * findOrphanedChannelSheets(). Only relinks pairs that have BOTH a BANK and
 * a BO sheet, to avoid guessing at a display name / config for a half sheet.
 * New rows default to 1 day / 0.01 amount tolerance and Active = true;
 * adjust those afterwards in the Channels table if needed.
 */
function relinkOrphanedChannels() {
  var ss = getDataSpreadsheet();
  var sheet = ensureConfigSheet(ss);
  var orphans = findOrphanedChannelSheets();
  var now = new Date();
  var relinked = [];
  var skippedIncomplete = [];

  orphans.forEach(function (o) {
    if (!o.complete) { skippedIncomplete.push(o.key); return; }
    sheet.appendRow([o.key, o.key, 1, 0.01, true, now]);
    relinked.push(o.key);
  });

  return { relinked: relinked, skippedIncomplete: skippedIncomplete };
}