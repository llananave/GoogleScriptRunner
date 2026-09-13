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
    var type = String(row.Type || '').trim().toUpperCase();
    if ([CHANNEL_TYPE_BANK, CHANNEL_TYPE_BO, CHANNEL_TYPE_BOTH].indexOf(type) === -1) type = CHANNEL_TYPE_BOTH;
    return {
      key: row.ChannelKey,
      displayName: row.DisplayName,
      toleranceDays: Number(row.ToleranceDays) || 1,
      amountTolerance: Number(row.AmountTolerance) || 0.01,
      active: row.Active === true || row.Active === 'TRUE' || row.Active === 'true',
      // Channels created before this feature existed have a blank Type cell -
      // treated as BOTH, which is exactly their original (only) behavior.
      type: type,
      // Never send a raw Date object across google.script.run inside an array
      // of objects - it can fail to serialize and cause the WHOLE response to
      // come back as null on the client, wiping out every channel at once.
      createdAt: (row.CreatedAt instanceof Date) ? row.CreatedAt.toISOString() : (row.CreatedAt || null),
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
 * Adds a new channel. `type` is 'BANK', 'BO', or 'BOTH' (default 'BOTH' for
 * backward compatibility with callers - like the default-channel seeder -
 * that don't specify one). Only the sheet(s) implied by `type` are created:
 * a 'BANK' channel gets just a BANK sheet, a 'BO' channel gets just a BO
 * sheet, and 'BOTH' gets the classic paired sheets.
 */
function addChannel(displayName, toleranceDays, amountTolerance, type) {
  displayName = String(displayName || '').trim();
  if (!displayName) throw new Error('Channel name is required.');

  type = String(type || CHANNEL_TYPE_BOTH).trim().toUpperCase();
  if ([CHANNEL_TYPE_BANK, CHANNEL_TYPE_BO, CHANNEL_TYPE_BOTH].indexOf(type) === -1) {
    throw new Error('Channel type must be Bank, Back Office, or Both.');
  }

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
    new Date(),
    type
  ]);

  ensureChannelSheets(key, type);
  return getChannel(key);
}

/**
 * Updates a channel's settings. Supported keys in `updates`:
 *   toleranceDays, amountTolerance, active, displayName,
 *   type ('BANK' | 'BO' | 'BOTH') - switching type will create any newly
 *     required sheet automatically. If `deleteUnusedSide` is also true and
 *     the change drops a side (e.g. BOTH -> BANK), that side's now-unused
 *     sheet is deleted too - this is destructive, the client must confirm
 *     with the user first.
 */
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
  if (updates.hasOwnProperty('type')) {
    var newType = String(updates.type || '').trim().toUpperCase();
    if ([CHANNEL_TYPE_BANK, CHANNEL_TYPE_BO, CHANNEL_TYPE_BOTH].indexOf(newType) === -1) {
      throw new Error('Channel type must be Bank, Back Office, or Both.');
    }
    var oldTypeRaw = String(target.Type || '').trim().toUpperCase();
    var oldType = [CHANNEL_TYPE_BANK, CHANNEL_TYPE_BO, CHANNEL_TYPE_BOTH].indexOf(oldTypeRaw) === -1 ? CHANNEL_TYPE_BOTH : oldTypeRaw;

    sheet.getRange(target._row, CONFIG_COLS.Type).setValue(newType);
    ensureChannelSheets(channelKey, newType); // create any newly-needed side

    if (updates.deleteUnusedSide) {
      var droppedBank = (oldType === CHANNEL_TYPE_BANK || oldType === CHANNEL_TYPE_BOTH) &&
        (newType === CHANNEL_TYPE_BO);
      var droppedBo = (oldType === CHANNEL_TYPE_BO || oldType === CHANNEL_TYPE_BOTH) &&
        (newType === CHANNEL_TYPE_BANK);
      if (droppedBank) {
        var bs = ss.getSheetByName(bankSheetName(channelKey));
        if (bs) ss.deleteSheet(bs);
      }
      if (droppedBo) {
        var bo = ss.getSheetByName(boSheetName(channelKey));
        if (bo) ss.deleteSheet(bo);
      }
    }
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
 * Recreates Config rows for orphaned channel sheets found by
 * findOrphanedChannelSheets(). A pair (both BANK and BO sheets present) is
 * relinked as type BOTH; a lone sheet is relinked as type BANK or BO to
 * match whichever side actually exists.
 */
function relinkOrphanedChannels() {
  var ss = getDataSpreadsheet();
  var sheet = ensureConfigSheet(ss);
  var orphans = findOrphanedChannelSheets();
  var now = new Date();
  var relinked = [];

  orphans.forEach(function (o) {
    var type = o.complete ? CHANNEL_TYPE_BOTH : (o.hasBank ? CHANNEL_TYPE_BANK : CHANNEL_TYPE_BO);
    sheet.appendRow([o.key, o.key, 1, 0.01, true, now, type]);
    relinked.push({ key: o.key, type: type });
  });

  return { relinked: relinked };
}