const APP_CONFIG = {
  title: 'Tennis Schedule Manager',
  playersSheet: 'Players',
  scheduleSheet: 'Schedule',
  settingsSheet: 'Settings',
  magicLinkMinutes: 15,
  sessionHours: 6,
  defaultPlayDays: [1, 3, 5],
  scheduleSessionCount: 21
};

function doGet(e) {
  const page = HtmlService.createTemplateFromFile('Index');
  page.magicToken = e && e.parameter ? String(e.parameter.token || '') : '';
  return page.evaluate()
    .setTitle(APP_CONFIG.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Run once from the Apps Script editor as the project owner.
function setupProject() {
  _requireOwnerEditor_();
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('DATA_SPREADSHEET_ID');
  let book;
  if (spreadsheetId) {
    book = SpreadsheetApp.openById(spreadsheetId);
  } else {
    book = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create(APP_CONFIG.title + ' Data');
    props.setProperty('DATA_SPREADSHEET_ID', book.getId());
  }
  _ensureSheet_(book, APP_CONFIG.playersSheet, ['PlayerID', 'Name', 'Email', 'Active']);
  _ensureSheet_(book, APP_CONFIG.scheduleSheet, ['Date', 'PlayerID', 'Status']);
  _ensureSheet_(book, APP_CONFIG.settingsSheet, ['Key', 'Value']);
  if (!_getSetting_('PlayDays')) _setSetting_('PlayDays', JSON.stringify(APP_CONFIG.defaultPlayDays));
  if (!_getSetting_('ScheduleDates')) {
    _setSetting_('ScheduleDates', JSON.stringify(_generateDates_(APP_CONFIG.scheduleSessionCount, APP_CONFIG.defaultPlayDays)));
  }
  if (_getSetting_('ManagerPlayerID') === null) _setSetting_('ManagerPlayerID', '');
  return { spreadsheetUrl: book.getUrl(), message: 'Data sheets are ready. Add the roster and set the manager in Settings.' };
}

// Run once from the Apps Script editor as the project owner to approve sending mail.
function authorizeEmailSending() {
  _requireOwnerEditor_();
  return MailApp.getRemainingDailyQuota();
}

function requestMagicLink(name, email) {
  const cleanName = String(name || '').trim();
  const cleanEmail = _cleanEmail_(email);
  if (!cleanName || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleanEmail)) {
    throw new Error('Enter your roster name and a valid email address.');
  }
  const rosterPlayer = _findPlayerByEmail_(cleanEmail);
  // Return the same message for unknown names/emails to avoid disclosing roster membership.
  if (rosterPlayer && rosterPlayer.name.toLowerCase() === cleanName.toLowerCase()) {
    _sendMagicLink_(rosterPlayer);
  }
  return 'If those details match the player list, a sign-in link has been sent.';
}

function consumeMagicLink(token) {
  const rawToken = String(token || '');
  if (!rawToken) throw new Error('This sign-in link is missing or invalid. Request a new link.');
  const cache = CacheService.getScriptCache();
  const key = 'magic_' + _hash_(rawToken);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let payload;
  try {
    payload = cache.get(key);
    if (payload) cache.remove(key); // A link can create a session once only.
  } finally {
    lock.releaseLock();
  }
  if (!payload) throw new Error('This sign-in link has expired or was already used. Request a new link.');
  const ticket = JSON.parse(payload);
  if (ticket.expiresAt < Date.now()) throw new Error('This sign-in link has expired. Request a new link.');
  const player = _findPlayerById_(ticket.playerID);
  if (!player) throw new Error('This player is no longer on the active roster.');
  const sessionToken = _newToken_();
  const session = { playerID: player.playerID, expiresAt: Date.now() + APP_CONFIG.sessionHours * 60 * 60 * 1000 };
  cache.put('session_' + _hash_(sessionToken), JSON.stringify(session), APP_CONFIG.sessionHours * 60 * 60);
  return { token: sessionToken, user: _publicUser_(player) };
}

function getSession(sessionToken) {
  const session = _getSession_(sessionToken);
  return _publicUser_(_findPlayerById_(session.playerID));
}

function signOut(sessionToken) {
  if (sessionToken) CacheService.getScriptCache().remove('session_' + _hash_(String(sessionToken)));
  return true;
}

function getAppData(sessionToken) {
  const session = _getSession_(sessionToken);
  const current = _findPlayerById_(session.playerID);
  const players = _readPlayers_().filter(player => player.active);
  const dates = _getScheduleDates_();
  const availability = {};
  players.forEach(player => availability[player.playerID] = {});
  _readSchedule_().forEach(item => {
    if (availability[item.playerID]) availability[item.playerID][item.date] = item.status;
  });
  players.forEach(player => {
    dates.forEach(date => {
      if (!(date in availability[player.playerID])) availability[player.playerID][date] = '';
    });
    player.availability = availability[player.playerID];
  });
  return {
    dates: dates,
    players: players.map(player => ({
      playerID: player.playerID,
      name: player.name,
      email: player.email,
      availability: player.availability
    })),
    playDays: _getPlayDays_(),
    user: _publicUser_(current)
  };
}

function savePlayerStatus(sessionToken, playerID, date, newStatus) {
  const session = _getSession_(sessionToken);
  const currentUser = _findPlayerById_(session.playerID);
  const manager = _isManager_(currentUser.playerID);
  const targetID = String(playerID || '');
  const status = String(newStatus || '');
  if (!manager && targetID !== currentUser.playerID) throw new Error('You can only update your own row.');
  if (!['', 'A', 'S', '/', 'B'].includes(status)) throw new Error('That player status is not allowed.');
  if (!manager && status === 'B') throw new Error('Only the Schedule Manager can mark a player as booked.');
  const activePlayer = _findPlayerById_(targetID);
  if (!activePlayer) throw new Error('That player is not on the active roster.');
  const dateKey = String(date || '');
  if (!_getScheduleDates_().includes(dateKey)) throw new Error('That date is not in the active schedule.');
  const prior = _getStatus_(targetID, dateKey);
  if (!manager && prior === 'B') throw new Error('Only the Schedule Manager can change or clear a booking.');
  _setStatus_(targetID, dateKey, status);
  return true;
}

function setPlayDays(sessionToken, days) {
  _requireManager_(sessionToken);
  const selected = Array.from(new Set((days || []).map(Number).filter(day => Number.isInteger(day) && day >= 0 && day <= 6))).sort((a, b) => a - b);
  if (!selected.length) throw new Error('Select at least one play day.');
  _setSetting_('PlayDays', JSON.stringify(selected));
  _setSetting_('ScheduleDates', JSON.stringify(_generateDates_(APP_CONFIG.scheduleSessionCount, selected)));
  return getAppData(sessionToken);
}

function rollForwardCurrentWeek(sessionToken) {
  _requireManager_(sessionToken);
  const dates = _getScheduleDates_();
  const today = new Date();
  const weekStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  weekStart.setDate(weekStart.getDate() - ((weekStart.getDay() + 6) % 7)); // Monday start
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekEnd.getDate() + 7);
  const startKey = _formatDate_(weekStart);
  const endKey = _formatDate_(weekEnd);
  const rolling = dates.filter(date => date >= startKey && date < endKey);
  if (!rolling.length) throw new Error('There are no scheduled play days in the current week to roll forward.');
  const remaining = dates.filter(date => !rolling.includes(date));
  const lastDate = remaining.length ? remaining[remaining.length - 1] : _formatDate_(today);
  const added = _generateAfter_(rolling.length, _parseDate_(lastDate), _getPlayDays_());
  _setSetting_('ScheduleDates', JSON.stringify(remaining.concat(added).sort()));
  return getAppData(sessionToken);
}

function importPlayerRoster(sessionToken, csvText) {
  _requireManager_(sessionToken);
  const rows = Utilities.parseCsv(String(csvText || ''));
  if (rows.length < 2) throw new Error('The CSV needs a header and at least one player.');
  const headers = rows[0].map(value => String(value).replace(/^\uFEFF/, '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''));
  const idCol = headers.indexOf('playerid');
  const nameCol = headers.indexOf('name');
  const emailCol = headers.indexOf('email');
  if (idCol < 0 || nameCol < 0 || emailCol < 0) throw new Error('CSV headers must include PlayerID, Name, and Email.');
  const seenIDs = {};
  const seenEmails = {};
  const imported = [];
  rows.slice(1).forEach((row, index) => {
    if (row.every(value => !String(value).trim())) return;
    const player = {
      playerID: String(row[idCol] || '').trim(),
      name: String(row[nameCol] || '').trim(),
      email: _cleanEmail_(row[emailCol])
    };
    if (!player.playerID || !player.name || !player.email) throw new Error('CSV row ' + (index + 2) + ' needs PlayerID, Name, and Email.');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(player.email)) throw new Error('CSV row ' + (index + 2) + ' has an invalid email.');
    if (seenIDs[player.playerID] || seenEmails[player.email]) throw new Error('CSV row ' + (index + 2) + ' repeats a PlayerID or email.');
    seenIDs[player.playerID] = true;
    seenEmails[player.email] = true;
    imported.push(player);
  });
  const book = _getBook_();
  const sheet = book.getSheetByName(APP_CONFIG.playersSheet);
  const existing = _readPlayers_();
  const byID = {};
  const byEmail = {};
  existing.forEach((player, index) => {
    byID[player.playerID] = index;
    if (player.email) byEmail[player.email] = index;
  });
  imported.forEach(player => {
    const index = byID[player.playerID] !== undefined ? byID[player.playerID] : byEmail[player.email];
    if (index !== undefined && existing[index].playerID !== player.playerID) {
      throw new Error('PlayerID is a stable identifier and cannot be changed for an existing email. Keep the same PlayerID in future roster uploads.');
    }
  });
  let added = 0;
  let updated = 0;
  imported.forEach(player => {
    let index = byID[player.playerID];
    if (index === undefined) index = byEmail[player.email];
    if (index !== undefined) {
      const old = existing[index];
      const rowNumber = old.row;
      sheet.getRange(rowNumber, 1, 1, 4).setValues([[player.playerID, player.name, player.email, true]]);
      old.playerID = player.playerID;
      old.name = player.name;
      old.email = player.email;
      updated++;
    } else {
      sheet.appendRow([player.playerID, player.name, player.email, true]);
      added++;
    }
  });
  return { added: added, updated: updated, data: getAppData(sessionToken) };
}

function importScheduleCsv(sessionToken, csvText) {
  _requireManager_(sessionToken);
  const rows = Utilities.parseCsv(String(csvText || ''));
  if (rows.length < 2) throw new Error('The schedule CSV needs a header and at least one player row.');
  const headers = rows[0].map(value => String(value).replace(/^\uFEFF/, '').trim());
  const players = _readPlayers_().filter(player => player.active);
  const keyHeader = String(headers[0] || '').toLowerCase();
  const dates = headers.slice(1);
  if (keyHeader !== 'player' && keyHeader !== 'playerid') throw new Error('The first schedule CSV column must be Player or PlayerID.');
  if (!dates.length || dates.some(date => !/^\d{4}-\d{2}-\d{2}$/.test(date))) throw new Error('Schedule date columns must use YYYY-MM-DD format.');
  const imported = [];
  rows.slice(1).forEach((row, index) => {
    if (row.every(value => !String(value).trim())) return;
    const identifier = String(row[0] || '').trim();
    const player = keyHeader === 'playerid'
      ? players.find(item => item.playerID === identifier)
      : players.find(item => item.name === identifier);
    if (!player) throw new Error('Schedule CSV row ' + (index + 2) + ' does not match an active player.');
    dates.forEach((date, column) => {
      let status = String(row[column + 1] || '').trim().toUpperCase();
      if (status === 'P') status = 'A'; // Keep older schedule exports compatible.
      if (!['', 'A', 'S', '/', 'B'].includes(status)) throw new Error('Unsupported status on schedule CSV row ' + (index + 2) + '.');
      if (status) imported.push([date, player.playerID, status]);
    });
  });
  const sheet = _getBook_().getSheetByName(APP_CONFIG.scheduleSheet);
  if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  if (imported.length) sheet.getRange(2, 1, imported.length, 3).setValues(imported);
  _setSetting_('ScheduleDates', JSON.stringify(dates));
  return getAppData(sessionToken);
}

function addPlayer(sessionToken, playerID, name, email) {
  _requireManager_(sessionToken);
  const record = _validatePlayer_({ playerID: playerID, name: name, email: email });
  if (_findPlayerById_(record.playerID, true)) throw new Error('That PlayerID already exists.');
  if (_findPlayerByEmail_(record.email, true)) throw new Error('That email already exists.');
  _getBook_().getSheetByName(APP_CONFIG.playersSheet).appendRow([record.playerID, record.name, record.email, true]);
  return getAppData(sessionToken);
}

function removePlayer(sessionToken, playerID) {
  _requireManager_(sessionToken);
  const player = _findPlayerById_(String(playerID || ''), true);
  if (!player) throw new Error('Player not found.');
  if (_isManager_(player.playerID)) throw new Error('Assign a different Schedule Manager before removing this player.');
  _getBook_().getSheetByName(APP_CONFIG.playersSheet).getRange(player.row, 4).setValue(false);
  return getAppData(sessionToken);
}

function resendPlayerLink(sessionToken, playerID) {
  _requireManager_(sessionToken);
  const player = _findPlayerById_(String(playerID || ''));
  if (!player) throw new Error('Player not found on the active roster.');
  _sendMagicLink_(player);
  return 'A fresh sign-in link was sent to ' + player.email + '.';
}

function _sendMagicLink_(player) {
  const cache = CacheService.getScriptCache();
  const rateKey = 'mail_' + _hash_(player.email.toLowerCase());
  if (cache.get(rateKey)) throw new Error('A link was just sent to this email. Please wait a minute before trying again.');
  const appUrl = ScriptApp.getService().getUrl();
  if (!appUrl) throw new Error('Deploy the Apps Script project as a web app before sending links.');
  const token = _newToken_();
  const cacheKey = 'magic_' + _hash_(token);
  const expiresAt = Date.now() + APP_CONFIG.magicLinkMinutes * 60 * 1000;
  cache.put(cacheKey, JSON.stringify({ playerID: player.playerID, expiresAt: expiresAt }), APP_CONFIG.magicLinkMinutes * 60);
  const link = appUrl + '?token=' + encodeURIComponent(token);
  try {
    MailApp.sendEmail({
      to: player.email,
      subject: 'Your Tennis Schedule sign-in link',
      body: 'Hi ' + player.name + ',\n\nUse this one-time link to sign in to the Tennis Schedule Manager. It expires in ' + APP_CONFIG.magicLinkMinutes + ' minutes.\n\n' + link + '\n\nIf you did not request this link, you can ignore this email.',
      name: APP_CONFIG.title
    });
    cache.put(rateKey, '1', 60);
  } catch (error) {
    cache.remove(cacheKey);
    throw new Error('Google Apps Script could not send the email. Authorize MailApp and check the sender account quota.');
  }
}

function _publicUser_(player) {
  if (!player) throw new Error('This player is no longer active.');
  return { playerID: player.playerID, name: player.name, email: player.email, isScheduleManager: _isManager_(player.playerID) };
}

function _getSession_(token) {
  const raw = String(token || '');
  if (!raw) throw new Error('Sign in with your email link to continue.');
  const value = CacheService.getScriptCache().get('session_' + _hash_(raw));
  if (!value) throw new Error('Your sign-in session expired. Request a new link.');
  const session = JSON.parse(value);
  if (session.expiresAt < Date.now()) throw new Error('Your sign-in session expired. Request a new link.');
  const player = _findPlayerById_(session.playerID);
  if (!player) throw new Error('This player is no longer active.');
  return session;
}

function _requireManager_(token) {
  const session = _getSession_(token);
  if (!_isManager_(session.playerID)) throw new Error('Only the Schedule Manager can do that.');
  return session;
}

function _isManager_(playerID) {
  return String(_getSetting_('ManagerPlayerID') || '') === String(playerID || '');
}

function _getBook_() {
  const id = PropertiesService.getScriptProperties().getProperty('DATA_SPREADSHEET_ID');
  if (!id) throw new Error('Run setupProject from the Apps Script editor first.');
  return SpreadsheetApp.openById(id);
}

function _ensureSheet_(book, name, headers) {
  let sheet = book.getSheetByName(name);
  if (!sheet) sheet = book.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
}

function _readPlayers_() {
  const sheet = _getBook_().getSheetByName(APP_CONFIG.playersSheet);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues().map((row, index) => ({
    playerID: String(row[0] || '').trim(), name: String(row[1] || '').trim(), email: _cleanEmail_(row[2]),
    active: row[3] !== false && String(row[3]).toUpperCase() !== 'FALSE', row: index + 2
  })).filter(player => player.playerID && player.name && player.email);
}

function _findPlayerById_(playerID, includeInactive) {
  return _readPlayers_().find(player => player.playerID === String(playerID || '') && (includeInactive || player.active)) || null;
}

function _findPlayerByEmail_(email, includeInactive) {
  const clean = _cleanEmail_(email);
  return _readPlayers_().find(player => player.email === clean && (includeInactive || player.active)) || null;
}

function _validatePlayer_(player) {
  const value = { playerID: String(player.playerID || '').trim(), name: String(player.name || '').trim(), email: _cleanEmail_(player.email) };
  if (!value.playerID || !value.name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.email)) throw new Error('Enter a PlayerID, name, and valid email.');
  return value;
}

function _cleanEmail_(email) { return String(email || '').trim().toLowerCase(); }

function _readSchedule_() {
  const sheet = _getBook_().getSheetByName(APP_CONFIG.scheduleSheet);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().map(row => ({
    date: row[0] instanceof Date ? _formatDate_(row[0]) : String(row[0] || ''),
    playerID: String(row[1] || ''), status: String(row[2] || '')
  })).filter(item => item.date && item.playerID && item.status);
}

function _getStatus_(playerID, date) {
  const match = _readSchedule_().find(item => item.playerID === playerID && item.date === date);
  return match ? match.status : '';
}

function _setStatus_(playerID, date, status) {
  const sheet = _getBook_().getSheetByName(APP_CONFIG.scheduleSheet);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const count = Math.max(0, sheet.getLastRow() - 1);
    if (count) {
      const values = sheet.getRange(2, 1, count, 3).getValues();
      for (let index = 0; index < values.length; index++) {
        const rowDate = values[index][0] instanceof Date ? _formatDate_(values[index][0]) : String(values[index][0] || '');
        if (String(values[index][1]) === playerID && rowDate === date) {
          if (status) sheet.getRange(index + 2, 3).setValue(status);
          else sheet.deleteRow(index + 2);
          return;
        }
      }
    }
    if (status) sheet.appendRow([date, playerID, status]);
  } finally {
    lock.releaseLock();
  }
}

function _getPlayDays_() {
  try { return JSON.parse(_getSetting_('PlayDays') || 'null') || APP_CONFIG.defaultPlayDays; }
  catch (error) { return APP_CONFIG.defaultPlayDays; }
}

function _getScheduleDates_() {
  try {
    const dates = JSON.parse(_getSetting_('ScheduleDates') || '[]');
    if (Array.isArray(dates) && dates.length) return dates.map(String).sort();
  } catch (error) {}
  const dates = _generateDates_(APP_CONFIG.scheduleSessionCount, _getPlayDays_());
  _setSetting_('ScheduleDates', JSON.stringify(dates));
  return dates;
}

function _generateDates_(count, days) {
  return _generateAfter_(count, new Date(), days);
}

function _generateAfter_(count, afterDate, days) {
  const allowed = days && days.length ? days : APP_CONFIG.defaultPlayDays;
  const cursor = new Date(afterDate.getFullYear(), afterDate.getMonth(), afterDate.getDate());
  const dates = [];
  while (dates.length < count) {
    cursor.setDate(cursor.getDate() + 1);
    if (allowed.indexOf(cursor.getDay()) !== -1) dates.push(_formatDate_(cursor));
  }
  return dates;
}

function _parseDate_(value) {
  const parts = String(value).split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function _formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function _getSetting_(key) {
  const sheet = _getBook_().getSheetByName(APP_CONFIG.settingsSheet);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (let index = 0; index < rows.length; index++) if (String(rows[index][0]) === key) return String(rows[index][1] == null ? '' : rows[index][1]);
  return null;
}

function _setSetting_(key, value) {
  const sheet = _getBook_().getSheetByName(APP_CONFIG.settingsSheet);
  if (!sheet) throw new Error('Run setupProject from the Apps Script editor first.');
  const count = Math.max(0, sheet.getLastRow() - 1);
  if (count) {
    const rows = sheet.getRange(2, 1, count, 2).getValues();
    for (let index = 0; index < rows.length; index++) {
      if (String(rows[index][0]) === key) {
        sheet.getRange(index + 2, 2).setValue(value);
        return;
      }
    }
  }
  sheet.appendRow([key, value]);
}

function _newToken_() { return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); }

function _hash_(value) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(digest).replace(/=+$/, '');
}

function _requireOwnerEditor_() {
  const active = Session.getActiveUser().getEmail();
  const effective = Session.getEffectiveUser().getEmail();
  if (!active || !effective || active.toLowerCase() !== effective.toLowerCase()) {
    throw new Error('Run setupProject from the Apps Script editor while signed in as the project owner.');
  }
}
