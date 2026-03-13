// ═══════════════════════════════════════════════════════
// NDS Campaign Dashboard — Google Apps Script Middleware
// ═══════════════════════════════════════════════════════
// Separate Code.gs for AT&T NDS offices (Air + Cell only).
// Accepts officeId + sheetId params to route to correct tabs/sheets.
// Deploy as Web App: Execute as ME, Anyone can access.
// Set API_KEY in Script Properties (Project Settings > Script Properties).


// === PER-OFFICE TAB BASES ===
const TAB = {
  SALES: '_Sales',
  ROSTER: '_Roster',
  TEAMS: '_Teams',
  TEAM_CUSTOM: '_TeamCustom',
  OVERRIDES: '_Overrides',
  UNLOCKS: '_Unlocks',
  SETTINGS: '_Settings'
};

// Shared tabs (no officeId suffix — one copy per campaign sheet)
const TABLEAU_TAB = '_TableauOrderLog';
const CHURN_REPORT_TAB = '_TableauChurnReport';

// Build per-office tab name: _Sales + _off_001 = _Sales_off_001
function officeTab(base, officeId) {
  return base + '_' + officeId;
}

// Default officeId when none is provided
const DEFAULT_OFFICE_ID = 'off_006';


// === Build team emoji maps dynamically from _Teams tab ===

function buildTeamEmojiMaps(ss, officeId) {
  var emojiMap = {};   // emoji → name
  var nameMap = {};    // name → emoji
  var sheet = ss.getSheetByName(officeTab(TAB.TEAMS, officeId));
  if (!sheet) return { emojiMap: emojiMap, nameMap: nameMap };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][1] || '').trim();
    var emoji = String(data[i][4] || '').trim();
    if (name && emoji) {
      emojiMap[emoji] = name;
      nameMap[name] = emoji;
    }
  }
  return { emojiMap: emojiMap, nameMap: nameMap };
}

// Ranks that appear in the "Leaders" section of the leaderboard
const LEADER_RANKS = ['owner', 'manager', 'jd', 'l1'];


// === _Sales column indices (0-based) ===
// NDS simplified schema: no Fiber, VoIP, DTV columns

const OL = {
  TIMESTAMP: 0,
  EMAIL: 1,
  REP_NAME: 2,
  DATE_OF_SALE: 3,
  CAMPAIGN: 4,
  DSI: 5,  // SPM in UI — kept as DSI internally for backward compat
  ACCOUNT_TYPE: 6,
  CLIENT_NAME: 7,
  TRAINEE: 8,
  TRAINEE_NAME: 9,
  AIR: 10,
  NEW_PHONES: 11,
  BYODS: 12,
  CELL: 13,
  OOMA_PACKAGE: 14,
  ACCOUNT_NOTES: 15,
  ACTIVATION_SUPPORT: 16,
  TEAM_EMOJI: 17,
  YESES: 18,
  UNITS: 19,
  STATUS: 20,
  NOTES: 21,
  PAID_OUT: 22,
  TICKETS: 23,
  ORDER_CHANNEL: 24,
  CODES_USED_BY: 25
};


// === _TableauOrderLog — header-based column lookup ===

const TOL_HEADER_MAP = {
  'owner & office': 'OWNER_OFFICE',
  'owner and office': 'OWNER_OFFICE',
  'rep': 'REP',
  'icd.lead rep id': 'LEAD_REP_ID',
  'lead rep id': 'LEAD_REP_ID',
  'rep.rep number': 'REP_NUMBER',
  'rep number': 'REP_NUMBER',
  'sp.order date (copy)': 'ORDER_DATE',
  'order date (copy)': 'ORDER_DATE',
  'order date': 'ORDER_DATE',
  'order time (timezone)': 'ORDER_TIME',
  'order time': 'ORDER_TIME',
  'sp.spm number': 'SPM',
  'spm number': 'SPM',
  'spm': 'SPM',
  'dsi': 'SPM',
  'spe.name': 'SPE',
  'spe name': 'SPE',
  'spe': 'SPE',
  'spe.account ban': 'BAN',
  'account ban': 'BAN',
  'ban': 'BAN',
  'product type (broken out)': 'PRODUCT_TYPE',
  'product type': 'PRODUCT_TYPE',
  'cru/iru': 'CRU_IRU',
  'cru / iru': 'CRU_IRU',
  'dtr status (enriched)': 'DTR_STATUS',
  'dtr status': 'DTR_STATUS',
  'disconnect reason (consolidated)': 'DISCO_REASON',
  'disconnect reason': 'DISCO_REASON',
  'disco reason': 'DISCO_REASON',
  'spe.port carrier': 'PORT_CARRIER',
  'port carrier': 'PORT_CARRIER',
  'notes.note': 'NOTES',
  'notes': 'NOTES',
  'dtr status date': 'DTR_STATUS_DATE',
  'order status': 'ORDER_STATUS',
  'spe.dtr posted date (copy)': 'POSTED_DATE',
  'posted date (copy)': 'POSTED_DATE',
  'posted date': 'POSTED_DATE',
  'max posted': 'MAX_POSTED',
  'first streaming date': 'FIRST_STREAMING',
  'first streaming': 'FIRST_STREAMING',
  'voice line count': 'VOICE_LINE_COUNT',
  'spe.tn type': 'TN_TYPE',
  'tn type': 'TN_TYPE',
  'spe.phone': 'PHONE',
  'phone': 'PHONE',
  'spe.install date': 'INSTALL_DATE',
  'install date': 'INSTALL_DATE',
  'b2b rep volume bonus tiers': 'BONUS_TIERS',
  'bonus tiers': 'BONUS_TIERS',
  'tier bonus payout/dnq reason': 'PAYOUT_REASON',
  'payout reason': 'PAYOUT_REASON',
  'unit count': 'UNIT_COUNT',
  'total volume': 'TOTAL_VOLUME',
  'total activations': 'TOTAL_ACTS',
  'total acts': 'TOTAL_ACTS'
};

// Build column index map from header row
function buildTableauColumnMap(headerRow) {
  var col = {};
  for (var i = 0; i < headerRow.length; i++) {
    var raw = String(headerRow[i] || '').trim().toLowerCase();
    var key = TOL_HEADER_MAP[raw];
    if (key && !col.hasOwnProperty(key)) {
      col[key] = i;
    }
  }
  return col;
}

// Safe getter — returns '' for missing columns instead of crashing
function tCol(row, col, key) {
  return col.hasOwnProperty(key) ? row[col[key]] : '';
}


// === UTILITIES ===

function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('API_KEY') || '';
}

function validateKey(key) {
  const expected = getApiKey();
  if (!expected) return true;
  return key === expected;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Shard-ready: open correct spreadsheet based on sheetId param
function getSheet(params) {
  var sheetId = (params && params.sheetId) || '';
  if (sheetId) {
    return SpreadsheetApp.openById(sheetId);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// Create or get a per-office tab, using baseName for header template
function getOrCreateSheet(ss, tabName, baseName) {
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    switch (baseName) {
      case TAB.SALES:
        sheet.appendRow([
          'Timestamp', 'Email', 'Rep Name', 'Date of Sale', 'Campaign',
          'SPM', 'Account Type', 'Client Name', 'Trainee', 'Trainee Name',
          'Air', 'New Phones', 'BYODs', 'Cell',
          'Ooma Package', 'Account Notes', 'Activation Support', 'Team Emoji',
          'Yeses', 'Units', 'Status', 'Notes', 'Paid Out', 'Tickets',
          'Order Channel', 'Codes Used By'
        ]);
        break;
      case TAB.ROSTER:
        sheet.appendRow(['email', 'name', 'team', 'rank', 'deactivated', 'dateAdded', 'pinHash', 'phone', 'tableauName']);
        break;
      case TAB.OVERRIDES:
        sheet.appendRow(['key', 'product', 'status', 'date', 'order', 'notes_json']);
        break;
      case TAB.TEAM_CUSTOM:
        sheet.appendRow(['persona', 'emoji', 'displayName']);
        break;
      case TAB.UNLOCKS:
        sheet.appendRow(['persona', 'status']);
        break;
      case TAB.TEAMS:
        sheet.appendRow(['teamId', 'name', 'parentId', 'leaderId', 'emoji', 'createdDate']);
        break;
      case TAB.SETTINGS:
        sheet.appendRow(['key', 'value']);
        break;
    }
  }
  return sheet;
}

function findRow(sheet, colIndex, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]).trim() === String(value).trim()) {
      return i + 1;
    }
  }
  return -1;
}

// Case-insensitive findRow for email matching
function findRowCI(sheet, colIndex, value) {
  const data = sheet.getDataRange().getValues();
  const target = String(value).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]).trim().toLowerCase() === target) {
      return i + 1;
    }
  }
  return -1;
}

// Hash a PIN with email as salt → hex string
function hashPin(email, pin) {
  var input = String(email).trim().toLowerCase() + ':' + String(pin).trim();
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  return digest.map(function(b) {
    return ('0' + ((b + 256) % 256).toString(16)).slice(-2);
  }).join('');
}


// === doGet() — READ ALL DATA ===

function doGet(e) {
  const key = (e && e.parameter && e.parameter.key) || '';
  if (!validateKey(key)) {
    return jsonResponse({ error: 'unauthorized' });
  }

  try {
    const action = (e && e.parameter && e.parameter.action) || '';
    const officeId = (e && e.parameter && e.parameter.officeId) || DEFAULT_OFFICE_ID;
    const ss = getSheet(e && e.parameter);

    // Orders-specific request
    if (action === 'readOrders') {
      const filterEmail = (e.parameter && e.parameter.email) || '';
      const orders = readOrders(ss, officeId, filterEmail || null);
      return jsonResponse({ orders: orders });
    }

    // Payroll orders — filtered by payrollMode (commission-split or flat-rate)
    if (action === 'readPayrollOrders') {
      var payrollMode = (e.parameter && e.parameter.payrollMode) || 'commission-split';
      const orders = readPayrollOrders(ss, officeId, payrollMode);
      return jsonResponse({ orders: orders });
    }

    // Tableau summary (shared tab, but needs officeId for DSI→email mapping)
    if (action === 'readTableauSummary') {
      return jsonResponse(getTableauSummaryWithCache(ss, officeId));
    }

    // Tableau device detail for a single DSI (shared tab)
    if (action === 'readTableauDetail') {
      const dsi = (e.parameter && e.parameter.dsi) || '';
      return jsonResponse({ devices: readTableauDetail(ss, dsi) });
    }

    // Leaderboard snapshot (used by Make.com for daily Discord post)
    if (action === 'leaderboard') {
      return jsonResponse(readLeaderboard(ss, officeId));
    }

    // Leaderboard as pre-built HTML (for HTML-to-Image rendering)
    if (action === 'leaderboardHtml') {
      var lb = readLeaderboard(ss, officeId);
      return jsonResponse({ html: buildLeaderboardHtml(lb) });
    }

    // Leaderboard as emoji text (called by centralized AdminCode scheduler)
    if (action === 'leaderboardText') {
      var officeName = (e.parameter && e.parameter.officeName) || 'OFFICE';
      return jsonResponse({ text: buildLeaderboardText(ss, officeId, officeName) });
    }

    // Default: return full dashboard data (includes Tableau summary)

    // Auto-add office owner to roster if not already present
    var ownerEmail = (e.parameter && e.parameter.ownerEmail || '').trim().toLowerCase();
    var ownerName = (e.parameter && e.parameter.ownerName || '').trim();
    if (ownerEmail) {
      var rosterSheet = ss.getSheetByName(officeTab(TAB.ROSTER, officeId));
      if (rosterSheet) {
        var rosterData = rosterSheet.getDataRange().getValues();
        var ownerFound = false;
        for (var ri = 1; ri < rosterData.length; ri++) {
          if (String(rosterData[ri][0] || '').trim().toLowerCase() === ownerEmail) {
            ownerFound = true;
            break;
          }
        }
        if (!ownerFound) {
          rosterSheet.appendRow([ownerEmail, ownerName || ownerEmail, '', 'owner', false, new Date().toISOString(), '', '', '']);
          Logger.log('[AutoRoster] Added owner to roster: ' + ownerEmail);
        }
      }
    }

    let roster = readRoster(ss, officeId);
    const teamMaps = buildTeamEmojiMaps(ss, officeId);
    const peopleResult = readPeople(ss, officeId, roster, teamMaps.nameMap);
    const tableauSummary = getTableauSummaryWithCache(ss, officeId);

    // Auto-assign Tableau names (managers first, then down the ranks)
    roster = autoAssignTableauNames(ss, officeId, roster, tableauSummary.possibleTableauNames);

    const data = {
      people: peopleResult.people || peopleResult,
      roster: roster,
      teamMap: teamMaps.emojiMap,
      teams: readTeams(ss, officeId),
      orderOverrides: readOrderOverrides(ss, officeId),
      teamCustomizations: readTeamCustomizations(ss, officeId),
      unlockRequests: readUnlockRequests(ss, officeId),
      settings: readSettings(ss, officeId),
      tableauSummary: tableauSummary,
      churnReport: readChurnReport(ss)
    };
    return jsonResponse(data);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}


// === PERIOD HELPERS (NDS: Air + Cell only) ===

function emptyPeriod() {
  return { y: 0, air: 0, cell: 0, units: 0 };
}

function addToPeriod(target, sale) {
  target.y += sale.y;
  target.air += sale.air;
  target.cell += sale.cell;
  target.units += sale.units;
}

function sumDays(days) {
  const r = emptyPeriod();
  days.forEach(d => addToPeriod(r, d));
  return r;
}

function sumAllPeriods(agg) {
  const r = emptyPeriod();
  agg.days.forEach(d => addToPeriod(r, d));
  addToPeriod(r, agg.priorWeek);
  addToPeriod(r, agg.twoWkPrior);
  addToPeriod(r, agg.threeWkPrior);
  addToPeriod(r, agg.fourWkPrior);
  addToPeriod(r, agg.fiveWkPrior);
  return r;
}

function getWeekStart() {
  const now = new Date();
  const dow = now.getDay(); // 0=Sun
  const daysFromMon = dow === 0 ? 6 : dow - 1;
  const mon = new Date(now);
  mon.setDate(now.getDate() - daysFromMon);
  mon.setHours(0, 0, 0, 0);
  return mon;
}


// === READ HELPERS ===

function readRoster(ss, officeId) {
  const sheet = ss.getSheetByName(officeTab(TAB.ROSTER, officeId));
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][0] || '').trim().toLowerCase();
    if (!email) continue;
    var pinVal = String(data[i][6] || '').trim();
    result[email] = {
      name: String(data[i][1] || '').trim(),
      team: String(data[i][2] || '').trim(),
      rank: String(data[i][3] || 'rep').trim(),
      deactivated: data[i][4] === true || String(data[i][4]).toUpperCase() === 'TRUE',
      dateAdded: data[i][5] || '',
      hasPin: pinVal.length > 0 && pinVal !== 'undefined',
      phone: String(data[i][7] || '').trim(),
      tableauName: String(data[i][8] || '').trim()
    };
  }
  return result;
}

function readPeople(ss, officeId, roster, teamNameToEmoji) {
  const olSheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!olSheet) return { people: [], _debug: { error: 'No Sales sheet' } };
  if (!roster || Object.keys(roster).length === 0) return { people: [], _debug: { error: 'Empty roster' } };

  const olData = olSheet.getDataRange().getValues();
  const hasSalesData = olData.length >= 2;

  var _dbg = {
    totalRows: olData.length - 1,
    noEmail: 0,
    emailNotInRoster: 0,
    noDate: 0,
    badDate: 0,
    tooOld: 0,
    matched: 0,
    headerRow: olData[0].slice(0, 10).map(String),
    sampleRow2: olData.length > 1 ? [
      'col0=' + String(olData[1][0]),
      'col1(email)=' + String(olData[1][OL.EMAIL]),
      'col3(date)=' + String(olData[1][OL.DATE_OF_SALE]),
      'col2(name)=' + String(olData[1][OL.REP_NAME]),
      'col18(yeses)=' + String(olData[1][OL.YESES]),
      'col19(units)=' + String(olData[1][OL.UNITS])
    ] : [],
    rosterEmails: Object.keys(roster).slice(0, 5),
    olEmails: []
  };

  var seenEmails = {};
  for (var d = 1; d < olData.length && _dbg.olEmails.length < 5; d++) {
    var de = String(olData[d][OL.EMAIL] || '').trim().toLowerCase();
    if (de && !seenEmails[de]) { seenEmails[de] = true; _dbg.olEmails.push(de); }
  }

  // Week boundaries
  const thisWeekStart = getWeekStart();
  const DAY_MS = 86400000;
  const priorWeekStart = new Date(thisWeekStart.getTime() - 7 * DAY_MS);
  const twoWkStart = new Date(thisWeekStart.getTime() - 14 * DAY_MS);
  const threeWkStart = new Date(thisWeekStart.getTime() - 21 * DAY_MS);
  const fourWkStart = new Date(thisWeekStart.getTime() - 28 * DAY_MS);
  const fiveWkStart = new Date(thisWeekStart.getTime() - 35 * DAY_MS);

  _dbg.thisWeekStart = thisWeekStart.toISOString();
  _dbg.fiveWkStart = fiveWkStart.toISOString();

  // Initialize aggregation buckets for each rostered person
  const agg = {};
  Object.keys(roster).forEach(email => {
    agg[email] = {
      days: Array.from({ length: 7 }, () => emptyPeriod()),
      lwDays: Array.from({ length: 7 }, () => emptyPeriod()),
      w2Days: Array.from({ length: 7 }, () => emptyPeriod()),
      w3Days: Array.from({ length: 7 }, () => emptyPeriod()),
      w4Days: Array.from({ length: 7 }, () => emptyPeriod()),
      w5Days: Array.from({ length: 7 }, () => emptyPeriod()),
      priorWeek: emptyPeriod(),
      twoWkPrior: emptyPeriod(),
      threeWkPrior: emptyPeriod(),
      fourWkPrior: emptyPeriod(),
      fiveWkPrior: emptyPeriod(),
      recentTime: [0, 0, 0, 0],
      fw4Time: [0, 0, 0, 0]
    };
  });

  // Process each sale row (skip header at index 0)
  if (hasSalesData) for (let i = 1; i < olData.length; i++) {
    const row = olData[i];
    const email = String(row[OL.EMAIL] || '').trim().toLowerCase();
    if (!email) { _dbg.noEmail++; continue; }

    // Skip sales from people not in the roster
    if (!roster[email]) { _dbg.emailNotInRoster++; continue; }

    // Skip Tower orders — tracked but excluded from leaderboard
    var orderChannel = String(row[OL.ORDER_CHANNEL] || 'Sara').trim();
    if (orderChannel === 'Tower') continue;

    // Parse date
    let rawDate = row[OL.DATE_OF_SALE];
    if (!rawDate) { _dbg.noDate++; continue; }
    let saleDate = new Date(rawDate);
    if (isNaN(saleDate.getTime())) { _dbg.badDate++; continue; }
    saleDate.setHours(0, 0, 0, 0);

    // Extract sale counters (NDS: Air + Cell only)
    const sale = {
      y:     Number(row[OL.YESES]) || 0,
      air:   Number(row[OL.AIR]) || 0,
      cell:  Number(row[OL.CELL]) || 0,
      units: Number(row[OL.UNITS]) || 0
    };

    const pa = agg[email];

    var weekOffset = -1;
    if (saleDate >= thisWeekStart) {
      weekOffset = 0;
      const dayIdx = Math.floor((saleDate.getTime() - thisWeekStart.getTime()) / DAY_MS);
      if (dayIdx >= 0 && dayIdx < 7) {
        addToPeriod(pa.days[dayIdx], sale);
        _dbg.matched++;
      }
    } else if (saleDate >= priorWeekStart) {
      weekOffset = 1;
      addToPeriod(pa.priorWeek, sale);
      var lwDayIdx = Math.floor((saleDate.getTime() - priorWeekStart.getTime()) / DAY_MS);
      if (lwDayIdx >= 0 && lwDayIdx < 7) addToPeriod(pa.lwDays[lwDayIdx], sale);
      _dbg.matched++;
    } else if (saleDate >= twoWkStart) {
      weekOffset = 2;
      addToPeriod(pa.twoWkPrior, sale);
      var w2DayIdx = Math.floor((saleDate.getTime() - twoWkStart.getTime()) / DAY_MS);
      if (w2DayIdx >= 0 && w2DayIdx < 7) addToPeriod(pa.w2Days[w2DayIdx], sale);
      _dbg.matched++;
    } else if (saleDate >= threeWkStart) {
      weekOffset = 3;
      addToPeriod(pa.threeWkPrior, sale);
      var w3DayIdx = Math.floor((saleDate.getTime() - threeWkStart.getTime()) / DAY_MS);
      if (w3DayIdx >= 0 && w3DayIdx < 7) addToPeriod(pa.w3Days[w3DayIdx], sale);
      _dbg.matched++;
    } else if (saleDate >= fourWkStart) {
      weekOffset = 4;
      addToPeriod(pa.fourWkPrior, sale);
      var w4DayIdx = Math.floor((saleDate.getTime() - fourWkStart.getTime()) / DAY_MS);
      if (w4DayIdx >= 0 && w4DayIdx < 7) addToPeriod(pa.w4Days[w4DayIdx], sale);
      _dbg.matched++;
    } else if (saleDate >= fiveWkStart) {
      weekOffset = 5;
      addToPeriod(pa.fiveWkPrior, sale);
      var w5DayIdx = Math.floor((saleDate.getTime() - fiveWkStart.getTime()) / DAY_MS);
      if (w5DayIdx >= 0 && w5DayIdx < 7) addToPeriod(pa.w5Days[w5DayIdx], sale);
      _dbg.matched++;
    } else {
      _dbg.tooOld++;
    }

    // Time-of-sale bucketing
    if (weekOffset >= 0) {
      var ts = row[OL.TIMESTAMP];
      if (ts instanceof Date) {
        var h = ts.getHours() + ts.getMinutes() / 60;
        var slotIdx = -1;
        if (h >= 10.5 && h < 16)  slotIdx = 0;
        else if (h >= 16 && h < 18) slotIdx = 1;
        else if (h >= 18 && h < 21) slotIdx = 2;
        else if (h >= 21 || h < 10.5) slotIdx = 3;
        if (slotIdx >= 0) {
          if (weekOffset <= 1) pa.recentTime[slotIdx]++;
          if (weekOffset >= 2 && weekOffset <= 5) pa.fw4Time[slotIdx]++;
        }
      }
    }
  }

  // Build person response objects
  const people = [];
  Object.entries(roster).forEach(([email, info]) => {
    const pa = agg[email];
    const type = LEADER_RANKS.includes(info.rank) ? 'leader' : 'rep';
    const teamEmoji = (teamNameToEmoji || {})[info.team] || '';

    people.push({
      name: info.name,
      type: type,
      email: email,
      rank: info.rank,
      teamEmoji: teamEmoji,
      team: info.team,
      deactivated: info.deactivated || false,
      days: pa.days,
      lwDays: pa.lwDays,
      w2Days: pa.w2Days,
      w3Days: pa.w3Days,
      w4Days: pa.w4Days,
      w5Days: pa.w5Days,
      thisWeek: sumDays(pa.days),
      priorWeek: pa.priorWeek,
      twoWkPrior: pa.twoWkPrior,
      threeWkPrior: pa.threeWkPrior,
      fourWkPrior: pa.fourWkPrior,
      fiveWkPrior: pa.fiveWkPrior,
      fourWkRunning: sumAllPeriods(pa),
      recentTime: pa.recentTime,
      fw4Time: pa.fw4Time
    });
  });

  return { people: people, _debug: _dbg };
}


// === readOrders() — Individual order rows for past 30 days ===

function readOrders(ss, officeId, filterEmail) {
  const olSheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!olSheet) return [];

  const olData = olSheet.getDataRange().getValues();
  if (olData.length < 2) return [];

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 30);
  cutoff.setHours(0, 0, 0, 0);

  const targetEmail = filterEmail ? String(filterEmail).trim().toLowerCase() : null;
  const orders = [];

  for (let i = 1; i < olData.length; i++) {
    const row = olData[i];
    const email = String(row[OL.EMAIL] || '').trim().toLowerCase();
    if (!email) continue;
    if (targetEmail && email !== targetEmail) continue;

    const rawDate = row[OL.DATE_OF_SALE];
    if (!rawDate) continue;
    const saleDate = new Date(rawDate);
    if (isNaN(saleDate.getTime())) continue;
    saleDate.setHours(0, 0, 0, 0);
    if (saleDate < cutoff) continue;

    orders.push({
      rowIndex: i + 1,
      email: email,
      repName: String(row[OL.REP_NAME] || '').trim(),
      dsi: String(row[OL.DSI] || '').trim(),
      dateOfSale: saleDate.toISOString().split('T')[0],
      campaign: String(row[OL.CAMPAIGN] || '').trim(),
      accountType: String(row[OL.ACCOUNT_TYPE] || '').trim(),
      clientName: String(row[OL.CLIENT_NAME] || '').trim(),
      air:   Number(row[OL.AIR]) || 0,
      newPhones: Number(row[OL.NEW_PHONES]) || 0,
      byods: Number(row[OL.BYODS]) || 0,
      cell:  Number(row[OL.CELL]) || 0,
      units: Number(row[OL.UNITS]) || 0,
      status: String(row[OL.STATUS] || 'Pending').trim(),
      notes:  String(row[OL.NOTES] || '').trim(),
      tickets: (function() { try { return JSON.parse(row[OL.TICKETS] || '[]'); } catch(e) { return []; } })(),
      orderChannel: String(row[OL.ORDER_CHANNEL] || 'Sara').trim(),
      codesUsedBy: String(row[OL.CODES_USED_BY] || '').trim().toLowerCase()
    });
  }

  orders.sort((a, b) => b.dateOfSale.localeCompare(a.dateOfSale));
  return orders;
}


// === readPayrollOrders() — Trainee orders past 2 months ===

function readPayrollOrders(ss, officeId, payrollMode) {
  const olSheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!olSheet) return [];

  const olData = olSheet.getDataRange().getValues();
  if (olData.length < 2) return [];

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 60);
  cutoff.setHours(0, 0, 0, 0);

  var mode = String(payrollMode || 'commission-split').trim();
  const orders = [];

  for (let i = 1; i < olData.length; i++) {
    const row = olData[i];
    const email = String(row[OL.EMAIL] || '').trim().toLowerCase();
    if (!email) continue;

    const trainee = String(row[OL.TRAINEE] || '').trim().toLowerCase();
    const codesUsedBy = String(row[OL.CODES_USED_BY] || '').trim().toLowerCase();
    var isTraineeOrder = (trainee === 'yes');
    var isCodesSwap = (codesUsedBy !== '');

    if (mode === 'flat-rate') {
      if (!isCodesSwap) continue;
    } else {
      if (!isTraineeOrder && !isCodesSwap) continue;
    }

    const rawDate = row[OL.DATE_OF_SALE];
    if (!rawDate) continue;
    const saleDate = new Date(rawDate);
    if (isNaN(saleDate.getTime())) continue;
    saleDate.setHours(0, 0, 0, 0);
    if (saleDate < cutoff) continue;

    let paidOut = {};
    try {
      const rawPaid = String(row[OL.PAID_OUT] || '').trim();
      if (rawPaid) paidOut = JSON.parse(rawPaid);
    } catch (e) { paidOut = {}; }

    orders.push({
      rowIndex: i + 1,
      email: email,
      repName: String(row[OL.REP_NAME] || '').trim(),
      traineeName: String(row[OL.TRAINEE_NAME] || '').trim(),
      dsi: String(row[OL.DSI] || '').trim(),
      dateOfSale: saleDate.toISOString().split('T')[0],
      air:   Number(row[OL.AIR]) || 0,
      cell:  Number(row[OL.CELL]) || 0,
      units: Number(row[OL.UNITS]) || 0,
      status: String(row[OL.STATUS] || 'Pending').trim(),
      notes:  String(row[OL.NOTES] || '').trim(),
      paidOut: paidOut,
      orderChannel: String(row[OL.ORDER_CHANNEL] || 'Sara').trim(),
      codesUsedBy: codesUsedBy
    });
  }

  // Enrich payroll orders with Tableau SPE data
  var tableauSummary = getTableauSummaryWithCache(ss, officeId);
  var dsiSummary = tableauSummary.dsiSummary || {};
  orders.forEach(function(order) {
    var ts = dsiSummary[order.dsi];
    if (ts) {
      order.speList = ts.speList || [];
      order.tableauStatusCounts = ts.statusCounts || {};
    }
  });

  orders.sort((a, b) => b.dateOfSale.localeCompare(a.dateOfSale));
  return orders;
}


// === readSettings() — Key-value settings from per-office _Settings tab ===

function readSettings(ss, officeId) {
  const sheet = ss.getSheetByName(officeTab(TAB.SETTINGS, officeId));
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    if (key) result[key] = String(data[i][1] || '').trim();
  }
  return result;
}


// === readChurnReport() — Read shared _TableauChurnReport tab ===

function readChurnReport(ss) {
  const sheet = ss.getSheetByName(CHURN_REPORT_TAB);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const rawHeaders = data[0].map((h) => String(h).trim());

  var metricTypeCol = -1;
  var KNOWN_METRICS = ['Activated SPE/SP', 'Disconnect count (SPE/SP)', 'Churn Rate'];
  for (var ri = 1; ri < Math.min(data.length, 20); ri++) {
    for (var ci = 0; ci < data[ri].length; ci++) {
      var cellVal = String(data[ri][ci] || '').trim();
      if (KNOWN_METRICS.indexOf(cellVal) !== -1) {
        metricTypeCol = ci;
        break;
      }
    }
    if (metricTypeCol !== -1) break;
  }

  const headers = rawHeaders.map((h, j) => {
    if (j === metricTypeCol) return 'metricType';
    return h === '' ? ('_blank_' + j) : h;
  });
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, j) => {
      row[h] = data[i][j] !== undefined && data[i][j] !== null ? data[i][j] : '';
    });
    rows.push(row);
  }
  return rows;
}


// === writeSetting() — Upsert a key-value setting ===

function writeSetting(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.SETTINGS, officeId), TAB.SETTINGS);
  const key = String(body.key || '').trim();
  const value = String(body.value || '').trim();
  if (!key) return { error: 'missing key' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return { ok: true };
    }
  }
  sheet.appendRow([key, value]);
  return { ok: true };
}


function readOrderOverrides(ss, officeId) {
  const sheet = ss.getSheetByName(officeTab(TAB.OVERRIDES, officeId));
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    if (!key) continue;
    let notes = [];
    try { notes = JSON.parse(data[i][5] || '[]'); } catch (e) { notes = []; }
    result[key] = {
      product: data[i][1] || '',
      status:  data[i][2] || '',
      date:    data[i][3] || '',
      order:   data[i][4] || '',
      notes
    };
  }
  return result;
}

function readTeamCustomizations(ss, officeId) {
  const sheet = ss.getSheetByName(officeTab(TAB.TEAM_CUSTOM, officeId));
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const persona = String(data[i][0] || '').trim();
    if (!persona) continue;
    result[persona] = {
      emoji: data[i][1] || '⚡',
      name:  data[i][2] || ''
    };
  }
  return result;
}

function readUnlockRequests(ss, officeId) {
  const sheet = ss.getSheetByName(officeTab(TAB.UNLOCKS, officeId));
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const persona = String(data[i][0] || '').trim();
    if (!persona) continue;
    result[persona] = data[i][1] || 'pending';
  }
  return result;
}

function readTeams(ss, officeId) {
  const sheet = ss.getSheetByName(officeTab(TAB.TEAMS, officeId));
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const teamId = String(data[i][0] || '').trim();
    if (!teamId) continue;
    result[teamId] = {
      teamId: teamId,
      name: String(data[i][1] || '').trim(),
      parentId: String(data[i][2] || '').trim(),
      leaderId: String(data[i][3] || '').trim(),
      emoji: String(data[i][4] || '').trim(),
      createdDate: data[i][5] || ''
    };
  }
  return result;
}


// === TABLEAU ORDER LOG ===

function buildDsiEmailMap(ss, officeId) {
  var olSheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!olSheet) return { dsiToEmail: {}, emailToDsis: {} };
  var olData = olSheet.getDataRange().getValues();
  var dsiToEmail = {};
  var emailToDsis = {};
  for (var i = 1; i < olData.length; i++) {
    var dsi = String(olData[i][OL.DSI] || '').trim();
    var email = String(olData[i][OL.EMAIL] || '').trim().toLowerCase();
    if (dsi && email) {
      if (!dsiToEmail[dsi]) dsiToEmail[dsi] = email;
      if (!emailToDsis[email]) emailToDsis[email] = {};
      emailToDsis[email][dsi] = true;
    }
  }
  return { dsiToEmail: dsiToEmail, emailToDsis: emailToDsis };
}

function readTableauSummary(ss, officeId) {
  var sheet = ss.getSheetByName(TABLEAU_TAB);
  if (!sheet) return { dsiSummary: {}, repSummary: {} };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { dsiSummary: {}, repSummary: {} };

  var col = buildTableauColumnMap(data[0]);
  var maps = buildDsiEmailMap(ss, officeId);
  var dsiEmailMap = maps.dsiToEmail;
  var emailToDsis = maps.emailToDsis;

  var thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  thirtyDaysAgo.setHours(0, 0, 0, 0);

  var dsiSummary = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dsi = String(tCol(row, col, 'DSI') || '').trim();
    if (!dsi) continue;

    var ownerOffice = String(tCol(row, col, 'OWNER_OFFICE') || '').trim();
    if (ownerOffice.toLowerCase() === 'total' || dsi.toLowerCase() === 'total') continue;

    var spe = String(tCol(row, col, 'SPE') || '').trim();
    var productType = String(tCol(row, col, 'PRODUCT_TYPE') || '').trim();
    var dtrStatus = String(tCol(row, col, 'DTR_STATUS') || '').trim();
    var discoReason = String(tCol(row, col, 'DISCO_REASON') || '').trim();
    var totalVolume = Number(tCol(row, col, 'TOTAL_VOLUME')) || 0;
    var totalActs = Number(tCol(row, col, 'TOTAL_ACTS')) || 0;
    var tableauRep = String(tCol(row, col, 'REP') || '').trim();
    var orderStatus = String(tCol(row, col, 'ORDER_STATUS') || '').trim();
    var orderDate = tCol(row, col, 'ORDER_DATE');

    if (!dsiSummary[dsi]) {
      dsiSummary[dsi] = {
        tableauRep: tableauRep,
        totalDevices: 0,
        totalActivations: 0,
        totalVolume: 0,
        statusCounts: {},
        productCounts: {},
        disconnectReasons: {},
        speList: [],
        devices: [],
        monthWirelessSPEs: {}
      };
    }

    var s = dsiSummary[dsi];
    s.totalDevices++;
    s.totalActivations += totalActs;
    s.totalVolume += totalVolume;

    if (dtrStatus) s.statusCounts[dtrStatus] = (s.statusCounts[dtrStatus] || 0) + 1;
    if (productType) s.productCounts[productType] = (s.productCounts[productType] || 0) + 1;
    if (discoReason) s.disconnectReasons[discoReason] = (s.disconnectReasons[discoReason] || 0) + 1;
    if (spe) {
      s.speList.push(spe);
      var ptUpper = productType.toUpperCase();
      var isWireless = ptUpper === 'WIRELESS';
      var isInMonth = orderDate instanceof Date && orderDate >= thirtyDaysAgo;
      if (isInMonth && isWireless) {
        s.monthWirelessSPEs[spe] = { orderStatus: orderStatus.toLowerCase(), dtrStatus: dtrStatus };
      }
    }
    s.devices.push({
      spe: spe,
      productType: productType,
      cruIru: String(tCol(row, col, 'CRU_IRU') || '').trim(),
      dtrStatus: dtrStatus,
      discoReason: discoReason,
      phone: String(tCol(row, col, 'PHONE') || '').trim(),
      tnType: String(tCol(row, col, 'TN_TYPE') || '').trim(),
      orderStatus: String(tCol(row, col, 'ORDER_STATUS') || '').trim(),
      postedDate: String(tCol(row, col, 'POSTED_DATE') || '').trim(),
      installDate: String(tCol(row, col, 'INSTALL_DATE') || '').trim(),
      firstStreaming: String(tCol(row, col, 'FIRST_STREAMING') || '').trim()
    });
  }

  // Aggregate per-rep using DSI → email join
  var repSummary = {};
  Object.keys(dsiSummary).forEach(function(dsi) {
    var email = dsiEmailMap[dsi];
    if (!email) return;
    var ds = dsiSummary[dsi];
    if (!repSummary[email]) {
      repSummary[email] = {
        totalDevices: 0, totalActivations: 0, totalVolume: 0,
        statusCounts: {}, productCounts: {},
        tableauName: ds.tableauRep, monthWirelessSPEs: {}
      };
    }
    var rs = repSummary[email];
    rs.totalDevices += ds.totalDevices;
    rs.totalActivations += ds.totalActivations;
    rs.totalVolume += ds.totalVolume;
    Object.keys(ds.statusCounts).forEach(function(st) { rs.statusCounts[st] = (rs.statusCounts[st] || 0) + ds.statusCounts[st]; });
    Object.keys(ds.productCounts).forEach(function(pt) { rs.productCounts[pt] = (rs.productCounts[pt] || 0) + ds.productCounts[pt]; });
    Object.keys(ds.monthWirelessSPEs).forEach(function(spe) { rs.monthWirelessSPEs[spe] = ds.monthWirelessSPEs[spe]; });
  });

  // Convert monthWirelessSPEs to counts
  Object.keys(repSummary).forEach(function(email) {
    var rs = repSummary[email];
    var spes = Object.keys(rs.monthWirelessSPEs);
    rs.monthTotalSPEs = spes.length;
    rs.monthApprovedSPEs = 0; rs.monthPendingSPEs = 0; rs.monthCanceledSPEs = 0; rs.monthDiscoSPEs = 0;
    spes.forEach(function(spe) {
      var info = rs.monthWirelessSPEs[spe];
      if (info.orderStatus === 'approved') rs.monthApprovedSPEs++;
      else if (info.orderStatus === 'pending') rs.monthPendingSPEs++;
      else if (info.orderStatus === 'canceled' || info.orderStatus === 'cancelled') rs.monthCanceledSPEs++;
      if (info.dtrStatus === 'Disconnected') rs.monthDiscoSPEs++;
    });
    delete rs.monthWirelessSPEs;
  });

  // Build possibleTableauNames
  var possibleTableauNames = {};
  Object.keys(emailToDsis).forEach(function(email) {
    var names = {};
    Object.keys(emailToDsis[email]).forEach(function(dsi) {
      if (dsiSummary[dsi] && dsiSummary[dsi].tableauRep) names[dsiSummary[dsi].tableauRep] = true;
    });
    var nameList = Object.keys(names);
    if (nameList.length > 0) possibleTableauNames[email] = nameList;
  });

  // Build repByName
  var repByName = {};
  Object.keys(dsiSummary).forEach(function(dsi) {
    var ds = dsiSummary[dsi];
    var name = ds.tableauRep;
    if (!name) return;
    if (!repByName[name]) {
      repByName[name] = {
        totalDevices: 0, totalActivations: 0, totalVolume: 0,
        statusCounts: {}, productCounts: {}, tableauName: name, monthWirelessSPEs: {}
      };
    }
    var rn = repByName[name];
    rn.totalDevices += ds.totalDevices;
    rn.totalActivations += ds.totalActivations;
    rn.totalVolume += ds.totalVolume;
    Object.keys(ds.statusCounts).forEach(function(st) { rn.statusCounts[st] = (rn.statusCounts[st] || 0) + ds.statusCounts[st]; });
    Object.keys(ds.productCounts).forEach(function(pt) { rn.productCounts[pt] = (rn.productCounts[pt] || 0) + ds.productCounts[pt]; });
    Object.keys(ds.monthWirelessSPEs).forEach(function(spe) { rn.monthWirelessSPEs[spe] = ds.monthWirelessSPEs[spe]; });
  });

  Object.keys(repByName).forEach(function(name) {
    var rn = repByName[name];
    var spes = Object.keys(rn.monthWirelessSPEs);
    rn.monthTotalSPEs = spes.length;
    rn.monthApprovedSPEs = 0; rn.monthPendingSPEs = 0; rn.monthCanceledSPEs = 0; rn.monthDiscoSPEs = 0;
    spes.forEach(function(spe) {
      var info = rn.monthWirelessSPEs[spe];
      if (info.orderStatus === 'approved') rn.monthApprovedSPEs++;
      else if (info.orderStatus === 'pending') rn.monthPendingSPEs++;
      else if (info.orderStatus === 'canceled' || info.orderStatus === 'cancelled') rn.monthCanceledSPEs++;
      if (info.dtrStatus === 'Disconnected') rn.monthDiscoSPEs++;
    });
    delete rn.monthWirelessSPEs;
  });

  return {
    dsiSummary: dsiSummary,
    repSummary: repSummary,
    repByName: repByName,
    possibleTableauNames: possibleTableauNames
  };
}

function autoAssignTableauNames(ss, officeId, roster, possibleTableauNames) {
  if (!possibleTableauNames || Object.keys(possibleTableauNames).length === 0) return roster;

  var RANK_ORDER = ['rep', 'l1', 'jd', 'manager', 'admin', 'owner', 'superadmin'];
  var unassigned = [];
  Object.keys(roster).forEach(function(email) {
    var r = roster[email];
    if (!r.tableauName && possibleTableauNames[email]) {
      unassigned.push({ email: email, rank: r.rank || 'rep' });
    }
  });
  unassigned.sort(function(a, b) {
    return RANK_ORDER.indexOf(b.rank) - RANK_ORDER.indexOf(a.rank);
  });

  if (unassigned.length === 0) return roster;

  var claimed = {};
  Object.keys(roster).forEach(function(email) {
    if (roster[email].tableauName) claimed[roster[email].tableauName] = true;
  });

  var sheet = ss.getSheetByName(officeTab(TAB.ROSTER, officeId));
  var newAssignments = [];

  unassigned.forEach(function(entry) {
    var possible = possibleTableauNames[entry.email] || [];
    var unclaimed = possible.filter(function(n) { return !claimed[n]; });
    if (unclaimed.length === 1) {
      var name = unclaimed[0];
      claimed[name] = true;
      roster[entry.email].tableauName = name;
      newAssignments.push({ email: entry.email, name: name });
    }
  });

  if (newAssignments.length > 0 && sheet) {
    var data = sheet.getDataRange().getValues();
    newAssignments.forEach(function(a) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim().toLowerCase() === a.email) {
          sheet.getRange(i + 1, 9).setValue(a.name);
          break;
        }
      }
    });
  }

  return roster;
}

function readTableauDetail(ss, dsi) {
  var sheet = ss.getSheetByName(TABLEAU_TAB);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var targetDsi = String(dsi || '').trim();
  if (!targetDsi) return [];
  var col = buildTableauColumnMap(data[0]);
  var devices = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowDsi = String(tCol(row, col, 'DSI') || '').trim();
    if (rowDsi !== targetDsi) continue;
    devices.push({
      spe: String(tCol(row, col, 'SPE') || '').trim(),
      productType: String(tCol(row, col, 'PRODUCT_TYPE') || '').trim(),
      cruIru: String(tCol(row, col, 'CRU_IRU') || '').trim(),
      dtrStatus: String(tCol(row, col, 'DTR_STATUS') || '').trim(),
      discoReason: String(tCol(row, col, 'DISCO_REASON') || '').trim(),
      phone: String(tCol(row, col, 'PHONE') || '').trim(),
      tnType: String(tCol(row, col, 'TN_TYPE') || '').trim(),
      orderStatus: String(tCol(row, col, 'ORDER_STATUS') || '').trim(),
      postedDate: String(tCol(row, col, 'POSTED_DATE') || '').trim(),
      installDate: String(tCol(row, col, 'INSTALL_DATE') || '').trim()
    });
  }
  return devices;
}

function getTableauSummaryWithCache(ss, officeId) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'tableauSummary_v5_' + officeId;
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* fall through */ }
  }
  var summary = readTableauSummary(ss, officeId);
  try {
    var json = JSON.stringify(summary);
    if (json.length < 100000) cache.put(cacheKey, json, 21600);
  } catch (e) { /* caching failed */ }
  return summary;
}

function writeBustTableauCache(officeId) {
  try { CacheService.getScriptCache().remove('tableauSummary_v5_' + officeId); } catch (e) { }
  return { ok: true, message: 'Tableau cache cleared' };
}


// === doPost() — WRITE ACTIONS ===

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'invalid JSON' });
  }

  if (!validateKey(body.key || '')) {
    return jsonResponse({ error: 'unauthorized' });
  }

  const officeId = body.officeId || DEFAULT_OFFICE_ID;
  const ss = getSheet(body);

  try {
    let result;
    switch (body.action) {
      case 'addRosterEntry':      result = writeAddRosterEntry(body, ss, officeId); break;
      case 'updateRosterEntry':   result = writeUpdateRosterEntry(body, ss, officeId); break;
      case 'setTableauName':      result = writeSetTableauName(body, ss, officeId); break;
      case 'deleteRosterEntry':   result = writeDeleteRosterEntry(body, ss, officeId); break;
      case 'toggleDeactivate':    result = writeToggleDeactivate(body, ss, officeId); break;
      case 'saveOrderOverride':    result = writeSaveOrderOverride(body, ss, officeId); break;
      case 'setTeamCustomization': result = writeSetTeamCustomization(body, ss, officeId); break;
      case 'setUnlockRequest':     result = writeSetUnlockRequest(body, ss, officeId); break;
      case 'deleteUnlockRequest':  result = writeDeleteUnlockRequest(body, ss, officeId); break;
      case 'addTeam':              result = writeAddTeam(body, ss, officeId); break;
      case 'updateTeam':           result = writeUpdateTeam(body, ss, officeId); break;
      case 'deleteTeam':           result = writeDeleteTeam(body, ss, officeId); break;
      case 'setPin':               result = writeSetPin(body, ss, officeId); break;
      case 'validatePin':          result = writeValidatePin(body, ss, officeId); break;
      case 'changePin':            result = writeChangePin(body, ss, officeId); break;
      case 'writeOrderNote':       result = writeOrderNote(body, ss, officeId); break;
      case 'setOrderStatus':       result = writeSetOrderStatus(body, ss, officeId); break;
      case 'updateOrder':          result = writeUpdateOrder(body, ss, officeId); break;
      case 'setSetting':           result = writeSetting(body, ss, officeId); break;
      case 'savePaidOut':          result = writeSavePaidOut(body, ss, officeId); break;
      case 'addTicket':            result = writeAddTicket(body, ss, officeId); break;
      case 'toggleTicket':         result = writeToggleTicket(body, ss, officeId); break;
      case 'addSale':              result = writeAddSale(body, ss, officeId); break;
      case 'replayWebhook':        result = replayWebhook(body, ss, officeId); break;
      case 'bustTableauCache':     result = writeBustTableauCache(officeId); break;
      case 'createOfficeTabs':     result = createOfficeTabs(body, ss); break;
      default: result = { error: 'unknown action: ' + body.action };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}


// === ROSTER WRITE HELPERS ===

function writeAddRosterEntry(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.ROSTER, officeId), TAB.ROSTER);
  const email = String(body.email || '').trim().toLowerCase();
  if (!email) return { error: 'missing email' };
  const existing = findRowCI(sheet, 0, email);
  if (existing > 0) return { error: 'email already exists' };
  sheet.appendRow([
    email, body.name || '', body.team || '', body.rank || 'rep',
    false, new Date().toISOString().split('T')[0], '', body.phone || '', ''
  ]);
  return { ok: true };
}

function writeUpdateRosterEntry(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.ROSTER, officeId), TAB.ROSTER);
  const email = String(body.email || '').trim().toLowerCase();
  if (!email) return { error: 'missing email' };
  const rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx < 0) return { error: 'email not found' };
  const newEmail = body.newEmail ? String(body.newEmail).trim().toLowerCase() : '';
  if (newEmail && newEmail !== email) {
    const conflict = findRowCI(sheet, 0, newEmail);
    if (conflict > 0) return { error: 'new email already exists in roster' };
  }
  const cur = sheet.getRange(rowIdx, 1, 1, 9).getValues()[0];
  const rowData = [
    newEmail || email,
    body.name !== undefined ? body.name : cur[1],
    body.team !== undefined ? body.team : cur[2],
    body.rank !== undefined ? body.rank : cur[3],
    body.deactivated !== undefined ? body.deactivated : cur[4],
    cur[5], cur[6],
    body.phone !== undefined ? body.phone : (cur[7] || ''),
    body.tableauName !== undefined ? body.tableauName : (cur[8] || '')
  ];
  sheet.getRange(rowIdx, 1, 1, 9).setValues([rowData]);
  return { ok: true };
}

function writeDeleteRosterEntry(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.ROSTER, officeId), TAB.ROSTER);
  const email = String(body.email || '').trim().toLowerCase();
  if (!email) return { error: 'missing email' };
  const rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx > 0) sheet.deleteRow(rowIdx);
  return { ok: true };
}

function writeToggleDeactivate(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.ROSTER, officeId), TAB.ROSTER);
  const email = String(body.email || '').trim().toLowerCase();
  if (!email) return { error: 'missing email' };
  const rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 5).setValue(body.deactivated === true || body.deactivated === 'true');
  }
  return { ok: true };
}

function writeSetTableauName(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.ROSTER, officeId), TAB.ROSTER);
  const email = String(body.email || '').trim().toLowerCase();
  if (!email) return { error: 'missing email' };
  const tableauName = String(body.tableauName || '').trim();
  if (!tableauName) return { error: 'missing tableauName' };
  const rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx < 0) return { error: 'email not found' };
  sheet.getRange(rowIdx, 9).setValue(tableauName);
  return { ok: true };
}


// === EXISTING WRITE HELPERS ===

function writeSaveOrderOverride(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.OVERRIDES, officeId), TAB.OVERRIDES);
  const key = String(body.overrideKey || '').trim();
  if (!key) return { error: 'missing key' };
  const rowIdx = findRow(sheet, 0, key);
  const rowData = [key, body.product || '', body.status || '', body.date || '', body.order || '',
    JSON.stringify(body.notes || [])];
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, 6).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return { ok: true };
}

function writeSetTeamCustomization(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.TEAM_CUSTOM, officeId), TAB.TEAM_CUSTOM);
  const persona = String(body.persona || '').trim();
  if (!persona) return { error: 'missing persona' };
  const rowIdx = findRow(sheet, 0, persona);
  const rowData = [persona, body.emoji || '⚡', body.displayName || ''];
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, 3).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return { ok: true };
}

function writeSetUnlockRequest(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.UNLOCKS, officeId), TAB.UNLOCKS);
  const persona = String(body.persona || '').trim();
  if (!persona) return { error: 'missing persona' };
  const rowIdx = findRow(sheet, 0, persona);
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 2).setValue(body.status || 'pending');
  } else {
    sheet.appendRow([persona, body.status || 'pending']);
  }
  return { ok: true };
}

function writeDeleteUnlockRequest(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.UNLOCKS, officeId), TAB.UNLOCKS);
  const persona = String(body.persona || '').trim();
  if (!persona) return { error: 'missing persona' };
  const rowIdx = findRow(sheet, 0, persona);
  if (rowIdx > 0) sheet.deleteRow(rowIdx);
  return { ok: true };
}

// Team hierarchy management
function writeAddTeam(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.TEAMS, officeId), TAB.TEAMS);
  const teamId = 'team_' + Date.now();
  sheet.appendRow([
    teamId,
    String(body.name || '').trim(),
    String(body.parentId || '').trim(),
    String(body.leaderId || '').trim(),
    String(body.emoji || '').trim(),
    new Date().toISOString().split('T')[0]
  ]);
  return { ok: true, teamId: teamId };
}

function writeUpdateTeam(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.TEAMS, officeId), TAB.TEAMS);
  const teamId = String(body.teamId || '').trim();
  if (!teamId) return { error: 'missing teamId' };
  const rowIdx = findRow(sheet, 0, teamId);
  if (rowIdx < 0) return { error: 'team not found' };
  const cur = sheet.getRange(rowIdx, 1, 1, 6).getValues()[0];
  sheet.getRange(rowIdx, 1, 1, 6).setValues([[
    teamId,
    body.name !== undefined ? String(body.name).trim() : cur[1],
    body.parentId !== undefined ? String(body.parentId).trim() : cur[2],
    body.leaderId !== undefined ? String(body.leaderId).trim() : cur[3],
    body.emoji !== undefined ? String(body.emoji).trim() : cur[4],
    cur[5]
  ]]);
  return { ok: true };
}

function writeDeleteTeam(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.TEAMS, officeId), TAB.TEAMS);
  const teamId = String(body.teamId || '').trim();
  if (!teamId) return { error: 'missing teamId' };
  const rowIdx = findRow(sheet, 0, teamId);
  if (rowIdx > 0) sheet.deleteRow(rowIdx);
  return { ok: true };
}

// PIN management
function writeSetPin(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.ROSTER, officeId), TAB.ROSTER);
  const email = String(body.email || '').trim().toLowerCase();
  const pin = String(body.pin || '').trim();
  if (!email || !pin) return { error: 'missing email or pin' };
  const rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx < 0) return { error: 'email not found' };
  var existingHash = String(sheet.getRange(rowIdx, 7).getValue() || '').trim();
  if (existingHash && existingHash !== 'undefined') return { error: 'PIN already set' };
  sheet.getRange(rowIdx, 7).setValue(hashPin(email, pin));
  return { ok: true };
}

function writeValidatePin(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.ROSTER, officeId), TAB.ROSTER);
  const email = String(body.email || '').trim().toLowerCase();
  const pin = String(body.pin || '').trim();
  if (!email || !pin) return { error: 'missing email or pin' };
  const rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx < 0) return { error: 'email not found' };
  var stored = String(sheet.getRange(rowIdx, 7).getValue() || '').trim();
  if (!stored || stored === 'undefined') return { error: 'no PIN set' };
  var attempt = hashPin(email, pin);
  if (attempt === stored) return { ok: true, valid: true };
  return { ok: true, valid: false };
}

function writeChangePin(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.ROSTER, officeId), TAB.ROSTER);
  const email = String(body.email || '').trim().toLowerCase();
  const oldPin = String(body.oldPin || '').trim();
  const newPin = String(body.newPin || '').trim();
  if (!email || !oldPin || !newPin) return { error: 'missing fields' };
  const rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx < 0) return { error: 'email not found' };
  var stored = String(sheet.getRange(rowIdx, 7).getValue() || '').trim();
  if (!stored || stored === 'undefined') return { error: 'no PIN set' };
  if (hashPin(email, oldPin) !== stored) return { error: 'incorrect current PIN' };
  sheet.getRange(rowIdx, 7).setValue(hashPin(email, newPin));
  return { ok: true };
}

// Order management
function writeOrderNote(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.SALES, officeId), TAB.SALES);
  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'invalid rowIndex' };
  var noteCol = OL.NOTES + 1;
  var existing = String(sheet.getRange(rowIndex, noteCol).getValue() || '').trim();
  var newNote = String(body.note || '').trim();
  var combined = existing ? existing + '\n---\n' + newNote : newNote;
  sheet.getRange(rowIndex, noteCol).setValue(combined);
  return { ok: true };
}

function writeSetOrderStatus(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.SALES, officeId), TAB.SALES);
  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'invalid rowIndex' };
  sheet.getRange(rowIndex, OL.STATUS + 1).setValue(String(body.status || 'Pending').trim());
  return { ok: true };
}

function writeUpdateOrder(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.SALES, officeId), TAB.SALES);
  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'invalid rowIndex' };

  if (body.dateOfSale !== undefined) sheet.getRange(rowIndex, OL.DATE_OF_SALE + 1).setValue(new Date(body.dateOfSale + 'T12:00:00'));
  if (body.dsi !== undefined) sheet.getRange(rowIndex, OL.DSI + 1).setValue(String(body.dsi).trim());
  if (body.accountType !== undefined) sheet.getRange(rowIndex, OL.ACCOUNT_TYPE + 1).setValue(String(body.accountType).trim());
  if (body.air !== undefined) sheet.getRange(rowIndex, OL.AIR + 1).setValue(Number(body.air) || 0);
  if (body.newPhones !== undefined) sheet.getRange(rowIndex, OL.NEW_PHONES + 1).setValue(Number(body.newPhones) || 0);
  if (body.byods !== undefined) sheet.getRange(rowIndex, OL.BYODS + 1).setValue(Number(body.byods) || 0);

  // Recalculate Cell, Yeses, Units
  var air = Number(sheet.getRange(rowIndex, OL.AIR + 1).getValue()) || 0;
  var newPhones = Number(sheet.getRange(rowIndex, OL.NEW_PHONES + 1).getValue()) || 0;
  var byods = Number(sheet.getRange(rowIndex, OL.BYODS + 1).getValue()) || 0;
  var cell = newPhones + byods;
  sheet.getRange(rowIndex, OL.CELL + 1).setValue(cell);

  // NDS yeses: Air + Cell only
  var yeses = 0;
  if (air > 0) yeses++;
  if (cell > 0) yeses++;
  sheet.getRange(rowIndex, OL.YESES + 1).setValue(yeses);

  // NDS units: Air + Cell only
  var units = air + cell;
  sheet.getRange(rowIndex, OL.UNITS + 1).setValue(units);

  return { ok: true };
}

function writeSavePaidOut(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.SALES, officeId), TAB.SALES);
  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'invalid rowIndex' };
  sheet.getRange(rowIndex, OL.PAID_OUT + 1).setValue(JSON.stringify(body.paidOut || {}));
  return { ok: true };
}

// Ticket management
function writeAddTicket(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.SALES, officeId), TAB.SALES);
  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'invalid rowIndex' };
  var ticketsCol = OL.TICKETS + 1;
  var existing = [];
  try { existing = JSON.parse(sheet.getRange(rowIndex, ticketsCol).getValue() || '[]'); } catch(e) { existing = []; }
  existing.push({
    id: 'tkt_' + Date.now(),
    text: String(body.text || '').trim(),
    completed: false,
    createdBy: String(body.createdBy || '').trim(),
    createdAt: new Date().toISOString()
  });
  sheet.getRange(rowIndex, ticketsCol).setValue(JSON.stringify(existing));
  return { ok: true };
}

function writeToggleTicket(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.SALES, officeId), TAB.SALES);
  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'invalid rowIndex' };
  var ticketsCol = OL.TICKETS + 1;
  var existing = [];
  try { existing = JSON.parse(sheet.getRange(rowIndex, ticketsCol).getValue() || '[]'); } catch(e) { existing = []; }
  var ticketId = String(body.ticketId || '').trim();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].id === ticketId) {
      existing[i].completed = !existing[i].completed;
      break;
    }
  }
  sheet.getRange(rowIndex, ticketsCol).setValue(JSON.stringify(existing));
  return { ok: true };
}


// === writeAddSale() — NDS: Air + Cell only ===

function writeAddSale(body, ss, officeId) {
  var sheet = getOrCreateSheet(ss, officeTab(TAB.SALES, officeId), TAB.SALES);

  var email = String(body.email || '').trim().toLowerCase();
  if (!email) return { error: 'Missing email' };
  var dateOfSale = body.dateOfSale;
  if (!dateOfSale) return { error: 'Missing date of sale' };

  // Look up team emoji from Roster + Teams
  var teamEmoji = '';
  try {
    var rosterSheet = ss.getSheetByName(officeTab(TAB.ROSTER, officeId));
    if (rosterSheet) {
      var rosterData = rosterSheet.getDataRange().getValues();
      var teamName = '';
      for (var r = 1; r < rosterData.length; r++) {
        if (String(rosterData[r][0]).trim().toLowerCase() === email) {
          teamName = String(rosterData[r][2] || '').trim();
          break;
        }
      }
      if (teamName) {
        var teamsSheet = ss.getSheetByName(officeTab(TAB.TEAMS, officeId));
        if (teamsSheet) {
          var teamsData = teamsSheet.getDataRange().getValues();
          for (var t = 1; t < teamsData.length; t++) {
            if (String(teamsData[t][1] || '').trim() === teamName) {
              teamEmoji = String(teamsData[t][4] || '').trim();
              break;
            }
          }
        }
      }
    }
  } catch (e) { /* emoji lookup failed, non-critical */ }

  // Calculate product counts (NDS: Air + Cell only)
  var air = Number(body.air) || 0;
  var newPhones = Number(body.newPhones) || 0;
  var byods = Number(body.byods) || 0;
  var cell = newPhones + byods;

  // YESES: count of products sold (Air + Cell only)
  var yeses = 0;
  if (air > 0) yeses++;
  if (cell > 0) yeses++;

  // UNITS: sum of products (Air + Cell only)
  var units = air + cell;

  // Build 26-column row (NDS schema — no Fiber/VoIP/DTV columns)
  var newRow = [
    new Date(),                                                    // 0  Timestamp
    email,                                                         // 1  Email
    String(body.repName || '').trim(),                              // 2  Rep Name
    new Date(dateOfSale + 'T12:00:00'),                            // 3  Date of Sale
    String(body.campaign || 'attb2b').trim(),                       // 4  Campaign
    String(body.dsi || '').trim(),                                  // 5  DSI
    String(body.accountType || '').trim(),                          // 6  Account Type
    String(body.clientName || '').trim(),                           // 7  Client Name
    body.trainee ? 'Yes' : 'No',                                   // 8  Trainee
    body.trainee ? String(body.traineeName || '').trim() : '',      // 9  Trainee Name
    air,                                                           // 10 Air
    newPhones,                                                     // 11 New Phones
    byods,                                                         // 12 BYODs
    cell,                                                          // 13 Cell
    '',                                                            // 14 Ooma Package
    String(body.accountNotes || '').trim(),                         // 15 Account Notes
    body.activationSupport ? 'Yes' : 'No',                         // 16 Activation Support
    teamEmoji,                                                     // 17 Team Emoji
    yeses,                                                         // 18 Yeses
    units,                                                         // 19 Units
    'Pending',                                                     // 20 Status
    '',                                                            // 21 Notes
    '',                                                            // 22 Paid Out
    '[]',                                                          // 23 Tickets
    String(body.orderChannel || 'Sara').trim(),                     // 24 Order Channel
    String(body.codesUsedBy || '').trim().toLowerCase()             // 25 Codes Used By
  ];

  sheet.appendRow(newRow);

  // Bust the Tableau cache so fresh data loads
  try {
    CacheService.getScriptCache().remove('tableauSummary_v5_' + officeId);
  } catch (e) { /* non-critical */ }

  // Fire Discord/GroupMe webhook server-side
  var webhookResult = 'skipped';
  try {
    webhookResult = _fireWebhook(body, units, teamEmoji);
  } catch (e) {
    webhookResult = 'error: ' + e.message;
  }

  return {
    ok: true,
    rowIndex: sheet.getLastRow(),
    units: units,
    yeses: yeses,
    webhook: webhookResult
  };
}


// === DISCORD / GROUPME WEBHOOK (server-side) ===

function _fireWebhook(body, units, teamEmoji) {
  var platform = String(body.chatPlatform || 'discord').toLowerCase();
  var webhookUrl = String(body.discordWebhookUrl || '').trim();
  if (!webhookUrl) return 'no_url';
  if (platform === 'none') return 'platform_none';

  var bold = (platform === 'discord') ? '**' : '';
  var repName = String(body.repName || '').trim();
  var traineeName = (body.trainee === true || body.trainee === 'Yes') ? String(body.traineeName || '').trim() : '';
  var who = traineeName ? (repName + ' and ' + traineeName) : repName;
  var campaign = String(body.campaign || '').trim();
  var msg = '';

  if (campaign === 'attb2b') {
    msg += bold + who + bold + ' made a sale with AT&T: NDS!\n';
    msg += (body.accountType || 'Consumer') + ' Account\n';
    msg += String(body.dsi || '') + '\n';
    if (Number(body.air) > 0) msg += '• Internet Air\n';
    var np = Number(body.newPhones) || 0;
    var by = Number(body.byods) || 0;
    if (np > 0 || by > 0) msg += '• ' + np + ' New Phone(s)|' + by + ' BYOD(s)\n';
  }

  var tags = String(body.hashtags || '').trim();
  if (tags) msg += tags + '\n';

  if (teamEmoji && units > 0) {
    var count = Math.min(units, 20);
    for (var i = 0; i < count; i++) msg += teamEmoji;
  }

  msg = msg.trim();
  if (!msg) return 'empty_message';

  var payload, url;
  if (platform === 'groupme') {
    url = 'https://api.groupme.com/v3/bots/post';
    payload = JSON.stringify({ bot_id: webhookUrl, text: msg });
  } else {
    url = webhookUrl;
    payload = JSON.stringify({ content: msg });
  }

  var fetchOpts = {
    method: 'post',
    contentType: 'application/json',
    payload: payload,
    muteHttpExceptions: true
  };

  // Retry up to 2 times on failure (rate-limit 429, server errors 5xx, or fetch exceptions)
  var maxAttempts = 3;
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      var resp = UrlFetchApp.fetch(url, fetchOpts);
      var code = resp.getResponseCode();
      if (code >= 200 && code < 300) {
        return 'sent_' + code + (attempt > 1 ? '_retry' + (attempt - 1) : '');
      }
      // Rate limited — Discord sends Retry-After header (seconds)
      if (code === 429 && attempt < maxAttempts) {
        var retryAfter = 2;
        try {
          var ra = resp.getHeaders()['Retry-After'] || resp.getHeaders()['retry-after'];
          if (ra) retryAfter = Math.min(Math.ceil(Number(ra)), 5);
        } catch (e) { /* use default */ }
        Utilities.sleep(retryAfter * 1000);
        continue;
      }
      // Server error — wait and retry
      if (code >= 500 && attempt < maxAttempts) {
        Utilities.sleep(2000);
        continue;
      }
      return 'http_' + code + '_attempt' + attempt + ':' + resp.getContentText().substring(0, 80);
    } catch (e) {
      if (attempt < maxAttempts) {
        Utilities.sleep(2000);
        continue;
      }
      return 'fetch_error_attempt' + attempt + ': ' + e.message;
    }
  }
  return 'exhausted_retries';
}


// === REPLAY WEBHOOK (re-fire Discord/GroupMe for existing sales) ===

function replayWebhook(body, ss, officeId) {
  var dsi = String(body.dsi || '').trim();
  if (!dsi) return { error: 'missing dsi' };

  var sheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!sheet) return { error: 'no sales sheet' };

  var data = sheet.getDataRange().getValues();
  var results = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][OL.DSI] || '').trim() === dsi) {
      var row = data[i];
      var saleBody = {
        repName: row[OL.REP_NAME],
        campaign: row[OL.CAMPAIGN] || 'attb2b',
        accountType: row[OL.ACCOUNT_TYPE] || 'Consumer',
        dsi: row[OL.DSI],
        trainee: String(row[OL.TRAINEE] || '').trim(),
        traineeName: String(row[OL.TRAINEE_NAME] || '').trim(),
        air: row[OL.AIR],
        newPhones: row[OL.NEW_PHONES],
        byods: row[OL.BYODS],
        hashtags: '',
        discordWebhookUrl: body.discordWebhookUrl || '',
        chatPlatform: body.chatPlatform || 'discord'
      };
      var units = (Number(row[OL.AIR]) || 0) + (Number(row[OL.NEW_PHONES]) || 0) + (Number(row[OL.BYODS]) || 0);
      var teamEmoji = String(row[OL.TEAM_EMOJI] || '').trim();
      var wh = _fireWebhook(saleBody, units, teamEmoji);
      results.push({ dsi: dsi, webhook: wh });
    }
  }
  if (results.length === 0) return { error: 'DSI not found: ' + dsi };
  return { ok: true, results: results };
}


// === OFFICE PROVISIONING (called by admin portal) ===

function createOfficeTabs(body, ss) {
  var oid = String(body.officeId || '').trim();
  if (!oid) return { error: 'Missing officeId' };

  var created = [];
  [TAB.SALES, TAB.ROSTER, TAB.TEAMS, TAB.TEAM_CUSTOM,
   TAB.OVERRIDES, TAB.UNLOCKS, TAB.SETTINGS].forEach(function(base) {
    var tabName = officeTab(base, oid);
    getOrCreateSheet(ss, tabName, base);
    created.push(tabName);
  });

  return { success: true, officeId: oid, tabs: created };
}


// === LEADERBOARD SNAPSHOT (for Make.com → Discord) ===

function readLeaderboard(ss, officeId) {
  var roster = readRoster(ss, officeId);
  var teamMaps = buildTeamEmojiMaps(ss, officeId);
  var result = readPeople(ss, officeId, roster, teamMaps.nameMap);
  var people = result.people || result;

  // Exclude non-sales roles (superadmin/admin are test-only)
  var NON_SALES_LB = { superadmin: true, admin: true };

  var reps = [];
  var leaders = [];
  people.forEach(function(p) {
    if (p.deactivated) return;
    var rank = (p.rank || 'rep').toLowerCase();
    if (NON_SALES_LB[rank]) return;
    var entry = {
      name: p.name,
      rank: p.rank,
      team: p.team,
      teamEmoji: p.teamEmoji,
      tw: p.thisWeek,
      lw: p.priorWeek,
      w2: p.twoWkPrior,
      w3: p.threeWkPrior
    };
    if (p.type === 'leader') leaders.push(entry);
    else reps.push(entry);
  });

  // Sort by units desc
  reps.sort(function(a, b) { return (b.tw.units || 0) - (a.tw.units || 0); });
  leaders.sort(function(a, b) { return (b.tw.units || 0) - (a.tw.units || 0); });

  return { reps: reps, leaders: leaders };
}

function buildLeaderboardHtml(lb) {
  var html = '<html><body style="font-family:sans-serif;padding:20px;background:#f0f4f8">';
  html += '<h2 style="text-align:center">Daily Leaderboard</h2>';

  // Reps table
  html += '<h3>Reps</h3>';
  html += '<table style="width:100%;border-collapse:collapse;background:white">';
  html += '<tr style="background:#0099cc;color:white"><th style="padding:8px;text-align:left">#</th><th style="padding:8px;text-align:left">Name</th><th style="padding:8px">Air</th><th style="padding:8px">Cell</th><th style="padding:8px">Units</th></tr>';

  lb.reps.forEach(function(r, i) {
    var bg = i % 2 === 0 ? '#ffffff' : '#f5f5f5';
    html += '<tr style="background:' + bg + '">';
    html += '<td style="padding:6px 8px">' + (i + 1) + '</td>';
    html += '<td style="padding:6px 8px">' + (r.teamEmoji || '') + ' ' + r.name + '</td>';
    html += '<td style="padding:6px 8px;text-align:center">' + (r.tw.air || 0) + '</td>';
    html += '<td style="padding:6px 8px;text-align:center">' + (r.tw.cell || 0) + '</td>';
    html += '<td style="padding:6px 8px;text-align:center;font-weight:bold">' + (r.tw.units || 0) + '</td>';
    html += '</tr>';
  });
  html += '</table>';

  // Leaders table
  if (lb.leaders.length > 0) {
    html += '<h3 style="margin-top:20px">Leaders</h3>';
    html += '<table style="width:100%;border-collapse:collapse;background:white">';
    html += '<tr style="background:#004466;color:white"><th style="padding:8px;text-align:left">#</th><th style="padding:8px;text-align:left">Name</th><th style="padding:8px">Air</th><th style="padding:8px">Cell</th><th style="padding:8px">Units</th></tr>';

    lb.leaders.forEach(function(r, i) {
      var bg = i % 2 === 0 ? '#ffffff' : '#f5f5f5';
      html += '<tr style="background:' + bg + '">';
      html += '<td style="padding:6px 8px">' + (i + 1) + '</td>';
      html += '<td style="padding:6px 8px">' + (r.teamEmoji || '') + ' ' + r.name + '</td>';
      html += '<td style="padding:6px 8px;text-align:center">' + (r.tw.air || 0) + '</td>';
      html += '<td style="padding:6px 8px;text-align:center">' + (r.tw.cell || 0) + '</td>';
      html += '<td style="padding:6px 8px;text-align:center;font-weight:bold">' + (r.tw.units || 0) + '</td>';
      html += '</tr>';
    });
    html += '</table>';
  }

  html += '</body></html>';
  return html;
}


// ═══════════════════════════════════════════════════════════════
// LEADERBOARD TEXT — Emoji-format message for Discord/GroupMe
// ═══════════════════════════════════════════════════════════════
// Called by AdminCode.gs centralized scheduler via ?action=leaderboardText
// Can also be called manually via postLeaderboardToChat() for testing.

/**
 * Build emoji-formatted leaderboard text from this office's data.
 * @param {Spreadsheet} ss - The office spreadsheet
 * @param {string} officeId - The office ID for tab suffixes
 * @param {string} officeName - Display name for the header
 * @returns {string} The formatted message text
 */
function buildLeaderboardText(ss, officeId, officeName) {
  var lb = readLeaderboard(ss, officeId);
  var allPeople = (lb.reps || []).concat(lb.leaders || []);

  // Filter out non-sales roles (superadmin/admin are test-only)
  var NON_SALES = { superadmin: true, admin: true };
  allPeople = allPeople.filter(function(p) {
    var rank = (p.rank || 'rep').toLowerCase();
    return !NON_SALES[rank];
  });

  allPeople.sort(function(a, b) { return (b.tw.units || 0) - (a.tw.units || 0); });

  var medals = ['🥇', '🥈', '🥉'];
  var lines = [];
  lines.push('🔥 ' + officeName.toUpperCase() + ' WEEKLY 🔥');
  lines.push('━━━━━━━━━━━━━━━━━━━');
  lines.push('');

  var officeTotal = 0;
  var shown = 0;
  for (var i = 0; i < allPeople.length; i++) {
    var p = allPeople[i];
    var units = p.tw.units || 0;
    officeTotal += units;
    if (units <= 0) continue;
    var medal = shown < 3 ? medals[shown] : '🏅';
    lines.push(medal + ' ' + p.name + ' — ' + units);
    shown++;
  }

  if (shown === 0) {
    lines.push('No sales this week yet.');
  }

  // Team rankings
  var teamTotals = {};
  for (var j = 0; j < allPeople.length; j++) {
    var person = allPeople[j];
    var team = person.team || 'Unassigned';
    if (team === 'Unassigned' || !team) continue;
    if (!teamTotals[team]) teamTotals[team] = 0;
    teamTotals[team] += (person.tw.units || 0);
  }

  var teamArr = [];
  for (var tName in teamTotals) {
    if (teamTotals[tName] > 0) {
      teamArr.push({ name: tName, units: teamTotals[tName] });
    }
  }
  teamArr.sort(function(a, b) { return b.units - a.units; });

  if (teamArr.length > 0) {
    lines.push('');
    lines.push('━━━━━━━━━━━━━━━━━━━');
    lines.push('⚔️ TEAM RANKINGS ⚔️');
    lines.push('━━━━━━━━━━━━━━━━━━━');
    lines.push('');
    for (var t = 0; t < teamArr.length; t++) {
      var tMedal = t < 3 ? medals[t] : '🏅';
      lines.push(tMedal + ' ' + teamArr[t].name + ' — ' + teamArr[t].units);
    }
  }

  lines.push('');
  lines.push('📊 Office Total: ' + officeTotal + ' units');

  return lines.join('\n');
}

/**
 * Manual test — posts leaderboard to the webhook configured in Script Properties.
 * For centralized scheduling, use AdminCode.gs checkLeaderboardPosts() instead.
 */
function postLeaderboardToChat() {
  var props = PropertiesService.getScriptProperties();
  var webhookUrl = (props.getProperty('CHAT_WEBHOOK_URL') || '').trim();
  var platform   = (props.getProperty('CHAT_PLATFORM') || 'discord').trim().toLowerCase();
  if (!webhookUrl || platform === 'none') {
    Logger.log('[LeaderboardPost] No webhook configured — skipping');
    return;
  }

  var ss = getSheet({});
  var officeId = DEFAULT_OFFICE_ID;
  var officeName = props.getProperty('OFFICE_NAME') || 'OFFICE';

  var message = buildLeaderboardText(ss, officeId, officeName);

  // Post to webhook
  var url, payload;
  if (platform === 'groupme') {
    url = 'https://api.groupme.com/v3/bots/post';
    payload = JSON.stringify({ bot_id: webhookUrl, text: message });
  } else {
    url = webhookUrl;
    payload = JSON.stringify({ content: message });
  }

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: payload,
    muteHttpExceptions: true
  };

  var resp = UrlFetchApp.fetch(url, options);
  Logger.log('[LeaderboardPost] ' + platform + ' response: ' + resp.getResponseCode());
}
