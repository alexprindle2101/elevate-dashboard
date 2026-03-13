// ═══════════════════════════════════════════════════════
// Campaign Dashboard — Google Apps Script Middleware
// ═══════════════════════════════════════════════════════
// Single shared Code.gs for all AT&T B2B offices.
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

// Default officeId when none is provided (backwards compat for existing Elevate office)
const DEFAULT_OFFICE_ID = 'off_001';


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


// === _Sales column indices (0-based, sequential A–AD, 30 columns) ===

const OL = {
  TIMESTAMP: 0,
  EMAIL: 1,
  REP_NAME: 2,
  DATE_OF_SALE: 3,
  CAMPAIGN: 4,
  DSI: 5,
  ACCOUNT_TYPE: 6,
  CLIENT_NAME: 7,
  TRAINEE: 8,
  TRAINEE_NAME: 9,
  AIR: 10,
  NEW_PHONES: 11,
  BYODS: 12,
  CELL: 13,
  FIBER: 14,
  FIBER_PACKAGE: 15,
  INSTALL_DATE: 16,
  VOIP_QTY: 17,
  DTV: 18,
  DTV_PACKAGE: 19,
  OOMA_PACKAGE: 20,
  ACCOUNT_NOTES: 21,
  ACTIVATION_SUPPORT: 22,
  TEAM_EMOJI: 23,
  YESES: 24,
  UNITS: 25,
  STATUS: 26,
  NOTES: 27,
  PAID_OUT: 28,
  TICKETS: 29,
  ORDER_CHANNEL: 30,
  CODES_USED_BY: 31
};

// Legacy Order Log column indices (for migration only — remove after migration)
const OL_LEGACY = {
  TIMESTAMP: 0,
  EMAIL: 2,
  DATE_OF_SALE: 3,
  REP_NAME: 4,
  DSI: 5,
  TRAINEE: 6,
  TRAINEE_NAME: 7,
  VOIP_QTY: 20,
  TEAM_EMOJI: 26,
  FIBER: 31,
  AIR: 32,
  DTV: 33,
  YESES: 34,
  CELL: 35,
  UNITS: 36,
  STATUS: 38,
  NOTES: 39,
  PAID_OUT: 40,
  TICKETS: 41
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
  'sp.spm number': 'DSI',
  'spm number': 'DSI',
  'dsi': 'DSI',
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
          'DSI', 'Account Type', 'Client Name', 'Trainee', 'Trainee Name',
          'Air', 'New Phones', 'BYODs', 'Cell', 'Fiber',
          'Fiber Package', 'Install Date', 'VoIP Qty', 'DTV', 'DTV Package',
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


// === PERIOD HELPERS ===

function emptyPeriod() {
  return { y: 0, air: 0, cell: 0, fiber: 0, voip: 0, dtv: 0, units: 0 };
}

function addToPeriod(target, sale) {
  target.y += sale.y;
  target.air += sale.air;
  target.cell += sale.cell;
  target.fiber += sale.fiber;
  target.voip += sale.voip;
  target.dtv += sale.dtv;
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

  // Debug counters
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
      'col24(yeses)=' + String(olData[1][OL.YESES]),
      'col25(units)=' + String(olData[1][OL.UNITS])
    ] : [],
    rosterEmails: Object.keys(roster).slice(0, 5),
    olEmails: []
  };

  // Collect first 5 unique OL emails for debug
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

    // Parse date — always force new Date() to handle Apps Script Date-like objects
    let rawDate = row[OL.DATE_OF_SALE];
    if (!rawDate) { _dbg.noDate++; continue; }
    let saleDate = new Date(rawDate);
    if (isNaN(saleDate.getTime())) { _dbg.badDate++; continue; }
    saleDate.setHours(0, 0, 0, 0);

    // Extract sale counters
    const sale = {
      y:     Number(row[OL.YESES]) || 0,
      air:   Number(row[OL.AIR]) || 0,
      cell:  Number(row[OL.CELL]) || 0,
      fiber: Number(row[OL.FIBER]) || 0,
      voip:  Number(row[OL.VOIP_QTY]) || 0,
      dtv:   Number(row[OL.DTV]) || 0,
      units: Number(row[OL.UNITS]) || 0
    };

    const pa = agg[email];

    // weekOffset: 0=this week, 1=last week, ... 5=five weeks ago
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
        if (h >= 10.5 && h < 16)  slotIdx = 0;  // 10:30-4:00
        else if (h >= 16 && h < 18) slotIdx = 1; // 4:00-6:00
        else if (h >= 18 && h < 21) slotIdx = 2; // 6:00-9:00
        else if (h >= 21 || h < 10.5) slotIdx = 3; // 9:00-10:30 (late/early)
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
      fiber: Number(row[OL.FIBER]) || 0,
      fiberPackage: String(row[OL.FIBER_PACKAGE] || '').trim(),
      installDate: String(row[OL.INSTALL_DATE] || '').trim(),
      voip:  Number(row[OL.VOIP_QTY]) || 0,
      dtv:   Number(row[OL.DTV]) || 0,
      dtvPackage: String(row[OL.DTV_PACKAGE] || '').trim(),
      oomaPackage: String(row[OL.OOMA_PACKAGE] || '').trim(),
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

    // Determine if this row is payroll-relevant based on mode
    const trainee = String(row[OL.TRAINEE] || '').trim().toLowerCase();
    const codesUsedBy = String(row[OL.CODES_USED_BY] || '').trim().toLowerCase();
    var isTraineeOrder = (trainee === 'yes');
    var isCodesSwap = (codesUsedBy !== '');

    if (mode === 'flat-rate') {
      // Flat rate: only codes-swap orders go to payroll (trainee irrelevant)
      if (!isCodesSwap) continue;
    } else {
      // Commission split (default): trainee orders AND codes-swap orders
      if (!isTraineeOrder && !isCodesSwap) continue;
    }

    const rawDate = row[OL.DATE_OF_SALE];
    if (!rawDate) continue;
    const saleDate = new Date(rawDate);
    if (isNaN(saleDate.getTime())) continue;
    saleDate.setHours(0, 0, 0, 0);
    if (saleDate < cutoff) continue;

    // Parse paid-out JSON — defaults to empty object
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
      fiber: Number(row[OL.FIBER]) || 0,
      voip:  Number(row[OL.VOIP_QTY]) || 0,
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

  // Map headers by name; detect the metricType column dynamically.
  // The metric type column (containing 'Activated SPE/SP', 'Disconnect count (SPE/SP)',
  // 'Churn Rate') may have a blank header. Find it by scanning data rows for known values.
  const rawHeaders = data[0].map((h) => String(h).trim());

  // Find which column index contains the metric type values
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
  // Not found — append new row
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

// Build DSI → email map from per-office Sales tab (for joining Tableau DSIs to roster emails)
function buildDsiEmailMap(ss, officeId) {
  var olSheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!olSheet) return { dsiToEmail: {}, emailToDsis: {} };
  var olData = olSheet.getDataRange().getValues();
  var dsiToEmail = {};   // DSI → first email (first-wins)
  var emailToDsis = {};  // email → { dsi: true, ... } (ALL DSIs per rep)
  for (var i = 1; i < olData.length; i++) {
    var dsi = String(olData[i][OL.DSI] || '').trim();
    var email = String(olData[i][OL.EMAIL] || '').trim().toLowerCase();
    if (dsi && email) {
      if (!dsiToEmail[dsi]) dsiToEmail[dsi] = email;  // first-wins
      if (!emailToDsis[email]) emailToDsis[email] = {};
      emailToDsis[email][dsi] = true;
    }
  }
  return { dsiToEmail: dsiToEmail, emailToDsis: emailToDsis };
}

// Read and aggregate shared _TableauOrderLog by DSI and by rep email
function readTableauSummary(ss, officeId) {
  var sheet = ss.getSheetByName(TABLEAU_TAB);
  if (!sheet) return { dsiSummary: {}, repSummary: {} };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { dsiSummary: {}, repSummary: {} };

  // Build column map from header row
  var col = buildTableauColumnMap(data[0]);

  // Build DSI → email map from this office's Sales tab
  var maps = buildDsiEmailMap(ss, officeId);
  var dsiEmailMap = maps.dsiToEmail;
  var emailToDsis = maps.emailToDsis;

  // Month window cutoff for Active % calc
  var thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  thirtyDaysAgo.setHours(0, 0, 0, 0);

  // Current-week Monday (Mon=0 week) for tier bonus filtering
  var now = new Date();
  var dayOfWeek = (now.getDay() + 6) % 7; // Mon=0 … Sun=6
  var thisMonday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayOfWeek);
  thisMonday.setHours(0, 0, 0, 0);

  var dsiSummary = {};
  var repTierData = {}; // keyed by tableau rep name, current-week only

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dsi = String(tCol(row, col, 'DSI') || '').trim();
    if (!dsi) continue;

    // Skip the "Total" grand total row
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
    var rawOrderDate = tCol(row, col, 'ORDER_DATE');
    // Handle both native Date objects (manual paste) and strings (CSV automation)
    var orderDate = rawOrderDate instanceof Date ? rawOrderDate : (rawOrderDate ? new Date(rawOrderDate) : null);
    if (orderDate && isNaN(orderDate.getTime())) orderDate = null;
    var bonusTier = String(tCol(row, col, 'BONUS_TIERS') || '').trim();
    var payoutReason = String(tCol(row, col, 'PAYOUT_REASON') || '').trim();

    // Collect current-week tier bonus data per rep name
    if (bonusTier && tableauRep && orderDate && orderDate >= thisMonday) {
      repTierData[tableauRep] = { bonusTier: bonusTier, payoutReason: payoutReason };
    }

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

    if (dtrStatus) {
      s.statusCounts[dtrStatus] = (s.statusCounts[dtrStatus] || 0) + 1;
    }
    if (productType) {
      s.productCounts[productType] = (s.productCounts[productType] || 0) + 1;
    }
    if (discoReason) {
      s.disconnectReasons[discoReason] = (s.disconnectReasons[discoReason] || 0) + 1;
    }
    if (spe) {
      s.speList.push(spe);
      var ptUpper = productType.toUpperCase();
      var isWireless = ptUpper === 'WIRELESS';
      var isInMonth = orderDate && orderDate >= thirtyDaysAgo;
      if (isInMonth && isWireless) {
        s.monthWirelessSPEs[spe] = {
          orderStatus: orderStatus.toLowerCase(),
          dtrStatus: dtrStatus
        };
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
      var tierInfo = repTierData[ds.tableauRep] || {};
      repSummary[email] = {
        totalDevices: 0,
        totalActivations: 0,
        totalVolume: 0,
        statusCounts: {},
        productCounts: {},
        tableauName: ds.tableauRep,
        monthWirelessSPEs: {},
        bonusTier: tierInfo.bonusTier || '',
        payoutReason: tierInfo.payoutReason || ''
      };
    }

    var rs = repSummary[email];
    rs.totalDevices += ds.totalDevices;
    rs.totalActivations += ds.totalActivations;
    rs.totalVolume += ds.totalVolume;

    Object.keys(ds.statusCounts).forEach(function(st) {
      rs.statusCounts[st] = (rs.statusCounts[st] || 0) + ds.statusCounts[st];
    });
    Object.keys(ds.productCounts).forEach(function(pt) {
      rs.productCounts[pt] = (rs.productCounts[pt] || 0) + ds.productCounts[pt];
    });
    Object.keys(ds.monthWirelessSPEs).forEach(function(spe) {
      rs.monthWirelessSPEs[spe] = ds.monthWirelessSPEs[spe];
    });
  });

  // Convert monthWirelessSPEs to counts for the response
  Object.keys(repSummary).forEach(function(email) {
    var rs = repSummary[email];
    var spes = Object.keys(rs.monthWirelessSPEs);
    rs.monthTotalSPEs = spes.length;
    rs.monthApprovedSPEs = 0;
    rs.monthPendingSPEs = 0;
    rs.monthCanceledSPEs = 0;
    rs.monthDiscoSPEs = 0;
    spes.forEach(function(spe) {
      var info = rs.monthWirelessSPEs[spe];
      if (info.orderStatus === 'approved') rs.monthApprovedSPEs++;
      else if (info.orderStatus === 'pending') rs.monthPendingSPEs++;
      else if (info.orderStatus === 'canceled' || info.orderStatus === 'cancelled') rs.monthCanceledSPEs++;
      if (info.dtrStatus === 'Disconnected') rs.monthDiscoSPEs++;
    });
    delete rs.monthWirelessSPEs;
  });

  // Build possibleTableauNames: email → [unique Tableau REP names from their DSIs]
  var possibleTableauNames = {};
  Object.keys(emailToDsis).forEach(function(email) {
    var names = {};
    Object.keys(emailToDsis[email]).forEach(function(dsi) {
      if (dsiSummary[dsi] && dsiSummary[dsi].tableauRep) {
        names[dsiSummary[dsi].tableauRep] = true;
      }
    });
    var nameList = Object.keys(names);
    if (nameList.length > 0) {
      possibleTableauNames[email] = nameList;
    }
  });

  // Build repByName: Tableau REP name → aggregated summary (for stored tableauName lookups)
  var repByName = {};
  Object.keys(dsiSummary).forEach(function(dsi) {
    var ds = dsiSummary[dsi];
    var name = ds.tableauRep;
    if (!name) return;

    if (!repByName[name]) {
      var tierInfoByName = repTierData[name] || {};
      repByName[name] = {
        totalDevices: 0, totalActivations: 0, totalVolume: 0,
        statusCounts: {}, productCounts: {}, tableauName: name,
        monthWirelessSPEs: {},
        bonusTier: tierInfoByName.bonusTier || '',
        payoutReason: tierInfoByName.payoutReason || ''
      };
    }

    var rn = repByName[name];
    rn.totalDevices += ds.totalDevices;
    rn.totalActivations += ds.totalActivations;
    rn.totalVolume += ds.totalVolume;

    Object.keys(ds.statusCounts).forEach(function(st) {
      rn.statusCounts[st] = (rn.statusCounts[st] || 0) + ds.statusCounts[st];
    });
    Object.keys(ds.productCounts).forEach(function(pt) {
      rn.productCounts[pt] = (rn.productCounts[pt] || 0) + ds.productCounts[pt];
    });
    Object.keys(ds.monthWirelessSPEs).forEach(function(spe) {
      rn.monthWirelessSPEs[spe] = ds.monthWirelessSPEs[spe];
    });
  });

  // Convert monthWirelessSPEs to counts for repByName
  Object.keys(repByName).forEach(function(name) {
    var rn = repByName[name];
    var spes = Object.keys(rn.monthWirelessSPEs);
    rn.monthTotalSPEs = spes.length;
    rn.monthApprovedSPEs = 0;
    rn.monthPendingSPEs = 0;
    rn.monthCanceledSPEs = 0;
    rn.monthDiscoSPEs = 0;
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

// Auto-assign tableauNames to roster members who don't have one yet.
// Processes top-down by rank so managers claim first, then reps get what's left.
function autoAssignTableauNames(ss, officeId, roster, possibleTableauNames) {
  if (!possibleTableauNames || Object.keys(possibleTableauNames).length === 0) return roster;

  // Rank priority: higher index = claims first
  var RANK_ORDER = ['rep', 'l1', 'jd', 'manager', 'admin', 'owner', 'superadmin'];

  // Build list of roster members who need assignment, sorted by rank (highest first)
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

  // Build set of already-claimed names
  var claimed = {};
  Object.keys(roster).forEach(function(email) {
    if (roster[email].tableauName) {
      claimed[roster[email].tableauName] = true;
    }
  });

  // Process top-down: claim the sole unclaimed name if exactly one exists
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

  // Write all new assignments to the sheet in one pass
  if (newAssignments.length > 0 && sheet) {
    var data = sheet.getDataRange().getValues();
    newAssignments.forEach(function(a) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim().toLowerCase() === a.email) {
          sheet.getRange(i + 1, 9).setValue(a.name);  // col 9 = tableauName
          break;
        }
      }
    });
    Logger.log('[AutoAssign] Assigned ' + newAssignments.length + ' Tableau names: ' +
      newAssignments.map(function(a) { return a.email + ' → ' + a.name; }).join(', '));
  }

  return roster;
}


// Lazy-load per-device detail rows for a single DSI (shared Tableau tab)
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

// Cached wrapper for readTableauSummary (6-hour TTL, per-office cache key)
function getTableauSummaryWithCache(ss, officeId) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'tableauSummary_v6_' + officeId;
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* fall through */ }
  }

  var summary = readTableauSummary(ss, officeId);
  try {
    var json = JSON.stringify(summary);
    // CacheService limit is 100KB per value
    if (json.length < 100000) {
      cache.put(cacheKey, json, 21600); // 6 hours
    }
  } catch (e) { /* caching failed, that's ok */ }
  return summary;
}

// Manual Tableau cache bust (owner/admin action)
function writeBustTableauCache(officeId) {
  try {
    CacheService.getScriptCache().remove('tableauSummary_v6_' + officeId);
  } catch (e) { /* non-critical */ }
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
  const ss = getSheet(body);  // uses body.sheetId if provided

  try {
    let result;
    switch (body.action) {
      // Roster management (JD+ only — enforced client-side)
      case 'addRosterEntry':      result = writeAddRosterEntry(body, ss, officeId); break;
      case 'updateRosterEntry':   result = writeUpdateRosterEntry(body, ss, officeId); break;
      case 'setTableauName':      result = writeSetTableauName(body, ss, officeId); break;
      case 'deleteRosterEntry':   result = writeDeleteRosterEntry(body, ss, officeId); break;
      case 'toggleDeactivate':    result = writeToggleDeactivate(body, ss, officeId); break;
      // Existing actions
      case 'saveOrderOverride':    result = writeSaveOrderOverride(body, ss, officeId); break;
      case 'setTeamCustomization': result = writeSetTeamCustomization(body, ss, officeId); break;
      case 'setUnlockRequest':     result = writeSetUnlockRequest(body, ss, officeId); break;
      case 'deleteUnlockRequest':  result = writeDeleteUnlockRequest(body, ss, officeId); break;
      // Team hierarchy management
      case 'addTeam':              result = writeAddTeam(body, ss, officeId); break;
      case 'updateTeam':           result = writeUpdateTeam(body, ss, officeId); break;
      case 'deleteTeam':           result = writeDeleteTeam(body, ss, officeId); break;
      // PIN authentication
      case 'setPin':               result = writeSetPin(body, ss, officeId); break;
      case 'validatePin':          result = writeValidatePin(body, ss, officeId); break;
      case 'changePin':            result = writeChangePin(body, ss, officeId); break;
      // Order management
      case 'writeOrderNote':       result = writeOrderNote(body, ss, officeId); break;
      case 'setOrderStatus':       result = writeSetOrderStatus(body, ss, officeId); break;
      case 'updateOrder':          result = writeUpdateOrder(body, ss, officeId); break;
      case 'setSetting':           result = writeSetting(body, ss, officeId); break;
      case 'savePaidOut':          result = writeSavePaidOut(body, ss, officeId); break;
      // Ticket management
      case 'addTicket':            result = writeAddTicket(body, ss, officeId); break;
      case 'toggleTicket':         result = writeToggleTicket(body, ss, officeId); break;
      // Sale submission
      case 'addSale':              result = writeAddSale(body, ss, officeId); break;
      case 'replayWebhook':        result = replayWebhook(body, ss, officeId); break;
      case 'bustTableauCache':     result = writeBustTableauCache(officeId); break;
      // Office provisioning (called by admin portal)
      case 'createOfficeTabs':     result = createOfficeTabs(body, ss); break;
      // One-time migration: old tabs → new per-office tabs
      case 'migrateFromLegacy':    result = migrateFromLegacy(ss, officeId); break;
      case 'migrateFromExternal': result = migrateFromExternal(body, ss); break;
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
    email,
    body.name || '',
    body.team || '',
    body.rank || 'rep',
    false,
    new Date().toISOString().split('T')[0],
    '',  // pinHash — empty until first login
    body.phone || '',
    ''   // tableauName — empty until rep picks from popup
  ]);
  return { ok: true };
}

function writeUpdateRosterEntry(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.ROSTER, officeId), TAB.ROSTER);
  const email = String(body.email || '').trim().toLowerCase();
  if (!email) return { error: 'missing email' };

  const rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx < 0) return { error: 'email not found' };

  // If changing email, check new email doesn't already exist
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
    cur[5], // preserve dateAdded
    cur[6], // preserve pinHash
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
  const key = String(body.key || '').trim();
  if (!key) return { error: 'missing key' };

  const rowData = [
    key,
    body.product || '',
    body.status || '',
    body.date || '',
    body.order || '',
    JSON.stringify(body.notes || [])
  ];

  const rowIdx = findRow(sheet, 0, key);
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return { ok: true };
}

function writeSetTeamCustomization(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.TEAM_CUSTOM, officeId), TAB.TEAM_CUSTOM);
  const persona = String(body.persona || '').trim();
  if (!persona) return { error: 'missing persona' };

  const rowData = [persona, body.emoji || '⚡', body.displayName || ''];
  const rowIdx = findRow(sheet, 0, persona);
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return { ok: true };
}

function writeSetUnlockRequest(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.UNLOCKS, officeId), TAB.UNLOCKS);
  const persona = String(body.persona || '').trim();
  if (!persona) return { error: 'missing persona' };

  const rowData = [persona, body.status || 'pending'];
  const rowIdx = findRow(sheet, 0, persona);
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
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


// === PIN AUTHENTICATION ===

function writeSetPin(body, ss, officeId) {
  var sheet = getOrCreateSheet(ss, officeTab(TAB.ROSTER, officeId), TAB.ROSTER);
  var email = String(body.email || '').trim().toLowerCase();
  var pin = String(body.pin || '').trim();

  if (!email) return { error: 'missing email' };
  if (!pin) return { error: 'missing pin' };
  if (!/^\d{4,6}$/.test(pin)) return { error: 'PIN must be 4-6 digits' };

  var rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx < 0) return { error: 'email not found' };

  var existingHash = String(sheet.getRange(rowIdx, 7).getValue() || '').trim();
  if (existingHash.length > 0 && existingHash !== 'undefined') {
    return { error: 'PIN already set. Use change PIN instead.' };
  }

  sheet.getRange(rowIdx, 7).setValue(hashPin(email, pin));
  return { ok: true };
}

function writeValidatePin(body, ss, officeId) {
  var sheet = getOrCreateSheet(ss, officeTab(TAB.ROSTER, officeId), TAB.ROSTER);
  var email = String(body.email || '').trim().toLowerCase();
  var pin = String(body.pin || '').trim();

  if (!email) return { error: 'missing email' };
  if (!pin) return { error: 'missing pin' };

  var rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx < 0) return { error: 'email not found' };

  var rowData = sheet.getRange(rowIdx, 1, 1, 7).getValues()[0];
  var deactivated = rowData[4] === true || String(rowData[4]).toUpperCase() === 'TRUE';
  if (deactivated) return { error: 'Account deactivated. Contact your Admin.' };

  var storedHash = String(rowData[6] || '').trim();
  if (!storedHash || storedHash === 'undefined') {
    return { error: 'No PIN set for this account' };
  }

  if (hashPin(email, pin) === storedHash) {
    return { ok: true, valid: true };
  } else {
    return { ok: true, valid: false, error: 'Incorrect PIN' };
  }
}

// === TEAM HIERARCHY WRITE HELPERS ===

function writeAddTeam(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.TEAMS, officeId), TAB.TEAMS);
  const teamId = String(body.teamId || '').trim();
  const name = String(body.name || '').trim();
  if (!teamId || !name) return { error: 'missing teamId or name' };

  // Check teamId uniqueness
  if (findRow(sheet, 0, teamId) > 0) return { error: 'teamId already exists' };
  // Check name uniqueness
  if (findRow(sheet, 1, name) > 0) return { error: 'team name already exists' };

  sheet.appendRow([
    teamId,
    name,
    body.parentId || '',
    body.leaderId || '',
    body.emoji || '',
    new Date().toISOString().split('T')[0]
  ]);
  return { ok: true };
}

function writeUpdateTeam(body, ss, officeId) {
  const sheet = getOrCreateSheet(ss, officeTab(TAB.TEAMS, officeId), TAB.TEAMS);
  const teamId = String(body.teamId || '').trim();
  if (!teamId) return { error: 'missing teamId' };

  const rowIdx = findRow(sheet, 0, teamId);
  if (rowIdx < 0) return { error: 'team not found' };

  const cur = sheet.getRange(rowIdx, 1, 1, 6).getValues()[0];

  // Cycle prevention: if parentId is being set, walk up the chain
  var newParentId = body.parentId !== undefined ? String(body.parentId || '').trim() : String(cur[2] || '').trim();
  if (newParentId) {
    var visited = {};
    visited[teamId] = true;
    var walkId = newParentId;
    var allData = sheet.getDataRange().getValues();
    var idToParent = {};
    for (var i = 1; i < allData.length; i++) {
      idToParent[String(allData[i][0] || '').trim()] = String(allData[i][2] || '').trim();
    }
    while (walkId) {
      if (visited[walkId]) return { error: 'circular parent reference' };
      visited[walkId] = true;
      walkId = idToParent[walkId] || '';
    }
  }

  // Check name uniqueness if name is being changed
  if (body.name !== undefined) {
    var newName = String(body.name).trim();
    var nameRow = findRow(sheet, 1, newName);
    if (nameRow > 0 && nameRow !== rowIdx) return { error: 'team name already exists' };
  }

  var rowData = [
    teamId,
    body.name !== undefined ? body.name : cur[1],
    body.parentId !== undefined ? body.parentId : cur[2],
    body.leaderId !== undefined ? body.leaderId : cur[3],
    body.emoji !== undefined ? body.emoji : cur[4],
    cur[5] // preserve createdDate
  ];
  sheet.getRange(rowIdx, 1, 1, 6).setValues([rowData]);
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


function writeChangePin(body, ss, officeId) {
  var sheet = getOrCreateSheet(ss, officeTab(TAB.ROSTER, officeId), TAB.ROSTER);
  var email = String(body.email || '').trim().toLowerCase();
  var currentPin = String(body.currentPin || '').trim();
  var newPin = String(body.newPin || '').trim();

  if (!email || !currentPin || !newPin) return { error: 'missing fields' };
  if (!/^\d{4,6}$/.test(newPin)) return { error: 'New PIN must be 4-6 digits' };

  var rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx < 0) return { error: 'email not found' };

  var storedHash = String(sheet.getRange(rowIdx, 7).getValue() || '').trim();
  if (!storedHash || storedHash === 'undefined') return { error: 'No PIN currently set' };

  if (hashPin(email, currentPin) !== storedHash) {
    return { error: 'Current PIN is incorrect' };
  }

  sheet.getRange(rowIdx, 7).setValue(hashPin(email, newPin));
  return { ok: true };
}


// === ORDER NOTES & STATUS ===

function writeOrderNote(body, ss, officeId) {
  var sheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!sheet) return { error: 'Sales sheet not found' };

  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'Invalid row' };

  var authorName = String(body.authorName || '').trim();
  var noteText = String(body.noteText || '').trim();
  if (!noteText) return { error: 'Empty note' };

  var now = new Date();
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var dateStr = months[now.getMonth()] + ' ' + now.getDate();
  var newEntry = '[' + dateStr + ' \u2014 ' + authorName + '] ' + noteText;

  var existing = String(sheet.getRange(rowIndex, OL.NOTES + 1).getValue() || '').trim();
  var updated = existing ? existing + '\n' + newEntry : newEntry;

  sheet.getRange(rowIndex, OL.NOTES + 1).setValue(updated);
  return { ok: true, notes: updated };
}


// === TICKET MANAGEMENT ===

function writeAddTicket(body, ss, officeId) {
  var sheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!sheet) return { error: 'Sales sheet not found' };

  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'Invalid row' };

  var ticketId = String(body.ticketId || '').trim();
  var ticketText = String(body.ticketText || '').trim();
  var authorName = String(body.authorName || '').trim();
  if (!ticketId) return { error: 'Ticket ID required' };

  var now = new Date();
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var dateStr = months[now.getMonth()] + ' ' + now.getDate();

  var existing = String(sheet.getRange(rowIndex, OL.TICKETS + 1).getValue() || '').trim();
  var tickets;
  try { tickets = JSON.parse(existing || '[]'); } catch(e) { tickets = []; }

  tickets.push({
    id: ticketId,
    text: ticketText,
    author: authorName,
    date: dateStr,
    resolved: false
  });

  sheet.getRange(rowIndex, OL.TICKETS + 1).setValue(JSON.stringify(tickets));
  return { ok: true, tickets: tickets };
}

function writeToggleTicket(body, ss, officeId) {
  var sheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!sheet) return { error: 'Sales sheet not found' };

  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'Invalid row' };

  var ticketId = String(body.ticketId || '').trim();
  if (!ticketId) return { error: 'Ticket ID required' };

  var existing = String(sheet.getRange(rowIndex, OL.TICKETS + 1).getValue() || '').trim();
  var tickets;
  try { tickets = JSON.parse(existing || '[]'); } catch(e) { tickets = []; }

  var found = false;
  for (var i = 0; i < tickets.length; i++) {
    if (tickets[i].id === ticketId) {
      tickets[i].resolved = !tickets[i].resolved;
      found = true;
      break;
    }
  }
  if (!found) return { error: 'Ticket not found' };

  sheet.getRange(rowIndex, OL.TICKETS + 1).setValue(JSON.stringify(tickets));
  return { ok: true, tickets: tickets };
}

function writeSetOrderStatus(body, ss, officeId) {
  var sheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!sheet) return { error: 'Sales sheet not found' };

  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'Invalid row' };

  var status = String(body.status || 'Pending').trim();
  sheet.getRange(rowIndex, OL.STATUS + 1).setValue(status);
  return { ok: true };
}

function writeUpdateOrder(body, ss, officeId) {
  var sheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!sheet) return { error: 'Sales sheet not found' };

  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'Invalid row' };

  // Update each field if provided
  if (body.repName !== undefined)   sheet.getRange(rowIndex, OL.REP_NAME + 1).setValue(String(body.repName).trim());
  if (body.dsi !== undefined)       sheet.getRange(rowIndex, OL.DSI + 1).setValue(String(body.dsi).trim());
  if (body.dateOfSale !== undefined) sheet.getRange(rowIndex, OL.DATE_OF_SALE + 1).setValue(new Date(body.dateOfSale));
  if (body.air !== undefined)       sheet.getRange(rowIndex, OL.AIR + 1).setValue(Number(body.air) || 0);
  if (body.cell !== undefined)      sheet.getRange(rowIndex, OL.CELL + 1).setValue(Number(body.cell) || 0);
  if (body.fiber !== undefined)     sheet.getRange(rowIndex, OL.FIBER + 1).setValue(Number(body.fiber) || 0);
  if (body.voip !== undefined)      sheet.getRange(rowIndex, OL.VOIP_QTY + 1).setValue(Number(body.voip) || 0);
  if (body.status !== undefined)    sheet.getRange(rowIndex, OL.STATUS + 1).setValue(String(body.status).trim());

  return { ok: true };
}


// === PAYROLL PAID-OUT ===

function writeSavePaidOut(body, ss, officeId) {
  var sheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!sheet) return { error: 'Sales sheet not found' };

  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'Invalid row' };

  var paidOutJson = JSON.stringify(body.paidOut || {});
  sheet.getRange(rowIndex, OL.PAID_OUT + 1).setValue(paidOutJson);
  return { ok: true };
}


// === LEADERBOARD SNAPSHOT (for daily Discord post) ===

function readLeaderboard(ss, officeId) {
  const roster = readRoster(ss, officeId);
  const teams = readTeams(ss, officeId);

  // Read Sales tab
  const olSheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!olSheet) return { error: 'No Sales sheet' };
  const olData = olSheet.getDataRange().getValues();

  // Current week boundary (Monday 00:00)
  const thisWeekStart = getWeekStart();

  // Aggregate this-week units and yeses per email
  const personAgg = {};
  Object.keys(roster).forEach(function(email) {
    if (roster[email].deactivated) return;
    personAgg[email] = { units: 0, yeses: 0 };
  });

  for (var i = 1; i < olData.length; i++) {
    var row = olData[i];
    var email = String(row[OL.EMAIL] || '').trim().toLowerCase();
    if (!email || !personAgg[email]) continue;

    var rawDate = row[OL.DATE_OF_SALE];
    if (!rawDate) continue;
    var saleDate = new Date(rawDate);
    if (isNaN(saleDate.getTime())) continue;
    saleDate.setHours(0, 0, 0, 0);

    if (saleDate >= thisWeekStart) {
      personAgg[email].units += Number(row[OL.UNITS]) || 0;
      personAgg[email].yeses += Number(row[OL.YESES]) || 0;
    }
  }

  // Role display labels
  var ROLE_LABELS = {
    owner: 'Owner', manager: 'Manager', jd: 'Jr. Director',
    l1: 'Team Leader', rep: 'Client Rep'
  };
  var NON_SALES = { superadmin: true, admin: true };

  // Build sorted individual list (exclude non-sales, deactivated, owners w/ 0 units)
  var individuals = [];
  Object.keys(personAgg).forEach(function(email) {
    var info = roster[email];
    var rank = (info.rank || 'rep').toLowerCase();
    if (NON_SALES[rank]) return;
    var agg = personAgg[email];
    if (rank === 'owner' && agg.units === 0) return;
    individuals.push({
      name: info.name,
      rank: ROLE_LABELS[rank] || rank,
      team: info.team,
      units: agg.units,
      yeses: agg.yeses
    });
  });
  individuals.sort(function(a, b) { return b.units - a.units || b.yeses - a.yeses; });

  // Week totals
  var weekUnits = 0, weekYeses = 0;
  individuals.forEach(function(p) { weekUnits += p.units; weekYeses += p.yeses; });

  // Team aggregation
  var teamAgg = {};
  individuals.forEach(function(p) {
    var t = p.team || 'Unassigned';
    if (!teamAgg[t]) teamAgg[t] = { name: t, emoji: '', units: 0, yeses: 0 };
    teamAgg[t].units += p.units;
    teamAgg[t].yeses += p.yeses;
  });

  // Set emojis from _Teams
  Object.keys(teams).forEach(function(tid) {
    var t = teams[tid];
    if (teamAgg[t.name]) teamAgg[t.name].emoji = t.emoji || '';
  });

  var sortedTeams = [];
  Object.keys(teamAgg).forEach(function(k) {
    if (k !== 'Unassigned') sortedTeams.push(teamAgg[k]);
  });
  sortedTeams.sort(function(a, b) { return b.units - a.units || b.yeses - a.yeses; });

  // Format date
  var now = new Date();
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var dateStr = months[now.getMonth()] + ' ' + now.getDate() + ', ' + now.getFullYear();

  return {
    weekUnits: weekUnits,
    weekYeses: weekYeses,
    topIndividuals: individuals.slice(0, 3),
    topTeams: sortedTeams.slice(0, 3),
    date: dateStr
  };
}

function buildLeaderboardHtml(data) {
  var wY = data.weekYeses, wU = data.weekUnits;
  var top3 = data.topIndividuals || [];
  var topT = data.topTeams || [];
  var dateStr = data.date || '';

  // Podium order: 2nd, 1st, 3rd (visual: silver-gold-bronze)
  var indOrder = top3.length >= 3 ? [top3[1], top3[0], top3[2]] : top3;
  var teamOrder = topT.length >= 3 ? [topT[1], topT[0], topT[2]] : topT;

  var medals = ['\uD83E\uDD48', '\uD83E\uDD47', '\uD83E\uDD49']; // 🥈🥇🥉
  var topGrad = [
    'linear-gradient(90deg,#C0C0C0,#94a3b8)',
    'linear-gradient(90deg,#FFD700,#fbbf24)',
    'linear-gradient(90deg,#CD7F32,#b45309)'
  ];

  function card(item, idx, isTeam) {
    var isGold = idx === 1;
    var bg = isGold ? 'linear-gradient(180deg,rgba(255,215,0,0.06) 0%,#F5F2EE 60%)' : 'rgba(255,255,255,0.5)';
    var bdr = isGold ? 'rgba(255,215,0,0.4)' : 'rgba(0,0,0,0.2)';
    var pad = isGold ? '32px 20px 20px' : '24px 20px 20px';
    var sub = isTeam
      ? '<div style="font-size:28px;line-height:1;">' + (item.emoji || '') + '</div>'
      : '<div style="font-family:Helvetica Neue,Inter,sans-serif;font-size:11px;font-weight:600;letter-spacing:0.5px;text-transform:uppercase;color:#708090;">' + (item.rank || '') + '</div>';
    return '<div style="flex:1;max-width:340px;border-radius:12px;padding:' + pad + ';display:flex;flex-direction:column;align-items:center;gap:8px;position:relative;overflow:hidden;border:1px solid ' + bdr + ';background:' + bg + ';">' +
      '<div style="position:absolute;top:0;left:0;right:0;height:3px;background:' + topGrad[idx] + ';"></div>' +
      '<div style="font-size:28px;line-height:1;">' + medals[idx] + '</div>' +
      '<div style="font-family:Inter,sans-serif;font-size:22px;font-weight:700;letter-spacing:1px;color:#242124;text-align:center;">' + item.name + '</div>' +
      sub +
      '<div style="display:flex;gap:20px;margin-top:8px;">' +
        '<div style="display:flex;flex-direction:column;align-items:center;gap:2px;">' +
          '<div style="font-family:Inter,sans-serif;font-size:32px;font-weight:700;line-height:1;color:#43B3AE;">' + item.units + '</div>' +
          '<div style="font-family:Helvetica Neue,Inter,sans-serif;font-size:10px;font-weight:600;letter-spacing:0.5px;text-transform:uppercase;color:#708090;">Units</div></div>' +
        '<div style="display:flex;flex-direction:column;align-items:center;gap:2px;">' +
          '<div style="font-family:Inter,sans-serif;font-size:32px;font-weight:700;line-height:1;color:#4A5568;">' + item.yeses + '</div>' +
          '<div style="font-family:Helvetica Neue,Inter,sans-serif;font-size:10px;font-weight:600;letter-spacing:0.5px;text-transform:uppercase;color:#708090;">Yeses</div></div>' +
      '</div></div>';
  }

  function podiumRow(items, isTeam) {
    var h = '';
    for (var i = 0; i < items.length; i++) h += card(items[i], i, isTeam);
    return '<div style="display:flex;gap:16px;align-items:flex-end;justify-content:center;margin-bottom:28px;">' + h + '</div>';
  }

  return '<div style="width:800px;background:#FEFAF3;padding:32px 32px 20px;font-family:Inter,sans-serif;">' +
    '<div style="display:flex;align-items:center;background:rgba(255,255,255,0.5);border:1px solid rgba(0,0,0,0.3);border-radius:14px;margin-bottom:32px;overflow:hidden;">' +
      '<div style="flex:1;display:flex;flex-direction:column;align-items:center;padding:24px 40px;gap:6px;">' +
        '<div style="font-family:Helvetica Neue,Inter,sans-serif;font-size:12px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#708090;">Week Yeses</div>' +
        '<div style="font-family:Inter,sans-serif;font-size:56px;font-weight:700;line-height:1;color:#242124;">' + wY + '</div></div>' +
      '<div style="width:1px;height:60px;background:rgba(0,0,0,0.3);"></div>' +
      '<div style="flex:1;display:flex;flex-direction:column;align-items:center;padding:24px 40px;gap:6px;">' +
        '<div style="font-family:Helvetica Neue,Inter,sans-serif;font-size:12px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#708090;">Week Units</div>' +
        '<div style="font-family:Inter,sans-serif;font-size:56px;font-weight:700;line-height:1;color:#43B3AE;">' + wU + '</div></div>' +
    '</div>' +
    '<div style="display:flex;align-items:center;gap:16px;margin-bottom:20px;">' +
      '<div style="font-family:Helvetica Neue,Inter,sans-serif;font-size:16px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#242124;">Top Performers</div>' +
      '<div style="flex:1;height:1px;background:rgba(0,0,0,0.15);"></div>' +
      '<div style="font-size:13px;font-style:italic;color:#708090;">This week</div></div>' +
    '<div style="font-family:Helvetica Neue,Inter,sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:#708090;margin-bottom:12px;">Individuals</div>' +
    podiumRow(indOrder, false) +
    '<div style="font-family:Helvetica Neue,Inter,sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:#708090;margin-bottom:12px;">Teams</div>' +
    podiumRow(teamOrder, true) +
    '<div style="text-align:center;padding-top:8px;border-top:1px solid rgba(0,0,0,0.08);">' +
      '<div style="font-size:11px;color:#708090;letter-spacing:0.5px;">End of Day Snapshot \u00B7 ' + dateStr + '</div></div>' +
  '</div>';
}


// === SALE SUBMISSION (30-column _Sales schema) ===

function writeAddSale(body, ss, officeId) {
  var sheet = getOrCreateSheet(ss, officeTab(TAB.SALES, officeId), TAB.SALES);

  // Validate required fields
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

  // Calculate product counts
  var air = Number(body.air) || 0;
  var newPhones = Number(body.newPhones) || 0;
  var byods = Number(body.byods) || 0;
  var cell = newPhones + byods;
  var fiber = Number(body.fiber) || 0;
  var voipQty = Number(body.voipQty) || 0;
  var dtv = Number(body.dtv) || 0;

  // YESES: count of products sold
  var yeses = 0;
  if (air > 0) yeses++;
  if (cell > 0) yeses++;
  if (fiber > 0) yeses++;
  if (voipQty > 0) yeses++;
  if (dtv > 0) yeses++;

  // UNITS: sum of products (DTV excluded per config)
  var units = air + cell + fiber + voipQty;

  // Build 30-column row (A–AD)
  var newRow = [
    new Date(),                                                    // 0  Timestamp
    email,                                                         // 1  Email
    String(body.repName || '').trim(),                              // 2  Rep Name
    new Date(dateOfSale + 'T12:00:00'),                            // 3  Date of Sale
    String(body.campaign || '').trim(),                             // 4  Campaign
    String(body.dsi || '').trim(),                                  // 5  DSI
    String(body.accountType || '').trim(),                          // 6  Account Type
    String(body.clientName || '').trim(),                           // 7  Client Name
    body.trainee ? 'Yes' : 'No',                                   // 8  Trainee
    body.trainee ? String(body.traineeName || '').trim() : '',      // 9  Trainee Name
    air,                                                           // 10 Air
    newPhones,                                                     // 11 New Phones
    byods,                                                         // 12 BYODs
    cell,                                                          // 13 Cell
    fiber,                                                         // 14 Fiber
    String(body.fiberPackage || '').trim(),                         // 15 Fiber Package
    String(body.installDate || '').trim(),                          // 16 Install Date
    voipQty,                                                       // 17 VoIP Qty
    dtv,                                                           // 18 DTV
    String(body.dtvPackage || '').trim(),                           // 19 DTV Package
    String(body.oomaPackage || '').trim(),                          // 20 Ooma Package
    String(body.accountNotes || '').trim(),                         // 21 Account Notes
    body.activationSupport ? 'Yes' : 'No',                         // 22 Activation Support
    teamEmoji,                                                     // 23 Team Emoji
    yeses,                                                         // 24 Yeses
    units,                                                         // 25 Units
    'Pending',                                                     // 26 Status
    '',                                                            // 27 Notes
    '',                                                            // 28 Paid Out
    '[]',                                                          // 29 Tickets
    String(body.orderChannel || 'Sara').trim(),                     // 30 Order Channel
    String(body.codesUsedBy || '').trim().toLowerCase()             // 31 Codes Used By
  ];

  sheet.appendRow(newRow);

  // Bust the Tableau cache so fresh data loads
  try {
    CacheService.getScriptCache().remove('tableauSummary_v6_' + officeId);
  } catch (e) { /* non-critical */ }

  // Fire Discord/GroupMe webhook server-side (fire-and-forget)
  var webhookDebug = '';
  try {
    webhookDebug = _fireWebhook(body, units, teamEmoji);
  } catch (e) { webhookDebug = 'CATCH: ' + e.message; }

  return {
    ok: true,
    rowIndex: sheet.getLastRow(),
    units: units,
    yeses: yeses,
    webhookDebug: webhookDebug || 'no-return'
  };
}


// === DISCORD / GROUPME WEBHOOK (server-side) ===

function _fireWebhook(body, units, teamEmoji) {
  var platform = String(body.chatPlatform || 'discord').toLowerCase();
  var webhookUrl = String(body.discordWebhookUrl || '').trim();
  if (!webhookUrl || platform === 'none') return 'SKIP: url=' + !!webhookUrl + ' platform=' + platform;

  var bold = (platform === 'discord') ? '**' : '';
  var repName = String(body.repName || '').trim();
  var campaign = String(body.campaign || '').trim();
  var msg = '';

  if (campaign === 'attb2b') {
    msg += bold + repName + bold + ' made a sale with AT&T: B2B!\n';
    msg += (body.accountType || 'Business') + ' Account\n';
    msg += String(body.dsi || '') + '\n';
    if (Number(body.air) > 0) msg += '• Internet Air\n';
    var np = Number(body.newPhones) || 0;
    var by = Number(body.byods) || 0;
    if (np > 0 || by > 0) msg += '• ' + np + ' New Phone(s)|' + by + ' BYOD(s)\n';
    if (Number(body.fiber) > 0) msg += '• ' + (body.fiberPackage || 'Fiber') + '\n';
    var vq = Number(body.voipQty) || 0;
    if (vq > 0) msg += '• ' + vq + ' VoIP(s)\n';
    if (Number(body.dtv) > 0) msg += '• DIRECTV ' + (body.dtvPackage || '') + '\n';
  } else if (campaign === 'ooma') {
    msg += bold + repName + bold + ' made a sale with Ooma!\n';
    msg += String(body.clientName || '') + '\n';
    msg += '• ' + (body.oomaPackage || 'Ooma Pro') + '\n';
  }

  var tags = String(body.hashtags || '').trim();
  if (tags) msg += tags + '\n';

  if (teamEmoji && units > 0) {
    var count = Math.min(units, 20);
    for (var i = 0; i < count; i++) msg += teamEmoji;
  }

  msg = msg.trim();
  if (!msg) return 'SKIP: empty msg';

  var fetchPayload, url;
  if (platform === 'groupme') {
    url = 'https://api.groupme.com/v3/bots/post';
    fetchPayload = JSON.stringify({ bot_id: webhookUrl, text: msg });
  } else {
    url = webhookUrl;
    fetchPayload = JSON.stringify({ content: msg });
  }

  var fetchOpts = {
    method: 'post',
    contentType: 'application/json',
    payload: fetchPayload,
    muteHttpExceptions: true
  };

  // Retry up to 2 times on failure (rate-limit 429, server errors 5xx, or fetch exceptions)
  var maxAttempts = 3;
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      var resp = UrlFetchApp.fetch(url, fetchOpts);
      var code = resp.getResponseCode();
      if (code >= 200 && code < 300) {
        return 'HTTP ' + code + (attempt > 1 ? ' (retry ' + (attempt - 1) + ')' : '') + ' | msg=' + msg.substring(0, 50);
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
      return 'HTTP ' + code + ' attempt ' + attempt + ' | ' + resp.getContentText().substring(0, 80);
    } catch (e) {
      if (attempt < maxAttempts) {
        Utilities.sleep(2000);
        continue;
      }
      return 'FETCH_ERROR attempt ' + attempt + ': ' + e.message;
    }
  }
  return 'EXHAUSTED_RETRIES';
}


// === REPLAY WEBHOOK FOR MISSED NOTIFICATIONS ===

function replayWebhook(body, ss, officeId) {
  var dsi = String(body.dsi || '').trim();
  if (!dsi) return { error: 'Missing dsi' };

  var webhookUrl = String(body.discordWebhookUrl || '').trim();
  var chatPlatform = String(body.chatPlatform || 'discord').toLowerCase();
  if (!webhookUrl) return { error: 'Missing discordWebhookUrl' };

  var sheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!sheet) return { error: 'Sales tab not found' };

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var found = null;
  // Col 5 (index 5) = DSI
  for (var r = data.length - 1; r >= 1; r--) {
    if (String(data[r][5]).trim().toUpperCase() === dsi.toUpperCase()) {
      found = data[r];
      break;
    }
  }
  if (!found) return { error: 'DSI not found: ' + dsi };

  // Rebuild body from row data
  var replayBody = {
    repName:          String(found[2] || ''),
    campaign:         String(found[4] || ''),
    dsi:              String(found[5] || ''),
    accountType:      String(found[6] || ''),
    clientName:       String(found[7] || ''),
    air:              Number(found[10]) || 0,
    newPhones:        Number(found[11]) || 0,
    byods:            Number(found[12]) || 0,
    fiber:            Number(found[14]) || 0,
    fiberPackage:     String(found[15] || ''),
    voipQty:          Number(found[17]) || 0,
    dtv:              Number(found[18]) || 0,
    dtvPackage:       String(found[19] || ''),
    oomaPackage:      String(found[20] || ''),
    hashtags:         '',
    discordWebhookUrl: webhookUrl,
    chatPlatform:     chatPlatform
  };

  var units = Number(found[25]) || 0;
  var teamEmoji = String(found[23] || '');

  var result = '';
  try {
    result = _fireWebhook(replayBody, units, teamEmoji);
  } catch (e) {
    result = 'ERROR: ' + e.message;
  }

  return { ok: true, dsi: dsi, webhook: result };
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


// === ONE-TIME MIGRATION: Old tabs → Per-office _off_001 tabs ===
// Can be called via API: POST { action: 'migrateFromLegacy' }
// Or run directly from Apps Script editor: migrateFromLegacy()
//
// What it does:
// 1. Reads old "Order Log" (42 sparse cols) → writes to _Sales_off_001 (30 clean cols)
// 2. Copies data from _Roster, _Teams, etc. into _Roster_off_001, _Teams_off_001, etc.
// 3. Renames old tabs with _Legacy suffix (preserves data, doesn't delete)
//
// Safe to run if new _off_001 tabs already exist — overwrites with old data.

function migrateFromLegacy(ss, officeId) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!officeId) officeId = 'off_001';
  var log = [];

  log.push('[Migration] Starting legacy → per-office migration for ' + officeId);

  // ── Step 1: Migrate Order Log → _Sales_off_001 ──
  var oldOL = ss.getSheetByName('Order Log');
  var salesTabName = officeTab(TAB.SALES, officeId);
  var salesSheet = ss.getSheetByName(salesTabName);
  if (!salesSheet) salesSheet = getOrCreateSheet(ss, salesTabName, TAB.SALES);

  if (oldOL) {
    var oldData = oldOL.getDataRange().getValues();
    var oldRowCount = oldData.length - 1;
    log.push('[Migration] Order Log: ' + oldRowCount + ' data rows');

    if (oldRowCount > 0) {
      // Map old 42-col sparse schema → new 30-col sequential schema
      var newRows = [];
      for (var i = 1; i < oldData.length; i++) {
        var old = oldData[i];
        var email = String(old[OL_LEGACY.EMAIL] || '').trim();
        if (!email) continue;  // skip blank rows

        newRows.push([
          old[OL_LEGACY.TIMESTAMP] || '',                        // 0  Timestamp
          email.toLowerCase(),                                    // 1  Email
          String(old[OL_LEGACY.REP_NAME] || '').trim(),          // 2  Rep Name
          old[OL_LEGACY.DATE_OF_SALE] || '',                     // 3  Date of Sale
          'attb2b',                                               // 4  Campaign
          String(old[OL_LEGACY.DSI] || '').trim(),               // 5  DSI
          '',                                                     // 6  Account Type
          '',                                                     // 7  Client Name
          String(old[OL_LEGACY.TRAINEE] || '').trim(),           // 8  Trainee
          String(old[OL_LEGACY.TRAINEE_NAME] || '').trim(),      // 9  Trainee Name
          Number(old[OL_LEGACY.AIR]) || 0,                       // 10 Air
          0,                                                      // 11 New Phones
          0,                                                      // 12 BYODs
          Number(old[OL_LEGACY.CELL]) || 0,                      // 13 Cell
          Number(old[OL_LEGACY.FIBER]) || 0,                     // 14 Fiber
          '',                                                     // 15 Fiber Package
          '',                                                     // 16 Install Date
          Number(old[OL_LEGACY.VOIP_QTY]) || 0,                 // 17 VoIP Qty
          Number(old[OL_LEGACY.DTV]) || 0,                       // 18 DTV
          '',                                                     // 19 DTV Package
          '',                                                     // 20 Ooma Package
          '',                                                     // 21 Account Notes
          '',                                                     // 22 Activation Support
          String(old[OL_LEGACY.TEAM_EMOJI] || '').trim(),        // 23 Team Emoji
          Number(old[OL_LEGACY.YESES]) || 0,                     // 24 Yeses
          Number(old[OL_LEGACY.UNITS]) || 0,                     // 25 Units
          String(old[OL_LEGACY.STATUS] || 'Pending').trim(),     // 26 Status
          String(old[OL_LEGACY.NOTES] || '').trim(),             // 27 Notes
          String(old[OL_LEGACY.PAID_OUT] || '').trim(),          // 28 Paid Out
          String(old[OL_LEGACY.TICKETS] || '[]').trim()          // 29 Tickets
        ]);
      }

      if (newRows.length > 0) {
        // Clear any existing data in _Sales_off_001 (keep header), then write
        if (salesSheet.getLastRow() > 1) {
          salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, 30).clearContent();
        }
        salesSheet.getRange(2, 1, newRows.length, 30).setValues(newRows);
        log.push('[Migration] Wrote ' + newRows.length + ' orders to ' + salesTabName);
      }
    }

    // Rename old Order Log → _OrderLog_Legacy
    oldOL.setName('_OrderLog_Legacy');
    log.push('[Migration] Renamed "Order Log" → "_OrderLog_Legacy"');
  } else {
    log.push('[Migration] No "Order Log" tab found — skipping sales migration');
  }

  // ── Step 2: Copy data from old tabs into new _off_001 tabs ──
  var copyMap = [
    { old: '_Roster',             newTab: officeTab(TAB.ROSTER, officeId) },
    { old: '_Teams',              newTab: officeTab(TAB.TEAMS, officeId) },
    { old: '_TeamCustomizations', newTab: officeTab(TAB.TEAM_CUSTOM, officeId) },
    { old: '_OrderOverrides',     newTab: officeTab(TAB.OVERRIDES, officeId) },
    { old: '_UnlockRequests',     newTab: officeTab(TAB.UNLOCKS, officeId) },
    { old: '_Settings',           newTab: officeTab(TAB.SETTINGS, officeId) }
  ];

  copyMap.forEach(function(entry) {
    var oldSheet = ss.getSheetByName(entry.old);
    if (!oldSheet) {
      log.push('[Migration] ' + entry.old + ' not found — skipping');
      return;
    }

    var newSheet = ss.getSheetByName(entry.newTab);
    if (!newSheet) {
      // New tab doesn't exist — just rename the old one directly
      oldSheet.setName(entry.newTab);
      log.push('[Migration] Renamed: ' + entry.old + ' → ' + entry.newTab);
      return;
    }

    // Both tabs exist — copy data from old into new (skip old header row)
    var oldData = oldSheet.getDataRange().getValues();
    if (oldData.length > 1) {
      var dataRows = oldData.slice(1);
      var cols = dataRows[0].length;
      // Clear existing data in new tab (keep header)
      if (newSheet.getLastRow() > 1) {
        newSheet.getRange(2, 1, newSheet.getLastRow() - 1, cols).clearContent();
      }
      newSheet.getRange(2, 1, dataRows.length, cols).setValues(dataRows);
      log.push('[Migration] Copied ' + dataRows.length + ' rows: ' + entry.old + ' → ' + entry.newTab);
    } else {
      log.push('[Migration] ' + entry.old + ' has no data rows — skipping copy');
    }

    // Rename old tab to _Legacy suffix
    oldSheet.setName(entry.old + '_Legacy');
    log.push('[Migration] Renamed: ' + entry.old + ' → ' + entry.old + '_Legacy');
  });

  // ── Step 3: Summary ──
  var salesFinal = ss.getSheetByName(salesTabName);
  var rosterFinal = ss.getSheetByName(officeTab(TAB.ROSTER, officeId));

  var summary = {
    success: true,
    officeId: officeId,
    salesRows: salesFinal ? salesFinal.getLastRow() - 1 : 0,
    rosterRows: rosterFinal ? rosterFinal.getLastRow() - 1 : 0,
    log: log
  };

  log.push('[Migration] Complete! Sales: ' + summary.salesRows + ', Roster: ' + summary.rosterRows);
  Logger.log(log.join('\n'));

  return summary;
}


// ═══════════════════════════════════════════════════════════════
// migrateFromExternal — Import sales data from a DIFFERENT Google Sheet
// ═══════════════════════════════════════════════════════════════
//
// Called via doPost: action='migrateFromExternal'
// Required body params:
//   sourceSheetId  — Google Sheet ID of the external source (e.g. Ignite)
//   officeId       — target office ID (e.g. 'off_002')
// Optional body params:
//   sourceTabName  — name of Order Log tab in source (default 'Order Log')
//
// Maps external 38-column Ignite format → new 30-column _Sales schema.
// Creates all 7 per-office tabs in the target campaign sheet.
// Copies _Roster and _Teams data if they exist in source.

function migrateFromExternal(body, ss) {
  var sourceSheetId = body.sourceSheetId;
  var officeId = body.officeId;
  var sourceTabName = body.sourceTabName || 'Order Log';
  var salesOnly = body.salesOnly || false;  // if true, skip roster/teams (preserve existing)

  if (!sourceSheetId) return { error: 'Missing sourceSheetId' };
  if (!officeId) return { error: 'Missing officeId' };
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  var log = [];
  log.push('[Migration] External migration: source=' + sourceSheetId + ', target officeId=' + officeId);

  // Open source spreadsheet
  var sourceSS;
  try {
    sourceSS = SpreadsheetApp.openById(sourceSheetId);
  } catch (e) {
    return { error: 'Cannot open source sheet: ' + e.message };
  }

  // ── Step 1: Migrate Order Log → _Sales_{officeId} ──
  var sourceOL = sourceSS.getSheetByName(sourceTabName);
  if (!sourceOL) return { error: 'Source tab "' + sourceTabName + '" not found' };

  var salesTabName = officeTab(TAB.SALES, officeId);
  var salesSheet = getOrCreateSheet(ss, salesTabName, TAB.SALES);

  var sourceData = sourceOL.getDataRange().getValues();
  var sourceRowCount = sourceData.length - 1;
  log.push('[Migration] Source Order Log: ' + sourceRowCount + ' data rows');

  if (sourceRowCount > 0) {
    // Ignite source column indices (0-based, 38 columns A–AL)
    var IGN = {
      PRODUCTION_POST: 0,
      TIMESTAMP: 1,
      EMAIL: 2,
      DATE_OF_SALE: 3,
      REP_NAME: 4,
      TRAINEE: 5,           // "Did you train someone?" Yes/No
      TRAINEE_NAME: 6,
      CAMPAIGN: 7,           // "Which Campaign?" e.g. "AT&T: B2B", "Ooma"
      DSI: 8,
      ACCOUNT_NOTES: 9,
      ACCOUNT_TYPE: 10,      // "Type of Account" Consumer/Business
      AIR_SOLD: 11,          // "Was Internet Air Sold?" Yes/No
      WIRELESS_SOLD: 12,     // "Were Wireless Lines Sold?" Yes/No
      NEW_PHONES: 13,        // Quantity of New Phones
      BYODS: 14,             // Quantity of BYODs
      FIBER_SOLD: 15,        // "Was Fiber Optic Internet Sold?" Yes/No
      FIBER_PACKAGE: 16,     // "Which Package Was Sold?"
      INSTALL_DATE: 17,
      VOIP_SOLD: 18,         // "Were VoIP Lines Sold?" Yes/No
      VOIP_QTY: 19,          // Quantity Sold (VoIP)
      DTV_SOLD: 20,          // "Was DIRECTV Sold?" Yes/No
      DTV_PACKAGE: 21,       // Package Sold (DTV)
      CLIENT_NAME: 22,
      OOMA_PACKAGE: 23,      // Which Ooma package?
      // 24-29: computed/display columns (not used)
      FIBER_COUNTER: 30,     // 0/1
      AIR_COUNTER: 31,       // 0/1
      DTV_COUNTER: 32,       // 0/1
      YES_COUNTER: 33,       // total yeses
      NUM_LINES: 34,         // # of wireless lines
      VOIP_COUNTER: 35,      // VoIP qty (may equal VOIP_QTY)
      UNITS: 36,             // total units
      UNITS_TODAY: 37        // units for today (not migrated)
    };

    var newRows = [];
    for (var i = 1; i < sourceData.length; i++) {
      var src = sourceData[i];
      var email = String(src[IGN.EMAIL] || '').trim();
      if (!email) continue;  // skip blank rows

      // Normalize campaign: "AT&T: B2B" / "AT&T B2B" → "attb2b", "Ooma" → "ooma"
      var rawCampaign = String(src[IGN.CAMPAIGN] || '').trim().toLowerCase();
      var campaign = 'attb2b';
      if (rawCampaign.indexOf('ooma') >= 0) campaign = 'ooma';

      var newPhones = Number(src[IGN.NEW_PHONES]) || 0;
      var byods = Number(src[IGN.BYODS]) || 0;
      var cell = newPhones + byods;

      // Use counter columns for binary flags (0/1)
      var air = Number(src[IGN.AIR_COUNTER]) || 0;
      var fiber = Number(src[IGN.FIBER_COUNTER]) || 0;
      var dtv = Number(src[IGN.DTV_COUNTER]) || 0;
      var voipQty = Number(src[IGN.VOIP_COUNTER]) || Number(src[IGN.VOIP_QTY]) || 0;

      newRows.push([
        src[IGN.TIMESTAMP] || '',                                  // 0  Timestamp
        email.toLowerCase(),                                        // 1  Email
        String(src[IGN.REP_NAME] || '').trim(),                    // 2  Rep Name
        src[IGN.DATE_OF_SALE] || '',                               // 3  Date of Sale
        campaign,                                                   // 4  Campaign
        String(src[IGN.DSI] || '').trim(),                         // 5  DSI
        String(src[IGN.ACCOUNT_TYPE] || '').trim(),                // 6  Account Type
        String(src[IGN.CLIENT_NAME] || '').trim(),                 // 7  Client Name
        String(src[IGN.TRAINEE] || '').trim(),                     // 8  Trainee
        String(src[IGN.TRAINEE_NAME] || '').trim(),                // 9  Trainee Name
        air,                                                        // 10 Air
        newPhones,                                                  // 11 New Phones
        byods,                                                      // 12 BYODs
        cell,                                                       // 13 Cell
        fiber,                                                      // 14 Fiber
        String(src[IGN.FIBER_PACKAGE] || '').trim(),               // 15 Fiber Package
        src[IGN.INSTALL_DATE] || '',                                // 16 Install Date
        voipQty,                                                    // 17 VoIP Qty
        dtv,                                                        // 18 DTV
        String(src[IGN.DTV_PACKAGE] || '').trim(),                 // 19 DTV Package
        String(src[IGN.OOMA_PACKAGE] || '').trim(),                // 20 Ooma Package
        String(src[IGN.ACCOUNT_NOTES] || '').trim(),               // 21 Account Notes
        '',                                                         // 22 Activation Support
        '',                                                         // 23 Team Emoji
        Number(src[IGN.YES_COUNTER]) || 0,                         // 24 Yeses
        Number(src[IGN.UNITS]) || 0,                               // 25 Units
        'Pending',                                                  // 26 Status
        '',                                                         // 27 Notes
        '',                                                         // 28 Paid Out
        '[]'                                                        // 29 Tickets
      ]);
    }

    if (newRows.length > 0) {
      // Clear any existing data in target (keep header), then write
      if (salesSheet.getLastRow() > 1) {
        salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, 30).clearContent();
      }
      salesSheet.getRange(2, 1, newRows.length, 30).setValues(newRows);
      log.push('[Migration] Wrote ' + newRows.length + ' orders to ' + salesTabName);
    } else {
      log.push('[Migration] No valid rows found in source Order Log');
    }
  }

  // ── Step 2: Copy _Roster → _Roster_{officeId} (skip if salesOnly) ──
  if (!salesOnly) {
    var sourceRoster = sourceSS.getSheetByName('_Roster');
    if (sourceRoster) {
      var rosterTabName = officeTab(TAB.ROSTER, officeId);
      var rosterSheet = getOrCreateSheet(ss, rosterTabName, TAB.ROSTER);

      var rosterData = sourceRoster.getDataRange().getValues();
      if (rosterData.length > 1) {
        var rosterRows = rosterData.slice(1);
        var rosterCols = rosterRows[0].length;
        if (rosterSheet.getLastRow() > 1) {
          rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, rosterCols).clearContent();
        }
        rosterSheet.getRange(2, 1, rosterRows.length, rosterCols).setValues(rosterRows);
        log.push('[Migration] Copied ' + rosterRows.length + ' roster rows to ' + rosterTabName);
      } else {
        log.push('[Migration] Source _Roster has no data rows');
      }
    } else {
      log.push('[Migration] No _Roster tab in source — creating empty ' + officeTab(TAB.ROSTER, officeId));
      getOrCreateSheet(ss, officeTab(TAB.ROSTER, officeId), TAB.ROSTER);
    }
  } else {
    log.push('[Migration] salesOnly=true — skipping roster (preserving existing)');
  }

  // ── Step 3: Copy _Teams if exists (skip if salesOnly) ──
  if (!salesOnly) {
    var sourceTeams = sourceSS.getSheetByName('_Teams');
    if (sourceTeams) {
      var teamsTabName = officeTab(TAB.TEAMS, officeId);
      var teamsSheet = getOrCreateSheet(ss, teamsTabName, TAB.TEAMS);

      var teamsData = sourceTeams.getDataRange().getValues();
      if (teamsData.length > 1) {
        var teamsRows = teamsData.slice(1);
        var teamsCols = teamsRows[0].length;
        if (teamsSheet.getLastRow() > 1) {
          teamsSheet.getRange(2, 1, teamsSheet.getLastRow() - 1, teamsCols).clearContent();
        }
        teamsSheet.getRange(2, 1, teamsRows.length, teamsCols).setValues(teamsRows);
        log.push('[Migration] Copied ' + teamsRows.length + ' teams to ' + teamsTabName);
      }
    } else {
      log.push('[Migration] No _Teams tab in source — creating empty ' + officeTab(TAB.TEAMS, officeId));
      getOrCreateSheet(ss, officeTab(TAB.TEAMS, officeId), TAB.TEAMS);
    }
  } else {
    log.push('[Migration] salesOnly=true — skipping teams (preserving existing)');
  }

  // ── Step 4: Create remaining empty tabs ──
  [TAB.TEAM_CUSTOM, TAB.OVERRIDES, TAB.UNLOCKS, TAB.SETTINGS].forEach(function(base) {
    getOrCreateSheet(ss, officeTab(base, officeId), base);
    log.push('[Migration] Ensured tab: ' + officeTab(base, officeId));
  });

  // ── Step 5: Summary ──
  var salesFinal = ss.getSheetByName(salesTabName);
  var rosterFinal = ss.getSheetByName(officeTab(TAB.ROSTER, officeId));

  var summary = {
    success: true,
    officeId: officeId,
    salesRows: salesFinal ? salesFinal.getLastRow() - 1 : 0,
    rosterRows: rosterFinal ? rosterFinal.getLastRow() - 1 : 0,
    log: log
  };

  log.push('[Migration] Complete! Sales: ' + summary.salesRows + ', Roster: ' + summary.rosterRows);
  Logger.log(log.join('\n'));

  return summary;
}


// ═══════════════════════════════════════════════════════
// LUNA DATA MIGRATION — Run ONCE from Apps Script editor
// ═══════════════════════════════════════════════════════
// Source: Luna's legacy sheet 1pOi6p5gsHCdn_SZpNwWJLzwLmLCn5D-6gKAHccAaqqU
// Target: _Sales_off_003 in campaign sheet
//
// Luna's source columns (0-indexed):
//  0: Production Post       1: Timestamp         2: Email Address
//  3: Date of Sale          4: Representative's Name
//  5: Was this sale made under someone else's codes?
//  6: Whose codes were used for the sale?
//  7: Which Campaign?       8: How was the order processed?
//  9: DSI Number           10: Ticket Number     11: Type of Account
// 12: Additional Account Notes      13: Was Internet Air Sold?
// 14: Were Wireless Lines Sold?     15: Quantity of New Phones
// 16: Quantity of BYODs             17: Was Fiber Optic Internet Sold?
// 18: Which Package Was Sold?       19: Install Date
// 20: Were VoIP Lines Sold?        21: Quantity Sold (VoIP)
// 22: Was DIRECTV Sold?            23: Package Sold (DTV)
// 24: Client Name                  25: Which package was sold? (Ooma)
// 26: Sale Screenshot              27: #'s
// 28: Package                      29: Rep and Trainee
// 30: AT&T Order    31: Vonage Order    32: Fiber Counter
// 33: Air Counter   34: DTV Counter     35: Yes Counter
// 36: # of Lines    37: VoIP Counter    38: Units    39: Units Today

function migrateLunaData() {
  var SOURCE_SHEET_ID = '1pOi6p5gsHCdn_SZpNwWJLzwLmLCn5D-6gKAHccAaqqU';
  var SOURCE_TAB = 'Production Post';
  var TARGET_OFFICE_ID = 'off_003';

  // Open source sheet
  var sourceSS = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  var sourceSheet = sourceSS.getSheetByName(SOURCE_TAB);
  if (!sourceSheet) {
    sourceSheet = sourceSS.getSheets()[0];
    Logger.log('[Luna] Tab not found, using first sheet: ' + sourceSheet.getName());
  }

  var sourceData = sourceSheet.getDataRange().getValues();
  Logger.log('[Luna] Source rows (incl header): ' + sourceData.length);

  // Target campaign sheet
  var targetSS = SpreadsheetApp.getActiveSpreadsheet();
  var salesTabName = officeTab(TAB.SALES, TARGET_OFFICE_ID);
  var targetSheet = getOrCreateSheet(targetSS, salesTabName, TAB.SALES);
  Logger.log('[Luna] Target tab: ' + salesTabName);

  // Build roster email → emoji map from _Roster_off_003
  var rosterSheet = targetSS.getSheetByName(officeTab(TAB.ROSTER, TARGET_OFFICE_ID));
  var emojiByEmail = {};
  if (rosterSheet) {
    var rosterData = rosterSheet.getDataRange().getValues();
    var teamMaps = buildTeamEmojiMaps(targetSS, TARGET_OFFICE_ID);
    for (var r = 1; r < rosterData.length; r++) {
      var rEmail = String(rosterData[r][0] || '').trim().toLowerCase();
      var tName = String(rosterData[r][2] || '').trim();
      if (rEmail && tName && teamMaps.nameMap[tName]) {
        emojiByEmail[rEmail] = teamMaps.nameMap[tName];
      }
    }
    Logger.log('[Luna] Built emoji map for ' + Object.keys(emojiByEmail).length + ' reps');
  }

  var newRows = [];
  var skipped = 0;

  for (var i = 1; i < sourceData.length; i++) {
    var row = sourceData[i];

    var timestamp = row[1] || '';
    var email = String(row[2] || '').trim().toLowerCase();
    if (!timestamp && !email) { skipped++; continue; }

    // Parse trainee from "Rep and Trainee" (col 29)
    var repAndTrainee = String(row[29] || '').trim();
    var isTrainee = 'No';
    var traineeName = '';
    if (repAndTrainee) {
      var rtLow = repAndTrainee.toLowerCase();
      if (rtLow.indexOf('trainee') !== -1 || rtLow.indexOf('training') !== -1) {
        isTrainee = 'Yes';
        var tMatch = repAndTrainee.match(/trainee[:\s]+(.+)/i);
        if (tMatch) traineeName = tMatch[1].replace(/[()]/g, '').trim();
      }
    }

    // Order channel from "How was the order processed?" (col 8)
    var ocRaw = String(row[8] || '').trim();
    var orderChannel = (ocRaw.toLowerCase().indexOf('tower') !== -1) ? 'Tower' : 'Sara';

    // Codes used by (col 5 = yes/no, col 6 = whose codes)
    var cuAnswer = String(row[5] || '').trim().toLowerCase();
    var codesUsedBy = '';
    if (cuAnswer === 'yes' || cuAnswer === 'true') {
      codesUsedBy = String(row[6] || '').trim().toLowerCase();
    }

    // Numeric fields
    var air       = parseInt(row[33]) || 0;
    var newPhones = parseInt(row[15]) || 0;
    var byods     = parseInt(row[16]) || 0;
    var cell      = parseInt(row[36]) || 0;
    if (cell === 0) cell = newPhones + byods;
    var fiber     = parseInt(row[32]) || 0;
    var voipQty   = parseInt(row[37]) || 0;
    var dtv       = parseInt(row[34]) || 0;
    var yeses     = parseInt(row[35]) || 0;
    var units     = parseInt(row[38]) || 0;

    var teamEmoji = emojiByEmail[email] || '';

    // Build 32-column target row
    newRows.push([
      timestamp,                                         //  0 Timestamp
      email,                                             //  1 Email
      String(row[4] || '').trim(),                      //  2 Rep Name
      row[3] || '',                                     //  3 Date of Sale
      String(row[7] || 'attb2b').trim() || 'attb2b',   //  4 Campaign
      String(row[9] || '').trim(),                      //  5 DSI
      String(row[11] || '').trim(),                     //  6 Account Type
      String(row[24] || '').trim(),                     //  7 Client Name
      isTrainee,                                        //  8 Trainee
      traineeName,                                      //  9 Trainee Name
      air,                                              // 10 Air
      newPhones,                                        // 11 New Phones
      byods,                                            // 12 BYODs
      cell,                                             // 13 Cell
      fiber,                                            // 14 Fiber
      String(row[18] || '').trim(),                     // 15 Fiber Package
      row[19] || '',                                    // 16 Install Date
      voipQty,                                          // 17 VoIP Qty
      dtv,                                              // 18 DTV
      String(row[23] || '').trim(),                     // 19 DTV Package
      '',                                               // 20 Ooma Package
      String(row[12] || '').trim(),                     // 21 Account Notes
      '',                                               // 22 Activation Support
      teamEmoji,                                        // 23 Team Emoji
      yeses,                                            // 24 Yeses
      units,                                            // 25 Units
      'Pending',                                        // 26 Status
      '',                                               // 27 Notes
      '',                                               // 28 Paid Out
      '[]',                                             // 29 Tickets
      orderChannel,                                     // 30 Order Channel
      codesUsedBy                                       // 31 Codes Used By
    ]);
  }

  Logger.log('[Luna] Rows to migrate: ' + newRows.length + ', Skipped empty: ' + skipped);

  if (newRows.length === 0) {
    Logger.log('[Luna] No data rows found! Check source tab name.');
    return { success: false, error: 'No data rows found' };
  }

  // Append all rows after existing data
  var startRow = targetSheet.getLastRow() + 1;
  targetSheet.getRange(startRow, 1, newRows.length, 32).setValues(newRows);

  Logger.log('[Luna] Complete! Migrated ' + newRows.length + ' rows to ' + salesTabName);
  Logger.log('[Luna] Target sheet now has ' + targetSheet.getLastRow() + ' rows total');

  // Spot-check first 3 rows
  for (var j = 0; j < Math.min(3, newRows.length); j++) {
    Logger.log('[Spot ' + (j+1) + '] ' + newRows[j][2] + ' | ' + newRows[j][3] + ' | Units:' + newRows[j][25] + ' | ' + newRows[j][30]);
  }

  return {
    success: true,
    rowsMigrated: newRows.length,
    rowsSkipped: skipped,
    targetTab: salesTabName,
    totalRows: targetSheet.getLastRow() - 1
  };
}


// ═══════════════════════════════════════════════════════════════
// LEADERBOARD TEXT — Emoji-format message for Discord/GroupMe
// ═══════════════════════════════════════════════════════════════
// Called by AdminCode.gs centralized scheduler via ?action=leaderboardText
// Can also be called manually via postLeaderboardToChat() for testing.

/**
 * Build emoji-formatted leaderboard text from this office's data.
 * Code.gs aggregates from roster + sales (full roster, not just top 3).
 * @param {Spreadsheet} ss - The office spreadsheet
 * @param {string} officeId - The office ID for tab suffixes
 * @param {string} officeName - Display name for the header
 * @returns {string} The formatted message text
 */
function buildLeaderboardText(ss, officeId, officeName) {
  var roster = readRoster(ss, officeId);

  var olSheet = ss.getSheetByName(officeTab(TAB.SALES, officeId));
  if (!olSheet) return 'No sales data available.';
  var olData = olSheet.getDataRange().getValues();
  var thisWeekStart = getWeekStart();

  var personAgg = {};
  Object.keys(roster).forEach(function(email) {
    if (roster[email].deactivated) return;
    personAgg[email] = { units: 0, yeses: 0 };
  });

  for (var i = 1; i < olData.length; i++) {
    var row = olData[i];
    var email = String(row[OL.EMAIL] || '').trim().toLowerCase();
    if (!email || !personAgg[email]) continue;
    var rawDate = row[OL.DATE_OF_SALE];
    if (!rawDate) continue;
    var saleDate = new Date(rawDate);
    if (isNaN(saleDate.getTime())) continue;
    saleDate.setHours(0, 0, 0, 0);
    if (saleDate >= thisWeekStart) {
      personAgg[email].units += Number(row[OL.UNITS]) || 0;
      personAgg[email].yeses += Number(row[OL.YESES]) || 0;
    }
  }

  var NON_SALES = { superadmin: true, admin: true };
  var individuals = [];
  Object.keys(personAgg).forEach(function(em) {
    var info = roster[em];
    var rank = (info.rank || 'rep').toLowerCase();
    if (NON_SALES[rank]) return;
    var agg = personAgg[em];
    if (rank === 'owner' && agg.units === 0) return;
    individuals.push({ name: info.name, team: info.team, units: agg.units });
  });
  individuals.sort(function(a, b) { return b.units - a.units; });

  // Build individual rankings
  var medals = ['🥇', '🥈', '🥉'];
  var lines = [];
  lines.push('🔥 ' + officeName.toUpperCase() + ' WEEKLY 🔥');
  lines.push('━━━━━━━━━━━━━━━━━━━');
  lines.push('');

  var officeTotal = 0;
  var shown = 0;
  for (var p = 0; p < individuals.length; p++) {
    var person = individuals[p];
    officeTotal += person.units;
    if (person.units <= 0) continue;
    var medal = shown < 3 ? medals[shown] : '🏅';
    lines.push(medal + ' ' + person.name + ' — ' + person.units);
    shown++;
  }

  if (shown === 0) {
    lines.push('No sales this week yet.');
  }

  // Team rankings
  var teamTotals = {};
  for (var j = 0; j < individuals.length; j++) {
    var ind = individuals[j];
    var teamName = ind.team || 'Unassigned';
    if (teamName === 'Unassigned' || !teamName) continue;
    if (!teamTotals[teamName]) teamTotals[teamName] = 0;
    teamTotals[teamName] += ind.units;
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
function postLeaderboardToChat(optOfficeId) {
  var props = PropertiesService.getScriptProperties();
  var webhookUrl = (props.getProperty('CHAT_WEBHOOK_URL') || '').trim();
  var platform   = (props.getProperty('CHAT_PLATFORM') || 'discord').trim().toLowerCase();
  if (!webhookUrl || platform === 'none') {
    Logger.log('[LeaderboardPost] No webhook configured — skipping');
    return;
  }

  var ss = getSheet({});
  var officeId = optOfficeId || DEFAULT_OFFICE_ID;
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
