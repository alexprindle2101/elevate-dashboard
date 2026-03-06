// ═══════════════════════════════════════════════════════
// ELEVATE Dashboard — Google Apps Script Middleware
// ═══════════════════════════════════════════════════════
// Deploy as Web App: Execute as ME, Anyone can access.
// Set API_KEY in Script Properties (Project Settings > Script Properties).

// === CONFIG ===
const ORDER_LOG_TAB = 'Order Log';
const ROSTER_TAB = '_Roster';
const ORDER_OVERRIDES_TAB = '_OrderOverrides';
const TEAM_CUSTOM_TAB = '_TeamCustomizations';
const UNLOCK_REQ_TAB = '_UnlockRequests';
const TEAMS_TAB = '_Teams';
const SETTINGS_TAB = '_Settings';
const CHURN_REPORT_TAB = '_TableauChurnReport';

// Build team emoji maps dynamically from _Teams tab
function buildTeamEmojiMaps(ss) {
  var emojiMap = {};   // emoji → name
  var nameMap = {};    // name → emoji
  var sheet = ss.getSheetByName(TEAMS_TAB);
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

// Order Log column indices (0-based)
const OL = {
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


// _TableauOrderLog — header-based column lookup (resilient to column inserts/reorders)
const TABLEAU_TAB = '_TableauOrderLog';

// Map header text (lowercased, trimmed) → internal key
// Multiple aliases per key so column renames/variations still match
const TOL_HEADER_MAP = {
  // Col 0 — Owner & Office
  'owner & office': 'OWNER_OFFICE',
  'owner and office': 'OWNER_OFFICE',
  // Col 1 — Rep
  'rep': 'REP',
  // Col 2 — icd.Lead Rep ID
  'icd.lead rep id': 'LEAD_REP_ID',
  'lead rep id': 'LEAD_REP_ID',
  // Col 3 — rep.Rep Number
  'rep.rep number': 'REP_NUMBER',
  'rep number': 'REP_NUMBER',
  // Col 4 — sp.Order Date (copy)
  'sp.order date (copy)': 'ORDER_DATE',
  'order date (copy)': 'ORDER_DATE',
  'order date': 'ORDER_DATE',
  // Col 5 — Order Time (Timezone)
  'order time (timezone)': 'ORDER_TIME',
  'order time': 'ORDER_TIME',
  // Col 6 — sp.SPM Number  (THIS IS THE DSI COLUMN)
  'sp.spm number': 'DSI',
  'spm number': 'DSI',
  'dsi': 'DSI',
  // Col 7 — spe.Name  (THIS IS THE SPE COLUMN)
  'spe.name': 'SPE',
  'spe name': 'SPE',
  'spe': 'SPE',
  // Col 8 — spe.Account BAN
  'spe.account ban': 'BAN',
  'account ban': 'BAN',
  'ban': 'BAN',
  // Col 9 — Product Type (Broken Out)
  'product type (broken out)': 'PRODUCT_TYPE',
  'product type': 'PRODUCT_TYPE',
  // Col 10 — CRU/IRU
  'cru/iru': 'CRU_IRU',
  'cru / iru': 'CRU_IRU',
  // Col 11 — DTR Status (enriched)
  'dtr status (enriched)': 'DTR_STATUS',
  'dtr status': 'DTR_STATUS',
  // Col 12 — Disconnect Reason (Consolidated)
  'disconnect reason (consolidated)': 'DISCO_REASON',
  'disconnect reason': 'DISCO_REASON',
  'disco reason': 'DISCO_REASON',
  // Col 13 — spe.Port Carrier
  'spe.port carrier': 'PORT_CARRIER',
  'port carrier': 'PORT_CARRIER',
  // Col 14 — Notes.Note
  'notes.note': 'NOTES',
  'notes': 'NOTES',
  // Col 15 — DTR Status Date
  'dtr status date': 'DTR_STATUS_DATE',
  // Col 16 — Order Status
  'order status': 'ORDER_STATUS',
  // Col 17 — spe.dtr Posted Date (copy)
  'spe.dtr posted date (copy)': 'POSTED_DATE',
  'posted date (copy)': 'POSTED_DATE',
  'posted date': 'POSTED_DATE',
  // Col 18 — Max Posted
  'max posted': 'MAX_POSTED',
  // Col 19 — First Streaming Date
  'first streaming date': 'FIRST_STREAMING',
  'first streaming': 'FIRST_STREAMING',
  // Col 20 — Voice Line Count
  'voice line count': 'VOICE_LINE_COUNT',
  // Col 21 — spe.TN Type
  'spe.tn type': 'TN_TYPE',
  'tn type': 'TN_TYPE',
  // Col 22 — spe.Phone
  'spe.phone': 'PHONE',
  'phone': 'PHONE',
  // Col 23 — spe.Install Date
  'spe.install date': 'INSTALL_DATE',
  'install date': 'INSTALL_DATE',
  // Col 24 — B2B Rep Volume Bonus Tiers
  'b2b rep volume bonus tiers': 'BONUS_TIERS',
  'bonus tiers': 'BONUS_TIERS',
  // Col 25 — Tier Bonus Payout/DNQ Reason
  'tier bonus payout/dnq reason': 'PAYOUT_REASON',
  'payout reason': 'PAYOUT_REASON',
  // Col 26 — Unit Count
  'unit count': 'UNIT_COUNT',
  // Col 27 — Total Volume
  'total volume': 'TOTAL_VOLUME',
  // Col 28 — Total Activations
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

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    switch (name) {
      case ROSTER_TAB:
        sheet.appendRow(['email', 'name', 'team', 'rank', 'deactivated', 'dateAdded', 'pinHash', 'phone', 'tableauName']);
        break;
      case ORDER_OVERRIDES_TAB:
        sheet.appendRow(['key', 'product', 'status', 'date', 'order', 'notes_json']);
        break;
      case TEAM_CUSTOM_TAB:
        sheet.appendRow(['persona', 'emoji', 'displayName']);
        break;
      case UNLOCK_REQ_TAB:
        sheet.appendRow(['persona', 'status']);
        break;
      case TEAMS_TAB:
        sheet.appendRow(['teamId', 'name', 'parentId', 'leaderId', 'emoji', 'createdDate']);
        break;
      case SETTINGS_TAB:
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Orders-specific request
    if (action === 'readOrders') {
      const filterEmail = (e.parameter && e.parameter.email) || '';
      const orders = readOrders(ss, filterEmail || null);
      return jsonResponse({ orders: orders });
    }

    // Payroll orders — trainee=Yes, past 2 months
    if (action === 'readPayrollOrders') {
      const orders = readPayrollOrders(ss);
      return jsonResponse({ orders: orders });
    }

    // Tableau summary (on-demand refresh)
    if (action === 'readTableauSummary') {
      return jsonResponse(getTableauSummaryWithCache(ss));
    }

    // Tableau device detail for a single DSI
    if (action === 'readTableauDetail') {
      const dsi = (e.parameter && e.parameter.dsi) || '';
      return jsonResponse({ devices: readTableauDetail(ss, dsi) });
    }

    // Leaderboard snapshot (used by Make.com for daily Discord post)
    if (action === 'leaderboard') {
      return jsonResponse(readLeaderboard(ss));
    }

    // Leaderboard as pre-built HTML (for HTML-to-Image rendering)
    if (action === 'leaderboardHtml') {
      var lb = readLeaderboard(ss);
      return jsonResponse({ html: buildLeaderboardHtml(lb) });
    }

    // Default: return full dashboard data (includes Tableau summary)
    let roster = readRoster(ss);
    const teamMaps = buildTeamEmojiMaps(ss);
    const peopleResult = readPeople(ss, roster, teamMaps.nameMap);
    const tableauSummary = getTableauSummaryWithCache(ss);

    // Auto-assign Tableau names (managers first, then down the ranks)
    roster = autoAssignTableauNames(ss, roster, tableauSummary.possibleTableauNames);

    const data = {
      people: peopleResult.people || peopleResult,
      roster: roster,
      teamMap: teamMaps.emojiMap,
      teams: readTeams(ss),
      orderOverrides: readOrderOverrides(ss),
      teamCustomizations: readTeamCustomizations(ss),
      unlockRequests: readUnlockRequests(ss),
      settings: readSettings(ss),
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

function readRoster(ss) {
  const sheet = ss.getSheetByName(ROSTER_TAB);
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

function readPeople(ss, roster, teamNameToEmoji) {
  const olSheet = ss.getSheetByName(ORDER_LOG_TAB);
  if (!olSheet) return { people: [], _debug: { error: 'No Order Log sheet' } };
  if (!roster || Object.keys(roster).length === 0) return { people: [], _debug: { error: 'Empty roster' } };

  const olData = olSheet.getDataRange().getValues();
  if (olData.length < 2) return { people: [], _debug: { error: 'Order Log has no data rows' } };

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
      'col2(email)=' + String(olData[1][OL.EMAIL]),
      'col3(date)=' + String(olData[1][OL.DATE_OF_SALE]),
      'col4(name)=' + String(olData[1][OL.REP_NAME]),
      'col34(yeses)=' + String(olData[1][OL.YESES]),
      'col36(units)=' + String(olData[1][OL.UNITS])
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
  for (let i = 1; i < olData.length; i++) {
    const row = olData[i];
    const email = String(row[OL.EMAIL] || '').trim().toLowerCase();
    if (!email) { _dbg.noEmail++; continue; }
    if (!roster[email]) { _dbg.emailNotInRoster++; continue; }

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

function readOrders(ss, filterEmail) {
  const olSheet = ss.getSheetByName(ORDER_LOG_TAB);
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
      air:   Number(row[OL.AIR]) || 0,
      cell:  Number(row[OL.CELL]) || 0,
      fiber: Number(row[OL.FIBER]) || 0,
      voip:  Number(row[OL.VOIP_QTY]) || 0,
      units: Number(row[OL.UNITS]) || 0,
      status: String(row[OL.STATUS] || 'Pending').trim(),
      notes:  String(row[OL.NOTES] || '').trim(),
      tickets: (function() { try { return JSON.parse(row[OL.TICKETS] || '[]'); } catch(e) { return []; } })()
    });
  }

  orders.sort((a, b) => b.dateOfSale.localeCompare(a.dateOfSale));
  return orders;
}


// === readPayrollOrders() — Trainee orders (col G = Yes) past 2 months ===

function readPayrollOrders(ss) {
  const olSheet = ss.getSheetByName(ORDER_LOG_TAB);
  if (!olSheet) return [];

  const olData = olSheet.getDataRange().getValues();
  if (olData.length < 2) return [];

  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - 60);
  cutoff.setHours(0, 0, 0, 0);

  const orders = [];

  for (let i = 1; i < olData.length; i++) {
    const row = olData[i];
    const email = String(row[OL.EMAIL] || '').trim().toLowerCase();
    if (!email) continue;

    // Only include rows where "Did you have a trainee?" = Yes
    const trainee = String(row[OL.TRAINEE] || '').trim().toLowerCase();
    if (trainee !== 'yes') continue;

    const rawDate = row[OL.DATE_OF_SALE];
    if (!rawDate) continue;
    const saleDate = new Date(rawDate);
    if (isNaN(saleDate.getTime())) continue;
    saleDate.setHours(0, 0, 0, 0);
    if (saleDate < cutoff) continue;

    // Parse paid-out JSON (col AO) — defaults to empty object
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
      paidOut: paidOut
    });
  }

  // Enrich payroll orders with Tableau SPE data
  var tableauSummary = getTableauSummaryWithCache(ss);
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


// === readSettings() — Key-value settings from _Settings tab ===

function readSettings(ss) {
  const sheet = ss.getSheetByName(SETTINGS_TAB);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    if (key) result[key] = String(data[i][1] || '').trim();
  }
  return result;
}


// === readChurnReport() — Read _TableauChurnReport tab ===

function readChurnReport(ss) {
  const sheet = ss.getSheetByName(CHURN_REPORT_TAB);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  // Map headers, assigning 'metricType' to the blank column (index 4)
  const headers = data[0].map((h, j) => {
    const trimmed = String(h).trim();
    return trimmed === '' ? 'metricType' : trimmed;
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

function writeSetting(body) {
  const sheet = getOrCreateSheet(SETTINGS_TAB);
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


function readOrderOverrides(ss) {
  const sheet = ss.getSheetByName(ORDER_OVERRIDES_TAB);
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

function readTeamCustomizations(ss) {
  const sheet = ss.getSheetByName(TEAM_CUSTOM_TAB);
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

function readUnlockRequests(ss) {
  const sheet = ss.getSheetByName(UNLOCK_REQ_TAB);
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

function readTeams(ss) {
  const sheet = ss.getSheetByName(TEAMS_TAB);
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

// Build DSI → email map from Order Log (for joining Tableau DSIs to roster emails)
function buildDsiEmailMap(ss) {
  var olSheet = ss.getSheetByName(ORDER_LOG_TAB);
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

// Read and aggregate _TableauOrderLog by DSI and by rep email
function readTableauSummary(ss) {
  var sheet = ss.getSheetByName(TABLEAU_TAB);
  if (!sheet) return { dsiSummary: {}, repSummary: {} };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { dsiSummary: {}, repSummary: {} };

  // Build column map from header row
  var col = buildTableauColumnMap(data[0]);

  // Build DSI → email map + email → DSIs map for rep aggregation
  var maps = buildDsiEmailMap(ss);
  var dsiEmailMap = maps.dsiToEmail;
  var emailToDsis = maps.emailToDsis;

  // Month window cutoff for Active % calc
  var thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
  thirtyDaysAgo.setHours(0, 0, 0, 0);

  var dsiSummary = {};

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
        monthWirelessSPEs: {}  // spe → { orderStatus, dtrStatus }
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
      // Track WIRELESS SPEs within 30-day window for Active/Pending/Cancel/Disco %
      var ptUpper = productType.toUpperCase();
      var isWireless = ptUpper === 'WIRELESS';
      var isInMonth = orderDate instanceof Date && orderDate >= thirtyDaysAgo;
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
      repSummary[email] = {
        totalDevices: 0,
        totalActivations: 0,
        totalVolume: 0,
        statusCounts: {},
        productCounts: {},
        tableauName: ds.tableauRep,
        monthWirelessSPEs: {}
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
    // Merge month-window wireless SPEs across DSIs
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
      repByName[name] = {
        totalDevices: 0, totalActivations: 0, totalVolume: 0,
        statusCounts: {}, productCounts: {}, tableauName: name,
        monthWirelessSPEs: {}
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
function autoAssignTableauNames(ss, roster, possibleTableauNames) {
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
  var sheet = ss.getSheetByName(ROSTER_TAB);
  var newAssignments = [];

  unassigned.forEach(function(entry) {
    var possible = possibleTableauNames[entry.email] || [];
    var unclaimed = possible.filter(function(n) { return !claimed[n]; });

    if (unclaimed.length === 1) {
      var name = unclaimed[0];
      claimed[name] = true;  // mark it so lower-rank reps can't take it
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


// Lazy-load per-device detail rows for a single DSI
function readTableauDetail(ss, dsi) {
  var sheet = ss.getSheetByName(TABLEAU_TAB);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var targetDsi = String(dsi || '').trim();
  if (!targetDsi) return [];

  // Build column map from header row
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

// Cached wrapper for readTableauSummary (6-hour TTL)
function getTableauSummaryWithCache(ss) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'tableauSummary_v5';
  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* fall through */ }
  }

  var summary = readTableauSummary(ss);
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
function writeBustTableauCache() {
  try {
    CacheService.getScriptCache().remove('tableauSummary_v5');
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

  try {
    let result;
    switch (body.action) {
      // Roster management (JD+ only — enforced client-side)
      case 'addRosterEntry':      result = writeAddRosterEntry(body); break;
      case 'updateRosterEntry':   result = writeUpdateRosterEntry(body); break;
      case 'setTableauName':      result = writeSetTableauName(body); break;
      case 'deleteRosterEntry':   result = writeDeleteRosterEntry(body); break;
      case 'toggleDeactivate':    result = writeToggleDeactivate(body); break;
      // Existing actions
      case 'saveOrderOverride':    result = writeSaveOrderOverride(body); break;
      case 'setTeamCustomization': result = writeSetTeamCustomization(body); break;
      case 'setUnlockRequest':     result = writeSetUnlockRequest(body); break;
      case 'deleteUnlockRequest':  result = writeDeleteUnlockRequest(body); break;
      // Team hierarchy management
      case 'addTeam':              result = writeAddTeam(body); break;
      case 'updateTeam':           result = writeUpdateTeam(body); break;
      case 'deleteTeam':           result = writeDeleteTeam(body); break;
      // PIN authentication
      case 'setPin':               result = writeSetPin(body); break;
      case 'validatePin':          result = writeValidatePin(body); break;
      case 'changePin':            result = writeChangePin(body); break;
      // Order management
      case 'writeOrderNote':       result = writeOrderNote(body); break;
      case 'setOrderStatus':       result = writeSetOrderStatus(body); break;
      case 'updateOrder':          result = writeUpdateOrder(body); break;
      case 'setSetting':           result = writeSetting(body); break;
      case 'savePaidOut':          result = writeSavePaidOut(body); break;
      // Ticket management
      case 'addTicket':            result = writeAddTicket(body); break;
      case 'toggleTicket':         result = writeToggleTicket(body); break;
      // Sale submission
      case 'addSale':              result = writeAddSale(body); break;
      case 'bustTableauCache':     result = writeBustTableauCache(); break;
      default: result = { error: 'unknown action: ' + body.action };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}


// === ROSTER WRITE HELPERS ===

function writeAddRosterEntry(body) {
  const sheet = getOrCreateSheet(ROSTER_TAB);
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
    body.phone || '',  // phone
    ''   // tableauName — empty until rep picks from popup
  ]);
  return { ok: true };
}

function writeUpdateRosterEntry(body) {
  const sheet = getOrCreateSheet(ROSTER_TAB);
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
    body.phone !== undefined ? body.phone : (cur[7] || ''),  // phone
    body.tableauName !== undefined ? body.tableauName : (cur[8] || '')  // tableauName
  ];
  sheet.getRange(rowIdx, 1, 1, 9).setValues([rowData]);
  return { ok: true };
}

function writeDeleteRosterEntry(body) {
  const sheet = getOrCreateSheet(ROSTER_TAB);
  const email = String(body.email || '').trim().toLowerCase();
  if (!email) return { error: 'missing email' };

  const rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx > 0) sheet.deleteRow(rowIdx);
  return { ok: true };
}

function writeToggleDeactivate(body) {
  const sheet = getOrCreateSheet(ROSTER_TAB);
  const email = String(body.email || '').trim().toLowerCase();
  if (!email) return { error: 'missing email' };

  const rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 5).setValue(body.deactivated === true || body.deactivated === 'true');
  }
  return { ok: true };
}


function writeSetTableauName(body) {
  const sheet = getOrCreateSheet(ROSTER_TAB);
  const email = String(body.email || '').trim().toLowerCase();
  if (!email) return { error: 'missing email' };
  const tableauName = String(body.tableauName || '').trim();
  if (!tableauName) return { error: 'missing tableauName' };

  const rowIdx = findRowCI(sheet, 0, email);
  if (rowIdx < 0) return { error: 'email not found' };

  // Column 9 = tableauName (1-based index)
  sheet.getRange(rowIdx, 9).setValue(tableauName);
  return { ok: true };
}


// === EXISTING WRITE HELPERS ===

function writeSaveOrderOverride(body) {
  const sheet = getOrCreateSheet(ORDER_OVERRIDES_TAB);
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

function writeSetTeamCustomization(body) {
  const sheet = getOrCreateSheet(TEAM_CUSTOM_TAB);
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

function writeSetUnlockRequest(body) {
  const sheet = getOrCreateSheet(UNLOCK_REQ_TAB);
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

function writeDeleteUnlockRequest(body) {
  const sheet = getOrCreateSheet(UNLOCK_REQ_TAB);
  const persona = String(body.persona || '').trim();
  if (!persona) return { error: 'missing persona' };

  const rowIdx = findRow(sheet, 0, persona);
  if (rowIdx > 0) sheet.deleteRow(rowIdx);
  return { ok: true };
}


// === PIN AUTHENTICATION ===

function writeSetPin(body) {
  var sheet = getOrCreateSheet(ROSTER_TAB);
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

function writeValidatePin(body) {
  var sheet = getOrCreateSheet(ROSTER_TAB);
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

function writeAddTeam(body) {
  const sheet = getOrCreateSheet(TEAMS_TAB);
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

function writeUpdateTeam(body) {
  const sheet = getOrCreateSheet(TEAMS_TAB);
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

function writeDeleteTeam(body) {
  const sheet = getOrCreateSheet(TEAMS_TAB);
  const teamId = String(body.teamId || '').trim();
  if (!teamId) return { error: 'missing teamId' };

  const rowIdx = findRow(sheet, 0, teamId);
  if (rowIdx > 0) sheet.deleteRow(rowIdx);
  return { ok: true };
}

// One-time migration: seed _Teams sheet from existing team config
function migrateTeams() {
  var sheet = getOrCreateSheet(TEAMS_TAB);
  var existingTeams = ['Aces', 'Squids', 'Grind Team', 'Queenz', 'Different Breed', 'Sharks', 'Dawgs'];
  var today = new Date().toISOString().split('T')[0];

  existingTeams.forEach(function(name) {
    var teamId = name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
    var emoji = TEAM_NAME_TO_EMOJI[name] || '';
    // Skip if already exists
    if (findRow(sheet, 0, teamId) > 0) return;
    sheet.appendRow([teamId, name, '', '', emoji, today]);
  });
}


function writeChangePin(body) {
  var sheet = getOrCreateSheet(ROSTER_TAB);
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

function writeOrderNote(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ORDER_LOG_TAB);
  if (!sheet) return { error: 'Order Log not found' };

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

function writeAddTicket(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ORDER_LOG_TAB);
  if (!sheet) return { error: 'Order Log not found' };

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

function writeToggleTicket(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ORDER_LOG_TAB);
  if (!sheet) return { error: 'Order Log not found' };

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

function writeSetOrderStatus(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ORDER_LOG_TAB);
  if (!sheet) return { error: 'Order Log not found' };

  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'Invalid row' };

  var status = String(body.status || 'Pending').trim();
  sheet.getRange(rowIndex, OL.STATUS + 1).setValue(status);
  return { ok: true };
}

function writeUpdateOrder(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ORDER_LOG_TAB);
  if (!sheet) return { error: 'Order Log not found' };

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

function writeSavePaidOut(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ORDER_LOG_TAB);
  if (!sheet) return { error: 'Order Log not found' };

  var rowIndex = Number(body.rowIndex);
  if (!rowIndex || rowIndex < 2) return { error: 'Invalid row' };

  var paidOutJson = JSON.stringify(body.paidOut || {});
  sheet.getRange(rowIndex, OL.PAID_OUT + 1).setValue(paidOutJson);
  return { ok: true };
}


// === LEADERBOARD SNAPSHOT (for daily Discord post) ===

function readLeaderboard(ss) {
  const roster = readRoster(ss);
  const teams = readTeams(ss);

  // Read Order Log
  const olSheet = ss.getSheetByName(ORDER_LOG_TAB);
  if (!olSheet) return { error: 'No Order Log' };
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
    // Hero stats banner
    '<div style="display:flex;align-items:center;background:rgba(255,255,255,0.5);border:1px solid rgba(0,0,0,0.3);border-radius:14px;margin-bottom:32px;overflow:hidden;">' +
      '<div style="flex:1;display:flex;flex-direction:column;align-items:center;padding:24px 40px;gap:6px;">' +
        '<div style="font-family:Helvetica Neue,Inter,sans-serif;font-size:12px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#708090;">Week Yeses</div>' +
        '<div style="font-family:Inter,sans-serif;font-size:56px;font-weight:700;line-height:1;color:#242124;">' + wY + '</div></div>' +
      '<div style="width:1px;height:60px;background:rgba(0,0,0,0.3);"></div>' +
      '<div style="flex:1;display:flex;flex-direction:column;align-items:center;padding:24px 40px;gap:6px;">' +
        '<div style="font-family:Helvetica Neue,Inter,sans-serif;font-size:12px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#708090;">Week Units</div>' +
        '<div style="font-family:Inter,sans-serif;font-size:56px;font-weight:700;line-height:1;color:#43B3AE;">' + wU + '</div></div>' +
    '</div>' +
    // TOP PERFORMERS heading
    '<div style="display:flex;align-items:center;gap:16px;margin-bottom:20px;">' +
      '<div style="font-family:Helvetica Neue,Inter,sans-serif;font-size:16px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:#242124;">Top Performers</div>' +
      '<div style="flex:1;height:1px;background:rgba(0,0,0,0.15);"></div>' +
      '<div style="font-size:13px;font-style:italic;color:#708090;">This week</div></div>' +
    // INDIVIDUALS
    '<div style="font-family:Helvetica Neue,Inter,sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:#708090;margin-bottom:12px;">Individuals</div>' +
    podiumRow(indOrder, false) +
    // TEAMS
    '<div style="font-family:Helvetica Neue,Inter,sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:#708090;margin-bottom:12px;">Teams</div>' +
    podiumRow(teamOrder, true) +
    // Footer
    '<div style="text-align:center;padding-top:8px;border-top:1px solid rgba(0,0,0,0.08);">' +
      '<div style="font-size:11px;color:#708090;letter-spacing:0.5px;">End of Day Snapshot \u00B7 ' + dateStr + '</div></div>' +
  '</div>';
}


// === SALE SUBMISSION ===

function writeAddSale(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ORDER_LOG_TAB);
  if (!sheet) return { error: 'Order Log not found' };

  // Validate required fields
  var email = String(body.email || '').trim().toLowerCase();
  if (!email) return { error: 'Missing email' };
  var dateOfSale = body.dateOfSale;
  if (!dateOfSale) return { error: 'Missing date of sale' };

  // Look up team emoji from _Roster + _Teams
  var teamEmoji = '';
  try {
    var rosterSheet = ss.getSheetByName(ROSTER_TAB);
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
        var teamsSheet = ss.getSheetByName(TEAMS_TAB);
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

  // Build row array — write only to known OL column indices
  var totalCols = OL.TICKETS + 2;  // ensure array is long enough
  var newRow = [];
  for (var i = 0; i < totalCols; i++) newRow.push('');

  newRow[OL.TIMESTAMP]    = new Date();
  newRow[OL.EMAIL]        = email;
  newRow[OL.DATE_OF_SALE] = new Date(dateOfSale + 'T12:00:00');
  newRow[OL.REP_NAME]     = String(body.repName || '').trim();
  newRow[OL.DSI]          = String(body.dsi || '').trim();
  newRow[OL.TRAINEE]      = body.trainee ? 'Yes' : 'No';
  newRow[OL.TRAINEE_NAME] = body.trainee ? String(body.traineeName || '').trim() : '';
  newRow[OL.VOIP_QTY]     = voipQty;
  newRow[OL.TEAM_EMOJI]   = teamEmoji;
  newRow[OL.FIBER]        = fiber;
  newRow[OL.AIR]          = air;
  newRow[OL.DTV]          = dtv;
  newRow[OL.YESES]        = yeses;
  newRow[OL.CELL]         = cell;
  newRow[OL.UNITS]        = units;
  newRow[OL.STATUS]       = 'Pending';
  newRow[OL.NOTES]        = '';
  newRow[OL.PAID_OUT]     = '';
  newRow[OL.TICKETS]      = '[]';

  sheet.appendRow(newRow);

  // Bust the Tableau cache so fresh data loads
  try {
    CacheService.getScriptCache().remove('tableauSummary_v5');
  } catch (e) { /* non-critical */ }

  return {
    ok: true,
    rowIndex: sheet.getLastRow(),
    units: units,
    yeses: yeses
  };
}
