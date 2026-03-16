// ═══════════════════════════════════════════════════════
// NationalCode.gs — National Consultant Dashboard Data Aggregator
// Reads from 4 external Google Sheets, returns unified JSON.
// Deploy as web app (Execute as: Me, Access: Anyone)
// ═══════════════════════════════════════════════════════

// ── Config ──
var NC_API_KEY = 'national-dash-2026-secret'; // Must match Script Properties > API_KEY

// External sheet IDs
var SHEETS = {
  RECRUITING_WEEKLY:  '1MNLqi8A329444SeZpKbYbcRe3dMxaOPLVdMy-7F1DPk',  // All Campaigns Stats Tracker 2026
  RECRUITING_DAILY:   '1ytTGen_AlzfDPW3HGYU1JKNLz1kfHrrhAFCVnmRS3fg',  // Recruiting Scoreboard Daily
  CAMPAIGN_TRACKER:   '1HvWJYox3JXvxmza63YBWAqKPtUGPFuaV-s-BOfbWGKM',  // ATT Campaign Tracker
  PERFORMANCE_AUDIT:  '15WCMzKnqvyyRMx2ae4tC1a12_-aoSRDh3McOuRAKuHk',  // Performance Audit
  NATIONAL:           '1eGkwjQRD9RV4n-JR_TTlgE6VY858WZID8cAF8soYSYM', // Ken's national recruiting sheet
  NLR_B2B:            '1sxauFjNjq4_rRYM2PAl5cyOyHF3Hg4OkO-t_hLKDJB8', // NLR's AT&T B2B 1-on-1's report
  NDS_ONE_ON_ONES:    '1kcUWR3EKgP-9wDct4vDyuQJ7IuS0cbcetY97dmVTY64', // AT&T NDS/Verizon Wireless One on Ones
  NLR_NDS:            '1u2iM7gfEGLUtxog5nxOLpjwCmJhWVsF5aT7_l87SCCg', // NLR's AT&T NDS 1-on-1's report (Sam Poles)
  INDEED_COSTS_FOLDER: '1r2lGOOjXQkvzz1we5k1Gn5drXrZKc42y'              // Drive folder: per-owner Indeed ad spend sheets
};

// Campaign configs
var CAMPAIGNS = {
  'att-b2b': {
    label: 'AT&T B2B',
    sectionHeader: 'AT&T Campaign Totals',   // Header text in recruiting sheets
    campaignTotalsTab: 'Campaign Totals',     // Tab in Campaign Tracker
    owners: [
      { name: 'Jay T',          tab: 'Jay T' },
      { name: 'Mason',          tab: 'Mason' },
      { name: 'Steven Sykes',   tab: 'Steven Sykes' },
      { name: 'Olin Salter',    tab: 'Olin Salter' },
      { name: 'Eric Martinez',  tab: 'Eric Martinez' },
      { name: 'Natalia Gwarda', tab: 'Natalia Gwarda' },
      { name: 'Nigel Gilbert',  tab: 'Nigel Gil' }
    ]
  }
};

// ══════════════════════════════════════════════════
// ENTRY POINT
// ══════════════════════════════════════════════════

function doGet(e) {
  const key = (e && e.parameter && e.parameter.key) || '';
  if (!validateKey(key)) return jsonResp({ error: 'unauthorized' });

  const action = (e && e.parameter && e.parameter.action) || '';
  const campaign = (e && e.parameter && e.parameter.campaign) || 'att-b2b';
  const owner = (e && e.parameter && e.parameter.owner) || '';

  try {
    // ── National recruiting data from Ken's sheet ──
    if (action === 'recruiting') {
      var weeks = parseInt(e.parameter.weeks) || 6;
      return jsonResp(readNationalRecruiting(weeks));
    }

    // ── Owner → Cam Company mapping from _OwnerCamMapping tab ──
    if (action === 'ownerCamMapping') {
      return jsonResp(readOwnerCamMapping());
    }

    // ── Online presence / audit data from Cam's sheet ──
    if (action === 'onlinePresence') {
      return jsonResp(readOnlinePresence());
    }

    // ── B2B headcount/production data (local copy in _B2B_Headcount tab) ──
    if (action === 'b2bHeadcount') {
      return jsonResp(readLocalHeadcount());
    }

    // ── B2B production/sales data (local copy in _B2B_Production tab) ──
    if (action === 'b2bProduction') {
      return jsonResp(readLocalProduction());
    }

    // ── NDS headcount data (local copy in _NDS_Headcount tab) ──
    if (action === 'ndsHeadcount') {
      return jsonResp(readLocalNDSHeadcount());
    }

    // ── List cost files in Drive folder (names + IDs only) ──
    if (action === 'listCostFiles') {
      return jsonResp(listCostFiles());
    }

    // ── Read cost data from a single spreadsheet by ID ──
    if (action === 'readCostSheet') {
      var sheetId = (e && e.parameter && e.parameter.sheetId) || '';
      return jsonResp(readCostSheet(sheetId));
    }

    // ── Weekly Indeed Tracking (prototype — Jackie Leroy's sheet) ──
    if (action === 'indeedTracking') {
      var ownerName = (e && e.parameter && e.parameter.owner) || '';
      return jsonResp(readIndeedTracking(ownerName));
    }

    // ── [DEPRECATED] Bulk Indeed costs — use readCostSheet instead ──
    if (action === 'indeedCosts') {
      var ownerFilter = (e && e.parameter && e.parameter.owners) || '';
      return jsonResp(readIndeedCosts(ownerFilter));
    }

    if (owner) {
      return jsonResp(loadOwnerDetail(campaign, owner));
    } else {
      return jsonResp(loadCampaignOverview(campaign));
    }
  } catch (err) {
    Logger.log('doGet error: ' + err.message + '\n' + err.stack);
    return jsonResp({ error: err.message });
  }
}

// ══════════════════════════════════════════════════
// POST ENTRY POINT
// ══════════════════════════════════════════════════

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResp({ error: 'invalid JSON' });
  }

  if (!validateKey(body.key || '')) {
    return jsonResp({ error: 'unauthorized' });
  }

  try {
    var result;
    switch (body.action) {
      case 'importRecruiting':
        result = importLatestRecruiting(typeof body.weeks === 'number' ? body.weeks : 1);
        break;
      case 'claimCompany':
        result = claimCompany(body.ownerName, body.companyName);
        break;
      case 'unclaimCompany':
        result = unclaimCompany(body.ownerName, body.companyName);
        break;
      case 'claimCostSheet':
        result = claimCostSheet(body.ownerName, body.sheetId);
        break;
      case 'unclaimCostSheet':
        result = unclaimCostSheet(body.ownerName);
        break;
      case 'importNLRHeadcount':
        result = importNLRHeadcount();
        break;
      case 'importNDSHeadcount':
        result = importNDSHeadcount();
        break;
      case 'updateHeadcount':
        result = updateHeadcountRow(body.ownerName, body.date,
                   body.active, body.leaders, body.dist, body.training);
        break;
      case 'updateProduction':
        result = updateProductionRow(body.ownerName, body.date,
                   body.productionLW, body.productionGoals);
        break;
      default:
        result = { error: 'unknown action: ' + body.action };
    }
    return jsonResp(result);
  } catch (err) {
    Logger.log('doPost error: ' + err.message + '\n' + err.stack);
    return jsonResp({ error: err.message });
  }
}

// ══════════════════════════════════════════════════
// IMPORT LATEST RECRUITING
// Copies newest tab from All Campaigns Stats Tracker
// into Ken's national sheet, then returns fresh data
// ══════════════════════════════════════════════════

function importLatestRecruiting(weekCount) {
  // weekCount: number of weeks to import (0 = all date-named tabs)
  weekCount = typeof weekCount === 'number' ? weekCount : 1;

  // 1. Open source sheet (Maddy's All Campaigns Stats Tracker)
  var srcSS = SpreadsheetApp.openById(SHEETS.RECRUITING_WEEKLY);
  var allSheets = srcSS.getSheets();
  if (!allSheets.length) return { error: 'Source sheet has no tabs' };

  // 2. Sort all tabs by date (newest first), keep only date-named tabs
  var tabInfos = [];
  for (var i = 0; i < allSheets.length; i++) {
    var name = allSheets[i].getName();
    var d = _parseTabDate(name);
    if (d) tabInfos.push({ sheet: allSheets[i], name: name, date: d });
  }
  if (!tabInfos.length) return { error: 'No date-named tabs found in source sheet' };

  tabInfos.sort(function(a, b) { return b.date.getTime() - a.date.getTime(); });

  // 3. Slice to requested week count (0 = all)
  var tabsToImport = weekCount === 0 ? tabInfos : tabInfos.slice(0, weekCount);

  // 4. Open destination sheet
  var dstSS = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var importedTabNames = [];
  var totalRows = 0;
  var totalSections = 0;

  // 5. Loop through each tab and copy
  for (var t = 0; t < tabsToImport.length; t++) {
    var tab = tabsToImport[t];
    var srcData = tab.sheet.getDataRange().getValues();
    var tabName = tab.name;

    var sections = _findCampaignSections(srcData);
    if (!sections.length) continue; // skip tabs with no recognizable sections

    // Build output rows — copy raw header + owner rows verbatim
    var outputRows = [];
    for (var s = 0; s < sections.length; s++) {
      var sec = sections[s];

      var headerRow = [];
      for (var c = 0; c < srcData[sec.headerRow].length; c++) {
        headerRow.push(srcData[sec.headerRow][c]);
      }
      outputRows.push(headerRow);

      for (var r = sec.startRow; r <= sec.endRow; r++) {
        var ownerRow = [];
        for (var c = 0; c < srcData[r].length; c++) {
          ownerRow.push(srcData[r][c]);
        }
        outputRows.push(ownerRow);
      }
      outputRows.push([]);
    }

    // Write to destination tab (create or overwrite)
    var dstSheet = dstSS.getSheetByName(tabName);
    if (dstSheet) {
      dstSheet.clearContents();
    } else {
      dstSheet = dstSS.insertSheet(tabName);
    }

    if (outputRows.length > 0) {
      var maxCols = 1;
      for (var i = 0; i < outputRows.length; i++) {
        if (outputRows[i].length > maxCols) maxCols = outputRows[i].length;
      }
      for (var i = 0; i < outputRows.length; i++) {
        while (outputRows[i].length < maxCols) outputRows[i].push('');
      }
      dstSheet.getRange(1, 1, outputRows.length, maxCols).setValues(outputRows);
    }

    importedTabNames.push(tabName);
    totalRows += outputRows.length;
    totalSections += sections.length;
  }

  if (!importedTabNames.length) return { error: 'No importable data found in source tabs' };

  // 6. Import applies/STL from NLR B2B sheet
  var appliesResult = {};
  try {
    appliesResult = importNLRApplies();
    Logger.log('NLR applies import: ' + JSON.stringify(appliesResult));
  } catch (err) {
    Logger.log('NLR applies import failed (non-fatal): ' + err.message);
  }

  // 6b. Import production data from NLR B2B sheet
  try {
    var prodResult = importNLRProduction();
    Logger.log('NLR production import: ' + JSON.stringify(prodResult));
  } catch (err) {
    Logger.log('NLR production import failed (non-fatal): ' + err.message);
  }

  // 7. Return success + fresh data
  var freshData = readNationalRecruiting(6);

  return {
    ok: true,
    imported: {
      tabCount: importedTabNames.length,
      tabNames: importedTabNames,
      tabName: importedTabNames[0], // backwards compat
      sections: totalSections,
      rows: totalRows
    },
    recruiting: freshData
  };
}

function validateKey(key) {
  // Check against script property first, fallback to hardcoded
  const stored = PropertiesService.getScriptProperties().getProperty('API_KEY');
  return key === (stored || NC_API_KEY);
}

function jsonResp(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════
// LOAD CAMPAIGN OVERVIEW
// Returns all owners with summary data
// ══════════════════════════════════════════════════

function loadCampaignOverview(campaignKey) {
  const cfg = CAMPAIGNS[campaignKey];
  if (!cfg) return { error: 'Unknown campaign: ' + campaignKey };

  const ctSS = SpreadsheetApp.openById(SHEETS.CAMPAIGN_TRACKER);
  const rwSS = SpreadsheetApp.openById(SHEETS.RECRUITING_WEEKLY);

  // 1. Read Campaign Totals tab for aggregate recruiting funnel
  const totals = readCampaignTotals(ctSS, cfg);

  // 2. Read each owner's tab in Campaign Tracker
  const owners = cfg.owners.map(function(ownerDef) {
    try {
      return loadOwnerFromCampaignTracker(ctSS, ownerDef, cfg);
    } catch (err) {
      Logger.log('Error loading owner ' + ownerDef.name + ': ' + err.message);
      return { name: ownerDef.name, tab: ownerDef.tab, error: err.message };
    }
  });

  // 3. Enrich with latest weekly recruiting data
  try {
    enrichWithWeeklyRecruiting(rwSS, owners, cfg);
  } catch (err) {
    Logger.log('Weekly recruiting enrichment error: ' + err.message);
  }

  return {
    campaign: campaignKey,
    label: cfg.label,
    totals: totals,
    owners: owners
  };
}

// ══════════════════════════════════════════════════
// CAMPAIGN TOTALS
// ══════════════════════════════════════════════════

function readCampaignTotals(ss, cfg) {
  var sheet = ss.getSheetByName(cfg.campaignTotalsTab);
  if (!sheet) return {};

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  // Header-based column lookup
  var headers = data[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var cols = findColumns(headers, {
    name: ['at&t campaign totals', 'name'],
    firstBooked: ['1st rounds booked'],
    firstShowed: ['1st rounds showed'],
    turnedTo2nd: ['turned to 2nd'],
    retention1: findNthRetention(headers, 1),
    conversion: ['conversion'],
    secondBooked: ['2nd rounds booked'],
    secondShowed: ['2nd rounds showed'],
    retention2: findNthRetention(headers, 2),
    newStarts: ['new start scheduled', 'new starts scheduled'],
    newStartsShowed: ['new starts showed'],
    retention3: findNthRetention(headers, 3),
    headcount: ['active selling headcount', 'active headcount']
  });

  // Sum all rep rows for campaign totals
  var totals = {
    headcount: 0, firstBooked: 0, newStarts: 0, retention: '—', production: '—'
  };
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = String(row[0] || '').trim();
    if (!name) continue;
    totals.headcount += num(row[cols.headcount]);
    totals.firstBooked += num(row[cols.firstBooked]);
    totals.newStarts += num(row[cols.newStarts]);
  }

  return totals;
}

// ══════════════════════════════════════════════════
// OWNER DATA FROM CAMPAIGN TRACKER
// Reads all 3 sections of an owner's tab
// ══════════════════════════════════════════════════

function loadOwnerFromCampaignTracker(ss, ownerDef, cfg) {
  var sheet = ss.getSheetByName(ownerDef.tab);
  if (!sheet) {
    return {
      name: ownerDef.name, tab: ownerDef.tab,
      health: { current: {}, trend: [] },
      recruiting: { funnel: {}, weekly: [], reps: [] },
      sales: { summary: {}, reps: [] },
      audit: { grades: {}, details: {} }
    };
  }

  var allData = sheet.getDataRange().getValues();

  // Find section boundaries by scanning for known header patterns
  var sections = findSections(allData);

  // SECTION 1: Office Health (Ken's manual input)
  var health = parseSection1(allData, sections.section1Start, sections.section1End);

  // SECTION 2: Recruiting by week
  var recruiting = parseSection2(allData, sections.section2Start, sections.section2End);

  // SECTION 3: Tableau sales data
  var sales = parseSection3(allData, sections.section3Start, sections.section3End);

  return {
    name: ownerDef.name,
    tab: ownerDef.tab,
    health: health,
    recruiting: { funnel: recruiting.totals || {}, weekly: recruiting.weekly || [], reps: [] },
    sales: sales,
    audit: { grades: { reviews: '—', website: '—', social: '—', seo: '—' }, details: {} }
  };
}

// ══════════════════════════════════════════════════
// SECTION DETECTION
// Scans the data for known header patterns to find
// where each of the 3 sections starts and ends.
// ══════════════════════════════════════════════════

function findSections(data) {
  var s1Start = 0, s1End = -1;
  var s2Start = -1, s2End = -1;
  var s3Start = -1, s3End = data.length - 1;

  for (var i = 0; i < data.length; i++) {
    var rowText = data[i].map(function(c) { return String(c).toLowerCase().trim(); }).join('|');

    // Section 1 starts at row 0 (Dates, Active, Leaders...)
    // Section 2 starts when we see "projected weekly" or "actual recruiting"
    if (rowText.indexOf('projected weekly') >= 0 || rowText.indexOf('actual recruiting') >= 0) {
      if (s1End < 0) s1End = i - 1;
      s2Start = i;
    }

    // Section 3 starts when we see Tableau-like headers
    if ((rowText.indexOf('new internet count') >= 0 || rowText.indexOf('sales (all)') >= 0) && s3Start < 0) {
      if (s2Start >= 0 && s2End < 0) s2End = i - 1;
      s3Start = i;
    }
  }

  // Fallbacks
  if (s1End < 0) s1End = (s2Start > 0 ? s2Start - 1 : data.length - 1);
  if (s2Start >= 0 && s2End < 0) s2End = (s3Start > 0 ? s3Start - 1 : data.length - 1);
  if (s3Start < 0) s3Start = data.length; // no section 3

  return {
    section1Start: s1Start, section1End: s1End,
    section2Start: s2Start, section2End: s2End,
    section3Start: s3Start, section3End: s3End
  };
}

// ══════════════════════════════════════════════════
// PARSE SECTION 1: Office Health
// ══════════════════════════════════════════════════

function parseSection1(data, start, end) {
  if (start < 0 || end < start) return { current: {}, trend: [] };

  // Find headers in row 0
  var headers = data[start].map(function(h) { return String(h).toLowerCase().trim(); });
  var colMap = {
    dates: findCol(headers, ['dates', 'date']),
    active: findCol(headers, ['active']),
    leaders: findCol(headers, ['leaders']),
    dist: findCol(headers, ['dist']),
    training: findCol(headers, ['training']),
    productionLW: findCol(headers, ['production lw', 'production']),
    dtv: findCol(headers, ['dtv']),
    goals: findCol(headers, ['production goals', 'goals'])
  };

  var trend = [];
  var lastGoodRow = null;

  for (var i = start + 1; i <= end; i++) {
    var row = data[i];
    var dateVal = row[colMap.dates];
    if (!dateVal) continue;

    var entry = {
      date: formatDate(dateVal),
      active: val(row[colMap.active]),
      leaders: val(row[colMap.leaders]),
      dist: val(row[colMap.dist]),
      training: val(row[colMap.training]),
      productionLW: val(row[colMap.productionLW]),
      dtv: val(row[colMap.dtv]),
      goals: val(row[colMap.goals])
    };
    trend.push(entry);
    lastGoodRow = entry;
  }

  return {
    current: lastGoodRow || { active: '—', leaders: '—', dist: '—', training: '—', productionLW: '—', dtv: '—', goals: '—' },
    trend: trend
  };
}

// ══════════════════════════════════════════════════
// PARSE SECTION 2: Recruiting by Week
// ══════════════════════════════════════════════════

function parseSection2(data, start, end) {
  if (start < 0 || end < start) return { totals: {}, weekly: [] };

  // Find the header row with week dates (e.g., "Feb-2", "Feb-9", etc.)
  var headerRowIdx = -1;
  for (var i = start; i <= Math.min(start + 3, end); i++) {
    var row = data[i];
    for (var j = 0; j < row.length; j++) {
      var cellText = String(row[j]).toLowerCase();
      if (cellText.indexOf('actual') >= 0 || cellText.indexOf('projected') >= 0) {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx >= 0) break;
  }

  // Look for the metrics rows: "Calls Received", "No List", "Booked", "Showed", etc.
  var weekly = [];
  var totals = {};
  var metricsMap = {};

  for (var i = (headerRowIdx >= 0 ? headerRowIdx + 1 : start); i <= end; i++) {
    var row = data[i];
    var label = String(row[0] || '').trim().toLowerCase();
    if (!label) continue;

    // Map known metric labels
    if (label.indexOf('calls received') >= 0 || label === 'calls received') metricsMap.callsReceived = row;
    else if (label.indexOf('no list') >= 0) metricsMap.noList = row;
    else if (label === 'booked' && !metricsMap.booked) metricsMap.booked = row;
    else if (label === 'showed' && !metricsMap.showed) metricsMap.showed = row;
    else if (label.indexOf('starts booked') >= 0) metricsMap.startsBooked = row;
    else if (label.indexOf('starts showed') >= 0) metricsMap.startsShowed = row;
    else if (label.indexOf('start retention') >= 0) metricsMap.startRetention = row;
    else if (label === 'total' || label.indexOf('month overview') >= 0) metricsMap.total = row;
  }

  return { totals: totals, weekly: [], rawMetrics: metricsMap };
}

// ══════════════════════════════════════════════════
// PARSE SECTION 3: Sales/Tableau Data
// ══════════════════════════════════════════════════

function parseSection3(data, start, end) {
  if (start < 0 || start >= data.length) return { summary: {}, reps: [] };

  // Find header row
  var headers = data[start].map(function(h) { return String(h).toLowerCase().trim(); });
  var colMap = {
    name: findCol(headers, ['name', '']),
    newInternet: findCol(headers, ['new internet count']),
    upgrade: findCol(headers, ['upgrade internet count', 'upgrade internet']),
    video: findCol(headers, ['video sales']),
    salesAll: findCol(headers, ['sales (all)', 'sales all']),
    abpMix: findCol(headers, ['new internet abp mix', 'abp mix']),
    gigMix: findCol(headers, ['new internet 1gig', '1gig+ mix']),
    techInstall: findCol(headers, ['tech install'])
  };

  // If col 0 name wasn't found by header, default to 0
  if (colMap.name < 0) colMap.name = 0;

  var reps = [];
  var sumSales = 0, sumInternet = 0, sumUpgrade = 0, sumVideo = 0;

  for (var i = start + 1; i <= end; i++) {
    var row = data[i];
    var name = String(row[colMap.name] || '').trim();
    if (!name) continue;

    var rep = {
      name: name,
      newInternet: num(row[colMap.newInternet]),
      upgrade: num(row[colMap.upgrade]),
      video: num(row[colMap.video]),
      salesAll: num(row[colMap.salesAll]),
      abpMix: pct(row[colMap.abpMix]),
      gigMix: pct(row[colMap.gigMix]),
      techInstall: pct(row[colMap.techInstall])
    };
    reps.push(rep);
    sumSales += rep.salesAll;
    sumInternet += rep.newInternet;
    sumUpgrade += rep.upgrade;
    sumVideo += rep.video;
  }

  return {
    summary: {
      totalSales: sumSales,
      newInternet: sumInternet,
      upgrades: sumUpgrade,
      videoSales: sumVideo,
      abpMix: '—',
      gigMix: '—'
    },
    reps: reps
  };
}

// ══════════════════════════════════════════════════
// WEEKLY RECRUITING ENRICHMENT
// Reads the latest tab from All Campaigns Stats Tracker
// and enriches owner data with per-rep recruiting numbers
// ══════════════════════════════════════════════════

function enrichWithWeeklyRecruiting(ss, owners, cfg) {
  // Find the most recent tab (tabs are named by date like "2-28-2026" or "03-02-26")
  var allSheets = ss.getSheets();
  if (!allSheets.length) return;

  // Use the first tab (most recent, per user's description)
  var latestSheet = allSheets[0];
  var data = latestSheet.getDataRange().getValues();

  // Find the AT&T section by looking for the section header
  var sectionStart = -1;
  for (var i = 0; i < data.length; i++) {
    var cellText = String(data[i][0] || '').toLowerCase().trim();
    if (cellText.indexOf('at&t campaign') >= 0 || cellText.indexOf('att campaign') >= 0) {
      sectionStart = i;
      break;
    }
  }
  if (sectionStart < 0) return;

  // Read headers from the section header row
  var headers = data[sectionStart].map(function(h) { return String(h).toLowerCase().trim(); });
  var colMap = {
    name: 0,
    firstBooked: findCol(headers, ['1st rounds booked']),
    firstShowed: findCol(headers, ['1st rounds showed']),
    turned2nd: findCol(headers, ['turned to 2nd']),
    conversion: findCol(headers, ['conversion']),
    secondBooked: findCol(headers, ['2nd rounds booked']),
    newStarts: findCol(headers, ['new start scheduled', 'new starts scheduled'])
  };

  // Build a map of rep name → recruiting data
  var repMap = {};
  for (var i = sectionStart + 1; i < data.length; i++) {
    var row = data[i];
    var name = String(row[0] || '').trim();
    if (!name) break; // end of section

    repMap[name.toLowerCase()] = {
      name: name,
      firstBooked: num(row[colMap.firstBooked]),
      firstShowed: num(row[colMap.firstShowed]),
      turned2nd: num(row[colMap.turned2nd]),
      conversion: pct(row[colMap.conversion]),
      secondBooked: num(row[colMap.secondBooked]),
      newStarts: num(row[colMap.newStarts])
    };
  }

  // Match reps to owners via the sales.reps list (rep names appear in both)
  // For now, attach all reps to the campaign data — we'll match by owner tab later
  // TODO: Cross-reference rep names between Campaign Tracker owner tabs and recruiting data
}

// ══════════════════════════════════════════════════
// NATIONAL RECRUITING — Read Ken's sheet
// Dynamically discovers campaigns & owners from
// Column A, reads recruiting actuals per week tab
// ══════════════════════════════════════════════════

// Known header patterns that identify a campaign section header row
var RECRUITING_HEADER_PATTERNS = [
  '1st rounds booked', '1st round booked', '1st rds booked',
  '1st rounds showed', '1st round showed', '1st rds showed',
  '2nd rounds booked', '2nd round booked', '2nd rds booked', '2nds booked',
  'retention', 'conversion', 'turned to 2nd',
  'new start scheduled', 'new starts scheduled', 'new start',
  'ns scheduled', 'ns showed'
];

function readNationalRecruiting(weekCount) {
  if (!SHEETS.NATIONAL) return { error: 'National sheet ID not configured' };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message };
  }

  var allTabs = ss.getSheets();
  if (!allTabs.length) return { error: 'National sheet has no tabs' };

  // Sort tabs: try to parse as dates (newest first), fall back to position
  var tabInfos = allTabs.map(function(sheet, idx) {
    var name = sheet.getName();
    var d = _parseTabDate(name);
    return { sheet: sheet, name: name, date: d, idx: idx };
  });
  tabInfos.sort(function(a, b) {
    if (a.date && b.date) return b.date.getTime() - a.date.getTime();
    if (a.date) return -1;
    if (b.date) return 1;
    return a.idx - b.idx;
  });

  // Take first N tabs
  var tabs = tabInfos.slice(0, weekCount);

  // Result: campaigns → { label, owners[], weeks[] }
  var campaigns = {};

  for (var t = 0; t < tabs.length; t++) {
    var sheet = tabs[t].sheet;
    var tabName = tabs[t].name;
    var data = sheet.getDataRange().getValues();

    // Find all campaign section headers in this tab
    var sections = _findCampaignSections(data);

    for (var s = 0; s < sections.length; s++) {
      var sec = sections[s];
      var key = _campaignSlug(sec.label);

      // Initialize campaign if first time seeing it
      if (!campaigns[key]) {
        var displayName = CAMPAIGN_DISPLAY_NAMES[key] || sec.label;
        campaigns[key] = { label: displayName, owners: [], weeks: [] };
      }

      // Parse owner rows for this section
      var weekData = _parseOwnerRecruiting(data, sec, key, tabName);

      // Collect unique owner names (from most recent tab)
      if (t === 0) {
        var ownerNames = Object.keys(weekData);
        for (var oi = 0; oi < ownerNames.length; oi++) {
          if (campaigns[key].owners.indexOf(ownerNames[oi]) < 0) {
            campaigns[key].owners.push(ownerNames[oi]);
          }
        }
      }

      // Add week data
      campaigns[key].weeks.push({
        tabName: tabName,
        data: weekData
      });
    }
  }

  // Merge applies/STL from _B2B_Applies tab into positions 0 and 1
  try {
    var appliesData = readLocalApplies();
    var appliesOwners = appliesData.owners || {};
    if (Object.keys(appliesOwners).length) {
      _mergeAppliesIntoCampaigns(campaigns, appliesOwners);
    }
  } catch (err) {
    Logger.log('Applies merge failed (non-fatal): ' + err.message);
  }

  return { campaigns: campaigns };
}

// ── Merge applies/STL into campaign week data ──
// appliesOwners: { "OwnerName": { "MM/DD/YYYY": { applies, stl } } }
// For each campaign week (tabName like "3/3"), find matching date in applies data.
function _mergeAppliesIntoCampaigns(campaigns, appliesOwners) {
  var keys = Object.keys(campaigns);
  for (var k = 0; k < keys.length; k++) {
    var camp = campaigns[keys[k]];
    for (var w = 0; w < camp.weeks.length; w++) {
      var week = camp.weeks[w];
      var tabDate = _parseTabDate(week.tabName);
      if (!tabDate) continue;

      // NLR dates are 1 week behind recruiting dates, so shift back 7 days to match
      var shiftedDate = new Date(tabDate.getTime() - 7 * 86400000);
      var normalizedShifted = _normalizeDate(shiftedDate);

      var ownerNames = Object.keys(week.data);
      for (var o = 0; o < ownerNames.length; o++) {
        var ownerName = ownerNames[o];
        var ownerApplies = appliesOwners[ownerName];
        if (!ownerApplies) continue;

        // Try exact match with shifted date (recruiting date - 7 days = NLR date)
        var match = ownerApplies[normalizedShifted];

        // If no exact match, try ±2 days around the shifted date
        if (!match) {
          var shiftedMs = shiftedDate.getTime();
          var bestKey = null;
          var bestDiff = Infinity;
          var aKeys = Object.keys(ownerApplies);
          for (var ai = 0; ai < aKeys.length; ai++) {
            var aDate = _parseTabDate(aKeys[ai]);
            if (!aDate) continue;
            var diff = Math.abs(aDate.getTime() - shiftedMs);
            if (diff < bestDiff && diff <= 2 * 86400000) { // ±2 days
              bestDiff = diff;
              bestKey = aKeys[ai];
            }
          }
          if (bestKey) match = ownerApplies[bestKey];
        }

        if (match) {
          week.data[ownerName][0] = match.applies || 0;
          week.data[ownerName][1] = match.stl || 0;
        }
      }
    }
  }
}

// ── Detect campaign section headers ──
// A section header has text in Column A AND has ≥2 recognizable
// recruiting metric headers in other columns of the same row.
function _findCampaignSections(data) {
  var sections = [];
  for (var i = 0; i < data.length; i++) {
    var colA = String(data[i][0] || '').trim();
    if (!colA) continue;

    // Count how many recognizable recruiting headers are in this row
    var headerCount = 0;
    var rowLower = data[i].map(function(c) { return String(c).toLowerCase().trim(); });
    for (var p = 0; p < RECRUITING_HEADER_PATTERNS.length; p++) {
      for (var c = 1; c < rowLower.length; c++) {
        if (rowLower[c].indexOf(RECRUITING_HEADER_PATTERNS[p]) >= 0) {
          headerCount++;
          break; // count each pattern once
        }
      }
      if (headerCount >= 2) break; // enough to confirm
    }

    if (headerCount >= 2) {
      // Find where this section ends (next section header or blank row or end of data)
      var endRow = data.length - 1;
      for (var j = i + 1; j < data.length; j++) {
        var nextColA = String(data[j][0] || '').trim();
        if (!nextColA) { endRow = j - 1; break; } // blank row = end

        // Check if next row is also a section header
        var nextHeaderCount = 0;
        var nextRowLower = data[j].map(function(c2) { return String(c2).toLowerCase().trim(); });
        for (var p2 = 0; p2 < RECRUITING_HEADER_PATTERNS.length; p2++) {
          for (var c2 = 1; c2 < nextRowLower.length; c2++) {
            if (nextRowLower[c2].indexOf(RECRUITING_HEADER_PATTERNS[p2]) >= 0) {
              nextHeaderCount++;
              break;
            }
          }
          if (nextHeaderCount >= 2) break;
        }
        if (nextHeaderCount >= 2) { endRow = j - 1; break; }
      }

      sections.push({
        label: colA,
        headerRow: i,
        startRow: i + 1,
        endRow: endRow,
        headers: rowLower
      });
    }
  }
  return sections;
}

// ── Parse owner rows within a campaign section ──
// Returns { "Owner Name": [12 values matching RECRUITING_LABELS order] }
function _parseOwnerRecruiting(data, section, campaignKey, tabName) {
  var headers = section.headers;

  // Map columns to the 12 RECRUITING_LABELS positions
  // Uses Nth-pattern matching for columns that repeat (retention) or vary in abbreviation
  // Fallbacks handle abbreviated headers across different tabs/campaigns:
  //   "new start scheduled" → "ns scheduled" | "2nd rounds booked" → "2nds booked"
  var ns1 = _findNthPattern(headers, 'new start', 1);
  var ns2 = _findNthPattern(headers, 'new start', 2);
  if (ns1 < 0) ns1 = _findNthPattern(headers, 'ns ', 1);
  if (ns2 < 0) ns2 = _findNthPattern(headers, 'ns ', 2);

  var r2_1 = _findNthPattern(headers, '2nd r', 1);
  var r2_2 = _findNthPattern(headers, '2nd r', 2);
  if (r2_1 < 0) r2_1 = _findNthPattern(headers, '2nds ', 1);
  if (r2_2 < 0) r2_2 = _findNthPattern(headers, '2nds ', 2);

  var convCol = findCol(headers, ['conversion', '% call list booked', 'turned to 2nd']);
  var rete2 = _findNthPattern(headers, 'rete', 2);
  var rete3 = _findNthPattern(headers, 'rete', 3);

  // Positional fallback: Maddy sometimes omits headers but data is still in the columns.
  // Standard layout: ... | Conversion | 2nd Booked | 2nd Showed | 2nd Retention | NS Booked | NS Showed | NS Retention | ...
  if (convCol >= 0 && r2_1 < 0) {
    r2_1 = convCol + 1;
  }
  if (convCol >= 0 && r2_2 < 0) {
    r2_2 = convCol + 2;
  }
  if (convCol >= 0 && rete2 < 0) {
    rete2 = convCol + 3;
  }
  if (rete2 >= 0 && ns1 < 0) {
    ns1 = rete2 + 1;
  }
  if (rete2 >= 0 && ns2 < 0) {
    ns2 = rete2 + 2;
  }
  if (rete2 >= 0 && rete3 < 0) {
    rete3 = rete2 + 3;
  }

  var colMap = [
    -1,                                                       // 0: Applies Received (not in sheet → 0)
    -1,                                                       // 1: Sent to List (not in sheet → 0)
    _findNthPattern(headers, '1st r', 1),                     // 2: 1st Rounds Booked
    _findNthPattern(headers, '1st r', 2),                     // 3: 1st Rounds Showed
    _findNthPattern(headers, 'rete', 1),                      // 4: 1st Retention
    convCol,                                                  // 5: Conversion
    r2_1,                                                     // 6: 2nd Rounds Booked
    r2_2,                                                     // 7: 2nd Rounds Showed
    rete2,                                                    // 8: 2nd Retention
    ns1,                                                      // 9: New Starts Booked/Scheduled
    ns2,                                                      // 10: New Starts Showed
    rete3                                                     // 11: New Start Retention
  ];

  var result = {};

  for (var i = section.startRow; i <= section.endRow; i++) {
    var row = data[i];
    var ownerName = String(row[0] || '').trim();
    if (!ownerName) continue;

    // Skip the last row if it's the sum row (always the final row in the section)
    if (i === section.endRow) continue;

    var values = [];
    for (var m = 0; m < 12; m++) {
      var ci = colMap[m];
      if (ci < 0) {
        values.push(0);
      } else {
        var cellVal = row[ci];
        // Rate rows (4, 5, 8, 11): read as percentage number
        if (m === 4 || m === 5 || m === 8 || m === 11) {
          values.push(_pctNum(cellVal));
        } else {
          values.push(num(cellVal));
        }
      }
    }
    result[ownerName] = values;
  }

  return result;
}

// ── Find Nth occurrence of a pattern in headers ──
function _findNthPattern(headers, pattern, n) {
  var count = 0;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].indexOf(pattern) >= 0) {
      count++;
      if (count === n) return i;
    }
  }
  return -1;
}

// ── Convert percentage cell to a number (e.g., 0.65→65, "65%"→65, 65→65) ──
function _pctNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  var s = String(v);
  if (s.indexOf('%') >= 0) {
    var n = parseFloat(s.replace('%', ''));
    return isNaN(n) ? 0 : Math.round(n);
  }
  var n = Number(v);
  if (isNaN(n)) return 0;
  // If decimal (0.65), convert to percentage
  if (n > 0 && n <= 1) return Math.round(n * 100);
  return Math.round(n);
}

// ── Parse tab name as date ──
function _parseTabDate(name) {
  // Try formats: "3-7-2026", "03-02-26", "Mar-7", "2-28-2026"
  var parts = name.split(/[-\/]/);
  if (parts.length >= 2) {
    var month = parseInt(parts[0]);
    var day = parseInt(parts[1]);
    var year = parts.length >= 3 ? parseInt(parts[2]) : new Date().getFullYear();
    if (year < 100) year += 2000;
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return new Date(year, month - 1, day);
    }
  }
  return null;
}

// ── Normalize campaign header text to a slug ──
// Known section headers from the sheet → slug + display name
var CAMPAIGN_DISPLAY_NAMES = {
  'att-b2b':          'AT&T: B2B',
  'att-nds':          'AT&T: NDS',
  'att-res':          'AT&T: Residential',
  'frontier':         'Frontier',
  'frontier-retail':  'Frontier: Retail',
  'leafguard':        'LeafGuard',
  'rogers':           'Rogers',
  'lumen':            'Lumen',
  'truconnect':       'TruConnect',
  'verizon':          'Verizon'
};

function _campaignSlug(label) {
  var lower = label.toLowerCase();
  // Known mappings — check most specific first
  var hasATT = lower.indexOf('at&t') >= 0 || lower.indexOf('att') >= 0 || lower.indexOf('at-t') >= 0;
  var hasB2B = lower.indexOf('b2b') >= 0;
  var hasOOF = lower.indexOf('oof') >= 0 || lower.indexOf('nds') >= 0;
  if (hasATT && hasB2B) return 'att-b2b';
  if (hasATT && hasOOF) return 'att-nds';
  if (hasATT) return 'att-res';
  if (lower.indexOf('frontier') >= 0 && lower.indexOf('retail') >= 0) return 'frontier-retail';
  if (lower.indexOf('frontier') >= 0) return 'frontier';
  if (lower.indexOf('rogers') >= 0) return 'rogers';
  if (lower.indexOf('leafguard') >= 0 || lower.indexOf('leaf guard') >= 0) return 'leafguard';
  if (lower.indexOf('truconnect') >= 0) return 'truconnect';
  if (lower.indexOf('lumen') >= 0) return 'lumen';
  if (lower.indexOf('verizon') >= 0) return 'verizon';
  // Fallback: slugify
  var s = lower
    .replace(/campaign\s*totals?/gi, '')
    .replace(/owners?/gi, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .trim();
  return s || 'unknown';
}

// ══════════════════════════════════════════════════
// OWNER → EXTERNAL DATA MAPPING
// Reads '_OwnerCamMapping' tab from Ken's national sheet.
// Three columns: Owner Name | Cam Company Name | Cost Sheet ID
// Returns { mapping: { "Owner Name": ["Company A", ...] }, costSheets: { "Owner Name": "sheetId" } }
// ══════════════════════════════════════════════════

function readOwnerCamMapping() {
  var ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var sheet = ss.getSheetByName('_OwnerCamMapping');
  if (!sheet) return { mapping: {}, costSheets: {} };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { mapping: {}, costSheets: {} };

  // Expect headers: Owner Name | Cam Company Name | Cost Sheet ID
  var mapping = {};
  var costSheets = {};
  for (var i = 1; i < data.length; i++) {
    var ownerName = String(data[i][0] || '').trim();
    var camCompany = String(data[i][1] || '').trim();
    var costSheetId = String(data[i][2] || '').trim();
    if (!ownerName) continue;

    // Cam company mapping (col B)
    if (camCompany) {
      if (!mapping[ownerName]) mapping[ownerName] = [];
      mapping[ownerName].push(camCompany);
    }

    // Cost sheet ID (col C) — take first non-empty value per owner
    if (costSheetId && !costSheets[ownerName]) {
      // Extract spreadsheet ID from URL if a full link was pasted
      var idMatch = costSheetId.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
      costSheets[ownerName] = idMatch ? idMatch[1] : costSheetId;
    }
  }

  return { mapping: mapping, costSheets: costSheets };
}


// ── Claim a company for an owner (append to _OwnerCamMapping) ──
function claimCompany(ownerName, companyName) {
  ownerName = String(ownerName || '').trim();
  companyName = String(companyName || '').trim();
  if (!ownerName || !companyName) return { error: 'ownerName and companyName are required' };

  var ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var sheet = ss.getSheetByName('_OwnerCamMapping');

  // Auto-create tab with headers if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('_OwnerCamMapping');
    sheet.getRange(1, 1, 1, 3).setValues([['Owner Name', 'Cam Company Name', 'Cost Sheet ID']]);
  }

  // Check for duplicate
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === ownerName.toLowerCase() &&
        String(data[i][1]).trim().toLowerCase() === companyName.toLowerCase()) {
      var r = readOwnerCamMapping();
      return { ok: true, message: 'Already claimed', mapping: r.mapping, costSheets: r.costSheets };
    }
  }

  // Append new row
  sheet.appendRow([ownerName, companyName]);

  var result = readOwnerCamMapping();
  return { ok: true, mapping: result.mapping, costSheets: result.costSheets };
}


// ── Unclaim a company from an owner (preserves Cost Sheet ID in col C) ──
function unclaimCompany(ownerName, companyName) {
  ownerName = String(ownerName || '').trim();
  companyName = String(companyName || '').trim();
  if (!ownerName || !companyName) return { error: 'ownerName and companyName are required' };

  var ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var sheet = ss.getSheetByName('_OwnerCamMapping');
  if (!sheet) return { ok: true, mapping: {}, costSheets: {} };

  var data = sheet.getDataRange().getValues();
  var ownerLc = ownerName.toLowerCase();

  // Find all rows for this owner and the target row to unclaim
  var ownerRows = [];    // all row indices (0-based) for this owner
  var targetRows = [];   // rows matching the company to remove
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === ownerLc) {
      ownerRows.push(i);
      if (String(data[i][1]).trim().toLowerCase() === companyName.toLowerCase()) {
        targetRows.push(i);
      }
    }
  }

  // Delete target rows bottom-to-top, but if a row has a cost sheet ID in col C,
  // migrate it to another surviving row first
  for (var t = targetRows.length - 1; t >= 0; t--) {
    var ri = targetRows[t];
    var costId = String(data[ri][2] || '').trim();

    if (costId) {
      // Find another row for this owner that isn't being deleted
      var otherRow = -1;
      for (var oi = 0; oi < ownerRows.length; oi++) {
        if (targetRows.indexOf(ownerRows[oi]) === -1) {
          otherRow = ownerRows[oi];
          break;
        }
      }
      if (otherRow >= 0) {
        // Move cost sheet ID to the surviving row
        sheet.getRange(otherRow + 1, 3).setValue(costId);
      } else {
        // No other rows — clear company but keep row so cost sheet ID survives
        sheet.getRange(ri + 1, 2).setValue('');
        continue; // don't delete
      }
    }

    sheet.deleteRow(ri + 1);
    // Adjust remaining indices
    for (var adj = 0; adj < ownerRows.length; adj++) {
      if (ownerRows[adj] > ri) ownerRows[adj]--;
    }
  }

  var result = readOwnerCamMapping();
  return { ok: true, mapping: result.mapping, costSheets: result.costSheets };
}


// ── List files in the Indeed/NLR costs Drive folder (names + IDs, no opens) ──
function listCostFiles() {
  var folder;
  try {
    folder = DriveApp.getFolderById(SHEETS.INDEED_COSTS_FOLDER);
  } catch (err) {
    Logger.log('listCostFiles: Cannot open folder: ' + err.message);
    return { files: [] };
  }

  var result = [];
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    var file = files.next();
    result.push({ name: file.getName().trim(), id: file.getId() });
  }
  // Sort alphabetically
  result.sort(function(a, b) { return a.name.localeCompare(b.name); });
  return { files: result };
}


// ── Read cost data from a single spreadsheet by ID ──
function readCostSheet(sheetId) {
  if (!sheetId) return { error: 'sheetId is required' };

  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var sheets = ss.getSheets();

    // Parse month tabs and sort newest-first
    var monthTabs = [];
    for (var i = 0; i < sheets.length; i++) {
      var tabName = sheets[i].getName().trim();
      var parsed = _parseMonthTab(tabName);
      if (parsed) {
        monthTabs.push({ sheet: sheets[i], name: tabName, date: parsed });
      }
    }
    monthTabs.sort(function(a, b) { return b.date - a.date; });

    if (!monthTabs.length) {
      return { months: [], platformOrder: [] };
    }

    // Extract data from each month tab
    var monthsData = [];
    for (var m = 0; m < monthTabs.length; m++) {
      var costData = _extractIndeedTable(monthTabs[m].sheet);
      if (costData) {
        costData.month = monthTabs[m].name;
        monthsData.push(costData);
      }
    }

    // Collect platform names
    var platSet = {};
    var platformOrder = [];
    for (var mi = 0; mi < monthsData.length; mi++) {
      var order = monthsData[mi].platformOrder || [];
      for (var pi = 0; pi < order.length; pi++) {
        if (!platSet[order[pi]]) {
          platSet[order[pi]] = true;
          platformOrder.push(order[pi]);
        }
      }
    }

    return { months: monthsData, platformOrder: platformOrder };
  } catch (err) {
    Logger.log('readCostSheet: Error reading ' + sheetId + ': ' + err.message);
    return { error: 'Cannot read spreadsheet: ' + err.message };
  }
}


// ── Claim a cost spreadsheet for an owner (write to col C of _OwnerCamMapping) ──
function claimCostSheet(ownerName, sheetId) {
  ownerName = String(ownerName || '').trim();
  sheetId = String(sheetId || '').trim();
  if (!ownerName || !sheetId) return { error: 'ownerName and sheetId are required' };

  var ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var sheet = ss.getSheetByName('_OwnerCamMapping');

  if (!sheet) {
    sheet = ss.insertSheet('_OwnerCamMapping');
    sheet.getRange(1, 1, 1, 3).setValues([['Owner Name', 'Cam Company Name', 'Cost Sheet ID']]);
  }

  var data = sheet.getDataRange().getValues();
  var foundRow = -1;

  // Find first row for this owner
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === ownerName.toLowerCase()) {
      foundRow = i + 1; // 1-indexed
      break;
    }
  }

  if (foundRow > 0) {
    // Update col C on existing row
    sheet.getRange(foundRow, 3).setValue(sheetId);
  } else {
    // Append new row with just owner name + cost sheet ID
    sheet.appendRow([ownerName, '', sheetId]);
  }

  var result = readOwnerCamMapping();
  return { ok: true, mapping: result.mapping, costSheets: result.costSheets };
}


// ── Unclaim a cost spreadsheet from an owner (clear col C in _OwnerCamMapping) ──
function unclaimCostSheet(ownerName) {
  ownerName = String(ownerName || '').trim();
  if (!ownerName) return { error: 'ownerName is required' };

  var ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var sheet = ss.getSheetByName('_OwnerCamMapping');
  if (!sheet) return { ok: true, costSheets: {} };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === ownerName.toLowerCase() &&
        String(data[i][2] || '').trim()) {
      sheet.getRange(i + 1, 3).setValue('');
    }
  }

  var result = readOwnerCamMapping();
  return { ok: true, mapping: result.mapping, costSheets: result.costSheets };
}


// ══════════════════════════════════════════════════
// ONLINE PRESENCE — Read Cam's Performance Audit sheet
// Returns all business rows with review platforms,
// Instagram, website, blog, and SEO data
// ══════════════════════════════════════════════════

function readOnlinePresence() {
  if (!SHEETS.PERFORMANCE_AUDIT) return { businesses: [] };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.PERFORMANCE_AUDIT);
  } catch (err) {
    return { error: 'Cannot open Performance Audit sheet: ' + err.message, businesses: [] };
  }

  // Read from 'vlookup' tab specifically
  var sheet = ss.getSheetByName('vlookup');
  if (!sheet) return { businesses: [] };

  var data = sheet.getDataRange().getValues();
  var headerRowIdx = -1;

  for (var r = 0; r < Math.min(5, data.length); r++) {
    for (var c = 0; c < data[r].length; c++) {
      if (String(data[r][c]).toLowerCase().trim() === 'client name') {
        headerRowIdx = r;
        break;
      }
    }
    if (headerRowIdx >= 0) break;
  }

  if (headerRowIdx < 0) return { businesses: [] };

  var headers = data[headerRowIdx].map(function(h) { return String(h).toLowerCase().trim(); });

  // ── Column mapping using anchor columns + Nth-occurrence for repeating headers ──
  var cols = {
    clientName:      findCol(headers, ['client name']),
    businessName:    findCol(headers, ['business name']),
    accountManager:  findCol(headers, ['account manager']),
    services:        findCol(headers, ['services']),
    auditMonth:      findCol(headers, ['audit month and year', 'audit month']),
    // Platform link anchors
    gblLink:         findCol(headers, ['gbl link']),
    glassdoorLink:   findCol(headers, ['glassdoor link']),
    indeedLink:      findCol(headers, ['indeed link']),
    igLink:          findCol(headers, ['ig link']),
    // Non-repeating columns
    shared:          findCol(headers, ['shared']),
    generated:       findCol(headers, ['generated']),
    followers:       findCol(headers, ['followers']),
    following:       findCol(headers, ['following']),
    website:         findCol(headers, ['website', 'web site', 'site url', 'company website', 'homepage']),
    updatedMonth:    findCol(headers, ['updated this month']),
    sitePhotos:      findCol(headers, ['site photos']),
    lastUpdated:     findCol(headers, ['last updated']),
    blog:            findCol(headers, ['blog']),
    lastBlogPost:    findCol(headers, ['last blog post']),
    threeMonthCount: findCol(headers, ['3 month count']),
    currentMonth:    findCol(headers, ['current month']),
    onQueue:         findCol(headers, ['on queue']),
    seoCheck:        findCol(headers, ['seo check']),
    otherNotes:      findCol(headers, ['other notes / follow-up', 'other notes']),
    full:            findCol(headers, ['full']),
    lite:            findCol(headers, ['lite']),
    status:          findCol(headers, ['status'])
  };

  // 4th review site: its "Link" column is the plain "link" (not prefixed)
  var otherLinkCol = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === 'link') { otherLinkCol = i; break; }
  }

  // Repeating columns: Rating (×4), # of Reviews (×4), Notes (×8)
  var ratingCols  = _findAllExact(headers, 'rating');        // [GBL, Glassdoor, Indeed, Other]
  var reviewCols  = _findAllExact(headers, '# of reviews');  // same order
  var notesCols   = _findAllExact(headers, 'notes');         // [GBL, GD, Indeed, Other, IG, Website, Blog, Final]

  // ── Parse each business row ──
  // Detect per-row column shift: the Performance Audit sheet sometimes has extra
  // cells in data rows that shift values right of where the headers say they are.
  // We detect this from the Instagram URL anchor and apply the shift to ALL
  // columns at or after the IG section (website, blog, SEO, etc).
  var businesses = [];
  var shiftStart = cols.igLink >= 0 ? cols.igLink : 9999;

  for (var i = headerRowIdx + 1; i < data.length; i++) {
    var row = data[i];
    var clientName = String(row[cols.clientName] || '').trim();
    if (!clientName) continue;

    var shift = _getColumnShift(row, cols);

    var biz = {
      clientName: clientName,
      businessName: _str(row[cols.businessName]),
      accountManager: _str(row[cols.accountManager]),
      services: _str(row[cols.services]),
      auditMonth: _str(row[cols.auditMonth]),

      // Reviews — before the IG section, unaffected by shift
      gbl: {
        link:    _str(row[cols.gblLink]),
        rating:  _safeNum(row[ratingCols[0]]),
        reviews: num(row[reviewCols[0]]),
        notes:   _str(row[notesCols[0]])
      },
      glassdoor: {
        link:    _str(row[cols.glassdoorLink]),
        rating:  _safeNum(row[ratingCols[1]]),
        reviews: num(row[reviewCols[1]]),
        notes:   _str(row[notesCols[1]])
      },
      indeed: {
        link:    _str(row[cols.indeedLink]),
        rating:  _safeNum(row[ratingCols[2]]),
        reviews: num(row[reviewCols[2]]),
        notes:   _str(row[notesCols[2]])
      },
      other: {
        link:     _str(row[otherLinkCol]),
        rating:   _safeNum(row[ratingCols[3]]),
        reviews:  num(row[reviewCols[3]]),
        notes:    _str(row[notesCols[3]]),
        platform: _detectPlatform(_str(row[otherLinkCol]))
      },

      // Instagram — uses anchor-based parsing (handles shift internally)
      instagram: _parseInstagram(row, cols, notesCols, shift, shiftStart),

      // Post-IG sections — all use _rs() to apply detected shift
      website: {
        url:          _str(_rs(row, cols.website, shift, shiftStart)),
        updatedMonth: _str(_rs(row, cols.updatedMonth, shift, shiftStart)),
        sitePhotos:   _str(_rs(row, cols.sitePhotos, shift, shiftStart)),
        lastUpdated:  _str(_rs(row, cols.lastUpdated, shift, shiftStart)),
        notes:        notesCols.length > 5 ? _str(_rs(row, notesCols[5], shift, shiftStart)) : ''
      },

      blog: {
        url:             _str(_rs(row, cols.blog, shift, shiftStart)),
        lastPost:        _str(_rs(row, cols.lastBlogPost, shift, shiftStart)),
        threeMonthCount: num(_rs(row, cols.threeMonthCount, shift, shiftStart)),
        currentMonth:    num(_rs(row, cols.currentMonth, shift, shiftStart)),
        onQueue:         num(_rs(row, cols.onQueue, shift, shiftStart)),
        notes:           notesCols.length > 6 ? _str(_rs(row, notesCols[6], shift, shiftStart)) : ''
      },

      seo: {
        check: _str(_rs(row, cols.seoCheck, shift, shiftStart))
      },

      serviceStatus: {
        full:   _str(_rs(row, cols.full, shift, shiftStart)),
        lite:   _str(_rs(row, cols.lite, shift, shiftStart)),
        status: _str(_rs(row, cols.status, shift, shiftStart)),
        notes:  notesCols.length > 7 ? _str(_rs(row, notesCols[7], shift, shiftStart)) : ''
      },

      otherNotes: _str(_rs(row, cols.otherNotes, shift, shiftStart))
    };

    businesses.push(biz);
  }

  // ── Deduplicate: keep only the most recent audit per client+business ──
  // Cam audits the same businesses monthly, creating duplicate rows.
  // Key on clientName|businessName, keep the row closest to the end of the sheet
  // (rows are in chronological order, newest audits at the bottom).
  var deduped = {};
  for (var d = 0; d < businesses.length; d++) {
    var bKey = (businesses[d].clientName + '|' + businesses[d].businessName).toLowerCase();
    deduped[bKey] = businesses[d];  // later rows overwrite earlier ones
  }
  var uniqueBiz = [];
  for (var key in deduped) {
    if (deduped.hasOwnProperty(key)) uniqueBiz.push(deduped[key]);
  }

  // Debug: column mapping, shift detection, and header info for troubleshooting
  var firstShift = 0;
  for (var dbg = headerRowIdx + 1; dbg < Math.min(data.length, headerRowIdx + 20); dbg++) {
    var testShift = _getColumnShift(data[dbg], cols);
    if (testShift > 0) { firstShift = testShift; break; }
  }
  var _debug = {
    headerRow: headerRowIdx,
    headerCount: headers.length,
    shiftDetected: firstShift,
    shiftStart: shiftStart,
    colIndices: cols,
    ratingCols: ratingCols,
    reviewCols: reviewCols,
    notesCols: notesCols,
    headers: headers,  // all headers for inspection
    totalBiz: uniqueBiz.length,
    totalRaw: businesses.length
  };

  // Extract unique company names for the claim dropdown
  var companySet = {};
  for (var ci = 0; ci < uniqueBiz.length; ci++) {
    var cn = String(uniqueBiz[ci].businessName || uniqueBiz[ci].clientName || '').trim();
    if (cn && cn !== 'Grand Total') companySet[cn] = true;
  }
  var allCompanyNames = Object.keys(companySet).sort();

  return { businesses: uniqueBiz, allCompanyNames: allCompanyNames, tabName: sheet.getName(), _debug: _debug };
}

// ══════════════════════════════════════════════════
// READ NLR B2B HEADCOUNT / PRODUCTION
// Each owner has a tab in NLR's spreadsheet.
// Row 1 = headers: Dates | Active | Leaders | Dist | Training |
//                  Personal Production | Production LW | Production Goals
// Returns { owners: { "TabName": { current: {...}, trend: [{...}, ...] } } }
// ══════════════════════════════════════════════════

function readNLRHeadcount() {
  if (!SHEETS.NLR_B2B) return { owners: {} };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NLR_B2B);
  } catch (err) {
    return { error: 'Cannot open NLR B2B sheet: ' + err.message, owners: {} };
  }

  // Tabs to skip (non-owner tabs)
  var SKIP_TABS = {
    'input - sales qual metrics': true,
    'input - comm per rep and total dd': true,
    'input - market metrics': true,
    'sheet2': true
  };

  var allSheets = ss.getSheets();
  var owners = {};

  for (var t = 0; t < allSheets.length; t++) {
    var sheet = allSheets[t];
    var tabName = sheet.getName().trim();

    // Skip non-owner tabs
    if (SKIP_TABS[tabName.toLowerCase()]) continue;

    // Read data — only need first ~30 rows (headcount section is at top)
    var lastRow = Math.min(sheet.getLastRow(), 30);
    if (lastRow < 2) continue;
    var data = sheet.getRange(1, 1, lastRow, 10).getValues();

    // Row 0 = headers
    var headers = data[0].map(function(h) { return String(h).toLowerCase().trim(); });

    // IMPORTANT: "production lw" must match before "production" to avoid
    // hitting "personal production" (col F) which is a different metric.
    // Same for "production goals" before generic "goals".
    var colMap = {
      dates:          findCol(headers, ['dates', 'date']),
      active:         findCol(headers, ['active']),
      leaders:        findCol(headers, ['leaders']),
      dist:           findCol(headers, ['dist']),
      training:       findCol(headers, ['training']),
      productionLW:   findCol(headers, ['production lw']),
      productionGoals:findCol(headers, ['production goals'])
    };

    // Must have at least dates + active columns to be a valid owner tab
    if (colMap.dates < 0 && colMap.active < 0) continue;

    // Walk ALL rows, collecting trend history + tracking last good row
    // ONLY include rows where column A is a real date (skip text like "Owner (Owner>Rep>Zip)")
    var trend = [];
    var lastGood = null;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // Date gate: column A must be a Date object or a parseable date string
      var dateCell = colMap.dates >= 0 ? row[colMap.dates] : null;
      if (!dateCell) continue;
      if (!(dateCell instanceof Date)) {
        var parsed = new Date(dateCell);
        if (isNaN(parsed.getTime())) continue;  // Not a date — skip row
      }

      // Check if row has any numeric data
      var hasActive = colMap.active >= 0 && row[colMap.active] !== '' && row[colMap.active] !== null;
      var hasProd   = colMap.productionLW >= 0 && row[colMap.productionLW] !== '' && row[colMap.productionLW] !== null;
      if (!hasActive && !hasProd) continue;

      var entry = {
        date:            colMap.dates >= 0 ? formatDate(row[colMap.dates]) : '',
        active:          colMap.active >= 0 ? num(row[colMap.active]) : 0,
        leaders:         colMap.leaders >= 0 ? num(row[colMap.leaders]) : 0,
        dist:            colMap.dist >= 0 ? num(row[colMap.dist]) : 0,
        training:        colMap.training >= 0 ? num(row[colMap.training]) : 0,
        productionLW:    colMap.productionLW >= 0 ? num(row[colMap.productionLW]) : 0,
        productionGoals: colMap.productionGoals >= 0 ? num(row[colMap.productionGoals]) : 0
      };
      trend.push(entry);
      lastGood = entry;
    }

    if (lastGood) {
      owners[tabName] = {
        current: lastGood,
        trend: trend
      };
    }
  }

  return { owners: owners };
}

// ══════════════════════════════════════════════════
// IMPORT NLR HEADCOUNT → LOCAL _B2B_Headcount TAB
// Reads NLR data once, writes it into Ken's national sheet
// so we're no longer dependent on the NLR spreadsheet.
// Tab format: Owner | Date | Active | Leaders | Dist | Training | ProductionLW | ProductionGoals
// ══════════════════════════════════════════════════

function importNLRHeadcount() {
  // 1. Read from NLR
  var nlrResult = readNLRHeadcount();
  if (nlrResult.error) return nlrResult;
  var nlrOwners = nlrResult.owners || {};
  if (!Object.keys(nlrOwners).length) return { error: 'No owner data found in NLR sheet' };

  // 2. Open Ken's national sheet
  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message };
  }

  // 3. Get or create _B2B_Headcount tab
  var TAB_NAME = '_B2B_Headcount';
  var sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_NAME);
    sheet.appendRow(['Owner', 'Date', 'Active', 'Leaders', 'Dist', 'Training', 'ProductionLW', 'ProductionGoals']);
  } else {
    // Clear existing data (keep header)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
    }
  }

  // 4. Build rows: one row per owner per date
  var rows = [];
  var ownerNames = Object.keys(nlrOwners).sort();
  for (var o = 0; o < ownerNames.length; o++) {
    var name = ownerNames[o];
    var ownerData = nlrOwners[name];
    var trend = ownerData.trend || [];
    for (var t = 0; t < trend.length; t++) {
      var entry = trend[t];
      rows.push([
        name,
        entry.date,
        entry.active,
        entry.leaders,
        entry.dist,
        entry.training,
        entry.productionLW,
        entry.productionGoals
      ]);
    }
  }

  // 5. Write all rows at once
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 8).setValues(rows);
  }

  return { success: true, ownersImported: ownerNames.length, rowsWritten: rows.length };
}

// ══════════════════════════════════════════════════
// READ NLR B2B APPLIES RECEIVED / SENT TO LIST
// Each owner tab has a section further down:
//   Row N-2: date headers (col C onward)
//   Row N:   "Applies Received" in col A, values in col C+
//   Row N+1: "Sent to List" in col A, values in col C+
// Returns { owners: { "TabName": { dates: [...], appliesReceived: [...], sentToList: [...] } } }
// ══════════════════════════════════════════════════

function readNLRApplies() {
  if (!SHEETS.NLR_B2B) return { owners: {} };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NLR_B2B);
  } catch (err) {
    return { error: 'Cannot open NLR B2B sheet: ' + err.message, owners: {} };
  }

  var SKIP_TABS = {
    'input - sales qual metrics': true,
    'input - comm per rep and total dd': true,
    'input - market metrics': true,
    'sheet2': true
  };

  var allSheets = ss.getSheets();
  var owners = {};

  for (var t = 0; t < allSheets.length; t++) {
    var sheet = allSheets[t];
    var tabName = sheet.getName().trim();
    if (SKIP_TABS[tabName.toLowerCase()]) continue;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 4 || lastCol < 3) continue;

    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    // Find "Applies Received" row in column A
    var appliesRow = -1;
    for (var i = 0; i < data.length; i++) {
      var cellA = String(data[i][0] || '').trim().toLowerCase();
      if (cellA === 'applies received') { appliesRow = i; break; }
    }
    if (appliesRow < 2) continue; // Need at least 2 rows above for headers

    // Headers are 2 rows above, starting from column C (index 2)
    var headerRow = appliesRow - 2;
    var dates = [];
    for (var c = 2; c < data[headerRow].length; c++) {
      var h = data[headerRow][c];
      if (!h && h !== 0) break;
      dates.push(_normalizeDate(h));
    }
    if (!dates.length) continue;

    // Read Applies Received values
    var appliesValues = [];
    for (var c = 2; c < 2 + dates.length; c++) {
      appliesValues.push(num(data[appliesRow][c]));
    }

    // Read Sent to List values (next row)
    var stlValues = [];
    var stlRow = appliesRow + 1;
    if (stlRow < data.length) {
      var stlLabel = String(data[stlRow][0] || '').trim().toLowerCase();
      if (stlLabel === 'sent to list') {
        for (var c = 2; c < 2 + dates.length; c++) {
          stlValues.push(num(data[stlRow][c]));
        }
      }
    }

    owners[tabName] = {
      dates: dates,
      appliesReceived: appliesValues,
      sentToList: stlValues
    };
  }

  return { owners: owners };
}

// ══════════════════════════════════════════════════
// IMPORT NLR APPLIES → LOCAL _B2B_Applies TAB
// Reads applies/STL from NLR B2B, writes to Ken's sheet.
// Tab format: Owner | Date | AppliesReceived | SentToList
// ══════════════════════════════════════════════════

function importNLRApplies() {
  var nlrResult = readNLRApplies();
  if (nlrResult.error) return nlrResult;
  var nlrOwners = nlrResult.owners || {};
  if (!Object.keys(nlrOwners).length) return { error: 'No applies data found in NLR sheet' };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message };
  }

  var TAB_NAME = '_B2B_Applies';
  var sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_NAME);
    sheet.appendRow(['Owner', 'Date', 'AppliesReceived', 'SentToList']);
  } else {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
    }
  }

  var rows = [];
  var ownerNames = Object.keys(nlrOwners).sort();
  for (var o = 0; o < ownerNames.length; o++) {
    var name = ownerNames[o];
    var d = nlrOwners[name];
    for (var i = 0; i < d.dates.length; i++) {
      rows.push([
        name,
        d.dates[i],
        d.appliesReceived[i] || 0,
        d.sentToList.length > i ? d.sentToList[i] : 0
      ]);
    }
  }

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }

  return { success: true, ownersImported: ownerNames.length, rowsWritten: rows.length };
}

// ══════════════════════════════════════════════════
// READ LOCAL APPLIES from _B2B_Applies tab
// Returns { owners: { "Name": { "MM/DD/YYYY": { applies: N, stl: N } } } }
// ══════════════════════════════════════════════════

function readLocalApplies() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { owners: {} };
  }

  var sheet = ss.getSheetByName('_B2B_Applies');
  if (!sheet) return { owners: {} };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { owners: {} };

  var owners = {};
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0] || '').trim();
    if (!name) continue;
    var date = _normalizeDate(data[i][1]);
    if (!date || date === '—') continue;

    if (!owners[name]) owners[name] = {};
    owners[name][date] = {
      applies: num(data[i][2]),
      stl: num(data[i][3])
    };
  }

  return { owners: owners };
}

// ══════════════════════════════════════════════════
// READ NLR PRODUCTION from NLR B2B owner tabs
// Each tab has a "Production" merged cell in column J (index 9).
// Below that: headers row with "Owner Name","Rep","Total Volume", etc.
// Then owner total row + per-rep rows.
// ══════════════════════════════════════════════════

function readNLRProduction() {
  if (!SHEETS.NLR_B2B) return { owners: {} };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NLR_B2B);
  } catch (err) {
    return { error: 'Cannot open NLR B2B sheet: ' + err.message, owners: {} };
  }

  var SKIP_TABS = {
    'input - sales qual metrics': true,
    'input - comm per rep and total dd': true,
    'input - market metrics': true,
    'sheet2': true
  };

  var allSheets = ss.getSheets();
  var owners = {};

  for (var t = 0; t < allSheets.length; t++) {
    var sheet = allSheets[t];
    var tabName = sheet.getName().trim();
    if (SKIP_TABS[tabName.toLowerCase()]) continue;

    var lastRow = sheet.getLastRow();
    var lastCol = Math.min(sheet.getLastColumn(), 29); // Cap at AC
    if (lastRow < 10 || lastCol < 15) continue;

    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    // Find "Production" in column J (index 9)
    var productionRow = -1;
    for (var i = 0; i < data.length; i++) {
      var cellJ = String(data[i][9] || '').trim().toLowerCase();
      if (cellJ === 'production') {
        productionRow = i;
        break;
      }
    }
    if (productionRow < 0) continue;

    // Find the main headers row: look for "Owner Name" in col J below the Production marker
    var headersRow = -1;
    for (var i = productionRow + 1; i < Math.min(data.length, productionRow + 15); i++) {
      var cellJ = String(data[i][9] || '').trim().toLowerCase();
      if (cellJ === 'owner name') {
        headersRow = i;
        break;
      }
    }
    if (headersRow < 0) continue;

    // Build column map from headers (col J onward = index 9+)
    var headers = [];
    for (var c = 9; c < data[headersRow].length; c++) {
      headers.push(String(data[headersRow][c] || '').trim().toLowerCase());
    }

    var colIdx = function(patterns) {
      for (var p = 0; p < patterns.length; p++) {
        for (var i = 0; i < headers.length; i++) {
          if (headers[i].indexOf(patterns[p]) >= 0) return i;
        }
      }
      return -1;
    };

    var cols = {
      ownerName:    colIdx(['owner name']),
      rep:          colIdx(['rep']),
      totalVolume:  colIdx(['total volume']),
      repCount:     colIdx(['rep count']),
      salesPerRep:  colIdx(['sales per rep']),
      orderCount:   colIdx(['order count']),
      ordersBefore: colIdx(['orders before']),
      earlyPct:     colIdx(['early order']),
      ordersAfter:  colIdx(['orders after']),
      latePct:      colIdx(['late order']),
      internet:     colIdx(['internet']),
      voip:         colIdx(['voip']),
      wireless:     colIdx(['wrls']),
      airAwb:       colIdx(['air/awb', 'air']),
      weekendPct:   colIdx(['weekend']),
      tierPct:      colIdx(['tier']),
      abpPct:       colIdx(['abp']),
      cruPct:       colIdx(['cru']),
      newWrlsPct:   colIdx(['new wireless']),
      byodPct:      colIdx(['byod'])
    };

    // Read data rows below headers (owner total + reps)
    var ownerTotal = null;
    var reps = [];

    for (var i = headersRow + 1; i < data.length; i++) {
      var row = data[i];
      var nameVal = cols.ownerName >= 0 ? String(row[9 + cols.ownerName] || '').trim() : '';
      var repVal = cols.rep >= 0 ? String(row[9 + cols.rep] || '').trim() : '';

      // Owner total row has owner name in col J + "Total" in col K
      // Rep rows have empty col J and rep name in col K
      if (!nameVal && !repVal) break; // Both empty = end of section

      var entry = {
        name: nameVal || repVal,
        rep: repVal,
        totalVolume:  cols.totalVolume >= 0 ? num(row[9 + cols.totalVolume]) : 0,
        repCount:     cols.repCount >= 0 ? num(row[9 + cols.repCount]) : 0,
        salesPerRep:  cols.salesPerRep >= 0 ? numDec(row[9 + cols.salesPerRep]) : 0,
        orderCount:   cols.orderCount >= 0 ? num(row[9 + cols.orderCount]) : 0,
        ordersBefore: cols.ordersBefore >= 0 ? num(row[9 + cols.ordersBefore]) : 0,
        earlyPct:     cols.earlyPct >= 0 ? numDec(row[9 + cols.earlyPct]) : 0,
        ordersAfter:  cols.ordersAfter >= 0 ? num(row[9 + cols.ordersAfter]) : 0,
        latePct:      cols.latePct >= 0 ? numDec(row[9 + cols.latePct]) : 0,
        internet:     cols.internet >= 0 ? num(row[9 + cols.internet]) : 0,
        voip:         cols.voip >= 0 ? num(row[9 + cols.voip]) : 0,
        wireless:     cols.wireless >= 0 ? num(row[9 + cols.wireless]) : 0,
        airAwb:       cols.airAwb >= 0 ? num(row[9 + cols.airAwb]) : 0,
        weekendPct:   cols.weekendPct >= 0 ? numDec(row[9 + cols.weekendPct]) : 0,
        tierPct:      cols.tierPct >= 0 ? numDec(row[9 + cols.tierPct]) : 0,
        abpPct:       cols.abpPct >= 0 ? numDec(row[9 + cols.abpPct]) : 0,
        cruPct:       cols.cruPct >= 0 ? numDec(row[9 + cols.cruPct]) : 0,
        newWrlsPct:   cols.newWrlsPct >= 0 ? numDec(row[9 + cols.newWrlsPct]) : 0,
        byodPct:      cols.byodPct >= 0 ? numDec(row[9 + cols.byodPct]) : 0
      };

      // First row is the owner total (rep column = "Total")
      if (repVal.toLowerCase() === 'total' || !ownerTotal) {
        ownerTotal = entry;
      } else {
        reps.push(entry);
      }
    }

    if (ownerTotal) {
      owners[tabName] = {
        summary: ownerTotal,
        reps: reps
      };
    }
  }

  return { owners: owners };
}

// ── Helper: num with decimal preservation ──
function numDec(v) {
  if (v === null || v === undefined || v === '') return 0;
  var n = Number(v);
  return isNaN(n) ? 0 : Math.round(n * 100) / 100;
}

// ══════════════════════════════════════════════════
// IMPORT NLR PRODUCTION → LOCAL _B2B_Production TAB
// ══════════════════════════════════════════════════

function importNLRProduction() {
  var nlrResult = readNLRProduction();
  if (nlrResult.error) return nlrResult;
  var nlrOwners = nlrResult.owners || {};
  if (!Object.keys(nlrOwners).length) return { error: 'No production data found in NLR sheet' };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message };
  }

  var TAB_NAME = '_B2B_Production';
  var HEADERS = ['Owner', 'Type', 'Name', 'Rep', 'TotalVolume', 'RepCount', 'SalesPerRep',
    'OrderCount', 'OrdersBefore', 'EarlyPct', 'OrdersAfter', 'LatePct',
    'Internet', 'VOIP', 'Wireless', 'AirAwb', 'WeekendPct', 'TierPct',
    'AbpPct', 'CruPct', 'NewWrlsPct', 'ByodPct'];

  var sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_NAME);
    sheet.appendRow(HEADERS);
  } else {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
    }
  }

  var rows = [];
  var ownerNames = Object.keys(nlrOwners).sort();
  for (var o = 0; o < ownerNames.length; o++) {
    var oName = ownerNames[o];
    var d = nlrOwners[oName];

    // Write summary row
    var s = d.summary;
    rows.push([oName, 'summary', s.name, s.rep, s.totalVolume, s.repCount, s.salesPerRep,
      s.orderCount, s.ordersBefore, s.earlyPct, s.ordersAfter, s.latePct,
      s.internet, s.voip, s.wireless, s.airAwb, s.weekendPct, s.tierPct,
      s.abpPct, s.cruPct, s.newWrlsPct, s.byodPct]);

    // Write rep rows
    for (var r = 0; r < d.reps.length; r++) {
      var rep = d.reps[r];
      rows.push([oName, 'rep', rep.name, rep.rep, rep.totalVolume, rep.repCount, rep.salesPerRep,
        rep.orderCount, rep.ordersBefore, rep.earlyPct, rep.ordersAfter, rep.latePct,
        rep.internet, rep.voip, rep.wireless, rep.airAwb, rep.weekendPct, rep.tierPct,
        rep.abpPct, rep.cruPct, rep.newWrlsPct, rep.byodPct]);
    }
  }

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  }

  return { success: true, ownersImported: ownerNames.length, rowsWritten: rows.length };
}

// ══════════════════════════════════════════════════
// READ LOCAL PRODUCTION from _B2B_Production tab
// Returns { owners: { "Name": { summary: {...}, reps: [{...}] } } }
// ══════════════════════════════════════════════════

function readLocalProduction() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { owners: {} };
  }

  var sheet = ss.getSheetByName('_B2B_Production');
  if (!sheet) return { owners: {} };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { owners: {} };

  var owners = {};
  for (var i = 1; i < data.length; i++) {
    var oName = String(data[i][0] || '').trim();
    if (!oName) continue;
    var type = String(data[i][1] || '').trim();

    var entry = {
      name:         String(data[i][2] || ''),
      rep:          String(data[i][3] || ''),
      totalVolume:  num(data[i][4]),
      repCount:     num(data[i][5]),
      salesPerRep:  numDec(data[i][6]),
      orderCount:   num(data[i][7]),
      ordersBefore: num(data[i][8]),
      earlyPct:     numDec(data[i][9]),
      ordersAfter:  num(data[i][10]),
      latePct:      numDec(data[i][11]),
      internet:     num(data[i][12]),
      voip:         num(data[i][13]),
      wireless:     num(data[i][14]),
      airAwb:       num(data[i][15]),
      weekendPct:   numDec(data[i][16]),
      tierPct:      numDec(data[i][17]),
      abpPct:       numDec(data[i][18]),
      cruPct:       numDec(data[i][19]),
      newWrlsPct:   numDec(data[i][20]),
      byodPct:      numDec(data[i][21])
    };

    if (!owners[oName]) {
      owners[oName] = { summary: null, reps: [] };
    }

    if (type === 'summary') {
      owners[oName].summary = entry;
    } else {
      owners[oName].reps.push(entry);
    }
  }

  return { owners: owners };
}

// ══════════════════════════════════════════════════
// READ LOCAL HEADCOUNT from _B2B_Headcount tab
// Returns same shape as readNLRHeadcount():
// { owners: { "Name": { current: {...}, trend: [{...}] } } }
// ══════════════════════════════════════════════════

function readLocalHeadcount() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message, owners: {} };
  }

  var sheet = ss.getSheetByName('_B2B_Headcount');
  if (!sheet) return { owners: {} };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { owners: {} };

  // Header row 0: Owner | Date | Active | Leaders | Dist | Training | ProductionLW | ProductionGoals
  var owners = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = String(row[0] || '').trim();
    if (!name) continue;

    var entry = {
      date:            _normalizeDate(row[1]),
      active:          num(row[2]),
      leaders:         num(row[3]),
      dist:            num(row[4]),
      training:        num(row[5]),
      productionLW:    num(row[6]),
      productionGoals: num(row[7])
    };

    if (!owners[name]) {
      owners[name] = { current: entry, trend: [] };
    }
    owners[name].trend.push(entry);
    owners[name].current = entry; // Last row wins as "current"
  }

  return { owners: owners };
}

// ══════════════════════════════════════════════════
// UPDATE HEADCOUNT ROW in _B2B_Headcount tab
// Finds row by Owner (col A) + Date (col B), updates cols C-F
// ══════════════════════════════════════════════════

function updateHeadcountRow(ownerName, date, active, leaders, dist, training) {
  if (!ownerName || !date) return { error: 'ownerName and date are required' };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message };
  }

  var sheet = ss.getSheetByName('_B2B_Headcount');
  if (!sheet) return { error: '_B2B_Headcount tab not found' };

  var data = sheet.getDataRange().getValues();
  var normDate = _normalizeDate(date);
  var targetRow = -1;

  for (var i = 1; i < data.length; i++) {
    var rowOwner = String(data[i][0] || '').trim();
    var rowDate  = _normalizeDate(data[i][1]);
    if (rowOwner === ownerName && rowDate === normDate) {
      targetRow = i + 1; // 1-based sheet row
      break;
    }
  }

  if (targetRow === -1) {
    return { error: 'Row not found for owner "' + ownerName + '" date "' + normDate + '"' };
  }

  // Update columns C-F (Active, Leaders, Dist, Training) — 1-based cols 3-6
  sheet.getRange(targetRow, 3, 1, 4).setValues([
    [parseInt(active) || 0, parseInt(leaders) || 0, parseInt(dist) || 0, parseInt(training) || 0]
  ]);

  return { ok: true, row: targetRow, owner: ownerName, date: normDate };
}

// ══════════════════════════════════════════════════
// UPDATE PRODUCTION ROW in _B2B_Headcount tab
// Finds row by Owner (col A) + Date (col B), updates cols G-H
// ══════════════════════════════════════════════════

function updateProductionRow(ownerName, date, productionLW, productionGoals) {
  if (!ownerName || !date) return { error: 'ownerName and date are required' };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message };
  }

  var sheet = ss.getSheetByName('_B2B_Headcount');
  if (!sheet) return { error: '_B2B_Headcount tab not found' };

  var data = sheet.getDataRange().getValues();
  var normDate = _normalizeDate(date);
  var targetRow = -1;

  for (var i = 1; i < data.length; i++) {
    var rowOwner = String(data[i][0] || '').trim();
    var rowDate  = _normalizeDate(data[i][1]);
    if (rowOwner === ownerName && rowDate === normDate) {
      targetRow = i + 1; // 1-based sheet row
      break;
    }
  }

  if (targetRow === -1) {
    return { error: 'Row not found for owner "' + ownerName + '" date "' + normDate + '"' };
  }

  // Update columns G-H (ProductionLW, ProductionGoals) — 1-based cols 7-8
  sheet.getRange(targetRow, 7, 1, 2).setValues([
    [parseInt(productionLW) || 0, parseInt(productionGoals) || 0]
  ]);

  return { ok: true, row: targetRow, owner: ownerName, date: normDate };
}

// ── Normalize any date value to MM/DD/YYYY string (no timezone) ──
function _normalizeDate(v) {
  if (!v) return '';
  // Already a Date object (Google Sheets auto-parsed)
  if (v instanceof Date) {
    var mm = ('0' + (v.getMonth() + 1)).slice(-2);
    var dd = ('0' + v.getDate()).slice(-2);
    return mm + '/' + dd + '/' + v.getFullYear();
  }
  // String — manually parse M-D-YYYY or M/D/YYYY to avoid timezone drift
  var s = String(v).trim();
  // Strip any trailing timezone text (e.g. " GMT-0500")
  s = s.replace(/\s+(GMT|UTC|EST|CST|MST|PST|EDT|CDT|MDT|PDT).*$/i, '');
  // Match M/D/YYYY or M-D-YYYY patterns
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) {
    var mm = ('0' + m[1]).slice(-2);
    var dd = ('0' + m[2]).slice(-2);
    return mm + '/' + dd + '/' + m[3];
  }
  return s; // fallback
}

// ── Find ALL exact occurrences of a header text ──
function _findAllExact(headers, text) {
  var indices = [];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === text) indices.push(i);
  }
  return indices;
}

// ── Safe string extraction ──
function _str(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return formatDate(v);
  var s = String(v).trim();
  // Filter out spreadsheet formula errors
  if (s.charAt(0) === '#' && (s.indexOf('REF') >= 0 || s.indexOf('N/A') >= 0 ||
      s.indexOf('VALUE') >= 0 || s.indexOf('DIV') >= 0 || s.indexOf('ERROR') >= 0)) return '';
  return s;
}

// ── Number or null (for ratings where 0 means "no rating") ──
function _safeNum(v) {
  if (v === null || v === undefined || v === '') return null;
  var n = Number(v);
  return isNaN(n) ? null : Math.round(n * 10) / 10;
}

// ── Instagram column parser — handles header/data misalignment ──
// Uses _rs() with header-detected column positions + per-row shift.
// The IG columns are NOT contiguous (notes col sits between link and shared),
// so we can't use anchor-relative offsets. Instead, read each metric from
// its proper header-defined column position with shift applied.
function _parseInstagram(row, cols, notesCols, shift, shiftStart) {
  if (cols.igLink < 0) return { link: '', shared: 0, generated: 0, followers: 0, following: 0, notes: '' };

  var s  = shift || 0;
  var ss = shiftStart || 9999;

  // IG notes — apply shift if the notes column is in the shifted region
  var igNotes = '';
  if (notesCols.length > 4 && notesCols[4] >= 0) {
    igNotes = _str(_rs(row, notesCols[4], s, ss));
  }

  return {
    link:      _str(_rs(row, cols.igLink, s, ss)),
    shared:    num(_rs(row, cols.shared, s, ss)),
    generated: num(_rs(row, cols.generated, s, ss)),
    followers: num(_rs(row, cols.followers, s, ss)),
    following: num(_rs(row, cols.following, s, ss)),
    notes:     igNotes
  };
}

/**
 * Detect column shift in data rows vs headers.
 * The Performance Audit sheet sometimes has extra cells in data rows
 * that shift values right of where the headers indicate.
 * Uses the Instagram URL as an anchor to detect the offset.
 */
function _getColumnShift(row, cols) {
  if (cols.igLink < 0) return 0;
  for (var offset = 0; offset <= 3; offset++) {
    var val = String(row[cols.igLink + offset] || '').toLowerCase();
    if (val.indexOf('instagram.com') >= 0 || val.indexOf('ig.com') >= 0) {
      return offset;
    }
  }
  return 0;
}

/**
 * Read a cell value with shift applied for columns at or past the shift start point.
 * Returns undefined for cols that are -1 (not found in headers).
 */
function _rs(row, colIdx, shift, shiftStart) {
  if (colIdx == null || colIdx < 0) return undefined;
  var actualIdx = colIdx >= shiftStart ? colIdx + shift : colIdx;
  return actualIdx < row.length ? row[actualIdx] : undefined;
}

function _detectPlatform(url) {
  if (!url) return '';
  var u = url.toLowerCase();
  if (u.indexOf('yelp') >= 0) return 'Yelp';
  if (u.indexOf('bbb') >= 0 || u.indexOf('betterbusiness') >= 0) return 'BBB';
  if (u.indexOf('facebook') >= 0 || u.indexOf('fb.com') >= 0) return 'Facebook';
  if (u.indexOf('google') >= 0) return 'Google';
  if (u.indexOf('trustpilot') >= 0) return 'Trustpilot';
  if (u.indexOf('angi') >= 0 || u.indexOf('angieslist') >= 0) return 'Angi';
  if (u.indexOf('nextdoor') >= 0) return 'Nextdoor';
  if (u.indexOf('linkedin') >= 0) return 'LinkedIn';
  if (u.indexOf('twitter') >= 0 || u.indexOf('x.com') >= 0) return 'X';
  return 'Other';
}

// ══════════════════════════════════════════════════
// HELPER: Header-based column finder
// Handles messy sheets where columns move around
// ══════════════════════════════════════════════════

function findCol(headers, patterns) {
  for (var p = 0; p < patterns.length; p++) {
    var pattern = patterns[p].toLowerCase();
    for (var i = 0; i < headers.length; i++) {
      if (headers[i].indexOf(pattern) >= 0) return i;
    }
  }
  return -1;
}

function findColumns(headers, spec) {
  var result = {};
  for (var key in spec) {
    if (Array.isArray(spec[key])) {
      result[key] = findCol(headers, spec[key]);
    } else {
      result[key] = spec[key]; // already resolved (e.g., from findNthRetention)
    }
  }
  return result;
}

// Find the Nth occurrence of "retention" in headers (since there are multiple)
function findNthRetention(headers, n) {
  var count = 0;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === 'retention') {
      count++;
      if (count === n) return i;
    }
  }
  return -1;
}

// ══════════════════════════════════════════════════
// HELPER: Value formatters
// ══════════════════════════════════════════════════

function num(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (v instanceof Date) return 0;  // Don't cast Date objects to epoch
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

function val(v) {
  if (v === null || v === undefined || v === '') return '—';
  return String(v);
}

function pct(v) {
  if (v === null || v === undefined || v === '') return '—';
  var s = String(v);
  if (s.indexOf('%') >= 0) return s;
  var n = Number(v);
  if (isNaN(n)) return s;
  // If it's a decimal (0.65), convert to percentage
  if (n > 0 && n <= 1) return Math.round(n * 100) + '%';
  return Math.round(n) + '%';
}

function formatDate(v) {
  if (!v) return '—';
  if (v instanceof Date) {
    var mm = ('0' + (v.getMonth() + 1)).slice(-2);
    var dd = ('0' + v.getDate()).slice(-2);
    return mm + '/' + dd + '/' + v.getFullYear();
  }
  // String — strip timezone, normalize M-D-YYYY or M/D/YYYY → MM/DD/YYYY
  var s = String(v).trim();
  s = s.replace(/\s+(GMT|UTC|EST|CST|MST|PST|EDT|CDT|MDT|PDT).*$/i, '');
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) {
    return ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + m[3];
  }
  return s;
}

// ══════════════════════════════════════════════════
// SINGLE OWNER DETAIL (for lazy loading)
// ══════════════════════════════════════════════════

function loadOwnerDetail(campaignKey, ownerName) {
  var cfg = CAMPAIGNS[campaignKey];
  if (!cfg) return { error: 'Unknown campaign' };

  var ownerDef = null;
  for (var i = 0; i < cfg.owners.length; i++) {
    if (cfg.owners[i].name.toLowerCase() === ownerName.toLowerCase() ||
        cfg.owners[i].tab.toLowerCase() === ownerName.toLowerCase()) {
      ownerDef = cfg.owners[i];
      break;
    }
  }
  if (!ownerDef) return { error: 'Owner not found: ' + ownerName };

  var ctSS = SpreadsheetApp.openById(SHEETS.CAMPAIGN_TRACKER);
  var owner = loadOwnerFromCampaignTracker(ctSS, ownerDef, cfg);

  // Also load Performance Audit data for this owner
  try {
    enrichWithAuditData(owner, ownerDef);
  } catch (err) {
    Logger.log('Audit enrichment error: ' + err.message);
  }

  return owner;
}

// ══════════════════════════════════════════════════
// PERFORMANCE AUDIT ENRICHMENT
// ══════════════════════════════════════════════════

function enrichWithAuditData(owner, ownerDef) {
  var ss = SpreadsheetApp.openById(SHEETS.PERFORMANCE_AUDIT);
  var auditSheet = ss.getSheetByName('Audit');
  if (!auditSheet) return;

  var data = auditSheet.getDataRange().getValues();
  if (data.length < 2) return;

  var headers = data[1].map(function(h) { return String(h).toLowerCase().trim(); }); // Row 2 is headers (row 1 is merged)
  var colMap = {
    clientName: findCol(headers, ['client name']),
    businessName: findCol(headers, ['business name']),
    accountManager: findCol(headers, ['account manager']),
    services: findCol(headers, ['services']),
    month: findCol(headers, ['month', 'edit month']),
    rating: findCol(headers, ['rating']),
    reviews: findCol(headers, ['# of reviews', 'number of reviews']),
    notes: findCol(headers, ['notes']),
    seoCheck: findCol(headers, ['seo check'])
  };

  // Search for businesses matching this owner
  // TODO: Need a mapping from owner name → business names in the audit sheet
  // For now, we'll search by partial name match
  var searchName = ownerDef.name.toLowerCase();
  var matches = [];

  for (var i = 2; i < data.length; i++) {
    var row = data[i];
    var client = String(row[colMap.clientName] || '').toLowerCase();
    var business = String(row[colMap.businessName] || '').toLowerCase();
    var manager = String(row[colMap.accountManager] || '').toLowerCase();

    // Match on client name, business name, or account manager
    if (client.indexOf(searchName) >= 0 || business.indexOf(searchName) >= 0 || manager.indexOf(searchName) >= 0) {
      matches.push({
        client: val(row[colMap.clientName]),
        business: val(row[colMap.businessName]),
        services: val(row[colMap.services]),
        rating: val(row[colMap.rating]),
        reviews: val(row[colMap.reviews]),
        notes: val(row[colMap.notes]),
        seo: val(row[colMap.seoCheck])
      });
    }
  }

  // Calculate simple grades from the data
  if (matches.length > 0) {
    var avgRating = 0, totalReviews = 0, ratingCount = 0;
    matches.forEach(function(m) {
      var r = parseFloat(m.rating);
      if (!isNaN(r)) { avgRating += r; ratingCount++; }
      totalReviews += num(m.reviews);
    });
    avgRating = ratingCount > 0 ? (avgRating / ratingCount) : 0;

    owner.audit = {
      grades: {
        reviews: ratingToGrade(avgRating),
        website: '—',
        social: '—',
        seo: '—'
      },
      details: {
        'Avg Rating': { value: avgRating.toFixed(1), notes: ratingCount + ' businesses' },
        'Total Reviews': { value: totalReviews, notes: '' },
        'Businesses Tracked': { value: matches.length, notes: '' }
      }
    };
  }
}

function ratingToGrade(rating) {
  if (rating >= 4.5) return 'A+';
  if (rating >= 4.0) return 'A';
  if (rating >= 3.5) return 'B';
  if (rating >= 3.0) return 'C';
  if (rating >= 2.0) return 'D';
  return 'F';
}

// ══════════════════════════════════════════════════
// INDEED AD SPEND — Per-Owner Drive Folder Reader
// ══════════════════════════════════════════════════

/**
 * Read Indeed ad cost tables from per-owner spreadsheets in a shared Drive folder.
 * Each spreadsheet is named after the owner and has monthly tabs with a fixed table.
 * Returns { owners: { "OwnerName": { months: [{ month, local, nlr, total }] } } }
 */
function readIndeedCosts(ownerFilter) {
  var folder;
  try {
    folder = DriveApp.getFolderById(SHEETS.INDEED_COSTS_FOLDER);
  } catch (err) {
    Logger.log('readIndeedCosts: Cannot open folder: ' + err.message);
    return { owners: {} };
  }

  // Build filter set from comma-separated owner names (lowercase for fuzzy match)
  var filterNames = [];
  if (ownerFilter && ownerFilter.trim()) {
    filterNames = ownerFilter.split(',').map(function(n) { return n.trim().toLowerCase(); });
  }

  var owners = {};
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var skipped = 0;

  while (files.hasNext()) {
    var file = files.next();
    var ownerName = file.getName().trim();

    // If filter provided, skip files that don't fuzzy-match any filter name
    if (filterNames.length > 0) {
      var fileLc = ownerName.toLowerCase();
      var matched = false;
      for (var fi = 0; fi < filterNames.length; fi++) {
        if (fileLc === filterNames[fi] ||
            fileLc.indexOf(filterNames[fi]) >= 0 ||
            filterNames[fi].indexOf(fileLc) >= 0) {
          matched = true;
          break;
        }
      }
      if (!matched) { skipped++; continue; }
    }

    try {
      var ss = SpreadsheetApp.openById(file.getId());
      var sheets = ss.getSheets();

      // Parse month tabs and sort newest-first
      var monthTabs = [];
      for (var i = 0; i < sheets.length; i++) {
        var tabName = sheets[i].getName().trim();
        var parsed = _parseMonthTab(tabName);
        if (parsed) {
          monthTabs.push({ sheet: sheets[i], name: tabName, date: parsed });
        }
      }
      monthTabs.sort(function(a, b) { return b.date - a.date; });

      if (!monthTabs.length) {
        Logger.log('readIndeedCosts: No month tabs found for ' + ownerName);
        continue;
      }

      // Extract data from each month tab
      var monthsData = [];
      for (var m = 0; m < monthTabs.length; m++) {
        var costData = _extractIndeedTable(monthTabs[m].sheet);
        if (costData) {
          costData.month = monthTabs[m].name;
          monthsData.push(costData);
        }
      }

      if (monthsData.length) {
        owners[ownerName] = { months: monthsData };
      }
    } catch (err) {
      Logger.log('readIndeedCosts: Error reading ' + ownerName + ': ' + err.message);
    }
  }

  // Collect union of all platform names across all owners/months (preserving order)
  var allPlatformsSet = {};
  var allPlatforms = [];
  var ownerKeys = Object.keys(owners);
  for (var oi = 0; oi < ownerKeys.length; oi++) {
    var months = owners[ownerKeys[oi]].months || [];
    for (var mi = 0; mi < months.length; mi++) {
      var order = months[mi].platformOrder || [];
      for (var pi = 0; pi < order.length; pi++) {
        if (!allPlatformsSet[order[pi]]) {
          allPlatformsSet[order[pi]] = true;
          allPlatforms.push(order[pi]);
        }
      }
    }
  }

  Logger.log('readIndeedCosts: Found data for ' + Object.keys(owners).length + ' owners (skipped ' + skipped + '), platforms: ' + allPlatforms.join(', '));
  return { owners: owners, allPlatforms: allPlatforms };
}

/**
 * Parse a tab name as a month. Handles:
 * "January 2026", "Jan 2026", "Feb", "March 26", "01/2026", "2026-03"
 */
function _parseMonthTab(name) {
  var s = name.trim();
  var MONTHS = {
    'jan': 0, 'january': 0, 'feb': 1, 'february': 1,
    'mar': 2, 'march': 2, 'apr': 3, 'april': 3,
    'may': 4, 'jun': 5, 'june': 5, 'jul': 6, 'july': 6,
    'aug': 7, 'august': 7, 'sep': 8, 'sept': 8, 'september': 8,
    'oct': 9, 'october': 9, 'nov': 10, 'november': 10, 'dec': 11, 'december': 11
  };

  // "Month Year" or just "Month" (e.g., "January 2026", "Feb 26", "March")
  var m1 = s.match(/^([a-zA-Z]+)\s*(\d{2,4})?$/);
  if (m1) {
    var idx = MONTHS[m1[1].toLowerCase()];
    if (idx !== undefined) {
      var yr = m1[2] ? parseInt(m1[2]) : new Date().getFullYear();
      if (yr < 100) yr += 2000;
      return new Date(yr, idx, 1);
    }
  }

  // "MM/YYYY" or "M-YYYY"
  var m2 = s.match(/^(\d{1,2})[\/\-](\d{4})$/);
  if (m2) {
    var mo = parseInt(m2[1]) - 1;
    if (mo >= 0 && mo <= 11) return new Date(parseInt(m2[2]), mo, 1);
  }

  // "YYYY-MM"
  var m3 = s.match(/^(\d{4})[\/\-](\d{1,2})$/);
  if (m3) {
    var mo2 = parseInt(m3[2]) - 1;
    if (mo2 >= 0 && mo2 <= 11) return new Date(parseInt(m3[1]), mo2, 1);
  }

  return null;
}

/**
 * Extract the recruiting cost table from a monthly sheet tab.
 * Dynamically detects all platform columns (Local Indeed, Careerbuilder,
 * Handshake, JazzHR, etc.) by finding the TOTAL column and capturing
 * everything between the label column and TOTAL.
 * Returns { platforms: { "Local Indeed": {...}, ... }, total: {...}, platformOrder: [...] }
 */
function _extractIndeedTable(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 3) return null;

  var labelCol = -1, totalCol = -1;
  var platformCols = []; // { col: index, name: "Local Indeed" }
  var tableStartRow = -1;

  // Find the header row by locating the "TOTAL" column
  for (var r = 0; r < data.length; r++) {
    var rowTotalCol = -1;
    for (var c = 0; c < data[r].length; c++) {
      var h = String(data[r][c] || '').toLowerCase().trim();
      if (h === 'total') { rowTotalCol = c; break; }
    }
    if (rowTotalCol < 0) continue;

    // Found TOTAL — now detect label col and all platform columns between them
    // Walk left from TOTAL to find the first non-empty column header as label boundary
    totalCol = rowTotalCol;
    var firstPlatformCol = -1;
    var tempPlatforms = [];

    for (var c2 = rowTotalCol - 1; c2 >= 0; c2--) {
      var val = String(data[r][c2] || '').trim();
      if (!val) {
        // Empty cell — this is the gap before the label column (or between label and platforms)
        if (tempPlatforms.length > 0) {
          labelCol = c2; // Label col is the last non-empty col before the gap, or col 0
          // Check if there's a label column to the left
          for (var c3 = c2; c3 >= 0; c3--) {
            var lv = String(data[r][c3] || '').trim();
            if (lv) { labelCol = c3; break; }
            if (c3 === 0) labelCol = 0;
          }
          break;
        }
        continue;
      }
      // Non-empty cell before TOTAL — it's a platform header
      tempPlatforms.unshift({ col: c2, name: val });
      firstPlatformCol = c2;
    }

    // If no empty gap found, label is the column before the first platform
    if (labelCol < 0 && tempPlatforms.length > 0) {
      labelCol = firstPlatformCol > 0 ? firstPlatformCol - 1 : 0;
    }

    if (tempPlatforms.length > 0 || totalCol >= 0) {
      platformCols = tempPlatforms;
      tableStartRow = r + 1;
      break;
    }
  }

  if (tableStartRow < 0) return null;
  if (labelCol < 0) labelCol = 0;

  // Row label patterns → keys
  var ROW_DEFS = [
    { patterns: ['# of ads', 'number of ads'],                          key: 'numAds' },
    { patterns: ['total cost'],                                          key: 'totalCost' },
    { patterns: ['# of applies', 'number of applies', 'applies'],       key: 'numApplies' },
    { patterns: ['cost/apply', 'cost per apply', 'cost / apply'],        key: 'costPerApply' },
    { patterns: ['# of 2nd', 'number of 2nd', '2nds'],                  key: 'num2nds' },
    { patterns: ['cost/second', 'cost/2nd', 'cost per 2nd', 'cost per second'], key: 'costPer2nd' },
    { patterns: ['# of new start', 'number of new start', 'new starts'], key: 'numNewStarts' },
    { patterns: ['cost/new start', 'cost per new start'],                key: 'costPerNewStart' }
  ];

  // Build per-platform data objects
  var platforms = {};
  var platformOrder = [];
  for (var pi = 0; pi < platformCols.length; pi++) {
    platforms[platformCols[pi].name] = {};
    platformOrder.push(platformCols[pi].name);
  }
  var totalData = {};

  for (var r = tableStartRow; r < Math.min(tableStartRow + 15, data.length); r++) {
    var label = String(data[r][labelCol] || '').toLowerCase().trim();
    if (!label) continue;

    var matchedKey = null;
    for (var d = 0; d < ROW_DEFS.length; d++) {
      for (var p = 0; p < ROW_DEFS[d].patterns.length; p++) {
        if (label.indexOf(ROW_DEFS[d].patterns[p]) >= 0) {
          matchedKey = ROW_DEFS[d].key;
          break;
        }
      }
      if (matchedKey) break;
    }
    if (!matchedKey) continue;

    // Read each platform column
    for (var pi = 0; pi < platformCols.length; pi++) {
      platforms[platformCols[pi].name][matchedKey] = _numVal(data[r][platformCols[pi].col]);
    }
    if (totalCol >= 0) totalData[matchedKey] = _numVal(data[r][totalCol]);
  }

  // Backward compat: also set local/nlr/total if those platforms exist
  var result = { platforms: platforms, total: totalData, platformOrder: platformOrder };
  for (var pi = 0; pi < platformOrder.length; pi++) {
    var lc = platformOrder[pi].toLowerCase();
    if (lc.indexOf('local') >= 0 && lc.indexOf('indeed') >= 0) result.local = platforms[platformOrder[pi]];
    if (lc.indexOf('nlr') >= 0 && lc.indexOf('indeed') >= 0) result.nlr = platforms[platformOrder[pi]];
  }
  if (!result.local) result.local = {};
  if (!result.nlr) result.nlr = {};

  return result;
}

/** Convert cell value to number, stripping $ and , */
function _numVal(v) {
  if (typeof v === 'number') return v;
  if (!v) return 0;
  var s = String(v).replace(/[$,\s]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}


// ══════════════════════════════════════════════════
// INDEED TRACKING — Weekly Ad Spend (prototype)
// Reads "Indeed Tracking 2026" tab from per-owner sheets.
// For prototype: hardcoded mapping of owner → sheet ID.
// ══════════════════════════════════════════════════

// Prototype mapping: owner name → { sheetId, tab }
var INDEED_TRACKING_SHEETS = {
  'Jackie Leroy': {
    sheetId: '1FnekrOkKwTCNgzSQEf2-Fz5P0BH7hfeHefsLmYc7A_g',
    tab: 'Indeed Tracking 2026'
  }
};

/**
 * readIndeedTracking(ownerName)
 * Returns: { weeks: [...], accounts: [...], error? }
 *
 * Each week object:
 *   { weekOf, totalSpend, totalApplies, total2nds, totalNewStarts,
 *     cpa, cpns, numAds,
 *     ads: [ { account, postedBy, adTitle, location, datePosted,
 *              spend, applies, seconds, newStarts, cpa, cpns, plan, adBudget } ] }
 *
 * accounts: unique Indeed Account names found across all weeks
 */
function readIndeedTracking(ownerName) {
  ownerName = String(ownerName || '').trim();
  if (!ownerName) return { error: 'owner parameter is required' };

  // Case-insensitive lookup
  var config = INDEED_TRACKING_SHEETS[ownerName];
  if (!config) {
    var lc = ownerName.toLowerCase();
    for (var k in INDEED_TRACKING_SHEETS) {
      if (k.toLowerCase() === lc) { config = INDEED_TRACKING_SHEETS[k]; break; }
    }
  }
  if (!config) return { error: 'No Indeed Tracking sheet configured for ' + ownerName, weeks: [] };

  // ── Check cache first (5 min TTL) ──
  var cacheKey = 'indeedTracking_' + ownerName.replace(/\s+/g, '_');
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      Logger.log('readIndeedTracking: cache HIT for ' + ownerName);
      return JSON.parse(cached);
    } catch (e) { /* parse failed, fetch fresh */ }
  }

  try {
    var ss = SpreadsheetApp.openById(config.sheetId);
    var sheet = ss.getSheetByName(config.tab);
    if (!sheet) return { error: 'Tab "' + config.tab + '" not found', weeks: [] };

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 3) return { weeks: [] };
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    if (data.length < 3) return { weeks: [] };

    // ── Parse weekly blocks ──
    // Structure: Row with "WEEK OF MM/DD" → header row → data rows → blank → next week
    var weeks = [];
    var accountSet = {};
    var r = 0;

    while (r < data.length) {
      // Look for "WEEK OF" header
      var weekOf = _findWeekHeader(data, r);
      if (!weekOf) { r++; continue; }

      r++; // skip to header row
      if (r >= data.length) break;

      // Verify header row has expected columns
      var colMap = _mapWeekColumns(data[r]);
      if (!colMap) { r++; continue; }

      r++; // skip to first data row

      // Read data rows until blank row or next week header
      var ads = [];
      while (r < data.length) {
        // Check if this is a blank separator or next week header
        var cellA = String(data[r][0] || '').trim();
        var cellAny = _rowHasData(data[r], colMap);

        if (!cellAny && !cellA) { r++; break; } // blank row = end of block
        if (_findWeekHeader(data, r)) break;     // next week header

        // Skip rows that are FLAGGED/PAUSED in Plan column (status rows, not real ads)
        var plan = colMap.plan >= 0 ? String(data[r][colMap.plan] || '').trim() : '';
        var spend = colMap.totalSpend >= 0 ? _numVal(data[r][colMap.totalSpend]) : 0;
        var applies = colMap.applies >= 0 ? _numVal(data[r][colMap.applies]) : 0;
        var seconds = colMap.seconds >= 0 ? _numVal(data[r][colMap.seconds]) : 0;
        var newStarts = colMap.newStarts >= 0 ? _numVal(data[r][colMap.newStarts]) : 0;
        var account = colMap.account >= 0 ? String(data[r][colMap.account] || '').trim() : '';
        var postedBy = colMap.postedBy >= 0 ? String(data[r][colMap.postedBy] || '').trim() : '';
        var adTitle = colMap.adTitle >= 0 ? String(data[r][colMap.adTitle] || '').trim() : '';
        var datePosted = colMap.datePosted >= 0 ? _formatCellDate(data[r][colMap.datePosted]) : '';
        var location = colMap.location >= 0 ? String(data[r][colMap.location] || '').trim() : '';
        var adBudget = colMap.adBudget >= 0 ? String(data[r][colMap.adBudget] || '').trim() : '';

        // Skip entirely empty data rows or rows without a title
        if (!adTitle && !account && applies === 0 && spend === 0) { r++; continue; }

        // Track account names
        if (account) accountSet[account] = true;

        // Carry forward account name from preceding rows (merged cells in sheet)
        var effectiveAccount = account;
        if (!effectiveAccount && ads.length > 0) {
          effectiveAccount = ads[ads.length - 1].account || '';
        }

        var adCpa = applies > 0 ? spend / applies : 0;
        var adCpns = newStarts > 0 ? spend / newStarts : 0;

        ads.push({
          account: effectiveAccount,
          postedBy: postedBy,
          adTitle: adTitle,
          datePosted: datePosted,
          location: location,
          spend: spend,
          applies: applies,
          seconds: seconds,
          newStarts: newStarts,
          cpa: Math.round(adCpa * 100) / 100,
          cpns: Math.round(adCpns * 100) / 100,
          plan: plan,
          adBudget: adBudget
        });

        r++;
      }

      // Compute week totals
      var wkSpend = 0, wkApplies = 0, wk2nds = 0, wkNewStarts = 0;
      for (var a = 0; a < ads.length; a++) {
        wkSpend += ads[a].spend;
        wkApplies += ads[a].applies;
        wk2nds += ads[a].seconds;
        wkNewStarts += ads[a].newStarts;
      }

      var wkCpa = wkApplies > 0 ? wkSpend / wkApplies : 0;
      var wkCpns = wkNewStarts > 0 ? wkSpend / wkNewStarts : 0;

      weeks.push({
        weekOf: weekOf,
        totalSpend: Math.round(wkSpend * 100) / 100,
        totalApplies: wkApplies,
        total2nds: wk2nds,
        totalNewStarts: wkNewStarts,
        cpa: Math.round(wkCpa * 100) / 100,
        cpns: Math.round(wkCpns * 100) / 100,
        numAds: ads.length,
        ads: ads
      });
    }

    // Sort weeks chronologically (oldest first → newest last)
    weeks.sort(function(a, b) {
      return _parseWeekDate(a.weekOf).getTime() - _parseWeekDate(b.weekOf).getTime();
    });

    // ── Compute WoW deltas for effectiveness tracking ──
    for (var wi = 0; wi < weeks.length; wi++) {
      if (wi === 0) {
        weeks[wi].delta = null;
      } else {
        var prev = weeks[wi - 1];
        var curr = weeks[wi];
        weeks[wi].delta = {
          spend: Math.round((curr.totalSpend - prev.totalSpend) * 100) / 100,
          applies: curr.totalApplies - prev.totalApplies,
          seconds: curr.total2nds - prev.total2nds,
          newStarts: curr.totalNewStarts - prev.totalNewStarts,
          cpa: Math.round((curr.cpa - prev.cpa) * 100) / 100,
          cpns: Math.round((curr.cpns - prev.cpns) * 100) / 100,
          spendPct: prev.totalSpend > 0 ? Math.round(((curr.totalSpend - prev.totalSpend) / prev.totalSpend) * 1000) / 10 : null,
          appliesPct: prev.totalApplies > 0 ? Math.round(((curr.totalApplies - prev.totalApplies) / prev.totalApplies) * 1000) / 10 : null,
          cpaPct: prev.cpa > 0 ? Math.round(((curr.cpa - prev.cpa) / prev.cpa) * 1000) / 10 : null,
          cpnsPct: prev.cpns > 0 ? Math.round(((curr.cpns - prev.cpns) / prev.cpns) * 1000) / 10 : null
        };
      }
    }

    var accounts = Object.keys(accountSet).sort();
    Logger.log('readIndeedTracking: ' + weeks.length + ' weeks, ' + accounts.length + ' accounts for ' + ownerName);

    // ── Cache result for 5 minutes ──
    var result = { weeks: weeks, accounts: accounts };
    try {
      var json = JSON.stringify(result);
      if (json.length < 100000) { // CacheService limit is 100KB per key
        cache.put(cacheKey, json, 300); // 5 min
        Logger.log('readIndeedTracking: cached (' + json.length + ' bytes)');
      }
    } catch (e) { Logger.log('Cache write failed: ' + e.message); }

    return result;

  } catch (err) {
    Logger.log('readIndeedTracking error: ' + err.message + '\n' + err.stack);
    return { error: 'Failed to read Indeed Tracking: ' + err.message, weeks: [] };
  }
}


// ── Helper: find "WEEK OF MM/DD" in row ──
function _findWeekHeader(data, r) {
  for (var c = 0; c < Math.min(data[r].length, 12); c++) {
    var v = String(data[r][c] || '').trim().toUpperCase();
    if (v.indexOf('WEEK OF') === 0) {
      // Extract date portion: "WEEK OF 03/09" → "03/09"
      return v.replace('WEEK OF', '').trim();
    }
  }
  return null;
}


// ── Helper: map column headers to indices ──
function _mapWeekColumns(row) {
  var map = {
    account: -1, postedBy: -1, adTitle: -1, datePosted: -1,
    location: -1, totalSpend: -1, applies: -1, seconds: -1,
    newStarts: -1, cpa: -1, cpns: -1, plan: -1, adBudget: -1
  };

  var found = 0;
  for (var c = 0; c < row.length; c++) {
    var h = String(row[c] || '').toLowerCase().trim().replace(/[:\s]+$/, '');
    if (h === 'indeed account')                          { map.account = c; found++; }
    else if (h === 'posted by')                          { map.postedBy = c; found++; }
    else if (h === 'ad title')                           { map.adTitle = c; found++; }
    else if (h === 'date posted')                        { map.datePosted = c; found++; }
    else if (h === 'current location' || h === 'location') { map.location = c; found++; }
    else if (h === 'total spend')                        { map.totalSpend = c; found++; }
    else if (h === 'applies')                            { map.applies = c; found++; }
    else if (h === '2nds')                               { map.seconds = c; found++; }
    else if (h === 'new starts')                         { map.newStarts = c; found++; }
    else if (h === 'cpa')                                { map.cpa = c; found++; }
    else if (h === 'cpns')                               { map.cpns = c; found++; }
    else if (h === 'plan')                               { map.plan = c; found++; }
    else if (h === 'ad budget')                          { map.adBudget = c; found++; }
  }

  // Need at least a couple key columns to consider valid
  return found >= 3 ? map : null;
}


// ── Helper: check if a data row has any useful data ──
function _rowHasData(row, colMap) {
  // Check key data columns for non-empty values
  var cols = [colMap.adTitle, colMap.applies, colMap.totalSpend, colMap.newStarts, colMap.account, colMap.plan];
  for (var i = 0; i < cols.length; i++) {
    if (cols[i] >= 0 && cols[i] < row.length) {
      var v = row[cols[i]];
      if (v !== null && v !== undefined && String(v).trim() !== '') return true;
    }
  }
  return false;
}


// ── Helper: parse "MM/DD" week date string into a Date for sorting ──
function _parseWeekDate(dateStr) {
  // dateStr like "03/09" or "01/26"
  var parts = String(dateStr).split('/');
  if (parts.length < 2) return new Date(0);
  var month = parseInt(parts[0], 10) - 1;
  var day = parseInt(parts[1], 10);
  var year = new Date().getFullYear(); // assume current year
  return new Date(year, month, day);
}


// ── Helper: format Date objects or date strings from cells ──
function _formatCellDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    var m = v.getMonth() + 1;
    var d = v.getDate();
    return (m < 10 ? '0' : '') + m + '/' + (d < 10 ? '0' : '') + d;
  }
  return String(v).trim();
}

// ══════════════════════════════════════════════════
// READ NDS HEADCOUNT from external NDS One-on-Ones sheet
// Reads specified owner tabs (or all owner tabs) and returns
// headcount/production data in the same shape as readNLRHeadcount().
// ══════════════════════════════════════════════════

// NDS data sources: One-on-Ones sheet + NLR NDS report
// Each entry: { sheetId, tabs } — tabs is array of tab names to read
var NDS_SOURCES = [
  { sheetId: 'NDS_ONE_ON_ONES', tabs: ['Sam Poles'] },
  { sheetId: 'NLR_NDS',         tabs: null }  // null = read ALL owner tabs (skip known non-owner tabs)
];

var NDS_SKIP_TABS = {
  'input - sales qual metrics': true,
  'input - comm per rep and total dd': true,
  'input - market metrics': true,
  'sheet1': true, 'sheet2': true, 'sheet3': true
};

function readNDSHeadcount() {
  var owners = {};

  for (var s = 0; s < NDS_SOURCES.length; s++) {
    var src = NDS_SOURCES[s];
    var id = SHEETS[src.sheetId];
    if (!id) continue;

    var ss;
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (err) {
      Logger.log('Cannot open ' + src.sheetId + ': ' + err.message);
      continue;
    }

    // Determine which tabs to read
    var tabsToRead = src.tabs;
    if (!tabsToRead) {
      // Read all tabs, skip known non-owner tabs
      tabsToRead = ss.getSheets().map(function(sh) { return sh.getName(); })
        .filter(function(name) { return !NDS_SKIP_TABS[name.toLowerCase().trim()]; });
    }

    for (var t = 0; t < tabsToRead.length; t++) {
      var tabName = tabsToRead[t];
      var sheet = ss.getSheetByName(tabName);
      if (!sheet) { Logger.log('NDS tab not found: ' + tabName + ' in ' + src.sheetId); continue; }

    // Read top ~50 rows (office health section is at top, recruiting starts around row 48)
    var lastRow = Math.min(sheet.getLastRow(), 50);
    if (lastRow < 2) continue;
    var lastCol = Math.min(sheet.getLastColumn(), 18); // cols A-R
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    // Row 0 = headers — use header-based column mapping
    var headers = data[0].map(function(h) { return String(h).toLowerCase().trim(); });

    var colMap = {
      dates:          findCol(headers, ['dates', 'date']),
      active:         findCol(headers, ['active']),
      leaders:        findCol(headers, ['leaders']),
      dist:           findCol(headers, ['dist']),
      training:       findCol(headers, ['training']),
      productionLW:   findCol(headers, ['production lw']),
      productionGoals:findCol(headers, ['production goals'])
    };

    if (colMap.dates < 0 && colMap.active < 0) continue;

    // Also map recruiting funnel columns (same row, cols H onward)
    var rcMap = {
      firstRoundsBooked:  findCol(headers, ['1st rounds booked', '1st round booked']),
      firstRoundsShowed:  findCol(headers, ['1st rounds showed', '1st round showed']),
      turnedTo2nd:        findCol(headers, ['turned to 2nd']),
      retention1:         _findNthCol(headers, ['retention'], 1),
      conversion:         findCol(headers, ['conversion']),
      secondRoundsBooked: findCol(headers, ['2nd rounds booked', '2nd round booked']),
      secondRoundsShowed: findCol(headers, ['2nd rounds showed', '2nd round showed']),
      retention2:         _findNthCol(headers, ['retention'], 2),
      newStartScheduled:  findCol(headers, ['new start scheduled', 'new starts scheduled']),
      newStartsShowed:    findCol(headers, ['new starts showed']),
      retention3:         _findNthCol(headers, ['retention'], 3)
    };

    var trend = [];
    var lastGood = null;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      // Date gate
      var dateCell = colMap.dates >= 0 ? row[colMap.dates] : null;
      if (!dateCell) continue;
      if (!(dateCell instanceof Date)) {
        var parsed = new Date(dateCell);
        if (isNaN(parsed.getTime())) continue;
      }

      var hasActive = colMap.active >= 0 && row[colMap.active] !== '' && row[colMap.active] !== null;
      var hasProd   = colMap.productionLW >= 0 && row[colMap.productionLW] !== '' && row[colMap.productionLW] !== null;
      if (!hasActive && !hasProd) continue;

      var entry = {
        date:            colMap.dates >= 0 ? formatDate(row[colMap.dates]) : '',
        active:          colMap.active >= 0 ? num(row[colMap.active]) : 0,
        leaders:         colMap.leaders >= 0 ? num(row[colMap.leaders]) : 0,
        dist:            colMap.dist >= 0 ? num(row[colMap.dist]) : 0,
        training:        colMap.training >= 0 ? num(row[colMap.training]) : 0,
        productionLW:    colMap.productionLW >= 0 ? num(row[colMap.productionLW]) : 0,
        productionGoals: colMap.productionGoals >= 0 ? num(row[colMap.productionGoals]) : 0,
        // Recruiting funnel
        firstRoundsBooked:  rcMap.firstRoundsBooked >= 0 ? num(row[rcMap.firstRoundsBooked]) : 0,
        firstRoundsShowed:  rcMap.firstRoundsShowed >= 0 ? num(row[rcMap.firstRoundsShowed]) : 0,
        turnedTo2nd:        rcMap.turnedTo2nd >= 0 ? num(row[rcMap.turnedTo2nd]) : 0,
        retention1:         rcMap.retention1 >= 0 ? numDec(row[rcMap.retention1]) : 0,
        conversion:         rcMap.conversion >= 0 ? numDec(row[rcMap.conversion]) : 0,
        secondRoundsBooked: rcMap.secondRoundsBooked >= 0 ? num(row[rcMap.secondRoundsBooked]) : 0,
        secondRoundsShowed: rcMap.secondRoundsShowed >= 0 ? num(row[rcMap.secondRoundsShowed]) : 0,
        retention2:         rcMap.retention2 >= 0 ? numDec(row[rcMap.retention2]) : 0,
        newStartScheduled:  rcMap.newStartScheduled >= 0 ? num(row[rcMap.newStartScheduled]) : 0,
        newStartsShowed:    rcMap.newStartsShowed >= 0 ? num(row[rcMap.newStartsShowed]) : 0,
        retention3:         rcMap.retention3 >= 0 ? numDec(row[rcMap.retention3]) : 0
      };
      trend.push(entry);
      lastGood = entry;
    }

    if (lastGood) {
      owners[tabName] = { current: lastGood, trend: trend };
    }
    } // end tabs loop
  } // end sources loop

  return { owners: owners };
}

// ── Helper: find Nth occurrence of a header (for repeated "Retention" columns) ──
function _findNthCol(headers, patterns, n) {
  var count = 0;
  for (var i = 0; i < headers.length; i++) {
    for (var p = 0; p < patterns.length; p++) {
      if (headers[i] === patterns[p]) {
        count++;
        if (count === n) return i;
      }
    }
  }
  return -1;
}

// ══════════════════════════════════════════════════
// IMPORT NDS HEADCOUNT → LOCAL _NDS_Headcount TAB
// Reads from NDS One-on-Ones sheet, writes into Ken's national sheet.
// Tab format: Owner | Date | Active | Leaders | Dist | Training | ProductionLW | ProductionGoals |
//             1stBooked | 1stShowed | TurnedTo2nd | Ret1 | Conversion |
//             2ndBooked | 2ndShowed | Ret2 | NewStartSched | NewStartsShowed | Ret3
// ══════════════════════════════════════════════════

function importNDSHeadcount() {
  // 1. Read from NDS spreadsheet
  var ndsResult = readNDSHeadcount();
  if (ndsResult.error) return ndsResult;
  var ndsOwners = ndsResult.owners || {};
  if (!Object.keys(ndsOwners).length) return { error: 'No owner data found in NDS sheet' };

  // 2. Open Ken's national sheet
  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message };
  }

  // 3. Get or create _NDS_Headcount tab
  var TAB_NAME = '_NDS_Headcount';
  var HEADERS = [
    'Owner', 'Date', 'Active', 'Leaders', 'Dist', 'Training', 'ProductionLW', 'ProductionGoals',
    '1stBooked', '1stShowed', 'TurnedTo2nd', 'Ret1', 'Conversion',
    '2ndBooked', '2ndShowed', 'Ret2', 'NewStartSched', 'NewStartsShowed', 'Ret3'
  ];
  var sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_NAME);
    sheet.appendRow(HEADERS);
  } else {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, HEADERS.length).clearContent();
    }
  }

  // Name aliases: tab name → canonical owner name (for matching with recruiting data)
  var NAME_ALIASES = {
    'Sam Poles': 'Samih Poles'
  };

  // 4. Build rows
  var rows = [];
  var ownerNames = Object.keys(ndsOwners).sort();
  for (var o = 0; o < ownerNames.length; o++) {
    var name = NAME_ALIASES[ownerNames[o]] || ownerNames[o];
    var ownerData = ndsOwners[name];
    var trend = ownerData.trend || [];
    for (var t = 0; t < trend.length; t++) {
      var e = trend[t];
      rows.push([
        name, e.date, e.active, e.leaders, e.dist, e.training, e.productionLW, e.productionGoals,
        e.firstRoundsBooked, e.firstRoundsShowed, e.turnedTo2nd, e.retention1, e.conversion,
        e.secondRoundsBooked, e.secondRoundsShowed, e.retention2, e.newStartScheduled, e.newStartsShowed, e.retention3
      ]);
    }
  }

  // 5. Write all rows
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  }

  return { success: true, ownersImported: ownerNames.length, rowsWritten: rows.length };
}

// ══════════════════════════════════════════════════
// READ LOCAL NDS HEADCOUNT from _NDS_Headcount tab
// Returns same shape as readNLRHeadcount():
// { owners: { "Name": { current: {...}, trend: [{...}] } } }
// ══════════════════════════════════════════════════

function readLocalNDSHeadcount() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message, owners: {} };
  }

  var sheet = ss.getSheetByName('_NDS_Headcount');
  if (!sheet) return { owners: {} };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { owners: {} };

  var owners = {};
  for (var i = 1; i < data.length; i++) {
    var oName = String(data[i][0] || '').trim();
    if (!oName) continue;

    var entry = {
      date:            formatDate(data[i][1]),
      active:          num(data[i][2]),
      leaders:         num(data[i][3]),
      dist:            num(data[i][4]),
      training:        num(data[i][5]),
      productionLW:    num(data[i][6]),
      productionGoals: num(data[i][7]),
      firstRoundsBooked:  num(data[i][8]),
      firstRoundsShowed:  num(data[i][9]),
      turnedTo2nd:        num(data[i][10]),
      retention1:         numDec(data[i][11]),
      conversion:         numDec(data[i][12]),
      secondRoundsBooked: num(data[i][13]),
      secondRoundsShowed: num(data[i][14]),
      retention2:         numDec(data[i][15]),
      newStartScheduled:  num(data[i][16]),
      newStartsShowed:    num(data[i][17]),
      retention3:         numDec(data[i][18])
    };

    if (!owners[oName]) owners[oName] = { current: null, trend: [] };
    owners[oName].trend.push(entry);
    owners[oName].current = entry; // last row wins
  }

  return { owners: owners };
}

