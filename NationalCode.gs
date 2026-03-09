// ═══════════════════════════════════════════════════════
// NationalCode.gs — National Consultant Dashboard Data Aggregator
// Reads from 4 external Google Sheets, returns unified JSON.
// Deploy as web app (Execute as: Me, Access: Anyone)
// ═══════════════════════════════════════════════════════

// ── Config ──
const NC_API_KEY = 'national-dash-2026-secret'; // Must match Script Properties > API_KEY

// External sheet IDs
const SHEETS = {
  RECRUITING_WEEKLY:  '1MNLqi8A329444SeZpKbYbcRe3dMxaOPLVdMy-7F1DPk',  // All Campaigns Stats Tracker 2026
  RECRUITING_DAILY:   '1ytTGen_AlzfDPW3HGYU1JKNLz1kfHrrhAFCVnmRS3fg',  // Recruiting Scoreboard Daily
  CAMPAIGN_TRACKER:   '1HvWJYox3JXvxmza63YBWAqKPtUGPFuaV-s-BOfbWGKM',  // ATT Campaign Tracker
  PERFORMANCE_AUDIT:  '15WCMzKnqvyyRMx2ae4tC1a12_-aoSRDh3McOuRAKuHk',  // Performance Audit
  NATIONAL:           '1eGkwjQRD9RV4n-JR_TTlgE6VY858WZID8cAF8soYSYM', // Ken's national recruiting sheet
  NLR_B2B:            '1sxauFjNjq4_rRYM2PAl5cyOyHF3Hg4OkO-t_hLKDJB8'  // NLR's AT&T B2B 1-on-1's report
};

// Campaign configs
const CAMPAIGNS = {
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

    if (owner) {
      // Single owner detail request
      const data = loadOwnerDetail(campaign, owner);
      return jsonResp(data);
    } else {
      // Full campaign overview
      const data = loadCampaignOverview(campaign);
      return jsonResp(data);
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
      case 'importNLRHeadcount':
        result = importNLRHeadcount();
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

  // 6. Return success + fresh data
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

  return { campaigns: campaigns };
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
// OWNER → CAM COMPANY MAPPING
// Reads '_OwnerCamMapping' tab from Ken's national sheet.
// Two columns: Owner Name | Cam Company Name
// Returns { mapping: { "Owner Name": ["Company A", "Company B"], ... } }
// ══════════════════════════════════════════════════

function readOwnerCamMapping() {
  var ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var sheet = ss.getSheetByName('_OwnerCamMapping');
  if (!sheet) return { mapping: {} };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { mapping: {} };

  // Expect headers: Owner Name | Cam Company Name
  var mapping = {};
  for (var i = 1; i < data.length; i++) {
    var ownerName = String(data[i][0] || '').trim();
    var camCompany = String(data[i][1] || '').trim();
    if (!ownerName || !camCompany) continue;
    if (!mapping[ownerName]) mapping[ownerName] = [];
    mapping[ownerName].push(camCompany);
  }

  return { mapping: mapping };
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
    sheet.getRange(1, 1, 1, 2).setValues([['Owner Name', 'Cam Company Name']]);
  }

  // Check for duplicate
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === ownerName.toLowerCase() &&
        String(data[i][1]).trim().toLowerCase() === companyName.toLowerCase()) {
      return { ok: true, message: 'Already claimed', mapping: readOwnerCamMapping().mapping };
    }
  }

  // Append new row
  sheet.appendRow([ownerName, companyName]);

  return { ok: true, mapping: readOwnerCamMapping().mapping };
}


// ── Unclaim a company from an owner (delete from _OwnerCamMapping) ──
function unclaimCompany(ownerName, companyName) {
  ownerName = String(ownerName || '').trim();
  companyName = String(companyName || '').trim();
  if (!ownerName || !companyName) return { error: 'ownerName and companyName are required' };

  var ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var sheet = ss.getSheetByName('_OwnerCamMapping');
  if (!sheet) return { ok: true, mapping: {} };

  var data = sheet.getDataRange().getValues();
  // Search from bottom to top so row indices don't shift when deleting
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim().toLowerCase() === ownerName.toLowerCase() &&
        String(data[i][1]).trim().toLowerCase() === companyName.toLowerCase()) {
      sheet.deleteRow(i + 1); // sheet rows are 1-indexed
    }
  }

  return { ok: true, mapping: readOwnerCamMapping().mapping };
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
      date:            String(row[1] || ''),
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
  return String(v);
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
