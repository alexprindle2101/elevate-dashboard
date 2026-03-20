// ============================================================
//  TableauSyncRes.gs — AT&T Residential D2D Tableau → National Sheet
//  Pulls PRODUCT SALES SUMMARY 4WK, aggregates by ICD Owner, ranks by total sales.
//  Writes compact ranking to _D2D_Res_Ranking tab in the National sheet.
// ============================================================
//
//  SETUP (one-time):
//    1. Create a new standalone Apps Script project
//    2. Paste this file
//    3. Set Script Properties (File → Project settings → Script properties):
//       - TABLEAU_EMAIL        → your Tableau login email
//       - TABLEAU_PASSWORD     → your Tableau login password
//       - TABLEAU_SITE         → "sci"
//       - TABLEAU_SERVER       → "https://us-east-1.online.tableau.com"
//       - SHEET_ID             → National sheet: 1eGkwjQRD9RV4n-JR_TTlgE6VY858WZID8cAF8soYSYM
//    4. Run syncD2DRes() manually to pull initial data
//    5. Run setupDailyTrigger() once to schedule automatic syncs
// ============================================================

var TABLEAU_API_VERSION = '3.24';

// === CONFIGURATION ===

function _getConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    email:    props.getProperty('TABLEAU_EMAIL')    || '',
    password: props.getProperty('TABLEAU_PASSWORD') || '',
    site:     props.getProperty('TABLEAU_SITE')     || 'sci',
    server:   props.getProperty('TABLEAU_SERVER')   || 'https://us-east-1.online.tableau.com',
    sheetId:  props.getProperty('SHEET_ID')         || ''
  };
}

// === REPORT CONFIG ===

var REPORT = {
  // Dashboard view (used to discover the workbook ID)
  dashboardContentUrl: 'ATTTracker2_1-D2DV2/D2D1-PAGERV3',
  // Target sheet: pre-aggregated product sales summary (rep-level, we roll up to owner)
  targetSheet: 'PRODUCT SALES SUMMARY 4WK',
  // Output tab in the National sheet
  tabName: '_D2D_Res_Ranking'
};


// === TABLEAU REST API ===

function _tableauSignIn(config) {
  var url = config.server + '/api/' + TABLEAU_API_VERSION + '/auth/signin';
  var payload = {
    credentials: {
      name: config.email,
      password: config.password,
      site: { contentUrl: config.site }
    }
  };

  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Accept': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error('Tableau sign-in failed (HTTP ' + resp.getResponseCode() + '): ' + resp.getContentText().substring(0, 300));
  }

  var json = JSON.parse(resp.getContentText());
  return { token: json.credentials.token, siteId: json.credentials.site.id };
}

function _tableauSignOut(config, token) {
  try {
    UrlFetchApp.fetch(config.server + '/api/' + TABLEAU_API_VERSION + '/auth/signout', {
      method: 'post',
      headers: { 'X-Tableau-Auth': token },
      muteHttpExceptions: true
    });
  } catch (e) { /* non-critical */ }
}

function _findView(config, token, siteId, viewContentUrl) {
  var parts = viewContentUrl.split('/');
  var viewName = parts[parts.length - 1];

  var url = config.server + '/api/' + TABLEAU_API_VERSION
    + '/sites/' + siteId + '/views'
    + '?filter=viewUrlName:eq:' + encodeURIComponent(viewName)
    + '&pageSize=100';

  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-Tableau-Auth': token, 'Accept': 'application/json' },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error('Failed to query views (HTTP ' + resp.getResponseCode() + '): ' + resp.getContentText().substring(0, 300));
  }

  var json = JSON.parse(resp.getContentText());
  var views = json.views && json.views.view ? json.views.view : [];

  // Match by full content URL
  for (var i = 0; i < views.length; i++) {
    if (views[i].contentUrl === viewContentUrl) {
      return { viewId: views[i].id, workbookId: views[i].workbook ? views[i].workbook.id : null };
    }
  }

  // Partial match
  var workbookName = parts[0];
  for (var i = 0; i < views.length; i++) {
    if (views[i].contentUrl && views[i].contentUrl.indexOf(workbookName) !== -1) {
      return { viewId: views[i].id, workbookId: views[i].workbook ? views[i].workbook.id : null };
    }
  }

  if (views.length === 1) {
    return { viewId: views[0].id, workbookId: views[0].workbook ? views[0].workbook.id : null };
  }

  throw new Error('View not found: ' + viewContentUrl);
}

function _downloadViewData(config, token, siteId, viewId) {
  var url = config.server + '/api/' + TABLEAU_API_VERSION
    + '/sites/' + siteId + '/views/' + viewId + '/data';

  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-Tableau-Auth': token },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error('Failed to download view data (HTTP ' + resp.getResponseCode() + '): ' + resp.getContentText().substring(0, 200));
  }

  return resp.getContentText();
}


// === SHEET WRITING ===

function _writeToSheet(sheetId, tabName, headers, rows) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) sheet = ss.insertSheet(tabName);

  sheet.clearContents();

  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  for (var i = 1; i <= Math.min(headers.length, 25); i++) {
    sheet.autoResizeColumn(i);
  }

  return { tab: tabName, headers: headers.length, rows: rows.length };
}


// === COLUMN FINDER HELPERS ===

/**
 * Find a column index by trying exact matches first, then case-insensitive partial.
 * Returns -1 if not found.
 */
function _findCol(headers, exactPatterns, fuzzyKeywords) {
  // Exact match
  for (var p = 0; p < exactPatterns.length; p++) {
    var idx = headers.indexOf(exactPatterns[p]);
    if (idx !== -1) return idx;
  }
  // Case-insensitive partial
  if (fuzzyKeywords) {
    for (var h = 0; h < headers.length; h++) {
      var hLower = headers[h].toLowerCase();
      var allMatch = true;
      for (var k = 0; k < fuzzyKeywords.length; k++) {
        if (hLower.indexOf(fuzzyKeywords[k]) === -1) { allMatch = false; break; }
      }
      if (allMatch) return h;
    }
  }
  return -1;
}

/**
 * Parse a cell value as a number (handles commas, %, empty strings)
 */
function _num(val) {
  if (!val && val !== 0) return 0;
  var s = String(val).replace(/[,%$]/g, '').trim();
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}


// === MAIN SYNC ===

/**
 * Pull PRODUCT SALES SUMMARY 4WK from Tableau, aggregate rep-level data
 * up to ICD Owner level, rank by total sales, and write to _D2D_Res_Ranking.
 *
 * The view has rep-level rows. Multiple reps belong to one owner.
 * We sum Sales (All) per owner, plus per-product breakdowns.
 */
function syncD2DRes() {
  var config = _getConfig();
  if (!config.email || !config.password) throw new Error('Tableau credentials not set');
  if (!config.sheetId) throw new Error('SHEET_ID not set');

  var auth = _tableauSignIn(config);

  try {
    // ── Find the target sheet view in the workbook ──
    Logger.log('Finding dashboard: ' + REPORT.dashboardContentUrl);
    var dashInfo = _findView(config, auth.token, auth.siteId, REPORT.dashboardContentUrl);
    Logger.log('Dashboard: viewId=' + dashInfo.viewId + ' workbookId=' + dashInfo.workbookId);

    if (!dashInfo.workbookId) throw new Error('Could not get workbook ID');

    // List all views in the workbook
    var wbUrl = config.server + '/api/' + TABLEAU_API_VERSION
      + '/sites/' + auth.siteId + '/workbooks/' + dashInfo.workbookId + '/views';
    var wbResp = UrlFetchApp.fetch(wbUrl, {
      method: 'get',
      headers: { 'X-Tableau-Auth': auth.token, 'Accept': 'application/json' },
      muteHttpExceptions: true
    });
    if (wbResp.getResponseCode() !== 200) {
      throw new Error('Failed to list workbook views: ' + wbResp.getContentText().substring(0, 300));
    }

    var views = JSON.parse(wbResp.getContentText());
    var viewList = views.views && views.views.view ? views.views.view : [];

    // Find the target sheet
    var targetViewId = null;
    for (var i = 0; i < viewList.length; i++) {
      if (viewList[i].name === REPORT.targetSheet) {
        targetViewId = viewList[i].id;
        break;
      }
    }
    if (!targetViewId) {
      // Case-insensitive fallback
      var targetLower = REPORT.targetSheet.toLowerCase();
      for (var i = 0; i < viewList.length; i++) {
        if (viewList[i].name.toLowerCase() === targetLower) {
          targetViewId = viewList[i].id;
          break;
        }
      }
    }
    if (!targetViewId) {
      throw new Error('Sheet "' + REPORT.targetSheet + '" not found. Available: '
        + viewList.map(function(v) { return '"' + v.name + '"'; }).join(', '));
    }

    Logger.log('Found target: "' + REPORT.targetSheet + '" → ' + targetViewId);

    // ── Download the data ──
    var csv = _downloadViewData(config, auth.token, auth.siteId, targetViewId);
    Logger.log('Downloaded: ' + csv.length + ' chars');

    var data = Utilities.parseCsv(csv);
    Logger.log('Parsed: ' + data.length + ' rows (incl header)');
    if (data.length < 2) throw new Error('No data rows returned');

    var headers = data[0];

    // ── Log all headers for debugging ──
    Logger.log('=== HEADERS (' + headers.length + ') ===');
    for (var h = 0; h < headers.length; h++) {
      Logger.log('  [' + h + '] "' + headers[h] + '"');
    }

    // ── Detect key columns ──
    // Owner column — the view is rep-level, but we need to find which column has the owner
    // Could be: "ICD Owner Name", or we may need to look at the data
    var ownerIdx = _findCol(headers,
      ['ICD Owner Name', 'Owner Name', 'Owner'],
      ['owner']);

    // Rep column
    var repIdx = _findCol(headers,
      ['Rep Name', 'Rep'],
      ['rep', 'name']);

    // Total sales column
    var salesIdx = _findCol(headers,
      ['Sales (All)', 'Sales (All) (Metrics)', 'Sales All', 'Total'],
      ['sales', 'all']);

    // Individual product columns
    var newInternetIdx = _findCol(headers,
      ['New Internet Count (Metrics)', 'New Internet Count'],
      ['new', 'internet', 'count']);
    var upgradeInternetIdx = _findCol(headers,
      ['Upgrade Internet Count (Metrics)', 'Upgrade Internet Count'],
      ['upgrade', 'internet', 'count']);
    var videoIdx = _findCol(headers,
      ['Video Sales (Metrics)', 'Video Sales'],
      ['video', 'sales']);

    Logger.log('Columns — owner:' + ownerIdx + ' rep:' + repIdx + ' sales:' + salesIdx
      + ' newInt:' + newInternetIdx + ' upgInt:' + upgradeInternetIdx + ' video:' + videoIdx);

    // If no owner column, we'll need to figure out an alternative
    // The view might only have Rep Name — log a warning
    if (ownerIdx === -1) {
      Logger.log('WARNING: No owner column found. Will try to use rep column or first text column.');
      // If the view doesn't have an owner column, maybe the rows ARE owner-level
      // (the "Rep Name" in the screenshot might actually be the owner/ICD name)
      if (repIdx !== -1) {
        ownerIdx = repIdx;
        Logger.log('Using rep column as owner column');
      } else {
        // Use first text column
        ownerIdx = 0;
        Logger.log('Falling back to column 0: "' + headers[0] + '"');
      }
    }

    if (salesIdx === -1) {
      Logger.log('WARNING: No sales total column found — will count rows per owner');
    }

    // ── Log sample data ──
    var sampleCount = Math.min(data.length - 1, 5);
    for (var s = 0; s < sampleCount; s++) {
      var row = data[s + 1];
      Logger.log('Sample [' + s + ']: owner="' + (row[ownerIdx]||'') + '" sales="' + (salesIdx >= 0 ? row[salesIdx] : 'N/A') + '"');
    }

    // ── Aggregate to owner level ──
    // Sum sales and product counts per owner
    var ownerData = {};

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var owner = String(row[ownerIdx] || '').trim();
      if (!owner || owner.toLowerCase() === 'total' || owner.toLowerCase() === 'grand total') continue;

      var ownerKey = owner.toUpperCase();
      if (!ownerData[ownerKey]) {
        ownerData[ownerKey] = {
          name: owner,
          totalSales: 0,
          newInternet: 0,
          upgradeInternet: 0,
          video: 0,
          repCount: 0
        };
      }

      ownerData[ownerKey].repCount++;

      if (salesIdx !== -1) {
        ownerData[ownerKey].totalSales += _num(row[salesIdx]);
      } else {
        ownerData[ownerKey].totalSales++; // count rows if no sales column
      }

      if (newInternetIdx !== -1) ownerData[ownerKey].newInternet += _num(row[newInternetIdx]);
      if (upgradeInternetIdx !== -1) ownerData[ownerKey].upgradeInternet += _num(row[upgradeInternetIdx]);
      if (videoIdx !== -1) ownerData[ownerKey].video += _num(row[videoIdx]);
    }

    // ── Rank by total sales descending ──
    var ownerList = Object.keys(ownerData).map(function(k) { return ownerData[k]; });
    ownerList.sort(function(a, b) { return b.totalSales - a.totalSales; });

    Logger.log('=== OWNER RANKING (' + ownerList.length + ' owners) ===');

    var rankHeaders = ['Rank', 'Owner', 'Total Sales', 'New Internet', 'Upgrade Internet', 'Video', 'Reps'];
    var rankRows = [];

    for (var i = 0; i < ownerList.length; i++) {
      var o = ownerList[i];
      rankRows.push([
        i + 1,
        o.name,
        o.totalSales,
        o.newInternet,
        o.upgradeInternet,
        o.video,
        o.repCount
      ]);
      if (i < 15) {
        Logger.log('  #' + (i+1) + ' ' + o.name + ' — ' + o.totalSales + ' total'
          + ' (NI:' + o.newInternet + ' UI:' + o.upgradeInternet + ' V:' + o.video + ')');
      }
    }
    Logger.log('=== END RANKING ===');

    // ── Write to sheet ──
    var result = _writeToSheet(config.sheetId, REPORT.tabName, rankHeaders, rankRows);
    Logger.log('Wrote: ' + JSON.stringify(result));

    // Sync metadata
    var ss = SpreadsheetApp.openById(config.sheetId);
    var sheet = ss.getSheetByName(REPORT.tabName);
    var metaRow = rankRows.length + 3;
    sheet.getRange(metaRow, 1, 1, 4).setValues([[
      'Last Sync', new Date().toISOString(), 'Source: ' + REPORT.targetSheet, 'Raw Rows: ' + (data.length - 1)
    ]]);
    sheet.getRange(metaRow, 1, 1, 4).setFontStyle('italic').setFontColor('#999999');

    return {
      ok: true,
      owners: ownerList.length,
      topOwner: ownerList.length > 0 ? '#1 ' + ownerList[0].name + ' (' + ownerList[0].totalSales + ')' : 'none'
    };

  } finally {
    _tableauSignOut(config, auth.token);
  }
}


// === TRIGGER SETUP ===

function setupDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'syncD2DRes') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('syncD2DRes')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  Logger.log('Daily trigger set: syncD2DRes at ~6 AM');
}
