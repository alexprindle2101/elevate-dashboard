// ============================================================
//  TableauSync.gs — Standalone Tableau → Google Sheets Pipeline
//  Central project for all campaign report syncs
// ============================================================
//
//  SETUP (one-time):
//    1. Create a new standalone Apps Script project
//    2. Paste this file
//    3. Set Script Properties (File → Project settings → Script properties):
//       - TABLEAU_EMAIL        → your Tableau login email
//       - TABLEAU_PASSWORD     → your Tableau login password
//       - TABLEAU_SITE         → site content URL (e.g. "sci")
//       - TABLEAU_SERVER       → server URL (e.g. "https://us-east-1.online.tableau.com")
//       - SHEET_ID             → Google Sheet ID for Tableau data output
//    4. Run setupDailyTrigger() once from the editor
//    5. Done — runs automatically every morning
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

// === REPORT REGISTRY ===
// Add new reports here. Each entry defines:
//   - viewContentUrl: workbook/view path as it appears in Tableau URLs
//   - tabName: the Google Sheet tab to write data to
//   - columns: array of column names to keep (case-sensitive, must match Tableau field names)
//   - dateFilterField: (optional) Tableau field name for date filtering
//   - dateRangeDays: (optional) rolling window in days (default 30)

var REPORTS = {

  'b2b-order-log': {
    viewContentUrl: 'ATTTRACKER-B2B/sheets/ORDERLOG',
    customViewId: 'c9cd7be0-1aac-4c89-aa0d-0928ce4c6150',  // "2026" custom view — locked to 1/1/2026–12/31/2026
    tabName: 'B2B Order Log',
    dateFilterStart: 'Start Date',                                    // Tableau parameter display name
    dateFilterEnd: 'End Date',                                         // Tableau parameter display name
    dateRangeDays: 30,
    dateFilterColumn: 'sp.Order Date (copy)',  // Client-side date filter (rolling 30-day window)
    timeColumns: ['Order Time (Timezone)'],    // Extract time-only from datetime values
    columns: [
      'Owner & Office',
      'Rep',
      'sp.Order Date (copy)',
      'Order Time (Timezone)',
      'sp.SPM Number',
      'spe.Name',
      'spe.Account BAN',
      'Product Type (Broken Out)',
      'CRU/IRU',
      'DTR Status (enriched)',
      'Disconnect Reason (Consolidated)',
      'spe.Port Carrier',
      'Order Status',
      'Voice Line Count',
      'spe.TN Type',
      'spe.Phone',
      'IF/OOF',
      'Package',
      'spe.Install Date',
      'Auto Bill Pay',
      'B2B Rep Volume Bonus Tiers',
      'Tier Bonus Payout/DNQ Reason',
      'DD Date',
      'Measure Values'
    ]
  }

  // Add more reports here:
  // 'churn-rates': {
  //   viewContentUrl: 'ATTTRACKER-B2B/CHURNRATES',
  //   tabName: 'Churn Rates',
  //   dateFilterField: 'Order Date',
  //   dateRangeDays: 90,
  //   columns: [ ... ]
  // }
};


// === TABLEAU REST API ===

/**
 * Sign in to Tableau REST API and return { token, siteId }
 */
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

  var code = resp.getResponseCode();
  var body = resp.getContentText();
  if (code !== 200) {
    throw new Error('Tableau sign-in failed (HTTP ' + code + '): ' + body.substring(0, 300));
  }

  var json = JSON.parse(body);
  return {
    token: json.credentials.token,
    siteId: json.credentials.site.id
  };
}

/**
 * Sign out of Tableau (invalidate token)
 */
function _tableauSignOut(config, token) {
  try {
    UrlFetchApp.fetch(config.server + '/api/' + TABLEAU_API_VERSION + '/auth/signout', {
      method: 'post',
      headers: { 'X-Tableau-Auth': token },
      muteHttpExceptions: true
    });
  } catch (e) { /* non-critical */ }
}

/**
 * Find view LUID by content URL path (e.g. "ATTTRACKER-B2B/ORDERLOG")
 * Returns { viewId, workbookId }
 */
function _findView(config, token, siteId, viewContentUrl) {
  // Split into workbook and view name
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

  var code = resp.getResponseCode();
  var body = resp.getContentText();
  if (code !== 200) {
    throw new Error('Failed to query views (HTTP ' + code + '): ' + body.substring(0, 300));
  }

  var json = JSON.parse(body);
  var views = json.views && json.views.view ? json.views.view : [];

  // Log all found views for debugging
  for (var i = 0; i < views.length; i++) {
    Logger.log('View [' + i + ']: contentUrl="' + views[i].contentUrl + '" id="' + views[i].id + '" workbook="' + (views[i].workbook ? views[i].workbook.id : 'n/a') + '"');
  }

  // Match by full content URL (workbook/view)
  for (var i = 0; i < views.length; i++) {
    if (views[i].contentUrl === viewContentUrl) {
      return { viewId: views[i].id, workbookId: views[i].workbook ? views[i].workbook.id : null };
    }
  }

  // Fallback: partial match — check if contentUrl contains the workbook name
  var workbookName = parts[0]; // e.g. "ATTTRACKER-B2B"
  for (var i = 0; i < views.length; i++) {
    if (views[i].contentUrl && views[i].contentUrl.indexOf(workbookName) !== -1) {
      Logger.log('Partial match: ' + views[i].contentUrl);
      return { viewId: views[i].id, workbookId: views[i].workbook ? views[i].workbook.id : null };
    }
  }

  // Fallback: single result
  if (views.length === 1) {
    return { viewId: views[0].id, workbookId: views[0].workbook ? views[0].workbook.id : null };
  }

  throw new Error('View not found: ' + viewContentUrl + '. Found ' + views.length + ' views named "' + viewName + '"');
}


// === CUSTOM VIEW DATA DOWNLOAD (primary) ===

/**
 * Download CSV data from a Tableau Custom View.
 * Custom views have their own saved parameter state (e.g. locked date range),
 * so the REST API returns data with those parameters applied — no vp_ needed.
 */
function _downloadCustomViewData(config, token, siteId, customViewId) {
  var url = config.server + '/api/' + TABLEAU_API_VERSION
    + '/sites/' + siteId + '/customviews/' + customViewId + '/data';

  Logger.log('Downloading custom view data: ' + customViewId);

  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-Tableau-Auth': token },
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('Custom view download failed (HTTP ' + code + '): ' + resp.getContentText().substring(0, 300));
  }

  return resp.getContentText();
}


// === REST API DATA DOWNLOAD (fallback) ===

/**
 * Download view data as CSV, with optional date range filter.
 * Supports two modes:
 *   1. Separate Start/End filters (dateFilterStart + dateFilterEnd) — matches Tableau views
 *      with distinct Start Date / End Date filter controls
 *   2. Single field filter (dateFilterField) — legacy, uses vf_ discrete match
 */
function _downloadViewData(config, token, siteId, viewId, report) {
  var url = config.server + '/api/' + TABLEAU_API_VERSION
    + '/sites/' + siteId + '/views/' + viewId + '/data';

  var dateRangeDays = report.dateRangeDays || 30;
  var endDate = new Date();
  var startDate = new Date();
  startDate.setDate(startDate.getDate() - dateRangeDays);
  var startStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'M/d/yyyy');
  var endStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'M/d/yyyy');

  // Mode 1: Separate Start Date / End Date parameters (vp_ = Tableau Parameters)
  if (report.dateFilterStart && report.dateFilterEnd) {
    url += '?vp_' + encodeURIComponent(report.dateFilterStart) + '=' + startStr
         + '&vp_' + encodeURIComponent(report.dateFilterEnd) + '=' + endStr;
    Logger.log('Date params (vp_): ' + report.dateFilterStart + '=' + startStr
      + ', ' + report.dateFilterEnd + '=' + endStr);
  }
  // Mode 2: Single field (legacy discrete filter)
  else if (report.dateFilterField) {
    url += '?vf_' + encodeURIComponent(report.dateFilterField) + '=' + encodeURIComponent(startStr + ',' + endStr);
    Logger.log('Date filter (vf_): ' + report.dateFilterField + '=' + startStr + ',' + endStr);
  }

  Logger.log('Download URL: ' + url);

  var resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-Tableau-Auth': token },
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('Failed to download view data (HTTP ' + code + '): ' + resp.getContentText().substring(0, 200));
  }

  return resp.getContentText();
}


// === CSV PARSING & COLUMN FILTERING ===

/**
 * Parse CSV text into 2D array, handling quoted fields with commas/newlines
 */
function _parseCsv(csvText) {
  // Use Utilities.parseCsv for robust parsing
  return Utilities.parseCsv(csvText);
}

/**
 * Filter a 2D array to only include specified columns (by header name)
 * Returns { headers: [...], rows: [[...], ...] }
 */
function _filterColumns(data, columnsToKeep) {
  if (!data || data.length === 0) return { headers: [], rows: [] };

  var headerRow = data[0];
  var keepIndices = [];
  var filteredHeaders = [];

  for (var c = 0; c < columnsToKeep.length; c++) {
    var colName = columnsToKeep[c];
    var idx = headerRow.indexOf(colName);
    if (idx !== -1) {
      keepIndices.push(idx);
      filteredHeaders.push(colName);
    } else {
      Logger.log('WARNING: Column not found in Tableau data: "' + colName + '"');
    }
  }

  var filteredRows = [];
  for (var r = 1; r < data.length; r++) {
    var row = [];
    for (var k = 0; k < keepIndices.length; k++) {
      row.push(data[r][keepIndices[k]] || '');
    }
    filteredRows.push(row);
  }

  return { headers: filteredHeaders, rows: filteredRows };
}

/**
 * Filter rows to only include those within the last N days (client-side).
 * The Tableau API vf_ date filter uses discrete matching, not ranges,
 * so this is the reliable way to enforce a rolling date window.
 */
function _filterByDateRange(filtered, dateColumn, rangeDays) {
  var idx = filtered.headers.indexOf(dateColumn);
  if (idx === -1) {
    Logger.log('WARNING: dateFilterColumn "' + dateColumn + '" not found in filtered headers — skipping date filter');
    return filtered;
  }

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - rangeDays);
  cutoff.setHours(0, 0, 0, 0);
  Logger.log('Date filter cutoff: ' + cutoff.toISOString());

  // Debug: log first 5 raw date values to diagnose format issues
  var sampleCount = Math.min(filtered.rows.length, 5);
  for (var s = 0; s < sampleCount; s++) {
    var rawVal = filtered.rows[s][idx];
    var parsed = rawVal ? new Date(rawVal) : null;
    Logger.log('  Sample date [' + s + ']: raw="' + rawVal + '" type=' + typeof rawVal
      + ' parsed=' + (parsed && !isNaN(parsed.getTime()) ? parsed.toISOString() : 'INVALID')
      + ' >= cutoff? ' + (parsed && !isNaN(parsed.getTime()) ? (parsed >= cutoff) : 'N/A'));
  }

  var beforeCount = filtered.rows.length;
  filtered.rows = filtered.rows.filter(function(row) {
    var val = row[idx];
    if (!val) return false;
    var d = new Date(val);
    return !isNaN(d.getTime()) && d >= cutoff;
  });

  Logger.log('Date filter (' + dateColumn + ', last ' + rangeDays + 'd): ' + beforeCount + ' → ' + filtered.rows.length + ' rows');
  return filtered;
}

/**
 * Extract time-only (HH:MM) from datetime values in specified columns.
 * Tableau CSV may export time fields as full datetimes (e.g. "2026-03-10 14:30:00").
 * This converts them to just the time portion for cleaner display.
 */
function _formatTimeColumns(filtered, timeColumns) {
  if (!timeColumns || timeColumns.length === 0) return;

  var indices = [];
  for (var c = 0; c < timeColumns.length; c++) {
    var idx = filtered.headers.indexOf(timeColumns[c]);
    if (idx !== -1) {
      indices.push(idx);
    } else {
      Logger.log('WARNING: timeColumn "' + timeColumns[c] + '" not found — skipping');
    }
  }
  if (indices.length === 0) return;

  var formatted = 0;
  for (var r = 0; r < filtered.rows.length; r++) {
    for (var i = 0; i < indices.length; i++) {
      var val = filtered.rows[r][indices[i]];
      if (!val) continue;
      var str = String(val).trim();
      // Already time-only (e.g. "14:30" or "2:30 PM") — leave it
      if (/^\d{1,2}:\d{2}/.test(str) && !/^\d{4}-/.test(str)) continue;
      // Try parsing as datetime and extract time
      var d = new Date(str);
      if (!isNaN(d.getTime())) {
        var h = d.getHours(), m = d.getMinutes();
        var ampm = h >= 12 ? 'PM' : 'AM';
        var h12 = h % 12 || 12;
        filtered.rows[r][indices[i]] = h12 + ':' + String(m).padStart(2, '0') + ' ' + ampm;
        formatted++;
      }
    }
  }
  Logger.log('Formatted ' + formatted + ' time values across ' + indices.length + ' column(s)');
}


// === GOOGLE SHEET WRITING ===

/**
 * Write filtered data to a specific tab in the output sheet
 * Clears existing data first, then writes headers + rows
 */
function _writeToSheet(sheetId, tabName, headers, rows) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName(tabName);

  // Create tab if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
  }

  // Clear existing data
  sheet.clearContents();

  // Write headers
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Bold + freeze header row
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Write data rows
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // Auto-resize columns (up to 25 to avoid timeout)
  var resizeCols = Math.min(headers.length, 25);
  for (var i = 1; i <= resizeCols; i++) {
    sheet.autoResizeColumn(i);
  }

  return { tab: tabName, headers: headers.length, rows: rows.length };
}


// === SYNC FUNCTIONS ===

/**
 * Sync a single report by key (e.g. 'b2b-order-log')
 */
function syncReport(reportKey) {
  var report = REPORTS[reportKey];
  if (!report) throw new Error('Unknown report: ' + reportKey);

  var config = _getConfig();
  if (!config.email || !config.password) throw new Error('Tableau credentials not set in Script Properties');
  if (!config.sheetId) throw new Error('SHEET_ID not set in Script Properties');

  var auth = _tableauSignIn(config);
  try {
    var csv;
    var source;

    // Primary: Custom View (has locked parameters — date range baked in)
    if (report.customViewId) {
      try {
        csv = _downloadCustomViewData(config, auth.token, auth.siteId, report.customViewId);
        source = 'customView';
        Logger.log('Custom view download: ' + csv.length + ' chars');
      } catch (cvErr) {
        Logger.log('Custom view failed: ' + cvErr.message + ' — falling back to default view');
        csv = null;
      }
    }

    // Fallback: Default view via REST API (uses view's default parameters)
    if (!csv) {
      var viewInfo = _findView(config, auth.token, auth.siteId, report.viewContentUrl);
      var viewId = viewInfo.viewId;
      Logger.log('Found view: ' + viewId);
      csv = _downloadViewData(config, auth.token, auth.siteId, viewId, report);
      source = 'restApi';
      Logger.log('REST API download: ' + csv.length + ' chars');
    }

    // Parse and filter
    var data = _parseCsv(csv);
    Logger.log('Parsed ' + data.length + ' rows (including header), source: ' + source);

    var filtered = _filterColumns(data, report.columns);
    Logger.log('Filtered to ' + filtered.headers.length + ' columns, ' + filtered.rows.length + ' data rows');

    // Client-side date filter — trims the wide custom view range to a rolling window
    if (report.dateFilterColumn && report.dateRangeDays) {
      filtered = _filterByDateRange(filtered, report.dateFilterColumn, report.dateRangeDays);
    }

    // Extract time-only from datetime columns
    if (report.timeColumns) {
      _formatTimeColumns(filtered, report.timeColumns);
    }

    // Write to sheet
    var result = _writeToSheet(config.sheetId, report.tabName, filtered.headers, filtered.rows);
    Logger.log('Wrote to sheet: ' + JSON.stringify(result));

    return {
      ok: true,
      report: reportKey,
      source: source,
      totalRowsFromTableau: data.length - 1,
      filteredColumns: filtered.headers.length,
      rowsWritten: filtered.rows.length
    };
  } finally {
    _tableauSignOut(config, auth.token);
  }
}

/**
 * Sync ALL registered reports
 */
function syncAllReports() {
  var results = [];
  var keys = Object.keys(REPORTS);
  for (var i = 0; i < keys.length; i++) {
    try {
      var result = syncReport(keys[i]);
      results.push(result);
      Logger.log('✓ ' + keys[i] + ': ' + result.rowsWritten + ' rows');
    } catch (e) {
      results.push({ ok: false, report: keys[i], error: e.message });
      Logger.log('✗ ' + keys[i] + ': ' + e.message);
    }
  }
  return results;
}


// === DISTRIBUTION TARGETS ===
// Each target maps a synced report → a destination campaign sheet + caches to bust.
//
// WHY campaign-based (not per-office)?
//   B2B offices share the SAME Google Sheet and _TableauOrderLog tab.
//   Code.gs readTableauSummary already scopes data per-office by matching DSIs.
//   So we write ALL rows once → bust cache for each office.
//
// filterColumn:     null = write ALL rows (shared sheets, Code.gs handles scoping)
//                   'Owner & Office' = filter per-office using ownerOfficeMatch
// ownerOfficeMatch: only needed when filterColumn is set (substring, case-insensitive)

var TARGETS = [
  {
    reportKey: 'b2b-order-log',       // Must match a REPORTS key
    sourceTab: 'B2B Order Log',       // Tab name in the Tableau Data sheet
    sheetId: '1wxM6Htwfy8LrD_o_C7gmvnZEmkfV3FTCVjJU6IITZFc',  // Shared B2B campaign sheet
    tabName: '_TableauOrderLog',      // Destination tab in the campaign sheet
    filterColumn: null,               // null = write all rows (shared sheet)
    offices: [
      {
        officeId: 'off_001',
        appsScriptUrl: 'https://script.google.com/macros/s/AKfycbwPx0jfdYdLKurHPlfQhOkYu70vVpirTISYrR3I2EIszVrVaRNwwjBvauSIO69thKFe/exec',
        apiKey: 'elevate-dash-2026-secret'
      },
      {
        officeId: 'off_002',
        appsScriptUrl: 'https://script.google.com/macros/s/AKfycbwPx0jfdYdLKurHPlfQhOkYu70vVpirTISYrR3I2EIszVrVaRNwwjBvauSIO69thKFe/exec',
        apiKey: 'elevate-dash-2026-secret'
      }
    ]
  }
  // Future: add NDS target, separate-sheet offices, etc.
  // Example with per-office filtering (if an office has its own sheet):
  // {
  //   reportKey: 'b2b-order-log',
  //   sourceTab: 'B2B Order Log',
  //   sheetId: 'SEPARATE_SHEET_ID',
  //   tabName: '_TableauOrderLog',
  //   filterColumn: 'Owner & Office',
  //   ownerOfficeMatch: 'some office name',
  //   offices: [
  //     { officeId: 'off_003', appsScriptUrl: '...', apiKey: '...' }
  //   ]
  // }
];


// === DISTRIBUTE TO OFFICES ===

/**
 * Read synced Tableau data from the shared Tableau Data sheet,
 * write to each target's campaign sheet, and bust office caches.
 *
 * Handles two modes:
 *   1. Shared sheet (filterColumn=null): write ALL rows once, bust multiple caches
 *   2. Filtered (filterColumn set): filter by ownerOfficeMatch, write per-target
 */
function distributeToOffices() {
  var config = _getConfig();
  if (!config.sheetId) throw new Error('SHEET_ID not set');
  if (!TARGETS.length) {
    Logger.log('No targets configured — skipping distribution');
    return [];
  }

  var tableauSS = SpreadsheetApp.openById(config.sheetId);
  var results = [];

  for (var t = 0; t < TARGETS.length; t++) {
    var target = TARGETS[t];
    try {
      // Read source data from the Tableau Data sheet
      var srcSheet = tableauSS.getSheetByName(target.sourceTab);
      if (!srcSheet || srcSheet.getLastRow() < 2) {
        Logger.log(target.reportKey + ': no data in "' + target.sourceTab + '" tab — skipping');
        results.push({ ok: false, reportKey: target.reportKey, error: 'No source data' });
        continue;
      }

      var data = srcSheet.getDataRange().getValues();
      var headers = data[0];
      var rows = data.slice(1);

      // Apply filter if configured
      if (target.filterColumn && target.ownerOfficeMatch) {
        var filterIdx = headers.indexOf(target.filterColumn);
        if (filterIdx === -1) {
          throw new Error('"' + target.filterColumn + '" column not found');
        }
        var match = target.ownerOfficeMatch.toLowerCase();
        rows = rows.filter(function(row) {
          return String(row[filterIdx] || '').toLowerCase().indexOf(match) !== -1;
        });
        Logger.log(target.reportKey + ': filtered to ' + rows.length + ' rows (match: "' + match + '")');
      } else {
        Logger.log(target.reportKey + ': writing all ' + rows.length + ' rows (shared sheet, no filter)');
      }

      // Write to the destination campaign sheet
      var destSS = SpreadsheetApp.openById(target.sheetId);
      var destTab = destSS.getSheetByName(target.tabName);
      if (!destTab) {
        destTab = destSS.insertSheet(target.tabName);
      }

      destTab.clearContents();
      if (headers.length > 0) {
        destTab.getRange(1, 1, 1, headers.length).setValues([headers]);
        destTab.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        destTab.setFrozenRows(1);
      }
      if (rows.length > 0) {
        destTab.getRange(2, 1, rows.length, headers.length).setValues(rows);
      }

      Logger.log(target.reportKey + ': wrote ' + rows.length + ' rows to ' + target.tabName);

      // Bust cache for each office that reads from this sheet
      for (var o = 0; o < target.offices.length; o++) {
        var office = target.offices[o];
        var bustResult = _bustOfficeCache(office);
        Logger.log('  ' + office.officeId + ': cache bust → ' + bustResult);
      }

      results.push({
        ok: true,
        reportKey: target.reportKey,
        rowsWritten: rows.length,
        officesBusted: target.offices.map(function(o) { return o.officeId; })
      });
    } catch (e) {
      Logger.log(target.reportKey + ': ERROR — ' + e.message);
      results.push({ ok: false, reportKey: target.reportKey, error: e.message });
    }
  }

  return results;
}

/**
 * Bust an office's Tableau cache via their Apps Script URL
 */
function _bustOfficeCache(office) {
  try {
    var url = office.appsScriptUrl
      + '?key=' + encodeURIComponent(office.apiKey)
      + '&officeId=' + encodeURIComponent(office.officeId)
      + '&action=bustTableauCache';

    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });
    return 'HTTP ' + resp.getResponseCode();
  } catch (e) {
    return 'ERROR: ' + e.message;
  }
}


// === FULL NIGHTLY PIPELINE ===

/**
 * Master function: sync from Tableau, then distribute to offices.
 * Use this as the single daily trigger.
 */
function nightlySync() {
  Logger.log('=== NIGHTLY SYNC START ===');

  // Step 1: Pull from Tableau API → shared sheet
  var syncResults = syncAllReports();
  Logger.log('Sync: ' + JSON.stringify(syncResults));

  // Step 2: Distribute to each office's sheet + bust caches
  var distResults = distributeToOffices();
  Logger.log('Distribute: ' + JSON.stringify(distResults));

  Logger.log('=== NIGHTLY SYNC COMPLETE ===');
  return { sync: syncResults, distribute: distResults };
}


// === DAILY TRIGGER ===

/**
 * Run this ONCE from the editor to set up the nightly trigger.
 * Replaces any existing triggers for syncAllReports or nightlySync.
 */
function setupDailyTrigger() {
  // Remove existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'syncAllReports' || fn === 'nightlySync') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Create new daily trigger at 1 AM — runs full pipeline (sync + distribute)
  ScriptApp.newTrigger('nightlySync')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .nearMinute(0)
    .create();

  Logger.log('Daily trigger set: nightlySync at ~1:00 AM (sync + distribute to offices)');
}

/**
 * Remove the daily trigger
 */
function removeDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'syncAllReports' || fn === 'nightlySync') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  Logger.log('Removed ' + removed + ' trigger(s)');
}


// === WEB APP ENDPOINT (optional — for dashboard refresh button) ===

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'syncAll';
  var key = (e && e.parameter && e.parameter.key) || '';

  // Simple API key check (set SYNC_API_KEY in Script Properties)
  var apiKey = PropertiesService.getScriptProperties().getProperty('SYNC_API_KEY') || '';
  if (apiKey && key !== apiKey) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid API key' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var result;
  try {
    if (action === 'syncAll') {
      result = syncAllReports();
    } else if (action === 'sync' && e.parameter.report) {
      result = syncReport(e.parameter.report);
    } else if (action === 'listReports') {
      result = { reports: Object.keys(REPORTS) };
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


// === MANUAL TEST FUNCTIONS ===

/**
 * Test Tableau connection (run from editor)
 */
function testConnection() {
  var config = _getConfig();
  Logger.log('Connecting to: ' + config.server + ' as ' + config.email + ' (site: ' + config.site + ')');

  var auth = _tableauSignIn(config);
  Logger.log('SUCCESS — Token: ' + auth.token.substring(0, 10) + '... | Site ID: ' + auth.siteId);
  _tableauSignOut(config, auth.token);
  Logger.log('Signed out.');
}

/**
 * Test syncing just the B2B Order Log (run from editor)
 */
function testSyncB2BOrderLog() {
  var result = syncReport('b2b-order-log');
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * List unique "Owner & Office" values from synced data (run from editor).
 * Useful for configuring ownerOfficeMatch in TARGETS.
 */
function listUniqueOwnerOffice() {
  var config = _getConfig();
  var ss = SpreadsheetApp.openById(config.sheetId);
  var sheet = ss.getSheetByName('B2B Order Log');
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('No data in B2B Order Log');
    return;
  }
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idx = headers.indexOf('Owner & Office');
  if (idx === -1) {
    Logger.log('"Owner & Office" column not found');
    return;
  }
  var unique = {};
  for (var r = 1; r < data.length; r++) {
    var val = String(data[r][idx] || '').trim();
    if (val) unique[val] = (unique[val] || 0) + 1;
  }
  var keys = Object.keys(unique).sort();
  Logger.log('=== ' + keys.length + ' UNIQUE "Owner & Office" VALUES ===');
  for (var i = 0; i < keys.length; i++) {
    Logger.log('[' + i + '] "' + keys[i] + '" (' + unique[keys[i]] + ' rows)');
  }
}

/**
 * Test the distribution pipeline (run from editor)
 */
function testDistribute() {
  var results = distributeToOffices();
  Logger.log(JSON.stringify(results, null, 2));
}

/**
 * Log all column headers from the custom view CSV (run from editor to find exact column names).
 * Also logs which REPORTS columns matched and which didn't.
 */
function logColumnHeaders() {
  var config = _getConfig();
  var auth = _tableauSignIn(config);
  try {
    var report = REPORTS['b2b-order-log'];
    var csv;
    if (report.customViewId) {
      csv = _downloadCustomViewData(config, auth.token, auth.siteId, report.customViewId);
      Logger.log('Source: Custom View ' + report.customViewId);
    } else {
      var viewInfo = _findView(config, auth.token, auth.siteId, report.viewContentUrl);
      csv = _downloadViewData(config, auth.token, auth.siteId, viewInfo.viewId, report);
      Logger.log('Source: Default REST API view');
    }
    var data = _parseCsv(csv);
    var headers = data[0];
    Logger.log('=== ALL ' + headers.length + ' COLUMN HEADERS FROM CSV ===');
    for (var i = 0; i < headers.length; i++) {
      Logger.log('[' + i + '] "' + headers[i] + '"');
    }

    // Check which REPORTS columns match
    Logger.log('=== COLUMN MATCHING ===');
    var reportCols = report.columns || [];
    for (var c = 0; c < reportCols.length; c++) {
      var found = headers.indexOf(reportCols[c]) !== -1;
      Logger.log((found ? '✓' : '✗') + ' "' + reportCols[c] + '"');
    }

    // Log sample data row
    if (data.length > 1) {
      Logger.log('=== SAMPLE ROW (row 2) ===');
      for (var j = 0; j < headers.length; j++) {
        Logger.log('  ' + headers[j] + ': "' + String(data[1][j] || '').substring(0, 80) + '"');
      }
    }
  } finally {
    _tableauSignOut(config, auth.token);
  }
}

/**
 * Diagnose the full pipeline: check _TableauOrderLog headers, row count, and DSI matching.
 * Run from editor to see why statuses aren't showing on the dashboard.
 */
function diagnosePipeline() {
  var config = _getConfig();
  var ss = SpreadsheetApp.openById(config.sheetId);  // Tableau Data sheet

  // Step 1: Check intermediate "B2B Order Log" tab (in Tableau Data sheet)
  var srcSheet = ss.getSheetByName('B2B Order Log');
  if (!srcSheet || srcSheet.getLastRow() < 2) {
    Logger.log('❌ "B2B Order Log" tab has NO DATA — run testSyncB2BOrderLog() first');
    return;
  }
  var srcData = srcSheet.getDataRange().getValues();
  Logger.log('✓ "B2B Order Log" (Tableau Data sheet): ' + (srcData.length - 1) + ' rows, ' + srcData[0].length + ' columns');
  Logger.log('  Headers: ' + srcData[0].join(', '));

  // Step 2: Check _TableauOrderLog tab in the CAMPAIGN sheet (where Code.gs reads from)
  // This is a DIFFERENT spreadsheet — TARGETS[0].sheetId, not config.sheetId
  var campaignSheetId = TARGETS.length > 0 ? TARGETS[0].sheetId : config.sheetId;
  Logger.log('Campaign sheet ID: ' + campaignSheetId);
  Logger.log('Tableau Data sheet ID: ' + config.sheetId);
  Logger.log('Same sheet? ' + (campaignSheetId === config.sheetId ? 'YES' : 'NO — distribution required'));
  var campaignSS = SpreadsheetApp.openById(campaignSheetId);
  var tolSheet = campaignSS.getSheetByName('_TableauOrderLog');
  if (!tolSheet || tolSheet.getLastRow() < 2) {
    Logger.log('❌ "_TableauOrderLog" tab has NO DATA in campaign sheet — run testDistribute()');
    return;
  }
  var tolData = tolSheet.getDataRange().getValues();
  Logger.log('✓ "_TableauOrderLog": ' + (tolData.length - 1) + ' rows, ' + tolData[0].length + ' columns');
  Logger.log('  Headers: ' + tolData[0].join(', '));

  // Step 3: Check DSI column existence
  var headers = tolData[0].map(function(h) { return String(h || '').trim().toLowerCase(); });
  var dsiIdx = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === 'sp.spm number' || headers[i] === 'spm number' || headers[i] === 'dsi') {
      dsiIdx = i;
      break;
    }
  }
  if (dsiIdx === -1) {
    Logger.log('❌ No DSI column found in _TableauOrderLog! Headers: ' + headers.join(', '));
    return;
  }
  Logger.log('✓ DSI column found at index ' + dsiIdx + ' ("' + tolData[0][dsiIdx] + '")');

  // Step 4: Check DTR Status column
  var statusIdx = -1;
  for (var j = 0; j < headers.length; j++) {
    if (headers[j] === 'dtr status (enriched)' || headers[j] === 'dtr status') {
      statusIdx = j;
      break;
    }
  }
  if (statusIdx === -1) {
    Logger.log('❌ No DTR Status column found!');
  } else {
    // Count unique statuses
    var statusCounts = {};
    for (var r = 1; r < tolData.length; r++) {
      var st = String(tolData[r][statusIdx] || '').trim();
      statusCounts[st] = (statusCounts[st] || 0) + 1;
    }
    Logger.log('✓ DTR Status column found. Status distribution:');
    Object.keys(statusCounts).sort().forEach(function(s) {
      Logger.log('  ' + (s || '(empty)') + ': ' + statusCounts[s]);
    });
  }

  // Step 5: Sample some DSIs from _TableauOrderLog
  var sampleDsis = [];
  for (var s = 1; s < Math.min(6, tolData.length); s++) {
    sampleDsis.push(String(tolData[s][dsiIdx] || '').trim());
  }
  Logger.log('Sample DSIs from _TableauOrderLog: ' + sampleDsis.join(', '));

  // Step 6: Check _Sales_off_001 for DSI matching (in campaign sheet)
  var salesSheet = campaignSS.getSheetByName('_Sales_off_001');
  if (!salesSheet || salesSheet.getLastRow() < 2) {
    Logger.log('❌ "_Sales_off_001" tab has NO DATA — DSI→email mapping will fail');
    return;
  }
  var salesData = salesSheet.getDataRange().getValues();
  Logger.log('✓ "_Sales_off_001": ' + (salesData.length - 1) + ' rows');

  // DSI is at column index 5 (OL.DSI in Code.gs), EMAIL at index 1
  var salesDsiCol = 5;
  var salesEmailCol = 1;
  var salesDsiMap = {};  // DSI → email
  var salesDsis = [];
  for (var sd = 1; sd < salesData.length; sd++) {
    var sDsi = String(salesData[sd][salesDsiCol] || '').trim();
    var sEmail = String(salesData[sd][salesEmailCol] || '').trim().toLowerCase();
    if (sDsi && sEmail) {
      salesDsiMap[sDsi] = sEmail;
    }
    if (sd <= 5) salesDsis.push(sDsi);
  }
  Logger.log('  Sales DSI→email map: ' + Object.keys(salesDsiMap).length + ' entries');
  Logger.log('  Sample Sales DSIs: ' + salesDsis.join(', '));

  // Step 7: Test the actual join — how many Tableau DSIs match Sales DSIs?
  var matched = 0;
  var unmatched = 0;
  var unmatchedSamples = [];
  for (var m = 1; m < tolData.length; m++) {
    var tDsi = String(tolData[m][dsiIdx] || '').trim();
    if (!tDsi) continue;
    if (salesDsiMap[tDsi]) {
      matched++;
    } else {
      unmatched++;
      if (unmatchedSamples.length < 5) unmatchedSamples.push(tDsi);
    }
  }
  Logger.log('=== DSI JOIN RESULTS ===');
  Logger.log('  Matched: ' + matched + ' / ' + (matched + unmatched));
  Logger.log('  Unmatched: ' + unmatched);
  if (unmatchedSamples.length > 0) {
    Logger.log('  Unmatched samples: ' + unmatchedSamples.join(', '));
  }
  if (matched === 0 && unmatched > 0) {
    Logger.log('❌ ZERO DSI matches! Format mismatch between Tableau and Sales DSIs.');
    Logger.log('  Tableau format: "' + sampleDsis[0] + '"');
    Logger.log('  Sales format:   "' + salesDsis[0] + '"');
  }

  // Step 8: Check cache
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('tableauSummary_v6_off_001');
    if (cached) {
      var parsed = JSON.parse(cached);
      var dsiCount = Object.keys(parsed.dsiSummary || {}).length;
      var repCount = Object.keys(parsed.repSummary || {}).length;
      Logger.log('✓ Cache exists: ' + dsiCount + ' DSIs, ' + repCount + ' reps');
    } else {
      Logger.log('ℹ️ No cache entry — will be built on next dashboard load');
    }
  } catch (e) {
    Logger.log('⚠️ Cache check failed: ' + e.message);
  }

  Logger.log('=== DIAGNOSIS COMPLETE ===');
}
