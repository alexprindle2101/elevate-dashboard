// ═══════════════════════════════════════════════════════
// NationalCode.gs — National Consultant Dashboard Data Aggregator
// Reads from 4 external Google Sheets, returns unified JSON.
// Deploy as web app (Execute as: Me, Access: Anyone)
// ═══════════════════════════════════════════════════════

// ── Config ──
var NC_API_KEY = 'national-dash-2026-secret'; // Must match Script Properties > API_KEY

// External sheet IDs
// ── Owner Development (OD) Config ──
var OD_CAMPAIGNS = {
  'frontier':        { label: 'Frontier',           sheetId: '1WWpLQTCyvPmJbx3jjowFszwOF_JUnjS6tzu-eAASwk0', excludeOwners: ['Collier Ricks', 'Stephen Cancino'] },
  'verizon-fios':    { label: 'Verizon Fios',       sheetId: '12J3HBdFQrqq5D7YwWEp93US40vmZz5n9mS-KMKXaXxA' },
  'att-nds':         { label: 'AT&T NDS/Verizon',   sheetId: '1kcUWR3EKgP-9wDct4vDyuQJ7IuS0cbcetY97dmVTY64' },
  'att-res':         { label: 'AT&T Residential',   sheetId: '1HvWJYox3JXvxmza63YBWAqKPtUGPFuaV-s-BOfbWGKM' },
  'rogers':          { label: 'Rogers',             sheetId: '1o1MPKrAzzeaU2JWMODkR9M3uY5rOhIKo-Q64armeTvE' },
  'leafguard':       { label: 'LeafGuard',          sheetId: '10Fy5XFWCuBmDwvpl4PG4FJT4krwX2ZqN12ARvQLpSuM', ownerSource: 'tabs', excludeTabs: ['Blank Copy'] },
  'lumen':           { label: 'Lumen',              sheetId: '1P4DYlcV1hgNkaAapk3tWD7ytcRXw4K1n7R6EMKPCoSA', sourceTab: 'Campaign', sectionHeader: 'LUMEN' },
  'att-b2b':         { label: 'AT&T B2B',           sheetId: '1sxauFjNjq4_rRYM2PAl5cyOyHF3Hg4OkO-t_hLKDJB8', ownerSource: 'visible-tabs', excludeTabs: ["B2B Enery 1:1's", "B2B Energy 1:1's"] },
  'box-energy':      { label: 'Box Energy',          sheetId: '1_PVzLcmlo6EzySNRfah-r-NLka3IthxzLZb4TjrDgtE', ownerSource: 'visible-tabs', hidden: true }
};
var OD_NLR_FOLDER = '1hARjh3UH48CWhbYrYBJxFVwgynxapCjG';

// External sheet IDs
var SHEETS = {
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
// CACHE HELPER
// ══════════════════════════════════════════════════

var CACHE_CHUNK_SIZE = 95000; // ~95KB per chunk (under 100KB CacheService limit)

/**
 * Store a JSON string across multiple cache keys if it exceeds the chunk size.
 * Keys: baseKey_0, baseKey_1, ... baseKey_N, plus baseKey_meta with chunk count.
 */
function _cacheStore(cache, baseKey, json, ttlSeconds) {
  var chunks = Math.ceil(json.length / CACHE_CHUNK_SIZE);
  if (chunks === 1) {
    cache.put(baseKey, json, ttlSeconds);
  } else {
    var pairs = {};
    for (var i = 0; i < chunks; i++) {
      pairs[baseKey + '_' + i] = json.substring(i * CACHE_CHUNK_SIZE, (i + 1) * CACHE_CHUNK_SIZE);
    }
    pairs[baseKey + '_meta'] = String(chunks);
    cache.putAll(pairs, ttlSeconds);
  }
  Logger.log('_cacheStore: ' + baseKey + ' (' + json.length + ' bytes, ' + chunks + ' chunk' + (chunks > 1 ? 's' : '') + ', ' + ttlSeconds + 's TTL)');
}

/**
 * Read a potentially multi-chunk cached value. Returns null on miss.
 */
function _cacheRead(cache, baseKey) {
  // Try single-key first (small values)
  var single = cache.get(baseKey);
  if (single) return single;

  // Try multi-chunk
  var metaVal = cache.get(baseKey + '_meta');
  if (!metaVal) return null;

  var chunks = parseInt(metaVal) || 0;
  if (chunks < 1) return null;

  var chunkKeys = [];
  for (var i = 0; i < chunks; i++) chunkKeys.push(baseKey + '_' + i);
  var parts = cache.getAll(chunkKeys);

  var assembled = '';
  for (var i = 0; i < chunks; i++) {
    var part = parts[baseKey + '_' + i];
    if (!part) return null; // partial miss — treat as full miss
    assembled += part;
  }
  return assembled;
}

/**
 * Remove a potentially multi-chunk cached value.
 */
function _cacheRemove(cache, baseKey) {
  var metaVal = cache.get(baseKey + '_meta');
  var keysToRemove = [baseKey, baseKey + '_meta'];
  if (metaVal) {
    var chunks = parseInt(metaVal) || 0;
    for (var i = 0; i < chunks; i++) keysToRemove.push(baseKey + '_' + i);
  }
  cache.removeAll(keysToRemove);
}

/**
 * Cache-through wrapper for any read function.
 * Supports chunked storage for values >100KB.
 * @param {string} cacheKey   - Unique cache key
 * @param {number} ttlSeconds - TTL in seconds (max 21600 = 6 hours)
 * @param {Function} readFn   - Zero-arg function returning data
 * @param {boolean} [bustCache] - If true, skip cache lookup and force fresh read
 * @returns {*} Cached or freshly-read data
 */
function _cachedRead(cacheKey, ttlSeconds, readFn, bustCache) {
  var cache = CacheService.getScriptCache();

  // Try cache first (unless busting)
  if (!bustCache) {
    try {
      var cached = _cacheRead(cache, cacheKey);
      if (cached) {
        var parsed = JSON.parse(cached);
        Logger.log('_cachedRead HIT: ' + cacheKey);
        return parsed;
      }
    } catch (e) {
      Logger.log('_cachedRead parse error for ' + cacheKey + ': ' + e.message);
    }
  }

  // Cache miss or bust — do the actual read
  Logger.log('_cachedRead MISS: ' + cacheKey + (bustCache ? ' (bust)' : ''));
  var result = readFn();

  // Store in cache (chunked if needed)
  try {
    var json = JSON.stringify(result);
    _cacheStore(cache, cacheKey, json, ttlSeconds);
  } catch (e) {
    Logger.log('_cachedRead store error for ' + cacheKey + ': ' + e.message);
  }

  return result;
}

// ══════════════════════════════════════════════════
// ENTRY POINT
// ══════════════════════════════════════════════════

function doGet(e) {
  const key = (e && e.parameter && e.parameter.key) || '';
  if (!validateKey(key)) return jsonResp({ error: 'unauthorized' });

  const action = (e && e.parameter && e.parameter.action) || '';
  const campaign = (e && e.parameter && e.parameter.campaign) || 'att-b2b';
  const owner = (e && e.parameter && e.parameter.owner) || '';
  const bustCache = (e && e.parameter && e.parameter.bustCache) === 'true';

  try {
    // ── Consolidated recruiting data (per-campaign tabs, cached per-campaign) ──
    if (action === 'recruiting') {
      var weeks = parseInt(e.parameter.weeks) || 6;
      var campaignFilter = e.parameter.campaign || '';
      return jsonResp(readConsolidatedRecruiting(weeks, campaignFilter, bustCache));
    }

    // ── Refresh: pull latest data from source spreadsheets into consolidated tabs ──
    if (action === 'refreshCampaigns') {
      return jsonResp(refreshAllCampaigns());
    }
    if (action === 'refreshCampaign') {
      var campaignKey = e.parameter.campaign;
      return jsonResp(refreshSingleCampaign(campaignKey));
    }

    // ── Owner → Cam Company mapping (reads _OD_Mappings + legacy _OwnerCamMapping) ──
    if (action === 'ownerCamMapping') {
      return jsonResp(_cachedRead('cache_ownerCamMapping', 300, readOwnerCamMapping, bustCache));
    }

    // ── Online presence / audit data from Cam's sheet ──
    if (action === 'onlinePresence') {
      return jsonResp(_cachedRead('cache_onlinePresence', 600, readOnlinePresence, bustCache));
    }

    // ── Per-owner NLR data: reads single owner's mapped NLR workbook ──
    if (action === 'ownerNlrData') {
      var ownerParam = (e.parameter.owner) || '';
      var campaignParam = (e.parameter.campaign) || '';
      return jsonResp(readOwnerNlrData(ownerParam, campaignParam));
    }

    // ── NDS production/sales data (read directly from NDS One-on-Ones sheet) ──
    if (action === 'ndsProduction') {
      return jsonResp(readNDSProduction());
    }

    // ── NDS per-owner sales data (fetched on demand when opening owner detail) ──
    if (action === 'ndsOwnerSales') {
      var ownerParam = e.parameter.owner || '';
      return jsonResp(readNDSOwnerSales(ownerParam));
    }

    // ── Verizon FIOS per-owner sales data (from Credico sync sheet) ──
    if (action === 'fiosOwnerSales') {
      var ownerParam = e.parameter.owner || '';
      return jsonResp(readFiosOwnerSales(ownerParam));
    }

    // ── AT&T B2B per-owner sales data (from Tableau sync) ──
    if (action === 'b2bOwnerSales') {
      var ownerParam = e.parameter.owner || '';
      return jsonResp(readB2BOwnerSales(ownerParam));
    }

    // ── AT&T Res per-owner sales data (Campaign Tracker Section 3) ──
    if (action === 'resOwnerSales') {
      var ownerParam = e.parameter.owner || '';
      return jsonResp(readResOwnerSales(ownerParam));
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

    // ── OD: Campaign owners (Column A from each campaign sheet) ──
    // Also returns tab names for each campaign so frontend can show dropdowns
    if (action === 'odCampaignOwners') {
      return jsonResp(odGetCampaignOwners());
    }

    // ── OD: Campaign tab mappings (owner → tab name overrides) ──
    if (action === 'odGetCampaignTabMap') {
      return jsonResp(odGetCampaignTabMap());
    }

    // ── OD: List NLR workbooks in Drive folder ──
    // ?includeTabs=true to also fetch tab names (slower — opens each spreadsheet)
    if (action === 'odNlrWorkbooks') {
      var includeTabs = (e && e.parameter && e.parameter.includeTabs) === 'true';
      return jsonResp(odGetNlrWorkbooks(includeTabs));
    }

    // ── OD: Get tab names from a spreadsheet ──
    if (action === 'odNlrTabs') {
      var sheetId = (e && e.parameter && e.parameter.sheetId) || '';
      if (!sheetId) return jsonResp({ error: 'sheetId parameter required' });
      return jsonResp(odGetNlrTabs(sheetId));
    }

    // ── OD: Get mappings ──
    if (action === 'odGetMappings') {
      return jsonResp(odGetMappings());
    }

    // ── OD: Get users (excludes pinHash) ──
    if (action === 'odGetUsers') {
      return jsonResp(odGetUsers());
    }

    // ── OD: Check if user exists (email only, no PIN) ──
    if (action === 'odCheckUser') {
      var email = (e && e.parameter && e.parameter.email) || '';
      if (!email) return jsonResp({ success: false, message: 'Email is required' });
      return jsonResp(odCheckUser(email));
    }

    // ── OD: Login ──
    if (action === 'odLogin') {
      var email = (e && e.parameter && e.parameter.email) || '';
      var pin = (e && e.parameter && e.parameter.pin) || '';
      if (!email || !pin) return jsonResp({ success: false, message: 'email and pin required' });
      return jsonResp(odLogin(email, pin));
    }

    // ── OD: Cam companies from Performance Audit ──
    if (action === 'odCamCompanies') {
      return jsonResp(odGetCamCompanies());
    }

    // ── D2D Residential ranking from _D2D_Res_Ranking tab ──
    if (action === 'd2dResRanking') {
      return jsonResp(readD2DResRanking());
    }

    // ── OD: Get weekly planning schedule ──
    if (action === 'odGetPlanning') {
      return jsonResp(odGetPlanning());
    }

    // ── OD: Get flagged reps (unresolved only) ──
    if (action === 'odGetFlaggedReps') {
      return jsonResp(odGetFlaggedReps());
    }

    // ── OD: Get owner notes log ──
    if (action === 'odGetNotes') {
      return jsonResp(odGetNotes_());
    }

    // ── OD: Campaign ownership (who owns which campaigns) ──
    if (action === 'odGetCampaignOwnership') {
      return jsonResp(odGetCampaignOwnership());
    }

    // ── OD: Access grants ──
    if (action === 'odGetAccessGrants') {
      var callerEmail = (e && e.parameter && e.parameter.email) || '';
      return jsonResp(odGetAccessGrants(callerEmail));
    }

    // ── OD: Resolve visible campaigns for a user ──
    if (action === 'odGetVisibleCampaigns') {
      var email = (e && e.parameter && e.parameter.email) || '';
      var role = (e && e.parameter && e.parameter.role) || '';
      return jsonResp(odResolveVisibleCampaigns(email, role));
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
      case 'refreshCampaigns':
        result = refreshAllCampaigns();
        break;
      case 'refreshCampaign':
        result = refreshSingleCampaign(body.campaign);
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
      case 'updateHeadcount':
        result = updateHeadcountRow(body.ownerName, body.date,
                   body.active, body.leaders, body.dist, body.training, body.campaignLabel);
        break;
      case 'updateProduction':
        result = updateProductionRow(body.ownerName, body.date, body.campaignLabel, body.products,
                   body.internet, body.wireless, body.dtv, body.goals);
        break;
      case 'saveGoals':
        result = saveGoalsRow_(body.ownerName, body.campaignLabel, body.campaignKey, body.goals);
        break;
      case 'addOwnerNote':
        result = addOwnerNote_(body.campaign, body.ownerName, body.coachName, body.coachEmail, body.text);
        break;
      case 'deleteOwnerNote':
        result = deleteOwnerNote_(body.noteId);
        break;
      case 'odCheckUser':
        result = odCheckUser(body.email || '');
        break;
      case 'odLogin':
        result = odLogin(body.email || '', body.pin || '', body.createPin || false);
        break;
      case 'odSaveMapping':
        result = odSaveMapping(body);
        break;
      case 'odSaveCampaignTabMap':
        result = odSaveCampaignTabMap(body);
        break;
      case 'odBatchSaveCampaignTabMap':
        result = odBatchSaveCampaignTabMap(body);
        break;
      case 'odBatchSaveMappings':
        result = odBatchSaveMappings(body);
        break;
      case 'odSaveUser':
        result = odSaveUser(body);
        break;
      case 'odDeleteUser':
        result = odDeleteUser(body);
        break;
      case 'odSavePlanning':
        result = odSavePlanning(body);
        break;
      case 'odFlagRep':
        result = odFlagRep(body);
        break;
      case 'odUnflagRep':
        result = odUnflagRep(body);
        break;
      case 'odSaveCampaignOwnership':
        result = odSaveCampaignOwnership(body);
        break;
      case 'odSaveAccessGrant':
        result = odSaveAccessGrant(body);
        break;
      case 'odDeleteAccessGrant':
        result = odDeleteAccessGrant(body);
        break;
      default:
        result = { error: 'unknown action: ' + body.action };
    }

    // ── Cache-bust after writes that affect cached data (supports chunked keys) ──
    try {
      var _cache = CacheService.getScriptCache();
      var _bustKeys = [];
      switch (body.action) {
        case 'claimCompany':
        case 'unclaimCompany':
        case 'claimCostSheet':
        case 'unclaimCostSheet':
          _bustKeys = ['cache_ownerCamMapping'];
          break;
        case 'updateHeadcount':
        case 'updateProduction':
        case 'saveGoals':
          var _ck = String(body.campaignKey || body.campaignLabel || '').toLowerCase().replace(/\s+/g, '-');
          if (_ck) _bustKeys.push('cache_recruiting_' + _ck);
          break;
        case 'refreshCampaigns':
          var _allKeys = Object.keys(typeof OD_CAMPAIGNS !== 'undefined' ? OD_CAMPAIGNS : {});
          for (var _b = 0; _b < _allKeys.length; _b++) _bustKeys.push('cache_recruiting_' + _allKeys[_b]);
          break;
        case 'refreshCampaign':
          if (body.campaign) _bustKeys.push('cache_recruiting_' + body.campaign);
          break;
      }
      if (_bustKeys.length) {
        for (var _r = 0; _r < _bustKeys.length; _r++) _cacheRemove(_cache, _bustKeys[_r]);
        Logger.log('Cache busted: ' + _bustKeys.join(', '));
      }
    } catch (e) { Logger.log('Cache bust error: ' + e.message); }

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
  // DEPRECATED — old Maddy's sheet import removed. Redirects to new pipeline.
  return refreshAllCampaigns();
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

  // 3. Weekly recruiting enrichment (deprecated — data now comes from consolidated tabs)
  try {
    // Skip — old Maddy's sheet no longer used
  } catch (err) {
    Logger.log('Weekly recruiting enrichment skipped: ' + err.message);
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
// PRODUCTION HELPERS: Internet/Wireless/DTV + Smart Goals
// Goals column uses "/" separator mapped left-to-right to
// whichever production categories the office tracks.
// E.g. if no Wireless data, "40/20" → Internet=40, DTV=20
// ══════════════════════════════════════════════════

/** Determine which production categories exist from colMap */
function _getActiveCategories(colMap) {
  var cats = [];
  if (colMap.internet >= 0) cats.push('internet');
  if (colMap.wireless >= 0) cats.push('wireless');
  if (colMap.dtv >= 0) cats.push('dtv');
  return cats;
}

/** Display names for production categories */
var _PROD_DISPLAY = { internet: 'Internet', wireless: 'Wireless', dtv: 'DTV' };

/**
 * Parse production actuals + goals into per-category object.
 * Uses { production, goals } keys to match frontend expectations.
 * @returns e.g. { Internet: { production: 15, goals: 40 }, DTV: { production: 8, goals: 20 } }
 */
function _parseProductionGoals(activeCategories, inet, wrls, dtvVal, goalsRaw) {
  var goalParts = String(goalsRaw || '').split('/').map(function(g) { return num(g.trim()); });
  var production = {};
  for (var i = 0; i < activeCategories.length; i++) {
    var cat = activeCategories[i];
    var actual = cat === 'internet' ? inet : cat === 'wireless' ? wrls : dtvVal;
    production[_PROD_DISPLAY[cat]] = {
      production: actual || 0,
      goals: i < goalParts.length ? goalParts[i] : 0
    };
  }
  return production;
}

/** Sum all "/" separated goal parts into a single total */
function _sumGoals(goalsRaw) {
  return String(goalsRaw || '').split('/').reduce(function(s, g) { return s + (num(g.trim()) || 0); }, 0);
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
    dist: findCol(headers, ['dist', 'distributors', 'distibutors']),
    training: findCol(headers, ['training']),
    internet: findCol(headers, ['internet']),
    wireless: findCol(headers, ['wireless']),
    dtv: findCol(headers, ['dtv']),
    goals: findCol(headers, ['goals', 'production goals'])
  };

  // Determine which production categories exist based on column headers
  var activeCategories = _getActiveCategories(colMap);

  var trend = [];
  var lastGoodRow = null;

  for (var i = start + 1; i <= end; i++) {
    var row = data[i];
    var dateVal = row[colMap.dates];
    if (!dateVal) continue;

    var inet = colMap.internet >= 0 ? val(row[colMap.internet]) : 0;
    var wrls = colMap.wireless >= 0 ? val(row[colMap.wireless]) : 0;
    var dtvVal = colMap.dtv >= 0 ? val(row[colMap.dtv]) : 0;
    var goalsRaw = colMap.goals >= 0 ? String(row[colMap.goals] || '') : '';
    var production = _parseProductionGoals(activeCategories, inet, wrls, dtvVal, goalsRaw);

    var entry = {
      date: formatDate(dateVal),
      active: val(row[colMap.active]),
      leaders: val(row[colMap.leaders]),
      dist: val(row[colMap.dist]),
      training: val(row[colMap.training]),
      internet: inet,
      wireless: wrls,
      dtv: dtvVal,
      goalsRaw: goalsRaw,
      production: production,
      productionLW: (inet || 0) + (wrls || 0) + (dtvVal || 0),
      goals: _sumGoals(goalsRaw)
    };
    trend.push(entry);
    lastGoodRow = entry;
  }

  return {
    current: lastGoodRow || { active: '—', leaders: '—', dist: '—', training: '—', internet: 0, wireless: 0, dtv: 0, production: {}, productionLW: 0, goals: 0 },
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
  // Strip common prefixes: "WE 1/3" → "1/3", "Week 2/28" → "2/28"
  var cleaned = name.replace(/^(WE|Week|Wk)\s+/i, '').trim();

  // Try formats: "3-7-2026", "03-02-26", "Mar-7", "2-28-2026", "1/3", "2/28"
  var parts = cleaned.split(/[-\/]/);
  if (parts.length >= 2) {
    var month = parseInt(parts[0]);
    var day = parseInt(parts[1]);
    var year = parts.length >= 3 ? parseInt(parts[2]) : 0;
    if (year < 100 && year > 0) year += 2000;

    // No year provided — infer from context
    if (year === 0) {
      var now = new Date();
      var curYear = now.getFullYear();
      // Try current year first; if that date is more than 1 week in the future,
      // assume it belongs to the previous year
      var candidate = new Date(curYear, month - 1, day);
      var oneWeekAhead = new Date(now.getTime() + 7 * 86400000);
      year = (candidate.getTime() > oneWeekAhead.getTime()) ? curYear - 1 : curYear;
    }

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

// ══════════════════════════════════════════════════
// D2D RESIDENTIAL RANKING — reads _D2D_Res_Ranking tab
// Returns { ranking: [ { rank, owner, totalUnits, products: { AIR, NEW INTERNET, ... } } ] }
// ══════════════════════════════════════════════════

function readD2DResRanking() {
  var ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var sheet = ss.getSheetByName('AT&T Res Metrics');
  if (!sheet) return { ranking: [], error: 'AT&T Res Metrics tab not found' };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ranking: [] };

  var headers = data[0].map(function(h) { return String(h).trim(); });

  var nameIdx = headers.indexOf('Name');
  var niIdx = headers.indexOf('New Internet');
  var upgIdx = headers.indexOf('Upgrade Internet');
  var wirelessIdx = headers.indexOf('Wireless');
  var dtvIdx = headers.indexOf('DTV');

  if (nameIdx === -1) return { ranking: [], error: 'Name column not found' };

  // Collect owner rows only (non-indented names = owner summary rows)
  var owners = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var name = String(row[nameIdx] || '');
    // Skip rep rows (indented with leading spaces)
    if (name.charAt(0) === ' ') continue;
    name = name.trim();
    if (!name) continue;

    var ni = niIdx !== -1 ? (Number(row[niIdx]) || 0) : 0;
    var upg = upgIdx !== -1 ? (Number(row[upgIdx]) || 0) : 0;
    var wireless = wirelessIdx !== -1 ? (Number(row[wirelessIdx]) || 0) : 0;
    var dtv = dtvIdx !== -1 ? (Number(row[dtvIdx]) || 0) : 0;
    var totalUnits = ni + upg + wireless + dtv;

    owners.push({
      owner: name,
      totalUnits: totalUnits,
      products: {
        'NEW INTERNET': ni,
        'UPGRADE INTERNET': upg,
        'WIRELESS': wireless,
        'VIDEO': dtv
      }
    });
  }

  // Sort by total units descending and assign rank
  owners.sort(function(a, b) { return b.totalUnits - a.totalUnits; });
  for (var i = 0; i < owners.length; i++) {
    owners[i].rank = i + 1;
  }

  return { ranking: owners };
}

function readOwnerCamMapping() {
  var ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var mapping = {};
  var costSheets = {};

  // Try legacy _OwnerCamMapping tab first
  var sheet = ss.getSheetByName('_OwnerCamMapping');
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var ownerName = String(data[i][0] || '').trim();
      var camCompany = String(data[i][1] || '').trim();
      var costSheetId = String(data[i][2] || '').trim();
      if (!ownerName) continue;
      if (camCompany) {
        if (!mapping[ownerName]) mapping[ownerName] = [];
        mapping[ownerName].push(camCompany);
      }
      if (costSheetId && !costSheets[ownerName]) {
        var idMatch = costSheetId.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
        costSheets[ownerName] = idMatch ? idMatch[1] : costSheetId;
      }
    }
  }

  // Also read camCompany from _OD_Mappings tab
  var odSheet = ss.getSheetByName('_OD_Mappings');
  if (odSheet) {
    var odData = odSheet.getDataRange().getValues();
    if (odData.length >= 2) {
      var odHeaders = odData[0].map(function(h) { return String(h).toLowerCase().trim(); });
      var colOwner = findCol(odHeaders, ['ownername', 'owner']);
      var colCam = findCol(odHeaders, ['camcompany', 'cam company']);
      if (colOwner >= 0 && colCam >= 0) {
        for (var j = 1; j < odData.length; j++) {
          var oName = String(odData[j][colOwner] || '').trim();
          var cName = String(odData[j][colCam] || '').trim();
          if (!oName || !cName) continue;
          // Only add if not already mapped from _OwnerCamMapping
          if (!mapping[oName]) mapping[oName] = [];
          if (mapping[oName].indexOf(cName) < 0) {
            mapping[oName].push(cName);
          }
        }
      }
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

// ── Helper: num with decimal preservation ──
function numDec(v) {
  if (v === null || v === undefined || v === '') return 0;
  var n = Number(v);
  return isNaN(n) ? 0 : Math.round(n * 100) / 100;
}




// ══════════════════════════════════════════════════
// OWNER NLR DATA — lazy load single owner from mapped NLR workbook
// Reads _OD_Mappings to find owner's nlrWorkbookId + nlrTab,
// opens just that one spreadsheet/tab, extracts health data.
// Called when user clicks on an individual owner.
// ══════════════════════════════════════════════════

function readOwnerNlrData(ownerName, campaignFilter) {
  if (!ownerName) return { error: 'ownerName required' };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message };
  }

  // Find this owner's mapping in _OD_Mappings
  var mapTab = ss.getSheetByName('_OD_Mappings');
  if (!mapTab) return { error: '_OD_Mappings not found' };
  var mapData = mapTab.getDataRange().getValues();
  if (mapData.length < 2) return { error: '_OD_Mappings empty' };

  var mapHeaders = mapData[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var colCampaign    = findCol(mapHeaders, ['campaign']);
  var colOwnerName   = findCol(mapHeaders, ['ownername', 'owner']);
  var colWorkbookId  = findCol(mapHeaders, ['nlrworkbookid']);
  var colNlrTab      = findCol(mapHeaders, ['nlrtab']);

  if (colOwnerName < 0 || colWorkbookId < 0 || colNlrTab < 0) {
    return { error: 'Missing columns in _OD_Mappings' };
  }

  // Find matching row (case-insensitive)
  var targetLower = ownerName.toLowerCase();
  var workbookId = '', tabName = '';
  for (var i = 1; i < mapData.length; i++) {
    var row = mapData[i];
    var name = String(row[colOwnerName] || '').trim();
    var campaign = String(row[colCampaign] || '').trim().toLowerCase();
    if (name.toLowerCase() !== targetLower) continue;
    if (campaignFilter && campaign !== campaignFilter.toLowerCase()) continue;
    workbookId = String(row[colWorkbookId] || '').trim();
    tabName = String(row[colNlrTab] || '').trim();
    break;
  }

  if (!workbookId || !tabName) {
    return { mapped: false, owner: ownerName };
  }

  // Open the mapped workbook and read the mapped tab directly
  try {
    var wb = SpreadsheetApp.openById(workbookId);
    var tab = wb.getSheetByName(tabName);
    if (!tab) return { error: 'Tab "' + tabName + '" not found in workbook' };

    var data = tab.getDataRange().getValues();
    if (data.length < 2) return { mapped: true, owner: ownerName, trend: [] };

    // Parse weekly sections: "WEEK OF X/XX" header, column headers, then data rows
    var weeks = [];
    var currentWeek = null;
    var colHeaders = null;

    for (var i = 0; i < data.length; i++) {
      var cell0 = String(data[i][0] || '').trim();
      var cell0Lower = cell0.toLowerCase();

      // Detect week header rows
      var dateMatch = cell0.match(/(\d{1,2}\/\d{1,2}(?:\/\d{2,4})?)/);
      if (cell0Lower.indexOf('week') >= 0 && dateMatch) {
        if (currentWeek && currentWeek.rows.length > 0) weeks.push(currentWeek);
        currentWeek = { date: dateMatch[1], rows: [], colHeaders: null };
        colHeaders = null;
        continue;
      }

      // Detect column header row — any row where multiple cells look like text headers
      if (!colHeaders && currentWeek) {
        var textCells = 0;
        for (var tc = 0; tc < Math.min(data[i].length, 10); tc++) {
          var v = String(data[i][tc] || '').trim();
          if (v.length > 1 && isNaN(Number(v))) textCells++;
        }
        if (textCells >= 3) {
          // Dedup headers: first "total spend" keeps its name, subsequent get "__2", "__3" etc.
          // This prevents the rightmost running-total column from overwriting the per-ad spend.
          var rawHeaders = data[i].map(function(h) { return String(h).toLowerCase().trim().replace(/[:\s]+$/, ''); });
          var seen = {};
          colHeaders = rawHeaders.map(function(h) {
            if (!h) return h;
            if (seen[h]) {
              seen[h]++;
              return h + '__' + seen[h]; // e.g. "total spend__2"
            }
            seen[h] = 1;
            return h;
          });
          if (currentWeek) currentWeek.colHeaders = colHeaders;
          continue;
        }
      }

      // Data row — check for any meaningful data, not just col A (merged cells leave it blank)
      if (currentWeek && colHeaders) {
        var hasData = false;
        for (var dc = 0; dc < Math.min(colHeaders.length, 10); dc++) {
          var dv = data[i][dc];
          if (dv !== null && dv !== undefined && String(dv).trim() !== '') { hasData = true; break; }
        }
        if (hasData) {
          var rowObj = {};
          for (var c = 0; c < colHeaders.length; c++) {
            if (colHeaders[c]) rowObj[colHeaders[c]] = data[i][c];
          }
          currentWeek.rows.push(rowObj);
        }
      }
    }
    if (currentWeek && currentWeek.rows.length > 0) weeks.push(currentWeek);

    // Merge same-date weeks (e.g. "WEEK 03/16" + "WEEK 03/16 (Phase 2)")
    var mergedMap = {};
    var mergedOrder = [];
    for (var mi = 0; mi < weeks.length; mi++) {
      var d = weeks[mi].date;
      if (mergedMap[d]) {
        mergedMap[d].rows = mergedMap[d].rows.concat(weeks[mi].rows);
        // Keep the first phase's colHeaders (column structure is the same)
      } else {
        mergedMap[d] = { date: d, rows: weeks[mi].rows.slice(), colHeaders: weeks[mi].colHeaders };
        mergedOrder.push(d);
      }
    }
    weeks = mergedOrder.map(function(d) { return mergedMap[d]; });

    // Aggregate each week + include individual ad rows for breakdown
    var trend = [];
    for (var w = 0; w < weeks.length; w++) {
      var wk = weeks[w];
      var summary = { date: wk.date, numAds: wk.rows.length, ads: [] };

      // Auto-detect numeric columns from first row (use per-week colHeaders)
      var wkHeaders = wk.colHeaders || colHeaders;
      if (wk.rows.length > 0 && wkHeaders) {
        for (var ci = 0; ci < wkHeaders.length; ci++) {
          var colName = wkHeaders[ci];
          if (!colName) continue;
          var colTotal = 0;
          var isNumeric = false;
          for (var ri = 0; ri < wk.rows.length; ri++) {
            var val = wk.rows[ri][colName];
            if (typeof val === 'number' || (typeof val === 'string' && !isNaN(Number(val)) && val.trim() !== '')) {
              colTotal += Number(val) || 0;
              isNumeric = true;
            }
          }
          if (isNumeric) {
            var key = colName.replace(/[^a-z0-9]+/g, ' ').trim().replace(/ ([a-z])/g, function(m, c) { return c.toUpperCase(); });
            summary[key] = Math.round(colTotal * 100) / 100;
          }
        }
      }

      // Override CPA/CPNS — these are ratios, not sums
      var wkSpend = summary.totalSpend || 0;
      var wkApplies = summary.applies || 0;
      var wkNS = summary.newStarts || 0;
      summary.cpa = wkApplies > 0 ? Math.round(wkSpend / wkApplies * 100) / 100 : 0;
      summary.cpns = wkNS > 0 ? Math.round(wkSpend / wkNS * 100) / 100 : 0;

      // Include individual ad rows for the Ad Breakdown table
      var lastAccount = '';
      for (var ai = 0; ai < wk.rows.length; ai++) {
        var r = wk.rows[ai];
        var adSpend = num(r['total spend']);
        var adApplies = num(r['applies']);
        var adNS = num(r['new starts']);
        var adTitle = String(r['ad title'] || '').trim();

        // Skip rows with no title and no spend (separator/empty rows)
        if (!adTitle && adSpend === 0 && adApplies === 0) continue;

        // Carry forward account name from merged cells
        var account = String(r['indeed account'] || '').trim();
        if (account) { lastAccount = account; } else { account = lastAccount; }

        summary.ads.push({
          account:   account,
          adTitle:   adTitle,
          location:  String(r['current location'] || r['location'] || '').trim(),
          spend:     adSpend,
          applies:   adApplies,
          seconds:   num(r['2nds']),
          newStarts: adNS,
          cpa:       adApplies > 0 ? Math.round(adSpend / adApplies * 100) / 100 : 0,
          cpns:      adNS > 0 ? Math.round(adSpend / adNS * 100) / 100 : 0,
          plan:      String(r['plan'] || r['action'] || '').trim()
        });
      }
      trend.push(summary);
    }

    return { mapped: true, owner: ownerName, trend: trend };
  } catch (err) {
    return { error: 'Failed to read NLR workbook: ' + err.message };
  }
}

// ══════════════════════════════════════════════════
// UPDATE HEADCOUNT ROW in consolidated campaign tab
// Finds row by Owner + Date, updates Active HC / Leaders / Dist / Training
// ══════════════════════════════════════════════════

function updateHeadcountRow(ownerName, date, active, leaders, dist, training, campaignLabel) {
  if (!ownerName) return { error: 'ownerName is required' };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message };
  }

  // Write to consolidated campaign tab (e.g. "Frontier", "AT&T NDS/Verizon")
  if (campaignLabel) {
    var sheet = ss.getSheetByName(campaignLabel);
    if (!sheet) return { error: 'Tab "' + campaignLabel + '" not found' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var colMap = {};
    for (var c = 0; c < headers.length; c++) colMap[headers[c]] = c;

    var colWeek = colMap['Week'];
    var colOwner = colMap['Owner'];
    var colActive = colMap['Active HC'];
    var colLeaders = colMap['Leaders'];
    var colDist = colMap['Dist'];
    var colTraining = colMap['Training'];

    if (colOwner === undefined) return { error: 'Owner column not found in "' + campaignLabel + '"' };

    // Find the row for this owner matching the requested date
    // If date is provided, match owner+date; otherwise fall back to most recent row
    var targetRow = -1;
    var targetDate = '';
    var fallbackRow = -1;
    var fallbackDate = '';
    var normDate = date ? _normalizeDate(date) : null;
    for (var i = 1; i < data.length; i++) {
      var rowOwner = String(data[i][colOwner] || '').trim();
      if (rowOwner.toLowerCase() !== ownerName.toLowerCase()) continue;
      var rowDate = colWeek !== undefined ? data[i][colWeek] : null;
      // If date requested, try exact match first
      if (normDate && rowDate) {
        var rowNorm = _normalizeDate(rowDate);
        if (rowNorm === normDate) {
          targetRow = i + 1;
          targetDate = rowDate;
          break;
        }
      }
      // Track first (most recent) row as fallback
      if (fallbackRow === -1) {
        fallbackRow = i + 1;
        fallbackDate = rowDate;
      }
    }
    // If a specific date was requested but not found, create a new row instead of falling back
    if (targetRow === -1 && normDate) {
      // Insert new row after header
      var numCols = headers.length;
      var newRow = [];
      for (var nr = 0; nr < numCols; nr++) newRow.push(0);
      // Set date and owner
      var dateParts = normDate.split('/');
      if (colWeek !== undefined) newRow[colWeek] = new Date(Number(dateParts[2]), Number(dateParts[0]) - 1, Number(dateParts[1]), 12, 0, 0);
      if (colOwner !== undefined) newRow[colOwner] = ownerName;
      if (colActive !== undefined) newRow[colActive] = parseInt(active) || 0;
      if (colLeaders !== undefined) newRow[colLeaders] = parseInt(leaders) || 0;
      if (colDist !== undefined) newRow[colDist] = parseInt(dist) || 0;
      if (colTraining !== undefined) newRow[colTraining] = parseInt(training) || 0;
      sheet.insertRowAfter(1);
      sheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);
      return { ok: true, row: 2, owner: ownerName, date: normDate, tab: campaignLabel, created: true };
    }

    // Use fallback if no date was specified
    if (targetRow === -1 && fallbackRow !== -1) {
      targetRow = fallbackRow;
      targetDate = fallbackDate;
    }

    if (targetRow === -1) {
      return { error: 'No row found for owner "' + ownerName + '" in "' + campaignLabel + '"' };
    }

    // Update headcount columns
    if (colActive !== undefined) sheet.getRange(targetRow, colActive + 1).setValue(parseInt(active) || 0);
    if (colLeaders !== undefined) sheet.getRange(targetRow, colLeaders + 1).setValue(parseInt(leaders) || 0);
    if (colDist !== undefined) sheet.getRange(targetRow, colDist + 1).setValue(parseInt(dist) || 0);
    if (colTraining !== undefined) sheet.getRange(targetRow, colTraining + 1).setValue(parseInt(training) || 0);

    return { ok: true, row: targetRow, owner: ownerName, date: formatDate(targetDate), tab: campaignLabel };
  }

  // campaignLabel is required — no fallback
  return { error: 'campaignLabel is required to update headcount' };
}

/**
 * Save goals for next week into the consolidated campaign tab.
 * Finds or creates a row for next week's Monday with the owner name,
 * then writes goal values into the Goal: <Product> columns.
 * @param {string} ownerName
 * @param {string} campaignLabel - e.g. "Frontier", "AT&T NDS/Verizon"
 * @param {string} campaignKey - e.g. "frontier", "att-nds"
 * @param {Object} goals - { productName: goalValue, ... }
 */
function saveGoalsRow_(ownerName, campaignLabel, campaignKey, goals) {
  if (!ownerName || !campaignLabel || !goals) return { error: 'Missing required params' };

  var ss;
  try { ss = SpreadsheetApp.openById(SHEETS.NATIONAL); }
  catch (e) { return { error: 'Cannot open sheet: ' + e.message }; }

  var sheet = ss.getSheetByName(campaignLabel);
  if (!sheet) return { error: 'Tab "' + campaignLabel + '" not found' };

  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var colMap = {};
  for (var c = 0; c < headers.length; c++) colMap[headers[c]] = c;

  var colWeek = colMap['Week'];
  var colOwner = colMap['Owner'];
  if (colOwner === undefined) return { error: 'Owner column not found' };

  // Goals go on the NEXT week's row. When looking at 3/22 data, the goal is for 3/29.
  // Step 1: Find the most recent past Sunday (the displayed week).
  // Step 2: Add 7 days to get the target week for goals.
  // Step 3: Find or create that row.
  var now = new Date();
  now.setHours(23, 59, 59, 999);
  var ownerLc = ownerName.toLowerCase();

  // Find the most recent lock day (<= today): Sunday for most campaigns, Monday for att-b2b
  var dayOfWeek = now.getDay(); // 0=Sun, 1=Mon, ...
  var goalTarget;
  if (campaignKey === 'att-b2b') {
    // Monday-lock: find most recent Monday, then add 7 for next Monday
    var monOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    var lastMonday = new Date(now.getFullYear(), now.getMonth(), now.getDate() + monOffset);
    goalTarget = new Date(lastMonday.getFullYear(), lastMonday.getMonth(), lastMonday.getDate() + 7);
  } else {
    var lastSunday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayOfWeek);
    goalTarget = new Date(lastSunday.getFullYear(), lastSunday.getMonth(), lastSunday.getDate() + 7);
  }
  var goalKey = _normalizeDateKey_(goalTarget, campaignKey);

  // Search for existing row matching owner + goal week
  var targetRow = -1;
  var targetDateKey = goalKey;
  for (var i = 1; i < data.length; i++) {
    var rowOwner = String(data[i][colOwner] || '').trim().toLowerCase();
    if (rowOwner !== ownerLc) continue;
    var rowDate = data[i][colWeek];
    var parsed = rowDate instanceof Date ? rowDate : _parseTabDate(String(rowDate));
    if (!parsed) continue;
    var rowKey = _normalizeDateKey_(parsed, campaignKey);
    if (rowKey === goalKey) {
      targetRow = i + 1; // 1-based
      break;
    }
  }

  // If no row exists for the goal week, create one
  if (targetRow === -1) {
    var numCols = headers.length;
    var newRow = [];
    for (var h2 = 0; h2 < numCols; h2++) newRow.push('');
    newRow[colWeek] = formatDate(goalTarget);
    newRow[colOwner] = ownerName;

    // Find insertion point: after the last row of the same week, or at top
    var insertAfter = -1;
    for (var i2 = 1; i2 < data.length; i2++) {
      var rd = data[i2][colWeek];
      var rk = _normalizeDateKey_(rd instanceof Date ? rd : _parseTabDate(String(rd)), campaignKey);
      if (rk === goalKey) insertAfter = i2 + 1;
    }

    if (insertAfter > 0) {
      sheet.insertRowAfter(insertAfter);
      sheet.getRange(insertAfter + 1, 1, 1, newRow.length).setValues([newRow]);
      targetRow = insertAfter + 1;
    } else {
      // No rows for this week — insert after the most recent week's rows
      var latestRowIdx = 1;
      for (var i3 = 1; i3 < data.length; i3++) {
        if (String(data[i3][colOwner] || '').trim()) latestRowIdx = i3 + 1;
      }
      sheet.insertRowAfter(latestRowIdx);
      sheet.getRange(latestRowIdx + 1, 1, 1, newRow.length).setValues([newRow]);
      targetRow = latestRowIdx + 1;
    }

    data = sheet.getDataRange().getValues();
    headers = data[0].map(function(h) { return String(h).trim(); });
    colMap = {};
    for (var c2 = 0; c2 < headers.length; c2++) colMap[headers[c2]] = c2;
  }

  // Write goals into Goal: <product> columns — ONLY if the cell is currently blank or 0.
  // Never overwrite existing non-zero data.
  var written = [];
  var skipped = [];
  for (var product in goals) {
    var goalCol = colMap['Goal: ' + product];
    if (goalCol === undefined) continue;
    var existing = sheet.getRange(targetRow, goalCol + 1).getValue();
    var existingStr = String(existing || '').trim();
    var existingNum = parseFloat(existingStr);
    if (existingStr !== '' && !isNaN(existingNum) && existingNum !== 0) {
      skipped.push(product + ' (existing: ' + existingStr + ')');
      continue;
    }
    sheet.getRange(targetRow, goalCol + 1).setValue(parseInt(goals[product]) || 0);
    written.push(product);
  }

  return { ok: true, row: targetRow, owner: ownerName, date: targetDateKey, written: written, skipped: skipped };
}

function updateProductionRow(ownerName, date, campaignLabel, products, internet, wireless, dtv, goals) {
  if (!ownerName) return { error: 'ownerName is required' };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message };
  }

  // Write to consolidated campaign tab (same pattern as updateHeadcountRow)
  if (campaignLabel) {
    var sheet = ss.getSheetByName(campaignLabel);
    if (!sheet) return { error: 'Tab "' + campaignLabel + '" not found' };

    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var colMap = {};
    for (var c = 0; c < headers.length; c++) colMap[headers[c]] = c;

    var colOwner = colMap['Owner'];
    if (colOwner === undefined) return { error: 'Owner column not found in "' + campaignLabel + '"' };

    // Find the row: if date provided, match owner+date; otherwise first match (most recent)
    var targetRow = -1;
    var targetDate = '';
    var normDate = date ? _normalizeDate(date) : null;
    for (var i = 1; i < data.length; i++) {
      var rowOwner = String(data[i][colOwner] || '').trim();
      if (rowOwner.toLowerCase() === ownerName.toLowerCase()) {
        if (normDate) {
          var colWeek = colMap['Week'];
          var rowDate = colWeek !== undefined ? _normalizeDate(data[i][colWeek]) : '';
          if (rowDate === normDate) {
            targetRow = i + 1;
            targetDate = data[i][colWeek];
            break;
          }
        } else {
          targetRow = i + 1;
          targetDate = colMap['Week'] !== undefined ? data[i][colMap['Week']] : '';
          break;
        }
      }
    }

    // After finding a date-matched row, check if a newer row exists for this owner.
    // If it does and has no production yet, redirect the write there — guards against
    // stale frontend dates writing to an older week when a newer row already exists.
    if (targetRow !== -1 && normDate) {
      var colWeekCheck = colMap['Week'];
      var targetDateMs = new Date(normDate).getTime();
      var newerRow = -1;
      var newerDateMs = targetDateMs;
      for (var i = 1; i < data.length; i++) {
        var ro = String(data[i][colOwner] || '').trim();
        if (ro.toLowerCase() !== ownerName.toLowerCase()) continue;
        var rd = colWeekCheck !== undefined ? _normalizeDate(data[i][colWeekCheck]) : '';
        if (!rd) continue;
        var rdMs = new Date(rd).getTime();
        if (rdMs > newerDateMs) { newerRow = i + 1; newerDateMs = rdMs; }
      }
      if (newerRow !== -1) {
        // A newer row exists — check its production
        var colProdCheck = colMap['Prod: Units'] !== undefined ? colMap['Prod: Units']
          : (colMap['Production LW'] !== undefined ? colMap['Production LW'] : colMap['Production']);
        if (colProdCheck === undefined) {
          for (var hk in colMap) { if (hk.indexOf('Prod: ') === 0) { colProdCheck = colMap[hk]; break; } }
        }
        var newerProd = colProdCheck !== undefined ? (parseInt(data[newerRow - 1][colProdCheck]) || 0) : 0;
        if (newerProd > 0) {
          // Newer row already has production — don't overwrite it, and don't write to the stale row either
          Logger.log('updateProductionRow: skipping "' + ownerName + '" — newer row at row ' + newerRow + ' already has prod ' + newerProd);
          return { ok: true, skipped: true, reason: 'newer-row-has-production', owner: ownerName };
        }
        // Newer row has no production — redirect the write there
        Logger.log('updateProductionRow: redirecting "' + ownerName + '" from stale row ' + targetRow + ' to newer row ' + newerRow);
        targetRow = newerRow;
        targetDate = data[newerRow - 1][colWeekCheck] || '';
      }
    }

    if (targetRow === -1) {
      // If no date was provided, fall back to the first row for this owner
      if (!normDate) {
        for (var i = 1; i < data.length; i++) {
          var rowOwner = String(data[i][colOwner] || '').trim();
          if (rowOwner.toLowerCase() === ownerName.toLowerCase()) {
            targetRow = i + 1;
            targetDate = colMap['Week'] !== undefined ? data[i][colMap['Week']] : '';
            break;
          }
        }
      }
      // Date was provided but no matching row found, or no rows exist — append a new one
      if (targetRow === -1) {
        var newRow = new Array(headers.length).fill('');
        newRow[colOwner] = ownerName;
        var colWeek = colMap['Week'];
        if (colWeek !== undefined && normDate) newRow[colWeek] = date;
        sheet.appendRow(newRow);
        // Re-read to get the new row index
        data = sheet.getDataRange().getValues();
        targetRow = data.length; // last row
        targetDate = normDate || '';
        Logger.log('updateProductionRow: no existing row for "' + ownerName + '" — appended new row at ' + targetRow);
      }
    }

    // Write per-product values using "Prod: <Product>" and "Goal: <Product>" columns
    var written = [];
    if (products && typeof products === 'object') {
      for (var pName in products) {
        var prodCol = colMap['Prod: ' + pName];
        var goalCol = colMap['Goal: ' + pName];
        if (prodCol !== undefined) {
          sheet.getRange(targetRow, prodCol + 1).setValue(parseInt(products[pName].actual) || 0);
          written.push('Prod: ' + pName + '=' + (products[pName].actual || 0));
        }
        if (goalCol !== undefined && products[pName].goal) {
          sheet.getRange(targetRow, goalCol + 1).setValue(parseInt(products[pName].goal) || 0);
          written.push('Goal: ' + pName + '=' + (products[pName].goal || 0));
        }
      }
    } else {
      // Legacy fallback: single Production LW / Production Goals columns
      var colProd = colMap['Production LW'] !== undefined ? colMap['Production LW'] : colMap['Production'];
      // Further fallback: any 'Prod: *' column (e.g. att-b2b uses 'Prod: Units')
      if (colProd === undefined) {
        for (var hk in colMap) {
          if (hk.indexOf('Prod: ') === 0) { colProd = colMap[hk]; break; }
        }
      }
      var colGoals = colMap['Production Goals'] !== undefined ? colMap['Production Goals'] : colMap['Goals'];
      if (colGoals === undefined) {
        for (var hk in colMap) {
          if (hk.indexOf('Goal: ') === 0) { colGoals = colMap[hk]; break; }
        }
      }
      if (colProd !== undefined) sheet.getRange(targetRow, colProd + 1).setValue(parseInt(internet) || 0);
      if (colGoals !== undefined) sheet.getRange(targetRow, colGoals + 1).setValue(parseInt(goals) || 0);
      written.push('prod=' + (internet || 0) + ' goals=' + (goals || 0));
    }

    return { ok: true, row: targetRow, owner: ownerName, date: formatDate(targetDate), tab: campaignLabel, written: written };
  }

  return { error: 'campaignLabel is required' };
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
  },
  'Samih Poles': {
    sheetId: '1u2iM7gfEGLUtxog5nxOLpjwCmJhWVsF5aT7_l87SCCg',
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

// NDS data sources: One-on-Ones sheet + NLR NDS report
// Each entry: { sheetId, tabs } — tabs is array of tab names to read
// Used by readNDSProduction() below
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




// ══════════════════════════════════════════════════
// READ NDS PRODUCTION/SALES from NDS One-on-Ones sheet
// Finds "Rep" in column A, reads headers + data rows.
// Returns { owners: { "Tab Name": { summary: {...}, reps: [{...}] } } }
// ══════════════════════════════════════════════════

function readNDSProduction() {
  if (!SHEETS.NDS_ONE_ON_ONES) return { owners: {} };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NDS_ONE_ON_ONES);
  } catch (err) {
    return { error: 'Cannot open NDS One-on-Ones sheet: ' + err.message, owners: {} };
  }

  var owners = {};

  // Use the same owner tabs as NDS_SOURCES[0] (One-on-Ones sheet)
  var ownerTabs = NDS_SOURCES[0].tabs || [];
  for (var t = 0; t < ownerTabs.length; t++) {
    var tabName = ownerTabs[t];
    var sheet = ss.getSheetByName(tabName);
    if (!sheet) continue;

    var lastRow = sheet.getLastRow();
    var lastCol = Math.min(sheet.getLastColumn(), 20);
    if (lastRow < 10 || lastCol < 5) continue;

    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();

    // Find "Rep" header row in column A
    var headersRow = -1;
    for (var i = 30; i < data.length; i++) { // sales section is below row 30
      var cellA = String(data[i][0] || '').trim().toLowerCase();
      if (cellA === 'rep') {
        headersRow = i;
        break;
      }
    }
    if (headersRow < 0) continue;

    // Build column map from headers
    var hdrs = data[headersRow].map(function(h) { return String(h || '').trim().toLowerCase(); });

    var colIdx = function(patterns) {
      for (var p = 0; p < patterns.length; p++) {
        for (var c = 0; c < hdrs.length; c++) {
          if (hdrs[c].indexOf(patterns[p]) >= 0) return c;
        }
      }
      return -1;
    };

    var cols = {
      rep:              0, // column A
      newPorts:         colIdx(['new/ports', 'new ports', 'sold last week']),
      orderCount:       colIdx(['order count', 'tier bonus']),
      cancelFraudPct:   colIdx(['cancel fraud', 'cancel']),
      extraPremiumPct:  colIdx(['extra', 'premium']),
      nextUpPct:        colIdx(['next up']),
      abpPct:           colIdx(['abp']),
      byodPct:          colIdx(['byod']),
      newOfNewPortsPct: colIdx(['new %', 'new % of']),
      insurancePct:     colIdx(['insurance']),
      highMedCreditPct: colIdx(['high/med', 'credit rating']),
      awayFromDoorsPct: colIdx(['away from door']),
      before3pmPct:     colIdx(['before 3']),
      after730pmPct:    colIdx(['after 7'])
    };

    // Read data rows below headers
    var summary = null;
    var reps = [];

    for (var i = headersRow + 1; i < data.length; i++) {
      var row = data[i];
      var repName = String(row[cols.rep] || '').trim();
      if (!repName) break; // empty row = end of section

      var entry = {
        name:             repName,
        rep:              repName,
        newPorts:         cols.newPorts >= 0 ? num(row[cols.newPorts]) : 0,
        orderCount:       cols.orderCount >= 0 ? num(row[cols.orderCount]) : 0,
        cancelFraudPct:   cols.cancelFraudPct >= 0 ? numDec(row[cols.cancelFraudPct]) : 0,
        extraPremiumPct:  cols.extraPremiumPct >= 0 ? numDec(row[cols.extraPremiumPct]) : 0,
        nextUpPct:        cols.nextUpPct >= 0 ? numDec(row[cols.nextUpPct]) : 0,
        abpPct:           cols.abpPct >= 0 ? numDec(row[cols.abpPct]) : 0,
        byodPct:          cols.byodPct >= 0 ? numDec(row[cols.byodPct]) : 0,
        newOfNewPortsPct: cols.newOfNewPortsPct >= 0 ? numDec(row[cols.newOfNewPortsPct]) : 0,
        insurancePct:     cols.insurancePct >= 0 ? numDec(row[cols.insurancePct]) : 0,
        highMedCreditPct: cols.highMedCreditPct >= 0 ? numDec(row[cols.highMedCreditPct]) : 0,
        awayFromDoorsPct: cols.awayFromDoorsPct >= 0 ? numDec(row[cols.awayFromDoorsPct]) : 0,
        before3pmPct:     cols.before3pmPct >= 0 ? numDec(row[cols.before3pmPct]) : 0,
        after730pmPct:    cols.after730pmPct >= 0 ? numDec(row[cols.after730pmPct]) : 0,
        // Map to B2B-compatible fields
        totalVolume:      cols.newPorts >= 0 ? num(row[cols.newPorts]) : 0,
        salesPerRep:      0,
        repCount:         0
      };

      if (repName.toLowerCase() === 'total') {
        summary = entry;
      } else {
        reps.push(entry);
      }
    }

    if (summary) {
      summary.repCount = reps.length;
    }

    if (summary || reps.length) {
      // Use canonical name (alias) if available
      var NDS_NAME_ALIASES = { 'Sam Poles': 'Samih Poles' };
      var ownerKey = NDS_NAME_ALIASES[tabName] || tabName;
      owners[ownerKey] = { summary: summary, reps: reps };
    }
  }

  return { owners: owners };
}

/**
 * Read sales data for a SINGLE NDS owner on demand.
 * Looks for the owner's tab in the NDS One-on-Ones sheet using fuzzy matching.
 * Returns: { summary: {...}, reps: [...] } or { error: '...' }
 */
function readNDSOwnerSales(ownerName) {
  if (!ownerName) return { error: 'No owner specified' };
  if (!SHEETS.NDS_ONE_ON_ONES) return { error: 'NDS sheet not configured' };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NDS_ONE_ON_ONES);
  } catch (err) {
    return { error: 'Cannot open NDS sheet: ' + err.message };
  }

  // Try exact match first, then fuzzy
  var ownerLower = ownerName.toLowerCase().trim();
  var allSheets = ss.getSheets();
  var tab = null;

  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName().toLowerCase().trim() === ownerLower) {
      tab = allSheets[i];
      break;
    }
  }

  if (!tab) {
    // Fuzzy: tab contains owner name or owner name contains tab name
    var allTabNames = allSheets.map(function(s) { return s.getName(); });
    var tabByName = {};
    allSheets.forEach(function(s) { tabByName[s.getName().toLowerCase().trim()] = s; });
    tab = _fuzzyFindTab_(ownerName, allTabNames, tabByName);
  }

  if (!tab) return { error: 'No tab found for: ' + ownerName };

  var lastRow = tab.getLastRow();
  var lastCol = Math.min(tab.getLastColumn(), 20);
  if (lastRow < 10 || lastCol < 5) return { error: 'Tab too small: ' + tab.getName() };

  var data = tab.getRange(1, 1, lastRow, lastCol).getValues();

  // Find "Rep" header row in column A (sales section is below row 30)
  var headersRow = -1;
  for (var i = 30; i < data.length; i++) {
    var cellA = String(data[i][0] || '').trim().toLowerCase();
    if (cellA === 'rep') {
      headersRow = i;
      break;
    }
  }
  if (headersRow < 0) return { error: 'No "Rep" header found in: ' + tab.getName() };

  var hdrs = data[headersRow].map(function(h) { return String(h || '').trim().toLowerCase(); });

  var colIdx = function(patterns) {
    for (var p = 0; p < patterns.length; p++) {
      for (var c = 0; c < hdrs.length; c++) {
        if (hdrs[c].indexOf(patterns[p]) >= 0) return c;
      }
    }
    return -1;
  };

  var cols = {
    rep:              0,
    newPorts:         colIdx(['new/ports', 'new ports', 'sold last week']),
    orderCount:       colIdx(['order count', 'tier bonus']),
    cancelFraudPct:   colIdx(['cancel fraud', 'cancel']),
    extraPremiumPct:  colIdx(['extra', 'premium']),
    nextUpPct:        colIdx(['next up']),
    abpPct:           colIdx(['abp']),
    byodPct:          colIdx(['byod']),
    newOfNewPortsPct: colIdx(['new %', 'new % of']),
    insurancePct:     colIdx(['insurance']),
    highMedCreditPct: colIdx(['high/med', 'credit rating']),
    awayFromDoorsPct: colIdx(['away from door']),
    before3pmPct:     colIdx(['before 3']),
    after730pmPct:    colIdx(['after 7'])
  };

  var summary = null;
  var reps = [];

  for (var i = headersRow + 1; i < data.length; i++) {
    var row = data[i];
    var repName = String(row[cols.rep] || '').trim();
    if (!repName) break;

    var entry = {
      name:             repName,
      rep:              repName,
      newPorts:         cols.newPorts >= 0 ? num(row[cols.newPorts]) : 0,
      orderCount:       cols.orderCount >= 0 ? num(row[cols.orderCount]) : 0,
      cancelFraudPct:   cols.cancelFraudPct >= 0 ? numDec(row[cols.cancelFraudPct]) : 0,
      extraPremiumPct:  cols.extraPremiumPct >= 0 ? numDec(row[cols.extraPremiumPct]) : 0,
      nextUpPct:        cols.nextUpPct >= 0 ? numDec(row[cols.nextUpPct]) : 0,
      abpPct:           cols.abpPct >= 0 ? numDec(row[cols.abpPct]) : 0,
      byodPct:          cols.byodPct >= 0 ? numDec(row[cols.byodPct]) : 0,
      newOfNewPortsPct: cols.newOfNewPortsPct >= 0 ? numDec(row[cols.newOfNewPortsPct]) : 0,
      insurancePct:     cols.insurancePct >= 0 ? numDec(row[cols.insurancePct]) : 0,
      highMedCreditPct: cols.highMedCreditPct >= 0 ? numDec(row[cols.highMedCreditPct]) : 0,
      awayFromDoorsPct: cols.awayFromDoorsPct >= 0 ? numDec(row[cols.awayFromDoorsPct]) : 0,
      before3pmPct:     cols.before3pmPct >= 0 ? numDec(row[cols.before3pmPct]) : 0,
      after730pmPct:    cols.after730pmPct >= 0 ? numDec(row[cols.after730pmPct]) : 0,
      totalVolume:      cols.newPorts >= 0 ? num(row[cols.newPorts]) : 0,
      salesPerRep:      0,
      repCount:         0
    };

    if (repName.toLowerCase() === 'total') {
      summary = entry;
    } else {
      reps.push(entry);
    }
  }

  if (summary) summary.repCount = reps.length;

  return { summary: summary, reps: reps, tab: tab.getName() };
}

// ── AT&T B2B per-owner sales (from Tableau sync "B2B Sales Metrics" tab) ──
function readB2BOwnerSales(ownerName) {
  if (!ownerName) return { error: 'No owner specified' };

  var ss;
  try { ss = SpreadsheetApp.openById(SHEETS.NATIONAL); }
  catch (err) { return { error: 'Cannot open National sheet: ' + err.message }; }

  var tab = ss.getSheetByName('B2B Sales Metrics');
  if (!tab) return { error: 'B2B Sales Metrics tab not found' };

  var data = tab.getDataRange().getValues();
  if (data.length < 2) return { error: 'No data in B2B Sales Metrics' };

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var colMap = {};
  for (var c = 0; c < headers.length; c++) colMap[headers[c]] = c;

  // The tab format: owner rows (not indented) followed by indented rep rows ("  RepName")
  // Find the owner section by matching Name column
  // Supports fuzzy matching: exact → starts-with → contains (handles "Alex Badawi" ↔ "Alexander Badawi")
  var ownerLower = ownerName.toLowerCase().trim();
  var summary = null;
  var reps = [];
  var inOwnerSection = false;

  // First pass: collect all owner names for fuzzy matching
  var ownerRows = [];
  for (var i = 1; i < data.length; i++) {
    var rawName = String(data[i][colMap['Name'] || 0] || '');
    if (rawName.charAt(0) !== ' ' && rawName.trim()) {
      ownerRows.push({ idx: i, name: rawName.trim(), nameLc: rawName.trim().toLowerCase() });
    }
  }

  // Find best match: exact → contains → last name match
  var matchIdx = -1;
  for (var m = 0; m < ownerRows.length; m++) {
    if (ownerRows[m].nameLc === ownerLower) { matchIdx = ownerRows[m].idx; break; }
  }
  if (matchIdx < 0) {
    for (var m = 0; m < ownerRows.length; m++) {
      if (ownerRows[m].nameLc.indexOf(ownerLower) >= 0 || ownerLower.indexOf(ownerRows[m].nameLc) >= 0) { matchIdx = ownerRows[m].idx; break; }
    }
  }
  // Last name match: "Alex Badawi" ↔ "Alexander Badawi" (same last name)
  if (matchIdx < 0) {
    var ownerParts = ownerLower.split(/\s+/);
    var ownerLast = ownerParts.length > 1 ? ownerParts[ownerParts.length - 1] : '';
    if (ownerLast) {
      for (var m = 0; m < ownerRows.length; m++) {
        var rowParts = ownerRows[m].nameLc.split(/\s+/);
        var rowLast = rowParts.length > 1 ? rowParts[rowParts.length - 1] : '';
        if (rowLast && rowLast === ownerLast) { matchIdx = ownerRows[m].idx; break; }
      }
    }
  }

  for (var i = 1; i < data.length; i++) {
    var rawName = String(data[i][colMap['Name'] || 0] || '');
    var cleanName = rawName.trim();
    if (!cleanName) continue;

    var isIndented = rawName.charAt(0) === ' ';

    if (!isIndented) {
      // Owner row — check if this is our target
      if (inOwnerSection) break; // We were in our section, now hit next owner — done
      if (i === matchIdx) {
        inOwnerSection = true;
        summary = _parseB2BRow_(data[i], colMap, cleanName);
      }
      continue;
    }

    // Indented rep row — only collect if we're in our owner's section
    if (inOwnerSection) {
      reps.push(_parseB2BRow_(data[i], colMap, cleanName));
    }
  }

  // ── MCOE Sales — extracted unconditionally, before any early returns ──
  // Must happen here so owners with no Tableau data still get their MCOE card.
  var mcoeSales = null;
  var srcSS = null;
  try {
    srcSS = SpreadsheetApp.openById(OD_CAMPAIGNS['att-b2b'].sheetId);
    mcoeSales = _extractMcoeSales_(srcSS, ownerName);
    Logger.log('readB2BOwnerSales: MCOE for "' + ownerName + '": ' + (mcoeSales ? mcoeSales.length + ' rows' : 'null'));
  } catch (e) {
    Logger.log('readB2BOwnerSales: MCOE read failed for "' + ownerName + '": ' + e.message);
  }

  if (!summary) {
    Logger.log('readB2BOwnerSales: no Tableau match for "' + ownerName + '" — available owners: ' + ownerRows.map(function(r) { return r.name; }).join(', '));
    // Still return MCOE if we have it — don't block the card on Tableau data
    if (mcoeSales) {
      return { summary: null, reps: [], dailyActivity: [], marketFulfillment: null, avgPaychecks: null, mcoeSales: mcoeSales, tab: 'B2B Sales Metrics' };
    }
    return { error: 'Owner not found: ' + ownerName };
  }
  summary.repCount = reps.length;

  // Read daily activity data
  var dailyActivity = [];
  var dailyTab = ss.getSheetByName('B2B Daily Activity');
  if (dailyTab && dailyTab.getLastRow() > 1) {
    var dailyData = dailyTab.getDataRange().getValues();
    var dHeaders = dailyData[0].map(function(h) { return String(h).trim(); });
    var dOwnerCol = dHeaders.indexOf('Owner');
    var dDayCol = dHeaders.indexOf('Day');
    var dFirstCol = dHeaders.indexOf('First Order');
    var dLastCol = dHeaders.indexOf('Last Order');
    var dOrdersCol = dHeaders.indexOf('Orders');

    for (var di = 1; di < dailyData.length; di++) {
      var dOwner = String(dailyData[di][dOwnerCol] || '').trim();
      if (dOwner.toLowerCase() !== ownerLower) continue;
      dailyActivity.push({
        day: String(dailyData[di][dDayCol] || '').trim(),
        firstOrder: String(dailyData[di][dFirstCol] || '').trim(),
        lastOrder: String(dailyData[di][dLastCol] || '').trim(),
        orders: parseInt(dailyData[di][dOrdersCol]) || 0
      });
    }
  }

  // Read Market Fulfillment and Average Paychecks from owner's source tab
  var marketFulfillment = null;
  var avgPaychecks = null;
  try {
    if (!srcSS) srcSS = SpreadsheetApp.openById(OD_CAMPAIGNS['att-b2b'].sheetId);
    var ownerTab = _findOwnerTab_(srcSS, ownerName);
    if (ownerTab) {
      var range = ownerTab.getDataRange();
      var srcData = range.getValues();
      var srcDisplay = range.getDisplayValues();
      marketFulfillment = _extractMarketFulfillment_(srcData, ownerName, srcDisplay);
      avgPaychecks = _extractAvgPaychecks_(srcData, srcDisplay);
    }
  } catch (e) {
    Logger.log('readB2BOwnerSales: source tab read failed: ' + e.message);
  }

  return {
    summary: summary, reps: reps, dailyActivity: dailyActivity,
    marketFulfillment: marketFulfillment, avgPaychecks: avgPaychecks,
    mcoeSales: mcoeSales,
    tab: 'B2B Sales Metrics'
  };
}

/**
 * Extract MCOE Sales rows for an owner from the "INPUT - MCOE Weekly Sales" tab.
 * Returns an array of { office, aiaCnt, byodCnt, phoneCnt, totalLines } objects,
 * one per matching row (owner may have multiple ICD offices).
 * Skips the "Grand Total" row.
 */
function _extractMcoeSales_(srcSS, ownerName) {
  var tab = srcSS.getSheetByName('INPUT - MCOE Weekly Sales')
    || srcSS.getSheetByName(' INPUT - MCOE Weekly Sales');
  if (!tab) { Logger.log('_extractMcoeSales_: tab "INPUT - MCOE Weekly Sales" not found'); return null; }
  if (tab.getLastRow() < 2) { Logger.log('_extractMcoeSales_: tab is empty'); return null; }

  var data = tab.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  Logger.log('_extractMcoeSales_: headers = ' + JSON.stringify(headers));

  var nameCol    = headers.indexOf('Name (ICD Office)');
  var aiaCol     = headers.indexOf('AIA Cnt');
  var byodCol    = headers.indexOf('BYOD Cnt');
  var phoneCol   = headers.indexOf('Phone Cnt');
  var totalCol   = headers.indexOf('Total Lines Count');

  if (nameCol < 0) {
    Logger.log('_extractMcoeSales_: "Name (ICD Office)" column not found in headers: ' + JSON.stringify(headers));
    return null;
  }

  var ownerLower = ownerName.toLowerCase().trim();
  var ownerParts = ownerLower.split(/\s+/);
  var ownerLast  = ownerParts.length > 1 ? ownerParts[ownerParts.length - 1] : '';

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var cellName = String(data[i][nameCol] || '').trim();
    if (!cellName) continue;
    if (cellName.toLowerCase() === 'grand total') continue;

    // Match: name cell contains the owner's full name or last name
    var cellLower = cellName.toLowerCase();
    var matched = cellLower === ownerLower
      || cellLower.indexOf(ownerLower) >= 0
      || ownerLower.indexOf(cellLower) >= 0
      || (ownerLast && cellLower.indexOf(ownerLast) >= 0);

    if (!matched) continue;

    rows.push({
      office:     cellName,
      aiaCnt:     parseInt(data[i][aiaCol])   || 0,
      byodCnt:    parseInt(data[i][byodCol])  || 0,
      phoneCnt:   parseInt(data[i][phoneCol]) || 0,
      totalLines: parseInt(data[i][totalCol]) || 0
    });
  }

  return rows.length > 0 ? rows : null;
}

/** Format helpers for B2B source data */
function _fmtPct_(v) {
  var n = parseFloat(String(v).replace(/[%,]/g, ''));
  if (isNaN(n)) return v;
  if (n > 0 && n < 1) return (n * 100).toFixed(2) + '%';
  return n.toFixed(2) + '%';
}
function _fmtNum_(v) {
  var n = parseFloat(String(v).replace(/[$,]/g, ''));
  if (isNaN(n)) return v;
  return Math.round(n).toLocaleString();
}
function _fmtDollar_(v) {
  var n = parseFloat(String(v).replace(/[$,]/g, ''));
  if (isNaN(n)) return v;
  return '$' + Math.round(n).toLocaleString();
}
function _fmtShortDate_(v) {
  var s = String(v || '').trim();
  // Handle Date objects or full datetime strings → M/D/YYYY
  var d = new Date(s);
  if (!isNaN(d.getTime()) && d.getFullYear() > 1900) {
    return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
  }
  // Already short format
  return s.replace(/\/\d{4}$/, function(m) { return m; }).substring(0, 15);
}

/** Find an owner's tab in the source spreadsheet (case-insensitive) */
function _findOwnerTab_(ss, ownerName) {
  var ownerLower = ownerName.toLowerCase().trim();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase().trim() === ownerLower) return sheets[i];
  }
  // Fuzzy: tab contains owner name
  for (var i = 0; i < sheets.length; i++) {
    var tabLower = sheets[i].getName().toLowerCase().trim();
    if (tabLower.indexOf(ownerLower) >= 0 || ownerLower.indexOf(tabLower) >= 0) return sheets[i];
  }
  return null;
}

/** Extract Market Fulfillment section from owner tab data + display data */
function _extractMarketFulfillment_(data, ownerName, displayData) {
  // Find header row: "Owner (Owner>Rep>Zip)" in column A
  var headerRow = -1;
  for (var i = 0; i < data.length; i++) {
    var cell = String(data[i][0] || '').trim().toLowerCase();
    if (cell.indexOf('owner (owner>rep>zip)') >= 0 || cell.indexOf('owner(owner>rep>zip)') >= 0) {
      headerRow = i;
      break;
    }
  }
  if (headerRow < 0) return null;

  // Use display values for headers (dates show as "3/22/26" not serial numbers)
  var headers = (displayData && displayData[headerRow] ? displayData[headerRow] : data[headerRow])
    .map(function(h) { return String(h || '').trim(); });

  // Find column indices
  var colDMA = -1, colWorkable = -1, colTotal = -1, colPen = -1;
  var colWeeklyTotal = -1, colWeeklyCRU = -1;
  var dateCols = []; // { idx, label }

  for (var c = 0; c < headers.length; c++) {
    var h = headers[c].toLowerCase();
    if (h.indexOf('dma') >= 0) colDMA = c;
    else if (h.indexOf('total workable') >= 0) colWorkable = c;
    else if (h.indexOf('total') >= 0 && h.indexOf('workable') < 0 && h.indexOf('weekly') < 0 && colTotal < 0) colTotal = c;
    else if (h.indexOf('actual pen') >= 0 || h.indexOf('pen %') >= 0 || h.indexOf('% pen') >= 0 || h.indexOf('pen rate') >= 0 || (h.indexOf('penetration') >= 0 && h.indexOf('%') >= 0)) colPen = c;
    else if (h.indexOf('weekly total') >= 0) colWeeklyTotal = c;
    else if (h.indexOf('weekly cru') >= 0) colWeeklyCRU = c;
    else if (h.match(/^\d{1,2}\/\d{1,2}/)) dateCols.push({ idx: c, label: headers[c] });
    else if (h.indexOf('4 wk avg') >= 0 || h.indexOf('4wk avg') >= 0) dateCols.push({ idx: c, label: '4 Wk Avg' });
    else if (h.indexOf('last wk') >= 0 || h.indexOf('last week') >= 0) dateCols.push({ idx: c, label: 'Last Wk' });
  }

  // Read data rows until blank — use display values for pre-formatted output
  var markets = [];
  for (var i = headerRow + 1; i < data.length; i++) {
    var name = String(data[i][0] || '').trim();
    if (!name) break;
    var dRow = displayData && displayData[i] ? displayData[i] : data[i];
    var market = {
      dma: colDMA >= 0 ? String(dRow[colDMA] || '').trim() : '',
      totalWorkable: colWorkable >= 0 ? String(dRow[colWorkable] || '').trim() : '',
      total: colTotal >= 0 ? String(dRow[colTotal] || '').trim() : '',
      penRate: colPen >= 0 ? String(dRow[colPen] || '').trim() : '',
      weeklyTotal: colWeeklyTotal >= 0 ? String(dRow[colWeeklyTotal] || '').trim() : '',
      weeklyCRU: colWeeklyCRU >= 0 ? String(dRow[colWeeklyCRU] || '').trim() : '',
      weeks: []
    };
    for (var d = 0; d < dateCols.length; d++) {
      market.weeks.push({ label: dateCols[d].label, value: String(dRow[dateCols[d].idx] || '').trim() });
    }
    markets.push(market);
  }

  return markets.length > 0 ? markets : null;
}

/** Extract Average Paychecks section from owner tab data + display data */
function _extractAvgPaychecks_(data, displayData) {
  // Find "Comm per Rep" row
  var commRow = -1;
  for (var i = 0; i < data.length; i++) {
    var cell = String(data[i][0] || '').trim().toLowerCase();
    if (cell === 'comm per rep') { commRow = i; break; }
  }
  if (commRow < 0) return null;

  // Header row is next row (has cl.ICD_Owner_Name + dates)
  var headerRow = commRow + 1;
  if (headerRow >= data.length) return null;
  // Use display values for headers (dates as formatted strings)
  var headers = (displayData && displayData[headerRow] ? displayData[headerRow] : data[headerRow])
    .map(function(h) { return String(h || '').trim(); });

  // Data row is next — use display values for pre-formatted numbers
  var dataRow = headerRow + 1;
  if (dataRow >= data.length) return null;
  var row = displayData && displayData[dataRow] ? displayData[dataRow] : data[dataRow];

  // Parse Comm per Rep columns (columns 1 until empty gap or "Last 4 Wk Avg")
  var commDates = [];
  var commAvg = '';
  for (var c = 1; c < headers.length; c++) {
    var h = headers[c].toLowerCase().trim();
    if (!h) break; // gap = end of Comm section
    if (h.indexOf('last') >= 0 && h.indexOf('avg') >= 0) {
      commAvg = String(row[c] || '').trim();
      break;
    }
    if (headers[c]) {
      commDates.push({ label: headers[c], value: String(row[c] || '').trim() });
    }
  }

  // Find Total DD section
  var totalDDDates = [];
  var totalDDAvg = '';
  var ddStartCol = -1;
  for (var c = 0; c < data[commRow].length; c++) {
    if (String(data[commRow][c] || '').trim().toLowerCase() === 'total dd') {
      ddStartCol = c;
      break;
    }
  }
  if (ddStartCol >= 0) {
    for (var c = ddStartCol; c < headers.length; c++) {
      var h = headers[c].toLowerCase().trim();
      if (!h || h === 'total dd' || h === 'cl.icd_owner_name') continue;
      var val = String(row[c] || '').trim();
      if (h.indexOf('last') >= 0 && h.indexOf('avg') >= 0) {
        totalDDAvg = val;
        break;
      }
      if (headers[c]) {
        totalDDDates.push({ label: headers[c], value: val });
      }
    }
  }

  return {
    commPerRep: { weeks: commDates, avg: commAvg },
    totalDD: { weeks: totalDDDates, avg: totalDDAvg }
  };
}

function _parseB2BRow_(row, colMap, name) {
  function v(colName) {
    var idx = colMap[colName];
    if (idx === undefined) return 0;
    var val = row[idx];
    if (typeof val === 'string') {
      val = val.replace(/[%,]/g, '');
      return parseFloat(val) || 0;
    }
    return Number(val) || 0;
  }
  return {
    name: name,
    rep: name,
    totalVolume:  v('Total Volume'),
    repCount:     v('Rep Count'),
    salesPerRep:  v('Sales Per Rep'),
    orderCount:   v('Order Count'),
    ordersBefore: v('Before 12 PM'),
    earlyPct:     v('Early %') / 100,
    ordersAfter:  v('After 5 PM'),
    latePct:      v('Late %') / 100,
    internet:     v('Internet'),
    voip:         v('VOIP'),
    wireless:     v('Wireless'),
    airAwb:       v('AIR/AWB'),
    weekendPct:   v('Weekend Selling %') / 100,
    tierPct:      v('Tier Attainment %') / 100,
    abpPct:       v('ABP %') / 100,
    cruPct:       v('CRU %') / 100,
    newWrlsPct:   v('New Wireless %') / 100,
    byodPct:      v('BYOD %') / 100
  };
}

// ── Verizon FIOS per-owner sales (from Credico sync "Verizon Sales" tab) ──
function readFiosOwnerSales(ownerName) {
  if (!ownerName) return { error: 'No owner specified' };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message };
  }

  var sheet = ss.getSheetByName('Verizon Sales');
  if (!sheet) return { error: 'Verizon Sales tab not found' };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { error: 'No data in Verizon Sales' };

  // Headers: Name, Fios, TV, Orders, Gig %, Autobill %, Avg Days Past 1st Avail, Lines, Port %, NEW, CPO, BYOD, Scoring HC, Productive HC, Avg Units/Rep
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var colMap = {};
  for (var c = 0; c < headers.length; c++) colMap[headers[c]] = c;

  var ownerLower = ownerName.toLowerCase().trim();
  // Build "Last, First" version for matching Credico format
  var ownerParts = ownerLower.split(/\s+/);
  var ownerLastFirst = ownerParts.length >= 2
    ? ownerParts[ownerParts.length - 1] + ', ' + ownerParts.slice(0, -1).join(' ')
    : ownerLower;

  // Reverse alias lookup: consolidated name → Credico name
  var FIOS_ALIASES = {
    'abhriham potturu': ['potturu, abhriham', 'potturu, abhiram'],
    'gent ademaj': ['ademaj, gent'],
    'day dobson-diaz': ['dobson-diaz, day', "dobson-diaz, ja'darry-a"],
    'jp morrone': ['morrone, john philip'],
    'ricky madureira': ['madureira, richard']
  };
  var aliasNames = FIOS_ALIASES[ownerLower] || [];

  // Find the owner row (not indented) and collect rep rows (indented) below it
  var ownerRowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0] || '').trim();
    // Owner rows are not indented; skip if starts with space
    if (name.charAt(0) === ' ') continue;
    var nameLower = name.toLowerCase();
    if (nameLower === ownerLower || nameLower === ownerLastFirst || aliasNames.indexOf(nameLower) >= 0) {
      ownerRowIdx = i;
      break;
    }
  }

  if (ownerRowIdx < 0) return { error: 'Owner not found: ' + ownerName };

  var ownerRow = data[ownerRowIdx];
  var val = function(colName) {
    var idx = colMap[colName];
    if (idx === undefined) return 0;
    var v = ownerRow[idx];
    if (typeof v === 'string') v = v.replace('%', '');
    return Number(v) || 0;
  };

  var summary = {
    fios:         val('Fios'),
    tv:           val('TV'),
    orderCount:   val('Orders'),
    gigPct:       val('Gig %'),
    autobillPct:  val('Autobill %'),
    avgDaysPast:  val('Avg Days Past 1st Avail'),
    lines:        val('Lines'),
    portPct:      val('Port %'),
    newPhones:    val('NEW'),
    cpo:          val('CPO'),
    byod:         val('BYOD'),
    scoringHC:    val('Scoring HC'),
    productiveHC: val('Productive HC'),
    avgUnitsRep:  val('Avg Units/Rep'),
    totalVolume:  val('Fios') + val('TV'),
    repCount:     val('Scoring HC')
  };

  // Collect rep rows (indented, immediately after owner row)
  var reps = [];
  for (var r = ownerRowIdx + 1; r < data.length; r++) {
    var repName = String(data[r][0] || '');
    if (!repName || repName.charAt(0) !== ' ') break; // next owner or end
    repName = repName.trim();

    var repVal = function(colName) {
      var idx = colMap[colName];
      if (idx === undefined) return 0;
      var v = data[r][idx];
      if (typeof v === 'string') v = v.replace('%', '');
      return Number(v) || 0;
    };

    reps.push({
      name:         repName,
      fios:         repVal('Fios'),
      tv:           repVal('TV'),
      orderCount:   repVal('Orders'),
      gigPct:       repVal('Gig %'),
      autobillPct:  repVal('Autobill %'),
      avgDaysPast:  repVal('Avg Days Past 1st Avail'),
      lines:        repVal('Lines'),
      portPct:      repVal('Port %'),
      newPhones:    repVal('NEW'),
      cpo:          repVal('CPO'),
      byod:         repVal('BYOD'),
      totalUnits:   repVal('Fios') + repVal('TV') + repVal('Lines')
    });
  }

  return { summary: summary, reps: reps, tab: 'Verizon Sales' };
}

// ── AT&T Residential per-owner sales (from Tableau sync "AT&T Res Metrics" tab) ──
function readResOwnerSales(ownerName) {
  if (!ownerName) return { error: 'No owner specified' };

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { error: 'Cannot open national sheet: ' + err.message };
  }

  var sheet = ss.getSheetByName('AT&T Res Metrics');
  if (!sheet) return { error: 'AT&T Res Metrics tab not found' };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { error: 'No data in AT&T Res Metrics' };

  // Headers: Name, New Internet, Upgrade Internet, Wireless, DTV, ABP Mix %, 1Gig+ Mix %, Tech Install %, ...
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var colMap = {};
  for (var c = 0; c < headers.length; c++) colMap[headers[c]] = c;

  var ownerLower = ownerName.toLowerCase().trim();

  // Find the owner row (not indented)
  var ownerRowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0] || '').trim();
    if (name.charAt(0) === ' ') continue; // skip rep rows
    if (name.toLowerCase() === ownerLower) {
      ownerRowIdx = i;
      break;
    }
  }

  if (ownerRowIdx < 0) return { error: 'Owner not found: ' + ownerName };

  var ownerRow = data[ownerRowIdx];
  var val = function(colName) {
    var idx = colMap[colName];
    if (idx === undefined) return 0;
    var v = ownerRow[idx];
    if (typeof v === 'string') v = v.replace('%', '');
    return Number(v) || 0;
  };

  var summary = {
    newInternet:     val('New Internet'),
    upgradeInternet: val('Upgrade Internet'),
    wirelessSales:   val('Wireless'),
    videoSales:      val('DTV'),
    totalVolume:     val('New Internet') + val('Upgrade Internet') + val('Wireless') + val('DTV'),
    abpMix:          val('ABP Mix %'),
    gigMix:          val('1Gig+ Mix %'),
    techInstall:     val('Tech Install %'),
    jepNI:           val('Jep NI (4wk)'),
    pastDueNI:       val('Past Due NI (4wk)'),
    sched6Days:      val('Sched 6+ Days (4wk)'),
    sales30d:        val('0-30d NI Sales'),
    cancels30d:      val('0-30d NI Cancels'),
    activations30d:  val('0-30d NI Activations'),
    disconnects30d:  val('0-30d NI Disconnects'),
    churnRate30d:    val('0-30d Churn Rate'),
    actRate3060d:    val('30-60d Activation Rate')
  };
  summary.repCount = 0;

  // Collect rep rows (indented, immediately after owner row)
  var reps = [];
  for (var r = ownerRowIdx + 1; r < data.length; r++) {
    var repName = String(data[r][0] || '');
    if (!repName || repName.charAt(0) !== ' ') break;
    repName = repName.trim();

    var repVal = function(colName) {
      var idx = colMap[colName];
      if (idx === undefined) return 0;
      var v = data[r][idx];
      if (typeof v === 'string') v = v.replace('%', '');
      return Number(v) || 0;
    };

    reps.push({
      name:            repName,
      rep:             repName,
      newInternet:     repVal('New Internet'),
      upgradeInternet: repVal('Upgrade Internet'),
      wirelessSales:   repVal('Wireless'),
      videoSales:      repVal('DTV'),
      totalVolume:     repVal('New Internet') + repVal('Upgrade Internet') + repVal('Wireless') + repVal('DTV'),
      abpMix:          repVal('ABP Mix %'),
      gigMix:          repVal('1Gig+ Mix %'),
      techInstall:     repVal('Tech Install %'),
      jepNI:           repVal('Jep NI (4wk)'),
      pastDueNI:       repVal('Past Due NI (4wk)'),
      sched6Days:      repVal('Sched 6+ Days (4wk)'),
      sales30d:        repVal('0-30d NI Sales'),
      cancels30d:      repVal('0-30d NI Cancels'),
      activations30d:  repVal('0-30d NI Activations'),
      disconnects30d:  repVal('0-30d NI Disconnects'),
      churnRate30d:    repVal('0-30d Churn Rate'),
      actRate3060d:    repVal('30-60d Activation Rate')
    });
  }

  summary.repCount = reps.length;

  return { summary: summary, reps: reps, tab: 'AT&T Res Metrics' };
}

// ═══════════════════════════════════════════════════════
// OWNER DEVELOPMENT (OD) — Helper Functions
// ═══════════════════════════════════════════════════════

/**
 * Get or create a tab in the script's bound spreadsheet.
 * If the tab doesn't exist, creates it with the given headers in row 1.
 */
function odGetOrCreateTab(tabName, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sheet;
}

/**
 * SHA-256 hash a PIN string. Returns hex string.
 */
function odHashPin(pin) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin));
  return raw.map(function(b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

/**
 * Read all rows from a tab and return an array of objects keyed by header row.
 */
function odReadTab(tabName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // Use header row to determine column count (getDataRange can miss trailing empty cols)
  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var headerVals = headerRange.getValues()[0];
  var numCols = headerVals.length;
  // Extend to at least the last header with a value
  for (var hc = headerVals.length - 1; hc >= 0; hc--) {
    if (String(headerVals[hc]).trim()) { numCols = hc + 1; break; }
  }
  var data = sheet.getRange(1, 1, lastRow, numCols).getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      obj[headers[c]] = data[i][c] !== undefined ? data[i][c] : '';
    }
    rows.push(obj);
  }
  return rows;
}

// ═══════════════════════════════════════════════════════
// OWNER DEVELOPMENT (OD) — Campaign Ownership & Access Grants
// ═══════════════════════════════════════════════════════

var OD_OWNERSHIP_HEADERS_ = ['campaign', 'ownerEmail', 'ownerName', 'updatedAt'];
var OD_GRANTS_HEADERS_ = ['campaign', 'grantedToEmail', 'grantedToName', 'accessLevel', 'grantedByEmail', 'grantedAt'];
var OD_USERS_HEADERS_ = ['email', 'name', 'team', 'role', 'pinHash', 'dateAdded', 'deactivated', 'managedBy'];

/** Roles that see ALL campaigns (functional/oversight roles) */
var OD_GLOBAL_VIEW_ROLES_ = { nlr: true, nlr_manager: true, bis: true, bis_manager: true, aptel: true, superadmin: true };

/**
 * Resolve which campaigns a user can see based on role + ownership + grants.
 * Returns { visible: [campaignKeys], editable: [campaignKeys], role, accessMap }
 * accessMap: { campaignKey: 'view'|'edit'|'own' }
 */
function odResolveVisibleCampaigns(email, role) {
  if (!email) return { visible: [], editable: [], role: role || '', accessMap: {} };
  email = email.toString().trim().toLowerCase();
  role = (role || '').toString().trim().toLowerCase();

  // Global-view roles see everything
  if (OD_GLOBAL_VIEW_ROLES_[role]) {
    // Get ALL campaign keys from ownership table
    var ownershipRows = odReadTab('_OD_Campaign_Ownership');
    var allCampaigns = [];
    var accessMap = {};
    for (var i = 0; i < ownershipRows.length; i++) {
      var ck = String(ownershipRows[i].campaign || '').trim();
      if (ck && !(OD_CAMPAIGNS[ck] && OD_CAMPAIGNS[ck].hidden)) {
        allCampaigns.push(ck);
        accessMap[ck] = role === 'superadmin' ? 'edit' : 'view';
      }
    }
    // Also include campaigns from odGetCampaignOwners that may not be in ownership table yet
    return { visible: allCampaigns, editable: role === 'superadmin' ? allCampaigns : [], role: role, accessMap: accessMap };
  }

  var accessMap = {};

  // Org Manager: owns campaigns
  if (role === 'org_manager') {
    var ownershipRows = odReadTab('_OD_Campaign_Ownership');
    for (var i = 0; i < ownershipRows.length; i++) {
      var ownerEmail = String(ownershipRows[i].ownerEmail || '').trim().toLowerCase();
      if (ownerEmail === email) {
        var ck = String(ownershipRows[i].campaign || '').trim();
        if (ck) accessMap[ck] = 'own';
      }
    }
  }

  // Admin: inherits their Org Manager's campaigns
  if (role === 'admin') {
    var managedBy = _odGetManagedBy(email);
    if (managedBy) {
      var ownershipRows = odReadTab('_OD_Campaign_Ownership');
      for (var i = 0; i < ownershipRows.length; i++) {
        var ownerEmail = String(ownershipRows[i].ownerEmail || '').trim().toLowerCase();
        if (ownerEmail === managedBy) {
          var ck = String(ownershipRows[i].campaign || '').trim();
          if (ck) accessMap[ck] = 'edit'; // admins can edit their OM's campaigns
        }
      }
    }
  }

  // National Consultant: auto-access to campaigns owned by their Org Managers.
  // An OM's managedBy field points to their National's email.
  if (role === 'national') {
    var usersRows = odReadTab('_OD_Users');
    var myOrgManagers = [];
    for (var u = 0; u < usersRows.length; u++) {
      var uRole = String(usersRows[u].role || '').trim().toLowerCase();
      var uManagedBy = String(usersRows[u].managedBy || '').trim().toLowerCase();
      if (uRole === 'org_manager' && uManagedBy === email) {
        myOrgManagers.push(String(usersRows[u].email || '').trim().toLowerCase());
      }
    }
    if (myOrgManagers.length > 0) {
      var ownershipRows = odReadTab('_OD_Campaign_Ownership');
      for (var i = 0; i < ownershipRows.length; i++) {
        var ownerEmail = String(ownershipRows[i].ownerEmail || '').trim().toLowerCase();
        if (myOrgManagers.indexOf(ownerEmail) !== -1) {
          var ck = String(ownershipRows[i].campaign || '').trim();
          if (ck && !accessMap[ck]) accessMap[ck] = 'auto'; // auto-granted, view+edit via OM relationship
        }
      }
    }
  }

  // All roles: check explicit access grants (direct grants to this user)
  var grantRows = odReadTab('_OD_Access_Grants');
  for (var i = 0; i < grantRows.length; i++) {
    var grantee = String(grantRows[i].grantedToEmail || '').trim().toLowerCase();
    if (grantee !== email) continue;
    var ck = String(grantRows[i].campaign || '').trim();
    var level = String(grantRows[i].accessLevel || '').trim().toLowerCase();
    if (!ck || !level) continue;
    // Don't downgrade: own > auto > edit > view
    var current = accessMap[ck];
    if (current === 'own') continue;
    if (current === 'auto') continue;
    if (current === 'edit' && level === 'view') continue;
    accessMap[ck] = level;
  }

  // Org Manager / Admin: inherit grants given to their National.
  // Grants are given to Nationals — their team (OMs + Admins) inherits the same level.
  if (role === 'org_manager' || role === 'admin') {
    // Walk up the chain to find the National: OM → managedBy = National, Admin → managedBy = OM → OM.managedBy = National
    var nationalEmail = '';
    if (role === 'org_manager') {
      nationalEmail = _odGetManagedBy(email); // OM's managedBy = National
    } else if (role === 'admin') {
      var omEmail = _odGetManagedBy(email);   // Admin's managedBy = OM
      if (omEmail) nationalEmail = _odGetManagedBy(omEmail); // OM's managedBy = National
    }
    if (nationalEmail) {
      for (var g = 0; g < grantRows.length; g++) {
        var natGrantee = String(grantRows[g].grantedToEmail || '').trim().toLowerCase();
        if (natGrantee !== nationalEmail) continue;
        var gck = String(grantRows[g].campaign || '').trim();
        var glevel = String(grantRows[g].accessLevel || '').trim().toLowerCase();
        if (!gck || !glevel) continue;
        var gcurrent = accessMap[gck];
        if (gcurrent === 'own' || gcurrent === 'auto' || gcurrent === 'edit') continue;
        accessMap[gck] = glevel;
      }
    }
  }

  // Filter out campaigns marked hidden in OD_CAMPAIGNS config
  for (var hk in accessMap) {
    if (OD_CAMPAIGNS[hk] && OD_CAMPAIGNS[hk].hidden) delete accessMap[hk];
  }

  var visible = Object.keys(accessMap);
  var editable = [];
  for (var k in accessMap) {
    if (accessMap[k] === 'own' || accessMap[k] === 'edit' || accessMap[k] === 'auto') editable.push(k);
  }

  return { visible: visible, editable: editable, role: role, accessMap: accessMap };
}

/**
 * Look up a user's managedBy field from _OD_Users.
 */
function _odGetManagedBy(email) {
  if (!email) return '';
  var target = email.toString().trim().toLowerCase();
  var sheet = odGetOrCreateTab('_OD_Users', OD_USERS_HEADERS_);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return '';
  var headers = data[0];
  var emailCol = -1, mbCol = -1;
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).trim().toLowerCase();
    if (h === 'email') emailCol = c;
    else if (h === 'managedby') mbCol = c;
  }
  if (emailCol < 0 || mbCol < 0) return '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emailCol] || '').trim().toLowerCase() === target) {
      return String(data[i][mbCol] || '').trim().toLowerCase();
    }
  }
  return '';
}

/**
 * Look up a user's role from _OD_Users.
 */
function _odGetUserRole(email) {
  if (!email) return '';
  var target = email.toString().trim().toLowerCase();
  var rows = odReadTab('_OD_Users');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].email || '').trim().toLowerCase() === target) {
      return String(rows[i].role || '').trim().toLowerCase();
    }
  }
  return '';
}

/**
 * Check if a user can edit a specific mapping column for a campaign.
 * Returns true/false.
 *   - Source Tab: org_manager who owns it, admin under that OM, edit-granted users, superadmin
 *   - BIS col: bis_manager, bis, superadmin
 *   - NLR cols: nlr_manager, nlr, superadmin
 */
function _odCanEditColumn(email, role, campaign, column, accessMap) {
  if (role === 'superadmin') return true;

  // NLR columns
  if (column === 'nlrWorkbook' || column === 'nlrWorkbookId' || column === 'nlrWorkbookName' || column === 'nlrTab') {
    return role === 'nlr' || role === 'nlr_manager';
  }
  // BIS column
  if (column === 'camCompany') {
    return role === 'bis' || role === 'bis_manager';
  }
  // Source Tab (campaign tab map) — needs campaign ownership or edit grant
  if (column === 'sourceTab' || column === 'tabName') {
    var access = accessMap ? accessMap[campaign] : null;
    return access === 'own' || access === 'edit' || access === 'auto';
  }
  return false;
}

// ── Campaign Ownership CRUD ──

function odGetCampaignOwnership() {
  odGetOrCreateTab('_OD_Campaign_Ownership', OD_OWNERSHIP_HEADERS_);
  var rows = odReadTab('_OD_Campaign_Ownership');
  return { success: true, ownership: rows };
}

function odSaveCampaignOwnership(body) {
  var campaign = String(body.campaign || '').trim();
  var ownerEmail = String(body.ownerEmail || '').trim().toLowerCase();
  var ownerName = String(body.ownerName || '').trim();
  if (!campaign || !ownerEmail) return { error: 'campaign and ownerEmail required' };

  var sheet = odGetOrCreateTab('_OD_Campaign_Ownership', OD_OWNERSHIP_HEADERS_);
  var data = sheet.getDataRange().getValues();

  // Upsert by campaign key
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === campaign.toLowerCase()) {
      sheet.getRange(i + 1, 2).setValue(ownerEmail);
      sheet.getRange(i + 1, 3).setValue(ownerName);
      sheet.getRange(i + 1, 4).setValue(new Date().toISOString());
      return { success: true, updated: true };
    }
  }

  // Insert new
  sheet.appendRow([campaign, ownerEmail, ownerName, new Date().toISOString()]);
  return { success: true, created: true };
}

// ── Access Grants CRUD ──

function odGetAccessGrants(callerEmail) {
  odGetOrCreateTab('_OD_Access_Grants', OD_GRANTS_HEADERS_);
  var rows = odReadTab('_OD_Access_Grants');
  // If callerEmail provided, filter to grants relevant to that user
  // (grants they created OR grants given to them)
  if (callerEmail) {
    var target = callerEmail.toString().trim().toLowerCase();
    var filtered = [];
    for (var i = 0; i < rows.length; i++) {
      var grantee = String(rows[i].grantedToEmail || '').trim().toLowerCase();
      var granter = String(rows[i].grantedByEmail || '').trim().toLowerCase();
      if (grantee === target || granter === target) filtered.push(rows[i]);
    }
    return { success: true, grants: filtered };
  }
  return { success: true, grants: rows };
}

function odSaveAccessGrant(body) {
  var campaign = String(body.campaign || '').trim();
  var grantedToEmail = String(body.grantedToEmail || '').trim().toLowerCase();
  var grantedToName = String(body.grantedToName || '').trim();
  var accessLevel = String(body.accessLevel || '').trim().toLowerCase();
  var grantedByEmail = String(body.grantedByEmail || '').trim().toLowerCase();
  if (!campaign || !grantedToEmail || !accessLevel) return { error: 'campaign, grantedToEmail, and accessLevel required' };
  if (accessLevel !== 'view' && accessLevel !== 'edit') return { error: 'accessLevel must be view or edit' };

  // Permission check: caller must own this campaign or be superadmin
  var callerRole = _odGetUserRole(grantedByEmail);
  if (callerRole !== 'superadmin') {
    var ownershipRows = odReadTab('_OD_Campaign_Ownership');
    var isOwner = false;
    for (var i = 0; i < ownershipRows.length; i++) {
      if (String(ownershipRows[i].campaign || '').trim().toLowerCase() === campaign.toLowerCase() &&
          String(ownershipRows[i].ownerEmail || '').trim().toLowerCase() === grantedByEmail) {
        isOwner = true;
        break;
      }
    }
    if (!isOwner) return { error: 'Only the campaign owner or superadmin can grant access' };
  }

  var sheet = odGetOrCreateTab('_OD_Access_Grants', OD_GRANTS_HEADERS_);
  var data = sheet.getDataRange().getValues();

  // Upsert by campaign + grantedToEmail
  for (var i = 1; i < data.length; i++) {
    var existCamp = String(data[i][0] || '').trim().toLowerCase();
    var existGrantee = String(data[i][1] || '').trim().toLowerCase();
    if (existCamp === campaign.toLowerCase() && existGrantee === grantedToEmail) {
      sheet.getRange(i + 1, 3).setValue(grantedToName);
      sheet.getRange(i + 1, 4).setValue(accessLevel);
      sheet.getRange(i + 1, 5).setValue(grantedByEmail);
      sheet.getRange(i + 1, 6).setValue(new Date().toISOString());
      return { success: true, updated: true };
    }
  }

  // Insert new
  sheet.appendRow([campaign, grantedToEmail, grantedToName, accessLevel, grantedByEmail, new Date().toISOString()]);
  return { success: true, created: true };
}

function odDeleteAccessGrant(body) {
  var campaign = String(body.campaign || '').trim().toLowerCase();
  var grantedToEmail = String(body.grantedToEmail || '').trim().toLowerCase();
  var callerEmail = String(body.callerEmail || '').trim().toLowerCase();
  if (!campaign || !grantedToEmail) return { error: 'campaign and grantedToEmail required' };

  // Permission check: caller must own this campaign or be superadmin
  var callerRole = _odGetUserRole(callerEmail);
  if (callerRole !== 'superadmin') {
    var ownershipRows = odReadTab('_OD_Campaign_Ownership');
    var isOwner = false;
    for (var i = 0; i < ownershipRows.length; i++) {
      if (String(ownershipRows[i].campaign || '').trim().toLowerCase() === campaign &&
          String(ownershipRows[i].ownerEmail || '').trim().toLowerCase() === callerEmail) {
        isOwner = true;
        break;
      }
    }
    if (!isOwner) return { error: 'Only the campaign owner or superadmin can revoke access' };
  }

  var sheet = odGetOrCreateTab('_OD_Access_Grants', OD_GRANTS_HEADERS_);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var existCamp = String(data[i][0] || '').trim().toLowerCase();
    var existGrantee = String(data[i][1] || '').trim().toLowerCase();
    if (existCamp === campaign && existGrantee === grantedToEmail) {
      sheet.deleteRow(i + 1);
      return { success: true, deleted: true };
    }
  }

  return { success: true, deleted: false, message: 'Grant not found' };
}

// ═══════════════════════════════════════════════════════
// OWNER DEVELOPMENT (OD) — doGet Endpoints
// ═══════════════════════════════════════════════════════

/**
 * Extract owner names from a campaign spreadsheet.
 * Supports 3 strategies based on campaign config:
 *   1. Default: Column A from first tab (skip header + footer-like rows)
 *   2. ownerSource='tabs': Tab names ARE owner names (minus excludeTabs)
 *   3. sectionHeader: Find names under a section header in a specific tab
 * @param {Object} cfg - Campaign config from OD_CAMPAIGNS
 * @param {Spreadsheet} ss - Already-opened spreadsheet
 * @return {string[]} Array of owner names
 */
function getOwnerNamesForCampaign_(cfg, ss) {
  var sheets = ss.getSheets();
  if (!sheets.length) return [];

  // ── Strategy 2: Tab names as owners (LeafGuard) ──
  if (cfg.ownerSource === 'tabs') {
    var exclude = (cfg.excludeTabs || []).map(function(t) { return t.toLowerCase(); });
    var owners = [];
    for (var s = 0; s < sheets.length; s++) {
      var tName = sheets[s].getName().trim();
      if (!tName || tName.charAt(0) === '_') continue;
      if (exclude.indexOf(tName.toLowerCase()) >= 0) continue;
      if (SKIP_TABS_.indexOf(tName.toLowerCase()) >= 0) continue;
      owners.push(tName);
    }
    return _applyExcludeOwners(owners, cfg);
  }

  // ── Strategy 2b: Visible (non-hidden) tab names as owners (AT&T B2B) ──
  // Only includes tabs that are NOT hidden and not in excludeTabs/SKIP_TABS_.
  // Tabs containing "INPUT" (case-insensitive) are also excluded.
  if (cfg.ownerSource === 'visible-tabs') {
    var exclude2 = (cfg.excludeTabs || []).map(function(t) { return t.toLowerCase(); });
    var owners2 = [];
    for (var s2 = 0; s2 < sheets.length; s2++) {
      if (sheets[s2].isSheetHidden()) continue;
      var tName2 = sheets[s2].getName().trim();
      if (!tName2 || tName2.charAt(0) === '_') continue;
      if (tName2.toLowerCase().indexOf('input') >= 0) continue;
      if (exclude2.indexOf(tName2.toLowerCase()) >= 0) continue;
      if (SKIP_TABS_.indexOf(tName2.toLowerCase()) >= 0) continue;
      owners2.push(tName2);
    }
    Logger.log('getOwnerNames visible-tabs: ' + owners2.length + ' visible owners');
    return _applyExcludeOwners(owners2, cfg);
  }

  // ── Strategy 3: Section header in a specific tab (Lumen) ──
  if (cfg.sectionHeader) {
    var tab = null;
    if (cfg.sourceTab) {
      tab = ss.getSheetByName(cfg.sourceTab);
      Logger.log('getOwnerNames sectionHeader: sourceTab="' + cfg.sourceTab + '" found=' + !!tab);
    }
    if (!tab) {
      tab = sheets[0]; // fallback to first tab
      Logger.log('getOwnerNames sectionHeader: fell back to first tab "' + tab.getName() + '"');
    }

    var data = tab.getDataRange().getValues();
    var headerLower = cfg.sectionHeader.toLowerCase();
    var owners = [];
    var inSection = false;

    // Log first 5 rows of column A for debugging
    Logger.log('getOwnerNames sectionHeader: looking for "' + headerLower + '" in tab "' + tab.getName() + '" (' + data.length + ' rows)');
    for (var dbg = 0; dbg < Math.min(data.length, 20); dbg++) {
      var dbgVal = String(data[dbg][0] || '').trim();
      Logger.log('  row ' + dbg + ' col A: "' + dbgVal + '" (len=' + dbgVal.length + ', match=' + (dbgVal.toLowerCase() === headerLower) + ')');
    }

    for (var i = 0; i < data.length; i++) {
      var val = String(data[i][0] || '').trim();
      var valLower = val.toLowerCase();

      // Found our section header — start collecting from next row
      if (valLower === headerLower) {
        inSection = true;
        Logger.log('getOwnerNames sectionHeader: FOUND header at row ' + i);
        continue;
      }

      if (!inSection) continue;

      // Stop at empty row, another all-caps section header, or footer-like row
      if (!val) { Logger.log('getOwnerNames sectionHeader: BREAK empty at row ' + i); break; }
      if (val === val.toUpperCase() && val.length > 2 && !/\d/.test(val)) { Logger.log('getOwnerNames sectionHeader: BREAK section header at row ' + i + ': "' + val + '"'); break; }
      if (valLower.indexOf('total') >= 0 || valLower.indexOf('template') >= 0 ||
          valLower.indexOf('summary') >= 0 || valLower.indexOf('***') >= 0 ||
          valLower.indexOf('average') >= 0) { Logger.log('getOwnerNames sectionHeader: BREAK footer at row ' + i + ': "' + val + '"'); break; }

      owners.push(val);
    }
    Logger.log('getOwnerNames sectionHeader: returning ' + owners.length + ' owners: ' + owners.join(', '));
    return _applyExcludeOwners(owners, cfg);
  }

  // ── Strategy 1 (Default): Column A from first tab ──
  var firstTab = sheets[0];
  var data = firstTab.getDataRange().getValues();
  var owners = [];
  for (var i = 1; i < data.length; i++) {
    var val = String(data[i][0] || '').trim();
    if (!val) continue;
    var valLower = val.toLowerCase();
    if (valLower.indexOf('total') >= 0 || valLower.indexOf('template') >= 0 ||
        valLower.indexOf('campaign') >= 0 || valLower.indexOf('summary') >= 0 ||
        valLower.indexOf('sum') >= 0 || valLower.indexOf('***') >= 0 ||
        valLower.indexOf('header') >= 0 || valLower.indexOf('average') >= 0) continue;
    owners.push(val);
  }
  return _applyExcludeOwners(owners, cfg);
}

/**
 * Filter out explicitly excluded owner names (case-insensitive).
 * Works with any owner-source strategy — applied as a final pass.
 */
function _applyExcludeOwners(owners, cfg) {
  if (!cfg.excludeOwners || !cfg.excludeOwners.length) return owners;
  var excludeSet = {};
  for (var i = 0; i < cfg.excludeOwners.length; i++) {
    excludeSet[cfg.excludeOwners[i].toLowerCase()] = true;
  }
  return owners.filter(function(name) {
    return !excludeSet[name.toLowerCase()];
  });
}

/**
 * action=odCampaignOwners
 * For each campaign in OD_CAMPAIGNS, open the sheet and extract owners
 * using the appropriate strategy. Return owners list per campaign.
 */
function odGetCampaignOwners() {
  var campaigns = {};
  var keys = Object.keys(OD_CAMPAIGNS);
  for (var k = 0; k < keys.length; k++) {
    var key = keys[k];
    var cfg = OD_CAMPAIGNS[key];
    if (!cfg.sheetId) continue;
    try {
      var ss = SpreadsheetApp.openById(cfg.sheetId);
      var owners = getOwnerNamesForCampaign_(cfg, ss);
      // Also collect all tab names (for campaign tab mapping dropdown)
      var sheets = ss.getSheets();
      var tabNames = [];
      for (var s = 0; s < sheets.length; s++) {
        var tName = sheets[s].getName().trim();
        if (tName && tName.charAt(0) !== '_') tabNames.push(tName);
      }
      campaigns[key] = { label: cfg.label, owners: owners, tabs: tabNames };
    } catch (err) {
      Logger.log('odCampaignOwners error for ' + key + ': ' + err.message);
      campaigns[key] = { label: cfg.label, owners: [], tabs: [], error: err.message };
    }
  }
  return { success: true, campaigns: campaigns };
}

/**
 * action=odNlrWorkbooks
 * List all Google Sheets files in the NLR Drive folder.
 */
function odGetNlrWorkbooks(includeTabs) {
  var folder = DriveApp.getFolderById(OD_NLR_FOLDER);
  var result = [];
  odScanFolderForSheets_(folder, result, !!includeTabs);
  return { success: true, workbooks: result };
}

/**
 * Recursively scan a folder and all subfolders for Google Sheets.
 * @param {Folder} folder - Drive folder to scan
 * @param {Array} result - accumulator for { id, name, tabs? } objects
 * @param {boolean} includeTabs - if true, open each sheet and read tab names (slow)
 */
function odScanFolderForSheets_(folder, result, includeTabs) {
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    var f = files.next();
    var fId = f.getId();
    var entry = { id: fId, name: f.getName() };
    if (includeTabs) {
      var tabs = [];
      try {
        var ss = SpreadsheetApp.openById(fId);
        var sheets = ss.getSheets();
        for (var i = 0; i < sheets.length; i++) {
          tabs.push(sheets[i].getName());
        }
      } catch (e) {}
      entry.tabs = tabs;
    }
    result.push(entry);
  }
  // Recurse into subfolders
  var subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    odScanFolderForSheets_(subfolders.next(), result, includeTabs);
  }
}

/**
 * action=odNlrTabs (param: sheetId)
 * Return all tab names from the given spreadsheet.
 */
function odGetNlrTabs(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheets = ss.getSheets();
  var tabs = [];
  for (var i = 0; i < sheets.length; i++) {
    tabs.push(sheets[i].getName());
  }
  return { success: true, tabs: tabs };
}

/**
 * action=odGetMappings
 * Read _OD_Mappings tab, return all rows as objects.
 */
function odGetMappings() {
  odGetOrCreateTab('_OD_Mappings', ['campaign', 'ownerName', 'camCompany', 'nlrWorkbookId', 'nlrWorkbookName', 'nlrTab', 'updatedBy', 'updatedAt']);
  var rows = odReadTab('_OD_Mappings');
  return { success: true, mappings: rows };
}

/**
 * action=odGetUsers
 * Read _OD_Users tab, return all rows as objects but EXCLUDE pinHash.
 */
function odGetUsers() {
  odGetOrCreateTab('_OD_Users', OD_USERS_HEADERS_);
  var rows = odReadTab('_OD_Users');
  var safe = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    safe.push({
      email: r.email,
      name: r.name,
      team: r.team,
      role: r.role,
      dateAdded: r.dateAdded,
      deactivated: r.deactivated,
      managedBy: r.managedBy || ''
    });
  }
  return { success: true, users: safe };
}

/**
 * action=odCheckUser (param: email)
 * Check if user exists in _OD_Users. Returns { success, hasPin } or error.
 */
function odCheckUser(email) {
  if (!email) return { success: false, message: 'Email is required' };

  var sheet = odGetOrCreateTab('_OD_Users', OD_USERS_HEADERS_);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: false, message: 'Contact your team manager to get added.' };

  var headers = data[0];
  var emailCol = -1, pinCol = -1, deactCol = -1;
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).trim().toLowerCase();
    if (h === 'email') emailCol = c;
    else if (h === 'pinhash') pinCol = c;
    else if (h === 'deactivated') deactCol = c;
  }
  if (emailCol < 0) return { success: false, message: 'Tab misconfigured' };

  var target = email.toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][emailCol] || '').trim().toLowerCase();
    if (rowEmail !== target) continue;

    // Check deactivated
    if (deactCol >= 0) {
      var deact = String(data[i][deactCol] || '').trim().toLowerCase();
      if (deact === 'true' || deact === 'yes') {
        return { success: false, message: 'Account deactivated' };
      }
    }

    var storedHash = (pinCol >= 0) ? String(data[i][pinCol] || '').trim() : '';
    return { success: true, hasPin: !!storedHash };
  }

  return { success: false, message: 'Contact your team manager to get added.' };
}

/**
 * action=odLogin (params: email, pin, createPin?)
 * Find user by email (case-insensitive). If createPin is true and no pinHash stored,
 * set the PIN. Otherwise validate the hash.
 * Returns flat user fields (email, name, team, role) directly on the response object.
 */
function odLogin(email, pin, createPin) {
  var sheet = odGetOrCreateTab('_OD_Users', OD_USERS_HEADERS_);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: false, message: 'No users found' };

  var headers = data[0];
  var emailCol = -1, pinCol = -1, nameCol = -1, teamCol = -1, roleCol = -1, deactCol = -1, mbCol = -1;
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).trim().toLowerCase();
    if (h === 'email') emailCol = c;
    else if (h === 'pinhash') pinCol = c;
    else if (h === 'name') nameCol = c;
    else if (h === 'team') teamCol = c;
    else if (h === 'role') roleCol = c;
    else if (h === 'deactivated') deactCol = c;
    else if (h === 'managedby') mbCol = c;
  }
  if (emailCol < 0 || pinCol < 0) return { success: false, message: 'Tab misconfigured' };

  var target = email.toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var rowEmail = String(data[i][emailCol] || '').trim().toLowerCase();
    if (rowEmail !== target) continue;

    // Check deactivated
    var deact = String(data[i][deactCol] || '').trim().toLowerCase();
    if (deact === 'true' || deact === 'yes') {
      return { success: false, message: 'Account deactivated' };
    }

    var storedHash = String(data[i][pinCol] || '').trim();
    var inputHash = odHashPin(pin);

    var userFields = {
      email: String(data[i][emailCol] || '').trim(),
      name: String(data[i][nameCol] || '').trim(),
      team: String(data[i][teamCol] || '').trim(),
      role: String(data[i][roleCol] || '').trim(),
      managedBy: mbCol >= 0 ? String(data[i][mbCol] || '').trim() : ''
    };

    if (!storedHash && createPin) {
      // First login — set the PIN
      sheet.getRange(i + 1, pinCol + 1).setValue(inputHash);
      userFields.success = true;
      userFields.firstLogin = true;
      return userFields;
    }

    if (!storedHash) {
      return { success: false, message: 'PIN not set. Please create a PIN.' };
    }

    if (inputHash === storedHash) {
      userFields.success = true;
      return userFields;
    } else {
      return { success: false, message: 'Incorrect PIN' };
    }
  }

  return { success: false, message: 'User not found' };
}

/**
 * action=odCamCompanies
 * Read Performance Audit sheet, get unique Business Name values.
 * Falls back to Client Name if Business Name column not found.
 */
function odGetCamCompanies() {
  var ss = SpreadsheetApp.openById(SHEETS.PERFORMANCE_AUDIT);

  // Use 'vlookup' tab (same as readOnlinePresence), fall back to first sheet
  var sheet = ss.getSheetByName('vlookup') || ss.getSheets()[0];
  if (!sheet) return { success: true, companies: [] };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, companies: [] };

  // Scan first 5 rows for header row — look for both "Business Name" and "Client Name"
  var headerRowIdx = -1;
  var bizCol = -1;
  var clientCol = -1;
  for (var r = 0; r < Math.min(5, data.length); r++) {
    for (var c = 0; c < data[r].length; c++) {
      var h = String(data[r][c]).trim().toLowerCase();
      if (h === 'business name') { bizCol = c; headerRowIdx = r; }
      if (h === 'client name') { clientCol = c; if (headerRowIdx < 0) headerRowIdx = r; }
    }
    if (headerRowIdx >= 0) break;
  }

  // Prefer Business Name; fall back to Client Name for company list
  var useCol = bizCol >= 0 ? bizCol : clientCol;
  if (useCol < 0) return { success: true, companies: [], clientToBusinessMap: [] };

  var seen = {};
  var companies = [];
  // Also build client name → business name pairs for auto-mapping
  var clientToBusinessMap = [];
  for (var i = headerRowIdx + 1; i < data.length; i++) {
    var bizVal = String(data[i][useCol] || '').trim();
    if (bizVal && !seen[bizVal]) {
      seen[bizVal] = true;
      companies.push(bizVal);
    }
    // If both columns exist, build the mapping pairs
    if (bizCol >= 0 && clientCol >= 0) {
      var clientVal = String(data[i][clientCol] || '').trim();
      if (clientVal && bizVal) {
        clientToBusinessMap.push({ clientName: clientVal, businessName: bizVal });
      }
    }
  }
  companies.sort();
  return { success: true, companies: companies, clientToBusinessMap: clientToBusinessMap };
}

// ═══════════════════════════════════════════════════════
// OWNER DEVELOPMENT (OD) — doPost Endpoints
// ═══════════════════════════════════════════════════════

/**
 * action: 'odSaveMapping'
 * body: { campaign, ownerName, field, value, updatedBy }
 * Find existing row by campaign+ownerName. Update specific field, or create new row.
 */
function odSaveMapping(body) {
  var campaign = String(body.campaign || '').trim();
  var ownerName = String(body.ownerName || '').trim();
  var field = String(body.field || '').trim();
  var value = body.value !== undefined ? body.value : '';
  var updatedBy = String(body.updatedBy || '').trim();

  if (!campaign || !ownerName || !field) {
    return { success: false, message: 'campaign, ownerName, and field are required' };
  }

  var validFields = ['camCompany', 'nlrWorkbook', 'nlrWorkbookId', 'nlrWorkbookName', 'nlrTab'];
  if (validFields.indexOf(field) < 0) {
    return { success: false, message: 'Invalid field. Must be one of: ' + validFields.join(', ') };
  }

  var tabHeaders = ['campaign', 'ownerName', 'camCompany', 'nlrWorkbookId', 'nlrWorkbookName', 'nlrTab', 'updatedBy', 'updatedAt'];
  var sheet = odGetOrCreateTab('_OD_Mappings', tabHeaders);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  // Find column indices
  var colMap = {};
  for (var c = 0; c < headers.length; c++) {
    colMap[String(headers[c]).trim()] = c;
  }

  var now = new Date().toISOString();

  // Search for existing row by campaign + ownerName
  var foundRow = -1;
  for (var i = 1; i < data.length; i++) {
    var rowCampaign = String(data[i][colMap['campaign']] || '').trim().toLowerCase();
    var rowOwner = String(data[i][colMap['ownerName']] || '').trim().toLowerCase();
    if (rowCampaign === campaign.toLowerCase() && rowOwner === ownerName.toLowerCase()) {
      foundRow = i + 1; // 1-based for sheet API
      break;
    }
  }

  // For 'nlrWorkbook' field, update nlrWorkbookId + nlrWorkbookName + nlrTab together
  var isNlrWorkbook = (field === 'nlrWorkbook');
  var nlrWorkbookId = isNlrWorkbook ? String(body.nlrWorkbookId || '') : '';
  var nlrWorkbookName = isNlrWorkbook ? String(body.nlrWorkbookName || '') : '';
  var nlrTab = isNlrWorkbook ? String(body.nlrTab || '') : '';

  if (foundRow > 0) {
    // Update existing row
    if (isNlrWorkbook) {
      sheet.getRange(foundRow, colMap['nlrWorkbookId'] + 1).setValue(nlrWorkbookId);
      sheet.getRange(foundRow, colMap['nlrWorkbookName'] + 1).setValue(nlrWorkbookName);
      sheet.getRange(foundRow, colMap['nlrTab'] + 1).setValue(nlrTab);
    } else {
      sheet.getRange(foundRow, colMap[field] + 1).setValue(value);
    }
    sheet.getRange(foundRow, colMap['updatedBy'] + 1).setValue(updatedBy);
    sheet.getRange(foundRow, colMap['updatedAt'] + 1).setValue(now);
  } else {
    // Create new row
    var newRow = [];
    for (var c = 0; c < tabHeaders.length; c++) {
      var col = tabHeaders[c];
      if (col === 'campaign') newRow.push(campaign);
      else if (col === 'ownerName') newRow.push(ownerName);
      else if (isNlrWorkbook && col === 'nlrWorkbookId') newRow.push(nlrWorkbookId);
      else if (isNlrWorkbook && col === 'nlrWorkbookName') newRow.push(nlrWorkbookName);
      else if (isNlrWorkbook && col === 'nlrTab') newRow.push(nlrTab);
      else if (!isNlrWorkbook && col === field) newRow.push(value);
      else if (col === 'updatedBy') newRow.push(updatedBy);
      else if (col === 'updatedAt') newRow.push(now);
      else newRow.push('');
    }
    sheet.appendRow(newRow);
  }

  return { success: true };
}

/**
 * action: 'odBatchSaveMappings'
 * body: { mappings: [{ campaign, ownerName, camCompany?, nlrWorkbookId?, nlrWorkbookName?, nlrTab?, updatedBy }] }
 * Efficiently saves multiple mappings in one call using batch sheet writes.
 */
function odBatchSaveMappings(body) {
  var items = body.mappings;
  if (!items || !items.length) return { success: true, saved: 0 };

  var tabHeaders = ['campaign', 'ownerName', 'camCompany', 'nlrWorkbookId', 'nlrWorkbookName', 'nlrTab', 'updatedBy', 'updatedAt'];
  var sheet = odGetOrCreateTab('_OD_Mappings', tabHeaders);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  // Build column index map
  var colMap = {};
  for (var c = 0; c < headers.length; c++) {
    colMap[String(headers[c]).trim()] = c;
  }

  // Index existing rows by campaign|ownerName (lowercase) for fast lookup
  var existingIndex = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][colMap['campaign']] || '').trim().toLowerCase() + '|' +
              String(data[i][colMap['ownerName']] || '').trim().toLowerCase();
    existingIndex[key] = i + 1; // 1-based row number
  }

  var now = new Date().toISOString();
  var newRows = [];
  var updates = []; // [{ row, col, value }]

  for (var j = 0; j < items.length; j++) {
    var item = items[j];
    var campaign = String(item.campaign || '').trim();
    var ownerName = String(item.ownerName || '').trim();
    if (!campaign || !ownerName) continue;

    var key = campaign.toLowerCase() + '|' + ownerName.toLowerCase();
    var updatedBy = String(item.updatedBy || 'auto-map').trim();

    if (existingIndex[key]) {
      // Update existing row
      var rowNum = existingIndex[key];
      if (item.camCompany !== undefined) updates.push({ row: rowNum, col: colMap['camCompany'] + 1, value: item.camCompany });
      if (item.nlrWorkbookId !== undefined) updates.push({ row: rowNum, col: colMap['nlrWorkbookId'] + 1, value: item.nlrWorkbookId });
      if (item.nlrWorkbookName !== undefined) updates.push({ row: rowNum, col: colMap['nlrWorkbookName'] + 1, value: item.nlrWorkbookName });
      if (item.nlrTab !== undefined) updates.push({ row: rowNum, col: colMap['nlrTab'] + 1, value: item.nlrTab });
      updates.push({ row: rowNum, col: colMap['updatedBy'] + 1, value: updatedBy });
      updates.push({ row: rowNum, col: colMap['updatedAt'] + 1, value: now });
    } else {
      // New row
      var newRow = [
        campaign,
        ownerName,
        item.camCompany || '',
        item.nlrWorkbookId || '',
        item.nlrWorkbookName || '',
        item.nlrTab || '',
        updatedBy,
        now
      ];
      newRows.push(newRow);
      // Mark as existing in case of duplicates in same batch
      existingIndex[key] = -1;
    }
  }

  // Apply cell updates
  for (var u = 0; u < updates.length; u++) {
    sheet.getRange(updates[u].row, updates[u].col).setValue(updates[u].value);
  }

  // Append new rows in bulk
  if (newRows.length > 0) {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newRows.length, tabHeaders.length).setValues(newRows);
  }

  SpreadsheetApp.flush();
  return { success: true, saved: items.length, updated: updates.length > 0 ? 'yes' : 'no', appended: newRows.length };
}

/**
 * action: 'odSaveUser'
 * body: { email, name, team, role, pin?, managedBy? }
 * Add or update user in _OD_Users.
 */
function odSaveUser(body) {
  var email = String(body.email || '').trim();
  var name = String(body.name || '').trim();
  var team = String(body.team || '').trim();
  var role = String(body.role || '').trim();
  var pin = body.pin ? String(body.pin).trim() : '';
  var managedBy = String(body.managedBy || '').trim();

  if (!email) return { success: false, message: 'email is required' };

  var sheet = odGetOrCreateTab('_OD_Users', OD_USERS_HEADERS_);
  var data = sheet.getDataRange().getValues();

  // Find existing row by email (case-insensitive)
  var target = email.toLowerCase();
  var foundRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === target) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow > 0) {
    // Update existing user
    sheet.getRange(foundRow, 2).setValue(name);       // name
    sheet.getRange(foundRow, 3).setValue(team);       // team
    sheet.getRange(foundRow, 4).setValue(role);       // role
    if (pin) {
      sheet.getRange(foundRow, 5).setValue(odHashPin(pin)); // pinHash
    }
    sheet.getRange(foundRow, 7).setValue('');         // clear deactivated on save
    sheet.getRange(foundRow, 8).setValue(managedBy);  // managedBy
  } else {
    // Add new user
    var pinHash = pin ? odHashPin(pin) : '';
    var now = new Date().toISOString();
    sheet.appendRow([email, name, team, role, pinHash, now, '', managedBy]);
  }

  return { success: true };
}

/**
 * action: 'odDeleteUser'
 * body: { email }
 * Soft-delete: set deactivated=true in _OD_Users.
 */
function odDeleteUser(body) {
  var email = String(body.email || '').trim();
  if (!email) return { success: false, message: 'email is required' };

  var sheet = odGetOrCreateTab('_OD_Users', OD_USERS_HEADERS_);
  var data = sheet.getDataRange().getValues();

  var target = email.toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === target) {
      sheet.getRange(i + 1, 7).setValue('true'); // deactivated column
      return { success: true };
    }
  }

  return { success: false, message: 'User not found' };
}


// ══════════════════════════════════════════════════════════════
// CONSOLIDATED CAMPAIGN DATA
// Reads per-campaign source spreadsheets (OD_CAMPAIGNS),
// extracts health + recruiting per owner per week,
// writes flat rows to per-campaign tabs in CONSOLIDATED sheet.
// ══════════════════════════════════════════════════════════════

// Headers for consolidated per-campaign tabs
// ── Per-campaign production product config ──
// Products tracked in the Production LW / Goals columns of Section 1.
// Combined cells use "/" separator (e.g. "5/3/2" = Frontier/Cell/TV).
// Order matters: values are split in this order. If fewer values than products,
// missing ones get 0. "TV" wasn't always tracked so Frontier may have 2 or 3 values.
var CAMPAIGN_PRODUCTS = {
  'frontier':        ['Frontier', 'TV', 'Cell'],
  'verizon-fios':    ['Units', 'Wireless'],
  'att-nds':         ['Units'],
  'att-res':         ['Internet', 'Wireless', 'DTV'],
  'rogers':          ['Units', 'Mobility'],
  'leafguard':       ['Personal Prod', 'Gross Leads', 'Number of Sales', 'Gross Sales'],
  'lumen':           ['Lumen', 'DTV'],
  'att-b2b':         ['Units']
};

// ── Build consolidated headers dynamically per campaign ──
// Base columns before production: Week, Owner, Active HC, Leaders, Dist, Training
// Then per-product: "Prod: <Product>" and "Goal: <Product>" pairs
// Then recruiting columns (Section 2)
var CONSOLIDATED_BASE_HEADERS = [
  'Week',            // 0
  'Owner',           // 1
  'Active HC',       // 2
  'Leaders',         // 3
  'Dist',            // 4
  'Training'         // 5
];

// Campaigns where each product has its own column in the source tab (not "/" separated)
// Maps product name → { prod: [...headers], goal: [...headers] }
var CAMPAIGN_PROD_COLUMNS = {
  'leafguard': {
    'Personal Prod':    { prod: ['personal prod', 'personal production'], goal: [] },
    'Gross Leads':      { prod: ['gross leads'], goal: ['gross leads goal', 'leads goal'] },
    'Number of Sales':  { prod: ['number of sales', '# of sales'], goal: [] },
    'Gross Sales':      { prod: ['gross sales'], goal: ['gross sales goal', 'sales goal', 'goal'] }
  },
  'verizon-fios': {
    'Units':     { prod: ['production lw', 'production', 'fios sales'], goal: [] },
    'Wireless':  { prod: ['wireless'], goal: [] }
  },
  'rogers': {
    'Units':     { prod: ['production lw', 'production'], goal: [] },
    'Mobility':  { prod: ['mobility', 'mobilty'], goal: [] }
  },
  'att-res': {
    'Internet':  { prod: ['internet'], goal: [] },
    'Wireless':  { prod: ['wireless'], goal: [] },
    'DTV':       { prod: ['dtv'], goal: [] }
  },
  'att-b2b': {
    'Units':     { prod: ['production lw', 'production'], goal: ['production goals'] }
  }
};

// Which products have goals set (these show goal comparison)
var CAMPAIGN_GOAL_PRODUCTS = {
  'leafguard': ['Gross Sales', 'Gross Leads'],
  'verizon-fios': ['Units'],  // Production Goals is combined (units + wireless) — shown against Units
  'rogers': ['Units'],        // Production Goals is combined (units + mobility) — shown against Units
  'att-b2b': ['Units']        // Production Goals → Goal: Units
};

// Campaigns with extra headcount columns (Closers, Lead Gen)
var CAMPAIGN_EXTRA_HC = {
  'leafguard': ['Closers', 'Lead Gen']
};

var CONSOLIDATED_RECRUITING_HEADERS = [
  'Calls Received',  // from Section 2 (applies)
  'Sent to List',    // from Section 2 (no list)
  '1st Booked',
  '1st Showed',
  '1st Retention',
  'Conversion',
  '2nd Booked',
  '2nd Showed',
  '2nd Retention',
  'NS Booked',
  'NS Showed',
  'NS Retention'
];

/**
 * Build full header row for a given campaign.
 * Layout: [base...] + [Prod: X, Goal: X, Prod: Y, Goal: Y, ...] + [recruiting...]
 */
function getConsolidatedHeaders_(campaignKey) {
  var products = CAMPAIGN_PRODUCTS[campaignKey] || ['Total'];
  var headers = CONSOLIDATED_BASE_HEADERS.slice();
  // Campaign-specific extra headcount columns (e.g. Closers, Lead Gen for LeafGuard)
  var extraHC = CAMPAIGN_EXTRA_HC[campaignKey] || [];
  for (var e = 0; e < extraHC.length; e++) {
    headers.push(extraHC[e]);
  }
  for (var p = 0; p < products.length; p++) {
    headers.push('Prod: ' + products[p]);
    headers.push('Goal: ' + products[p]);
  }
  return headers.concat(CONSOLIDATED_RECRUITING_HEADERS);
}

// Legacy single-column headers (kept for readConsolidatedRecruiting compatibility)
var CONSOLIDATED_HEADERS = [
  'Week',            // 0: date (from Section 1 or Section 2)
  'Owner',           // 1: owner name (tab name in source sheet)
  'Active HC',       // 2: from Section 1
  'Leaders',         // 3: from Section 1
  'Dist',            // 4: from Section 1
  'Training',        // 5: from Section 1
  'Production',      // 6: from Section 1
  'Goals',           // 7: from Section 1
  'Calls Received',  // 8: from Section 2 (applies)
  'Sent to List',    // 9: from Section 2 (no list)
  '1st Booked',      // 10: from Section 2
  '1st Showed',      // 11: from Section 2
  '1st Retention',   // 12: from Section 2 (rate)
  'Conversion',      // 13: from Section 2 (rate)
  '2nd Booked',      // 14: from Section 2
  '2nd Showed',      // 15: from Section 2
  '2nd Retention',   // 16: from Section 2 (rate)
  'NS Booked',       // 17: from Section 2
  'NS Showed',       // 18: from Section 2
  'NS Retention'     // 19: from Section 2 (rate)
];

// Tabs to skip when scanning campaign spreadsheets for owner tabs
var SKIP_TABS_ = ['campaign totals', 'template', 'instructions', 'summary', 'master',
                  'data', 'config', 'sheet1', 'dashboard', 'overview', 'totals'];

/**
 * Master refresh: iterate all campaigns in OD_CAMPAIGNS,
 * read source spreadsheets, write consolidated tabs.
 */
function refreshAllCampaigns() {
  var destSS;
  try {
    destSS = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { ok: false, error: 'Cannot open consolidated sheet: ' + err.message };
  }

  var results = {};
  var keys = Object.keys(OD_CAMPAIGNS);

  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var campaign = OD_CAMPAIGNS[key];
    if (!campaign.sheetId) {
      results[key] = { ok: false, error: 'No sheetId configured' };
      continue;
    }
    try {
      var count = consolidateCampaign_(key, campaign, destSS);
      results[key] = { ok: true, rows: count };
    } catch (err) {
      Logger.log('refreshAllCampaigns error for ' + key + ': ' + err.message + '\n' + err.stack);
      results[key] = { ok: false, error: err.message };
    }
  }

  return { ok: true, results: results, timestamp: new Date().toISOString() };
}

/**
 * Refresh a single campaign by key. Much faster than refreshAllCampaigns
 * since it only opens one source spreadsheet.
 * @param {string} campaignKey - e.g. 'lumen', 'leafguard', 'frontier'
 */
function refreshSingleCampaign(campaignKey) {
  if (!campaignKey) return { ok: false, error: 'campaign key is required' };

  var campaign = OD_CAMPAIGNS[campaignKey];
  if (!campaign) return { ok: false, error: 'Unknown campaign: ' + campaignKey };
  if (!campaign.sheetId) return { ok: false, error: 'No sheetId configured for ' + campaignKey };

  var destSS;
  try {
    destSS = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    return { ok: false, error: 'Cannot open consolidated sheet: ' + err.message };
  }

  try {
    var count = consolidateCampaign_(campaignKey, campaign, destSS);
    return { ok: true, campaign: campaignKey, rows: count, timestamp: new Date().toISOString() };
  } catch (err) {
    Logger.log('refreshSingleCampaign error for ' + campaignKey + ': ' + err.message + '\n' + err.stack);
    return { ok: false, campaign: campaignKey, error: err.message };
  }
}

/**
 * Consolidate a single campaign: read source spreadsheet,
 * extract health + recruiting from each owner tab,
 * write flat rows to destination tab.
 */
function consolidateCampaign_(campaignKey, campaign, destSS) {
  Logger.log('consolidateCampaign_ START: key=' + campaignKey + ' sheetId=' + campaign.sheetId + ' sectionHeader=' + (campaign.sectionHeader || 'none') + ' ownerSource=' + (campaign.ownerSource || 'default'));
  var srcSS = SpreadsheetApp.openById(campaign.sheetId);
  var allTabs = srcSS.getSheets();
  Logger.log('consolidateCampaign_ opened source sheet, tabs: ' + allTabs.map(function(t) { return t.getName(); }).join(', '));

  // Get owner names using the shared extraction helper
  var ownerNames = getOwnerNamesForCampaign_(campaign, srcSS);
  Logger.log('consolidateCampaign_ ownerNames (' + ownerNames.length + '): ' + ownerNames.join(', '));

  // ── Load saved tab mappings from _Campaign_Tab_Map ──
  var savedTabMap = {}; // lowercase ownerName → tabName
  try {
    var mapTab = destSS.getSheetByName('_Campaign_Tab_Map');
    if (mapTab) {
      var mapData = mapTab.getDataRange().getValues();
      for (var m = 1; m < mapData.length; m++) {
        var mapCampaign = String(mapData[m][0] || '').trim().toLowerCase();
        var mapOwner    = String(mapData[m][1] || '').trim().toLowerCase();
        var mapTabName  = String(mapData[m][2] || '').trim();
        if (mapCampaign === campaignKey.toLowerCase() && mapOwner && mapTabName) {
          savedTabMap[mapOwner] = mapTabName;
        }
      }
    }
  } catch (e) {
    Logger.log('consolidateCampaign_: could not load tab map: ' + e.message);
  }

  // Build tab lookup: tabName (lowercase) → Sheet object
  var tabByName = {};
  for (var t = 0; t < allTabs.length; t++) {
    var tn = allTabs[t].getName().trim();
    if (tn) tabByName[tn.toLowerCase()] = allTabs[t];
  }

  // Build list of all tab names for fuzzy matching
  var allTabNames = [];
  for (var t = 0; t < allTabs.length; t++) {
    var tn = allTabs[t].getName().trim();
    if (tn && SKIP_TABS_.indexOf(tn.toLowerCase()) < 0) {
      allTabNames.push(tn);
    }
  }

  var rows = [];

  // Process each owner: saved mapping → exact match → skip (but still include in owner list with zeros)
  for (var oi = 0; oi < ownerNames.length; oi++) {
    var ownerName = ownerNames[oi];
    var ownerLower = ownerName.toLowerCase();

    // Determine which tab to read for this owner
    var targetTabName = savedTabMap[ownerLower] || null;
    var tab = null;

    if (targetTabName) {
      // Use saved mapping (skip "non-partner" entries)
      if (targetTabName.toLowerCase() !== 'non-partner') {
        tab = tabByName[targetTabName.toLowerCase()] || null;
      }
    }
    if (!tab && !targetTabName) {
      // Exact name match only — no fuzzy matching (avoids grabbing wrong tab)
      tab = tabByName[ownerLower] || null;
    }

    Logger.log('consolidateCampaign_ TAB MATCH: "' + ownerName + '" → ' + (tab ? '"' + tab.getName() + '"' : 'NONE') + (targetTabName ? ' (saved: ' + targetTabName + ')' : ''));

    if (!tab) {
      Logger.log('consolidateCampaign_ NO TAB for: "' + ownerName + '" in ' + campaignKey);
      // No tab found — still include owner with a minimal placeholder row
      // so they appear in the owner list (with no data)
      // Snap to Sunday (or Monday for att-b2b) so placeholder rows align with real data rows
      var campaignHeaders = getConsolidatedHeaders_(campaignKey);
      var now = new Date();
      var dayOfWeek = now.getDay(); // 0=Sun
      var placeholderDate;
      if (campaignKey === 'att-b2b') {
        // Snap to Monday
        var monOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
        placeholderDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() + monOffset, 12, 0, 0);
      } else {
        placeholderDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayOfWeek, 12, 0, 0);
      }
      var emptyRow = [placeholderDate, ownerName];
      for (var ei = 2; ei < campaignHeaders.length; ei++) emptyRow.push(0);
      rows.push(emptyRow);
      continue;
    }

    try {
      var range = tab.getDataRange();
      var data = range.getValues();
      // Also get display values — needed for production/goals columns
      // where Google Sheets may auto-parse "90/6" as a Date instead of text
      var displayData = range.getDisplayValues();
      if (!data || data.length < 2) continue;

      var sections = findSections(data);

      // Extract health rows (Section 1): [{date, active, leaders, dist, training, production, goals}]
      // Pass displayData so slash-separated values like "90/6" are read as text
      var healthRows = extractHealthRows_(data, sections.section1Start, sections.section1End, displayData, campaignKey);

      // Extract recruiting rows from Section 1 (horizontal — inline with health, same row per week)
      var recruitingRows = extractHorizontalRecruitingRows_(data, sections.section1Start, sections.section1End, campaignKey);
      Logger.log('consolidateCampaign_ ' + ownerName + ': s2Start=' + sections.section2Start + ', recruitingRows=' + recruitingRows.length + ', healthRows=' + healthRows.length);

      // All campaigns except LeafGuard and AT&T B2B have source dates offset by -7 days.
      // Shift dates forward by 7 days at extraction time so consolidated data has correct dates.
      // LeafGuard: dates are direct. AT&T B2B: Monday-lock, dates are direct.
      var NO_DATE_OFFSET_CAMPAIGNS = ['leafguard', 'att-b2b'];
      if (NO_DATE_OFFSET_CAMPAIGNS.indexOf(campaignKey) < 0) {
        var offset = 7 * 86400000;
        for (var hi = 0; hi < healthRows.length; hi++) {
          if (healthRows[hi].date instanceof Date) {
            healthRows[hi].date = new Date(healthRows[hi].date.getTime() + offset);
          }
        }
        for (var ri = 0; ri < recruitingRows.length; ri++) {
          if (recruitingRows[ri].date instanceof Date) {
            recruitingRows[ri].date = new Date(recruitingRows[ri].date.getTime() + offset);
          }
        }
      }

      // Merge by date → flat output rows (campaign-aware product splitting)
      var merged = mergeHealthRecruiting_(ownerName, healthRows, recruitingRows, campaignKey);

      for (var r = 0; r < merged.length; r++) {
        rows.push(merged[r]);
      }
    } catch (err) {
      Logger.log('consolidateCampaign_ tab error (' + campaignKey + '/' + ownerName + '): ' + err.message);
    }
  }

  // Get or create destination tab
  var destTab = destSS.getSheetByName(campaign.label);
  if (!destTab) {
    destTab = destSS.insertSheet(campaign.label);
  }

  // Build campaign-specific headers
  var campaignHeaders = getConsolidatedHeaders_(campaignKey);
  var products = CAMPAIGN_PRODUCTS[campaignKey] || ['Total'];
  var extraHC = CAMPAIGN_EXTRA_HC[campaignKey] || [];

  // ── Read existing goals + production before clearing (to preserve manual/synced entries) ──
  var existingGoals = {}; // 'ownerLower|dateKey' → { productName: goalValue, ... }
  var existingProd = {};  // 'ownerLower|dateKey' → { productName: prodValue, ... }
  if (destTab) {
    var existingData = destTab.getDataRange().getValues();
    if (existingData.length > 1) {
      var exHeaders = existingData[0].map(function(h) { return String(h).trim(); });
      var goalCols = [];
      var prodCols = [];
      for (var gc = 0; gc < exHeaders.length; gc++) {
        if (exHeaders[gc].indexOf('Goal: ') === 0) {
          goalCols.push({ colIdx: gc, productName: exHeaders[gc].replace('Goal: ', '') });
        }
        if (exHeaders[gc].indexOf('Prod: ') === 0) {
          prodCols.push({ colIdx: gc, productName: exHeaders[gc].replace('Prod: ', '') });
        }
      }
      var exOwnerCol = exHeaders.indexOf('Owner');
      var exWeekCol = exHeaders.indexOf('Week');
      if (exOwnerCol >= 0 && exWeekCol >= 0) {
        for (var ei = 1; ei < existingData.length; ei++) {
          var exOwner = String(existingData[ei][exOwnerCol] || '').trim().toLowerCase();
          var exDate = existingData[ei][exWeekCol];
          if (!exOwner || !exDate) continue;
          var exDateKey = _normalizeDateKey_(exDate instanceof Date ? exDate : _parseTabDate(String(exDate)), campaignKey);
          var ownerGoals = {}, ownerProd = {};
          var hasGoals = false, hasProd = false;
          for (var gci = 0; gci < goalCols.length; gci++) {
            var gVal = parseInt(existingData[ei][goalCols[gci].colIdx]) || 0;
            if (gVal) { ownerGoals[goalCols[gci].productName] = gVal; hasGoals = true; }
          }
          for (var pci = 0; pci < prodCols.length; pci++) {
            var pVal = parseInt(existingData[ei][prodCols[pci].colIdx]) || 0;
            if (pVal) { ownerProd[prodCols[pci].productName] = pVal; hasProd = true; }
          }
          if (hasGoals) existingGoals[exOwner + '|' + exDateKey] = ownerGoals;
          if (hasProd) existingProd[exOwner + '|' + exDateKey] = ownerProd;
        }
      }
    }
    Logger.log('consolidateCampaign_ preserved ' + Object.keys(existingGoals).length + ' goal, ' + Object.keys(existingProd).length + ' production entries');
  }

  // ── Restore goals + production into rows ──
  var prodStart = 6 + extraHC.length;
  for (var ri = 0; ri < rows.length; ri++) {
    var rOwner = String(rows[ri][1] || '').trim().toLowerCase();
    var rDateKey = _normalizeDateKey_(rows[ri][0], campaignKey);
    var savedGoals = existingGoals[rOwner + '|' + rDateKey];
    var savedProd = existingProd[rOwner + '|' + rDateKey];
    for (var pi = 0; pi < products.length; pi++) {
      var prodIdx = prodStart + pi * 2;
      var goalIdx = prodStart + pi * 2 + 1;
      // Production: keep existing if source is 0
      if (!rows[ri][prodIdx] && savedProd && savedProd[products[pi]]) {
        rows[ri][prodIdx] = savedProd[products[pi]];
      }
      // Goals: restore if current is 0
      if (!rows[ri][goalIdx] && savedGoals && savedGoals[products[pi]]) {
        rows[ri][goalIdx] = savedGoals[products[pi]];
      }
    }
  }

  // ── Restore orphaned goal/production rows (future weeks not in source data) ──
  var restoredKeys = {};
  for (var ri2 = 0; ri2 < rows.length; ri2++) {
    var rk2 = String(rows[ri2][1] || '').trim().toLowerCase() + '|' + _normalizeDateKey_(rows[ri2][0], campaignKey);
    restoredKeys[rk2] = true;
  }
  // Check existingGoals and existingProd for entries that weren't matched to any row
  var allPreservedKeys = {};
  for (var gk in existingGoals) allPreservedKeys[gk] = true;
  for (var pk in existingProd) allPreservedKeys[pk] = true;
  for (var orphanKey in allPreservedKeys) {
    if (restoredKeys[orphanKey]) continue; // already matched
    var parts = orphanKey.split('|');
    var orphanOwner = parts[0];
    var orphanDate = parts[1];
    if (!orphanOwner || !orphanDate) continue;
    // Find proper-cased owner name from existing rows
    var properOwner = orphanOwner;
    for (var ni = 0; ni < rows.length; ni++) {
      if (String(rows[ni][1] || '').trim().toLowerCase() === orphanOwner) {
        properOwner = String(rows[ni][1] || '').trim();
        break;
      }
    }
    // Build a new row with goals/production preserved
    var orphanRow = [];
    for (var oi2 = 0; oi2 < campaignHeaders.length; oi2++) orphanRow.push(0);
    var dateParts = orphanDate.split('/');
    orphanRow[0] = new Date(Number(dateParts[2]), Number(dateParts[0]) - 1, Number(dateParts[1]), 12, 0, 0);
    orphanRow[1] = properOwner;
    // Fill in goals and production
    var oGoals = existingGoals[orphanKey] || {};
    var oProd = existingProd[orphanKey] || {};
    for (var opi = 0; opi < products.length; opi++) {
      var oProdIdx = prodStart + opi * 2;
      var oGoalIdx = prodStart + opi * 2 + 1;
      if (oProd[products[opi]]) orphanRow[oProdIdx] = oProd[products[opi]];
      if (oGoals[products[opi]]) orphanRow[oGoalIdx] = oGoals[products[opi]];
    }
    rows.push(orphanRow);
    Logger.log('consolidateCampaign_ restored orphaned row: ' + properOwner + ' | ' + orphanDate);
  }

  // ── Dedup: merge duplicate owner+date rows ──
  var dedupMap = {};
  for (var di = 0; di < rows.length; di++) {
    var dOwner = String(rows[di][1] || '').trim().toLowerCase();
    var dDate = _normalizeDateKey_(rows[di][0], campaignKey);
    var dKey = dOwner + '|' + dDate;
    if (!dedupMap[dKey]) {
      dedupMap[dKey] = di;
    } else {
      var existIdx = dedupMap[dKey];
      for (var mi = 2; mi < rows[di].length; mi++) {
        var existVal = Number(rows[existIdx][mi]) || 0;
        var newVal = Number(rows[di][mi]) || 0;
        if (newVal > existVal) rows[existIdx][mi] = rows[di][mi];
      }
      rows.splice(di, 1);
      di--;
    }
  }
  Logger.log('consolidateCampaign_ after dedup: ' + rows.length + ' rows');

  // Clear and write
  destTab.clear();
  destTab.getRange(1, 1, 1, campaignHeaders.length).setValues([campaignHeaders]);

  if (rows.length > 0) {
    // Sort by date descending, then owner ascending
    rows.sort(function(a, b) {
      var dateA = (a[0] instanceof Date) ? a[0].getTime() : new Date(a[0]).getTime() || 0;
      var dateB = (b[0] instanceof Date) ? b[0].getTime() : new Date(b[0]).getTime() || 0;
      if (dateA !== dateB) return dateB - dateA;
      return String(a[1]).localeCompare(String(b[1]));
    });
    destTab.getRange(2, 1, rows.length, campaignHeaders.length).setValues(rows);
  }

  return rows.length;
}

/**
 * Extract health (Section 1) data as flat rows.
 * Returns: [{ date: Date, active, leaders, dist, training, production, goals }]
 */
function extractHealthRows_(data, start, end, displayData, campaignKey) {
  if (start < 0 || end < start) return [];

  // Find the actual header row — skip empty rows at the start
  var headerIdx = start;
  for (var hi = start; hi <= Math.min(start + 3, end); hi++) {
    var rowText = data[hi].map(function(c) { return String(c).trim(); }).join('');
    if (rowText.length > 0) { headerIdx = hi; break; }
  }

  var headers = data[headerIdx].map(function(h) { return String(h).toLowerCase().trim(); });
  var colMap = {
    dates: findCol(headers, ['dates', 'date']),
    active: findCol(headers, ['active', 'total agents', 'agents']),
    leaders: findCol(headers, ['leaders']),
    dist: findCol(headers, ['dist', 'distributors', 'distibutors']),
    closers: findCol(headers, ['closers']),
    leadGen: findCol(headers, ['lead gen', 'leadgen']),
    training: findCol(headers, ['training']),
    production: findCol(headers, ['production lw', 'production', 'total production']),
    goals: findCol(headers, ['production goals', 'production goal', 'goals'])
  };

  // Per-product column detection (LeafGuard: separate columns per product)
  // Per-product column detection (LeafGuard: separate columns per product)
  var perProdCols = null; // { productName: { prodCol, goalCol } }
  var prodColConfig = campaignKey && CAMPAIGN_PROD_COLUMNS[campaignKey];
  if (prodColConfig) {
    perProdCols = {};
    var prodNames = Object.keys(prodColConfig);
    for (var pp = 0; pp < prodNames.length; pp++) {
      var cfg = prodColConfig[prodNames[pp]];
      perProdCols[prodNames[pp]] = {
        prodCol: findCol(headers, cfg.prod),
        goalCol: cfg.goal && cfg.goal.length ? findCol(headers, cfg.goal) : -1
      };
    }
  }

  var result = [];
  for (var i = headerIdx + 1; i <= end; i++) {
    var row = data[i];
    var dateVal = row[colMap.dates];
    if (!dateVal) continue;
    var d = (dateVal instanceof Date) ? dateVal : _parseTabDate(String(dateVal));
    if (!d) continue;

    var prodRaw, goalsRaw;
    if (perProdCols) {
      // Per-product columns: build "/" separated string for compatibility with mergeHealthRecruiting_
      var prodVals = [], goalVals = [];
      var products = CAMPAIGN_PRODUCTS[campaignKey] || [];

      // Read the shared "goals" column — may be "/" separated (e.g., "60/30" for Units/Wireless)
      var sharedGoalRaw = '';
      if (colMap.goals >= 0) {
        sharedGoalRaw = displayData && displayData[i] ? String(displayData[i][colMap.goals]) : String(row[colMap.goals] || '');
      }
      var sharedGoalParts = _splitSlashValues(sharedGoalRaw);

      // First pass: read production values + detect which have dedicated goal columns
      // Use display data and take only the FIRST slash-part so that cells with multiple
      // outputs like "20/5" (e.g. Ayman's Wireless column) only count the first value.
      var hasGoalCol = [];
      for (var pp = 0; pp < products.length; pp++) {
        var ppc = perProdCols[products[pp]] || { prodCol: -1, goalCol: -1 };
        var pCellRaw = ppc.prodCol >= 0 ? (displayData && displayData[i] ? displayData[i][ppc.prodCol] : row[ppc.prodCol]) : '';
        prodVals.push(_splitSlashValues(pCellRaw)[0]);
        hasGoalCol.push(ppc.goalCol >= 0);
      }

      // Smart goal mapping: if fewer goal parts than products,
      // map goals left-to-right to products that have data (skipping empty ones).
      // This handles cases like Internet+DTV with no Wireless: "40/20" → Internet=40, DTV=20.
      var activeIndices = []; // product indices that have actual data
      for (var ai = 0; ai < prodVals.length; ai++) {
        if (!hasGoalCol[ai] && prodVals[ai] > 0) activeIndices.push(ai);
      }
      // If no products have data, fall back to all non-goal-col products
      if (!activeIndices.length) {
        for (var ai2 = 0; ai2 < products.length; ai2++) {
          if (!hasGoalCol[ai2]) activeIndices.push(ai2);
        }
      }

      // Second pass: assign goals
      var sharedGoalIdx = 0;
      for (var pp = 0; pp < products.length; pp++) {
        var ppc = perProdCols[products[pp]] || { prodCol: -1, goalCol: -1 };
        if (ppc.goalCol >= 0) {
          // Per-product goal column (e.g., LeafGuard Gross Leads Goal)
          var goalCellRaw = displayData && displayData[i] ? String(displayData[i][ppc.goalCol]) : String(row[ppc.goalCol] || '');
          goalVals.push(num(goalCellRaw));
        } else if (activeIndices.indexOf(pp) >= 0 && sharedGoalIdx < sharedGoalParts.length) {
          // Shared "/" goal mapped to this active product
          goalVals.push(sharedGoalParts[sharedGoalIdx++]);
        } else if (sharedGoalParts.length === 1 && pp === activeIndices[0]) {
          // Single goal value — assign to first active product
          goalVals.push(sharedGoalParts[0]);
          sharedGoalIdx++;
        } else {
          goalVals.push(0);
        }
      }
      prodRaw = prodVals.join('/');
      goalsRaw = goalVals.join('/');
    } else {
      // Standard: single production/goals column (may be "/" separated)
      if (displayData && displayData[i]) {
        prodRaw = (colMap.production >= 0) ? displayData[i][colMap.production] : '';
        goalsRaw = (colMap.goals >= 0) ? displayData[i][colMap.goals] : '';
      } else {
        prodRaw = (colMap.production >= 0) ? row[colMap.production] : '';
        goalsRaw = (colMap.goals >= 0) ? row[colMap.goals] : '';
      }
    }

    result.push({
      date: d,
      active: num(row[colMap.active]),
      leaders: num(row[colMap.leaders]),
      dist: num(row[colMap.dist]),
      closers: colMap.closers >= 0 ? num(row[colMap.closers]) : 0,
      leadGen: colMap.leadGen >= 0 ? num(row[colMap.leadGen]) : 0,
      training: num(row[colMap.training]),
      productionRaw: prodRaw,
      goalsRaw: goalsRaw
    });
  }

  // ── Sequential year correction for campaigns with no-year dates (e.g. Lumen "WE 3/20") ──
  // Only apply to campaigns whose source dates have no year component.
  // Dates are in chronological order. Walk backward from today and fix year rollovers.
  var NO_YEAR_CAMPAIGNS = ['lumen'];
  Logger.log('[YearFix-Health] campaignKey=' + campaignKey + ' resultLen=' + result.length + ' willRun=' + (campaignKey && NO_YEAR_CAMPAIGNS.indexOf(campaignKey) >= 0 && result.length >= 2));
  if (campaignKey && NO_YEAR_CAMPAIGNS.indexOf(campaignKey) >= 0 && result.length >= 2) {
    Logger.log('[YearFix-Health] ENTERING correction block');
    var now = new Date();
    var curYear = now.getFullYear();

    var last = result[result.length - 1].date;
    var candidateCur = new Date(curYear, last.getMonth(), last.getDate());
    var candidatePrev = new Date(curYear - 1, last.getMonth(), last.getDate());
    var fourWeeksAhead = new Date(now.getTime() + 28 * 86400000);
    if (candidateCur.getTime() <= fourWeeksAhead.getTime()) {
      result[result.length - 1].date = candidateCur;
    } else {
      result[result.length - 1].date = candidatePrev;
    }
    Logger.log('[YearFix-Health] anchor=' + result[result.length - 1].date.toLocaleDateString());

    for (var fix = result.length - 2; fix >= 0; fix--) {
      var curr = result[fix + 1].date;
      var prev = result[fix].date;
      var adjusted = new Date(curr.getFullYear(), prev.getMonth(), prev.getDate());
      if (adjusted.getTime() > curr.getTime()) {
        adjusted = new Date(curr.getFullYear() - 1, prev.getMonth(), prev.getDate());
      }
      if (adjusted.getTime() !== prev.getTime()) {
        Logger.log('[YearFix-Health] idx=' + fix + ' ' + prev.toLocaleDateString() + ' → ' + adjusted.toLocaleDateString());
      }
      result[fix].date = adjusted;
    }
    Logger.log('[YearFix-Health] correction complete');
  }

  return result;
}

/**
 * Extract recruiting data from HORIZONTAL columns in Section 1 (same row as health data).
 * Used when there is no separate Section 2 (e.g. Frontier tabs).
 * Returns: [{ date: Date, metrics: [12 values matching CONSOLIDATED_RECRUITING_HEADERS order] }]
 *
 * Expected source columns (after health columns):
 *   Calls Received / Applies Received, Sent to Call List / Sent to List,
 *   1st Rounds Booked, 1st Rounds Showed, Turned to 2nd / Retention, Conversion,
 *   2nd Rounds Booked, 2nd Rounds Showed, Retention,
 *   New Start Scheduled / NS Booked, New Starts Showed / NS Showed, Retention
 */
function extractHorizontalRecruitingRows_(data, start, end, campaignKey) {
  if (start < 0 || end < start) return [];

  // Find the actual header row — skip empty rows at the start
  var headerIdx = start;
  for (var hi = start; hi <= Math.min(start + 3, end); hi++) {
    var rowText = data[hi].map(function(c) { return String(c).trim(); }).join('');
    if (rowText.length > 0) { headerIdx = hi; break; }
  }

  var headers = data[headerIdx].map(function(h) { return String(h).toLowerCase().trim(); });

  Logger.log('extractHorizontalRecruiting_ headers: ' + JSON.stringify(headers));

  // Find recruiting columns by header name
  // Use nth-occurrence logic for "retention" (appears 3 times)
  var colCalls = findCol(headers, ['calls received', 'applies received', 'applies']);
  var colSentToList = findCol(headers, ['sent to call list', 'sent to list', 'no list']);
  var col1stBooked = findCol(headers, ['1st rounds booked', '1st booked', 'first booked', 'booked from call']);
  var col1stShowed = findCol(headers, ['1st rounds showed', '1st showed', 'first showed', 'rounds showed']);
  var colConversion = findCol(headers, ['conversion', 'turned to 2nd', '% of call']);

  // Find 2nd round columns (must contain "2nd")
  var col2ndBooked = -1, col2ndShowed = -1;
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].indexOf('2nd') >= 0 && headers[i].indexOf('book') >= 0) col2ndBooked = i;
    if (headers[i].indexOf('2nd') >= 0 && headers[i].indexOf('show') >= 0) col2ndShowed = i;
  }

  // Find NS columns
  var colNSBooked = -1, colNSShowed = -1;
  for (var i = 0; i < headers.length; i++) {
    if ((headers[i].indexOf('new start') >= 0 && (headers[i].indexOf('schedul') >= 0 || headers[i].indexOf('book') >= 0)) ||
        (headers[i].indexOf('ns') >= 0 && headers[i].indexOf('book') >= 0)) colNSBooked = i;
    if ((headers[i].indexOf('new start') >= 0 && headers[i].indexOf('show') >= 0) ||
        (headers[i].indexOf('ns') >= 0 && headers[i].indexOf('show') >= 0)) colNSShowed = i;
  }

  // Find retention columns (up to 3 occurrences in order)
  var retentionCols = [];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].indexOf('retention') >= 0) retentionCols.push(i);
  }
  var colRetention1 = retentionCols.length > 0 ? retentionCols[0] : -1;
  var colRetention2 = retentionCols.length > 1 ? retentionCols[1] : -1;
  var colRetention3 = retentionCols.length > 2 ? retentionCols[2] : -1;

  // Check if we found any recruiting columns at all
  var hasAny = col1stBooked >= 0 || col1stShowed >= 0 || col2ndBooked >= 0 || colNSBooked >= 0;
  Logger.log('extractHorizontalRecruiting_ cols: 1stBooked=' + col1stBooked + ' 1stShowed=' + col1stShowed + ' conv=' + colConversion + ' 2ndBooked=' + col2ndBooked + ' hasAny=' + hasAny);
  if (!hasAny) return [];

  var dateCol = findCol(headers, ['dates', 'date']);
  if (dateCol < 0) return [];

  var result = [];
  for (var i = headerIdx + 1; i <= end; i++) {
    var row = data[i];
    var dateVal = row[dateCol];
    if (!dateVal) continue;
    var d = (dateVal instanceof Date) ? dateVal : _parseTabDate(String(dateVal));
    if (!d) continue;

    var metrics = [
      num(colCalls >= 0 ? row[colCalls] : 0),           // Calls Received
      num(colSentToList >= 0 ? row[colSentToList] : 0),  // Sent to List
      num(col1stBooked >= 0 ? row[col1stBooked] : 0),    // 1st Booked
      num(col1stShowed >= 0 ? row[col1stShowed] : 0),    // 1st Showed
      _pctNum(colRetention1 >= 0 ? row[colRetention1] : 0), // 1st Retention
      _pctNum(colConversion >= 0 ? row[colConversion] : 0), // Conversion
      num(col2ndBooked >= 0 ? row[col2ndBooked] : 0),    // 2nd Booked
      num(col2ndShowed >= 0 ? row[col2ndShowed] : 0),    // 2nd Showed
      _pctNum(colRetention2 >= 0 ? row[colRetention2] : 0), // 2nd Retention
      num(colNSBooked >= 0 ? row[colNSBooked] : 0),      // NS Booked
      num(colNSShowed >= 0 ? row[colNSShowed] : 0),      // NS Showed
      _pctNum(colRetention3 >= 0 ? row[colRetention3] : 0)  // NS Retention
    ];

    var hasData = metrics.some(function(v) { return v !== 0; });
    if (hasData) {
      result.push({ date: d, metrics: metrics });
    }
  }

  // ── Sequential year correction for recruiting (same as health) ──
  var NO_YEAR_CAMPAIGNS_R = ['lumen'];
  if (campaignKey && NO_YEAR_CAMPAIGNS_R.indexOf(campaignKey) >= 0 && result.length >= 2) {
    var now = new Date();
    var curYear = now.getFullYear();
    var last = result[result.length - 1].date;
    var candidateCur = new Date(curYear, last.getMonth(), last.getDate());
    var fourWeeksAhead = new Date(now.getTime() + 28 * 86400000);
    result[result.length - 1].date = (candidateCur.getTime() <= fourWeeksAhead.getTime()) ? candidateCur : new Date(curYear - 1, last.getMonth(), last.getDate());
    for (var fix = result.length - 2; fix >= 0; fix--) {
      var curr = result[fix + 1].date;
      var prev = result[fix].date;
      var adjusted = new Date(curr.getFullYear(), prev.getMonth(), prev.getDate());
      if (adjusted.getTime() > curr.getTime()) adjusted = new Date(curr.getFullYear() - 1, prev.getMonth(), prev.getDate());
      result[fix].date = adjusted;
    }
  }

  return result;
}

/**
 * Extract recruiting (Section 2) data by transposing metric-rows × date-columns.
 * Returns: [{ date: Date, metrics: [12 values matching RECRUITING_LABELS order] }]
 *
 * Section 2 layout in source sheets:
 *   Row with "Actual"/"Projected" has week dates in columns B, C, D...
 *   Below that: metric labels in col A, values in corresponding date columns.
 */
function extractRecruitingRows_(data, start, end) {
  if (start < 0 || end < start) return [];

  // Step 1: Find the row with date columns
  var dateRowIdx = -1;
  var dateCols = []; // [{colIdx, date}]

  // Scan first few rows of section for date-like values in columns B+
  for (var i = start; i <= Math.min(start + 4, end); i++) {
    var foundDates = [];
    for (var j = 1; j < data[i].length; j++) {
      var cellVal = data[i][j];
      var d = null;
      if (cellVal instanceof Date) {
        d = cellVal;
      } else {
        var s = String(cellVal || '').trim();
        if (s) d = _parseTabDate(s);
      }
      if (d) foundDates.push({ colIdx: j, date: d });
    }
    if (foundDates.length >= 2) {
      dateRowIdx = i;
      dateCols = foundDates;
      break;
    }
  }

  if (dateCols.length === 0) return [];

  // Step 2: Find metric rows below the date row
  var metricsMap = {};
  var metricsStart = dateRowIdx + 1;

  for (var i = metricsStart; i <= end; i++) {
    var label = String(data[i][0] || '').trim().toLowerCase();
    if (!label) continue;

    if (label.indexOf('calls received') >= 0 || label.indexOf('applies') >= 0) metricsMap.callsReceived = data[i];
    else if (label.indexOf('no list') >= 0 || label.indexOf('sent to call') >= 0 || label.indexOf('sent to list') >= 0) metricsMap.noList = data[i];
    else if ((label === 'booked' || (label.indexOf('1st') >= 0 && label.indexOf('book') >= 0) || label.indexOf('first booked') >= 0 || label.indexOf('booked from call') >= 0) && !metricsMap.booked) metricsMap.booked = data[i];
    else if ((label === 'showed' || (label.indexOf('1st') >= 0 && label.indexOf('show') >= 0) || label.indexOf('first showed') >= 0) && !metricsMap.showed) metricsMap.showed = data[i];
    else if (label.indexOf('retention') >= 0 && !metricsMap.retention1) metricsMap.retention1 = data[i];
    else if (label.indexOf('conversion') >= 0 || label.indexOf('% call') >= 0 || label.indexOf('% of call') >= 0 || label.indexOf('turned to 2nd') >= 0) metricsMap.conversion = data[i];
    else if (label.indexOf('2nd') >= 0 && label.indexOf('book') >= 0) metricsMap.booked2 = data[i];
    else if (label.indexOf('2nd') >= 0 && label.indexOf('show') >= 0) metricsMap.showed2 = data[i];
    else if (label.indexOf('retention') >= 0 && metricsMap.retention1 && !metricsMap.retention2) metricsMap.retention2 = data[i];
    else if (label.indexOf('start') >= 0 && (label.indexOf('book') >= 0 || label.indexOf('schedul') >= 0)) metricsMap.startsBooked = data[i];
    else if (label.indexOf('start') >= 0 && label.indexOf('show') >= 0) metricsMap.startsShowed = data[i];
    else if ((label.indexOf('start') >= 0 && label.indexOf('retention') >= 0) || (label.indexOf('new start') >= 0 && label.indexOf('retention') >= 0)) metricsMap.startRetention = data[i];
    else if (label.indexOf('retention') >= 0 && metricsMap.retention1 && metricsMap.retention2 && !metricsMap.retention3) metricsMap.retention3 = data[i];
  }

  // Step 3: Transpose — for each date column, collect all metric values
  var result = [];
  for (var d = 0; d < dateCols.length; d++) {
    var colIdx = dateCols[d].colIdx;
    var date = dateCols[d].date;

    var metrics = [
      num(metricsMap.callsReceived ? metricsMap.callsReceived[colIdx] : 0),
      num(metricsMap.noList ? metricsMap.noList[colIdx] : 0),
      num(metricsMap.booked ? metricsMap.booked[colIdx] : 0),
      num(metricsMap.showed ? metricsMap.showed[colIdx] : 0),
      _pctNum(metricsMap.retention1 ? metricsMap.retention1[colIdx] : 0),
      _pctNum(metricsMap.conversion ? metricsMap.conversion[colIdx] : 0),
      num(metricsMap.booked2 ? metricsMap.booked2[colIdx] : 0),
      num(metricsMap.showed2 ? metricsMap.showed2[colIdx] : 0),
      _pctNum(metricsMap.retention2 ? metricsMap.retention2[colIdx] : 0),
      num(metricsMap.startsBooked ? metricsMap.startsBooked[colIdx] : 0),
      num(metricsMap.startsShowed ? metricsMap.startsShowed[colIdx] : 0),
      _pctNum((metricsMap.startRetention || metricsMap.retention3) ?
        (metricsMap.startRetention || metricsMap.retention3)[colIdx] : 0)
    ];

    // Only include if there's at least some non-zero data
    var hasData = metrics.some(function(v) { return v !== 0; });
    if (hasData) {
      result.push({ date: date, metrics: metrics });
    }
  }

  // ── Sequential year correction for campaigns with no-year dates (e.g. Lumen) ──
  var NO_YEAR_CAMPAIGNS = ['lumen'];
  if (campaignKey && NO_YEAR_CAMPAIGNS.indexOf(campaignKey) >= 0 && result.length >= 2) {
    var now = new Date();
    var curYear = now.getFullYear();
    var last = result[result.length - 1].date;
    var candidateCur = new Date(curYear, last.getMonth(), last.getDate());
    var candidatePrev = new Date(curYear - 1, last.getMonth(), last.getDate());
    var fourWeeksAhead = new Date(now.getTime() + 28 * 86400000);
    if (candidateCur.getTime() <= fourWeeksAhead.getTime()) {
      result[result.length - 1].date = candidateCur;
    } else {
      result[result.length - 1].date = candidatePrev;
    }
    for (var fix = result.length - 2; fix >= 0; fix--) {
      var curr = result[fix + 1].date;
      var prev = result[fix].date;
      var adjusted = new Date(curr.getFullYear(), prev.getMonth(), prev.getDate());
      if (adjusted.getTime() > curr.getTime()) {
        adjusted = new Date(curr.getFullYear() - 1, prev.getMonth(), prev.getDate());
      }
      result[fix].date = adjusted;
    }
  }

  return result;
}

/**
 * Split a "/" separated cell value into an array of numbers.
 * e.g. "5/3/2" → [5, 3, 2],  "10" → [10],  "" → [0]
 * Handles numbers stored as actual numbers (not strings).
 */
function _splitSlashValues(raw) {
  if (raw === null || raw === undefined || raw === '') return [0];
  // If it's a plain number, return as single element
  if (typeof raw === 'number') return [raw];
  var str = String(raw).trim();
  if (!str) return [0];
  // Check if it contains "/" separator
  if (str.indexOf('/') >= 0) {
    return str.split('/').map(function(v) { return num(v.trim()); });
  }
  return [num(str)];
}

/**
 * Merge health rows and recruiting rows by date.
 * campaignKey determines how many product columns to produce.
 * Produces flat arrays matching getConsolidatedHeaders_(campaignKey).
 */
function mergeHealthRecruiting_(ownerName, healthRows, recruitingRows, campaignKey) {
  var products = CAMPAIGN_PRODUCTS[campaignKey] || ['Total'];
  var dateMap = {};

  for (var i = 0; i < healthRows.length; i++) {
    var key = _normalizeDateKey_(healthRows[i].date, campaignKey);
    if (!dateMap[key]) dateMap[key] = { date: healthRows[i].date };
    // Only overwrite health if new row has more data (avoid zero row clobbering real data)
    var existing = dateMap[key].health;
    if (!existing || (healthRows[i].active || 0) + (healthRows[i].leaders || 0) >= (existing.active || 0) + (existing.leaders || 0)) {
      dateMap[key].health = healthRows[i];
    }
  }

  for (var i = 0; i < recruitingRows.length; i++) {
    var key = _normalizeDateKey_(recruitingRows[i].date, campaignKey);
    if (!dateMap[key]) dateMap[key] = { date: recruitingRows[i].date };
    // Only overwrite recruiting if new row has more data
    var existingR = dateMap[key].recruiting;
    if (!existingR) {
      dateMap[key].recruiting = recruitingRows[i];
    } else {
      var newSum = recruitingRows[i].metrics.reduce(function(a, b) { return a + b; }, 0);
      var oldSum = existingR.metrics.reduce(function(a, b) { return a + b; }, 0);
      if (newSum > oldSum) dateMap[key].recruiting = recruitingRows[i];
    }
  }

  // Fuzzy date matching (±3 days) — merge nearby entries that split due to timezone/parsing
  var dateKeys = Object.keys(dateMap);
  // Pass 1: health without recruiting → find nearby recruiting
  for (var i = 0; i < dateKeys.length; i++) {
    var entry = dateMap[dateKeys[i]];
    if (entry.health && !entry.recruiting) {
      for (var j = 0; j < recruitingRows.length; j++) {
        var diff = Math.abs(entry.date.getTime() - recruitingRows[j].date.getTime());
        if (diff <= 3 * 86400000) {
          entry.recruiting = recruitingRows[j];
          break;
        }
      }
    }
  }
  // Pass 2: recruiting without health → find nearby health entry and merge into it, then delete orphan
  dateKeys = Object.keys(dateMap);
  for (var i = 0; i < dateKeys.length; i++) {
    var entry = dateMap[dateKeys[i]];
    if (entry.recruiting && !entry.health) {
      // Look for a nearby health entry that already absorbed this recruiting, or merge into one
      var merged = false;
      for (var j = 0; j < dateKeys.length; j++) {
        if (i === j) continue;
        var other = dateMap[dateKeys[j]];
        if (!other.health) continue;
        var diff = Math.abs(entry.date.getTime() - other.date.getTime());
        if (diff <= 3 * 86400000) {
          // Merge recruiting into the health entry
          if (!other.recruiting) other.recruiting = entry.recruiting;
          delete dateMap[dateKeys[i]];
          merged = true;
          break;
        }
      }
    }
  }

  var result = [];
  var keys2 = Object.keys(dateMap);
  for (var i = 0; i < keys2.length; i++) {
    var entry = dateMap[keys2[i]];
    var h = entry.health || {};
    var r = entry.recruiting ? entry.recruiting.metrics : [0,0,0,0,0,0,0,0,0,0,0,0];

    // ── Split production/goals by "/" into per-product values ──
    var prodParts = _splitSlashValues(h.productionRaw);
    var goalParts = _splitSlashValues(h.goalsRaw);

    // Build the row: base columns first
    // Use snapped date so all owners share the same week date in consolidated tab
    // (Sunday for most campaigns, Monday for att-b2b)
    var snapKey = keys2[i]; // already "MM/DD/YYYY" format from _normalizeDateKey_
    var snapParts = snapKey.split('/');
    var snappedDate = new Date(Number(snapParts[2]), Number(snapParts[0]) - 1, Number(snapParts[1]), 12, 0, 0);
    var row = [
      snappedDate,         // 0: Week (Sunday or Monday of week)
      ownerName,           // 1: Owner
      h.active || 0,       // 2: Active HC
      h.leaders || 0,      // 3: Leaders
      h.dist || 0,         // 4: Dist
      h.training || 0      // 5: Training
    ];
    // Campaign-specific extra headcount columns
    var extraHC = CAMPAIGN_EXTRA_HC[campaignKey] || [];
    for (var ei = 0; ei < extraHC.length; ei++) {
      var ehName = extraHC[ei].toLowerCase().replace(/\s+/g, '');
      if (ehName === 'closers') row.push(h.closers || 0);
      else if (ehName === 'leadgen') row.push(h.leadGen || 0);
      else row.push(0);
    }

    // Per-product Prod pairs (goals go on NEXT week's row — separated below)
    var prodTotal = 0;
    for (var p = 0; p < products.length; p++) {
      var prodVal = (p < prodParts.length) ? prodParts[p] : 0;
      row.push(prodVal);
      row.push(0); // goal placeholder — actual goals written to next week's row below
      prodTotal += prodVal;
    }
    var goalTotal = goalParts.reduce(function(s, v) { return s + v; }, 0);

    // Stash goals to be placed on next week's row after the main loop
    if (goalTotal > 0) {
      var nextWeekDate = new Date(snappedDate.getTime() + 7 * 86400000);
      var nextKey = _normalizeDateKey_(nextWeekDate, campaignKey);
      if (!entry._goalForward) entry._goalForward = { nextKey: nextKey, nextDate: nextWeekDate, goalParts: goalParts };
    }
    // Recruiting metrics
    for (var ri = 0; ri < 12; ri++) {
      row.push(r[ri]);
    }

    result.push(row);
  }

  // ── Second pass: forward goals to next week's row ──
  // Goals from source row dated 3/15 (shifted to 3/22) belong on consolidated row 3/29.
  var extraHCCount = (CAMPAIGN_EXTRA_HC[campaignKey] || []).length;
  var baseCols = 6 + extraHCCount; // Week, Owner, Active, Leaders, Dist, Training + extras
  for (var i = 0; i < keys2.length; i++) {
    var entry = dateMap[keys2[i]];
    if (!entry._goalForward) continue;
    var fwd = entry._goalForward;

    // Find or create the next week's row in results
    var found = false;
    for (var ri = 0; ri < result.length; ri++) {
      var rKey = _normalizeDateKey_(result[ri][0], campaignKey);
      if (rKey === fwd.nextKey) {
        // Merge goals into this row's goal columns (don't overwrite non-zero)
        for (var gp = 0; gp < products.length; gp++) {
          var goalIdx = baseCols + gp * 2 + 1; // goal column for this product
          var existing = Number(result[ri][goalIdx]) || 0;
          var fwdGoal = (gp < fwd.goalParts.length) ? fwd.goalParts[gp] : 0;
          if (!existing && fwdGoal) result[ri][goalIdx] = fwdGoal;
        }
        found = true;
        break;
      }
    }

    if (!found) {
      // Create a new row for the goal week (no production data, just goals)
      var goalRow = [fwd.nextDate, ownerName];
      for (var gi = 2; gi < baseCols; gi++) goalRow.push(0); // HC zeros
      for (var gp = 0; gp < products.length; gp++) {
        goalRow.push(0); // prod = 0
        goalRow.push((gp < fwd.goalParts.length) ? fwd.goalParts[gp] : 0); // goal
      }
      for (var ri2 = 0; ri2 < 12; ri2++) goalRow.push(0); // recruiting zeros
      result.push(goalRow);
    }
  }

  return result;
}

/** Normalize date to key string for dedup */
/**
 * Fuzzy-match an owner name to a tab name in the spreadsheet.
 * Tries multiple strategies:
 *  1. Contains: tab contains owner name or vice versa
 *  2. First + Last name match: owner "Adam Cole" matches tab "Adam C" or "A. Cole"
 *  3. Token overlap: >50% of name tokens match
 * Returns the Sheet object or null.
 */
function _fuzzyFindTab_(ownerName, allTabNames, tabByName) {
  var ownerLower = ownerName.toLowerCase().trim();
  var ownerTokens = ownerLower.split(/\s+/);

  var bestTab = null;
  var bestScore = 0;

  for (var i = 0; i < allTabNames.length; i++) {
    var tabName = allTabNames[i];
    var tabLower = tabName.toLowerCase().trim();

    // Skip if tab is the totals/summary tab
    if (tabLower.indexOf('total') >= 0 || tabLower.indexOf('template') >= 0) continue;

    var score = 0;

    // Strategy 1: Contains match — longer overlap scores higher
    if (tabLower.indexOf(ownerLower) >= 0) {
      // Tab contains full owner name — perfect
      score = 95;
    } else if (ownerLower.indexOf(tabLower) >= 0) {
      // Owner name contains tab name — score by coverage (longer tab name = better match)
      // "Christian E" (11 chars) scores higher than "Christian" (9 chars) for "Christian Esposito"
      score = 50 + Math.round(tabLower.length / ownerLower.length * 40);
    }

    // Strategy 2: First name + last initial, or last name match
    if (score === 0 && ownerTokens.length >= 2) {
      var tabTokens = tabLower.split(/\s+/);
      var firstName = ownerTokens[0];
      var lastName = ownerTokens[ownerTokens.length - 1];

      // "Adam Cole" matches "Adam C" or "Adam C."
      if (tabTokens.length >= 2 && tabTokens[0] === firstName) {
        var tabLast = tabTokens[tabTokens.length - 1].replace(/\./g, '');
        if (tabLast === lastName || tabLast === lastName.charAt(0)) {
          score = 80;
        }
      }
      // "Adam Cole" matches "A Cole" or "A. Cole"
      if (score === 0 && tabTokens.length >= 2) {
        var tabFirst = tabTokens[0].replace(/\./g, '');
        var tabLastFull = tabTokens[tabTokens.length - 1];
        if ((tabFirst === firstName.charAt(0) || tabFirst === firstName) && tabLastFull === lastName) {
          score = 80;
        }
      }
    }

    // Strategy 3: Token overlap (>= 50% of owner tokens found in tab name)
    if (score === 0 && ownerTokens.length >= 2) {
      var matches = 0;
      for (var t = 0; t < ownerTokens.length; t++) {
        if (tabLower.indexOf(ownerTokens[t]) >= 0) matches++;
      }
      var overlap = matches / ownerTokens.length;
      if (overlap >= 0.5) {
        score = Math.round(overlap * 60);
      }
    }

    if (score > bestScore) {
      bestScore = score;
      bestTab = tabByName[tabLower] || null;
    }
  }

  // Only return if we have a decent match (score >= 30)
  return bestScore >= 30 ? bestTab : null;
}

function _normalizeDateKey_(d, campaignKey) {
  if (!d) return '';
  if (!(d instanceof Date)) d = new Date(d);

  // AT&T B2B uses Monday-lock dates — snap to Monday instead of Sunday
  if (campaignKey === 'att-b2b') {
    var day = d.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
    var offset = day === 0 ? -6 : 1 - day; // Sun→prev Mon, Mon→0, Tue→-1, ... Sat→-5
    var mon = new Date(d.getFullYear(), d.getMonth(), d.getDate() + offset, 12, 0, 0);
    return ('0' + (mon.getMonth() + 1)).slice(-2) + '/' + ('0' + mon.getDate()).slice(-2) + '/' + mon.getFullYear();
  }

  // Default: snap to Sunday of the same week (source spreadsheets use Sunday dates)
  var day = d.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
  var offset = -day; // Sun→0, Mon→-1, Tue→-2, ... Sat→-6
  var sun = new Date(d.getFullYear(), d.getMonth(), d.getDate() + offset, 12, 0, 0);
  return ('0' + (sun.getMonth() + 1)).slice(-2) + '/' + ('0' + sun.getDate()).slice(-2) + '/' + sun.getFullYear();
}

/**
 * Read consolidated per-campaign tabs and return data
 * in the same format the frontend expects:
 * { campaigns: { slug: { label, owners[], weeks[{ tabName, data }] } } }
 */
function readConsolidatedRecruiting(weekCount, campaignFilter, bustCache) {
  weekCount = weekCount || 6;

  var cache = CacheService.getScriptCache();

  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    Logger.log('readConsolidatedRecruiting error: ' + err.message);
    return { campaigns: {}, error: 'Cannot open national sheet' };
  }

  var campaigns = {};
  var keys = Object.keys(OD_CAMPAIGNS);

  for (var k = 0; k < keys.length; k++) {
    var key = keys[k];
    var campaign = OD_CAMPAIGNS[key];
    if (!campaign.sheetId) continue;
    if (campaignFilter && key !== campaignFilter) continue;

    // ── Per-campaign cache check (supports chunked storage) ──
    var campaignCacheKey = 'cache_recruiting_' + key;
    if (!bustCache) {
      try {
        var cached = _cacheRead(cache, campaignCacheKey);
        if (cached) {
          campaigns[key] = JSON.parse(cached);
          Logger.log('readConsolidatedRecruiting cache HIT: ' + key);
          continue;
        }
      } catch (e) { /* parse error — fall through to fresh read */ }
    }

    var tab = ss.getSheetByName(campaign.label);
    if (!tab) {
      campaigns[key] = { label: campaign.label, owners: [], weeks: [] };
      continue;
    }

    var data = tab.getDataRange().getValues();
    if (data.length < 2) {
      campaigns[key] = { label: campaign.label, owners: [], weeks: [] };
      continue;
    }

    // Parse headers (row 0)
    var headers = data[0].map(function(h) { return String(h).toLowerCase().trim(); });
    var colWeek = findCol(headers, ['week']);
    var colOwner = findCol(headers, ['owner']);
    var colCalls = findCol(headers, ['calls received', 'applies']);
    var colSTL = findCol(headers, ['sent to list', 'no list']);
    var col1B = findCol(headers, ['1st booked']);
    var col1S = findCol(headers, ['1st showed']);
    var col1R = findCol(headers, ['1st retention']);
    var colConv = findCol(headers, ['conversion']);
    var col2B = findCol(headers, ['2nd booked']);
    var col2S = findCol(headers, ['2nd showed']);
    var col2R = findCol(headers, ['2nd retention']);
    var colNSB = findCol(headers, ['ns booked']);
    var colNSS = findCol(headers, ['ns showed']);
    var colNSR = findCol(headers, ['ns retention']);
    var colHC = findCol(headers, ['active hc', 'active']);
    var colLeaders = findCol(headers, ['leaders']);
    var colDist = findCol(headers, ['dist']);
    var colClosers = findCol(headers, ['closers']);
    var colLeadGen = findCol(headers, ['lead gen', 'leadgen']);
    var colTraining = findCol(headers, ['training']);
    // ── Find per-product production/goal columns ──
    // New format: "prod: frontier", "goal: frontier", "prod: cell", etc.
    // Fallback: old single "production" / "goals" columns
    var products = CAMPAIGN_PRODUCTS[key] || ['Total'];
    var prodCols = []; // [{product, prodCol, goalCol}]
    for (var pi = 0; pi < products.length; pi++) {
      var pName = products[pi].toLowerCase();
      var pc = findCol(headers, ['prod: ' + pName]);
      var gc = findCol(headers, ['goal: ' + pName]);
      if (pc >= 0 || gc >= 0) {
        prodCols.push({ product: products[pi], prodCol: pc, goalCol: gc });
      }
    }
    // (Total columns removed — each product tracked individually)
    // Fallback: old single-column format
    var colProdLegacy = findCol(headers, ['production']);
    var colGoalsLegacy = findCol(headers, ['goals']);

    // Group by week date
    var weekMap = {};
    var ownerSet = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var dateVal = row[colWeek];
      var owner = String(row[colOwner] || '').trim();
      if (!dateVal || !owner) continue;

      // Snap to week boundary — Sunday for most campaigns, Monday for att-b2b.
      // Must pass campaign key so att-b2b Monday-lock is respected here too.
      var dateKey = _normalizeDateKey_(dateVal, key);
      if (!dateKey) continue;
      ownerSet[owner] = true;

      if (!weekMap[dateKey]) {
        var parts = dateKey.split('/');
        weekMap[dateKey] = {
          tabName: dateKey,
          date: new Date(Number(parts[2]), Number(parts[0]) - 1, Number(parts[1]), 12, 0, 0),
          data: {}
        };
      }

      // 12-value recruiting metrics
      var metrics = [
        num(row[colCalls]),
        num(row[colSTL]),
        num(row[col1B]),
        num(row[col1S]),
        _pctNum(row[col1R]),
        _pctNum(row[colConv]),
        num(row[col2B]),
        num(row[col2S]),
        _pctNum(row[col2R]),
        num(row[colNSB]),
        num(row[colNSS]),
        _pctNum(row[colNSR])
      ];

      // Build per-product production/goals object
      var productionData = {};
      if (prodCols.length > 0) {
        // New per-product columns
        for (var pi = 0; pi < prodCols.length; pi++) {
          var p = prodCols[pi];
          productionData[p.product] = {
            production: p.prodCol >= 0 ? num(row[p.prodCol]) : 0,
            goals: p.goalCol >= 0 ? num(row[p.goalCol]) : 0
          };
        }
        // (No Total row — products tracked individually)
      } else if (colProdLegacy >= 0 || colGoalsLegacy >= 0) {
        // Legacy single-column: try splitting "/" values
        var prodParts = _splitSlashValues(row[colProdLegacy >= 0 ? colProdLegacy : 0]);
        var goalParts = _splitSlashValues(row[colGoalsLegacy >= 0 ? colGoalsLegacy : 0]);
        for (var pi = 0; pi < products.length; pi++) {
          productionData[products[pi]] = {
            production: pi < prodParts.length ? prodParts[pi] : 0,
            goals: pi < goalParts.length ? goalParts[pi] : 0
          };
        }
      }

      // Store as object so health survives JSON.stringify
      // (array properties are silently dropped by JSON.stringify)
      weekMap[dateKey].data[owner] = {
        metrics: metrics,
        health: {
          active: num(row[colHC]),
          leaders: num(row[colLeaders]),
          dist: num(row[colDist]),
          closers: colClosers >= 0 ? num(row[colClosers]) : 0,
          leadGen: colLeadGen >= 0 ? num(row[colLeadGen]) : 0,
          training: num(row[colTraining]),
          production: productionData,
          goals: productionData
        }
      };
    }

    // Filter out future dates (corrupted data) and sort by date descending
    var now = new Date();
    var cutoff = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 7); // allow up to 1 week ahead
    var weekEntries = Object.keys(weekMap).map(function(wk) { return weekMap[wk]; })
      .filter(function(w) { return w.date && w.date.getTime() <= cutoff.getTime(); });
    weekEntries.sort(function(a, b) {
      var da = a.date ? a.date.getTime() : 0;
      var db = b.date ? b.date.getTime() : 0;
      return db - da;
    });
    // No week limit — send all history so WoW health/production cards have full data.
    // The frontend recruiting table handles its own 4-week column display independently.

    campaigns[key] = {
      label: campaign.label,
      owners: Object.keys(ownerSet).sort(),
      weeks: weekEntries,
      products: CAMPAIGN_PRODUCTS[key] || ['Total']
    };

    // ── Store per-campaign cache (chunked if needed) ──
    try {
      var cJson = JSON.stringify(campaigns[key]);
      _cacheStore(cache, campaignCacheKey, cJson, 300);
    } catch (e) { Logger.log('readConsolidatedRecruiting cache store error: ' + e.message); }
  }

  // B2B: include Tableau product breakdown grand totals
  var b2bProductTotals = null;
  try {
    var metricsSheet = ss.getSheetByName('B2B Sales Metrics');
    if (metricsSheet && metricsSheet.getLastRow() > 1) {
      var mData = metricsSheet.getDataRange().getDisplayValues();
      var mHeaders = mData[0];
      for (var r = mData.length - 1; r >= 1; r--) {
        if (String(mData[r][0]).trim() === 'Grand Total') {
          var iInt = mHeaders.indexOf('Internet');
          var iVoip = mHeaders.indexOf('VOIP');
          var iWrls = mHeaders.indexOf('Wireless');
          var iAir = mHeaders.indexOf('AIR/AWB');
          b2bProductTotals = {
            internet: iInt >= 0 ? parseInt(mData[r][iInt]) || 0 : 0,
            voip: iVoip >= 0 ? parseInt(mData[r][iVoip]) || 0 : 0,
            wireless: iWrls >= 0 ? parseInt(mData[r][iWrls]) || 0 : 0,
            air: iAir >= 0 ? parseInt(mData[r][iAir]) || 0 : 0
          };
          break;
        }
      }
    }
  } catch (e) {
    Logger.log('B2B product totals read error: ' + e.message);
  }

  return { campaigns: campaigns, b2bProductTotals: b2bProductTotals };
}

/**
 * Setup a daily trigger to refresh campaign data at 1 AM.
 * Run this ONCE from the script editor to install the trigger.
 */
function setupDailyRefreshTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshAllCampaigns') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('refreshAllCampaigns')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .nearMinute(0)
    .create();

  Logger.log('Daily refresh trigger installed: refreshAllCampaigns at 1 AM');
  return { ok: true, message: 'Trigger set: refreshAllCampaigns runs daily at 1 AM' };
}

/**
 * Setup refresh trigger for specific weekdays only.
 * @param {number[]} days - Array of weekday numbers (1=Mon, 2=Tue, ..., 7=Sun)
 * Run from script editor: setupWeekdayRefreshTrigger([1, 3, 5]) for Mon/Wed/Fri
 */
function setupWeekdayRefreshTrigger(days) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'scheduledRefreshCampaigns_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  if (days && days.length > 0) {
    PropertiesService.getScriptProperties().setProperty('REFRESH_DAYS', JSON.stringify(days));
  }

  ScriptApp.newTrigger('scheduledRefreshCampaigns_')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .nearMinute(0)
    .create();

  Logger.log('Weekday refresh trigger installed for days: ' + JSON.stringify(days));
  return { ok: true, message: 'Trigger set for days: ' + JSON.stringify(days) };
}

/**
 * Trigger handler that checks if today is an allowed refresh day.
 * @deprecated — replaced by scheduledPlanningRefresh_
 */
function scheduledRefreshCampaigns_() {
  var daysJson = PropertiesService.getScriptProperties().getProperty('REFRESH_DAYS');
  if (daysJson) {
    var days = JSON.parse(daysJson);
    var today = new Date().getDay();
    var todayMon = today === 0 ? 7 : today; // 1=Mon...7=Sun
    if (days.indexOf(todayMon) < 0) {
      Logger.log('scheduledRefreshCampaigns_: skipping (today=' + todayMon + ')');
      return;
    }
  }
  refreshAllCampaigns();
}

// ═══════════════════════════════════════════════════════
// PLANNING-BASED AUTO-REFRESH (replaces manual refresh)
// ═══════════════════════════════════════════════════════

/**
 * Setup the daily 1 AM trigger based on _OD_Planning schedule.
 * Run once from Apps Script editor to install.
 * Removes old triggers for refreshAllCampaigns / scheduledRefreshCampaigns_.
 */
function setupPlanningRefreshTrigger() {
  // Remove old triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var handler = triggers[i].getHandlerFunction();
    if (handler === 'refreshAllCampaigns' || handler === 'scheduledRefreshCampaigns_' || handler === 'scheduledPlanningRefresh_') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('scheduledPlanningRefresh_')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .nearMinute(0)
    .create();

  Logger.log('Planning refresh trigger installed: scheduledPlanningRefresh_ at 1 AM daily');
  return { ok: true, message: 'Trigger set: scheduledPlanningRefresh_ runs daily at 1 AM' };
}

/**
 * Daily 1 AM trigger handler.
 * Reads _OD_Planning to find which campaigns are scheduled for today,
 * then runs consolidateCampaignSlim_ for each.
 */
function scheduledPlanningRefresh_() {
  var jsDay = new Date().getDay();
  var todayIdx = (jsDay + 6) % 7; // Mon=0, Sun=6

  var planningResult = odGetPlanning();
  var planning = (planningResult && planningResult.planning) || [];
  var todayCampaigns = [];
  for (var i = 0; i < planning.length; i++) {
    if (planning[i].day === todayIdx) todayCampaigns.push(planning[i]);
  }

  if (!todayCampaigns.length) {
    Logger.log('scheduledPlanningRefresh_: no campaigns scheduled for today (dayIdx=' + todayIdx + ')');
    return;
  }

  Logger.log('scheduledPlanningRefresh_: refreshing ' + todayCampaigns.length + ' campaigns for dayIdx=' + todayIdx);

  var destSS;
  try {
    destSS = SpreadsheetApp.openById(SHEETS.NATIONAL);
  } catch (err) {
    Logger.log('scheduledPlanningRefresh_: cannot open national sheet: ' + err.message);
    return;
  }

  for (var i = 0; i < todayCampaigns.length; i++) {
    var key = todayCampaigns[i].campaignKey;
    var campaign = OD_CAMPAIGNS[key];
    if (!campaign || !campaign.sheetId) {
      Logger.log('scheduledPlanningRefresh_: skipping unknown/unconfigured campaign: ' + key);
      continue;
    }
    try {
      var count = consolidateCampaignSlim_(key, campaign, destSS);
      Logger.log('scheduledPlanningRefresh_: ' + key + ' → ' + count + ' rows');
    } catch (err) {
      Logger.log('scheduledPlanningRefresh_: ERROR for ' + key + ': ' + err.message + '\n' + err.stack);
    }
  }

  // Warm caches for today's scheduled campaigns + global caches
  var todayKeys = todayCampaigns.map(function(c) { return c.campaignKey; });
  Logger.log('scheduledPlanningRefresh_: warming caches for ' + todayKeys.join(', '));
  warmAllCaches(todayKeys);
}

// ══════════════════════════════════════════════════
// CACHE WARMING
// ══════════════════════════════════════════════════

/**
 * Warm all CacheService caches. Called by scheduled trigger and manually.
 * @param {string[]} [campaignKeys] - If provided, only warm these campaign recruiting caches.
 *                                    If omitted, warms ALL campaigns.
 */
function warmAllCaches(campaignKeys) {
  var start = new Date().getTime();

  // If no keys provided (e.g. triggered by 5-min timer), check today's planning schedule
  if (!campaignKeys) {
    try {
      var jsDay = new Date().getDay();
      var todayIdx = (jsDay + 6) % 7; // Mon=0, Sun=6
      var planningResult = odGetPlanning();
      var planning = (planningResult && planningResult.planning) || [];
      var todayKeys = [];
      for (var p = 0; p < planning.length; p++) {
        if (planning[p].day === todayIdx) todayKeys.push(planning[p].campaignKey);
      }
      if (todayKeys.length) {
        campaignKeys = todayKeys;
        Logger.log('warmAllCaches: planning-aware — today\'s campaigns: ' + todayKeys.join(', '));
      } else {
        // No campaigns scheduled today — only warm global caches, skip recruiting
        Logger.log('warmAllCaches: no campaigns scheduled today, warming globals only');
        campaignKeys = [];
      }
    } catch (e) {
      Logger.log('warmAllCaches: planning read error, warming all: ' + e.message);
      campaignKeys = Object.keys(OD_CAMPAIGNS);
    }
  }

  // Global caches (always warm)
  try {
    _cachedRead('cache_onlinePresence', 600, readOnlinePresence, true);
    Logger.log('warmAllCaches: onlinePresence OK');
  } catch (e) { Logger.log('warmAllCaches: onlinePresence ERROR: ' + e.message); }

  try {
    _cachedRead('cache_ownerCamMapping', 300, readOwnerCamMapping, true);
    Logger.log('warmAllCaches: ownerCamMapping OK');
  } catch (e) { Logger.log('warmAllCaches: ownerCamMapping ERROR: ' + e.message); }

  // Per-campaign recruiting caches (only today's scheduled campaigns)
  if (campaignKeys.length) {
    for (var c = 0; c < campaignKeys.length; c++) {
      try {
        readConsolidatedRecruiting(6, campaignKeys[c], true);
        Logger.log('warmAllCaches: recruiting ' + campaignKeys[c] + ' OK');
      } catch (e) { Logger.log('warmAllCaches: recruiting ' + campaignKeys[c] + ' ERROR: ' + e.message); }
    }
  }

  var elapsed = new Date().getTime() - start;
  Logger.log('warmAllCaches: done in ' + elapsed + 'ms (' + campaignKeys.length + ' campaigns)');
}

/**
 * Setup an every-5-minute trigger to keep caches warm throughout the day.
 * Run once from the Apps Script editor to install.
 */
function setupCacheWarmingTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'warmAllCaches') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('warmAllCaches')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('Cache warming trigger installed: warmAllCaches every 5 minutes');
  return { ok: true, message: 'Trigger set: warmAllCaches runs every 5 minutes' };
}

/**
 * Slim consolidation: production + recruiting only (3 week lookback).
 * Headcount columns are 0 unless previously manually entered (preserved).
 * Goal columns are always 0.
 */
function consolidateCampaignSlim_(campaignKey, campaign, destSS) {
  Logger.log('consolidateCampaignSlim_ START: key=' + campaignKey);
  var srcSS = SpreadsheetApp.openById(campaign.sheetId);
  var allTabs = srcSS.getSheets();

  // Get owner names
  var ownerNames = getOwnerNamesForCampaign_(campaign, srcSS);
  Logger.log('consolidateCampaignSlim_ owners (' + ownerNames.length + '): ' + ownerNames.join(', '));

  // Load saved tab mappings
  var savedTabMap = {};
  try {
    var mapTab = destSS.getSheetByName('_Campaign_Tab_Map');
    if (mapTab) {
      var mapData = mapTab.getDataRange().getValues();
      for (var m = 1; m < mapData.length; m++) {
        var mapCampaign = String(mapData[m][0] || '').trim().toLowerCase();
        var mapOwner    = String(mapData[m][1] || '').trim().toLowerCase();
        var mapTabName  = String(mapData[m][2] || '').trim();
        if (mapCampaign === campaignKey.toLowerCase() && mapOwner && mapTabName) {
          savedTabMap[mapOwner] = mapTabName;
        }
      }
    }
  } catch (e) { Logger.log('consolidateCampaignSlim_: tab map load error: ' + e.message); }

  // Build tab lookup
  var tabByName = {};
  var allTabNames = [];
  for (var t = 0; t < allTabs.length; t++) {
    var tn = allTabs[t].getName().trim();
    if (tn) tabByName[tn.toLowerCase()] = allTabs[t];
    if (tn && SKIP_TABS_.indexOf(tn.toLowerCase()) < 0) allTabNames.push(tn);
  }

  // ── Read existing headcount + goals + production from destination tab (to preserve entries) ──
  var existingHC = {}; // 'ownerLower|dateKey' → { active, leaders, dist, training, closers, leadGen }
  var existingGoals = {}; // 'ownerLower|dateKey' → { productName: goalValue, ... }
  var existingProd = {}; // 'ownerLower|dateKey' → { productName: prodValue, ... }
  var destTab = destSS.getSheetByName(campaign.label);
  var campaignHeaders = getConsolidatedHeaders_(campaignKey);
  if (destTab) {
    var existingData = destTab.getDataRange().getValues();
    if (existingData.length > 1) {
      var exHeaders = existingData[0].map(function(h) { return String(h).trim(); });
      var exColMap = {};
      for (var c = 0; c < exHeaders.length; c++) exColMap[exHeaders[c]] = c;

      // Identify goal and production columns
      var goalCols = []; // [{ colIdx, productName }]
      var prodCols = []; // [{ colIdx, productName }]
      for (var gc = 0; gc < exHeaders.length; gc++) {
        if (exHeaders[gc].indexOf('Goal: ') === 0) {
          goalCols.push({ colIdx: gc, productName: exHeaders[gc].replace('Goal: ', '') });
        }
        if (exHeaders[gc].indexOf('Prod: ') === 0) {
          prodCols.push({ colIdx: gc, productName: exHeaders[gc].replace('Prod: ', '') });
        }
      }

      for (var ei = 1; ei < existingData.length; ei++) {
        var exOwner = String(existingData[ei][exColMap['Owner']] || '').trim().toLowerCase();
        var exDate = existingData[ei][exColMap['Week']];
        if (!exOwner || !exDate) continue;
        var exDateKey = _normalizeDateKey_(exDate instanceof Date ? exDate : _parseTabDate(String(exDate)));

        // Check if any headcount value is non-zero (manually entered)
        var exActive = parseInt(existingData[ei][exColMap['Active HC']]) || 0;
        var exLeaders = parseInt(existingData[ei][exColMap['Leaders']]) || 0;
        var exDist = parseInt(existingData[ei][exColMap['Dist']]) || 0;
        var exTraining = parseInt(existingData[ei][exColMap['Training']]) || 0;
        var exClosers = exColMap['Closers'] !== undefined ? (parseInt(existingData[ei][exColMap['Closers']]) || 0) : 0;
        var exLeadGen = exColMap['Lead Gen'] !== undefined ? (parseInt(existingData[ei][exColMap['Lead Gen']]) || 0) : 0;

        if (exActive || exLeaders || exDist || exTraining || exClosers || exLeadGen) {
          existingHC[exOwner + '|' + exDateKey] = {
            active: exActive, leaders: exLeaders, dist: exDist, training: exTraining,
            closers: exClosers, leadGen: exLeadGen
          };
        }

        // Preserve any non-zero goals and production
        var ownerGoals = {};
        var ownerProd = {};
        var hasGoals = false;
        var hasProd = false;
        for (var gci = 0; gci < goalCols.length; gci++) {
          var gVal = parseInt(existingData[ei][goalCols[gci].colIdx]) || 0;
          if (gVal) {
            ownerGoals[goalCols[gci].productName] = gVal;
            hasGoals = true;
          }
        }
        for (var pci = 0; pci < prodCols.length; pci++) {
          var pVal = parseInt(existingData[ei][prodCols[pci].colIdx]) || 0;
          if (pVal) {
            ownerProd[prodCols[pci].productName] = pVal;
            hasProd = true;
          }
        }
        if (hasGoals) {
          existingGoals[exOwner + '|' + exDateKey] = ownerGoals;
        }
        if (hasProd) {
          existingProd[exOwner + '|' + exDateKey] = ownerProd;
        }
      }
    }
    Logger.log('consolidateCampaignSlim_ preserved ' + Object.keys(existingHC).length + ' headcount, ' + Object.keys(existingGoals).length + ' goal, ' + Object.keys(existingProd).length + ' production entries');
  }

  var products = CAMPAIGN_PRODUCTS[campaignKey] || ['Total'];
  var extraHC = CAMPAIGN_EXTRA_HC[campaignKey] || [];
  var rows = [];

  for (var oi = 0; oi < ownerNames.length; oi++) {
    var ownerName = ownerNames[oi];
    var ownerLower = ownerName.toLowerCase();

    // Find tab (same logic as consolidateCampaign_)
    var targetTabName = savedTabMap[ownerLower] || null;
    var tab = null;
    if (targetTabName) {
      if (targetTabName.toLowerCase() !== 'non-partner') {
        tab = tabByName[targetTabName.toLowerCase()] || null;
      }
    }
    if (!tab && !targetTabName) {
      // Exact name match only — no fuzzy matching (avoids grabbing wrong tab)
      tab = tabByName[ownerLower] || null;
    }

    if (!tab) {
      // No tab — still include owner with empty row (snap to Sunday so dates stay consistent)
      var now = new Date();
      var dayOfWeek = now.getDay(); // 0=Sun … 6=Sat
      var sunday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayOfWeek, 12, 0, 0);
      var emptyRow = [sunday, ownerName];
      for (var ei2 = 2; ei2 < campaignHeaders.length; ei2++) emptyRow.push(0);
      rows.push(emptyRow);
      continue;
    }

    try {
      var range = tab.getDataRange();
      var data = range.getValues();
      var displayData = range.getDisplayValues();
      if (!data || data.length < 2) { Logger.log('consolidateCampaignSlim_ SKIP ' + ownerName + ': <2 rows'); continue; }

      var sections = findSections(data);

      // Extract health (for production only) and recruiting
      var healthRows = extractHealthRows_(data, sections.section1Start, sections.section1End, displayData, campaignKey);
      var recruitingRows = extractHorizontalRecruitingRows_(data, sections.section1Start, sections.section1End, campaignKey);

      // Debug: log extraction results for owners with no data
      if (!healthRows.length && !recruitingRows.length) {
        Logger.log('consolidateCampaignSlim_ EMPTY ' + ownerName + ': sections=' + JSON.stringify(sections) + ' rows=' + data.length + ' tab=' + tab.getName());
      }

      // +7 day offset: campaigns whose source dates are 1 week behind consolidated dates.
      // LeafGuard skips entirely. AT&T B2B is mixed: some owners (e.g. Van's team) date
      // entries as the current Monday → no offset needed. Others (e.g. Ken's team) date
      // entries as the previous Sunday/Monday → +7 is needed to align to the correct week.
      // Auto-detect per-owner: if the most recent source date is a Monday, skip +7 (direct
      // Monday convention). If it's a Sunday or other day, apply +7.
      var applyOffset = true;
      if (campaignKey === 'leafguard') {
        applyOffset = false;
      } else if (campaignKey === 'att-b2b') {
        // Check the day-of-week of the most recent health or recruiting date
        var sampleDate = null;
        if (healthRows.length > 0) sampleDate = healthRows[healthRows.length - 1].date;
        else if (recruitingRows.length > 0) sampleDate = recruitingRows[recruitingRows.length - 1].date;
        if (sampleDate instanceof Date) {
          // Monday (day=1) → owner uses current-Monday convention → no offset
          // Sunday (day=0) or any other day → offset needed
          applyOffset = sampleDate.getDay() !== 1;
        } else {
          applyOffset = false; // no dates to sample — leave as-is
        }
        Logger.log('consolidateCampaignSlim_ att-b2b ' + ownerName + ': sampleDate=' + (sampleDate ? sampleDate.toDateString() : 'none') + ' applyOffset=' + applyOffset);
      }
      if (applyOffset) {
        var offset = 7 * 86400000;
        for (var hi = 0; hi < healthRows.length; hi++) {
          if (healthRows[hi].date instanceof Date) healthRows[hi].date = new Date(healthRows[hi].date.getTime() + offset);
        }
        for (var ri = 0; ri < recruitingRows.length; ri++) {
          if (recruitingRows[ri].date instanceof Date) recruitingRows[ri].date = new Date(recruitingRows[ri].date.getTime() + offset);
        }
      }

      // Merge health + recruiting by date
      var merged = mergeHealthRecruiting_(ownerName, healthRows, recruitingRows, campaignKey);

      // Sort by date descending and limit to 3 most recent weeks
      merged.sort(function(a, b) {
        var da = (a[0] instanceof Date) ? a[0].getTime() : 0;
        var db = (b[0] instanceof Date) ? b[0].getTime() : 0;
        return db - da;
      });
      if (merged.length > 3) merged = merged.slice(0, 3);

      // Guarantee a row exists for the current week so coaches can always enter headcount
      var now = new Date();
      var currentWeekDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      // Adjust to Monday of this week
      var dayOfWeek = (currentWeekDate.getDay() + 6) % 7; // Mon=0
      currentWeekDate.setDate(currentWeekDate.getDate() - dayOfWeek);
      var currentWeekKey = _normalizeDateKey_(currentWeekDate);
      var hasCurrentWeek = merged.some(function(m) { return _normalizeDateKey_(m[0]) === currentWeekKey; });
      if (!hasCurrentWeek) {
        // Build a blank row for this week
        var blankRow = [currentWeekDate, ownerName];
        for (var bi = 2; bi < campaignHeaders.length; bi++) blankRow.push(0);
        // Restore preserved headcount if any
        var blankHcKey = ownerLower + '|' + currentWeekKey;
        if (existingHC[blankHcKey]) {
          var bp = existingHC[blankHcKey];
          blankRow[2] = bp.active; blankRow[3] = bp.leaders; blankRow[4] = bp.dist; blankRow[5] = bp.training;
        }
        merged.unshift(blankRow); // newest first
        if (merged.length > 3) merged = merged.slice(0, 3);
      }

      // Zero out headcount + goals, but preserve existing manual entries
      for (var r = 0; r < merged.length; r++) {
        var row = merged[r];
        var rowDateKey = _normalizeDateKey_(row[0]);
        var hcKey = ownerLower + '|' + rowDateKey;
        var preserved = existingHC[hcKey];

        // Headcount: use preserved values or 0
        row[2] = preserved ? preserved.active : 0;   // Active HC
        row[3] = preserved ? preserved.leaders : 0;   // Leaders
        row[4] = preserved ? preserved.dist : 0;      // Dist
        row[5] = preserved ? preserved.training : 0;  // Training

        // Extra HC columns (Closers, Lead Gen for LeafGuard)
        var extraStart = 6;
        for (var exi = 0; exi < extraHC.length; exi++) {
          var exName = extraHC[exi].toLowerCase().replace(/\s+/g, '');
          if (exName === 'closers') row[extraStart + exi] = preserved ? preserved.closers : 0;
          else if (exName === 'leadgen') row[extraStart + exi] = preserved ? preserved.leadGen : 0;
          else row[extraStart + exi] = 0;
        }

        // Restore existing production and goals (don't overwrite Credico-synced values)
        var prodStart = 6 + extraHC.length;
        var rowDateKey = _normalizeDateKey_(row[0]);
        var savedGoals = existingGoals[ownerLower + '|' + rowDateKey];
        var savedProd = existingProd[ownerLower + '|' + rowDateKey];
        for (var pi = 0; pi < products.length; pi++) {
          // Production: keep existing if source is 0
          if (!row[prodStart + pi * 2] && savedProd && savedProd[products[pi]]) {
            row[prodStart + pi * 2] = savedProd[products[pi]];
          }
          // Goals: restore from existing
          row[prodStart + pi * 2 + 1] = (savedGoals && savedGoals[products[pi]]) ? savedGoals[products[pi]] : 0;
        }

        rows.push(row);
      }
    } catch (err) {
      Logger.log('consolidateCampaignSlim_ tab error (' + campaignKey + '/' + ownerName + '): ' + err.message);
    }
  }

  // Preserve older rows (outside the 3-week window) that have headcount, production, or goals
  var newDateOwnerKeys = {};
  for (var ri2 = 0; ri2 < rows.length; ri2++) {
    var rOwner = String(rows[ri2][1] || '').trim().toLowerCase();
    var rDate = _normalizeDateKey_(rows[ri2][0]);
    newDateOwnerKeys[rOwner + '|' + rDate] = true;
  }

  // Collect all unique keys that have any data worth preserving
  var allPreserveKeys = {};
  for (var hcK in existingHC) allPreserveKeys[hcK] = true;
  for (var prK in existingProd) allPreserveKeys[prK] = true;
  for (var glK in existingGoals) allPreserveKeys[glK] = true;

  for (var pKey in allPreserveKeys) {
    if (newDateOwnerKeys[pKey]) continue; // already covered by fresh 3-week data
    var parts = pKey.split('|');
    var pOwner = parts[0];
    var pDateParts = parts[1].split('/');
    var pDate = new Date(Number(pDateParts[2]), Number(pDateParts[0]) - 1, Number(pDateParts[1]), 12, 0, 0);
    // Find the original owner name (proper casing)
    var pOwnerName = pOwner;
    for (var oni = 0; oni < ownerNames.length; oni++) {
      if (ownerNames[oni].toLowerCase() === pOwner) { pOwnerName = ownerNames[oni]; break; }
    }
    var hcEntry = existingHC[pKey] || { active: 0, leaders: 0, dist: 0, training: 0, closers: 0, leadGen: 0 };
    var preservedRow = [pDate, pOwnerName, hcEntry.active, hcEntry.leaders, hcEntry.dist, hcEntry.training];
    for (var exi2 = 0; exi2 < extraHC.length; exi2++) {
      var exn = extraHC[exi2].toLowerCase().replace(/\s+/g, '');
      if (exn === 'closers') preservedRow.push(hcEntry.closers || 0);
      else if (exn === 'leadgen') preservedRow.push(hcEntry.leadGen || 0);
      else preservedRow.push(0);
    }
    // Restore production + goals
    var oldProd = existingProd[pKey];
    var oldGoals = existingGoals[pKey];
    for (var pi2 = 0; pi2 < products.length; pi2++) {
      preservedRow.push((oldProd && oldProd[products[pi2]]) ? oldProd[products[pi2]] : 0);
      preservedRow.push((oldGoals && oldGoals[products[pi2]]) ? oldGoals[products[pi2]] : 0);
    }
    for (var rci = 0; rci < 12; rci++) preservedRow.push(0); // recruiting
    rows.push(preservedRow);
  }

  // ── Dedup: merge duplicate owner+date rows (keep the one with most data) ──
  var dedupMap = {};
  var prodStartDedup = 6 + extraHC.length;
  for (var di = 0; di < rows.length; di++) {
    var dOwner = String(rows[di][1] || '').trim().toLowerCase();
    var dDate = _normalizeDateKey_(rows[di][0]);
    var dKey = dOwner + '|' + dDate;
    if (!dedupMap[dKey]) {
      dedupMap[dKey] = di;
    } else {
      // Merge: keep whichever row has more non-zero values, preserving max of each field
      var existIdx = dedupMap[dKey];
      for (var mi = 2; mi < rows[di].length; mi++) {
        var existVal = Number(rows[existIdx][mi]) || 0;
        var newVal = Number(rows[di][mi]) || 0;
        if (newVal > existVal) rows[existIdx][mi] = rows[di][mi];
      }
      rows.splice(di, 1);
      di--; // re-check this index
    }
  }
  Logger.log('consolidateCampaignSlim_ after dedup: ' + rows.length + ' rows');

  // Write to destination tab
  if (!destTab) destTab = destSS.insertSheet(campaign.label);
  destTab.clear();
  destTab.getRange(1, 1, 1, campaignHeaders.length).setValues([campaignHeaders]);

  if (rows.length > 0) {
    rows.sort(function(a, b) {
      var dateA = (a[0] instanceof Date) ? a[0].getTime() : new Date(a[0]).getTime() || 0;
      var dateB = (b[0] instanceof Date) ? b[0].getTime() : new Date(b[0]).getTime() || 0;
      if (dateA !== dateB) return dateB - dateA;
      return String(a[1]).localeCompare(String(b[1]));
    });
    destTab.getRange(2, 1, rows.length, campaignHeaders.length).setValues(rows);
  }

  return rows.length;
}

// ══════════════════════════════════════════════════
// CAMPAIGN TAB MAP — owner→tab name overrides
// Stored in _Campaign_Tab_Map tab in NATIONAL sheet.
// Columns: campaign | ownerName | tabName | updatedBy | updatedAt
// ══════════════════════════════════════════════════

var CAMPAIGN_TAB_MAP_HEADERS = ['campaign', 'ownerName', 'tabName', 'updatedBy', 'updatedAt'];

/**
 * action=odGetCampaignTabMap
 * Returns all saved campaign→tab mappings plus the available tabs for each campaign spreadsheet.
 */
function odGetCampaignTabMap() {
  var ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var tab = ss.getSheetByName('_Campaign_Tab_Map');

  var mappings = [];
  if (tab) {
    var data = tab.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[1]) continue;
      mappings.push({
        campaign:  String(row[0] || '').trim(),
        ownerName: String(row[1] || '').trim(),
        tabName:   String(row[2] || '').trim(),
        updatedBy: String(row[3] || '').trim(),
        updatedAt: String(row[4] || '')
      });
    }
  }

  // Also gather available tabs per campaign
  var campaignTabs = {};
  var keys = Object.keys(OD_CAMPAIGNS);
  for (var k = 0; k < keys.length; k++) {
    var key = keys[k];
    var cfg = OD_CAMPAIGNS[key];
    if (!cfg.sheetId) continue;
    try {
      var srcSS = SpreadsheetApp.openById(cfg.sheetId);
      var sheets = srcSS.getSheets();
      var tabNames = [];
      for (var s = 0; s < sheets.length; s++) {
        var name = sheets[s].getName().trim();
        if (name && name.charAt(0) !== '_') tabNames.push(name);
      }
      campaignTabs[key] = tabNames;
    } catch (err) {
      Logger.log('odGetCampaignTabMap tab list error for ' + key + ': ' + err.message);
      campaignTabs[key] = [];
    }
  }

  return { success: true, mappings: mappings, campaignTabs: campaignTabs };
}

// ═══════════════════════════════════════════════════════
// FLAGGED REPS — Coach → Planning one-on-one requests
// ═══════════════════════════════════════════════════════

var OD_FLAGGED_HEADERS_ = ['repName', 'ownerName', 'campaign', 'flaggedBy', 'flaggedAt', 'resolved'];

/**
 * action: odGetFlaggedReps (doGet)
 * Returns unresolved flagged reps.
 */
function odGetFlaggedReps() {
  var sheet = odGetOrCreateTab('_OD_FlaggedReps', OD_FLAGGED_HEADERS_);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, reps: [] };

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var colMap = {};
  for (var c = 0; c < headers.length; c++) colMap[headers[c]] = c;

  var reps = [];
  for (var i = 1; i < data.length; i++) {
    var resolved = String(data[i][colMap['resolved']] || '').trim();
    if (resolved) continue; // skip resolved

    reps.push({
      repName: String(data[i][colMap['repName']] || '').trim(),
      ownerName: String(data[i][colMap['ownerName']] || '').trim(),
      campaign: String(data[i][colMap['campaign']] || '').trim(),
      flaggedBy: String(data[i][colMap['flaggedBy']] || '').trim(),
      flaggedAt: String(data[i][colMap['flaggedAt']] || '').trim()
    });
  }
  return { success: true, reps: reps };
}

/**
 * action: odFlagRep (doPost)
 * body: { repName, ownerName, campaign, flaggedBy }
 */
function odFlagRep(body) {
  var repName = String(body.repName || '').trim();
  var ownerName = String(body.ownerName || '').trim();
  var campaign = String(body.campaign || '').trim();
  var flaggedBy = String(body.flaggedBy || '').trim();
  if (!repName || !ownerName || !campaign) {
    return { success: false, message: 'repName, ownerName, and campaign are required' };
  }

  var sheet = odGetOrCreateTab('_OD_FlaggedReps', OD_FLAGGED_HEADERS_);

  // Check if already flagged (unresolved)
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var rn = String(data[i][0] || '').trim().toLowerCase();
    var on = String(data[i][1] || '').trim().toLowerCase();
    var camp = String(data[i][2] || '').trim().toLowerCase();
    var resolved = String(data[i][5] || '').trim();
    if (rn === repName.toLowerCase() && on === ownerName.toLowerCase() && camp === campaign.toLowerCase() && !resolved) {
      return { success: true, message: 'Already flagged' };
    }
  }

  sheet.appendRow([repName, ownerName, campaign, flaggedBy, new Date().toISOString(), '']);
  return { success: true };
}

/**
 * action: odUnflagRep (doPost)
 * body: { repName, ownerName, campaign }
 * Sets resolved = now on the matching unresolved row.
 */
function odUnflagRep(body) {
  var repName = String(body.repName || '').trim().toLowerCase();
  var ownerName = String(body.ownerName || '').trim().toLowerCase();
  var campaign = String(body.campaign || '').trim().toLowerCase();
  if (!repName || !ownerName || !campaign) {
    return { success: false, message: 'repName, ownerName, and campaign are required' };
  }

  var sheet = odGetOrCreateTab('_OD_FlaggedReps', OD_FLAGGED_HEADERS_);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var rn = String(data[i][0] || '').trim().toLowerCase();
    var on = String(data[i][1] || '').trim().toLowerCase();
    var camp = String(data[i][2] || '').trim().toLowerCase();
    var resolved = String(data[i][5] || '').trim();
    if (rn === repName && on === ownerName && camp === campaign && !resolved) {
      sheet.getRange(i + 1, 6).setValue(new Date().toISOString()); // col 6 = resolved
      return { success: true };
    }
  }
  return { success: false, message: 'Flagged rep not found' };
}

// ══════════════════════════════════════════════════
// OWNER NOTES LOG
// ══════════════════════════════════════════════════

var OD_NOTES_HEADERS_ = ['noteId', 'campaign', 'ownerName', 'coachName', 'coachEmail', 'text', 'timestamp'];

/** Read all notes from _OD_Notes tab */
function odGetNotes_() {
  var sheet = odGetOrCreateTab('_OD_Notes', OD_NOTES_HEADERS_);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, notes: [] };

  var notes = [];
  for (var i = 1; i < data.length; i++) {
    var noteId = String(data[i][0] || '').trim();
    if (!noteId) continue;
    notes.push({
      noteId: noteId,
      campaign: String(data[i][1] || '').trim(),
      ownerName: String(data[i][2] || '').trim(),
      coachName: String(data[i][3] || '').trim(),
      coachEmail: String(data[i][4] || '').trim(),
      text: String(data[i][5] || '').trim(),
      timestamp: String(data[i][6] || '').trim()
    });
  }
  return { success: true, notes: notes };
}

/** Add a note to _OD_Notes tab */
function addOwnerNote_(campaign, ownerName, coachName, coachEmail, text) {
  if (!ownerName || !text) return { error: 'ownerName and text are required' };
  var sheet = odGetOrCreateTab('_OD_Notes', OD_NOTES_HEADERS_);
  var noteId = 'n_' + Date.now() + '_' + Math.random().toString(36).substr(2, 4);
  var timestamp = new Date().toISOString();
  sheet.appendRow([noteId, campaign || '', ownerName, coachName || '', coachEmail || '', text, timestamp]);
  return { success: true, noteId: noteId, timestamp: timestamp };
}

/** Delete a note from _OD_Notes tab by noteId */
function deleteOwnerNote_(noteId) {
  if (!noteId) return { error: 'noteId is required' };
  var sheet = odGetOrCreateTab('_OD_Notes', OD_NOTES_HEADERS_);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === noteId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'Note not found' };
}

/**
 * action=odSaveCampaignTabMap (POST)
 * Save a single campaign→owner→tab mapping.
 * Body: { campaign, ownerName, tabName, updatedBy }
 */
function odSaveCampaignTabMap(body) {
  var campaign  = String(body.campaign  || '').trim();
  var ownerName = String(body.ownerName || '').trim();
  var tabName   = String(body.tabName   || '').trim();
  var updatedBy = String(body.updatedBy || '').trim();
  if (!campaign || !ownerName) return { success: false, message: 'campaign and ownerName required' };

  var ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var tab = ss.getSheetByName('_Campaign_Tab_Map');
  if (!tab) {
    tab = ss.insertSheet('_Campaign_Tab_Map');
    tab.getRange(1, 1, 1, CAMPAIGN_TAB_MAP_HEADERS.length).setValues([CAMPAIGN_TAB_MAP_HEADERS]);
  }

  var data = tab.getDataRange().getValues();
  var now = new Date().toISOString();

  // Find existing row
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === campaign.toLowerCase() &&
        String(data[i][1]).trim().toLowerCase() === ownerName.toLowerCase()) {
      tab.getRange(i + 1, 3, 1, 3).setValues([[tabName, updatedBy, now]]);
      return { success: true, updated: true };
    }
  }

  // Append new row
  tab.appendRow([campaign, ownerName, tabName, updatedBy, now]);
  return { success: true, created: true };
}

/**
 * action=odBatchSaveCampaignTabMap (POST)
 * Save multiple campaign→tab mappings at once.
 * Body: { mappings: [{ campaign, ownerName, tabName, updatedBy }] }
 */
function odBatchSaveCampaignTabMap(body) {
  var items = body.mappings || [];
  if (!items.length) return { success: true, saved: 0 };

  var ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var tab = ss.getSheetByName('_Campaign_Tab_Map');
  if (!tab) {
    tab = ss.insertSheet('_Campaign_Tab_Map');
    tab.getRange(1, 1, 1, CAMPAIGN_TAB_MAP_HEADERS.length).setValues([CAMPAIGN_TAB_MAP_HEADERS]);
  }

  var data = tab.getDataRange().getValues();
  var now = new Date().toISOString();

  // Build lookup for existing rows: "campaign|ownerName" → row index
  var lookup = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim().toLowerCase() + '|' + String(data[i][1]).trim().toLowerCase();
    lookup[key] = i + 1; // 1-indexed sheet row
  }

  var toAppend = [];
  var saved = 0;

  for (var j = 0; j < items.length; j++) {
    var item = items[j];
    var campaign  = String(item.campaign  || '').trim();
    var ownerName = String(item.ownerName || '').trim();
    var tabName   = String(item.tabName   || '').trim();
    var updatedBy = String(item.updatedBy || '').trim();
    if (!campaign || !ownerName) continue;

    var key = campaign.toLowerCase() + '|' + ownerName.toLowerCase();
    if (lookup[key]) {
      tab.getRange(lookup[key], 3, 1, 3).setValues([[tabName, updatedBy, now]]);
    } else {
      toAppend.push([campaign, ownerName, tabName, updatedBy, now]);
    }
    saved++;
  }

  if (toAppend.length > 0) {
    var lastRow = tab.getLastRow();
    tab.getRange(lastRow + 1, 1, toAppend.length, CAMPAIGN_TAB_MAP_HEADERS.length).setValues(toAppend);
  }

  return { success: true, saved: saved };
}


// ══════════════════════════════════════════════════
// TEMPORARY TEST — safe to delete after debugging
// Run directly from editor (▶ button) — does NOT affect deployed app
// ══════════════════════════════════════════════════
function TEST_lumen_extraction() {
  var ss = SpreadsheetApp.openById('1P4DYlcV1hgNkaAapk3tWD7ytcRXw4K1n7R6EMKPCoSA');
  var tab = ss.getSheetByName('Campaign');
  Logger.log('Tab found: ' + !!tab);
  if (!tab) {
    Logger.log('Available tabs: ' + ss.getSheets().map(function(t) { return t.getName(); }).join(', '));
    return;
  }
  var data = tab.getDataRange().getValues();
  Logger.log('Data rows: ' + data.length);
  for (var i = 0; i < Math.min(data.length, 20); i++) {
    var val = String(data[i][0] || '');
    Logger.log('Row ' + i + ': "' + val + '" (len=' + val.length + ', trimmed="' + val.trim() + '")');
  }
}

// Run the FULL refresh pipeline for Lumen and log the result
function TEST_lumen_refresh() {
  Logger.log('=== Testing refreshSingleCampaign("lumen") ===');
  Logger.log('OD_CAMPAIGNS.lumen = ' + JSON.stringify(OD_CAMPAIGNS['lumen']));
  var result = refreshSingleCampaign('lumen');
  Logger.log('Result: ' + JSON.stringify(result));
}

// Debug: check which tab Christian Esposito matches to
function TEST_christian_match() {
  var ss = SpreadsheetApp.openById('1P4DYlcV1hgNkaAapk3tWD7ytcRXw4K1n7R6EMKPCoSA');
  var allSheets = ss.getSheets();
  var allTabNames = allSheets.map(function(s) { return s.getName(); });
  var tabByName = {};
  allSheets.forEach(function(s) { tabByName[s.getName().toLowerCase()] = s; });
  Logger.log('All tabs: ' + allTabNames.join(', '));

  var ownerName = 'Christian Esposito';
  // Exact match
  var tab = tabByName[ownerName.toLowerCase()] || null;
  Logger.log('Exact match: ' + (tab ? tab.getName() : 'NONE'));

  // Fuzzy match
  if (!tab) {
    tab = _fuzzyFindTab_(ownerName, allTabNames, tabByName);
    Logger.log('Fuzzy match: ' + (tab ? tab.getName() : 'NONE'));
  }

  if (tab) {
    Logger.log('=== Reading from tab: ' + tab.getName() + ' ===');
    var data = tab.getDataRange().getValues();
    var sections = findSections(data);
    Logger.log('Sections: ' + JSON.stringify(sections));
    var healthRows = extractHealthRows_(data, sections.section1Start, sections.section1End, tab.getDataRange().getDisplayValues());
    Logger.log('Health rows: ' + healthRows.length);
    for (var i = Math.max(0, healthRows.length - 3); i < healthRows.length; i++) {
      var h = healthRows[i];
      Logger.log('Health[' + i + '] active=' + h.active + ' leaders=' + h.leaders + ' dist=' + h.dist + ' training=' + h.training);
    }
  }
}

// Debug: show column mapping + sample data from Christian E's tab
function TEST_angel_columns() {
  var ss = SpreadsheetApp.openById('1P4DYlcV1hgNkaAapk3tWD7ytcRXw4K1n7R6EMKPCoSA');
  var tab = ss.getSheetByName('Christian E');
  if (!tab) { Logger.log('No Christian E tab'); return; }
  var data = tab.getDataRange().getValues();
  var sections = findSections(data);
  Logger.log('Sections: ' + JSON.stringify(sections));

  // Show headers
  var headerIdx = sections.section1Start;
  for (var hi = sections.section1Start; hi <= Math.min(sections.section1Start + 3, sections.section1End); hi++) {
    var rowText = data[hi].map(function(c) { return String(c).trim(); }).join('');
    if (rowText.length > 0) { headerIdx = hi; break; }
  }
  var headers = data[headerIdx].map(function(h) { return String(h).toLowerCase().trim(); });
  Logger.log('Header row ' + headerIdx + ': ' + JSON.stringify(headers.slice(0, 15)));
  Logger.log('active col: ' + findCol(headers, ['active', 'total agents', 'agents']));
  Logger.log('leaders col: ' + findCol(headers, ['leaders']));
  Logger.log('dist col: ' + findCol(headers, ['dist', 'distributors', 'distibutors']));
  Logger.log('training col: ' + findCol(headers, ['training']));

  // Health extraction
  var healthRows = extractHealthRows_(data, sections.section1Start, sections.section1End, tab.getDataRange().getDisplayValues());
  Logger.log('Health rows: ' + healthRows.length);
  // Show last 3 rows (most recent)
  for (var i = Math.max(0, healthRows.length - 3); i < healthRows.length; i++) {
    var h = healthRows[i];
    Logger.log('Health[' + i + '] date=' + h.date + ' active=' + h.active + ' leaders=' + h.leaders + ' dist=' + h.dist + ' training=' + h.training + ' prod=' + h.productionRaw + ' goals=' + h.goalsRaw);
  }

  // Recruiting dates
  var recruitingRows = [];
  if (sections.section2Start >= 0) {
    recruitingRows = extractRecruitingRows_(data, sections.section2Start, sections.section2End);
  }
  if (recruitingRows.length === 0) {
    recruitingRows = extractHorizontalRecruitingRows_(data, sections.section1Start, sections.section1End);
  }
  Logger.log('Recruiting rows: ' + recruitingRows.length);
  for (var i = 0; i < Math.min(recruitingRows.length, 5); i++) {
    var r = recruitingRows[i];
    var d = r.date;
    Logger.log('Recruit[' + i + '] date=' + d + ' type=' + typeof d + ' isDate=' + (d instanceof Date) + ' key=' + _normalizeDateKey_(d));
  }
}

function TEST_rogers_pipeline() {
  var campaignKey = 'rogers';
  var ss = SpreadsheetApp.openById(OD_CAMPAIGNS[campaignKey].sheetId);
  var sheets = ss.getSheets();
  var tab = null;
  var skip = ['campaign', 'blank copy', '_ownertabmapping', 'template', 'rogers campaign stats', 'fios campaign stats', 'campaign totals'];
  for (var i = 0; i < sheets.length; i++) {
    var tName = sheets[i].getName().toLowerCase().trim();
    var isSkip = false;
    for (var s = 0; s < skip.length; s++) {
      if (tName === skip[s] || tName.indexOf(skip[s]) >= 0) { isSkip = true; break; }
    }
    if (!isSkip) { tab = sheets[i]; break; }
  }
  if (!tab) { Logger.log('No owner tab found'); return; }
  Logger.log('Using tab: ' + tab.getName());

  var data = tab.getDataRange().getValues();
  var displayData = tab.getDataRange().getDisplayValues();
  var sections = findSections(data);
  Logger.log('Sections: ' + JSON.stringify(sections));

  // Step 1: Extract health rows WITH campaignKey
  var healthRows = extractHealthRows_(data, sections.section1Start, sections.section1End, displayData, campaignKey);
  Logger.log('Health rows: ' + healthRows.length);

  // Log last 3 health rows — focus on productionRaw and goalsRaw
  for (var i = Math.max(0, healthRows.length - 3); i < healthRows.length; i++) {
    var h = healthRows[i];
    Logger.log('Health[' + i + '] date=' + h.date + ' active=' + h.active + ' leaders=' + h.leaders +
      ' productionRaw=' + h.productionRaw + ' goalsRaw=' + h.goalsRaw +
      ' (typeof prod=' + typeof h.productionRaw + ', typeof goals=' + typeof h.goalsRaw + ')');
  }

  // Step 2: Extract recruiting rows
  var recruitingRows = extractHorizontalRecruitingRows_(data, sections.section1Start, sections.section1End);
  if (recruitingRows.length === 0 && sections.section2Start >= 0) {
    recruitingRows = extractRecruitingRows_(data, sections.section2Start, sections.section2End);
  }
  Logger.log('Recruiting rows: ' + recruitingRows.length);

  // Step 3: Merge
  var merged = mergeHealthRecruiting_(tab.getName(), healthRows, recruitingRows, campaignKey);
  Logger.log('Merged rows: ' + merged.length);

  // Log last 3 merged rows — show all columns
  var headers = getConsolidatedHeaders_(campaignKey);
  Logger.log('Headers: ' + JSON.stringify(headers));
  for (var i = Math.max(0, merged.length - 3); i < merged.length; i++) {
    var row = merged[i];
    var parts = [];
    for (var c = 0; c < Math.min(row.length, headers.length); c++) {
      parts.push(headers[c] + '=' + row[c]);
    }
    Logger.log('Row[' + i + '] ' + parts.join(', '));
  }
}

// Debug: trace Lumen date parsing for Angel Padilla
function TEST_lumen_dates() {
  var ss = SpreadsheetApp.openById('1P4DYlcV1hgNkaAapk3tWD7ytcRXw4K1n7R6EMKPCoSA');
  var tab = ss.getSheetByName('Angel P');
  if (!tab) {
    // Try variations
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var n = sheets[i].getName().toLowerCase();
      if (n.indexOf('angel') >= 0) { tab = sheets[i]; break; }
    }
  }
  if (!tab) { Logger.log('No Angel tab found'); return; }
  Logger.log('Using tab: ' + tab.getName());

  var data = tab.getDataRange().getValues();
  var sections = findSections(data);
  Logger.log('Sections: ' + JSON.stringify(sections));

  // Find the dates column
  var headerIdx = sections.section1Start;
  for (var hi = sections.section1Start; hi <= Math.min(sections.section1Start + 3, sections.section1End); hi++) {
    var rowText = data[hi].map(function(c) { return String(c).trim(); }).join('');
    if (rowText.length > 0) { headerIdx = hi; break; }
  }
  var headers = data[headerIdx].map(function(h) { return String(h).toLowerCase().trim(); });
  var dateCol = findCol(headers, ['dates', 'date']);
  Logger.log('Date col: ' + dateCol);

  // Log EVERY date value, its type, and what _parseTabDate returns
  Logger.log('=== RAW DATE VALUES ===');
  for (var i = headerIdx + 1; i <= sections.section1End; i++) {
    var raw = data[i][dateCol];
    var isDate = raw instanceof Date;
    var parsed = isDate ? raw : _parseTabDate(String(raw || ''));
    Logger.log('Row ' + i + ': raw="' + raw + '" type=' + typeof raw + ' isDate=' + isDate +
      ' → parsed=' + (parsed ? parsed.toLocaleDateString() : 'null'));
  }

  // Now run full extraction with campaignKey='lumen' and log corrected dates
  var displayData = tab.getDataRange().getDisplayValues();
  var healthRows = extractHealthRows_(data, sections.section1Start, sections.section1End, displayData, 'lumen');
  Logger.log('=== CORRECTED HEALTH ROWS (' + healthRows.length + ') ===');
  for (var i = 0; i < healthRows.length; i++) {
    var h = healthRows[i];
    Logger.log('Health[' + i + '] date=' + h.date.toLocaleDateString() + ' active=' + h.active + ' prod=' + h.productionRaw);
  }
}

// Debug: list all tab names in Frontier source sheet, then inspect Alante's tab
function TEST_frontier_alante() {
  var ss = SpreadsheetApp.openById('1WWpLQTCyvPmJbx3jjowFszwOF_JUnjS6tzu-eAASwk0');
  var allTabs = ss.getSheets().map(function(s) { return s.getName(); });
  Logger.log('All tabs (' + allTabs.length + '): ' + allTabs.join(', '));

  // Find Alante's tab (fuzzy)
  var tab = null;
  for (var i = 0; i < allTabs.length; i++) {
    if (allTabs[i].toLowerCase().indexOf('alante') >= 0) {
      tab = ss.getSheetByName(allTabs[i]);
      Logger.log('Found Alante tab: "' + allTabs[i] + '"');
      break;
    }
  }
  if (!tab) { Logger.log('No tab containing "alante" found'); return; }

  var data = tab.getDataRange().getValues();
  Logger.log('Total rows: ' + data.length);
  var sections = findSections(data);
  Logger.log('Sections: ' + JSON.stringify(sections));
  // Log ALL health section rows with first 8 columns
  for (var r = 0; r <= Math.min(sections.section1End, data.length - 1); r++) {
    var cols = [];
    for (var c = 0; c < Math.min(8, data[r].length); c++) {
      var v = data[r][c];
      cols.push('[' + c + ']' + (v instanceof Date ? 'DATE:' + v.toISOString().slice(0,10) : v));
    }
    Logger.log('Row ' + r + ': ' + cols.join(' | '));
  }
}

// ═══════════════════════════════════════════════════════
// OWNER DEVELOPMENT (OD) — Planning Schedule
// ═══════════════════════════════════════════════════════

var OD_PLANNING_HEADERS_ = ['day', 'sortOrder', 'campaignKey', 'ownerOrder', 'updatedBy', 'updatedAt'];

/**
 * action: odGetPlanning (doGet)
 * Returns the weekly planning schedule from _OD_Planning tab.
 */
function odGetPlanning() {
  var sheet = odGetOrCreateTab('_OD_Planning', OD_PLANNING_HEADERS_);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, planning: [] };

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var colMap = {};
  for (var c = 0; c < headers.length; c++) colMap[headers[c]] = c;

  var planning = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var campaignKey = String(row[colMap['campaignKey']] || '').trim();
    if (!campaignKey) continue;

    var ownerOrderRaw = String(row[colMap['ownerOrder']] || '');
    var ownerOrder = [];
    try { ownerOrder = JSON.parse(ownerOrderRaw); } catch (e) { ownerOrder = []; }
    if (!Array.isArray(ownerOrder)) ownerOrder = [];

    planning.push({
      day: parseInt(row[colMap['day']]) || 0,
      sortOrder: parseInt(row[colMap['sortOrder']]) || 0,
      campaignKey: campaignKey,
      ownerOrder: ownerOrder
    });
  }
  return { success: true, planning: planning };
}

/**
 * action: odSavePlanning (doPost)
 * body: { planning: [{ day, sortOrder, campaignKey, ownerOrder }], email }
 * Full-replace strategy: clears all data rows, writes the new schedule.
 */
function odSavePlanning(body) {
  var items = body.planning;
  if (!items || !Array.isArray(items)) {
    return { success: false, message: 'planning array is required' };
  }

  var email = String(body.email || '').trim();
  var now = new Date().toISOString();
  var sheet = odGetOrCreateTab('_OD_Planning', OD_PLANNING_HEADERS_);

  // Clear existing data rows (keep header)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, OD_PLANNING_HEADERS_.length).clearContent();
  }

  // Write new rows
  if (items.length > 0) {
    var newRows = [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var ownerOrder = '';
      try { ownerOrder = JSON.stringify(item.ownerOrder || []); } catch (e) { ownerOrder = '[]'; }
      newRows.push([
        parseInt(item.day) || 0,
        parseInt(item.sortOrder) || 0,
        String(item.campaignKey || '').trim(),
        ownerOrder,
        email,
        now
      ]);
    }
    sheet.getRange(2, 1, newRows.length, OD_PLANNING_HEADERS_.length).setValues(newRows);
  }

  return { success: true, saved: items.length };
}


// ═══════════════════════════════════════════════════════
// SEED DATA — Run once from editor to populate initial roles
// After running, delete this function or comment it out.
// ═══════════════════════════════════════════════════════

function seedODRolesAndOwnership() {
  var now = new Date().toISOString();

  // ── 1. Seed _OD_Campaign_Ownership ──
  var ownershipSheet = odGetOrCreateTab('_OD_Campaign_Ownership', OD_OWNERSHIP_HEADERS_);
  ownershipSheet.clearContents();
  ownershipSheet.getRange(1, 1, 1, OD_OWNERSHIP_HEADERS_.length).setValues([OD_OWNERSHIP_HEADERS_]);

  var ownershipRows = [
    // Maddy's campaigns
    ['att-nds',       'mdanesi3@gmail.com',                    'Maddy Komis',          now],
    ['frontier',      'mdanesi3@gmail.com',                    'Maddy Komis',          now],
    ['verizon-fios',  'mdanesi3@gmail.com',                    'Maddy Komis',          now],
    ['att-res',       'mdanesi3@gmail.com',                    'Maddy Komis',          now],
    ['lumen',         'mdanesi3@gmail.com',                    'Maddy Komis',          now],
    ['rogers',        'mdanesi3@gmail.com',                    'Maddy Komis',          now],
    ['leafguard',     'mdanesi3@gmail.com',                    'Maddy Komis',          now],
    // Caitlyn's campaigns
    ['att-b2b',       'caitlyn@nextlevelrecruitinginc.com',    'Caitlyn Carrafiello',  now],
    ['box',           'caitlyn@nextlevelrecruitinginc.com',    'Caitlyn Carrafiello',  now]
  ];
  ownershipSheet.getRange(2, 1, ownershipRows.length, OD_OWNERSHIP_HEADERS_.length).setValues(ownershipRows);
  Logger.log('Seeded _OD_Campaign_Ownership: ' + ownershipRows.length + ' rows');

  // ── 2. Seed _OD_Users (upsert — preserves existing PIN hashes) ──
  var usersSheet = odGetOrCreateTab('_OD_Users', OD_USERS_HEADERS_);
  var existingData = usersSheet.getDataRange().getValues();
  var existingMap = {}; // email → row index (1-based)
  if (existingData.length > 1) {
    for (var i = 1; i < existingData.length; i++) {
      var em = String(existingData[i][0] || '').trim().toLowerCase();
      if (em) existingMap[em] = i + 1;
    }
  }

  // [email, name, team, role, pinHash, dateAdded, deactivated, managedBy]
  var users = [
    // National Consultants
    ['kweinraub@gmail.com',                     'Ken Weinraub',          'ken',     'national',     '', now, '', ''],
    ['vanberardino@gmail.com',                   'Van Berardino',         'van',     'national',     '', now, '', ''],
    ['raffi127@gmail.com',                       'Rafael Hidalgo',        'rafael',  'national',     '', now, '', ''],
    ['synergyadvmkg@gmail.com',                  'Sam Rabinowitz',        'sam',     'national',     '', now, '', ''],
    // Org Managers
    ['mdanesi3@gmail.com',                       'Maddy Komis',           'ken',     'org_manager',  '', now, '', 'kweinraub@gmail.com'],
    ['caitlyn@nextlevelrecruitinginc.com',       'Caitlyn Carrafiello',   'van',     'org_manager',  '', now, '', 'vanberardino@gmail.com'],
    // Admins
    ['bruhncloie7@gmail.com',                    'Cloie Bruhn',           'ken',     'admin',        '', now, '', 'mdanesi3@gmail.com'],
    // NLR
    ['erica@nextlevelrecruitinginc.com',         'Erica Gordon',          'nlr',     'nlr_manager',  '', now, '', ''],
    // BIS
    ['cam@betterimagesolutions.com',             'Cam Besseda',           'bis',     'bis_manager',  '', now, '', ''],
    // Aptel
    ['austin.rios@thesmartcircle.com',           'Austin Rios',           '',        'aptel',        '', now, '', ''],
    ['destiny@aptel.com',                        'Destiny Tuangco',       '',        'aptel',        '', now, '', ''],
    ['jorton@thesmartcircle.com',                'Josh Orton',            '',        'aptel',        '', now, '', ''],
    // Superadmin
    ['alex.aspirehr@gmail.com',                  'Alex Prindle',          '',        'superadmin',   '', now, '', '']
  ];

  var added = 0, updated = 0;
  for (var u = 0; u < users.length; u++) {
    var row = users[u];
    var email = row[0].toLowerCase();
    var existingRow = existingMap[email];

    if (existingRow) {
      // Update: name, team, role, managedBy — PRESERVE pinHash
      usersSheet.getRange(existingRow, 2).setValue(row[1]); // name
      usersSheet.getRange(existingRow, 3).setValue(row[2]); // team
      usersSheet.getRange(existingRow, 4).setValue(row[3]); // role
      usersSheet.getRange(existingRow, 7).setValue('');      // clear deactivated
      usersSheet.getRange(existingRow, 8).setValue(row[7]);  // managedBy
      updated++;
    } else {
      usersSheet.appendRow(row);
      added++;
    }
  }
  Logger.log('Seeded _OD_Users: ' + added + ' added, ' + updated + ' updated (PINs preserved)');

  // ── 3. Ensure _OD_Access_Grants tab exists ──
  odGetOrCreateTab('_OD_Access_Grants', OD_GRANTS_HEADERS_);
  Logger.log('_OD_Access_Grants tab ready');

  Logger.log('=== SEED COMPLETE ===');
  return { ownership: ownershipRows.length, usersAdded: added, usersUpdated: updated };
}

// ── Diagnostic: inspect Joey Ramirez's tab in the B2B source sheet ──
// Run from the Apps Script editor (▶ button). Safe — read-only, no writes.
function TEST_joey_ramirez() {
  var OWNER = 'Joey Ramirez';
  var campaign = OD_CAMPAIGNS['att-b2b'];
  var ss = SpreadsheetApp.openById(campaign.sheetId);

  // 1. Find the tab
  var tab = ss.getSheetByName(OWNER);
  if (!tab) {
    Logger.log('❌ Tab "' + OWNER + '" NOT FOUND.');
    Logger.log('Available tabs: ' + ss.getSheets().map(function(t) { return t.getName(); }).join(', '));
    return;
  }
  Logger.log('✅ Tab found: "' + tab.getName() + '"');

  // 2. Read raw data
  var data = tab.getDataRange().getValues();
  var displayData = tab.getDataRange().getDisplayValues();
  Logger.log('Total rows in tab: ' + data.length);

  // 3. Print first 5 rows so you can see the structure
  Logger.log('--- First 5 rows (raw) ---');
  for (var i = 0; i < Math.min(5, data.length); i++) {
    Logger.log('Row ' + i + ': ' + JSON.stringify(data[i].slice(0, 10)));
  }

  // 4. Detect sections
  var sections = findSections(data);
  Logger.log('Sections: s1=' + sections.section1Start + '-' + sections.section1End
    + '  s2=' + sections.section2Start + '-' + sections.section2End);

  // 5. Extract health rows
  var healthRows = extractHealthRows_(data, sections.section1Start, sections.section1End, displayData, 'att-b2b');
  Logger.log('Health rows extracted: ' + healthRows.length);
  for (var h = 0; h < healthRows.length; h++) {
    var hr = healthRows[h];
    Logger.log('  Health[' + h + '] date=' + (hr.date ? hr.date.toDateString() : 'null')
      + ' active=' + hr.active + ' leaders=' + hr.leaders
      + ' prod=' + hr.productionRaw + ' goals=' + hr.goalsRaw);
  }

  // 6. Extract recruiting rows
  var recruitingRows = extractHorizontalRecruitingRows_(data, sections.section1Start, sections.section1End, 'att-b2b');
  Logger.log('Recruiting rows extracted: ' + recruitingRows.length);
  for (var r = 0; r < recruitingRows.length; r++) {
    var rr = recruitingRows[r];
    Logger.log('  Recruiting[' + r + '] date=' + (rr.date ? rr.date.toDateString() : 'null')
      + ' metrics=' + JSON.stringify(rr.metrics));
  }

  // 7. Run the full merge and show output rows
  var merged = mergeHealthRecruiting_(OWNER, healthRows, recruitingRows, 'att-b2b');
  Logger.log('Merged rows: ' + merged.length);
  for (var m = 0; m < merged.length; m++) {
    Logger.log('  Row[' + m + '] week=' + (merged[m][0] instanceof Date ? merged[m][0].toDateString() : merged[m][0])
      + ' owner=' + merged[m][1]);
  }
}

// ── Diagnostic: inspect Ayman's Fios tab — check 3 tracked columns + slash-split production ──
// Run from the Apps Script editor (▶ button). Safe — read-only, no writes.
function TEST_ayman_fios() {
  var OWNER = 'Ayman';
  var campaign = OD_CAMPAIGNS['verizon-fios'];
  var ss = SpreadsheetApp.openById(campaign.sheetId);

  // 1. List all tabs and find Ayman's
  var allTabs = ss.getSheets().map(function(t) { return t.getName(); });
  Logger.log('All tabs: ' + allTabs.join(', '));

  // Try to find Ayman's tab (exact or partial match)
  var tab = ss.getSheetByName(OWNER);
  if (!tab) {
    // Try partial match
    for (var ti = 0; ti < allTabs.length; ti++) {
      if (allTabs[ti].toLowerCase().indexOf('ayman') >= 0) {
        tab = ss.getSheetByName(allTabs[ti]);
        Logger.log('Matched tab by partial name: "' + allTabs[ti] + '"');
        break;
      }
    }
  }
  if (!tab) {
    Logger.log('❌ No tab matching "Ayman" found. Available tabs listed above.');
    return;
  }
  Logger.log('✅ Tab: "' + tab.getName() + '"');

  // 2. Read raw + display data
  var data = tab.getDataRange().getValues();
  var displayData = tab.getDataRange().getDisplayValues();
  Logger.log('Total rows: ' + data.length);

  // 3. Print first 8 rows (raw) to see structure
  Logger.log('--- First 8 rows (raw, cols 0-12) ---');
  for (var i = 0; i < Math.min(8, data.length); i++) {
    Logger.log('Row ' + i + ': ' + JSON.stringify(data[i].slice(0, 13)));
  }
  Logger.log('--- First 8 rows (display, cols 0-12) ---');
  for (var i = 0; i < Math.min(8, data.length); i++) {
    Logger.log('Row ' + i + ': ' + JSON.stringify(displayData[i].slice(0, 13)));
  }

  // 4. Detect sections
  var sections = findSections(data);
  Logger.log('Sections: s1=' + sections.section1Start + '-' + sections.section1End
    + '  s2=' + sections.section2Start + '-' + sections.section2End);

  // 5. Show Section 1 headers + per-product column resolution
  if (sections.section1Start >= 0) {
    // Find actual header row (first non-empty row in section)
    var headerIdx = sections.section1Start;
    for (var hi = sections.section1Start; hi <= Math.min(sections.section1Start + 3, sections.section1End); hi++) {
      var rowText = data[hi].map(function(c) { return String(c).trim(); }).join('');
      if (rowText.length > 0) { headerIdx = hi; break; }
    }
    var headers = data[headerIdx].map(function(h) { return String(h).toLowerCase().trim(); });
    Logger.log('Section 1 headers (row ' + headerIdx + '): ' + JSON.stringify(headers));

    // Show what CAMPAIGN_PROD_COLUMNS resolves to for Fios
    var prodColConfig = CAMPAIGN_PROD_COLUMNS['verizon-fios'];
    Logger.log('--- Per-product column resolution (verizon-fios) ---');
    var prodNames = Object.keys(prodColConfig);
    for (var pp = 0; pp < prodNames.length; pp++) {
      var cfg = prodColConfig[prodNames[pp]];
      var colIdx = findCol(headers, cfg.prod);
      Logger.log('  Product "' + prodNames[pp] + '": searching ' + JSON.stringify(cfg.prod) + ' → col ' + colIdx
        + (colIdx >= 0 ? ' (header: "' + headers[colIdx] + '")' : ' NOT FOUND'));
    }

    // Also show ALL headers that contain 'wireless', 'production', 'units'
    Logger.log('--- Headers containing key terms ---');
    for (var ci = 0; ci < headers.length; ci++) {
      var h = headers[ci];
      if (h.indexOf('wireless') >= 0 || h.indexOf('production') >= 0 || h.indexOf('units') >= 0
          || h.indexOf('internet') >= 0 || h.indexOf('sales') >= 0 || h.indexOf('fios') >= 0) {
        Logger.log('  col[' + ci + '] = "' + h + '"');
      }
    }
  }

  // 6. Extract health rows and show raw vs display for production columns
  var healthRows = extractHealthRows_(data, sections.section1Start, sections.section1End, displayData, 'verizon-fios');
  Logger.log('Health rows extracted: ' + healthRows.length);
  for (var h = 0; h < healthRows.length; h++) {
    var hr = healthRows[h];
    Logger.log('  Health[' + h + '] date=' + (hr.date ? hr.date.toDateString() : 'null')
      + ' active=' + hr.active
      + ' prodRaw="' + hr.productionRaw + '" goalsRaw="' + hr.goalsRaw + '"');
  }

  // 7. Also show raw cell values for section 1 data rows (to see if slash-values are stored)
  if (sections.section1Start >= 0) {
    var headerIdx2 = sections.section1Start;
    for (var hi2 = sections.section1Start; hi2 <= Math.min(sections.section1Start + 3, sections.section1End); hi2++) {
      if (data[hi2].map(function(c) { return String(c).trim(); }).join('').length > 0) { headerIdx2 = hi2; break; }
    }
    var headers2 = data[headerIdx2].map(function(h) { return String(h).toLowerCase().trim(); });
    var unitCol = findCol(headers2, ['production lw', 'production']);
    var wlCol   = findCol(headers2, ['wireless']);
    Logger.log('--- Section 1 data rows: col[' + unitCol + ']=Units, col[' + wlCol + ']=Wireless ---');
    for (var di = headerIdx2 + 1; di <= sections.section1End; di++) {
      var dateV = data[di][findCol(headers2, ['dates', 'date'])];
      if (!dateV) continue;
      var unitRaw = unitCol >= 0 ? data[di][unitCol] : 'N/A';
      var unitDisp = unitCol >= 0 ? displayData[di][unitCol] : 'N/A';
      var wlRaw  = wlCol >= 0  ? data[di][wlCol]  : 'N/A';
      var wlDisp = wlCol >= 0  ? displayData[di][wlCol]  : 'N/A';
      Logger.log('  Row ' + di + ' date=' + dateV
        + '  Units raw=' + JSON.stringify(unitRaw) + ' disp=' + JSON.stringify(unitDisp)
        + '  Wireless raw=' + JSON.stringify(wlRaw) + ' disp=' + JSON.stringify(wlDisp));
    }
  }

  // 8. Show merged output
  var recruitingRows = extractHorizontalRecruitingRows_(data, sections.section1Start, sections.section1End, 'verizon-fios');
  Logger.log('Recruiting rows: ' + recruitingRows.length);
  var merged = mergeHealthRecruiting_(OWNER, healthRows, recruitingRows, 'verizon-fios');
  Logger.log('Merged rows: ' + merged.length);
  var mHeaders = getConsolidatedHeaders_('verizon-fios');
  var unitsIdx  = mHeaders.indexOf('Units');
  var wlIdx     = mHeaders.indexOf('Wireless');
  for (var m = 0; m < merged.length; m++) {
    var week = merged[m][0];
    Logger.log('  Row[' + m + '] week=' + (week instanceof Date ? week.toDateString() : week)
      + '  Units=' + merged[m][unitsIdx] + '  Wireless=' + merged[m][wlIdx]);
  }
}

// ── Diagnostic: compare column B (Total Volume) vs product breakdown sum for all B2B owners ──
// Read-only. Run from Apps Script editor to check if totalVol = internet+voip+wireless+air.
function TEST_b2b_volume_alignment() {
  var ss = SpreadsheetApp.openById(SHEETS.NATIONAL);
  var tab = ss.getSheetByName('B2B Sales Metrics');
  if (!tab) { Logger.log('❌ B2B Sales Metrics tab not found'); return; }

  var data = tab.getDataRange().getValues();
  if (data.length < 2) { Logger.log('❌ No data'); return; }

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var colName    = headers.indexOf('Name');
  var colTV      = headers.indexOf('Total Volume');
  var colInternet = headers.indexOf('Internet');
  var colVOIP    = headers.indexOf('VOIP');
  var colWireless = headers.indexOf('Wireless');
  var colAIR     = headers.indexOf('AIR/AWB');

  Logger.log('Columns — Name:' + colName + ' TV:' + colTV + ' Internet:' + colInternet
    + ' VOIP:' + colVOIP + ' Wireless:' + colWireless + ' AIR:' + colAIR);

  var mismatches = 0;
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][colName] || '').trim();
    if (!name || name.charAt(0) === ' ') continue; // skip rep rows (indented) and blanks
    var tv       = Number(data[i][colTV])      || 0;
    var internet = Number(data[i][colInternet]) || 0;
    var voip     = Number(data[i][colVOIP])     || 0;
    var wireless = Number(data[i][colWireless]) || 0;
    var air      = Number(data[i][colAIR])      || 0;
    var breakdown = internet + voip + wireless + air;
    var diff = tv - breakdown;
    var status = diff === 0 ? '✅' : '❌ diff=' + diff;
    Logger.log(status + '  ' + name + ':  TV=' + tv
      + '  internet=' + internet + '  voip=' + voip
      + '  wireless=' + wireless + '  air=' + air
      + '  breakdown_sum=' + breakdown);
    if (diff !== 0) mismatches++;
  }
  Logger.log('--- ' + mismatches + ' mismatches out of ' + (data.length - 1) + ' owner rows ---');
}
