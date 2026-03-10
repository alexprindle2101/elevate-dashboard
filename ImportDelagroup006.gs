// ═══════════════════════════════════════════════════════════
// One-Time Import: Delagroup Order Log → _Sales_off_nds_006
// ═══════════════════════════════════════════════════════════
// Run importDelagroupSales() from Apps Script editor.
// Source: Delagroup Order Log sheet "Training and Tracking" tab
// Dest:   NDS campaign sheet "_Sales_off_nds_006" tab
//
// After running, check Logger output (View > Logs) for:
//   - Total rows imported
//   - Rows skipped (no date or no rep name)
//   - Reps with generated placeholder emails (inactive)

var SOURCE_SHEET_ID = '1CEoyrg4-t1nxq7Je4PXoWUeaVId3nAIuPR5A5Vd8MOw';
var SOURCE_TAB_SALES = 'Training and Tracking';
var SOURCE_TAB_REPS = 'Active Sales Reps';

// Destination: NDS campaign sheet (same one used by NdsCode.gs)
var DEST_SHEET_ID = '1RQaw9XHdHXnr9laW0UtPDQxfCpHQVzeXfo6Z3Pz5VPA';
var DEST_TAB = '_Sales_off_nds_006';
var DEST_ROSTER_TAB = '_Roster_off_nds_006';
var DEST_TEAMS_TAB = '_Teams_off_nds_006';

var ROSTER_HEADERS = ['email', 'name', 'team', 'rank', 'deactivated', 'dateAdded', 'pinHash', 'phone', 'tableauName'];
var TEAMS_HEADERS = ['teamId', 'name', 'parentId', 'leaderId', 'emoji', 'createdDate'];

var SALES_HEADERS = [
  'Timestamp', 'Email', 'Rep Name', 'Date of Sale', 'Campaign',
  'DSI', 'Account Type', 'Client Name', 'Trainee', 'Trainee Name',
  'Air', 'New Phones', 'BYODs', 'Cell',
  'Ooma Package', 'Account Notes', 'Activation Support', 'Team Emoji',
  'Yeses', 'Units', 'Status', 'Notes', 'Paid Out', 'Tickets',
  'Order Channel', 'Codes Used By'
];


// ── Name → Email mapping from Active Sales Reps ──────────
// Built dynamically from the source sheet's "Active Sales Reps" tab.
// Alias overrides handle name variants found in the sales data.

var NAME_ALIASES = {
  'stephan k chhun': 'stephan chhun',
  'miranda': 'miranda smith',
  'austin roy': 'austin roy',          // placeholder — inactive
  'austin  roy': 'austin roy',
  'sadiya': 'sadiya ali',
  'sadiya and': 'sadiya ali',
  'amin': 'amin mirzaei',
  'melani payan': 'melani payan',
  'hussein': 'hussein',                 // placeholder — inactive
  'harold alex elbelle': 'harold alex elbelle',
  'memphis kealohi': 'memphis kealohi'
};


function importDelagroupSales() {
  var sourceSS = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  var destSS = SpreadsheetApp.openById(DEST_SHEET_ID);

  // ── 1. Build name → email map from Active Sales Reps ──
  var nameEmailMap = buildNameEmailMap_(sourceSS);
  Logger.log('Name→Email map built: ' + Object.keys(nameEmailMap).length + ' entries');

  // ── 2. Read source sales data ──
  var salesSheet = sourceSS.getSheetByName(SOURCE_TAB_SALES);
  if (!salesSheet) { Logger.log('ERROR: Tab "' + SOURCE_TAB_SALES + '" not found'); return; }
  var salesData = salesSheet.getDataRange().getValues();
  Logger.log('Source rows (incl header): ' + salesData.length);

  // ── 3. Get or create destination tab ──
  var destSheet = destSS.getSheetByName(DEST_TAB);
  if (!destSheet) {
    destSheet = destSS.insertSheet(DEST_TAB);
    destSheet.appendRow(SALES_HEADERS);
    Logger.log('Created tab: ' + DEST_TAB);
  }

  // ── 4. Transform and collect rows ──
  var rows = [];
  var skipped = 0;
  var unmatchedLog = {};
  var now = new Date();

  for (var i = 1; i < salesData.length; i++) {  // skip header row
    var srcRow = salesData[i];
    var dateRaw = srcRow[0];
    var repName = String(srcRow[1] || '').trim();
    var trainee = String(srcRow[2] || '').trim();
    var orderSummary = String(srcRow[3] || '').trim();
    var notes = String(srcRow[4] || '').trim();

    // Skip empty rows
    if (!dateRaw && !repName) { skipped++; continue; }
    if (!dateRaw) { skipped++; continue; }

    // Parse date
    var dateOfSale;
    if (dateRaw instanceof Date) {
      dateOfSale = dateRaw;
    } else {
      dateOfSale = new Date(String(dateRaw).trim());
    }
    if (isNaN(dateOfSale.getTime())) { skipped++; continue; }

    // Resolve email
    var email = resolveEmail_(repName, nameEmailMap);
    if (!email) {
      // Generate placeholder for inactive reps
      email = generatePlaceholderEmail_(repName);
      if (!unmatchedLog[repName]) unmatchedLog[repName] = 0;
      unmatchedLog[repName]++;
    }

    // Parse Order Summary
    var parsed = parseOrderSummary_(orderSummary);

    // Calculate products
    var air = parsed.air;
    var newPhones = parsed.newPhones;
    var byods = parsed.byods;
    var cell = newPhones + byods;

    var yeses = 0;
    if (air > 0) yeses++;
    if (cell > 0) yeses++;

    var units = air + cell;

    // Build 26-column row
    var newRow = [
      now,                          // 0  Timestamp
      email,                        // 1  Email
      repName,                      // 2  Rep Name
      dateOfSale,                   // 3  Date of Sale
      'attb2b',                     // 4  Campaign
      parsed.dsi,                   // 5  DSI
      '',                           // 6  Account Type
      parsed.clientName,            // 7  Client Name
      trainee ? 'Yes' : 'No',      // 8  Trainee
      trainee || '',                // 9  Trainee Name
      air,                          // 10 Air
      newPhones,                    // 11 New Phones
      byods,                        // 12 BYODs
      cell,                         // 13 Cell
      '',                           // 14 Ooma Package
      '',                           // 15 Account Notes
      'No',                         // 16 Activation Support
      '',                           // 17 Team Emoji (populated later by roster)
      yeses,                        // 18 Yeses
      units,                        // 19 Units
      'Pending',                    // 20 Status
      notes,                        // 21 Notes
      '',                           // 22 Paid Out
      '[]',                         // 23 Tickets
      'Sara',                       // 24 Order Channel
      ''                            // 25 Codes Used By
    ];

    rows.push(newRow);
  }

  // ── 5. Batch write to destination ──
  if (rows.length > 0) {
    var startRow = destSheet.getLastRow() + 1;
    destSheet.getRange(startRow, 1, rows.length, 26).setValues(rows);
    Logger.log('SUCCESS: Imported ' + rows.length + ' sales rows to ' + DEST_TAB);
  } else {
    Logger.log('WARNING: No rows to import');
  }

  Logger.log('Skipped rows: ' + skipped);

  // Log unmatched reps (inactive — got placeholder emails)
  var unmatchedKeys = Object.keys(unmatchedLog);
  if (unmatchedKeys.length > 0) {
    Logger.log('── Inactive reps (placeholder emails): ──');
    unmatchedKeys.sort(function(a, b) { return unmatchedLog[b] - unmatchedLog[a]; });
    for (var u = 0; u < unmatchedKeys.length; u++) {
      var name = unmatchedKeys[u];
      Logger.log('  ' + name + ': ' + unmatchedLog[name] + ' sales → ' + generatePlaceholderEmail_(name));
    }
  }
}


// ═══════════════════════════════════════════════════════════
// Helper Functions
// ═══════════════════════════════════════════════════════════

/**
 * Build name → email map from the "Active Sales Reps" tab.
 * Reads both Leader and Client Rep sections.
 */
function buildNameEmailMap_(sourceSS) {
  var map = {};
  var sheet = sourceSS.getSheetByName(SOURCE_TAB_REPS);
  if (!sheet) { Logger.log('WARNING: Active Sales Reps tab not found'); return map; }

  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    var name = String(data[i][1] || '').trim();  // Column B = Name
    var email = String(data[i][2] || '').trim();  // Column C = Email
    if (name && email && email.indexOf('@') > -1) {
      map[name.toLowerCase()] = email.toLowerCase();
    }
  }
  return map;
}


/**
 * Resolve a rep name to an email address.
 * Tries exact match first, then aliases, then fuzzy first-name match.
 */
function resolveEmail_(repName, nameEmailMap) {
  var key = repName.toLowerCase().trim();

  // Direct match
  if (nameEmailMap[key]) return nameEmailMap[key];

  // Alias match
  var aliasKey = NAME_ALIASES[key];
  if (aliasKey && nameEmailMap[aliasKey]) return nameEmailMap[aliasKey];

  // Fuzzy: try matching just the last name or first name
  // e.g., "Parker VanderBurgh" vs "Parker Vanderburgh" (case diff)
  var keyParts = key.split(/\s+/);
  for (var mapName in nameEmailMap) {
    var mapParts = mapName.split(/\s+/);
    // Match if first AND last name match (case-insensitive)
    if (keyParts.length >= 2 && mapParts.length >= 2) {
      if (keyParts[0] === mapParts[0] && keyParts[keyParts.length - 1] === mapParts[mapParts.length - 1]) {
        return nameEmailMap[mapName];
      }
    }
  }

  return null;  // No match — will get placeholder
}


/**
 * Generate a placeholder email for inactive reps.
 * Format: firstname.lastname@delagroup.inactive
 */
function generatePlaceholderEmail_(repName) {
  var parts = repName.toLowerCase().trim().replace(/[^a-z\s]/g, '').split(/\s+/).filter(function(p) { return p; });
  if (parts.length === 0) return 'unknown@delagroup.inactive';
  return parts.join('.') + '@delagroup.inactive';
}


/**
 * Parse the multi-line Order Summary field.
 * Format:
 *   Client Name
 *   DSI/SPM number
 *   - X New Phone(s)|Y BYOD(s)
 *   - Internet Air        (optional)
 *
 * Returns { clientName, dsi, newPhones, byods, air }
 */
function parseOrderSummary_(summary) {
  var result = { clientName: '', dsi: '', newPhones: 0, byods: 0, air: 0 };
  if (!summary) return result;

  var lines = summary.split('\n');

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;

    // Check for Internet Air
    if (/internet\s*air/i.test(line)) {
      result.air = 1;
      continue;
    }

    // Check for phones/BYODs line: "- X New Phone(s)|Y BYOD(s)"
    var phoneMatch = line.match(/(\d+)\s*New\s*Phone/i);
    var byodMatch = line.match(/(\d+)\s*BYOD/i);
    if (phoneMatch || byodMatch) {
      if (phoneMatch) result.newPhones = parseInt(phoneMatch[1], 10) || 0;
      if (byodMatch) result.byods = parseInt(byodMatch[1], 10) || 0;
      continue;
    }

    // Check for DSI/SPM number (starts with DSI or SPM, followed by digits)
    if (/^(DSI|SPM)\d/i.test(line)) {
      result.dsi = line;
      continue;
    }

    // First unrecognized line = client name
    if (!result.clientName) {
      result.clientName = line;
    }
  }

  return result;
}


/**
 * Dry-run: preview import without writing.
 * Logs the first 10 rows + summary stats.
 */
function previewImport() {
  var sourceSS = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  var nameEmailMap = buildNameEmailMap_(sourceSS);

  var salesSheet = sourceSS.getSheetByName(SOURCE_TAB_SALES);
  var salesData = salesSheet.getDataRange().getValues();

  var stats = { total: 0, air: 0, matched: 0, unmatched: 0, skipped: 0 };
  var sample = [];

  for (var i = 1; i < salesData.length; i++) {
    var srcRow = salesData[i];
    var dateRaw = srcRow[0];
    var repName = String(srcRow[1] || '').trim();
    var orderSummary = String(srcRow[3] || '').trim();

    if (!dateRaw || !repName) { stats.skipped++; continue; }
    stats.total++;

    var email = resolveEmail_(repName, nameEmailMap);
    if (email) { stats.matched++; } else { stats.unmatched++; }

    var parsed = parseOrderSummary_(orderSummary);
    if (parsed.air > 0) stats.air++;

    if (sample.length < 10) {
      sample.push({
        date: String(dateRaw),
        rep: repName,
        email: email || generatePlaceholderEmail_(repName) + ' (inactive)',
        client: parsed.clientName,
        dsi: parsed.dsi,
        phones: parsed.newPhones,
        byods: parsed.byods,
        air: parsed.air
      });
    }
  }

  Logger.log('═══ IMPORT PREVIEW ═══');
  Logger.log('Total sales: ' + stats.total);
  Logger.log('Air sales: ' + stats.air);
  Logger.log('Matched reps: ' + stats.matched);
  Logger.log('Unmatched (inactive): ' + stats.unmatched);
  Logger.log('Skipped: ' + stats.skipped);
  Logger.log('');
  Logger.log('── Sample rows (first 10): ──');
  for (var s = 0; s < sample.length; s++) {
    Logger.log(JSON.stringify(sample[s]));
  }
}


// ═══════════════════════════════════════════════════════════
// Roster + Teams Import
// ═══════════════════════════════════════════════════════════
// Run importDelagroupRoster() to populate _Roster_off_nds_006
// and _Teams_off_nds_006 from the source sheet's Active Sales Reps tab.

// Team structure from the Delagroup leaderboard
var DELAGROUP_TEAMS = [
  { id: 'team_sharks',    name: 'Sharks',    emoji: '🦈' },
  { id: 'team_sonder',    name: 'Sonder',    emoji: '🌀' },
  { id: 'team_ascension', name: 'Ascension', emoji: '🚀' },
  { id: 'team_prestige',  name: 'Prestige',  emoji: '👑' }
];

function importDelagroupRoster() {
  var sourceSS = SpreadsheetApp.openById(SOURCE_SHEET_ID);
  var destSS = SpreadsheetApp.openById(DEST_SHEET_ID);

  // ── 1. Read Active Sales Reps ──
  var repsSheet = sourceSS.getSheetByName(SOURCE_TAB_REPS);
  if (!repsSheet) { Logger.log('ERROR: Active Sales Reps tab not found'); return; }
  var repsData = repsSheet.getDataRange().getValues();

  // ── 2. Get or create _Roster_off_nds_006 tab ──
  var rosterSheet = destSS.getSheetByName(DEST_ROSTER_TAB);
  if (!rosterSheet) {
    rosterSheet = destSS.insertSheet(DEST_ROSTER_TAB);
    rosterSheet.appendRow(ROSTER_HEADERS);
    Logger.log('Created tab: ' + DEST_ROSTER_TAB);
  }

  // Check existing roster emails to avoid duplicates
  var existingData = rosterSheet.getDataRange().getValues();
  var existingEmails = {};
  for (var e = 1; e < existingData.length; e++) {
    var existEmail = String(existingData[e][0] || '').trim().toLowerCase();
    if (existEmail) existingEmails[existEmail] = true;
  }

  // ── 3. Parse Leaders (rows with "Leader" header) and Client Reps ──
  var rosterRows = [];
  var section = '';  // 'leader' or 'client'
  var now = new Date().toISOString().split('T')[0];
  var added = 0;
  var skipped = 0;

  for (var i = 0; i < repsData.length; i++) {
    var headerCell = String(repsData[i][1] || '').trim().toLowerCase();

    // Detect section headers
    if (headerCell === 'leader') { section = 'leader'; continue; }
    if (headerCell === 'client rep') { section = 'client'; continue; }

    var name = String(repsData[i][1] || '').trim();
    var email = String(repsData[i][2] || '').trim().toLowerCase();
    var team = String(repsData[i][3] || '').trim();  // Column D = Team

    // Skip rows without name or email
    if (!name || !email || email.indexOf('@') === -1) continue;

    // Skip if already in roster
    if (existingEmails[email]) {
      Logger.log('SKIP (exists): ' + email + ' — ' + name);
      skipped++;
      continue;
    }

    // Determine rank based on section
    var rank = 'rep';
    if (section === 'leader') rank = 'l1';

    // Build roster row: email, name, team, rank, deactivated, dateAdded, pinHash, phone, tableauName
    rosterRows.push([
      email,       // email
      name,        // name
      team,        // team
      rank,        // rank
      '',          // deactivated (empty = active)
      now,         // dateAdded
      '',          // pinHash (set via dashboard)
      '',          // phone
      ''           // tableauName
    ]);

    existingEmails[email] = true;
    added++;
  }

  // ── 4. Batch write roster rows ──
  if (rosterRows.length > 0) {
    var startRow = rosterSheet.getLastRow() + 1;
    rosterSheet.getRange(startRow, 1, rosterRows.length, 9).setValues(rosterRows);
    Logger.log('SUCCESS: Added ' + added + ' reps to ' + DEST_ROSTER_TAB);
  } else {
    Logger.log('WARNING: No new reps to add');
  }
  Logger.log('Skipped (already exist): ' + skipped);

  // ── 5. Create _Teams_off_nds_006 tab ──
  var teamsSheet = destSS.getSheetByName(DEST_TEAMS_TAB);
  if (!teamsSheet) {
    teamsSheet = destSS.insertSheet(DEST_TEAMS_TAB);
    teamsSheet.appendRow(TEAMS_HEADERS);
    Logger.log('Created tab: ' + DEST_TEAMS_TAB);

    // Add team rows
    var teamRows = [];
    for (var t = 0; t < DELAGROUP_TEAMS.length; t++) {
      var tm = DELAGROUP_TEAMS[t];
      teamRows.push([
        tm.id,     // teamId
        tm.name,   // name
        '',        // parentId (no hierarchy for now)
        '',        // leaderId
        tm.emoji,  // emoji
        now        // createdDate
      ]);
    }
    if (teamRows.length > 0) {
      teamsSheet.getRange(2, 1, teamRows.length, 6).setValues(teamRows);
      Logger.log('SUCCESS: Added ' + teamRows.length + ' teams to ' + DEST_TEAMS_TAB);
    }
  } else {
    Logger.log('Teams tab already exists — skipping team creation');
  }

  Logger.log('');
  Logger.log('═══ ROSTER IMPORT COMPLETE ═══');
  Logger.log('Next: Open the NDS dashboard and verify the leaderboard shows data.');
  Logger.log('Note: Reps may need PINs set via the dashboard People tab.');
}


// ═══════════════════════════════════════════════════════════
// Tab Rename: off_nds_006 → off_006
// ═══════════════════════════════════════════════════════════
// Run renameTabsToOff006() ONCE to align tab names with admin portal officeId.

function renameTabsToOff006() {
  var ss = SpreadsheetApp.openById(DEST_SHEET_ID);
  var renames = [
    { from: '_Sales_off_nds_006',  to: '_Sales_off_006' },
    { from: '_Roster_off_nds_006', to: '_Roster_off_006' },
    { from: '_Teams_off_nds_006',  to: '_Teams_off_006' }
  ];

  for (var i = 0; i < renames.length; i++) {
    var sheet = ss.getSheetByName(renames[i].from);
    if (sheet) {
      sheet.setName(renames[i].to);
      Logger.log('RENAMED: ' + renames[i].from + ' → ' + renames[i].to);
    } else {
      // Check if already renamed
      var existing = ss.getSheetByName(renames[i].to);
      if (existing) {
        Logger.log('ALREADY OK: ' + renames[i].to + ' (already exists)');
      } else {
        Logger.log('WARNING: ' + renames[i].from + ' not found');
      }
    }
  }

  Logger.log('');
  Logger.log('═══ TAB RENAME COMPLETE ═══');
  Logger.log('Tabs now use off_006 suffix to match admin portal.');
  Logger.log('Next: Redeploy NdsCode.gs and push nds-config.js to git.');
}
