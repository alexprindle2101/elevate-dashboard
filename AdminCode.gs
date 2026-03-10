// ═══════════════════════════════════════════════════════
// Aptel Admin Dashboard — Google Apps Script Middleware
// ═══════════════════════════════════════════════════════
// Deploy as Web App: Execute as ME, Anyone can access.
// Set API_KEY in Script Properties (Project Settings > Script Properties).
// Paste this into a NEW Apps Script project attached to the Admin master sheet.

// === CONFIG ===
const ADMIN_ROSTER_TAB = '_AdminRoster';
const OFFICES_TAB = '_Offices';
const OWNERS_TAB = '_Owners';

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
      case ADMIN_ROSTER_TAB:
        sheet.appendRow(['email', 'name', 'role', 'pinHash', 'dateAdded', 'deactivated', 'assignedOwner', 'assignedOffices', 'managedBy']);
        break;
      case OFFICES_TAB:
        sheet.appendRow([
          'officeId', 'name', 'templateType', 'sheetId', 'appsScriptUrl',
          'apiKey', 'status', 'ownerEmail', 'ownerName', 'ownerLevel',
          'logoUrl', 'logoIconUrl', 'brandColors', 'createdDate', 'discordWebhookUrl',
          'headerLogoStyle', 'payrollManagerEmail', 'payrollMode', 'chatPlatform',
          'leaderboardEnabled', 'leaderboardHour'
        ]);
        break;
      case OWNERS_TAB:
        sheet.appendRow(['email', 'name', 'level', 'uplineEmail', 'phone', 'notes', 'dateAdded', 'deactivated', 'pinHash']);
        break;
    }
  }
  return sheet;
}

// Auto-migrate _AdminRoster headers if sheet has fewer than 9 columns
function migrateAdminRosterHeaders(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var expected = ['email', 'name', 'role', 'pinHash', 'dateAdded', 'deactivated', 'assignedOwner', 'assignedOffices', 'managedBy'];
  if (headers.length < expected.length) {
    for (var c = headers.length; c < expected.length; c++) {
      sheet.getRange(1, c + 1).setValue(expected[c]);
    }
  }
}

function hashPin(pin) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pin);
  return raw.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

// Case-insensitive row finder by column value
function findRowCI(sheet, colIdx, value) {
  const data = sheet.getDataRange().getValues();
  const target = (value || '').toString().trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][colIdx] || '').toString().trim().toLowerCase() === target) {
      return { rowIndex: i + 1, rowData: data[i] };
    }
  }
  return null;
}

function generateOfficeId() {
  const sheet = getOrCreateSheet(OFFICES_TAB);
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const id = (data[i][0] || '').toString();
    const match = id.match(/^off_(\d+)$/);
    if (match) maxNum = Math.max(maxNum, parseInt(match[1]));
  }
  return 'off_' + String(maxNum + 1).padStart(3, '0');
}

// Recursively collect all downline owner emails for a given owner
function getDownlineEmails(ownerEmail, allOwners) {
  var result = {};
  var queue = [ownerEmail.toLowerCase()];
  while (queue.length > 0) {
    var current = queue.shift();
    var entries = Object.values(allOwners);
    for (var i = 0; i < entries.length; i++) {
      var o = entries[i];
      if (o.uplineEmail === current && !result[o.email]) {
        result[o.email] = true;
        queue.push(o.email);
      }
    }
  }
  return result;
}


// ═══════════════════════════════════════════════════════
// doGet — Read operations
// ═══════════════════════════════════════════════════════

function doGet(e) {
  const key = (e.parameter.key || '').trim();
  if (!validateKey(key)) {
    return jsonResponse({ error: 'Invalid API key' });
  }

  const action = (e.parameter.action || 'readAll').trim();

  try {
    // Auto-migrate _AdminRoster headers on every read
    var adminSheet = getOrCreateSheet(ADMIN_ROSTER_TAB);
    migrateAdminRosterHeaders(adminSheet);

    switch (action) {
      case 'readAll':
        return jsonResponse({
          adminRoster: readAdminRoster(),
          offices: readOffices(),
          owners: readOwners()
        });

      case 'readOffices':
        return jsonResponse({ offices: readOffices() });

      case 'readAdminRoster':
        return jsonResponse({ adminRoster: readAdminRoster() });

      case 'readOwners':
        return jsonResponse({ owners: readOwners() });

      case 'listOfficesBasic': {
        var email = (e.parameter.email || '').trim().toLowerCase();
        if (email) {
          return jsonResponse({ offices: readOfficesBasicScoped(email) });
        }
        return jsonResponse({ offices: readOfficesBasic() });
      }

      // ── Scoped read — returns role-filtered data for a specific admin ──
      case 'readScoped': {
        var scopeEmail = (e.parameter.email || '').trim().toLowerCase();
        if (!scopeEmail) return jsonResponse({ error: 'Email parameter required' });
        return jsonResponse(readScoped(scopeEmail));
      }

      // ── SSO validation — checks if email is a valid admin or owner ──
      case 'validateAdminAuth': {
        var authEmail = (e.parameter.email || '').trim().toLowerCase();
        if (!authEmail) return jsonResponse({ valid: false, error: 'Email required' });
        var roster = readAdminRoster();
        var admin = roster[authEmail];
        if (admin) {
          if (admin.deactivated) return jsonResponse({ valid: false, error: 'Account deactivated' });
          return jsonResponse({ valid: true, email: admin.email, name: admin.name, role: admin.role, userType: 'admin' });
        }
        // Check _Owners
        var owners = readOwners();
        var owner = owners[authEmail];
        if (owner) {
          if (owner.deactivated) return jsonResponse({ valid: false, error: 'Account deactivated' });
          return jsonResponse({ valid: true, email: owner.email, name: owner.name, role: owner.level, userType: 'owner' });
        }
        return jsonResponse({ valid: false, error: 'Account not found' });
      }

      // ── Universal login — find which office a rep belongs to ──
      case 'findRepOffice': {
        var repEmail = (e.parameter.email || '').trim().toLowerCase();
        if (!repEmail) return jsonResponse({ found: false, error: 'Email required' });
        return jsonResponse(findRepOffice(repEmail));
      }

      // ── Get office config by officeId ──
      case 'getOfficeConfig': {
        var reqOfficeId = (e.parameter.officeId || '').trim();
        if (!reqOfficeId) return jsonResponse({ found: false, error: 'officeId required' });
        var allOffices = readOffices();
        for (var oi = 0; oi < allOffices.length; oi++) {
          if (allOffices[oi].officeId === reqOfficeId && allOffices[oi].status === 'active') {
            var o = allOffices[oi];
            return jsonResponse({
              found: true,
              office: {
                officeId: o.officeId,
                name: o.name,
                sheetId: o.sheetId,
                appsScriptUrl: o.appsScriptUrl,
                apiKey: o.apiKey,
                logoUrl: o.logoUrl || '',
                logoIconUrl: o.logoIconUrl || '',
                discordWebhookUrl: o.discordWebhookUrl || '',
                chatPlatform: o.chatPlatform || 'none',
                headerLogoStyle: o.headerLogoStyle || 'icon',
                payrollManagerEmail: o.payrollManagerEmail || '',
                payrollMode: o.payrollMode || 'commission-split',
                ownerEmail: o.ownerEmail || '',
                ownerName: o.ownerName || '',
                leaderboardEnabled: o.leaderboardEnabled || 'false',
                leaderboardHour: o.leaderboardHour || '22'
              }
            });
          }
        }
        return jsonResponse({ found: false, error: 'Office not found or inactive' });
      }

      default:
        return jsonResponse({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}


// ═══════════════════════════════════════════════════════
// doPost — Write operations
// ═══════════════════════════════════════════════════════

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (_) {
    return jsonResponse({ error: 'Invalid JSON' });
  }

  const key = (body.key || '').trim();
  if (!validateKey(key)) {
    return jsonResponse({ error: 'Invalid API key' });
  }

  const action = (body.action || '').trim();

  try {
    switch (action) {

      // ── PIN VALIDATION (dual-source: _AdminRoster then _Owners) ──
      case 'validatePin': {
        const email = (body.email || '').trim().toLowerCase();
        const pin = (body.pin || '').toString().trim();

        // 1) Check _AdminRoster first
        const adminSheet = getOrCreateSheet(ADMIN_ROSTER_TAB);
        const adminFound = findRowCI(adminSheet, 0, email);
        if (adminFound) {
          const row = adminFound.rowData;
          if ((row[5] || '').toString().toUpperCase() === 'TRUE') {
            return jsonResponse({ success: false, error: 'Account deactivated' });
          }
          const storedHash = (row[3] || '').toString().trim();
          if (!storedHash) {
            var rawRole = (row[2] || 'a3').toString().trim();
            var mappedRole = (rawRole === 'superadmin') ? 'a3' : rawRole;
            return jsonResponse({ success: true, firstLogin: true, name: row[1], role: mappedRole, userType: 'admin' });
          }
          const inputHash = hashPin(pin);
          if (inputHash === storedHash) {
            var rawRole2 = (row[2] || 'a3').toString().trim();
            var mappedRole2 = (rawRole2 === 'superadmin') ? 'a3' : rawRole2;
            return jsonResponse({ success: true, firstLogin: false, name: row[1], role: mappedRole2, userType: 'admin' });
          } else {
            return jsonResponse({ success: false, error: 'Incorrect PIN' });
          }
        }

        // 2) Check _Owners tab
        const ownerSheet = getOrCreateSheet(OWNERS_TAB);
        const ownerFound = findRowCI(ownerSheet, 0, email);
        if (ownerFound) {
          const oRow = ownerFound.rowData;
          if ((oRow[7] || '').toString().toUpperCase() === 'TRUE') {
            return jsonResponse({ success: false, error: 'Account deactivated' });
          }
          const oStoredHash = (oRow[8] || '').toString().trim(); // pinHash = col 8
          if (!oStoredHash) {
            return jsonResponse({ success: true, firstLogin: true, name: oRow[1], role: (oRow[2] || 'o1').toString().trim(), userType: 'owner' });
          }
          const oInputHash = hashPin(pin);
          if (oInputHash === oStoredHash) {
            return jsonResponse({ success: true, firstLogin: false, name: oRow[1], role: (oRow[2] || 'o1').toString().trim(), userType: 'owner' });
          } else {
            return jsonResponse({ success: false, error: 'Incorrect PIN' });
          }
        }

        return jsonResponse({ success: false, error: 'Email not found' });
      }

      // ── CREATE PIN (dual-source: _AdminRoster then _Owners) ──
      case 'createPin': {
        const email = (body.email || '').trim().toLowerCase();
        const pin = (body.pin || '').toString().trim();
        const hashed = hashPin(pin);

        // Check _AdminRoster first
        const adminSheet2 = getOrCreateSheet(ADMIN_ROSTER_TAB);
        const adminFound2 = findRowCI(adminSheet2, 0, email);
        if (adminFound2) {
          adminSheet2.getRange(adminFound2.rowIndex, 4).setValue(hashed); // pinHash = col 4
          return jsonResponse({ success: true });
        }

        // Check _Owners
        const ownerSheet2 = getOrCreateSheet(OWNERS_TAB);
        const ownerFound2 = findRowCI(ownerSheet2, 0, email);
        if (ownerFound2) {
          ownerSheet2.getRange(ownerFound2.rowIndex, 9).setValue(hashed); // pinHash = col 9 (index 8)
          return jsonResponse({ success: true });
        }

        return jsonResponse({ success: false, error: 'Email not found' });
      }

      // ── OFFICE CRUD ──
      case 'addOffice': {
        const sheet = getOrCreateSheet(OFFICES_TAB);
        const officeId = generateOfficeId();
        sheet.appendRow([
          officeId,
          body.name || '',
          body.templateType || 'att-b2b',
          body.sheetId || '',
          body.appsScriptUrl || '',
          body.apiKey || '',
          body.status || 'setup',
          body.ownerEmail || '',
          body.ownerName || '',
          body.ownerLevel || 'o1',
          body.logoUrl || '',
          body.logoIconUrl || '',
          body.brandColors || '',
          new Date().toISOString(),
          body.discordWebhookUrl || '',
          body.headerLogoStyle || 'icon',
          body.payrollManagerEmail || '',
          body.payrollMode || 'commission-split',
          body.chatPlatform || 'none',
          body.leaderboardEnabled || 'false',
          body.leaderboardHour || '22'
        ]);
        return jsonResponse({ success: true, officeId: officeId });
      }

      case 'updateOffice': {
        const sheet = getOrCreateSheet(OFFICES_TAB);
        const found = findRowCI(sheet, 0, body.officeId);
        if (!found) return jsonResponse({ success: false, error: 'Office not found' });

        const row = found.rowIndex;
        if (body.name !== undefined) sheet.getRange(row, 2).setValue(body.name);
        if (body.templateType !== undefined) sheet.getRange(row, 3).setValue(body.templateType);
        if (body.sheetId !== undefined) sheet.getRange(row, 4).setValue(body.sheetId);
        if (body.appsScriptUrl !== undefined) sheet.getRange(row, 5).setValue(body.appsScriptUrl);
        if (body.apiKey !== undefined) sheet.getRange(row, 6).setValue(body.apiKey);
        if (body.status !== undefined) sheet.getRange(row, 7).setValue(body.status);
        if (body.ownerEmail !== undefined) sheet.getRange(row, 8).setValue(body.ownerEmail);
        if (body.ownerName !== undefined) sheet.getRange(row, 9).setValue(body.ownerName);
        if (body.ownerLevel !== undefined) sheet.getRange(row, 10).setValue(body.ownerLevel);
        if (body.logoUrl !== undefined) sheet.getRange(row, 11).setValue(body.logoUrl);
        if (body.logoIconUrl !== undefined) sheet.getRange(row, 12).setValue(body.logoIconUrl);
        if (body.brandColors !== undefined) sheet.getRange(row, 13).setValue(body.brandColors);
        if (body.discordWebhookUrl !== undefined) sheet.getRange(row, 15).setValue(body.discordWebhookUrl);
        if (body.headerLogoStyle !== undefined) sheet.getRange(row, 16).setValue(body.headerLogoStyle);
        if (body.payrollManagerEmail !== undefined) sheet.getRange(row, 17).setValue(body.payrollManagerEmail);
        if (body.payrollMode !== undefined) sheet.getRange(row, 18).setValue(body.payrollMode);
        if (body.chatPlatform !== undefined) sheet.getRange(row, 19).setValue(body.chatPlatform);
        if (body.leaderboardEnabled !== undefined) sheet.getRange(row, 20).setValue(body.leaderboardEnabled);
        if (body.leaderboardHour !== undefined) sheet.getRange(row, 21).setValue(body.leaderboardHour);
        return jsonResponse({ success: true });
      }

      case 'deleteOffice': {
        const sheet = getOrCreateSheet(OFFICES_TAB);
        const found = findRowCI(sheet, 0, body.officeId);
        if (!found) return jsonResponse({ success: false, error: 'Office not found' });
        sheet.deleteRow(found.rowIndex);
        return jsonResponse({ success: true });
      }

      // ── ADMIN ROSTER CRUD ──
      case 'addAdmin': {
        const sheet = getOrCreateSheet(ADMIN_ROSTER_TAB);
        const existing = findRowCI(sheet, 0, body.email);
        if (existing) return jsonResponse({ success: false, error: 'Email already exists' });
        sheet.appendRow([
          (body.email || '').trim().toLowerCase(),
          body.name || '',
          body.role || 'a3',
          '', // pinHash — set on first login
          new Date().toISOString(),
          'FALSE',
          (body.assignedOwner || '').trim().toLowerCase(),
          body.assignedOffices || '',
          (body.managedBy || '').trim().toLowerCase()
        ]);
        return jsonResponse({ success: true });
      }

      case 'updateAdmin': {
        const sheet = getOrCreateSheet(ADMIN_ROSTER_TAB);
        const found = findRowCI(sheet, 0, body.email);
        if (!found) return jsonResponse({ success: false, error: 'Admin not found' });

        const row = found.rowIndex;
        if (body.name !== undefined) sheet.getRange(row, 2).setValue(body.name);
        if (body.role !== undefined) sheet.getRange(row, 3).setValue(body.role);
        if (body.deactivated !== undefined) sheet.getRange(row, 6).setValue(body.deactivated);
        if (body.assignedOwner !== undefined) sheet.getRange(row, 7).setValue((body.assignedOwner || '').trim().toLowerCase());
        if (body.assignedOffices !== undefined) sheet.getRange(row, 8).setValue(body.assignedOffices);
        if (body.managedBy !== undefined) sheet.getRange(row, 9).setValue((body.managedBy || '').trim().toLowerCase());
        return jsonResponse({ success: true });
      }

      // ── QC OFFICE SELECTION ──
      case 'updateQcOffices': {
        const sheet = getOrCreateSheet(ADMIN_ROSTER_TAB);
        const emailQC = (body.email || '').trim().toLowerCase();
        const foundQC = findRowCI(sheet, 0, emailQC);
        if (!foundQC) return jsonResponse({ success: false, error: 'User not found' });
        var roleQC = (foundQC.rowData[2] || '').toString().trim();
        if (roleQC !== 'qc_manager') return jsonResponse({ success: false, error: 'Not a QC Manager' });
        sheet.getRange(foundQC.rowIndex, 8).setValue(body.assignedOffices || '');
        return jsonResponse({ success: true });
      }

      // ── OWNER CRUD ──
      case 'addOwner': {
        const sheet = getOrCreateSheet(OWNERS_TAB);
        const email = (body.email || '').trim().toLowerCase();
        if (!email) return jsonResponse({ success: false, error: 'Email required' });
        const existing = findRowCI(sheet, 0, email);
        if (existing) return jsonResponse({ success: false, error: 'Owner already exists' });
        sheet.appendRow([
          email,
          body.name || '',
          body.level || 'o1',
          (body.uplineEmail || '').trim().toLowerCase(),
          body.phone || '',
          body.notes || '',
          new Date().toISOString(),
          'FALSE'
        ]);
        return jsonResponse({ success: true });
      }

      case 'updateOwner': {
        const sheet = getOrCreateSheet(OWNERS_TAB);
        const email = (body.email || '').trim().toLowerCase();
        const found = findRowCI(sheet, 0, email);
        if (!found) return jsonResponse({ success: false, error: 'Owner not found' });
        const row = found.rowIndex;
        if (body.name !== undefined)        sheet.getRange(row, 2).setValue(body.name);
        if (body.level !== undefined)       sheet.getRange(row, 3).setValue(body.level);
        if (body.uplineEmail !== undefined) sheet.getRange(row, 4).setValue((body.uplineEmail || '').trim().toLowerCase());
        if (body.phone !== undefined)       sheet.getRange(row, 5).setValue(body.phone);
        if (body.notes !== undefined)       sheet.getRange(row, 6).setValue(body.notes);
        if (body.deactivated !== undefined) sheet.getRange(row, 8).setValue(body.deactivated);
        return jsonResponse({ success: true });
      }

      case 'deleteOwner': {
        const sheet = getOrCreateSheet(OWNERS_TAB);
        const email = (body.email || '').trim().toLowerCase();
        const found = findRowCI(sheet, 0, email);
        if (!found) return jsonResponse({ success: false, error: 'Owner not found' });
        sheet.deleteRow(found.rowIndex);
        return jsonResponse({ success: true });
      }

      default:
        return jsonResponse({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}


// ═══════════════════════════════════════════════════════
// READ HELPERS
// ═══════════════════════════════════════════════════════

function readAdminRoster() {
  const sheet = getOrCreateSheet(ADMIN_ROSTER_TAB);
  const data = sheet.getDataRange().getValues();
  const roster = {};

  for (let i = 1; i < data.length; i++) {
    const email = (data[i][0] || '').toString().trim().toLowerCase();
    if (!email) continue;
    var rawRole = (data[i][2] || 'a3').toString().trim();
    // Backward compat: map legacy 'superadmin' to 'a3'
    if (rawRole === 'superadmin') rawRole = 'a3';
    roster[email] = {
      email: email,
      name: (data[i][1] || '').toString().trim(),
      role: rawRole,
      hasPinSet: !!(data[i][3] || '').toString().trim(),
      dateAdded: (data[i][4] || '').toString(),
      deactivated: (data[i][5] || '').toString().toUpperCase() === 'TRUE',
      assignedOwner: (data[i][6] || '').toString().trim().toLowerCase(),
      assignedOffices: (data[i][7] || '').toString().trim(),
      managedBy: (data[i][8] || '').toString().trim().toLowerCase()
    };
  }
  return roster;
}

function readOwners() {
  const sheet = getOrCreateSheet(OWNERS_TAB);
  const data = sheet.getDataRange().getValues();
  const owners = {};

  for (let i = 1; i < data.length; i++) {
    const email = (data[i][0] || '').toString().trim().toLowerCase();
    if (!email) continue;
    owners[email] = {
      email: email,
      name: (data[i][1] || '').toString().trim(),
      level: (data[i][2] || 'o1').toString().trim(),
      uplineEmail: (data[i][3] || '').toString().trim().toLowerCase(),
      phone: (data[i][4] || '').toString().trim(),
      notes: (data[i][5] || '').toString().trim(),
      dateAdded: (data[i][6] || '').toString(),
      deactivated: (data[i][7] || '').toString().toUpperCase() === 'TRUE',
      hasPinSet: !!(data[i][8] || '').toString().trim()
    };
  }
  return owners;
}

function readOffices() {
  const sheet = getOrCreateSheet(OFFICES_TAB);
  const data = sheet.getDataRange().getValues();
  const offices = [];

  for (let i = 1; i < data.length; i++) {
    const officeId = (data[i][0] || '').toString().trim();
    if (!officeId) continue;
    offices.push({
      officeId: officeId,
      name: (data[i][1] || '').toString().trim(),
      templateType: (data[i][2] || 'att-b2b').toString().trim(),
      sheetId: (data[i][3] || '').toString().trim(),
      appsScriptUrl: (data[i][4] || '').toString().trim(),
      apiKey: (data[i][5] || '').toString().trim(),
      status: (data[i][6] || 'setup').toString().trim(),
      ownerEmail: (data[i][7] || '').toString().trim().toLowerCase(),
      ownerName: (data[i][8] || '').toString().trim(),
      ownerLevel: (data[i][9] || 'o1').toString().trim(),
      logoUrl: (data[i][10] || '').toString().trim(),
      logoIconUrl: (data[i][11] || '').toString().trim(),
      brandColors: (data[i][12] || '').toString().trim(),
      createdDate: (data[i][13] || '').toString(),
      discordWebhookUrl: (data[i][14] || '').toString().trim(),
      headerLogoStyle: (data[i][15] || '').toString().trim() || 'icon',
      payrollManagerEmail: (data[i][16] || '').toString().trim().toLowerCase(),
      payrollMode: (data[i][17] || 'commission-split').toString().trim(),
      chatPlatform: (data[i][18] || 'none').toString().trim(),
      leaderboardEnabled: (data[i][19] || 'false').toString().trim(),
      leaderboardHour: (data[i][20] || '22').toString().trim()
    });
  }
  return offices;
}

// Lightweight office list — all active offices (unscoped)
function readOfficesBasic() {
  const sheet = getOrCreateSheet(OFFICES_TAB);
  const data = sheet.getDataRange().getValues();
  const offices = [];
  for (let i = 1; i < data.length; i++) {
    const officeId = (data[i][0] || '').toString().trim();
    const status = (data[i][6] || 'setup').toString().trim();
    if (!officeId || status !== 'active') continue;
    offices.push({
      officeId: officeId,
      name: (data[i][1] || '').toString().trim(),
      sheetId: (data[i][3] || '').toString().trim(),
      appsScriptUrl: (data[i][4] || '').toString().trim(),
      apiKey: (data[i][5] || '').toString().trim(),
      logoUrl: (data[i][10] || '').toString().trim(),
      logoIconUrl: (data[i][11] || '').toString().trim(),
      discordWebhookUrl: (data[i][14] || '').toString().trim(),
      chatPlatform: (data[i][18] || 'none').toString().trim(),
      leaderboardEnabled: (data[i][19] || 'false').toString().trim(),
      leaderboardHour: (data[i][20] || '22').toString().trim()
    });
  }
  return offices;
}

// ── Universal login: find which office a rep email belongs to ──
// Scans _Roster_{officeId} tabs across all active offices.
// Groups by sheetId to avoid opening the same spreadsheet multiple times.
function findRepOffice(email) {
  var offices = readOffices();
  var activeOffices = [];
  for (var i = 0; i < offices.length; i++) {
    if (offices[i].status === 'active' && offices[i].sheetId) {
      activeOffices.push(offices[i]);
    }
  }
  if (!activeOffices.length) return { found: false, error: 'No active offices configured' };

  // Group offices by sheetId (all AT&T B2B offices share one sheet)
  var sheetGroups = {};
  for (var j = 0; j < activeOffices.length; j++) {
    var sid = activeOffices[j].sheetId;
    if (!sheetGroups[sid]) sheetGroups[sid] = [];
    sheetGroups[sid].push(activeOffices[j]);
  }

  // Scan each sheet's roster tabs
  for (var sheetId in sheetGroups) {
    var ss;
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch (err) {
      Logger.log('findRepOffice: Cannot open sheet ' + sheetId + ': ' + err.message);
      continue;
    }

    var groupOffices = sheetGroups[sheetId];
    for (var k = 0; k < groupOffices.length; k++) {
      var office = groupOffices[k];
      var rosterTab = ss.getSheetByName('_Roster_' + office.officeId);
      if (!rosterTab) continue;

      var rosterData = rosterTab.getDataRange().getValues();
      for (var r = 1; r < rosterData.length; r++) {
        var rowEmail = String(rosterData[r][0] || '').trim().toLowerCase();
        if (rowEmail !== email) continue;

        // Check deactivated (col 4)
        var deactivated = rosterData[r][4] === true || String(rosterData[r][4]).toUpperCase() === 'TRUE';
        if (deactivated) continue;

        // Found active rep
        var pinVal = String(rosterData[r][6] || '').trim();
        return {
          found: true,
          office: {
            officeId: office.officeId,
            name: office.name,
            sheetId: office.sheetId,
            appsScriptUrl: office.appsScriptUrl,
            apiKey: office.apiKey,
            logoUrl: office.logoUrl || '',
            logoIconUrl: office.logoIconUrl || ''
          },
          rosterEntry: {
            name: String(rosterData[r][1] || '').trim(),
            team: String(rosterData[r][2] || '').trim(),
            rank: String(rosterData[r][3] || 'rep').trim(),
            hasPin: pinVal.length > 0 && pinVal !== 'undefined',
            deactivated: false
          }
        };
      }
    }
  }

  return { found: false };
}

// Scoped office list — filtered by admin role
function readOfficesBasicScoped(adminEmail) {
  var roster = readAdminRoster();
  var admin = roster[adminEmail];
  if (!admin) return [];

  var allBasic = readOfficesBasic();

  // a3: all active offices
  if (admin.role === 'a3') return allBasic;

  // a1: only assigned offices
  if (admin.role === 'a1') {
    var ids = {};
    var parts = (admin.assignedOffices || '').split(',');
    for (var i = 0; i < parts.length; i++) {
      var id = parts[i].trim();
      if (id) ids[id] = true;
    }
    return allBasic.filter(function(o) { return ids[o.officeId]; });
  }

  // a2: offices under assigned owner + downline
  if (admin.role === 'a2' && admin.assignedOwner) {
    var allOwners = readOwners();
    var downline = getDownlineEmails(admin.assignedOwner, allOwners);
    downline[admin.assignedOwner] = true;
    var allOffices = readOffices();
    var activeIds = {};
    for (var j = 0; j < allOffices.length; j++) {
      var off = allOffices[j];
      if (off.status === 'active' && downline[off.ownerEmail]) {
        activeIds[off.officeId] = true;
      }
    }
    return allBasic.filter(function(o) { return activeIds[o.officeId]; });
  }

  return [];
}


// ═══════════════════════════════════════════════════════
// SCOPED READ — Role-filtered data for admin portal
// ═══════════════════════════════════════════════════════

function readScoped(email) {
  var roster = readAdminRoster();
  var admin = roster[email];

  // If not in _AdminRoster, check _Owners tab
  if (!admin) {
    var allOwners = readOwners();
    var owner = allOwners[email];
    if (!owner) return { error: 'Account not found' };
    if (owner.deactivated) return { error: 'Account deactivated' };

    var allOffices = readOffices();
    var ownerRole = owner.level || 'o1';

    // Owner sees only offices where they are the owner
    var ownerOffices = allOffices.filter(function(o) {
      return (o.ownerEmail || '').toLowerCase() === email;
    });

    // For o2+ owners, also include offices owned by their downline
    var downline = {};
    if (ownerRole === 'o2' || ownerRole === 'o3' || ownerRole === 'o4') {
      downline = getDownlineEmails(email, allOwners);
      ownerOffices = allOffices.filter(function(o) {
        var oe = (o.ownerEmail || '').toLowerCase();
        return oe === email || downline[oe];
      });
    }

    // Build set of owner's office IDs for admin filtering
    var ownerOfficeIds = {};
    for (var oi = 0; oi < ownerOffices.length; oi++) {
      ownerOfficeIds[ownerOffices[oi].officeId] = true;
    }

    // Return admins eligible for the owner's offices:
    // a3 (all offices), a2 assigned to this owner or downline, a1 assigned to owner's offices
    var scopedAdmins = {};
    var rosterKeys = Object.keys(roster);
    for (var ri = 0; ri < rosterKeys.length; ri++) {
      var a = roster[rosterKeys[ri]];
      if (a.deactivated) continue;
      // a3 — super admins see all offices
      if (a.role === 'a3') { scopedAdmins[rosterKeys[ri]] = a; continue; }
      // a2 — assigned to this owner or someone in their downline
      if (a.role === 'a2') {
        var ao = (a.assignedOwner || '').toLowerCase();
        if (ao === email || downline[ao]) { scopedAdmins[rosterKeys[ri]] = a; }
        continue;
      }
      // a1 — assigned to one of the owner's offices
      if (a.role === 'a1') {
        var assignedIds = (a.assignedOffices || '').split(',');
        for (var ai = 0; ai < assignedIds.length; ai++) {
          if (ownerOfficeIds[assignedIds[ai].trim()]) {
            scopedAdmins[rosterKeys[ri]] = a;
            break;
          }
        }
      }
    }

    // Scoped owners: self + downline
    var scopedOwners = {};
    scopedOwners[email] = owner;
    var ownerKeys = Object.keys(allOwners);
    for (var di = 0; di < ownerKeys.length; di++) {
      if (downline[ownerKeys[di]]) scopedOwners[ownerKeys[di]] = allOwners[ownerKeys[di]];
    }

    return {
      offices: ownerOffices,
      owners: scopedOwners,
      adminRoster: scopedAdmins,
      role: ownerRole,
      userType: 'owner'
    };
  }

  var allOffices = readOffices();
  var allOwners = readOwners();

  // ── a3 (Super Admin): full access ──
  if (admin.role === 'a3') {
    return {
      adminRoster: roster,
      offices: allOffices,
      owners: allOwners,
      role: 'a3',
      userType: 'admin'
    };
  }

  // ── a2 (Org Admin): scoped to assigned owner + downline ──
  if (admin.role === 'a2') {
    var ownerEmail = admin.assignedOwner || '';
    var downline = {};
    if (ownerEmail) {
      downline = getDownlineEmails(ownerEmail, allOwners);
      downline[ownerEmail] = true;
    }

    // Offices under owner org
    var scopedOffices = allOffices.filter(function(o) {
      return downline[o.ownerEmail];
    });

    // Owners in the tree
    var scopedOwners = {};
    var ownerKeys = Object.keys(allOwners);
    for (var i = 0; i < ownerKeys.length; i++) {
      if (downline[ownerKeys[i]]) {
        scopedOwners[ownerKeys[i]] = allOwners[ownerKeys[i]];
      }
    }

    // Admins: those managed by this a2, plus self
    var scopedAdmins = {};
    var rosterKeys = Object.keys(roster);
    for (var j = 0; j < rosterKeys.length; j++) {
      var a = roster[rosterKeys[j]];
      if (a.managedBy === email || rosterKeys[j] === email) {
        scopedAdmins[rosterKeys[j]] = a;
      }
    }

    return {
      adminRoster: scopedAdmins,
      offices: scopedOffices,
      owners: scopedOwners,
      role: 'a2',
      userType: 'admin',
      assignedOwner: ownerEmail
    };
  }

  // ── a1 (Admin): only assigned offices, team read-only ──
  if (admin.role === 'a1') {
    var ids = {};
    var parts = (admin.assignedOffices || '').split(',');
    for (var k = 0; k < parts.length; k++) {
      var id = parts[k].trim();
      if (id) ids[id] = true;
    }
    var scopedOffices1 = allOffices.filter(function(o) { return ids[o.officeId]; });

    // Admins with same managedBy (the a2 team), plus self
    var scopedAdmins1 = {};
    var rosterKeys1 = Object.keys(roster);
    for (var m = 0; m < rosterKeys1.length; m++) {
      var a1 = roster[rosterKeys1[m]];
      if (admin.managedBy && a1.managedBy === admin.managedBy) {
        scopedAdmins1[rosterKeys1[m]] = a1;
      }
    }
    scopedAdmins1[email] = roster[email]; // always include self

    return {
      adminRoster: scopedAdmins1,
      offices: scopedOffices1,
      owners: {},
      role: 'a1',
      userType: 'admin'
    };
  }

  // ── qc_manager: sees ALL offices (for toggle selection) + their QC team ──
  if (admin.role === 'qc_manager') {
    // QC team: admins where managedBy = this qc_manager AND role = 'qc'
    var scopedAdminsQM = {};
    var rosterKeysQM = Object.keys(roster);
    for (var qm = 0; qm < rosterKeysQM.length; qm++) {
      var aqm = roster[rosterKeysQM[qm]];
      if (aqm.role === 'qc' && aqm.managedBy === email) {
        scopedAdminsQM[rosterKeysQM[qm]] = aqm;
      }
    }
    scopedAdminsQM[email] = roster[email]; // include self

    return {
      adminRoster: scopedAdminsQM,
      offices: allOffices,
      owners: {},
      role: 'qc_manager',
      userType: 'admin',
      assignedOffices: admin.assignedOffices || ''
    };
  }

  // ── qc: sees only assigned offices ──
  if (admin.role === 'qc') {
    var idsQC = {};
    var partsQC = (admin.assignedOffices || '').split(',');
    for (var qc = 0; qc < partsQC.length; qc++) {
      var idQC = partsQC[qc].trim();
      if (idQC) idsQC[idQC] = true;
    }
    var scopedOfficesQC = allOffices.filter(function(o) { return idsQC[o.officeId]; });

    var scopedAdminsQC = {};
    scopedAdminsQC[email] = roster[email]; // self only

    return {
      adminRoster: scopedAdminsQC,
      offices: scopedOfficesQC,
      owners: {},
      role: 'qc',
      userType: 'admin',
      assignedOffices: admin.assignedOffices || '',
      managedBy: admin.managedBy || ''
    };
  }

  return { error: 'Unknown role: ' + admin.role };
}


// ═══════════════════════════════════════════════════════
// CENTRALIZED LEADERBOARD SCHEDULER
// ═══════════════════════════════════════════════════════
// Run setupLeaderboardScheduler() ONCE to create an hourly trigger.
// checkLeaderboardPosts() runs every hour — for each office with
// leaderboardEnabled=true, if the current hour matches leaderboardHour,
// it calls the office's ?action=leaderboardText endpoint and posts
// the result to the configured webhook (Discord/GroupMe).

/**
 * ONE-TIME SETUP — run from Apps Script editor to create hourly trigger.
 */
function setupLeaderboardScheduler() {
  // Remove existing leaderboard triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'checkLeaderboardPosts') {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log('Removed existing checkLeaderboardPosts trigger');
    }
  }

  // Create hourly trigger
  ScriptApp.newTrigger('checkLeaderboardPosts')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('✅ Leaderboard scheduler created — runs every hour.');
}

/**
 * Hourly check — loops all offices, posts leaderboard text where due.
 */
function checkLeaderboardPosts() {
  var now = new Date();
  var currentHour = now.getHours(); // 0-23 in script timezone

  var offices = readOffices();

  for (var i = 0; i < offices.length; i++) {
    var office = offices[i];

    // Skip inactive, disabled, or wrong hour
    if (office.status !== 'active') continue;
    if (office.leaderboardEnabled !== 'true') continue;
    if (String(office.leaderboardHour) !== String(currentHour)) continue;

    // Must have a webhook configured
    var platform = (office.chatPlatform || 'none').toLowerCase();
    var webhookUrl = (office.discordWebhookUrl || '').trim();
    if (!webhookUrl || platform === 'none') continue;

    // Must have an Apps Script URL to call
    var appsScriptUrl = (office.appsScriptUrl || '').trim();
    var apiKey = (office.apiKey || '').trim();
    if (!appsScriptUrl) continue;

    try {
      // Build the leaderboardText endpoint URL
      var url = 'https://script.google.com/macros/s/' + appsScriptUrl + '/exec'
        + '?action=leaderboardText'
        + '&key=' + encodeURIComponent(apiKey)
        + '&officeId=' + encodeURIComponent(office.officeId)
        + '&sheetId=' + encodeURIComponent(office.sheetId)
        + '&officeName=' + encodeURIComponent(office.name);

      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var code = resp.getResponseCode();
      if (code !== 200) {
        Logger.log('[Leaderboard] ' + office.name + ' — HTTP ' + code);
        continue;
      }

      var json = JSON.parse(resp.getContentText());
      var text = (json.text || '').trim();
      if (!text) {
        Logger.log('[Leaderboard] ' + office.name + ' — empty text');
        continue;
      }

      // Post to the office's webhook
      var postUrl, postPayload;
      if (platform === 'groupme') {
        postUrl = 'https://api.groupme.com/v3/bots/post';
        postPayload = JSON.stringify({ bot_id: webhookUrl, text: text });
      } else {
        postUrl = webhookUrl;
        postPayload = JSON.stringify({ content: text });
      }

      var postResp = UrlFetchApp.fetch(postUrl, {
        method: 'post',
        contentType: 'application/json',
        payload: postPayload,
        muteHttpExceptions: true
      });
      Logger.log('[Leaderboard] ' + office.name + ' — posted to ' + platform + ' (' + postResp.getResponseCode() + ')');
    } catch (err) {
      Logger.log('[Leaderboard] ' + office.name + ' — ERROR: ' + err.message);
    }
  }
}
