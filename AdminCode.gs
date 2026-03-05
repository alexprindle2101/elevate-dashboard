// ═══════════════════════════════════════════════════════
// Aptel Admin Dashboard — Google Apps Script Middleware
// ═══════════════════════════════════════════════════════
// Deploy as Web App: Execute as ME, Anyone can access.
// Set API_KEY in Script Properties (Project Settings > Script Properties).
// Paste this into a NEW Apps Script project attached to the Admin master sheet.

// === CONFIG ===
const ADMIN_ROSTER_TAB = '_AdminRoster';
const OFFICES_TAB = '_Offices';

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
        sheet.appendRow(['email', 'name', 'role', 'pinHash', 'dateAdded', 'deactivated']);
        break;
      case OFFICES_TAB:
        sheet.appendRow([
          'officeId', 'name', 'templateType', 'sheetId', 'appsScriptUrl',
          'apiKey', 'status', 'ownerEmail', 'ownerName', 'logoUrl',
          'brandColors', 'createdDate'
        ]);
        break;
    }
  }
  return sheet;
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
    switch (action) {
      case 'readAll':
        return jsonResponse({
          adminRoster: readAdminRoster(),
          offices: readOffices()
        });

      case 'readOffices':
        return jsonResponse({ offices: readOffices() });

      case 'readAdminRoster':
        return jsonResponse({ adminRoster: readAdminRoster() });

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

      // ── PIN VALIDATION ──
      case 'validatePin': {
        const email = (body.email || '').trim().toLowerCase();
        const pin = (body.pin || '').toString().trim();
        const sheet = getOrCreateSheet(ADMIN_ROSTER_TAB);
        const found = findRowCI(sheet, 0, email);
        if (!found) return jsonResponse({ success: false, error: 'Email not found' });

        const row = found.rowData;
        // columns: email(0), name(1), role(2), pinHash(3), dateAdded(4), deactivated(5)
        if ((row[5] || '').toString().toUpperCase() === 'TRUE') {
          return jsonResponse({ success: false, error: 'Account deactivated' });
        }

        const storedHash = (row[3] || '').toString().trim();
        if (!storedHash) {
          // No PIN set — first login
          return jsonResponse({ success: true, firstLogin: true, name: row[1], role: row[2] });
        }

        const inputHash = hashPin(pin);
        if (inputHash === storedHash) {
          return jsonResponse({ success: true, firstLogin: false, name: row[1], role: row[2] });
        } else {
          return jsonResponse({ success: false, error: 'Incorrect PIN' });
        }
      }

      // ── CREATE PIN ──
      case 'createPin': {
        const email = (body.email || '').trim().toLowerCase();
        const pin = (body.pin || '').toString().trim();
        const sheet = getOrCreateSheet(ADMIN_ROSTER_TAB);
        const found = findRowCI(sheet, 0, email);
        if (!found) return jsonResponse({ success: false, error: 'Email not found' });

        const hashed = hashPin(pin);
        sheet.getRange(found.rowIndex, 4).setValue(hashed); // pinHash = col 4
        return jsonResponse({ success: true });
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
          body.logoUrl || '',
          body.brandColors || '',
          new Date().toISOString()
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
        if (body.logoUrl !== undefined) sheet.getRange(row, 10).setValue(body.logoUrl);
        if (body.brandColors !== undefined) sheet.getRange(row, 11).setValue(body.brandColors);
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
          body.role || 'superadmin',
          '', // pinHash — set on first login
          new Date().toISOString(),
          'FALSE'
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
    roster[email] = {
      email: email,
      name: (data[i][1] || '').toString().trim(),
      role: (data[i][2] || 'superadmin').toString().trim(),
      hasPinSet: !!(data[i][3] || '').toString().trim(),
      dateAdded: (data[i][4] || '').toString(),
      deactivated: (data[i][5] || '').toString().toUpperCase() === 'TRUE'
    };
  }
  return roster;
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
      ownerEmail: (data[i][7] || '').toString().trim(),
      ownerName: (data[i][8] || '').toString().trim(),
      logoUrl: (data[i][9] || '').toString().trim(),
      brandColors: (data[i][10] || '').toString().trim(),
      createdDate: (data[i][11] || '').toString()
    });
  }
  return offices;
}
