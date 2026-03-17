// ═══════════════════════════════════════════════════════
// Owner Development Dashboard — Configuration
// ═══════════════════════════════════════════════════════
// Multi-team data mapping tool. Three teams (Maddie's, Cam's, NLR)
// each map owner names across their respective spreadsheets.

const OD_CONFIG = {
  // ── Apps Script Deployment ──
  appsScriptUrl: 'https://script.google.com/macros/s/AKfycbzX68JL0htLcwtyQYN7dttltwcEyHNQW3lcPeWzkS4PqKycCnvzDD6RQFbFvCvUcPfH/exec',
  apiKey: 'national-dash-2026-secret',

  // ── Session ──
  sessionKey: 'owner_dev_session',
  sessionDuration: 24 * 60 * 60 * 1000, // 24 hours

  // ── NLR Folder (Google Drive) ──
  nlrFolderId: '1hARjh3UH48CWhbYrYBJxFVwgynxapCjG',

  // ── Campaign Data Sources ──
  // Each campaign has a Google Sheet with owner data.
  // sheetId is blank for campaigns not yet connected.
  campaignSources: {
    'frontier':     { label: 'Frontier',           sheetId: '1OULSC_r8dCW2dlvIGeP6zLtNKANJ3mZMDPyE15tH3gM' },
    'verizon-fios': { label: 'Verizon Fios',       sheetId: '12J3HBdFQrqq5D7YwWEp93US40vmZz5n9mS-KMKXaXxA' },
    'att-nds':      { label: 'AT&T NDS/Verizon',   sheetId: '1kcUWR3EKgP-9wDct4vDyuQJ7IuS0cbcetY97dmVTY64' },
    'att-res':      { label: 'AT&T Residential',   sheetId: '1HvWJYox3JXvxmza63YBWAqKPtUGPFuaV-s-BOfbWGKM' },
    'rogers':       { label: 'Rogers',             sheetId: '1o1MPKrAzzeaU2JWMODkR9M3uY5rOhIKo-Q64armeTvE' },
    'leafguard':    { label: 'Leafguard',          sheetId: '' },
    'lumen':        { label: 'Lumen',              sheetId: '' }
  },

  // ── Teams ──
  // Each team has a label, accent color, and icon.
  teams: {
    maddie: { label: "Maddie's Team", color: '#8b5cf6', icon: '\u{1F451}' },
    cam:    { label: "Cam's Team",    color: '#3b82f6', icon: '\u{1F517}' },
    nlr:    { label: 'NLR Team',      color: '#0ea5a0', icon: '\u{1F4CB}' }
  }
};
