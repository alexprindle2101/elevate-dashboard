// ═══════════════════════════════════════════════════════
// Owner Development Dashboard — Configuration
// ═══════════════════════════════════════════════════════
// Multi-role, campaign-ownership-based access control.
// Org Managers own campaigns and grant access to Nationals + others.

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
    'frontier':        { label: 'Frontier',           sheetId: '1WWpLQTCyvPmJbx3jjowFszwOF_JUnjS6tzu-eAASwk0' },
    'verizon-fios':    { label: 'Verizon Fios',       sheetId: '12J3HBdFQrqq5D7YwWEp93US40vmZz5n9mS-KMKXaXxA' },
    'att-nds':         { label: 'AT&T NDS/Verizon',   sheetId: '1kcUWR3EKgP-9wDct4vDyuQJ7IuS0cbcetY97dmVTY64' },
    'att-res':         { label: 'AT&T Residential',   sheetId: '1HvWJYox3JXvxmza63YBWAqKPtUGPFuaV-s-BOfbWGKM' },
    'rogers':          { label: 'Rogers',             sheetId: '1o1MPKrAzzeaU2JWMODkR9M3uY5rOhIKo-Q64armeTvE' },
    'leafguard':       { label: 'Leafguard',          sheetId: '10Fy5XFWCuBmDwvpl4PG4FJT4krwX2ZqN12ARvQLpSuM' },
    'lumen':           { label: 'Lumen',              sheetId: '1P4DYlcV1hgNkaAapk3tWD7ytcRXw4K1n7R6EMKPCoSA' },
    'att-b2b':         { label: 'AT&T B2B',           sheetId: '1sxauFjNjq4_rRYM2PAl5cyOyHF3Hg4OkO-t_hLKDJB8' },
    'box-energy':      { label: 'Box Energy',          sheetId: '1_PVzLcmlo6EzySNRfah-r-NLka3IthxzLZb4TjrDgtE', hidden: true }
  },

  // ── Roles ──
  // Defines all valid roles and their base permissions.
  // Column edit rights:
  //   sourceTab  = campaign tab map (which tab the owner's data is on)
  //   camCompany = Better Image Solutions company column
  //   nlrFile    = NLR workbook + tab columns
  roles: {
    superadmin:    { label: 'Super Admin',           rank: 100 },
    aptel:         { label: 'Aptel',                 rank: 90  },
    national:      { label: 'National Consultant',   rank: 80  },
    org_manager:   { label: 'Org Manager',           rank: 70  },
    admin:         { label: 'Admin',                 rank: 60  },
    nlr_manager:   { label: 'NLR Manager',           rank: 50  },
    nlr:           { label: 'NLR',                   rank: 40  },
    bis_manager:   { label: 'BIS Manager',           rank: 50  },
    bis:           { label: 'BIS',                   rank: 40  }
  },

  // ── Tab Access by Role ──
  // Which nav tabs each role can see.
  // 'edit' = full access, 'view' = read-only, false = hidden
  tabAccess: {
    superadmin:    { mapping: 'edit', team: 'edit', coach: 'edit', planning: 'edit', tools: 'edit' },
    aptel:         { mapping: 'view', team: 'view', coach: 'view', planning: 'view', tools: false  },
    national:      { mapping: false,  team: false,  coach: 'edit', planning: false,  tools: false  },
    org_manager:   { mapping: 'edit', team: 'edit', coach: 'edit', planning: 'edit', tools: 'edit' },
    admin:         { mapping: 'edit', team: false,  coach: 'edit', planning: 'edit', tools: 'edit' },
    nlr_manager:   { mapping: 'edit', team: 'edit', coach: 'view', planning: 'view', tools: false  },
    nlr:           { mapping: 'edit', team: false,  coach: 'view', planning: 'view', tools: false  },
    bis_manager:   { mapping: 'edit', team: 'edit', coach: 'view', planning: 'view', tools: false  },
    bis:           { mapping: 'edit', team: false,  coach: 'view', planning: 'view', tools: false  }
  },

  // ── Column Edit Rights by Role ──
  // Which mapping columns each role can edit.
  // Org Managers + Admins edit sourceTab for their own campaigns (checked dynamically).
  // NLR roles edit nlrFile + nlrTab. BIS roles edit camCompany.
  columnEdit: {
    superadmin:    { sourceTab: true, camCompany: true, nlrFile: true, nlrTab: true },
    aptel:         { sourceTab: false, camCompany: false, nlrFile: false, nlrTab: false },
    national:      { sourceTab: false, camCompany: false, nlrFile: false, nlrTab: false }, // edit via grants checked dynamically
    org_manager:   { sourceTab: true, camCompany: false, nlrFile: false, nlrTab: false },  // only for owned campaigns
    admin:         { sourceTab: true, camCompany: false, nlrFile: false, nlrTab: false },  // inherits from OM
    nlr_manager:   { sourceTab: false, camCompany: false, nlrFile: true, nlrTab: true },
    nlr:           { sourceTab: false, camCompany: false, nlrFile: true, nlrTab: true },
    bis_manager:   { sourceTab: false, camCompany: true, nlrFile: false, nlrTab: false },
    bis:           { sourceTab: false, camCompany: true, nlrFile: false, nlrTab: false }
  },

  // ── Teams ──
  // National-based teams + functional teams.
  // National teams are dynamic (keyed by National's identifier).
  // Functional teams (nlr, bis) are cross-cutting.
  teams: {
    ken:  { label: "Ken's Team",                type: 'national', color: '#8b5cf6', icon: '\u{1F451}' },
    van:  { label: "Van's Team",                type: 'national', color: '#f59e0b', icon: '\u{1F31F}' },
    rafael: { label: "Rafael's Team",           type: 'national', color: '#10b981', icon: '\u{1F680}' },
    sam:  { label: "Sam's Team",                type: 'national', color: '#ef4444', icon: '\u{1F525}' },
    nlr:  { label: 'NLR Team',                  type: 'functional', color: '#0ea5a0', icon: '\u{1F4CB}' },
    bis:  { label: 'Better Image Solutions',    type: 'functional', color: '#3b82f6', icon: '\u{1F517}' }
  },

  // ── Superadmins ──
  // Full edit + View-As across all teams and campaigns.
  superadmins: [
    'alex.aspirehr@gmail.com'
  ],

  // ── Login aliases (shorthand → full email) ──
  loginAliases: {
    'alex': 'alex.aspirehr@gmail.com'
  }
};
