// ═══════════════════════════════════════════════════════
// NDS — Office Configuration
// ═══════════════════════════════════════════════════════
// Each office gets its own config. For multi-office,
// swap this file or load from a config sheet.

const OFFICE_CONFIG = {
  officeName: "Delagroup",
  officeId: "off_006",  // Per-office identifier — tabs are suffixed with this

  // ── Apps Script Middleware ──
  // All reads and writes go through Apps Script to keep the sheet private.
  // For shared campaign sheets, all AT&T B2B offices use the same Code.gs URL.
  appsScriptUrl: "https://script.google.com/macros/s/AKfycbyyG05ebEPT-MyrFHLvGtQwokFh5_HOP5B1OZh1Mc6FK_1_Rt800DHMw1o0MrCoEJM2hg/exec",
  apiKey: "nds-secret-key-2026",

  // ── Sheet reference ──
  // Passed to Code.gs so it opens the correct campaign sheet (shard support).
  sheetId: "1RQaw9XHdHXnr9laW0UtPDQxfCpHQVzeXfo6Z3Pz5VPA",
  salesTab:  "_Sales",    // Per-office: _Sales_{officeId}
  rosterTab: "_Roster",   // Per-office: _Roster_{officeId}

  // ── Column mappings ──
  // Maps new _Sales schema headers to internal product keys
  columns: {
    repName:    "Rep Name",
    dateOfSale: "Date of Sale",
    timestamp:  "Timestamp",  // For time-of-sale bucketing

    // Products: each entry defines how to read units from the sheet
    // type: "boolean" → 0/1 in sheet
    // type: "quantity" → parse integer from column
    // type: "sum" → sum multiple columns
    products: [
      { key: "air",   label: "Air",   type: "boolean",  column: "Air" },
      { key: "cell",  label: "Cell",  type: "sum",      columns: ["New Phones", "BYODs"] }
    ],

    // Columns that exist but are EXCLUDED from yeses and units
    excluded: []
  },

  // ── Teams ──
  // Loaded dynamically from _Teams tab. Fallback only if _Teams is empty.
  teams: [],

  // ── Logos ──
  logoUrl: "references/logos/delagroup-logo-full-standard.png",
  logoIconUrl: "references/logos/delagroup-logo-full-standard.png",
  headerLogoStyle: "full",  // 'full' = full logo image, 'icon' = icon + text name

  // ── Team emojis for customization ──
  // Large pool — UI shows a random subset with a reroll button
  teamEmojis: [
    '🦈','🐺','🦁','🔥','⚡','🏆','💎','👑','🎯','🚀','🦅','🐉','💪','🌊','🎪','🏹','⚔️','🛡️','🌟','🎲',
    '🐍','🦂','🐝','🦬','🐘','🦍','🐆','🦊','🐻','🦏','🐊','🦇','🦎','🐧','🦉','🐎','🦄','🐲','🦩','🦚',
    '🔱','💀','🧊','🌪️','🌋','☄️','💣','🧨','🪓','🗡️','🏴‍☠️','🎱','🃏','♠️','♦️','♣️','♥️','🪖','🛸','🧲',
    '🍀','🌵','🌶️','🥊','🏈','🏀','⚽','🎸','🎮','🧬','🔮','🪬','🗿','⛩️','🏔️','🌙','☀️','❄️','🌈','💫'
  ],

  // ── Roles ──
  roles: {
    superadmin: { label: "Super Admin",  rank: 7 },
    owner:      { label: "Owner",        rank: 6 },
    admin:      { label: "Admin",        rank: 5 },
    manager:    { label: "Manager",      rank: 4 },
    jd:         { label: "Jr. Director", rank: 3 },
    l1:         { label: "Team Leader",  rank: 2 },
    rep:        { label: "Client Rep",   rank: 1 }
  },

  // ── Time-of-sale buckets ──
  timeSlots: ['10:30–4:00', '4:00–6:00', '6:00–9:00', '9:00–10:00'],
  timeRanges: [
    { start: 10.5, end: 16 },   // 10:30 AM – 4:00 PM
    { start: 16,   end: 18 },   // 4:00 PM – 6:00 PM
    { start: 18,   end: 21 },   // 6:00 PM – 9:00 PM
    { start: 21,   end: 22 }    // 9:00 PM – 10:00 PM
  ],

  // ── Thresholds ──
  dailyGreenThreshold: 3,   // daily units >= 3 → green
  dailyYellowThreshold: 2,  // daily units 1..2 → yellow, else red (3+ = green)
  weeklyGreenThreshold: 10, // weekly units >= 10 → green
  weeklyYellowThreshold: 9, // weekly units >= 1..9 → yellow, else red (10+ = green)

  // ── Payroll ──
  payrollMode: "commission-split",  // "commission-split" | "flat-rate"

  // ── Refresh ──
  refreshInterval: 5 * 60 * 1000,  // 5 minutes

  // ── Session ──
  sessionDuration: 24 * 60 * 60 * 1000,  // 24 hours

  // ── Admin API (for office switcher) ──
  adminApiUrl: 'https://script.google.com/macros/s/AKfycbz1WJARKP4YZzZjbWyyBjgrAkUOkJWiMHkcJxr4qV3QwRuBfo6YyleBe2MwV_ruRHWo/exec',
  adminApiKey: 'aptel-admin-2026-secret',
};

// ── Roster sheet columns (email-keyed in _Roster tab) ──
const ROSTER_COLUMNS = {
  email:     "Email",
  name:      "Name",
  role:      "Role",
  team:      "Team",
  active:    "Active",
  dateAdded: "Date Added"
};

// Period labels for table headers
const PERIOD_LABELS = ['MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT', 'SUN', 'THIS WK', 'LW', '2WA', '3WA', '4W TOTAL'];

// Indices that are weekly summaries (not individual days)
const WEEK_PERIODS = new Set([7, 8, 9, 10, 11]);

// ═══════════════════════════════════════════════════════
// Multi-Office Config Override
// ═══════════════════════════════════════════════════════
// ?office= param supports two formats:
//   1. BASE64 JSON (from admin portal SSO) — applied immediately
//   2. Plain office ID like "off_002" — resolved at runtime via AdminCode.gs
// If no ?office= param, defaults above are used (off_001).

(function() {
  const params = new URLSearchParams(window.location.search);
  const officeParam = params.get('office');
  if (officeParam) {
    // Try base64 JSON first (admin portal SSO format)
    try {
      const cfg = JSON.parse(atob(officeParam));
      if (cfg.officeId) OFFICE_CONFIG.officeId = cfg.officeId;
      if (cfg.sheetId) OFFICE_CONFIG.sheetId = cfg.sheetId;
      if (cfg.appsScriptUrl) OFFICE_CONFIG.appsScriptUrl = cfg.appsScriptUrl;
      if (cfg.apiKey) OFFICE_CONFIG.apiKey = cfg.apiKey;
      if (cfg.officeName) OFFICE_CONFIG.officeName = cfg.officeName;
      if (cfg.logoUrl) OFFICE_CONFIG.logoUrl = cfg.logoUrl;
      if (cfg.logoIconUrl) OFFICE_CONFIG.logoIconUrl = cfg.logoIconUrl;
      if (cfg.headerLogoStyle) OFFICE_CONFIG.headerLogoStyle = cfg.headerLogoStyle;
      if (cfg.payrollManagerEmail) OFFICE_CONFIG.payrollManagerEmail = cfg.payrollManagerEmail;
      if (cfg.payrollMode) OFFICE_CONFIG.payrollMode = cfg.payrollMode;
      if (cfg.discordWebhookUrl) OFFICE_CONFIG.discordWebhookUrl = cfg.discordWebhookUrl;
      if (cfg.chatPlatform && cfg.chatPlatform !== 'none') OFFICE_CONFIG.chatPlatform = cfg.chatPlatform;
      if (cfg.ownerEmail) OFFICE_CONFIG.ownerEmail = cfg.ownerEmail;
      if (cfg.ownerName) OFFICE_CONFIG.ownerName = cfg.ownerName;
      console.log('[Multi-Office] Config overridden for:', cfg.officeName || 'Unknown office', '| officeId:', cfg.officeId || 'default');
    } catch(e) {
      // Not valid base64 JSON — treat as plain office ID (e.g. "off_002")
      // Will be resolved at runtime in app.js via AdminCode.gs
      OFFICE_CONFIG._pendingOfficeId = officeParam;
      console.log('[Multi-Office] Plain office ID detected:', officeParam, '— will resolve at runtime');
    }
  }

  // ── Admin SSO token ──
  // When opened from admin portal, URL includes ?adminAuth=BASE64
  // Validates source + 5-minute timestamp expiry
  const authParam = params.get('adminAuth');
  if (authParam) {
    try {
      const auth = JSON.parse(atob(authParam));
      if (auth.source === 'admin-portal' && (Date.now() - auth.timestamp) < 5 * 60 * 1000) {
        OFFICE_CONFIG._adminAuth = auth;
        console.log('[SSO] Admin auth token validated for:', auth.email);
      } else {
        console.warn('[SSO] Admin auth token expired or invalid source');
      }
    } catch(e) {
      console.warn('[SSO] Invalid adminAuth param');
    }
  }
})();
