// ═══════════════════════════════════════════════════════
// ELEVATE — Office Configuration
// ═══════════════════════════════════════════════════════
// Each office gets its own config. For multi-office,
// swap this file or load from a config sheet.

const OFFICE_CONFIG = {
  officeName: "Main Office",

  // ── Apps Script Middleware ──
  // All reads and writes go through Apps Script to keep the sheet private.
  appsScriptUrl: "https://script.google.com/macros/s/AKfycbw_kY7RVyydXgLlAA6_CdC-zBbhrv18VAJCJoUgymVTYBGh_uTlBZvbIKp805m7WA20/exec",
  apiKey: "elevate-dash-2026-secret",

  // ── Sheet reference (for documentation — reads go through Apps Script) ──
  sheetId: "1wxM6Htwfy8LrD_o_C7gmvnZEmkfV3FTCVjJU6IITZFc",
  salesTab:  "Order Log",
  rosterTab: "_Roster",

  // ── Column mappings ──
  // Maps Google Sheet column headers to internal product keys
  columns: {
    repName:    "Representative's Name",
    dateOfSale: "Date of Sale",
    timestamp:  "Timestamp",  // For time-of-sale bucketing

    // Products: each entry defines how to read units from the sheet
    // type: "boolean" → yes/no = 1/0
    // type: "quantity" → parse integer from column
    // type: "sum" → sum multiple columns
    products: [
      { key: "air",   label: "Air",   type: "boolean",  column: "Was Internet Air Sold?" },
      { key: "cell",  label: "Cell",  type: "sum",      columns: ["Quantity of New Phones", "Quantity of BYODs"] },
      { key: "fiber", label: "Fiber", type: "boolean",  column: "Was Fiber Optic Internet Sold?" },
      { key: "voip",  label: "VoIP",  type: "quantity", column: "Quantity Sold" }
    ],

    // Columns that exist but are EXCLUDED from yeses and units
    excluded: ["Was DIRECTV Sold?"]
  },

  // ── Teams ──
  teams: ["Aces", "Squids", "Grind Team", "Queenz", "Different Breed", "Sharks", "Dawgs"],

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
    manager:    { label: "Manager",      rank: 5 },
    admin:      { label: "Admin",        rank: 4 },
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
  dailyYellowThreshold: 1,  // daily units >= 1 → yellow, else red
  weeklyGreenThreshold: 10, // weekly units >= 10 → green
  weeklyYellowThreshold: 1, // weekly units >= 1 → yellow, else red

  // ── Refresh ──
  refreshInterval: 5 * 60 * 1000,  // 5 minutes

  // ── Session ──
  sessionDuration: 24 * 60 * 60 * 1000,  // 24 hours
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
