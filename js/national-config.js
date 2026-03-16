// ═══════════════════════════════════════════════════════
// National Consultant Dashboard — Configuration
// ═══════════════════════════════════════════════════════
// Data sources for the NC dashboard (Ken's one-on-one review tool).
// NationalCode.gs reads from these external sheets and returns unified JSON.

const NATIONAL_CONFIG = {
  // ── NationalCode.gs Deployment ──
  // Create a new Google Sheet, add NationalCode.gs, deploy as web app.
  appsScriptUrl: 'https://script.google.com/macros/s/AKfycbzX68JL0htLcwtyQYN7dttltwcEyHNQW3lcPeWzkS4PqKycCnvzDD6RQFbFvCvUcPfH/exec',
  apiKey: 'national-dash-2026-secret',

  // ── Session ──
  sessionKey: 'national_session',
  sessionDuration: 24 * 60 * 60 * 1000, // 24 hours

  // ── Data Source Sheets ──
  // These are external sheets shared with the NC. NationalCode.gs reads them via openById().
  sheets: {
    // Sheet 1: Maddy's weekly recruiting stats (one tab per week, named by date)
    recruitingWeekly: {
      id: '1MNLqi8A329444SeZpKbYbcRe3dMxaOPLVdMy-7F1DPk',
      label: 'All Campaigns Stats Tracker 2026'
    },
    // Sheet 1 Alt: Daily recruiting scoreboard
    recruitingDaily: {
      id: '1ytTGen_AlzfDPW3HGYU1JKNLz1kfHrrhAFCVnmRS3fg',
      label: 'Recruiting Scoreboard: All Programs Daily'
    },
    // Sheet 3: ATT Campaign Tracker — per-owner tabs with 3 sections each
    campaignTracker: {
      id: '1HvWJYox3JXvxmza63YBWAqKPtUGPFuaV-s-BOfbWGKM',
      label: 'ATT Campaign Tracker'
    },
    // Sheet 4: Performance Audit — online presence, reviews, social media
    performanceAudit: {
      id: '15WCMzKnqvyyRMx2ae4tC1a12_-aoSRDh3McOuRAKuHk',
      label: 'Performance Audit'
    },
    // Sheet 5: Ken's national recruiting data (copy of Maddy's weekly stats)
    // One tab per week, Column A has owners grouped under campaign headers.
    national: {
      id: '1eGkwjQRD9RV4n-JR_TTlgE6VY858WZID8cAF8soYSYM',
      label: 'National Recruiting Data'
    },
    // Sheet 7: AT&T NDS/Verizon Wireless One-on-Ones (per-owner tabs with headcount + recruiting + sales)
    ndsOneOnOnes: {
      id: '1kcUWR3EKgP-9wDct4vDyuQJ7IuS0cbcetY97dmVTY64',
      label: 'AT&T NDS One-on-Ones'
    },
    // Folder 6: Indeed ad cost spreadsheets (one per owner, named by owner name)
    indeedCostsFolder: {
      id: '1r2lGOOjXQkvzz1we5k1Gn5drXrZKc42y',
      label: 'Indeed Ad Spend'
    }
  },

  // ── Campaigns ──
  // Each campaign maps to specific data in the source sheets.
  // V1: AT&T B2B only. Future: Frontier, Ooma, etc.
  campaigns: {
    'att-b2b': {
      label: 'AT&T B2B',
      // Tab name pattern in Campaign Tracker (owner tabs are just owner names)
      campaignTotalsTab: 'Campaign Totals',
      // Section header text used to find AT&T data in the weekly recruiting sheet
      sectionHeader: 'AT&T Campaign Totals',
      // Number of weekly tabs to pull for trend data
      weeksToPull: 6
    }
  },

  // ── Owner List (AT&T B2B) ──
  // Maps owner display names to their tab names in the Campaign Tracker.
  // Also maps to business names in the Performance Audit sheet.
  // TODO: This could eventually be pulled dynamically from the _Owners tab in admin sheet.
  owners: {
    'att-b2b': [
      { name: 'Jay T',             tab: 'Jay T',             businesses: [] },
      { name: 'Mason',             tab: 'Mason',             businesses: [] },
      { name: 'Steven Sykes',      tab: 'Steven Sykes',      businesses: [] },
      { name: 'Olin Salter',       tab: 'Olin Salter',       businesses: [] },
      { name: 'Eric Martinez',     tab: 'Eric Martinez',     businesses: [] },
      { name: 'Natalia Gwarda',    tab: 'Natalia Gwarda',    businesses: [] },
      { name: 'Nigel Gilbert',     tab: 'Nigel Gil',         businesses: [] }
    ]
  },

  // ── Owner Name Aliases ──
  // Maps variations/nicknames from Cam's Performance Audit sheet (Client Name column)
  // to canonical owner names used in the recruiting/campaign data.
  // Key = lowercase alias, Value = canonical owner name from owners list.
  // Add entries here as mismatches are discovered.
  ownerAliases: {
    // Examples — update these once you see the actual Client Name values in Cam's sheet:
    // 'jay':            'Jay T',
    // 'jay thurston':   'Jay T',
    // 'steve':          'Steven Sykes',
    // 'steve sykes':    'Steven Sykes',
    // 'nat':            'Natalia Gwarda',
    // 'nigel':          'Nigel Gilbert',
    // 'nigel gil':      'Nigel Gilbert',
  },

  // ── Login ──
  // NC dashboard uses admin portal SSO (o4 role) or direct PIN login.
  // For now, simple alias-based access like admin portal.
  loginAliases: {
    'ken': 'ken@example.com'   // TODO: set real email
  },

  // ── Recruiting Funnel Columns ──
  // Header text patterns to find columns by name (not index) in messy sheets.
  // Used by NationalCode.gs for header-based lookup.
  recruitingHeaders: {
    name:               ['name', 'at&t campaign totals'],
    firstRoundsBooked:  ['1st rounds booked', '1st round booked'],
    firstRoundsShowed:  ['1st rounds showed', '1st round showed'],
    turnedTo2nd:        ['turned to 2nd'],
    retention1:         ['retention'],
    conversion:         ['conversion'],
    secondRoundsBooked: ['2nd rounds booked', '2nd round booked'],
    secondRoundsShowed: ['2nd rounds showed', '2nd round showed'],
    retention2:         ['retention'],
    newStartScheduled:  ['new start scheduled', 'new starts scheduled'],
    newStartsShowed:    ['new starts showed'],
    retention3:         ['retention'],
    activeHeadcount:    ['active selling headcount', 'active headcount']
  },

  // ── Office Health Headers (Campaign Tracker Section 1) ──
  officeHealthHeaders: {
    dates:           ['dates', 'date'],
    active:          ['active'],
    leaders:         ['leaders'],
    dist:            ['dist'],
    training:        ['training'],
    productionLW:    ['production lw', 'production'],
    dtv:             ['dtv'],
    productionGoals: ['production goals']
  },

  // ── Tableau/Sales Headers (Campaign Tracker Section 3) ──
  salesHeaders: {
    name:             ['name'],
    newInternet:      ['new internet count'],
    upgradeInternet:  ['upgrade internet count'],
    videoSales:       ['video sales'],
    salesAll:         ['sales (all)', 'sales all'],
    abpMix:           ['new internet abp mix', 'abp mix'],
    gigMix:           ['new internet 1gig', '1gig+ mix'],
    techInstall:      ['tech install']
  }
};
