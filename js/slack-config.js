// ═══════════════════════════════════════════════════════
// Aptel Slack Channel Auditor — Configuration
// ═══════════════════════════════════════════════════════

const SLACK_CONFIG = {
  // Cloudflare Worker proxy URL
  workerUrl: 'https://aptel-slack-proxy.aprindle.workers.dev',

  // localStorage keys
  excelStorageKey: 'aptel_slack_excel_data',

  // Excel sheet names
  expectedSheets: {
    people: 'People',
    departments: 'Departments',
    roles: 'Roles',
  },

  // Column header mappings — People sheet
  peopleColumns: {
    name: 'Name',
    email: 'Email',
    slackEmail: 'SlackEmail',
    department: 'Department',   // comma-separated if multiple
    level: 'Level',             // SWAT, Manager, Lead, Member
  },

  // Column header mappings — Departments sheet (base channels)
  deptColumns: {
    department: 'Department',
    channel: 'Channel',
  },

  // Column header mappings — Roles sheet (department + level specific channels)
  roleColumns: {
    department: 'Department',
    level: 'Level',
    channel: 'Channel',
  },

  // Status labels
  statusLabels: {
    match: 'OK',
    missing: 'Missing',
    extra: 'Extra',
    notFound: 'Not in Slack',
    noMapping: 'No Mapping',
  },

  // Status badge CSS classes
  statusColors: {
    match: 'ok',
    missing: 'missing',
    extra: 'extra',
    notFound: 'not-found',
    noMapping: 'no-role',
  },
};
