// ═══════════════════════════════════════════════════════
// Aptel Slack Channel Auditor — Configuration
// ═══════════════════════════════════════════════════════

const SLACK_CONFIG = {
  // Cloudflare Worker proxy URL
  workerUrl: 'https://aptel-slack-proxy.aprindle.workers.dev',

  // Column keys (match Google Sheet headers)
  columns: {
    name: 'Name',
    email: 'Email',
    slackEmail: 'SlackEmail',
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
