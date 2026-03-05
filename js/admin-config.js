// ═══════════════════════════════════════════════════════
// Aptel Admin Dashboard — Configuration
// ═══════════════════════════════════════════════════════
// Update these values after creating the Admin Google Sheet
// and deploying AdminCode.gs as a web app.

const ADMIN_CONFIG = {
  // Google Sheet ID for the admin master sheet
  sheetId: '',  // TODO: paste your admin sheet ID here

  // Deployed AdminCode.gs web app URL
  appsScriptUrl: '',  // TODO: paste your deployed URL here

  // API key (must match Script Properties > API_KEY in AdminCode.gs)
  apiKey: 'aptel-admin-2026-secret',

  // Session config
  sessionKey: 'aptel_admin_session',
  sessionDuration: 24 * 60 * 60 * 1000,  // 24 hours

  // Template types available for offices
  templates: {
    'att-b2b': {
      label: 'AT&T B2B',
      file: 'index.html',
      description: 'AT&T Business-to-Business sales dashboard'
    }
    // Future: 'residential': { label: 'Residential', file: 'residential.html', ... }
  }
};
