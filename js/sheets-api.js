// ═══════════════════════════════════════════════════════
// ELEVATE — Apps Script API Layer
// All reads and writes go through the deployed Apps Script
// web app, keeping the underlying Google Sheet private.
// ═══════════════════════════════════════════════════════

const SheetsAPI = {

  // ── Fetch all dashboard data via Apps Script doGet ──
  // Returns: { people, roster, teamMap, orderOverrides, teamCustomizations, unlockRequests }
  async fetchAllData(config) {
    const url = `${config.appsScriptUrl}?key=${encodeURIComponent(config.apiKey)}`;
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Apps Script HTTP ${resp.status}`);
    const data = await resp.json();
    if (data.error) throw new Error(data.error);
    return data;
  },

  // ── Write-back via Apps Script doPost ──
  // Actions: addRosterEntry, updateRosterEntry, deleteRosterEntry,
  //          toggleDeactivate, saveOrderOverride, setTeamCustomization,
  //          setUnlockRequest, deleteUnlockRequest
  async post(config, action, payload) {
    if (!this.isConfigured(config)) {
      console.warn('Apps Script URL not configured — write skipped');
      return { ok: false, error: 'not configured' };
    }
    try {
      const resp = await fetch(config.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' }, // Avoids CORS preflight
        body: JSON.stringify({ key: config.apiKey, action, ...payload })
      });
      const result = await resp.json().catch(() => ({ ok: true }));
      return { ok: true, data: result };
    } catch (err) {
      console.error('Write-back failed:', err);
      return { ok: false, error: err.message };
    }
  },

  // ── Fetch individual order rows (past 30 days) ──
  async fetchOrders(config, filterEmail) {
    let url = `${config.appsScriptUrl}?key=${encodeURIComponent(config.apiKey)}&action=readOrders`;
    if (filterEmail) {
      url += `&email=${encodeURIComponent(filterEmail)}`;
    }
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Apps Script HTTP ${resp.status}`);
    const data = await resp.json();
    if (data.error) throw new Error(data.error);
    return data.orders || [];
  },

  // ── Fetch payroll orders (trainee=Yes, past 2 months) ──
  async fetchPayrollOrders(config) {
    const url = `${config.appsScriptUrl}?key=${encodeURIComponent(config.apiKey)}&action=readPayrollOrders`;
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Apps Script HTTP ${resp.status}`);
    const data = await resp.json();
    if (data.error) throw new Error(data.error);
    return data.orders || [];
  },

  // ── Check if Apps Script URL is configured ──
  isConfigured(config) {
    return config.appsScriptUrl &&
           config.appsScriptUrl !== 'YOUR_APPS_SCRIPT_URL';
  }
};
