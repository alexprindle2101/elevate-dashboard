// ═══════════════════════════════════════════════════════
// ELEVATE — Apps Script API Layer
// All reads and writes go through the deployed Apps Script
// web app, keeping the underlying Google Sheet private.
//
// Multi-office: Every request includes officeId + sheetId
// so Code.gs can open the correct campaign sheet and
// read/write the correct per-office tabs.
// ═══════════════════════════════════════════════════════

const SheetsAPI = {

  // ── Build GET query string with office params ──
  _buildUrl(config, extraParams) {
    let url = `${config.appsScriptUrl}?key=${encodeURIComponent(config.apiKey)}`;
    url += `&officeId=${encodeURIComponent(config.officeId || '')}`;
    if (config.sheetId) {
      url += `&sheetId=${encodeURIComponent(config.sheetId)}`;
    }
    if (extraParams) {
      for (const [k, v] of Object.entries(extraParams)) {
        if (v !== undefined && v !== null && v !== '') {
          url += `&${k}=${encodeURIComponent(v)}`;
        }
      }
    }
    return url;
  },

  // ── Fetch all dashboard data via Apps Script doGet ──
  // Returns: { people, roster, teamMap, orderOverrides, teamCustomizations, unlockRequests }
  async fetchAllData(config) {
    const params = {};
    if (OFFICE_CONFIG.ownerEmail) params.ownerEmail = OFFICE_CONFIG.ownerEmail;
    if (OFFICE_CONFIG.ownerName) params.ownerName = OFFICE_CONFIG.ownerName;
    const url = this._buildUrl(config, params);
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Apps Script HTTP ${resp.status}`);
    const data = await resp.json();
    if (data.error) throw new Error(data.error);
    return data;
  },

  // ── Write-back via Apps Script doPost ──
  // Actions: addRosterEntry, updateRosterEntry, deleteRosterEntry,
  //          toggleDeactivate, saveOrderOverride, setTeamCustomization,
  //          setUnlockRequest, deleteUnlockRequest, addSale, etc.
  async post(config, action, payload) {
    if (!this.isConfigured(config)) {
      console.warn('Apps Script URL not configured — write skipped');
      return { ok: false, error: 'not configured' };
    }
    try {
      const resp = await fetch(config.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' }, // Avoids CORS preflight
        body: JSON.stringify({
          key: config.apiKey,
          action,
          officeId: config.officeId || '',
          sheetId: config.sheetId || '',
          ...payload
        })
      });
      const result = await resp.json().catch(() => ({}));
      const serverOk = result.ok === true && !result.error;
      if (!serverOk) {
        console.error('[SheetsAPI.post]', action, 'server responded:', JSON.stringify(result));
      } else {
        console.log('[SheetsAPI.post]', action, 'success:', JSON.stringify(result));
      }
      return { ok: serverOk, data: result };
    } catch (err) {
      console.error('[SheetsAPI.post]', action, 'network error:', err.message);
      return { ok: false, error: err.message };
    }
  },

  // ── Fetch individual order rows (past 30 days) ──
  async fetchOrders(config, filterEmail) {
    const params = { action: 'readOrders' };
    if (filterEmail) params.email = filterEmail;
    const url = this._buildUrl(config, params);
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Apps Script HTTP ${resp.status}`);
    const data = await resp.json();
    if (data.error) throw new Error(data.error);
    return data.orders || [];
  },

  // ── Fetch payroll orders (filtered by payrollMode) ──
  async fetchPayrollOrders(config) {
    const url = this._buildUrl(config, { action: 'readPayrollOrders', payrollMode: OFFICE_CONFIG.payrollMode || 'commission-split' });
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Apps Script HTTP ${resp.status}`);
    const data = await resp.json();
    if (data.error) throw new Error(data.error);
    return data.orders || [];
  },

  // ── Fetch Tableau summary (on-demand refresh) ──
  async fetchTableauSummary(config) {
    const url = this._buildUrl(config, { action: 'readTableauSummary' });
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Apps Script HTTP ${resp.status}`);
    const data = await resp.json();
    if (data.error) throw new Error(data.error);
    return data;
  },

  // ── Challenge ──
  async fetchChallengeConfig(config) {
    const url = this._buildUrl(config, { action: 'readChallengeConfig' });
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Apps Script HTTP ${resp.status}`);
    const data = await resp.json();
    if (data.error) throw new Error(data.error);
    return data.config || null;
  },

  async fetchChallengeSales(config, startDate, endDate) {
    const url = this._buildUrl(config, { action: 'readChallengeSales', startDate, endDate });
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Apps Script HTTP ${resp.status}`);
    const data = await resp.json();
    if (data.error) throw new Error(data.error);
    return data.sales || {};
  },

  async fetchChallengeBlood(config) {
    const url = this._buildUrl(config, { action: 'readChallengeBlood' });
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Apps Script HTTP ${resp.status}`);
    const data = await resp.json();
    if (data.error) throw new Error(data.error);
    return data.blood || {};
  },

  // ── Fetch Tableau device detail for a single DSI ──
  async fetchTableauDetail(config, dsi) {
    const url = this._buildUrl(config, { action: 'readTableauDetail', dsi });
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Apps Script HTTP ${resp.status}`);
    const data = await resp.json();
    if (data.error) throw new Error(data.error);
    return data.devices || [];
  },

  // ── Universal login: find which office a rep belongs to ──
  // Calls AdminCode.gs (admin API), not the per-office Code.gs
  async findRepOffice(email) {
    const url = `${OFFICE_CONFIG.adminApiUrl}?key=${encodeURIComponent(OFFICE_CONFIG.adminApiKey)}&action=findRepOffice&email=${encodeURIComponent(email)}`;
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Admin API HTTP ${resp.status}`);
    return await resp.json();
  },

  // ── Look up office config by plain office ID ──
  // Calls AdminCode.gs getOfficeConfig endpoint
  async getOfficeConfig(officeId) {
    const url = `${OFFICE_CONFIG.adminApiUrl}?key=${encodeURIComponent(OFFICE_CONFIG.adminApiKey)}&action=getOfficeConfig&officeId=${encodeURIComponent(officeId)}`;
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`Admin API HTTP ${resp.status}`);
    return await resp.json();
  },

  // ── Check if Apps Script URL is configured ──
  isConfigured(config) {
    return config.appsScriptUrl &&
           config.appsScriptUrl !== 'YOUR_APPS_SCRIPT_URL';
  }
};
