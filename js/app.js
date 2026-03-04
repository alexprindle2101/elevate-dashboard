// ═══════════════════════════════════════════════════════
// ELEVATE — Main Application
// Init, navigation, role switching, event wiring
// ═══════════════════════════════════════════════════════

const App = {
  state: {
    people: [],
    teams: [],
    teamsData: {},       // raw _Teams sheet data (for rebuilding hierarchy)
    roster: {},          // email-keyed roster from API
    settings: {},        // key-value settings from _Settings tab
    currentRole: 'rep',
    realRole: 'rep',       // actual logged-in role (stays superadmin during View-As)
    viewAsActive: false,   // whether superadmin View-As switcher is expanded
    currentPersona: '',
    currentEmail: '',
    currentNav: 'leaderboard',
    lastUpdated: null,
    refreshTimer: null
  },

  // Cached API data (used between fetch and login)
  _cachedApiData: null,

  // ══════════════════════════════════════════════
  // INIT
  // ══════════════════════════════════════════════
  async init() {
    await this.initProduction();
  },

  // ── Production Init ──
  async initProduction() {
    Auth.showLoading('Checking session...');

    // Check existing session
    const session = Auth.getSession();
    if (session) {
      this.state.currentRole = session.role;
      this.state.realRole = session.role;
      this.state.currentPersona = session.name;
      this.state.currentEmail = session.email || '';
      await this.loadData();
      // Show role-switcher for superadmin in production (collapsed — views as admin)
      if (this.state.realRole === 'superadmin') {
        const switcher = document.getElementById('role-switcher');
        if (switcher) switcher.style.display = 'flex';
        this._applySuperAdminViewAs();
      }
    } else {
      // Need login — fetch all data (includes roster for email validation)
      Auth.showLoading('Connecting...');
      try {
        const apiData = await SheetsAPI.fetchAllData(OFFICE_CONFIG);
        this.state.roster = apiData.roster || {};
        this._cachedApiData = apiData;
        console.info('Roster loaded:', Object.keys(this.state.roster).length, 'entries');
        Auth.hideLoading();
        Auth.showLoginScreen((email) => this.handleEmailStep(email));
      } catch (err) {
        Auth.hideLoading();
        this.showError('Unable to connect. Check your Apps Script deployment.', err);
      }
    }
  },

  // Temp login state (between email and PIN steps)
  _loginEmail: null,
  _loginRosterEntry: null,

  // ── Step 1: Email submitted ──
  handleEmailStep(email) {
    const result = Auth.checkEmail(email, this.state.roster);
    if (!result.ok) {
      Auth.showLoginError(result.error);
      return;
    }

    this._loginEmail = result.email;
    this._loginRosterEntry = result.rosterEntry;

    const onBack = () => {
      Auth.showLoginError('');
      Auth._showEmailStep((e) => this.handleEmailStep(e));
    };

    if (result.hasPin) {
      Auth.showPinStep(result.email, (pin) => this.handlePinValidation(pin), onBack);
    } else {
      Auth.showPinCreateStep(result.email, (pin, confirm) => this.handlePinCreation(pin, confirm), onBack);
    }
  },

  // ── Step 2a: Validate existing PIN ──
  async handlePinValidation(pin) {
    if (!pin) {
      Auth.showLoginError('Please enter your PIN');
      return;
    }

    const btn = document.getElementById('login-btn');
    if (btn) { btn.textContent = 'VERIFYING...'; btn.disabled = true; }

    const result = await Auth.validatePin(this._loginEmail, pin, OFFICE_CONFIG);

    if (btn) { btn.textContent = 'SIGN IN'; btn.disabled = false; }

    if (result.ok) {
      this._completeLogin();
    } else {
      Auth.showLoginError(result.error || 'Incorrect PIN');
    }
  },

  // ── Step 2b: Create new PIN (first login) ──
  async handlePinCreation(pin, confirmPin) {
    if (!pin || !confirmPin) {
      Auth.showLoginError('Please fill in both PIN fields');
      return;
    }
    if (!/^\d{4,6}$/.test(pin)) {
      Auth.showLoginError('PIN must be 4-6 digits');
      return;
    }
    if (pin !== confirmPin) {
      Auth.showLoginError('PINs do not match');
      return;
    }

    const btn = document.getElementById('login-btn');
    if (btn) { btn.textContent = 'SETTING UP...'; btn.disabled = true; }

    const result = await Auth.createPin(this._loginEmail, pin, OFFICE_CONFIG);

    if (btn) { btn.textContent = 'SET PIN & SIGN IN'; btn.disabled = false; }

    if (result.ok) {
      if (this.state.roster[this._loginEmail]) {
        this.state.roster[this._loginEmail].hasPin = true;
      }
      this._completeLogin();
    } else {
      Auth.showLoginError(result.error || 'Failed to set PIN');
    }
  },

  // ── Complete login after PIN validation/creation ──
  _completeLogin() {
    const session = Auth.createSession(this._loginEmail, this._loginRosterEntry);
    this.state.currentRole = session.role;
    this.state.realRole = session.role;
    this.state.currentPersona = session.name;
    this.state.currentEmail = session.email;
    Auth.hideLoginScreen();

    // Show role-switcher for superadmin (collapsed — views as admin)
    if (this.state.realRole === 'superadmin') {
      const switcher = document.getElementById('role-switcher');
      if (switcher) switcher.style.display = 'flex';
      this._applySuperAdminViewAs();
    }

    if (this._cachedApiData) {
      Auth.showLoading('Loading dashboard...');
      this._processApiData(this._cachedApiData);
      this._cachedApiData = null;
      Auth.hideLoading();
      this.startAutoRefresh();
    } else {
      this.loadData();
    }

    this._loginEmail = null;
    this._loginRosterEntry = null;
  },

  // ── Load Data from Apps Script (with retry) ──
  async loadData() {
    Auth.showLoading('Loading sales data...');
    const MAX_RETRIES = 2;
    let lastErr;
    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
      try {
        if (attempt > 0) Auth.showLoading(`Retrying (${attempt}/${MAX_RETRIES})...`);
        const apiData = await SheetsAPI.fetchAllData(OFFICE_CONFIG);
        this._processApiData(apiData);
        Auth.hideLoading();
        this.startAutoRefresh();
        return;
      } catch (err) {
        lastErr = err;
        console.warn(`Fetch attempt ${attempt + 1} failed:`, err.message);
        if (attempt < MAX_RETRIES) await new Promise(r => setTimeout(r, 2000));
      }
    }
    Auth.hideLoading();
    this.showError('Failed to load data from Google Sheets.', lastErr);
  },

  // ── Process API data into app state ──
  _processApiData(apiData) {
    if (apiData._debug) console.info('🔍 Server debug:', JSON.stringify(apiData._debug, null, 2));
    const result = DataPipeline.buildFromAppsScript(apiData, OFFICE_CONFIG);
    this.state.people = result.people;
    this.state.teams = result.teams;
    this.state.roster = apiData.roster || {};

    this.state.teamsData = apiData.teams || {};
    this.state.settings = apiData.settings || {};
    Roster.init(this.state.roster);
    Roster.initFromApi(apiData);
    TeamsManager.init(this.state.teamsData);
    this.updateNav();
    Render.renderAll(this.state.people, this.state.teams);
    // Re-render roster page if it's currently visible
    if (this.state.currentNav === 'roster') {
      Roster.renderRoster(this.state.people, this.state.currentRole, OFFICE_CONFIG);
    }
    // Re-render teams page if it's currently visible
    if (this.state.currentNav === 'teams') {
      TeamsManager.renderTeamsPage(this.state.people, this.state.currentRole, OFFICE_CONFIG);
    }
    // Re-fetch orders if orders page is visible
    if (this.state.currentNav === 'allOrders') this._loadAndRenderOrders('all');
    if (this.state.currentNav === 'myOrders') this._loadAndRenderOrders('my');
    if (this.state.currentNav === 'payroll') this._loadAndRenderPayroll();
    if (this.state.currentNav === 'office') this._renderOfficePage();
    // Ensure role-switcher stays visible for superadmin after refresh
    if (this.state.realRole === 'superadmin') {
      const switcher = document.getElementById('role-switcher');
      if (switcher) switcher.style.display = 'flex';
      // Re-apply collapsed/expanded state without resetting it
      const controls = document.getElementById('view-as-controls');
      const personaWrap = document.getElementById('role-persona-wrap');
      const toggleBtn = document.getElementById('view-as-toggle');
      if (toggleBtn) toggleBtn.style.display = '';
      if (!this.state.viewAsActive) {
        if (controls) controls.style.display = 'none';
        if (personaWrap) personaWrap.style.display = 'none';
      }
    }
    this.updateLastUpdated();
  },

  // ── Auto-Refresh ──
  startAutoRefresh() {
    if (this.state.refreshTimer) clearInterval(this.state.refreshTimer);
    this.state.refreshTimer = setInterval(() => {
      this.refreshData();
    }, OFFICE_CONFIG.refreshInterval);
  },

  _refreshing: false,

  async refreshData() {
    if (this._refreshing) return;
    this._refreshing = true;
    this._setRosterLoading(true);
    try {
      const apiData = await SheetsAPI.fetchAllData(OFFICE_CONFIG);
      this._processApiData(apiData);
    } catch (err) {
      console.error('Auto-refresh failed:', err);
    } finally {
      this._refreshing = false;
      this._setRosterLoading(false);
    }
  },

  async refreshRoster() {
    const btn = document.getElementById('roster-refresh-btn');
    if (btn) { btn.textContent = '⟳ Refreshing...'; btn.disabled = true; }
    await this.refreshData();
    // Re-render roster page with updated data
    Roster.renderRoster(this.state.people, this.state.currentRole, OFFICE_CONFIG);
    if (btn) { btn.innerHTML = '&#x21bb; Refresh'; btn.disabled = false; }
  },

  _setRosterLoading(loading) {
    const addBtn = document.getElementById('add-member-btn');
    if (addBtn) {
      addBtn.disabled = loading;
      addBtn.style.opacity = loading ? '0.5' : '';
      addBtn.style.pointerEvents = loading ? 'none' : '';
    }
  },

  updateLastUpdated() {
    this.state.lastUpdated = new Date();
    const el = document.getElementById('last-updated-text');
    if (el) {
      el.textContent = `Updated ${this.state.lastUpdated.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' })}`;
    }
  },

  // ══════════════════════════════════════════════
  // NAVIGATION
  // ══════════════════════════════════════════════
  updateNav() {
    const role = this.state.currentRole;
    const showProfile = true;
    const isSuperAdmin = (role === 'superadmin');
    const showTeam = isSuperAdmin || (role === 'rep' || role === 'l1' || role === 'jd' || role === 'manager');
    const showEdit = isSuperAdmin || (role === 'jd');
    const showRoster = isSuperAdmin || (role === 'owner' || role === 'manager' || role === 'admin' || role === 'jd');
    const showTeams = isSuperAdmin || (role === 'owner' || role === 'manager' || role === 'admin');
    const showAllOrders = isSuperAdmin || (role === 'owner' || role === 'manager' || role === 'admin');
    const showMyOrders = isSuperAdmin || (role === 'rep' || role === 'l1' || role === 'jd' || role === 'manager');
    const showOffice = isSuperAdmin || (role === 'owner');
    const curEmail = (this.state.currentEmail || '').toLowerCase();
    const payrollMgr = (this.state.settings.payrollManager || '').toLowerCase();
    const showPayroll = isSuperAdmin || (role === 'owner') || (payrollMgr && curEmail === payrollMgr);

    const setDisplay = (id, show) => {
      const el = document.getElementById(id);
      if (el) el.style.display = show ? '' : 'none';
    };

    setDisplay('nav-profile', showProfile);
    setDisplay('nav-team', showTeam);
    setDisplay('nav-team-edit', showEdit);
    setDisplay('nav-roster', showRoster);
    setDisplay('nav-teams', showTeams);
    setDisplay('nav-all-orders', showAllOrders);
    setDisplay('nav-my-orders', showMyOrders);
    setDisplay('nav-office', showOffice);
    setDisplay('nav-payroll', showPayroll);

    // Update team label
    if (showTeam) {
      const myTeam = Roster.getEffectiveTeam(this.state.currentPersona, this.state.people);
      const label = document.getElementById('nav-team-label');
      if (label && myTeam) {
        const d = Roster.getTeamDisplay(myTeam, this.state.people, this.state.teams);
        label.textContent = d.emoji + ' ' + d.name;
      }
    }

    // Unlock badge
    if (showEdit) {
      const badge = document.getElementById('edit-unlock-badge');
      if (badge) badge.style.display = Roster.getUnlockStatus(this.state.currentPersona) === 'pending' ? 'block' : 'none';
    }

    // Roster notification badge
    if (showRoster) {
      const pending = Roster.getPendingRequests().length;
      const badge = document.getElementById('roster-notif-badge');
      if (badge) {
        badge.style.display = pending > 0 ? 'flex' : 'none';
        badge.textContent = pending;
      }
    }

    // Active tab highlight
    const navIdMap = { allOrders: 'nav-all-orders', myOrders: 'nav-my-orders' };
    document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
    const activeTab = document.getElementById(navIdMap[this.state.currentNav] || ('nav-' + this.state.currentNav));
    if (activeTab) activeTab.classList.add('active');
  },

  navTo(tab) {
    // Hide overlay pages when navigating away
    const overlays = { roster: 'roster-page', teams: 'teams-page', allOrders: 'all-orders-page', myOrders: 'my-orders-page', office: 'office-page', payroll: 'payroll-page' };
    Object.entries(overlays).forEach(([key, id]) => {
      if (tab !== key) { const el = document.getElementById(id); if (el) el.style.display = 'none'; }
    });

    this.state.currentNav = tab;
    this.updateNav();

    if (tab === 'leaderboard') { Render.closeProfile(); return; }
    if (tab === 'profile') { this.openPersonProfile(this.state.currentPersona); return; }
    if (tab === 'team') {
      const myTeam = Roster.getEffectiveTeam(this.state.currentPersona, this.state.people);
      if (myTeam) this.openTeamProfile(myTeam);
      return;
    }
    if (tab === 'roster') {
      Render.closeProfile();
      Roster.renderRoster(this.state.people, this.state.currentRole, OFFICE_CONFIG);
      return;
    }
    if (tab === 'teams') {
      Render.closeProfile();
      TeamsManager.renderTeamsPage(this.state.people, this.state.currentRole, OFFICE_CONFIG);
      return;
    }
    if (tab === 'allOrders') {
      Render.closeProfile();
      this._loadAndRenderOrders('all');
      return;
    }
    if (tab === 'myOrders') {
      Render.closeProfile();
      this._loadAndRenderOrders('my');
      return;
    }
    if (tab === 'office') {
      Render.closeProfile();
      this._renderOfficePage();
      return;
    }
    if (tab === 'payroll') {
      Render.closeProfile();
      this._loadAndRenderPayroll();
      return;
    }
  },

  async _loadAndRenderOrders(mode) {
    const pageId = (mode === 'all') ? 'all-orders-page' : 'my-orders-page';
    const page = document.getElementById(pageId);
    if (page) page.style.display = 'block';

    const subtitle = document.getElementById(pageId + '-subtitle');
    if (subtitle) subtitle.textContent = 'Loading orders...';

    await Orders.fetchOrders(OFFICE_CONFIG, mode);
    Orders.renderOrdersPage(mode, OFFICE_CONFIG);
  },

  // ══════════════════════════════════════════════
  // OFFICE PAGE
  // ══════════════════════════════════════════════
  _renderOfficePage() {
    const page = document.getElementById('office-page');
    if (page) page.style.display = 'block';

    // Populate payroll manager dropdown with all admins
    const sel = document.getElementById('office-payroll-select');
    if (!sel) return;

    const admins = [];
    Object.entries(this.state.roster).forEach(([email, r]) => {
      if (r.rank === 'admin' && !r.deactivated) {
        admins.push({ email: email.toLowerCase(), name: r.name || email });
      }
    });
    admins.sort((a, b) => a.name.localeCompare(b.name));

    sel.innerHTML = '<option value="">— None —</option>'
      + admins.map(a => `<option value="${a.email}">${a.name}</option>`).join('');

    // Set current selection
    const current = (this.state.settings.payrollManager || '').toLowerCase();
    if (current) sel.value = current;

    // Hide saved indicator
    const saved = document.getElementById('office-payroll-saved');
    if (saved) saved.style.display = 'none';
  },

  async _setPayrollManager(email) {
    this.state.settings.payrollManager = email;
    this.updateNav();

    // Show saved indicator briefly
    const saved = document.getElementById('office-payroll-saved');
    if (saved) {
      saved.style.display = 'block';
      setTimeout(() => { saved.style.display = 'none'; }, 2000);
    }

    await SheetsAPI.post(OFFICE_CONFIG, 'setSetting', { key: 'payrollManager', value: email });
  },

  // ══════════════════════════════════════════════
  // PAYROLL PAGE
  // ══════════════════════════════════════════════
  _payrollOrders: [],

  async _loadAndRenderPayroll() {
    const page = document.getElementById('payroll-page');
    if (page) page.style.display = 'block';

    const subtitle = document.getElementById('payroll-page-subtitle');
    if (subtitle) subtitle.textContent = 'Loading payroll orders...';

    try {
      this._payrollOrders = await SheetsAPI.fetchPayrollOrders(OFFICE_CONFIG);
    } catch (err) {
      console.error('Failed to fetch payroll orders:', err);
      this._payrollOrders = [];
    }

    // Sync into Orders module so note modal can find orders by rowIndex
    Orders._orders = this._payrollOrders;
    Orders._mode = 'payroll';

    if (subtitle) subtitle.textContent = `${this._payrollOrders.length} trainee orders · Past 2 months`;
    this._filterPayrollOrders();
  },

  _filterPayrollOrders() {
    const search = (document.getElementById('payroll-search')?.value || '').toLowerCase().trim();
    let filtered = this._payrollOrders;
    if (search) {
      filtered = filtered.filter(o => {
        const haystack = (o.repName + ' ' + o.dsi).toLowerCase();
        return haystack.includes(search);
      });
    }

    const countEl = document.getElementById('payroll-count');
    if (countEl) {
      countEl.textContent = filtered.length === this._payrollOrders.length
        ? `Showing all ${this._payrollOrders.length}`
        : `Showing ${filtered.length} of ${this._payrollOrders.length}`;
    }

    this._renderPayrollRows(filtered);
  },

  _renderPayrollRows(orders) {
    const tbody = document.getElementById('payroll-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (orders.length === 0) {
      tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;color:var(--silver-dim);padding:32px;font-family:\'Barlow Condensed\',sans-serif;font-size:14px">No trainee orders found</td></tr>';
      return;
    }

    orders.forEach(o => {
      // Build product list
      const soldParts = [];
      OFFICE_CONFIG.columns.products.forEach(prod => {
        const val = o[prod.key] || 0;
        if (val > 0) {
          soldParts.push(prod.type === 'boolean' ? prod.label : `${prod.label} x${val}`);
        }
      });
      const soldStr = soldParts.length > 0 ? soldParts.join(', ') : '—';

      // Status badge
      const statusColors = { Active: '#22c55e', Pending: '#f0b429', Cancelled: '#e53535', Complete: '#0099cc' };
      const sColor = statusColors[o.status] || 'var(--silver-dim)';

      // Notes preview
      const noteLines = o.notes ? o.notes.split('\n').filter(l => l.trim()) : [];
      const notePreview = noteLines.length > 0
        ? `<span style="font-size:11px;color:var(--silver-dim)">${noteLines[noteLines.length - 1].substring(0, 40)}${noteLines[noteLines.length - 1].length > 40 ? '...' : ''}</span>
           ${noteLines.length > 1 ? `<span style="background:rgba(26,92,229,0.15);color:var(--blue-core);font-size:9px;font-weight:700;border-radius:4px;padding:1px 5px;margin-left:4px">${noteLines.length}</span>` : ''}`
        : '<span style="font-size:11px;color:var(--silver-dim)">—</span>';

      const escapedDsi = (o.dsi || '').replace(/'/g, "\\'");

      // Build paid-out checkboxes
      const paidOutHtml = this._buildPaidOutHtml(o);

      const tr = document.createElement('tr');
      tr.style.cssText = 'border-bottom:1px solid rgba(0,0,0,0.06)';
      tr.innerHTML = `
        <td style="padding:10px 16px;font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:600;color:var(--white)">${o.repName}</td>
        <td style="padding:10px 16px;font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:600;color:var(--sc-cyan)">${o.traineeName || '—'}</td>
        <td style="padding:10px 16px;font-family:'Barlow Condensed',sans-serif;font-size:13px;color:var(--silver)">${o.dsi}</td>
        <td style="padding:10px 16px;font-family:'Barlow Condensed',sans-serif;font-size:13px;color:var(--silver)">${o.dateOfSale}</td>
        <td style="padding:10px 16px;font-family:'Barlow Condensed',sans-serif;font-size:12px;color:var(--silver-dim)">${soldStr}</td>
        <td style="padding:10px 16px;text-align:center;font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:700;color:var(--white)">${o.units}</td>
        <td style="padding:10px 16px">${paidOutHtml}</td>
        <td style="padding:10px 16px"><span style="display:inline-block;padding:3px 10px;border-radius:6px;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;background:${sColor}22;color:${sColor};border:1px solid ${sColor}44">${o.status}</span></td>
        <td style="padding:10px 16px;text-align:center">
          <button onclick="Orders.openNoteModal(${o.rowIndex},'${escapedDsi}')"
            style="background:rgba(0,200,255,0.1);border:1px solid rgba(0,200,255,0.3);border-radius:6px;color:var(--sc-cyan);padding:4px 12px;font-family:'Barlow Condensed',sans-serif;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;cursor:pointer">Notes</button>
          <div style="margin-top:4px">${notePreview}</div>
        </td>`;
      tbody.appendChild(tr);
    });
  },

  // Build paid-out checkbox HTML for a single order
  _buildPaidOutHtml(order) {
    // Determine which products have units
    const products = [];
    OFFICE_CONFIG.columns.products.forEach(prod => {
      const qty = order[prod.key] || 0;
      if (qty > 0) {
        products.push({ key: prod.key, label: prod.label, qty: qty });
      }
    });

    if (products.length === 0) return '<span style="font-size:11px;color:var(--silver-dim)">—</span>';

    // Build/merge paid-out state from saved data
    const saved = order.paidOut || {};
    const state = {};
    products.forEach(p => {
      const savedArr = Array.isArray(saved[p.key]) ? saved[p.key] : [];
      state[p.key] = Array.from({ length: p.qty }, (_, i) => savedArr[i] === true);
    });

    // Check if all are paid
    const totalUnits = products.reduce((sum, p) => sum + p.qty, 0);
    const paidCount = products.reduce((sum, p) => sum + state[p.key].filter(v => v).length, 0);
    const allPaid = totalUnits > 0 && paidCount === totalUnits;
    const rowId = order.rowIndex;

    let html = '';

    // ALL checkbox
    html += `<label style="display:flex;align-items:center;gap:4px;cursor:pointer;margin-bottom:4px;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:${allPaid ? 'var(--green)' : 'var(--silver-dim)'}">
      <input type="checkbox" ${allPaid ? 'checked' : ''} onchange="App._togglePaidOutAll(${rowId},this.checked)"
        style="accent-color:#22c55e;cursor:pointer;width:14px;height:14px"> ALL
    </label>`;

    // Per-product checkboxes
    products.forEach(p => {
      html += `<div style="display:flex;align-items:center;gap:3px;margin-bottom:2px">
        <span style="font-family:'Barlow Condensed',sans-serif;font-size:10px;font-weight:600;letter-spacing:0.5px;color:var(--silver-dim);min-width:32px">${p.label}</span>`;
      for (let i = 0; i < p.qty; i++) {
        const checked = state[p.key][i] ? 'checked' : '';
        html += `<input type="checkbox" ${checked} onchange="App._togglePaidOutUnit(${rowId},'${p.key}',${i},this.checked)"
          style="accent-color:#22c55e;cursor:pointer;width:13px;height:13px">`;
      }
      html += `</div>`;
    });

    // Paid count summary
    if (paidCount > 0 && paidCount < totalUnits) {
      html += `<div style="font-family:'Barlow Condensed',sans-serif;font-size:9px;color:var(--sc-cyan);margin-top:2px">${paidCount}/${totalUnits} paid</div>`;
    } else if (allPaid) {
      html += `<div style="font-family:'Barlow Condensed',sans-serif;font-size:9px;color:var(--green);margin-top:2px">✓ All paid</div>`;
    }

    return html;
  },

  // Toggle ALL paid-out checkboxes for an order
  _togglePaidOutAll(rowIndex, checked) {
    const order = this._payrollOrders.find(o => o.rowIndex === rowIndex);
    if (!order) return;

    // Build state — set all to checked/unchecked
    const state = {};
    OFFICE_CONFIG.columns.products.forEach(prod => {
      const qty = order[prod.key] || 0;
      if (qty > 0) {
        state[prod.key] = Array.from({ length: qty }, () => checked);
      }
    });

    order.paidOut = state;
    this._persistPaidOut(order);
    this._filterPayrollOrders(); // re-render to update all checkbox visuals
  },

  // Toggle a single unit's paid-out checkbox
  _togglePaidOutUnit(rowIndex, prodKey, unitIdx, checked) {
    const order = this._payrollOrders.find(o => o.rowIndex === rowIndex);
    if (!order) return;

    // Initialize paidOut state if needed
    if (!order.paidOut || typeof order.paidOut !== 'object') order.paidOut = {};

    // Ensure array exists for this product with correct length
    const qty = order[prodKey] || 0;
    if (!Array.isArray(order.paidOut[prodKey])) {
      order.paidOut[prodKey] = Array.from({ length: qty }, () => false);
    }

    // Update the specific unit
    order.paidOut[prodKey][unitIdx] = checked;

    this._persistPaidOut(order);
    this._filterPayrollOrders(); // re-render to update ALL checkbox state
  },

  // Debounced save of paid-out state to server
  _paidOutDebounce: {},

  _persistPaidOut(order) {
    // Debounce saves per order (wait 600ms to batch rapid clicks)
    if (this._paidOutDebounce[order.rowIndex]) {
      clearTimeout(this._paidOutDebounce[order.rowIndex]);
    }
    this._paidOutDebounce[order.rowIndex] = setTimeout(async () => {
      try {
        await SheetsAPI.post(OFFICE_CONFIG, 'savePaidOut', {
          rowIndex: order.rowIndex,
          paidOut: order.paidOut
        });
      } catch (err) {
        console.error('Failed to save paid-out state:', err);
      }
    }, 600);
  },

  // ══════════════════════════════════════════════
  // PROFILE SHORTCUTS
  // ══════════════════════════════════════════════
  openPersonProfile(name) { Render.openPersonProfile(name); },
  openTeamProfile(name) { Render.openTeamProfile(name); },

  // ══════════════════════════════════════════════
  // ROLE SWITCHER (superadmin View-As)
  // ══════════════════════════════════════════════
  toggleViewAs() {
    this.state.viewAsActive = !this.state.viewAsActive;
    const controls = document.getElementById('view-as-controls');
    const personaWrap = document.getElementById('role-persona-wrap');
    const toggleBtn = document.getElementById('view-as-toggle');

    if (this.state.viewAsActive) {
      if (controls) controls.style.display = 'flex';
      if (personaWrap) personaWrap.style.display = '';
      if (toggleBtn) toggleBtn.textContent = '◂ Close';
      // Keep current view-as role as-is (default to admin on first open)
    } else {
      if (controls) controls.style.display = 'none';
      if (personaWrap) personaWrap.style.display = 'none';
      if (toggleBtn) toggleBtn.textContent = 'Admin ▸';
      // Collapse back to admin view
      this.setRole('admin');
    }
  },

  _applySuperAdminViewAs() {
    const toggleBtn = document.getElementById('view-as-toggle');
    const controls = document.getElementById('view-as-controls');
    const personaWrap = document.getElementById('role-persona-wrap');
    if (toggleBtn) toggleBtn.style.display = '';
    // Start collapsed — view as admin
    this.state.viewAsActive = false;
    if (controls) controls.style.display = 'none';
    if (personaWrap) personaWrap.style.display = 'none';
    if (toggleBtn) toggleBtn.textContent = 'Admin ▸';
    this.state.currentRole = 'admin';
    this.updateNav();
  },

  setRole(role) {
    this.state.currentRole = role;
    this.state.currentNav = 'leaderboard';
    document.querySelectorAll('.role-pill').forEach(b => b.classList.toggle('active', b.dataset.role === role));
    this.populatePersonaSelect(role);
    this.updateRoleIndicator();
    this.updateNav();
    Render.closeProfile();
    Render.renderAll(this.state.people, this.state.teams);
  },

  setPersona(name) {
    this.state.currentPersona = name;
    this.state.currentNav = 'leaderboard';
    this.updateRoleIndicator();
    this.updateNav();
    Render.closeProfile();
    Render.renderAll(this.state.people, this.state.teams);
  },

  populatePersonaSelect(role) {
    const sel = document.getElementById('role-persona');
    const wrap = document.getElementById('role-persona-wrap');
    if (!sel || !wrap) return;

    const roster = this.state.roster || {};
    let people = Object.entries(roster)
      .filter(([, r]) => r.rank === role && !r.deactivated)
      .map(([, r]) => r.name)
      .sort();

    if (people.length === 0) people = this.state.people.map(p => p.name).slice(0, 5);

    this.state.currentPersona = people[0] || 'Admin';
    if (people.length <= 1) {
      wrap.style.display = 'none';
    } else {
      wrap.style.display = 'block';
      sel.innerHTML = people.map(n => `<option value="${n}">${n}</option>`).join('');
      sel.value = this.state.currentPersona;
    }
  },

  updateRoleIndicator() {
    const team = Roster.getEffectiveTeam(this.state.currentPersona, this.state.people);
    const existing = document.getElementById('role-team-tag');
    if (existing) existing.remove();

    const role = this.state.currentRole;
    if (team && (role === 'jd' || role === 'l1' || role === 'rep')) {
      const tag = document.createElement('div');
      tag.id = 'role-team-tag';
      tag.style.cssText = 'font-family:"Barlow Condensed",sans-serif;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:var(--sc-cyan);background:rgba(26,92,229,0.15);border:1px solid rgba(26,92,229,0.3);border-radius:6px;padding:4px 10px;white-space:nowrap';
      tag.textContent = '⬡ ' + team;
      const switcher = document.getElementById('role-switcher');
      if (switcher) switcher.appendChild(tag);
    }
  },

  // ══════════════════════════════════════════════
  // ROSTER ACTIONS (called from onclick)
  // ══════════════════════════════════════════════
  async setPersonRole(name, role) {
    await Roster.setRole(name, role, OFFICE_CONFIG);
    // Update local person object
    const p = this.state.people.find(x => x.name === name);
    if (p) { p._roleKey = role; p.role = OFFICE_CONFIG.roles[role]?.label || role; }
    const email = Roster.getEmail(name);
    if (email && this.state.roster[email]) this.state.roster[email].rank = role;
  },

  async setPersonTeam(name, team) {
    await Roster.setTeam(name, team, this.state.people, OFFICE_CONFIG);
    const email = Roster.getEmail(name);
    if (email && this.state.roster[email]) this.state.roster[email].team = team;
    const hasHierarchy = Object.keys(this.state.teamsData).length > 0;
    this.state.teams = hasHierarchy
      ? DataPipeline.buildTeamHierarchy(this.state.people, this.state.teamsData, OFFICE_CONFIG)
      : DataPipeline.buildTeams(this.state.people, OFFICE_CONFIG);
    Render.renderTeamGrid(this.state.teams);
  },

  async savePersonInfo(oldEmail) {
    const nameInput = document.getElementById('roster-edit-name-' + oldEmail);
    const emailInput = document.getElementById('roster-edit-email-' + oldEmail);
    if (!nameInput || !emailInput) return;

    const newName = nameInput.value.trim();
    const newEmail = emailInput.value.trim().toLowerCase();

    if (!newName) { this.showToast('Name is required'); return; }
    if (!newEmail || !newEmail.includes('@')) { this.showToast('Valid email is required'); return; }

    // Check for email conflict (only if changed)
    if (newEmail !== oldEmail && this.state.roster[newEmail]) {
      this.showToast('That email is already in the roster');
      return;
    }

    try {
      const payload = { email: oldEmail, name: newName };
      if (newEmail !== oldEmail) payload.newEmail = newEmail;
      await SheetsAPI.post(OFFICE_CONFIG, 'updateRosterEntry', payload);

      // Update local state
      const rosterEntry = this.state.roster[oldEmail];
      if (rosterEntry) {
        rosterEntry.name = newName;
        if (newEmail !== oldEmail) {
          this.state.roster[newEmail] = rosterEntry;
          delete this.state.roster[oldEmail];
          // Update emailMap
          Roster.emailMap[newName] = newEmail;
        }
      }
      // Update people array
      const p = this.state.people.find(x => x.email === oldEmail || Roster.getEmail(x.name) === oldEmail);
      if (p) {
        p.name = newName;
        p.email = newEmail;
      }

      this.showToast('Updated ' + newName);
      // Re-render to show changes
      Roster.renderRoster(this.state.people, this.state.currentRole, OFFICE_CONFIG);
    } catch (err) {
      this.showToast('Failed to update: ' + err.message);
    }
  },

  async toggleDeactivate(name) {
    await Roster.toggleDeactivate(name, OFFICE_CONFIG);
    Roster.renderRoster(this.state.people, this.state.currentRole, OFFICE_CONFIG);
    Render.renderMainTable(this.state.people);
    Render.renderTeamGrid(this.state.teams);
    // Refresh manage team tab if visible
    Render._refreshManageTeam();
  },

  async approveUnlock(name) {
    await Roster.approveUnlock(name, OFFICE_CONFIG);
    this.updateNav();
    Roster.renderRoster(this.state.people, this.state.currentRole, OFFICE_CONFIG);
  },

  async denyUnlock(name) {
    await Roster.denyUnlock(name, OFFICE_CONFIG);
    this.updateNav();
    Roster.renderRoster(this.state.people, this.state.currentRole, OFFICE_CONFIG);
  },

  // ══════════════════════════════════════════════
  // ADD MEMBER (JD+ only)
  // ══════════════════════════════════════════════
  openAddMemberModal(presetTeam) {
    const modal = document.getElementById('add-member-modal');
    if (!modal) return;
    modal.style.display = 'flex';
    // Clear form
    const fields = ['add-member-email', 'add-member-name', 'add-member-error'];
    fields.forEach(id => {
      const el = document.getElementById(id);
      if (el) el.tagName === 'INPUT' ? el.value = '' : el.textContent = '';
    });
    // Dynamically populate team dropdown
    const teamSel = document.getElementById('add-member-team');
    if (teamSel) {
      const teamNames = (this.state.teams && this.state.teams.length > 0)
        ? this.state.teams.filter(t => t.teamId !== '_unassigned').map(t => t.name)
        : OFFICE_CONFIG.teams;
      teamSel.innerHTML = '<option value="Unassigned">Unassigned</option>'
        + teamNames.map(t => `<option value="${t}">${t}</option>`).join('');
      // Pre-select team if provided (from Manage Team tab)
      if (presetTeam) {
        teamSel.value = presetTeam;
      } else {
        teamSel.selectedIndex = 0;
      }
    }
    const rankSel = document.getElementById('add-member-rank');
    if (rankSel) rankSel.selectedIndex = 0;
  },

  closeAddMemberModal() {
    const modal = document.getElementById('add-member-modal');
    if (modal) modal.style.display = 'none';
  },

  _addingMember: false,

  async submitAddMember() {
    if (this._addingMember) return; // Prevent double-click

    const email = document.getElementById('add-member-email')?.value?.trim().toLowerCase();
    const name = document.getElementById('add-member-name')?.value?.trim();
    const team = document.getElementById('add-member-team')?.value;
    const rank = document.getElementById('add-member-rank')?.value;
    const errorEl = document.getElementById('add-member-error');
    const submitBtn = document.querySelector('#add-member-modal button[onclick*="submitAddMember"]');

    if (!email || !name) {
      if (errorEl) errorEl.textContent = 'Email and name are required';
      return;
    }
    if (!email.includes('@')) {
      if (errorEl) errorEl.textContent = 'Enter a valid email address';
      return;
    }
    if (this.state.roster[email]) {
      if (errorEl) errorEl.textContent = 'This email is already in the roster';
      return;
    }

    if (errorEl) errorEl.textContent = '';

    // Disable button while saving
    this._addingMember = true;
    if (submitBtn) { submitBtn.textContent = 'Adding...'; submitBtn.style.opacity = '0.5'; }

    try {
      await Roster.addNewPerson(email, name, team || 'Unassigned', rank || 'rep', OFFICE_CONFIG);
      // Update local state
      this.state.roster[email] = {
        name,
        team: team || 'Unassigned',
        rank: rank || 'rep',
        deactivated: false,
        dateAdded: new Date().toISOString().split('T')[0]
      };
      this.closeAddMemberModal();
      this.showToast(`${name} added to the roster`);
      // Refresh to get full updated data
      await this.refreshData();
    } catch (err) {
      if (errorEl) errorEl.textContent = 'Failed to add: ' + err.message;
    } finally {
      this._addingMember = false;
      if (submitBtn) { submitBtn.textContent = 'ADD MEMBER'; submitBtn.style.opacity = ''; }
    }
  },

  // ══════════════════════════════════════════════
  // TEAM CUSTOMIZATION
  // ══════════════════════════════════════════════
  handleTeamEditClick() {
    const status = Roster.getUnlockStatus(this.state.currentPersona);
    if (status === 'approved') {
      this.openTeamCustomize();
    } else if (status === 'pending') {
      this.showUnlockModal(`<div style="color:#f0b429;font-size:13px;margin-bottom:16px">⏳ Your request is pending approval.</div>`);
    } else {
      this.showUnlockModal('');
    }
  },

  showUnlockModal(statusHtml) {
    const modal = document.getElementById('unlock-request-modal');
    const status = document.getElementById('unlock-request-status');
    const btn = document.getElementById('unlock-send-btn');
    if (status) status.innerHTML = statusHtml;
    if (btn) btn.style.display = Roster.getUnlockStatus(this.state.currentPersona) === 'pending' ? 'none' : '';
    if (modal) modal.style.display = 'flex';
  },

  sendUnlockRequest() {
    Roster.sendUnlockRequest(this.state.currentPersona, OFFICE_CONFIG);
    const status = document.getElementById('unlock-request-status');
    const btn = document.getElementById('unlock-send-btn');
    if (status) status.innerHTML = `<div style="color:#22c55e;font-size:13px;margin-bottom:16px">✓ Request sent!</div>`;
    if (btn) btn.style.display = 'none';
    this.updateNav();
  },

  closeUnlockModal() {
    const modal = document.getElementById('unlock-request-modal');
    if (modal) modal.style.display = 'none';
  },

  // ── Emoji picker helpers ──
  _pickRandomEmojis(count, mustInclude) {
    const pool = [...OFFICE_CONFIG.teamEmojis];
    const result = new Set();
    if (mustInclude && pool.includes(mustInclude)) result.add(mustInclude);
    while (result.size < Math.min(count, pool.length)) {
      result.add(pool.splice(Math.floor(Math.random() * pool.length), 1)[0]);
    }
    return [...result];
  },

  _renderEmojiPicker(pickerId, displayId, selectFn, highlightFn) {
    const picker = document.getElementById(pickerId);
    if (!picker) return;
    const selected = selectFn === 'App.selectEmoji' ? this._selectedEmoji : TeamsManager._selectedTeamEmoji;
    const emojis = this._pickRandomEmojis(12, selected);
    picker.innerHTML = emojis.map(e =>
      `<span onclick="${selectFn}('${e}')" style="cursor:pointer;font-size:24px;padding:4px 6px;border-radius:8px;border:2px solid transparent;transition:all 0.1s" id="${pickerId === 'emoji-picker' ? 'emoji-opt-' : 'team-emoji-opt-'}${e.codePointAt(0)}">${e}</span>`
    ).join('')
    + `<span onclick="App._rerollEmojis('${pickerId}','${displayId}','${selectFn}')" style="cursor:pointer;font-size:18px;padding:4px 8px;border-radius:8px;background:rgba(0,0,0,0.06);border:1px solid rgba(26,92,229,0.2);color:var(--silver-dim);font-family:'Barlow Condensed',sans-serif;font-weight:700;letter-spacing:1px;display:inline-flex;align-items:center;gap:4px" title="Show more emojis">🔄</span>`;
    if (highlightFn) highlightFn();
  },

  _rerollEmojis(pickerId, displayId, selectFn) {
    const highlightFn = selectFn === 'App.selectEmoji'
      ? () => this.highlightSelectedEmoji()
      : () => TeamsManager.highlightTeamEmoji();
    this._renderEmojiPicker(pickerId, displayId, selectFn, highlightFn);
  },

  openTeamCustomize() {
    const modal = document.getElementById('team-customize-modal');
    if (!modal) return;
    modal.style.display = 'flex';
    const myTeam = Roster.getEffectiveTeam(this.state.currentPersona, this.state.people);
    const existing = Roster.teamCustomizations[this.state.currentPersona];
    this._selectedEmoji = existing?.emoji || '⚡';
    const input = document.getElementById('team-name-input');
    if (input) input.value = existing?.name || (myTeam || '');
    const display = document.getElementById('emoji-display');
    if (display) display.textContent = this._selectedEmoji;
    this._renderEmojiPicker('emoji-picker', 'emoji-display', 'App.selectEmoji', () => this.highlightSelectedEmoji());
  },

  _selectedEmoji: '⚡',

  selectEmoji(e) {
    this._selectedEmoji = e;
    const display = document.getElementById('emoji-display');
    if (display) display.textContent = e;
    this.highlightSelectedEmoji();
  },

  highlightSelectedEmoji() {
    document.querySelectorAll('#emoji-picker span[id^="emoji-opt-"]').forEach(s => {
      s.style.borderColor = 'transparent'; s.style.background = 'transparent';
    });
    const el = document.getElementById('emoji-opt-' + this._selectedEmoji.codePointAt(0));
    if (el) { el.style.borderColor = 'var(--sc-cyan)'; el.style.background = 'rgba(26,92,229,0.2)'; }
  },

  saveTeamCustomize() {
    const input = document.getElementById('team-name-input');
    const name = input?.value.trim();
    if (!name) return;
    Roster.setTeamCustomization(this.state.currentPersona, this._selectedEmoji || '⚡', name, OFFICE_CONFIG);
    this.closeTeamCustomize();
    this.updateNav();
    Render.renderTeamGrid(this.state.teams);
  },

  closeTeamCustomize() {
    const modal = document.getElementById('team-customize-modal');
    if (modal) modal.style.display = 'none';
  },

  // ══════════════════════════════════════════════
  // TEAM HIERARCHY CRUD
  // ══════════════════════════════════════════════
  openCreateTeamModal() {
    const modal = document.getElementById('team-crud-modal');
    if (!modal) return;
    modal.style.display = 'flex';
    TeamsManager.populateModal(null);
  },

  openEditTeamModal(teamId) {
    const modal = document.getElementById('team-crud-modal');
    if (!modal) return;
    modal.style.display = 'flex';
    TeamsManager.populateModal(teamId);
  },

  closeTeamCrudModal() {
    const modal = document.getElementById('team-crud-modal');
    if (modal) modal.style.display = 'none';
  },

  _savingTeam: false,

  async submitTeamModal() {
    if (this._savingTeam) return;

    const vals = TeamsManager.getModalValues();
    const errorEl = document.getElementById('team-modal-error');
    const submitBtn = document.getElementById('team-modal-submit');

    if (!vals.name) {
      if (errorEl) errorEl.textContent = 'Team name is required';
      return;
    }

    const isEdit = !!vals.teamId;
    const teamId = isEdit ? vals.teamId : TeamsManager.generateTeamId(vals.name);

    if (errorEl) errorEl.textContent = '';
    this._savingTeam = true;
    if (submitBtn) { submitBtn.textContent = 'Saving...'; submitBtn.style.opacity = '0.5'; }

    try {
      const action = isEdit ? 'updateTeam' : 'addTeam';
      const payload = {
        teamId,
        name: vals.name,
        parentId: vals.parentId,
        leaderId: vals.leaderId,
        emoji: vals.emoji
      };
      const result = await SheetsAPI.post(OFFICE_CONFIG, action, payload);
      if (result.data?.error) {
        if (errorEl) errorEl.textContent = result.data.error;
        return;
      }

      this.closeTeamCrudModal();
      this.showToast(isEdit ? `${vals.name} updated` : `${vals.name} created`);
      await this.refreshData();
      // Re-render teams page
      if (this.state.currentNav === 'teams') {
        TeamsManager.renderTeamsPage(this.state.people, this.state.currentRole, OFFICE_CONFIG);
      }
    } catch (err) {
      if (errorEl) errorEl.textContent = 'Failed: ' + err.message;
    } finally {
      this._savingTeam = false;
      if (submitBtn) {
        submitBtn.textContent = isEdit ? 'SAVE CHANGES' : 'CREATE TEAM';
        submitBtn.style.opacity = '';
      }
    }
  },

  async confirmDeleteTeam(teamId, teamName) {
    if (!confirm(`Delete team "${teamName}"? Members will become unassigned.`)) return;

    try {
      await SheetsAPI.post(OFFICE_CONFIG, 'deleteTeam', { teamId });
      this.showToast(`${teamName} deleted`);
      await this.refreshData();
      if (this.state.currentNav === 'teams') {
        TeamsManager.renderTeamsPage(this.state.people, this.state.currentRole, OFFICE_CONFIG);
      }
    } catch (err) {
      this.showToast('Failed to delete team');
    }
  },

  // ══════════════════════════════════════════════
  // UTILITIES
  // ══════════════════════════════════════════════
  showError(msg, err) {
    console.error(msg, err);
    Auth.hideLoading();
    const container = document.querySelector('.container');
    if (container) {
      container.innerHTML = `
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:120px 24px;text-align:center">
          <div style="font-size:48px;margin-bottom:16px">⚠️</div>
          <div style="font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:800;color:var(--white);margin-bottom:8px">${msg}</div>
          <div style="color:var(--silver-dim);font-size:13px;max-width:500px;line-height:1.6;margin-bottom:24px">${err?.message || ''}</div>
          <button onclick="location.reload()" style="background:var(--blue-core);border:none;border-radius:8px;padding:12px 24px;color:#fff;font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:700;letter-spacing:2px;text-transform:uppercase;cursor:pointer">Retry</button>
        </div>`;
    }
  },

  showToast(msg) {
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.textContent = msg;
    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), 3000);
  },

  // Manual refresh
  async manualRefresh() {
    await this.refreshData();
    this.showToast('Data refreshed');
  },

  // Logout
  logout() { Auth.logout(); }
};

// ══════════════════════════════════════════════
// BOOT
// ══════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => App.init());
