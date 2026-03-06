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
    currentNav: 'profile',
    lastUpdated: null,
    refreshTimer: null,
    tableauDsi: {},      // DSI-keyed Tableau summary
    tableauByRep: {},    // email-keyed Tableau rep summary
    tableauByName: {},   // Tableau REP name-keyed summary (for stored name lookups)
    possibleTableauNames: {}  // email → [Tableau REP names] for picker popup
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
      // Set up superadmin View-As (collapsed — views as admin)
      if (this.state.realRole === 'superadmin') {
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

    // Set up superadmin View-As (collapsed — views as admin)
    if (this.state.realRole === 'superadmin') {
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

    // Extract Tableau summary (included in default doGet response)
    if (apiData.tableauSummary) {
      this.state.tableauDsi = apiData.tableauSummary.dsiSummary || {};
      this.state.tableauByRep = apiData.tableauSummary.repSummary || {};
      this.state.tableauByName = apiData.tableauSummary.repByName || {};
      this.state.possibleTableauNames = apiData.tableauSummary.possibleTableauNames || {};
      console.log('[Tableau] DSIs loaded:', Object.keys(this.state.tableauDsi).length);
      console.log('[Tableau] Possible names:', Object.keys(this.state.possibleTableauNames).length);
      const firstDsi = Object.keys(this.state.tableauDsi)[0];
      if (firstDsi) console.log('[Tableau] Sample DSI:', firstDsi, this.state.tableauDsi[firstDsi]);
    } else {
      console.warn('[Tableau] No tableauSummary in API response. Keys:', Object.keys(apiData));
    }

    // Enrich person and team metrics with Tableau data (pass name-keyed + roster for stored name lookups)
    DataPipeline.enrichWithTableau(this.state.people, this.state.tableauByRep, this.state.tableauByName, this.state.roster);
    DataPipeline.enrichTeamsWithTableau(this.state.teams, this.state.tableauByRep);

    // Enrich churn buckets from _TableauChurnReport
    DataPipeline.enrichWithChurnReport(this.state.people, apiData.churnReport);
    DataPipeline.enrichTeamsWithChurn(this.state.teams);

    Roster.init(this.state.roster);
    Roster.initFromApi(apiData);
    TeamsManager.init(this.state.teamsData);
    this.updateNav();
    Render.renderAll(this.state.people, this.state.teams);
    // Show leaderboard section when on leaderboard, profile, or team views
    const lbSection = document.getElementById('leaderboard-section');
    if (lbSection && ['leaderboard', 'profile', 'team'].includes(this.state.currentNav)) {
      lbSection.style.display = '';
    }
    // Land on My Profile by default
    if (this.state.currentNav === 'profile') {
      this.openPersonProfile(this.state.currentPersona);
    }
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
    this.updateLastUpdated();

    // Auto-claim Tableau name if an unclaimed name is found
    this._autoClaimTableauName();
  },

  // ── Tableau Name Auto-Claim ──
  // Checks if the current user has an unclaimed Tableau REP name.
  // A name is "claimed" if any other roster member already has it stored.
  // If exactly one unclaimed name is found → auto-save it silently.
  _autoClaimTableauName() {
    const email = this.state.currentEmail;
    if (!email) return;

    // Only for sales roles
    const salesRoles = ['rep', 'l1', 'jd', 'manager'];
    if (!salesRoles.includes(this.state.currentRole)) return;

    // Skip if already has a stored tableauName
    const rosterEntry = this.state.roster[email];
    if (rosterEntry && rosterEntry.tableauName) return;

    // Get all Tableau names tied to this rep's posted DSIs
    const possibleNames = this.state.possibleTableauNames[email];
    if (!possibleNames || possibleNames.length === 0) return;

    // Build set of already-claimed names (stored by other roster members)
    const claimedNames = new Set();
    Object.entries(this.state.roster).forEach(([rosterEmail, entry]) => {
      if (rosterEmail !== email && entry.tableauName) {
        claimedNames.add(entry.tableauName);
      }
    });

    // Filter to unclaimed names only
    const unclaimed = possibleNames.filter(n => !claimedNames.has(n));
    console.log('[Tableau] Possible names for', email, ':', possibleNames, '| Unclaimed:', unclaimed);

    if (unclaimed.length === 1) {
      // Exactly one unclaimed name — auto-claim it
      this._saveTableauName(unclaimed[0]);
    }
    // 0 or 2+ unclaimed — do nothing, wait for the situation to resolve
  },

  async _saveTableauName(name) {
    const email = this.state.currentEmail;
    if (!email || !name) return;

    console.log('[Tableau] Auto-claiming name:', name, 'for', email);

    // Update local roster state immediately
    if (this.state.roster[email]) {
      this.state.roster[email].tableauName = name;
    }

    // Re-run enrichment with the new name
    DataPipeline.enrichWithTableau(this.state.people, this.state.tableauByRep, this.state.tableauByName, this.state.roster);
    DataPipeline.enrichTeamsWithTableau(this.state.teams, this.state.tableauByRep);

    // Re-render current view if on profile
    if (this.state.currentNav === 'profile') {
      this.openPersonProfile(this.state.currentPersona);
    }

    // Persist to backend
    try {
      await SheetsAPI.post(OFFICE_CONFIG, 'setTableauName', { email, tableauName: name });
      console.log('[Tableau] Saved tableauName:', name, 'for', email);
    } catch (err) {
      console.error('[Tableau] Failed to save tableauName:', err);
    }
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
    const isSA = (role === 'superadmin');

    // ── PERSONAL GROUP ──
    const showProfile = true;
    const showMyOrders = isSA || ['rep', 'l1', 'jd', 'manager', 'owner'].includes(role);

    // ── TEAM GROUP ──
    const showTeam = isSA || ['jd', 'manager'].includes(role);
    const showEdit = isSA || (role === 'jd');
    const showTeamRoster = isSA || ['jd', 'manager'].includes(role);

    // ── OFFICE GROUP ──
    const showAllOrders = isSA || ['owner', 'admin'].includes(role);
    const showPeople = isSA || ['owner', 'admin'].includes(role);
    const showTeams = isSA || ['owner', 'admin'].includes(role);
    const showLeaderboard = true;
    const showOffice = isSA || (role === 'owner');
    const curEmail = (this.state.currentEmail || '').toLowerCase();
    const payrollMgr = (this.state.settings.payrollManager || '').toLowerCase();
    const showPayroll = isSA || (role === 'owner') || (payrollMgr && curEmail === payrollMgr);

    // ── Separator visibility ──
    const hasTeamGroup = showTeam || showEdit || showTeamRoster;
    const showSep1 = hasTeamGroup;
    const showSep2 = true; // always separates last group from Office

    const setDisplay = (id, show) => {
      const el = document.getElementById(id);
      if (el) el.style.display = show ? '' : 'none';
    };

    // Personal
    setDisplay('nav-profile', showProfile);
    setDisplay('nav-my-orders', showMyOrders);

    // Separators
    setDisplay('nav-sep-1', showSep1);
    setDisplay('nav-sep-2', showSep2);

    // Team
    setDisplay('nav-team', showTeam);
    setDisplay('nav-team-edit', showEdit);
    setDisplay('nav-team-roster', showTeamRoster);

    // Office
    setDisplay('nav-all-orders', showAllOrders);
    setDisplay('nav-people', showPeople);
    setDisplay('nav-teams', showTeams);
    setDisplay('nav-leaderboard', showLeaderboard);
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

    // People notification badge (pending unlock requests)
    if (showPeople) {
      const pending = Roster.getPendingRequests().length;
      const badge = document.getElementById('roster-notif-badge');
      if (badge) {
        badge.style.display = pending > 0 ? 'flex' : 'none';
        badge.textContent = pending;
      }
    }

    // Active tab highlight
    const navIdMap = { allOrders: 'nav-all-orders', myOrders: 'nav-my-orders', roster: 'nav-people', teamRoster: 'nav-team-roster' };
    document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
    const activeTab = document.getElementById(navIdMap[this.state.currentNav] || ('nav-' + this.state.currentNav));
    if (activeTab) activeTab.classList.add('active');
  },

  navTo(tab) {
    // Hide ALL page sections
    const sections = {
      leaderboard: 'leaderboard-section',
      roster: 'roster-page',
      teams: 'teams-page',
      allOrders: 'all-orders-page',
      myOrders: 'my-orders-page',
      office: 'office-page',
      payroll: 'payroll-page',
      teamRoster: 'team-roster-page'
    };
    Object.values(sections).forEach(id => {
      const el = document.getElementById(id);
      if (el) el.style.display = 'none';
    });

    this.state.currentNav = tab;
    this.updateNav();

    // Profile & Team open overlay on top of leaderboard
    if (tab === 'profile') {
      const lb = document.getElementById('leaderboard-section');
      if (lb) lb.style.display = '';
      this.openPersonProfile(this.state.currentPersona);
      return;
    }
    if (tab === 'team') {
      const lb = document.getElementById('leaderboard-section');
      if (lb) lb.style.display = '';
      const myTeam = Roster.getEffectiveTeam(this.state.currentPersona, this.state.people);
      if (myTeam) this.openTeamProfile(myTeam);
      return;
    }

    // Close profile overlay for all other tabs
    Render.closeProfile();

    // Show the target section
    if (tab === 'leaderboard') {
      const lb = document.getElementById('leaderboard-section');
      if (lb) lb.style.display = '';
    } else if (tab === 'roster') {
      Roster.renderRoster(this.state.people, this.state.currentRole, OFFICE_CONFIG);
    } else if (tab === 'teamRoster') {
      this._renderTeamRoster();
    } else if (tab === 'teams') {
      TeamsManager.renderTeamsPage(this.state.people, this.state.currentRole, OFFICE_CONFIG);
    } else if (tab === 'allOrders') {
      this._loadAndRenderOrders('all');
    } else if (tab === 'myOrders') {
      this._loadAndRenderOrders('my');
    } else if (tab === 'office') {
      this._renderOfficePage();
    } else if (tab === 'payroll') {
      this._loadAndRenderPayroll();
    }

    // Scroll to top
    window.scrollTo({ top: 0, behavior: 'instant' });
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
        const speStr = (o.speList || []).join(' ');
        const haystack = (o.repName + ' ' + o.dsi + ' ' + speStr).toLowerCase();
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
      tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:var(--silver-dim);padding:32px;font-family:\'Cerebri Sans\',\'DM Sans\',\'Inter\',sans-serif;font-size:14px">No trainee orders found</td></tr>';
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

      // Notes preview
      const noteLines = o.notes ? o.notes.split('\n').filter(l => l.trim()) : [];
      const notePreview = noteLines.length > 0
        ? `<span style="font-size:11px;color:var(--silver-dim)">${noteLines[noteLines.length - 1].substring(0, 40)}${noteLines[noteLines.length - 1].length > 40 ? '...' : ''}</span>
           ${noteLines.length > 1 ? `<span style="background:rgba(0,0,0,0.15);color:var(--blue-core);font-size:9px;font-weight:700;border-radius:4px;padding:1px 5px;margin-left:4px">${noteLines.length}</span>` : ''}`
        : '<span style="font-size:11px;color:var(--silver-dim)">—</span>';

      const escapedDsi = (o.dsi || '').replace(/'/g, "\\'");

      // Build paid-out checkboxes
      const paidOutHtml = this._buildPaidOutHtml(o);

      const tr = document.createElement('tr');
      tr.style.cssText = 'border-bottom:1px solid rgba(0,0,0,0.06)';
      tr.innerHTML = `
        <td style="padding:10px 16px;font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:14px;font-weight:600;color:var(--white)">${o.repName}</td>
        <td style="padding:10px 16px;font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:14px;font-weight:600;color:var(--sc-cyan)">${o.traineeName || '—'}</td>
        <td style="padding:10px 16px;font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:13px;color:var(--silver)">${o.dsi}</td>
        <td style="padding:10px 16px;font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:13px;color:var(--silver)">${o.dateOfSale}</td>
        <td style="padding:10px 16px;font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:12px;color:var(--silver-dim)">${soldStr}</td>
        <td style="padding:10px 16px">${paidOutHtml}</td>
        <td style="padding:10px 16px;text-align:center">
          <button onclick="Orders.openNoteModal(${o.rowIndex},'${escapedDsi}')"
            style="background:rgba(44,110,106,0.1);border:1px solid rgba(44,110,106,0.3);border-radius:6px;color:var(--sc-cyan);padding:4px 12px;font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;cursor:pointer">Notes</button>
          <div style="margin-top:4px">${notePreview}</div>
        </td>`;
      tbody.appendChild(tr);
    });
  },

  // Build paid-out checkbox HTML for a single order
  _buildPaidOutHtml(order) {
    const speList = order.speList;

    // If SPE data available → use SPE-keyed UI
    if (speList && speList.length > 0) {
      return this._buildPaidOutHtmlSpe(order, speList);
    }
    // Fallback: legacy product-keyed UI
    return this._buildPaidOutHtmlLegacy(order);
  },

  // ── SPE-keyed paid-out UI ──
  _buildPaidOutHtmlSpe(order, speList) {
    // Migrate from legacy format if needed
    const saved = order.paidOut || {};
    if (!saved._v || saved._v < 2) {
      this._migratePaidOutToSpe(order, speList);
    }

    const rowId = order.rowIndex;
    const paidOut = order.paidOut || {};
    const totalSpe = speList.length;
    let paidCount = 0;
    speList.forEach(spe => { if (paidOut[spe] === true) paidCount++; });
    const allPaid = totalSpe > 0 && paidCount === totalSpe;

    let html = '';

    // ALL checkbox
    html += `<label style="display:flex;align-items:center;gap:4px;cursor:pointer;margin-bottom:4px;font-family:'Helvetica Neue','Inter',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:${allPaid ? 'var(--green)' : 'var(--silver-dim)'}">
      <input type="checkbox" ${allPaid ? 'checked' : ''} onchange="App._togglePaidOutAllSpe(${rowId},this.checked)"
        style="accent-color:#2E8B57;cursor:pointer;width:14px;height:14px"> ALL
    </label>`;

    // Per-SPE checkboxes
    speList.forEach(spe => {
      const checked = paidOut[spe] === true ? 'checked' : '';
      const shortSpe = spe.length > 16 ? '...' + spe.slice(-10) : spe;
      html += `<div style="display:flex;align-items:center;gap:3px;margin-bottom:2px">
        <input type="checkbox" ${checked} onchange="App._togglePaidOutSpe(${rowId},'${spe.replace(/'/g, "\\'")}',this.checked)"
          style="accent-color:#2E8B57;cursor:pointer;width:13px;height:13px">
        <span style="font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:9px;color:var(--silver-dim)" title="${spe}">${shortSpe}</span>
      </div>`;
    });

    // Count summary
    if (paidCount > 0 && paidCount < totalSpe) {
      html += `<div style="font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:9px;color:var(--sc-cyan);margin-top:2px">${paidCount}/${totalSpe} paid</div>`;
    } else if (allPaid) {
      html += `<div style="font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:9px;color:var(--green);margin-top:2px">\u2713 All paid</div>`;
    }

    return html;
  },

  // Migrate legacy product-keyed paidOut → SPE-keyed format
  _migratePaidOutToSpe(order, speList) {
    const old = order.paidOut || {};
    if (old._v === 2) return; // already migrated

    const newState = { _v: 2 };
    // Flatten old boolean arrays into a sequential list
    const flatBools = [];
    OFFICE_CONFIG.columns.products.forEach(prod => {
      if (Array.isArray(old[prod.key])) {
        old[prod.key].forEach(v => flatBools.push(v === true));
      }
    });

    // Map sequentially to SPE list (best-effort)
    speList.forEach((spe, i) => {
      newState[spe] = i < flatBools.length ? flatBools[i] : false;
    });

    order.paidOut = newState;
    // Persist the migrated format
    this._persistPaidOut(order);
  },

  // ── Legacy product-keyed paid-out UI (fallback) ──
  _buildPaidOutHtmlLegacy(order) {
    const products = [];
    OFFICE_CONFIG.columns.products.forEach(prod => {
      const qty = order[prod.key] || 0;
      if (qty > 0) {
        products.push({ key: prod.key, label: prod.label, qty: qty });
      }
    });

    if (products.length === 0) return '<span style="font-size:11px;color:var(--silver-dim)">\u2014</span>';

    const saved = order.paidOut || {};
    const state = {};
    products.forEach(p => {
      const savedArr = Array.isArray(saved[p.key]) ? saved[p.key] : [];
      state[p.key] = Array.from({ length: p.qty }, (_, i) => savedArr[i] === true);
    });

    const totalUnits = products.reduce((sum, p) => sum + p.qty, 0);
    const paidCount = products.reduce((sum, p) => sum + state[p.key].filter(v => v).length, 0);
    const allPaid = totalUnits > 0 && paidCount === totalUnits;
    const rowId = order.rowIndex;

    let html = '';

    html += `<label style="display:flex;align-items:center;gap:4px;cursor:pointer;margin-bottom:4px;font-family:'Helvetica Neue','Inter',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:${allPaid ? 'var(--green)' : 'var(--silver-dim)'}">
      <input type="checkbox" ${allPaid ? 'checked' : ''} onchange="App._togglePaidOutAll(${rowId},this.checked)"
        style="accent-color:#2E8B57;cursor:pointer;width:14px;height:14px"> ALL
    </label>`;

    products.forEach(p => {
      html += `<div style="display:flex;align-items:center;gap:3px;margin-bottom:2px">
        <span style="font-family:'Helvetica Neue','Inter',sans-serif;font-size:10px;font-weight:600;letter-spacing:0.5px;color:var(--silver-dim);min-width:32px">${p.label}</span>`;
      for (let i = 0; i < p.qty; i++) {
        const checked = state[p.key][i] ? 'checked' : '';
        html += `<input type="checkbox" ${checked} onchange="App._togglePaidOutUnit(${rowId},'${p.key}',${i},this.checked)"
          style="accent-color:#2E8B57;cursor:pointer;width:13px;height:13px">`;
      }
      html += `</div>`;
    });

    if (paidCount > 0 && paidCount < totalUnits) {
      html += `<div style="font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:9px;color:var(--sc-cyan);margin-top:2px">${paidCount}/${totalUnits} paid</div>`;
    } else if (allPaid) {
      html += `<div style="font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:9px;color:var(--green);margin-top:2px">\u2713 All paid</div>`;
    }

    return html;
  },

  // Toggle ALL paid-out checkboxes for an order (legacy product-keyed)
  _togglePaidOutAll(rowIndex, checked) {
    const order = this._payrollOrders.find(o => o.rowIndex === rowIndex);
    if (!order) return;

    const state = {};
    OFFICE_CONFIG.columns.products.forEach(prod => {
      const qty = order[prod.key] || 0;
      if (qty > 0) {
        state[prod.key] = Array.from({ length: qty }, () => checked);
      }
    });

    order.paidOut = state;
    this._persistPaidOut(order);
    this._filterPayrollOrders();
  },

  // Toggle a single unit's paid-out checkbox (legacy product-keyed)
  _togglePaidOutUnit(rowIndex, prodKey, unitIdx, checked) {
    const order = this._payrollOrders.find(o => o.rowIndex === rowIndex);
    if (!order) return;

    if (!order.paidOut || typeof order.paidOut !== 'object') order.paidOut = {};

    const qty = order[prodKey] || 0;
    if (!Array.isArray(order.paidOut[prodKey])) {
      order.paidOut[prodKey] = Array.from({ length: qty }, () => false);
    }

    order.paidOut[prodKey][unitIdx] = checked;

    this._persistPaidOut(order);
    this._filterPayrollOrders();
  },

  // Toggle ALL SPE paid-out checkboxes
  _togglePaidOutAllSpe(rowIndex, checked) {
    const order = this._payrollOrders.find(o => o.rowIndex === rowIndex);
    if (!order || !order.speList) return;

    const state = { _v: 2 };
    order.speList.forEach(spe => { state[spe] = checked; });

    order.paidOut = state;
    this._persistPaidOut(order);
    this._filterPayrollOrders();
  },

  // Toggle a single SPE paid-out checkbox
  _togglePaidOutSpe(rowIndex, spe, checked) {
    const order = this._payrollOrders.find(o => o.rowIndex === rowIndex);
    if (!order) return;

    if (!order.paidOut || typeof order.paidOut !== 'object') order.paidOut = { _v: 2 };
    order.paidOut._v = 2;
    order.paidOut[spe] = checked;

    this._persistPaidOut(order);
    this._filterPayrollOrders();
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
    this.state.currentNav = 'profile';
    document.querySelectorAll('.role-pill').forEach(b => b.classList.toggle('active', b.dataset.role === role));
    this.populatePersonaSelect(role);
    this.updateRoleIndicator();
    this.updateNav();
    Render.closeProfile();
    Render.renderAll(this.state.people, this.state.teams);
    this.openPersonProfile(this.state.currentPersona);
  },

  setPersona(name) {
    this.state.currentPersona = name;
    this.state.currentNav = 'profile';
    this.updateRoleIndicator();
    this.updateNav();
    Render.closeProfile();
    Render.renderAll(this.state.people, this.state.teams);
    this.openPersonProfile(name);
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
      tag.style.cssText = 'font-family:"Helvetica Neue","Inter",sans-serif;font-size:11px;font-weight:700;letter-spacing:0.5px;text-transform:uppercase;color:var(--sc-cyan);background:rgba(0,0,0,0.15);border:1px solid rgba(0,0,0,0.3);border-radius:6px;padding:4px 10px;white-space:nowrap';
      tag.textContent = '⬡ ' + team;
      const switcher = document.getElementById('role-switcher');
      if (switcher) switcher.appendChild(tag);
    }
  },

  // ══════════════════════════════════════════════
  // ROSTER ACTIONS (called from onclick)
  // ══════════════════════════════════════════════
  async setPersonRole(name, role) {
    // Permission validation: prevent assigning roles above current user's rank
    const myRole = this.state.currentRole || 'rep';
    const myRank = OFFICE_CONFIG.roles[myRole]?.rank || 0;
    const targetRank = OFFICE_CONFIG.roles[role]?.rank || 0;

    // Never allow assigning owner or superadmin
    if (role === 'owner' || role === 'superadmin') return;
    // Admin can only be assigned by owner, admin, or superadmin
    if (role === 'admin' && myRole !== 'owner' && myRole !== 'admin' && myRole !== 'superadmin') return;
    // Can't assign role at or above own rank (unless superadmin)
    if (myRole !== 'superadmin' && targetRank >= myRank) return;

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

  async savePersonPhone(email) {
    const phoneInput = document.getElementById('roster-edit-phone-' + email);
    if (!phoneInput) return;
    const phone = phoneInput.value.trim();
    const name = Object.keys(Roster.emailMap).find(n => Roster.emailMap[n] === email);
    if (!name) { this.showToast('Could not find person'); return; }
    try {
      await Roster.setPhone(name, phone, OFFICE_CONFIG);
      if (this.state.roster[email]) this.state.roster[email].phone = phone;
      this.showToast(phone ? 'Phone updated' : 'Phone removed');
      Roster.renderRoster(this.state.people, this.state.currentRole, OFFICE_CONFIG);
    } catch (err) {
      this.showToast('Failed to update phone: ' + err.message);
    }
  },

  async toggleDeactivate(name) {
    await Roster.toggleDeactivate(name, OFFICE_CONFIG);
    Roster.renderRoster(this.state.people, this.state.currentRole, OFFICE_CONFIG);
    Render.renderMainTable(this.state.people);
    Render.renderTeamGrid(this.state.teams);
    // Refresh team roster page if visible
    this._refreshTeamRoster();
  },

  // ══════════════════════════════════════════════
  // TEAM ROSTER (Scoped roster for Manager/JD)
  // ══════════════════════════════════════════════
  _teamRosterTeamName: null,
  _teamRosterStatusTab: 'active',  // 'active' or 'deactivated'

  _renderTeamRoster() {
    const page = document.getElementById('team-roster-page');
    if (!page) return;
    page.style.display = 'block';

    const persona = this.state.currentPersona;
    const myTeamName = Roster.getEffectiveTeam(persona, this.state.people);
    this._teamRosterTeamName = myTeamName;

    // Reset sub-tab to Active
    this._teamRosterStatusTab = 'active';
    const tabBtnActive = document.getElementById('team-roster-tab-active');
    const tabBtnDeactivated = document.getElementById('team-roster-tab-deactivated');
    if (tabBtnActive) tabBtnActive.classList.add('active');
    if (tabBtnDeactivated) tabBtnDeactivated.classList.remove('active');

    const team = this.state.teams.find(t => t.name === myTeamName);
    const subtitle = document.getElementById('team-roster-subtitle');
    const title = document.getElementById('team-roster-title');

    if (!team) {
      if (title) title.textContent = 'Roster';
      if (subtitle) subtitle.textContent = 'No team assigned';
      return;
    }

    // Title with emoji + name
    const display = Roster.getTeamDisplay(myTeamName, this.state.people, this.state.teams);
    if (title) title.textContent = (display.emoji || '') + ' ' + display.name + ' — Roster';

    const members = team.members || [];
    const activeCount = members.filter(p => !Roster.deactivated.has(p.name)).length;
    if (subtitle) subtitle.textContent = `${activeCount} active member${activeCount !== 1 ? 's' : ''}`;

    // Clear search
    const search = document.getElementById('team-roster-search');
    if (search) search.value = '';

    // Render only active members by default
    const activeMembers = members.filter(p => !Roster.deactivated.has(p.name));
    this._renderTeamRosterRows(activeMembers);
  },

  // ── Team Roster sub-tab toggle ──
  setTeamRosterStatusTab(tab) {
    this._teamRosterStatusTab = tab;
    const btnActive = document.getElementById('team-roster-tab-active');
    const btnDeactivated = document.getElementById('team-roster-tab-deactivated');
    if (btnActive) btnActive.classList.toggle('active', tab === 'active');
    if (btnDeactivated) btnDeactivated.classList.toggle('active', tab === 'deactivated');
    this.filterTeamRoster();
  },

  filterTeamRoster() {
    const myTeamName = this._teamRosterTeamName;
    if (!myTeamName) return;
    const team = this.state.teams.find(t => t.name === myTeamName);
    if (!team) return;

    const members = team.members || [];
    const search = (document.getElementById('team-roster-search')?.value || '').toLowerCase().trim();
    const statusTab = this._teamRosterStatusTab || 'active';

    let filtered = members.filter(p => {
      if (statusTab === 'active' && Roster.deactivated.has(p.name)) return false;
      if (statusTab === 'deactivated' && !Roster.deactivated.has(p.name)) return false;
      if (search && !p.name.toLowerCase().includes(search)) return false;
      return true;
    });

    // Total in current tab (before search filter)
    const totalInTab = members.filter(p =>
      statusTab === 'active' ? !Roster.deactivated.has(p.name) : Roster.deactivated.has(p.name)
    ).length;

    const subtitle = document.getElementById('team-roster-subtitle');
    if (subtitle) {
      subtitle.textContent = statusTab === 'active'
        ? `${totalInTab} active member${totalInTab !== 1 ? 's' : ''}`
        : `${totalInTab} deactivated member${totalInTab !== 1 ? 's' : ''}`;
    }

    const countEl = document.getElementById('team-roster-count');
    if (countEl) {
      countEl.textContent = filtered.length === totalInTab
        ? `Showing all ${totalInTab}`
        : `Showing ${filtered.length} of ${totalInTab}`;
    }

    this._renderTeamRosterRows(filtered);
  },

  _renderTeamRosterRows(filtered) {
    const tbody = document.getElementById('team-roster-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (filtered.length === 0) {
      tbody.innerHTML = '<tr><td colspan="3" style="text-align:center;color:var(--silver-dim);padding:32px;font-family:\'Cerebri Sans\',\'DM Sans\',\'Inter\',sans-serif;font-size:14px">No members found</td></tr>';
      return;
    }

    // Role permission: current user's rank determines what they can assign
    const myRole = this.state.currentRole || 'rep';
    const myRank = OFFICE_CONFIG.roles[myRole]?.rank || 0;

    // Sort alphabetically (sub-tabs already separate active/deactivated)
    const sorted = [...filtered].sort((a, b) => a.name.localeCompare(b.name));

    sorted.forEach(p => {
      const isDeactivated = Roster.deactivated.has(p.name);
      const safeName = p.name.replace(/'/g, "\\'");
      const email = p.email || Roster.getEmail(p.name) || '';
      const roleKey = p._roleKey || 'rep';

      // Build per-row role cell based on permission hierarchy
      const targetRank = OFFICE_CONFIG.roles[roleKey]?.rank || 0;
      const canChangeRole = myRank > targetRank || myRole === 'superadmin';

      let roleCell;
      if (!canChangeRole) {
        const roleLabel = OFFICE_CONFIG.roles[roleKey]?.label || roleKey;
        roleCell = `<span style="font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:12px;color:var(--silver-dim);padding:5px 0">${roleLabel}</span>`;
      } else {
        const roleOptions = Object.entries(OFFICE_CONFIG.roles)
          .filter(([key, val]) => {
            if (key === 'superadmin') return false;
            if (key === 'owner') return false;
            if (key === 'admin') return myRole === 'owner' || myRole === 'admin' || myRole === 'superadmin';
            return val.rank < myRank || myRole === 'superadmin';
          })
          .map(([key, val]) => ({ key, label: val.label }));
        const roleSelect = roleOptions.map(r =>
          `<option value="${r.key}"${roleKey === r.key ? ' selected' : ''}>${r.label}</option>`
        ).join('');
        roleCell = `<select onchange="App.setPersonRole('${safeName}',this.value)"
            style="background:rgba(0,0,0,0.06);border:1px solid rgba(0,0,0,0.25);border-radius:6px;color:var(--white);padding:5px 8px;font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:12px;cursor:pointer;outline:none">
            ${roleSelect}
          </select>`;
      }

      const tr = document.createElement('tr');
      tr.style.cssText = `border-bottom:1px solid rgba(0,0,0,0.06);`;
      tr.innerHTML = `
        <td style="padding:12px 16px">
          <div style="font-family:'Neue Montreal','Inter',sans-serif;font-size:15px;font-weight:700;color:var(--white)">
            ${p.name}
          </div>
          ${email ? `<div style="font-size:10px;color:var(--silver-dim);margin-top:2px">${email}</div>` : ''}
        </td>
        <td style="padding:12px 16px">
          ${roleCell}
        </td>
        <td style="padding:12px 16px;text-align:center">
          <button onclick="App.toggleDeactivate('${safeName}')"
            style="background:${isDeactivated ? 'rgba(46,139,87,0.1)' : 'rgba(229,86,74,0.1)'};border:1px solid ${isDeactivated ? 'rgba(46,139,87,0.3)' : 'rgba(229,86,74,0.3)'};border-radius:6px;color:${isDeactivated ? '#2E8B57' : '#E5564A'};padding:5px 14px;font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;cursor:pointer;text-transform:uppercase">
            ${isDeactivated ? 'Reactivate' : 'Deactivate'}
          </button>
        </td>`;
      tbody.appendChild(tr);
    });
  },

  _refreshTeamRoster() {
    const page = document.getElementById('team-roster-page');
    if (page && page.style.display !== 'none' && this._teamRosterTeamName) {
      this.filterTeamRoster();
    }
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
    const fields = ['add-member-email', 'add-member-name', 'add-member-phone', 'add-member-error'];
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
    const phone = document.getElementById('add-member-phone')?.value?.trim() || '';
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
      await Roster.addNewPerson(email, name, team || 'Unassigned', rank || 'rep', phone, OFFICE_CONFIG);
      // Update local state
      this.state.roster[email] = {
        name,
        team: team || 'Unassigned',
        rank: rank || 'rep',
        phone: phone || '',
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
    if (status) status.innerHTML = `<div style="color:#2E8B57;font-size:13px;margin-bottom:16px">✓ Request sent!</div>`;
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
    + `<span onclick="App._rerollEmojis('${pickerId}','${displayId}','${selectFn}')" style="cursor:pointer;font-size:18px;padding:4px 8px;border-radius:8px;background:rgba(0,0,0,0.06);border:1px solid rgba(0,0,0,0.2);color:var(--silver-dim);font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-weight:700;letter-spacing:1px;display:inline-flex;align-items:center;gap:4px" title="Show more emojis">🔄</span>`;
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
    if (el) { el.style.borderColor = 'var(--sc-cyan)'; el.style.background = 'rgba(0,0,0,0.2)'; }
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
          <div style="font-size:56px;margin-bottom:12px">😔</div>
          <div style="font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:24px;font-weight:700;color:var(--white);margin-bottom:6px">Something went wrong</div>
          <div style="color:var(--silver-dim);font-size:13px;max-width:400px;line-height:1.6;margin-bottom:28px">Don't worry — just tap retry and we'll get you back on track.</div>
          <button onclick="location.reload()" style="background:var(--blue-deep);border:none;border-radius:12px;padding:14px 36px;color:#fff;font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:15px;font-weight:700;letter-spacing:0.3px;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,0.15)">Retry</button>
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
