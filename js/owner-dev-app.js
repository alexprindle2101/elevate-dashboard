// ═══════════════════════════════════════════════════════
// Owner Development Dashboard — App Controller
// Multi-team owner mapping tool: Maddie's, Cam's, NLR
// ═══════════════════════════════════════════════════════

const OwnerDev = {

  // ── State ──
  state: {
    session: null,          // { email, name, team, role, loginTime }
    campaigns: {},          // { 'frontier': { label, owners: ['name1', ...] }, ... }
    mappings: [],           // [{ campaign, ownerName, camCompany, nlrWorkbookId, nlrWorkbookName, nlrTab }, ...]
    users: [],              // [{ email, name, team, role }, ...]
    camCompanies: [],       // ['Company1', ...] for Cam's autocomplete
    nlrWorkbooks: [],       // [{ id, name }, ...] from NLR Drive folder
    nlrTabsCache: {},       // { sheetId: ['Tab1', ...] }
    campaignTabMap: [],     // [{ campaign, ownerName, tabName }] saved source tab overrides
    campaignTabsCache: {},  // { campaignKey: ['Tab1', ...] } available tabs per campaign
    activeCampaign: 'all',  // filter key
    searchQuery: '',        // search filter
    activeTab: 'mapping',   // 'mapping' | 'team'
    sortCol: 'campaign',    // current sort column
    sortAsc: true,          // ascending sort
    _savingCells: new Set(),// tracks cells currently saving (prevents double-submit)

    // View-As state (superadmin only)
    viewAsTeam: null,       // which team we're impersonating (null = own team)
    isSuperadmin: false,    // is the logged-in user a superadmin?
    realTeam: null,         // actual logged-in team (preserved during View-As)

    // Planning tab state
    planningData: [],       // [{ day, sortOrder, campaignKey, ownerOrder }, ...]
    planningLoaded: false,  // whether planning data has been fetched
    _planningDirty: false   // whether unsaved changes exist
  },

  // ══════════════════════════════════════════════════════
  // INIT
  // ══════════════════════════════════════════════════════

  async init() {
    console.log('[OwnerDev] init');
    this.state.session = this._getSession();

    if (!this.state.session) {
      this._showLogin();
      return;
    }

    // Check superadmin status
    const email = (this.state.session.email || '').toLowerCase();
    this.state.isSuperadmin = (OD_CONFIG.superadmins || []).some(e => e.toLowerCase() === email);
    this.state.realTeam = this.state.session.team;
    const role = (this.state.session.role || '').toLowerCase();
    this.state.isCoachOnly = role === 'coach';

    // Populate user info in top bar
    this._renderUserInfo();

    // Coach-only mode: skip mapping/team, go straight to coach view
    if (this.state.isCoachOnly) {
      // Hide all nav tabs except coach
      document.querySelectorAll('.nav-tab').forEach(t => t.style.display = 'none');
      const coachTab = document.getElementById('nav-coach-tab');
      if (coachTab) { coachTab.style.display = ''; coachTab.classList.add('active'); }
      // Hide mapping view, show coach view
      document.getElementById('view-mapping').style.display = 'none';
      const coachView = document.getElementById('view-coach');
      if (coachView) coachView.style.display = '';
      document.getElementById('dashboard').style.display = 'block';
      this.state.activeTab = 'coach';
      if (typeof NationalApp !== 'undefined') {
        NationalApp.initCoachView({ email: this.state.session.email, name: this.state.session.name });
      }
      return;
    }

    // Show team tab if manager or superadmin
    if (role === 'manager' || this.state.isSuperadmin) {
      document.getElementById('nav-team-tab').style.display = '';
    }

    // Show coach tab for superadmins or managers
    if (this.state.isSuperadmin || role === 'manager') {
      const coachTab = document.getElementById('nav-coach-tab');
      if (coachTab) coachTab.style.display = '';
    }

    // Show planning tab for Maddie's team managers + superadmins
    const initTeam = (this.state.session.team || '').toLowerCase();
    if ((initTeam === 'maddie' && role === 'manager') || this.state.isSuperadmin) {
      const planTab = document.getElementById('nav-planning-tab');
      if (planTab) planTab.style.display = '';
    }

    // Build View-As bar for superadmins
    if (this.state.isSuperadmin) {
      this._buildViewAsBar();
    }

    // Load data (cache-first: shows cached data instantly, refreshes in background)
    const cache = this._readCache();
    const hasCache = cache?.campaigns && Object.values(cache.campaigns).some(c => c.owners?.length > 0);
    if (!hasCache) this._showLoading('Loading owner data...');
    try {
      await this._loadAllData();
    } catch (err) {
      console.error('[OwnerDev] Data load failed:', err);
      this._toast('Failed to load data. Please refresh.', 'error');
    }

    // If cache was used, _loadAllData applied it synchronously — show dashboard + render
    // If no cache, _refreshFromServer handles showing dashboard after owners arrive
    if (hasCache) {
      document.getElementById('dashboard').style.display = 'block';
      this._renderFilterPills();
      this._renderStats();
      this.renderMapping();
    }
  },

  // ══════════════════════════════════════════════════════
  // SESSION
  // ══════════════════════════════════════════════════════

  _getSession() {
    try {
      const raw = localStorage.getItem(OD_CONFIG.sessionKey);
      if (!raw) return null;
      const s = JSON.parse(raw);
      if (Date.now() - s.loginTime > OD_CONFIG.sessionDuration) {
        localStorage.removeItem(OD_CONFIG.sessionKey);
        return null;
      }
      return s;
    } catch { return null; }
  },

  _saveSession(data) {
    const s = { ...data, loginTime: Date.now() };
    localStorage.setItem(OD_CONFIG.sessionKey, JSON.stringify(s));
    return s;
  },

  logout() {
    localStorage.removeItem(OD_CONFIG.sessionKey);
    window.location.reload();
  },

  // ══════════════════════════════════════════════════════
  // LOGIN (two-step: email → PIN)
  // ══════════════════════════════════════════════════════

  _showLogin() {
    const screen = document.getElementById('login-screen');
    screen.style.display = 'flex';

    const emailWrap = document.getElementById('login-email-wrap');
    const pinWrap = document.getElementById('login-pin-wrap');
    const pinCreateWrap = document.getElementById('login-pin-create-wrap');
    const emailInput = document.getElementById('login-email');
    const pinInput = document.getElementById('login-pin');
    const pinNew = document.getElementById('login-pin-new');
    const pinConfirm = document.getElementById('login-pin-confirm');
    const btn = document.getElementById('login-btn');
    const error = document.getElementById('login-error');
    const backLink = document.getElementById('login-back-link');

    let loginStep = 'email'; // 'email' | 'pin' | 'pin-create'
    let loginEmail = '';

    // Reset state
    emailWrap.style.display = '';
    pinWrap.style.display = 'none';
    pinCreateWrap.style.display = 'none';
    backLink.style.display = 'none';
    error.textContent = '';
    btn.textContent = 'Continue';
    btn.disabled = false;

    const handleSubmit = async () => {
      error.textContent = '';
      btn.disabled = true;

      try {
        if (loginStep === 'email') {
          // ── Email step: verify user exists ──
          loginEmail = emailInput.value.trim().toLowerCase();
          if (!loginEmail) { error.textContent = 'Please enter your email'; btn.disabled = false; return; }

          // Resolve login aliases (e.g. 'alex' → 'alex.aspirehr@gmail.com')
          if (OD_CONFIG.loginAliases && OD_CONFIG.loginAliases[loginEmail]) {
            loginEmail = OD_CONFIG.loginAliases[loginEmail].toLowerCase();
          }

          // Check if superadmin — skip _OD_Users lookup but still require PIN
          const isSA = (OD_CONFIG.superadmins || []).some(e => e.toLowerCase() === loginEmail);

          if (!isSA) {
            // Normal users: verify they exist in _OD_Users
            const res = await this._post('odCheckUser', { email: loginEmail });
            if (!res.success) {
              error.textContent = res.message || 'Contact your team manager to get added.';
              btn.disabled = false;
              return;
            }
            this._loginHasPin = res.hasPin;
          } else {
            // Superadmins: check if they have a PIN set already via odCheckUser
            // (may not exist in _OD_Users yet — that's fine, treat as first-time)
            try {
              const res = await this._post('odCheckUser', { email: loginEmail });
              this._loginHasPin = res.success && res.hasPin;
            } catch {
              this._loginHasPin = false;
            }
          }
          this._loginIsSA = isSA;

          // Move to PIN step
          emailWrap.style.display = 'none';
          backLink.style.display = '';

          if (this._loginHasPin) {
            loginStep = 'pin';
            pinWrap.style.display = '';
            btn.textContent = 'Sign In';
            setTimeout(() => pinInput.focus(), 100);
          } else {
            loginStep = 'pin-create';
            pinCreateWrap.style.display = '';
            btn.textContent = 'Create PIN & Sign In';
            setTimeout(() => pinNew.focus(), 100);
          }
          btn.disabled = false;

        } else if (loginStep === 'pin') {
          // ── PIN step: authenticate ──
          const pin = pinInput.value.trim();
          if (!pin || pin.length < 4) { error.textContent = 'Enter your 4-6 digit PIN'; btn.disabled = false; return; }

          const res = await this._post('odLogin', { email: loginEmail, pin });
          if (!res.success) {
            error.textContent = res.message || 'Invalid PIN.';
            btn.disabled = false;
            return;
          }

          // Login success — superadmins get elevated session
          const saName = loginEmail.split('@')[0].replace(/[._]/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
          this.state.session = this._saveSession({
            email: res.email || loginEmail,
            name: res.name || saName,
            team: this._loginIsSA ? 'maddie' : res.team,
            role: this._loginIsSA ? 'manager' : (res.role || 'member')
          });
          screen.style.display = 'none';
          this.init();

        } else if (loginStep === 'pin-create') {
          // ── Create PIN step ──
          const pin1 = pinNew.value.trim();
          const pin2 = pinConfirm.value.trim();
          if (!pin1 || pin1.length < 4) { error.textContent = 'PIN must be 4-6 digits'; btn.disabled = false; return; }
          if (pin1 !== pin2) { error.textContent = 'PINs do not match'; btn.disabled = false; return; }

          if (this._loginIsSA) {
            // Superadmin first login: create user in _OD_Users via odSaveUser, then login
            await this._post('odSaveUser', {
              email: loginEmail,
              name: loginEmail.split('@')[0].replace(/[._]/g, ' ').replace(/\b\w/g, c => c.toUpperCase()),
              team: 'maddie',
              role: 'manager',
              pin: pin1
            });
          }

          const res = await this._post('odLogin', { email: loginEmail, pin: pin1, createPin: true });
          if (!res.success) {
            error.textContent = res.message || 'Failed to create PIN.';
            btn.disabled = false;
            return;
          }

          // Login success
          const saName2 = loginEmail.split('@')[0].replace(/[._]/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
          this.state.session = this._saveSession({
            email: res.email || loginEmail,
            name: res.name || saName2,
            team: this._loginIsSA ? 'maddie' : res.team,
            role: this._loginIsSA ? 'manager' : (res.role || 'member')
          });
          screen.style.display = 'none';
          this.init();
        }
      } catch (err) {
        console.error('[OwnerDev] Login error:', err);
        error.textContent = 'Connection failed. Please try again.';
        btn.disabled = false;
      }
    };

    // Back link: go back to email step
    backLink.onclick = () => {
      loginStep = 'email';
      emailWrap.style.display = '';
      pinWrap.style.display = 'none';
      pinCreateWrap.style.display = 'none';
      backLink.style.display = 'none';
      btn.textContent = 'Continue';
      btn.disabled = false;
      error.textContent = '';
      pinInput.value = '';
      pinNew.value = '';
      pinConfirm.value = '';
      setTimeout(() => emailInput.focus(), 100);
    };

    // Button click + Enter key
    btn.onclick = handleSubmit;
    emailInput.onkeydown = e => { if (e.key === 'Enter') handleSubmit(); };
    pinInput.onkeydown = e => { if (e.key === 'Enter') handleSubmit(); };
    pinConfirm.onkeydown = e => { if (e.key === 'Enter') handleSubmit(); };

    setTimeout(() => emailInput.focus(), 100);
  },

  // ══════════════════════════════════════════════════════
  // API LAYER
  // ══════════════════════════════════════════════════════

  /**
   * GET request to Apps Script
   * @param {string} action - API action name
   * @param {Object} params - Additional query parameters
   * @returns {Promise<Object>} parsed JSON response
   */
  async _api(action, params = {}, timeout) {
    const url = new URL(OD_CONFIG.appsScriptUrl);
    url.searchParams.set('key', OD_CONFIG.apiKey);
    url.searchParams.set('action', action);
    for (const [k, v] of Object.entries(params)) {
      url.searchParams.set(k, v);
    }
    const res = await this._fetchWithTimeout(
      fetch(url.toString()).then(r => r.json()),
      timeout
    );
    return res;
  },

  /**
   * POST request to Apps Script (text/plain to avoid CORS preflight)
   * @param {string} action - API action name
   * @param {Object} body - Data to send
   * @returns {Promise<Object>} parsed JSON response
   */
  async _post(action, body = {}) {
    const res = await this._fetchWithTimeout(
      fetch(OD_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({ action, key: OD_CONFIG.apiKey, ...body })
      }).then(r => r.json())
    );
    return res;
  },

  /**
   * Fetch with timeout wrapper (default 30s)
   */
  _fetchWithTimeout(promise, ms = 30000) {
    return Promise.race([
      promise,
      new Promise((_, reject) => setTimeout(() => reject(new Error('Request timeout')), ms))
    ]);
  },

  // ══════════════════════════════════════════════════════
  // DATA CACHE (localStorage, stale-while-revalidate)
  // ══════════════════════════════════════════════════════

  _CACHE_KEY: 'od_data_cache',
  _CACHE_MAX_AGE: 10 * 60 * 1000, // 10 min — treat as fresh, skip background refresh

  /**
   * Read the full data cache from localStorage
   */
  _readCache() {
    try {
      const raw = localStorage.getItem(this._CACHE_KEY);
      if (!raw) return null;
      return JSON.parse(raw);
    } catch { return null; }
  },

  /**
   * Write all cacheable data to localStorage
   */
  _writeCache(data) {
    try {
      data._ts = Date.now();
      localStorage.setItem(this._CACHE_KEY, JSON.stringify(data));
    } catch (err) {
      console.warn('[OwnerDev] Cache write failed:', err.message);
    }
  },

  /**
   * Apply cached data to state (same logic as processing API results)
   */
  _applyCacheToState(cache) {
    if (cache.campaigns) {
      this.state.campaigns = cache.campaigns;
      for (const [key, camp] of Object.entries(this.state.campaigns)) {
        if (camp.tabs && camp.tabs.length) {
          this.state.campaignTabsCache[key] = camp.tabs;
        }
      }
    }
    if (cache.mappings) this.state.mappings = cache.mappings;
    if (cache.users) this.state.users = cache.users;
    if (cache.camCompanies) this.state.camCompanies = cache.camCompanies;
    if (cache.clientToBusinessMap) this.state.clientToBusinessMap = cache.clientToBusinessMap;
    if (cache.nlrWorkbooks) {
      this.state.nlrWorkbooks = cache.nlrWorkbooks;
      for (const wb of this.state.nlrWorkbooks) {
        if (wb.tabs && wb.tabs.length) this.state.nlrTabsCache[wb.id] = wb.tabs;
      }
    }
    if (cache.campaignTabMap) this.state.campaignTabMap = cache.campaignTabMap;
    if (cache.campaignTabs) {
      for (const [key, tabs] of Object.entries(cache.campaignTabs)) {
        if (tabs.length) this.state.campaignTabsCache[key] = tabs;
      }
    }
  },

  // ══════════════════════════════════════════════════════
  // DATA LOADING (stale-while-revalidate)
  // ══════════════════════════════════════════════════════

  /**
   * Load data with cache-first strategy:
   * 1. If cache exists → render immediately from cache
   * 2. Fetch fresh data in background (or foreground if no cache)
   * 3. Update UI when fresh data arrives
   */
  async _loadAllData() {
    const cache = this._readCache();
    const hasCache = cache && cache.campaigns && Object.values(cache.campaigns).some(c => c.owners?.length > 0);
    const isFresh = hasCache && (Date.now() - (cache._ts || 0)) < this._CACHE_MAX_AGE;

    // ── Step 1: If we have cached data, apply it immediately ──
    if (hasCache) {
      console.log(`[OwnerDev] Rendering from cache (age: ${Math.round((Date.now() - (cache._ts || 0)) / 1000)}s)`);
      this._applyCacheToState(cache);
    }

    // ── Step 2: If cache is fresh enough, skip server fetch ──
    if (isFresh) {
      console.log('[OwnerDev] Cache is fresh, skipping server fetch');
      return;
    }

    // ── Step 3: Fetch fresh data (foreground if no cache, background if cached) ──
    if (hasCache) {
      // Background refresh — don't block the UI
      this._refreshFromServer();
    } else {
      // No cache — must wait for server
      await this._refreshFromServer();
    }
  },

  /**
   * Fetch all data from the server and update state + cache + UI.
   * Uses progressive loading: renders owners as soon as they arrive
   * instead of waiting for all 6 API calls to finish.
   */
  async _refreshFromServer() {
    console.log('[OwnerDev] Fetching fresh data from server...');
    const isFirstLoad = !Object.values(this.state.campaigns).some(c => c.owners?.length > 0);

    // ── Fire all requests in parallel ──
    const campaignOwners = this._api('odCampaignOwners', {}, 45000);
    const mappings       = this._api('odGetMappings');
    const users          = this._api('odGetUsers');
    const camCompanies   = this._api('odCamCompanies', {}, 30000);
    const nlrWorkbooks   = this._api('odNlrWorkbooks', {}, 30000);
    const campaignTabMap = this._api('odGetCampaignTabMap');

    const cacheData = {};

    // ── Priority 1: Campaign owners + mappings (needed to render the table) ──
    // Wait for these two first so we can show the table ASAP on first load
    const [ownersResult, mappingsResult] = await Promise.allSettled([campaignOwners, mappings]);

    if (ownersResult.status === 'fulfilled' && ownersResult.value.success) {
      this.state.campaigns = ownersResult.value.campaigns || {};
      cacheData.campaigns = this.state.campaigns;
      for (const [key, camp] of Object.entries(this.state.campaigns)) {
        if (camp.tabs && camp.tabs.length) this.state.campaignTabsCache[key] = camp.tabs;
      }
    } else {
      const reason = ownersResult.status === 'rejected'
        ? (ownersResult.reason?.message || 'Request timeout')
        : (ownersResult.value?.error || 'Unknown error');
      console.error('[OwnerDev] Failed to load campaign owners:', reason);
      if (isFirstLoad) {
        this._toast('Failed to load owner data: ' + reason, 'error');
        this.state.campaigns = {};
        for (const [key, cfg] of Object.entries(OD_CONFIG.campaignSources)) {
          this.state.campaigns[key] = { label: cfg.label, owners: [] };
        }
        this._showRetryBanner();
      }
    }

    if (mappingsResult.status === 'fulfilled' && mappingsResult.value.success) {
      this.state.mappings = mappingsResult.value.mappings || [];
      cacheData.mappings = this.state.mappings;
    } else if (!this.state.mappings.length) {
      this.state.mappings = [];
    }

    // ── Render immediately with owners + mappings ──
    if (isFirstLoad) {
      this._hideLoading();
      document.getElementById('dashboard').style.display = 'block';
    }
    this._renderFilterPills();
    this._renderStats();
    this.renderMapping();

    // ── Priority 2: Remaining data (settles in background) ──
    const [usersResult, camResult, nlrResult, tabMapResult] =
      await Promise.allSettled([users, camCompanies, nlrWorkbooks, campaignTabMap]);

    if (usersResult.status === 'fulfilled' && usersResult.value.success) {
      this.state.users = usersResult.value.users || [];
      cacheData.users = this.state.users;
    } else if (!this.state.users.length) {
      this.state.users = [];
    }

    if (camResult.status === 'fulfilled' && camResult.value.success) {
      this.state.camCompanies = camResult.value.companies || [];
      this.state.clientToBusinessMap = camResult.value.clientToBusinessMap || [];
      cacheData.camCompanies = this.state.camCompanies;
      cacheData.clientToBusinessMap = this.state.clientToBusinessMap;
    } else if (!this.state.camCompanies.length) {
      this.state.camCompanies = [];
      this.state.clientToBusinessMap = [];
    }

    if (nlrResult.status === 'fulfilled' && nlrResult.value.success) {
      this.state.nlrWorkbooks = nlrResult.value.workbooks || [];
      cacheData.nlrWorkbooks = this.state.nlrWorkbooks;
      for (const wb of this.state.nlrWorkbooks) {
        if (wb.tabs && wb.tabs.length) this.state.nlrTabsCache[wb.id] = wb.tabs;
      }
    } else if (!this.state.nlrWorkbooks.length) {
      this.state.nlrWorkbooks = [];
    }

    if (tabMapResult.status === 'fulfilled' && tabMapResult.value.success) {
      this.state.campaignTabMap = tabMapResult.value.mappings || [];
      cacheData.campaignTabMap = this.state.campaignTabMap;
      const ct = tabMapResult.value.campaignTabs || {};
      cacheData.campaignTabs = ct;
      for (const [key, tabs] of Object.entries(ct)) {
        if (tabs.length) this.state.campaignTabsCache[key] = tabs;
      }
    } else if (!this.state.campaignTabMap.length) {
      this.state.campaignTabMap = [];
    }

    console.log('[OwnerDev] Data loaded:', {
      campaigns: Object.keys(this.state.campaigns).length,
      mappings: this.state.mappings.length,
      users: this.state.users.length,
      camCompanies: this.state.camCompanies.length,
      nlrWorkbooks: this.state.nlrWorkbooks.length,
      clientToBusinessPairs: this.state.clientToBusinessMap.length,
      campaignTabMaps: this.state.campaignTabMap.length,
      campaignTabsCached: Object.keys(this.state.campaignTabsCache).length
    });

    // Write to cache if we got campaign data
    if (cacheData.campaigns) {
      this._writeCache(cacheData);
    }

    // Run auto-mapping
    const hasOwners = Object.values(this.state.campaigns).some(c => c.owners?.length > 0);
    if (hasOwners) {
      if (this.state.clientToBusinessMap.length > 0) await this._autoMapCamCompanies();
      if (this.state.nlrWorkbooks.length > 0) await this._autoMapNlrFiles();
      if (Object.keys(this.state.campaignTabsCache).length > 0) await this._autoMapCampaignTabs();
    }

    // Final re-render with all data (auto-mapping may have updated statuses)
    this._renderFilterPills();
    this._renderStats();
    this.renderMapping();
  },

  /**
   * Show a retry banner when data fails to load
   */
  _showRetryBanner() {
    const existing = document.getElementById('retry-banner');
    if (existing) existing.remove();

    const banner = document.createElement('div');
    banner.id = 'retry-banner';
    banner.style.cssText = `
      background:#fef3cd; border:1px solid #ffc107; border-radius:8px;
      padding:16px 24px; margin:16px auto; max-width:600px;
      display:flex; align-items:center; gap:12px; font-size:14px;
      color:#856404; font-family:var(--font,Inter,sans-serif);
    `;
    banner.innerHTML = `
      <span style="font-size:20px">&#x26A0;&#xFE0F;</span>
      <span style="flex:1">Owner data failed to load. The server may be slow — please try again.</span>
      <button onclick="OwnerDev.retryLoad()" style="
        background:#0099cc; color:#fff; border:none; border-radius:6px;
        padding:8px 20px; font-size:13px; font-weight:600; cursor:pointer;
      ">Retry</button>
    `;

    const content = document.getElementById('view-mapping');
    content.insertBefore(banner, content.firstChild);
  },

  /**
   * Retry loading all data (force server fetch, bypass cache)
   */
  async retryLoad() {
    const banner = document.getElementById('retry-banner');
    if (banner) banner.remove();

    this._showLoading('Retrying data load...');
    try {
      await this._refreshFromServer();
    } catch (err) {
      console.error('[OwnerDev] Retry failed:', err);
      this._toast('Retry failed. Please refresh the page.', 'error');
    }
    this._hideLoading();
  },

  // ══════════════════════════════════════════════════════
  // AUTO-MAPPING (Cam's companies by name similarity)
  // ══════════════════════════════════════════════════════

  /**
   * Normalize a name for fuzzy comparison: lowercase, strip punctuation, collapse whitespace
   */
  _normName(name) {
    return (name || '').toLowerCase().replace(/[^a-z0-9\s]/g, '').replace(/\s+/g, ' ').trim();
  },

  /**
   * Check if two names are a fuzzy match.
   * Matches if: exact, one contains the other, or first+last name tokens overlap.
   */
  _namesMatch(ownerName, clientName) {
    const a = this._normName(ownerName);
    const b = this._normName(clientName);
    if (!a || !b) return false;

    // Exact
    if (a === b) return true;

    // One contains the other
    if (a.includes(b) || b.includes(a)) return true;

    // Token overlap: if first and last name both match
    const aToks = a.split(' ');
    const bToks = b.split(' ');
    if (aToks.length >= 2 && bToks.length >= 2) {
      const aFirst = aToks[0], aLast = aToks[aToks.length - 1];
      const bFirst = bToks[0], bLast = bToks[bToks.length - 1];
      if (aFirst === bFirst && aLast === bLast) return true;
    }

    // First name + last initial (e.g. "Jay T" matches "Jay Thurston")
    if (aToks.length >= 1 && bToks.length >= 2) {
      if (aToks[0] === bToks[0] && aToks.length === 2 && aToks[1].length === 1 && bToks[bToks.length - 1].startsWith(aToks[1])) return true;
    }
    if (bToks.length >= 1 && aToks.length >= 2) {
      if (bToks[0] === aToks[0] && bToks.length === 2 && bToks[1].length === 1 && aToks[aToks.length - 1].startsWith(bToks[1])) return true;
    }

    return false;
  },

  /**
   * Auto-map unmapped owners to Cam's companies by matching owner name → client name.
   * Only fills in camCompany where there's no existing mapping.
   * Returns count of auto-mapped owners.
   */
  async _autoMapCamCompanies() {
    const pairs = this.state.clientToBusinessMap;
    if (!pairs.length) return 0;

    let autoMapped = 0;
    const toSave = [];

    // For each owner across all campaigns, check if they match a client name
    for (const [campaignKey, campaign] of Object.entries(this.state.campaigns)) {
      for (const ownerName of (campaign.owners || [])) {
        const existing = this._findMapping(campaignKey, ownerName);
        if (existing?.camCompany) continue; // already mapped

        // Find best match from client→business pairs
        let match = null;
        for (const pair of pairs) {
          if (this._namesMatch(ownerName, pair.clientName)) {
            match = pair.businessName;
            break;
          }
        }

        if (match) {
          // Update local state immediately
          this._upsertMapping(campaignKey, ownerName, { camCompany: match });
          toSave.push({ campaign: campaignKey, ownerName, value: match });
          autoMapped++;
        }
      }
    }

    // Batch-save all auto-mapped values in a single POST
    if (toSave.length > 0) {
      console.log(`[OwnerDev] Auto-mapped ${toSave.length} owners to Cam companies`);
      this._toast(`Auto-mapped ${toSave.length} owner${toSave.length > 1 ? 's' : ''} to companies`, 'success');

      const mappings = toSave.map(item => ({
        campaign: item.campaign,
        ownerName: item.ownerName,
        camCompany: item.value,
        updatedBy: 'auto-map'
      }));
      this._post('odBatchSaveMappings', { mappings })
        .then(res => console.log('[OwnerDev] Batch save Cam result:', res))
        .catch(err => console.warn('[OwnerDev] Batch save Cam error:', err.message));
    }

    return autoMapped;
  },

  /**
   * Auto-map NLR workbooks + tabs by matching owner names to file names.
   * If a workbook name fuzzy-matches an owner, assign it.
   * If the workbook has a tab called "Indeed Tracking 2026", auto-select it.
   */
  async _autoMapNlrFiles() {
    const workbooks = this.state.nlrWorkbooks;
    if (!workbooks.length) { console.log('[OwnerDev] NLR auto-map: no workbooks loaded'); return 0; }

    console.log('[OwnerDev] NLR auto-map: checking', workbooks.length, 'workbooks:', workbooks.map(w => w.name));

    // ── Step 1: Match owner names to workbook filenames ──
    const matched = []; // { campaignKey, ownerName, wb }
    const unmatched = [];

    for (const [campaignKey, campaign] of Object.entries(this.state.campaigns)) {
      for (const ownerName of (campaign.owners || [])) {
        const existing = this._findMapping(campaignKey, ownerName);
        if (existing?.nlrWorkbookId) continue; // already mapped

        let matchedWb = null;
        for (const wb of workbooks) {
          if (this._namesMatch(ownerName, wb.name)) {
            matchedWb = wb;
            break;
          }
          const normOwner = this._normName(ownerName);
          const normWb = this._normName(wb.name);
          if (normOwner && normWb) {
            const ownerTokens = normOwner.split(' ').filter(t => t.length >= 2);
            const matchCount = ownerTokens.filter(t => normWb.includes(t)).length;
            if (matchCount >= 2 || (ownerTokens.length === 1 && matchCount === 1 && normWb.includes(normOwner))) {
              matchedWb = wb;
              break;
            }
          }
        }

        if (matchedWb) {
          matched.push({ campaignKey, ownerName, wb: matchedWb });
        } else {
          unmatched.push(ownerName);
        }
      }
    }

    if (unmatched.length > 0) {
      console.log(`[OwnerDev] NLR unmatched owners (${unmatched.length}):`, unmatched.slice(0, 20));
    }
    if (!matched.length) { console.log('[OwnerDev] NLR auto-map: no new file matches'); return 0; }

    console.log(`[OwnerDev] NLR auto-map: ${matched.length} file matches, fetching tabs...`);

    // ── Step 2: Fetch tabs only for matched workbooks (in parallel, by unique ID) ──
    const uniqueIds = [...new Set(matched.map(m => m.wb.id))];
    const idsToFetch = uniqueIds.filter(id => !this.state.nlrTabsCache[id]);

    if (idsToFetch.length > 0) {
      const tabResults = await Promise.allSettled(
        idsToFetch.map(id =>
          this._api('odNlrTabs', { sheetId: id }).then(res => ({ id, tabs: res.success ? (res.tabs || []) : [] }))
        )
      );
      for (const result of tabResults) {
        if (result.status === 'fulfilled') {
          this.state.nlrTabsCache[result.value.id] = result.value.tabs;
        }
      }
      console.log(`[OwnerDev] Fetched tabs for ${tabResults.filter(r => r.status === 'fulfilled').length}/${idsToFetch.length} workbooks`);
    }

    // ── Step 3: Build mappings with auto-selected "Indeed Tracking" tab ──
    let autoMapped = 0;
    const toSave = [];

    for (const { campaignKey, ownerName, wb } of matched) {
      const tabs = this.state.nlrTabsCache[wb.id] || [];
      const indeedTab = tabs.find(t => t.toLowerCase().includes('indeed tracking'));
      const autoTab = indeedTab || '';

      this._upsertMapping(campaignKey, ownerName, {
        nlrWorkbookId: wb.id,
        nlrWorkbookName: wb.name,
        nlrTab: autoTab
      });

      toSave.push({
        campaign: campaignKey,
        ownerName,
        nlrWorkbookId: wb.id,
        nlrWorkbookName: wb.name,
        nlrTab: autoTab,
        updatedBy: 'auto-map'
      });
      autoMapped++;
    }

    // ── Step 4: Batch save to backend ──
    if (toSave.length > 0) {
      console.log(`[OwnerDev] Auto-mapped ${toSave.length} owners to NLR files (${toSave.filter(s => s.nlrTab).length} with tabs)`);
      this._post('odBatchSaveMappings', { mappings: toSave })
        .then(res => console.log('[OwnerDev] Batch save NLR result:', res))
        .catch(err => console.warn('[OwnerDev] Batch save NLR error:', err.message));
    }

    return autoMapped;
  },

  /**
   * Manual auto-map trigger (from the UI button) — runs both Cam + NLR
   */
  async runAutoMap() {
    const btn = document.getElementById('btn-automap');
    if (btn) { btn.disabled = true; btn.textContent = '⚡ Mapping...'; }

    const tabCount = await this._autoMapCampaignTabs();
    const camCount = await this._autoMapCamCompanies();
    const nlrCount = await this._autoMapNlrFiles();
    const total = tabCount + camCount + nlrCount;

    if (total === 0) {
      this._toast('No new matches found — remaining owners need manual mapping', 'error');
    } else {
      const parts = [];
      if (tabCount) parts.push(`${tabCount} source tab${tabCount > 1 ? 's' : ''}`);
      if (camCount) parts.push(`${camCount} compan${camCount > 1 ? 'ies' : 'y'}`);
      if (nlrCount) parts.push(`${nlrCount} NLR file${nlrCount > 1 ? 's' : ''}`);
      this._toast(`Auto-mapped ${parts.join(' + ')}`, 'success');
    }

    // Re-render everything
    this._renderStats();
    this.renderMapping();

    if (btn) { btn.disabled = false; btn.textContent = '⚡ Auto-Map'; }
  },

  // ══════════════════════════════════════════════════════
  // NAVIGATION
  // ══════════════════════════════════════════════════════

  /**
   * Switch between Mapping and Team tabs
   */
  switchTab(tab) {
    this.state.activeTab = tab;

    // Update nav tab active state
    document.querySelectorAll('.nav-tab').forEach(t => {
      t.classList.toggle('active', t.dataset.tab === tab);
    });

    // Show/hide views
    document.getElementById('view-mapping').style.display = tab === 'mapping' ? '' : 'none';
    const teamView = document.getElementById('view-team');
    teamView.classList.toggle('active', tab === 'team');
    const coachView = document.getElementById('view-coach');
    if (coachView) coachView.style.display = tab === 'coach' ? '' : 'none';
    const planningView = document.getElementById('view-planning');
    if (planningView) planningView.style.display = tab === 'planning' ? '' : 'none';

    if (tab === 'team') {
      this.renderTeam();
    }
    if (tab === 'coach') {
      if (typeof NationalApp !== 'undefined') {
        if (NationalApp._coachInitDone) {
          // Already initialized — go back to landing page
          NationalApp.backToLanding();
        } else {
          // First open — lazy-init
          NationalApp.initCoachView({
            email: this.state.session.email,
            name: this.state.session.name,
            campaign: (this.state.activeCampaign && this.state.activeCampaign !== 'all') ? this.state.activeCampaign : null
          });
        }
      }
    }
    if (tab === 'planning') {
      this._initPlanning();
    }
  },

  // ══════════════════════════════════════════════════════
  // FILTER & SEARCH
  // ══════════════════════════════════════════════════════

  /**
   * Set active campaign filter and re-render
   */
  filterCampaign(key) {
    this.state.activeCampaign = key;
    this._updateFilterPillsActive();
    this._renderStats();
    this.renderMapping();
  },

  /**
   * Filter by owner name search
   */
  filterSearch(query) {
    this.state.searchQuery = query.trim().toLowerCase();
    this._renderStats();
    this.renderMapping();
  },

  /**
   * Sort mapping table by column
   */
  sortBy(col) {
    if (this.state.sortCol === col) {
      this.state.sortAsc = !this.state.sortAsc;
    } else {
      this.state.sortCol = col;
      this.state.sortAsc = true;
    }
    this._updateSortArrows();
    this.renderMapping();
  },

  // ══════════════════════════════════════════════════════
  // RENDERING — User Info
  // ══════════════════════════════════════════════════════

  _renderUserInfo() {
    const s = this.state.session;
    if (!s) return;

    document.getElementById('user-name').textContent = s.name || s.email;

    const badge = document.getElementById('user-team-badge');
    const effectiveTeam = this._getEffectiveTeam();
    const teamCfg = OD_CONFIG.teams[effectiveTeam];
    if (teamCfg) {
      badge.style.display = '';
      badge.style.background = teamCfg.color;
      badge.textContent = teamCfg.icon + ' ' + teamCfg.label;
    }

    // If superadmin, show indicator
    if (this.state.isSuperadmin && this.state.viewAsTeam) {
      badge.textContent = '👁️ ' + badge.textContent;
    }
  },

  // ══════════════════════════════════════════════════════
  // RENDERING — Filter Pills
  // ══════════════════════════════════════════════════════

  _renderFilterPills() {
    const container = document.getElementById('filter-pills');
    const rows = this._getFilteredRows(true); // all rows (ignore campaign filter for counts)

    // Count owners per campaign
    const counts = {};
    let total = 0;
    for (const row of rows) {
      counts[row.campaign] = (counts[row.campaign] || 0) + 1;
      total++;
    }

    // Build pills
    let html = `<button class="filter-pill active" onclick="OwnerDev.filterCampaign('all')">All <span class="pill-count">${total}</span></button>`;

    for (const [key, cfg] of Object.entries(OD_CONFIG.campaignSources)) {
      const count = counts[key] || 0;
      if (count === 0 && !cfg.sheetId) continue; // skip unconfigured empty campaigns
      html += `<button class="filter-pill" data-campaign="${key}" onclick="OwnerDev.filterCampaign('${key}')">${cfg.label} <span class="pill-count">${count}</span></button>`;
    }

    container.innerHTML = html;
    this._updateFilterPillsActive();
  },

  _updateFilterPillsActive() {
    document.querySelectorAll('.filter-pill').forEach(pill => {
      const key = pill.dataset.campaign || 'all';
      pill.classList.toggle('active', key === this.state.activeCampaign);
    });
  },

  // ══════════════════════════════════════════════════════
  // RENDERING — Stats
  // ══════════════════════════════════════════════════════

  _renderStats() {
    const rows = this._getFilteredRows();
    let mapped = 0, partial = 0, unmapped = 0;

    for (const row of rows) {
      const s = this._getRowStatus(row);
      if (s === 'mapped') mapped++;
      else if (s === 'partial') partial++;
      else unmapped++;
    }

    document.getElementById('stat-total').textContent = rows.length;
    document.getElementById('stat-mapped').textContent = mapped;
    document.getElementById('stat-partial').textContent = partial;
    document.getElementById('stat-unmapped').textContent = unmapped;
  },

  // ══════════════════════════════════════════════════════
  // RENDERING — Mapping Table
  // ══════════════════════════════════════════════════════

  renderMapping() {
    const rows = this._getFilteredRows();
    const sorted = this._sortRows(rows);
    const tbody = document.getElementById('mapping-tbody');
    const empty = document.getElementById('mapping-empty');

    if (sorted.length === 0) {
      tbody.innerHTML = '';
      empty.style.display = '';
      return;
    }
    empty.style.display = 'none';

    const team = this._getEffectiveTeam();
    const isCam = team === 'cam';
    const isNlr = team === 'nlr';
    const isMaddie = team === 'maddie';

    let html = '';
    for (const row of sorted) {
      const mapping = this._findMapping(row.campaign, row.ownerName);
      const status = this._getRowStatus(row, mapping);
      const statusLabel = status === 'mapped' ? 'Mapped' : status === 'partial' ? 'Partial' : 'Unmapped';
      const statusIcon = status === 'mapped' ? '\u2705' : status === 'partial' ? '\u26A0\uFE0F' : '\u274C';
      const rowId = this._rowId(row.campaign, row.ownerName);
      const campaignLabel = OD_CONFIG.campaignSources[row.campaign]?.label || row.campaign;

      html += `<tr data-row-id="${rowId}">`;

      // Campaign pill
      html += `<td><span class="campaign-pill">${this._esc(campaignLabel)}</span></td>`;

      // Owner name
      html += `<td><strong>${this._esc(row.ownerName)}</strong></td>`;

      // Source Tab column (editable only when viewing as Maddie's team)
      if (isMaddie) {
        html += `<td class="cell-editable">${this._renderSourceTabSelect(row)}</td>`;
      } else {
        const tabMapping = this._findCampaignTabMap(row.campaign, row.ownerName);
        const val = tabMapping?.tabName || '';
        html += `<td class="cell-readonly">${val ? '<span class="readonly-value">' + this._esc(val) + '</span>' : '<span class="input-needed">Input Needed from Maddy</span>'}</td>`;
      }

      // Cam's Company column
      if (isCam) {
        html += `<td class="cell-editable">${this._renderCamSelect(row, mapping)}</td>`;
      } else {
        const val = mapping?.camCompany || '';
        html += `<td class="cell-readonly">${val ? '<span class="readonly-value">' + this._esc(val) + '</span>' : '<span class="input-needed">Input Needed from BIS</span>'}</td>`;
      }

      // NLR File column
      if (isNlr) {
        html += `<td class="cell-editable">${this._renderNlrFileSelect(row, mapping)}</td>`;
      } else {
        const val = mapping?.nlrWorkbookName || '';
        html += `<td class="cell-readonly">${val ? '<span class="readonly-value">' + this._esc(val) + '</span>' : '<span class="input-needed">Input Needed from NLR</span>'}</td>`;
      }

      // NLR Tab column
      if (isNlr) {
        html += `<td class="cell-editable">${this._renderNlrTabSelect(row, mapping)}</td>`;
      } else {
        const val = mapping?.nlrTab || '';
        html += `<td class="cell-readonly">${val ? '<span class="readonly-value">' + this._esc(val) + '</span>' : '<span class="input-needed">Input Needed from NLR</span>'}</td>`;
      }

      // Status badge
      html += `<td><span class="status-badge ${status}">${statusIcon} ${statusLabel}</span></td>`;

      html += `</tr>`;
    }

    tbody.innerHTML = html;
  },

  /**
   * Render a searchable dropdown trigger (shared by all 3 column types)
   * @param {string} cellId - unique ID for this dropdown
   * @param {string} displayVal - current display text (or '' for placeholder)
   * @param {string} placeholder - placeholder text
   * @param {boolean} disabled - whether the dropdown is disabled
   * @param {string} onClickFn - JS expression for opening the dropdown
   * @param {string} onClearFn - JS expression for clearing the value
   */
  _renderSearchDropdown(cellId, displayVal, placeholder, disabled, onClickFn, onClearFn) {
    const hasVal = displayVal ? ' has-value' : '';
    const disabledClass = disabled ? ' disabled' : '';

    return `<div class="sd-wrap" id="sd-wrap-${cellId}">
      <div class="sd-trigger${hasVal}${disabledClass}" onclick="if(!this.classList.contains('disabled')){${onClickFn}}">
        <span style="overflow:hidden;text-overflow:ellipsis">${displayVal ? this._esc(displayVal) : placeholder}</span>
        <span class="sd-clear" onclick="event.stopPropagation();${onClearFn}" title="Clear">&times;</span>
        <span class="sd-arrow">▾</span>
      </div>
      <div class="sd-dropdown" id="sd-dd-${cellId}"></div>
    </div>`;
  },

  /**
   * Open a searchable dropdown with given options
   * @param {string} cellId - matches the dropdown wrapper ID
   * @param {Array} options - [{ value, label }]
   * @param {string} currentVal - currently selected value
   * @param {Function} onSelect - callback(value, label) when an option is picked
   */
  _openSearchDropdown(cellId, options, currentVal, onSelect) {
    // Close any other open dropdowns first
    this._closeAllDropdowns();

    const dd = document.getElementById(`sd-dd-${cellId}`);
    if (!dd) return;

    // Build the dropdown content
    dd.innerHTML = `
      <div class="sd-search-wrap">
        <input class="sd-search" type="text" placeholder="Type to search..." autocomplete="off">
      </div>
      <div class="sd-options"></div>
    `;
    dd.classList.remove('flip-up');
    dd.classList.add('open');

    // Flip upward if dropdown would overflow viewport
    requestAnimationFrame(() => {
      const rect = dd.getBoundingClientRect();
      if (rect.bottom > window.innerHeight - 10) {
        dd.classList.add('flip-up');
      }
    });

    const searchInput = dd.querySelector('.sd-search');
    const optionsContainer = dd.querySelector('.sd-options');

    const renderOptions = (filter = '') => {
      const q = filter.toLowerCase();
      const filtered = q ? options.filter(o => o.label.toLowerCase().includes(q)) : options;

      if (filtered.length === 0) {
        optionsContainer.innerHTML = `<div class="sd-no-results">No matches</div>`;
        return;
      }

      optionsContainer.innerHTML = filtered.map(o => {
        const sel = o.value === currentVal ? ' selected' : '';
        return `<div class="sd-option${sel}" data-value="${this._esc(o.value)}" data-label="${this._esc(o.label)}">${this._esc(o.label)}</div>`;
      }).join('');

      // Click handlers on options
      optionsContainer.querySelectorAll('.sd-option').forEach(el => {
        el.onclick = () => {
          const val = el.dataset.value;
          const label = el.dataset.label;
          dd.classList.remove('open');
          onSelect(val, label);
        };
      });
    };

    renderOptions();

    // Search filtering
    searchInput.oninput = () => renderOptions(searchInput.value);
    searchInput.onkeydown = (e) => {
      if (e.key === 'Escape') { dd.classList.remove('open'); }
    };

    // Focus the search input
    setTimeout(() => searchInput.focus(), 50);

    // Close on outside click (one-time listener)
    setTimeout(() => {
      const closeHandler = (e) => {
        if (!dd.contains(e.target) && !dd.previousElementSibling?.contains(e.target)) {
          dd.classList.remove('open');
          document.removeEventListener('mousedown', closeHandler);
        }
      };
      document.addEventListener('mousedown', closeHandler);
    }, 10);
  },

  /**
   * Close all open searchable dropdowns
   */
  _closeAllDropdowns() {
    document.querySelectorAll('.sd-dropdown.open').forEach(dd => dd.classList.remove('open'));
  },

  /**
   * Render Cam's company searchable dropdown
   */
  _renderCamSelect(row, mapping) {
    const val = mapping?.camCompany || '';
    const cellId = `cam-${this._rowId(row.campaign, row.ownerName)}`;
    const campaign = this._esc(row.campaign);
    const owner = this._esc(row.ownerName);

    return this._renderSearchDropdown(
      cellId, val, '-- Select Company --', false,
      `OwnerDev._openCamDropdown('${campaign}','${owner}','${cellId}')`,
      `OwnerDev._clearCamCompany('${campaign}','${owner}')`
    );
  },

  /**
   * Open Cam's company searchable dropdown
   */
  _openCamDropdown(campaign, ownerName, cellId) {
    const mapping = this._findMapping(campaign, ownerName);
    const currentVal = mapping?.camCompany || '';
    const options = [
      { value: 'Non-Partner', label: '⛔ Non-Partner' },
      ...this.state.camCompanies.map(c => ({ value: c, label: c }))
    ];

    this._openSearchDropdown(cellId, options, currentVal, (value, label) => {
      // Update the trigger display
      const trigger = document.querySelector(`#sd-wrap-${cellId} .sd-trigger`);
      if (trigger) {
        trigger.querySelector('span').textContent = value || '-- Select Company --';
        trigger.classList.toggle('has-value', !!value);
      }
      // Fire the save
      this._onCamCompanyChange(campaign, ownerName, value);
    });
  },

  /**
   * Clear Cam's company value
   */
  _clearCamCompany(campaign, ownerName) {
    const cellId = `cam-${this._rowId(campaign, ownerName)}`;
    const trigger = document.querySelector(`#sd-wrap-${cellId} .sd-trigger`);
    if (trigger) {
      trigger.querySelector('span').textContent = '-- Select Company --';
      trigger.classList.remove('has-value');
    }
    this._onCamCompanyChange(campaign, ownerName, '');
  },

  /**
   * Render NLR workbook file searchable dropdown
   */
  _renderNlrFileSelect(row, mapping) {
    const val = mapping?.nlrWorkbookId || '';
    const displayVal = mapping?.nlrWorkbookName || '';
    const cellId = `nlr-file-${this._rowId(row.campaign, row.ownerName)}`;
    const campaign = this._esc(row.campaign);
    const owner = this._esc(row.ownerName);

    return this._renderSearchDropdown(
      cellId, displayVal, '-- Select File --', false,
      `OwnerDev._openNlrFileDropdown('${campaign}','${owner}','${cellId}')`,
      `OwnerDev._clearNlrFile('${campaign}','${owner}')`
    );
  },

  /**
   * Open NLR file searchable dropdown
   */
  _openNlrFileDropdown(campaign, ownerName, cellId) {
    const mapping = this._findMapping(campaign, ownerName);
    const currentVal = mapping?.nlrWorkbookId || '';
    const options = [
      { value: 'Non-Partner', label: '⛔ Non-Partner' },
      ...this.state.nlrWorkbooks.map(wb => ({ value: wb.id, label: wb.name }))
    ];

    this._openSearchDropdown(cellId, options, currentVal, (value, label) => {
      const trigger = document.querySelector(`#sd-wrap-${cellId} .sd-trigger`);
      if (trigger) {
        trigger.querySelector('span').textContent = value === 'Non-Partner' ? 'Non-Partner' : (value ? label : '-- Select File --');
        trigger.classList.toggle('has-value', !!value);
      }
      // Fire save (this also fetches tabs for the new workbook)
      this._onNlrFileChangeSearchable(campaign, ownerName, value, label);
    });
  },

  /**
   * Clear NLR file value
   */
  _clearNlrFile(campaign, ownerName) {
    const cellId = `nlr-file-${this._rowId(campaign, ownerName)}`;
    const trigger = document.querySelector(`#sd-wrap-${cellId} .sd-trigger`);
    if (trigger) {
      trigger.querySelector('span').textContent = '-- Select File --';
      trigger.classList.remove('has-value');
    }
    this._onNlrFileChangeSearchable(campaign, ownerName, '', '');
  },

  /**
   * Render NLR tab searchable dropdown
   */
  _renderNlrTabSelect(row, mapping) {
    const wbId = mapping?.nlrWorkbookId || '';
    const val = mapping?.nlrTab || '';
    const cellId = `nlr-tab-${this._rowId(row.campaign, row.ownerName)}`;
    const campaign = this._esc(row.campaign);
    const owner = this._esc(row.ownerName);
    const disabled = !wbId;

    return this._renderSearchDropdown(
      cellId, val, '-- Select Tab --', disabled,
      `OwnerDev._openNlrTabDropdown('${campaign}','${owner}','${cellId}')`,
      `OwnerDev._clearNlrTab('${campaign}','${owner}')`
    );
  },

  /**
   * Open NLR tab searchable dropdown — fetches tabs on demand if not cached
   */
  async _openNlrTabDropdown(campaign, ownerName, cellId) {
    const mapping = this._findMapping(campaign, ownerName);
    const wbId = mapping?.nlrWorkbookId || '';
    const currentVal = mapping?.nlrTab || '';

    // If tabs aren't cached yet for this workbook, fetch them on demand
    if (wbId && !this.state.nlrTabsCache[wbId]) {
      const trigger = document.querySelector(`#sd-wrap-${cellId} .sd-trigger`);
      if (trigger) trigger.querySelector('span').textContent = 'Loading tabs...';
      try {
        const tabsRes = await this._api('odNlrTabs', { sheetId: wbId });
        this.state.nlrTabsCache[wbId] = tabsRes.success ? (tabsRes.tabs || []) : [];
      } catch (err) {
        console.error('[OwnerDev] Tab fetch error for', wbId, err);
        this.state.nlrTabsCache[wbId] = [];
      }
      if (trigger) trigger.querySelector('span').textContent = currentVal || '-- Select Tab --';
    }

    const tabs = (wbId && this.state.nlrTabsCache[wbId]) || [];
    const options = [
      { value: 'Non-Partner', label: '⛔ Non-Partner' },
      ...tabs.map(t => ({ value: t, label: t }))
    ];

    this._openSearchDropdown(cellId, options, currentVal, (value, label) => {
      const trigger = document.querySelector(`#sd-wrap-${cellId} .sd-trigger`);
      if (trigger) {
        trigger.querySelector('span').textContent = value || '-- Select Tab --';
        trigger.classList.toggle('has-value', !!value);
      }
      this._onNlrTabChangeSearchable(campaign, ownerName, value);
    });
  },

  /**
   * Clear NLR tab value
   */
  _clearNlrTab(campaign, ownerName) {
    const cellId = `nlr-tab-${this._rowId(campaign, ownerName)}`;
    const trigger = document.querySelector(`#sd-wrap-${cellId} .sd-trigger`);
    if (trigger) {
      trigger.querySelector('span').textContent = '-- Select Tab --';
      trigger.classList.remove('has-value');
    }
    this._onNlrTabChangeSearchable(campaign, ownerName, '');
  },

  // ══════════════════════════════════════════════════════
  // SOURCE TAB — Campaign spreadsheet tab mapping
  // ══════════════════════════════════════════════════════

  /**
   * Find an existing campaign tab mapping for a campaign+owner
   */
  _findCampaignTabMap(campaign, ownerName) {
    return this.state.campaignTabMap.find(
      m => m.campaign === campaign && m.ownerName.toLowerCase() === ownerName.toLowerCase()
    ) || null;
  },

  /**
   * Upsert campaign tab mapping in local state
   */
  _upsertCampaignTabMap(campaign, ownerName, tabName) {
    let mapping = this._findCampaignTabMap(campaign, ownerName);
    if (mapping) {
      mapping.tabName = tabName;
    } else {
      this.state.campaignTabMap.push({ campaign, ownerName, tabName });
    }
  },

  /**
   * Render Source Tab searchable dropdown
   */
  _renderSourceTabSelect(row) {
    const tabMapping = this._findCampaignTabMap(row.campaign, row.ownerName);
    const val = tabMapping?.tabName || '';
    const cellId = `src-tab-${this._rowId(row.campaign, row.ownerName)}`;
    const campaign = this._esc(row.campaign);
    const owner = this._esc(row.ownerName);
    const hasTabs = (this.state.campaignTabsCache[row.campaign] || []).length > 0;

    return this._renderSearchDropdown(
      cellId, val, '-- Select Tab --', !hasTabs,
      `OwnerDev._openSourceTabDropdown('${campaign}','${owner}','${cellId}')`,
      `OwnerDev._clearSourceTab('${campaign}','${owner}')`
    );
  },

  /**
   * Open Source Tab searchable dropdown
   */
  _openSourceTabDropdown(campaign, ownerName, cellId) {
    const tabMapping = this._findCampaignTabMap(campaign, ownerName);
    const currentVal = tabMapping?.tabName || '';
    const tabs = this.state.campaignTabsCache[campaign] || [];
    const options = [
      { value: 'Non-Partner', label: '⛔ Non-Partner' },
      ...tabs.map(t => ({ value: t, label: t }))
    ];

    this._openSearchDropdown(cellId, options, currentVal, (value, label) => {
      const trigger = document.querySelector(`#sd-wrap-${cellId} .sd-trigger`);
      if (trigger) {
        trigger.querySelector('span').textContent = value || '-- Select Tab --';
        trigger.classList.toggle('has-value', !!value);
      }
      this._onSourceTabChange(campaign, ownerName, value);
    });
  },

  /**
   * Clear Source Tab value
   */
  _clearSourceTab(campaign, ownerName) {
    const cellId = `src-tab-${this._rowId(campaign, ownerName)}`;
    const trigger = document.querySelector(`#sd-wrap-${cellId} .sd-trigger`);
    if (trigger) {
      trigger.querySelector('span').textContent = '-- Select Tab --';
      trigger.classList.remove('has-value');
    }
    this._onSourceTabChange(campaign, ownerName, '');
  },

  /**
   * Auto-map campaign tabs by fuzzy-matching owner names to tab names.
   * Only fills in where there's no existing saved mapping.
   * Returns count of auto-mapped owners.
   */
  async _autoMapCampaignTabs() {
    let autoMapped = 0;
    const toSave = [];

    for (const [campaignKey, campaign] of Object.entries(this.state.campaigns)) {
      const tabs = this.state.campaignTabsCache[campaignKey] || [];
      if (!tabs.length) continue;

      for (const ownerName of (campaign.owners || [])) {
        // Skip if already has a saved mapping
        const existing = this._findCampaignTabMap(campaignKey, ownerName);
        if (existing?.tabName) continue;

        // Try to find a matching tab
        let matchedTab = null;

        // 1. Exact match (case-insensitive)
        const ownerLower = ownerName.toLowerCase().trim();
        matchedTab = tabs.find(t => t.toLowerCase().trim() === ownerLower);

        // 2. Fuzzy match using _namesMatch
        if (!matchedTab) {
          matchedTab = tabs.find(t => this._namesMatch(ownerName, t));
        }

        // 3. Token overlap: owner name tokens found inside tab name
        if (!matchedTab) {
          const ownerTokens = this._normName(ownerName).split(' ').filter(t => t.length >= 2);
          if (ownerTokens.length >= 1) {
            matchedTab = tabs.find(t => {
              const normTab = this._normName(t);
              const matchCount = ownerTokens.filter(tok => normTab.includes(tok)).length;
              return matchCount >= 2 || (ownerTokens.length === 1 && matchCount === 1);
            });
          }
        }

        if (matchedTab) {
          this._upsertCampaignTabMap(campaignKey, ownerName, matchedTab);
          toSave.push({ campaign: campaignKey, ownerName, tabName: matchedTab, updatedBy: 'auto-map' });
          autoMapped++;
        }
      }
    }

    // Batch save all auto-mapped tab mappings
    if (toSave.length > 0) {
      console.log(`[OwnerDev] Auto-mapped ${toSave.length} owners to campaign tabs`);
      this._post('odBatchSaveCampaignTabMap', { mappings: toSave })
        .then(res => console.log('[OwnerDev] Batch save campaign tab map result:', res))
        .catch(err => console.warn('[OwnerDev] Batch save campaign tab map error:', err.message));
    }

    return autoMapped;
  },

  // ══════════════════════════════════════════════════════
  // SAVE HANDLERS (auto-save on dropdown change)
  // ══════════════════════════════════════════════════════

  /**
   * Handle Source Tab save (called from searchable dropdown)
   */
  async _onSourceTabChange(campaign, ownerName, value) {
    const cellKey = `src-tab-${campaign}-${ownerName}`;
    if (this.state._savingCells.has(cellKey)) return;
    this.state._savingCells.add(cellKey);

    try {
      const res = await this._post('odSaveCampaignTabMap', {
        campaign,
        ownerName,
        tabName: value,
        updatedBy: this.state.session.email
      });

      if (res.success) {
        this._upsertCampaignTabMap(campaign, ownerName, value);
        this._toast('Source tab saved', 'success');
      } else {
        this._toast(res.message || 'Save failed', 'error');
      }
    } catch (err) {
      console.error('[OwnerDev] Source tab save error:', err);
      this._toast('Save failed. Please try again.', 'error');
    } finally {
      this.state._savingCells.delete(cellKey);
    }
  },

  /**
   * Handle Cam's Company save (called from searchable dropdown)
   */
  async _onCamCompanyChange(campaign, ownerName, value) {
    const cellKey = `cam-${campaign}-${ownerName}`;
    if (this.state._savingCells.has(cellKey)) return;
    this.state._savingCells.add(cellKey);

    try {
      const res = await this._post('odSaveMapping', {
        campaign,
        ownerName,
        field: 'camCompany',
        value,
        updatedBy: this.state.session.email
      });

      if (res.success) {
        this._upsertMapping(campaign, ownerName, { camCompany: value });
        this._renderStats();
        this._toast('Saved', 'success');
        this._updateRowStatus(campaign, ownerName);
      } else {
        this._toast(res.message || 'Save failed', 'error');
      }
    } catch (err) {
      console.error('[OwnerDev] Save error:', err);
      this._toast('Save failed. Please try again.', 'error');
    } finally {
      this.state._savingCells.delete(cellKey);
    }
  },

  /**
   * Handle NLR File save + fetch tabs for the workbook (searchable dropdown version)
   */
  async _onNlrFileChangeSearchable(campaign, ownerName, wbId, wbName) {
    const cellKey = `nlr-file-${campaign}-${ownerName}`;
    if (this.state._savingCells.has(cellKey)) return;
    this.state._savingCells.add(cellKey);

    // Reset the tab dropdown trigger
    const tabCellId = `nlr-tab-${this._rowId(campaign, ownerName)}`;
    const tabTrigger = document.querySelector(`#sd-wrap-${tabCellId} .sd-trigger`);

    try {
      // Save workbook selection (clear tab when file changes)
      const res = await this._post('odSaveMapping', {
        campaign,
        ownerName,
        field: 'nlrWorkbook',
        nlrWorkbookId: wbId,
        nlrWorkbookName: wbName,
        nlrTab: '',
        updatedBy: this.state.session.email
      });

      if (res.success) {
        this._upsertMapping(campaign, ownerName, {
          nlrWorkbookId: wbId,
          nlrWorkbookName: wbName,
          nlrTab: ''
        });
        this._renderStats();
        this._toast('File saved', 'success');
        this._updateRowStatus(campaign, ownerName);
      } else {
        this._toast(res.message || 'Save failed', 'error');
      }

      // Fetch tabs for the newly selected workbook
      if (wbId) {
        if (tabTrigger) {
          tabTrigger.querySelector('span').textContent = 'Loading tabs...';
          tabTrigger.classList.add('disabled');
        }

        if (!this.state.nlrTabsCache[wbId]) {
          try {
            const tabsRes = await this._api('odNlrTabs', { sheetId: wbId });
            this.state.nlrTabsCache[wbId] = tabsRes.success ? (tabsRes.tabs || []) : [];
          } catch (err) {
            console.error('[OwnerDev] Tab fetch error:', err);
            this.state.nlrTabsCache[wbId] = [];
          }
        }

        // Re-enable the tab dropdown
        if (tabTrigger) {
          tabTrigger.querySelector('span').textContent = '-- Select Tab --';
          tabTrigger.classList.remove('has-value', 'disabled');
        }
      } else if (tabTrigger) {
        tabTrigger.querySelector('span').textContent = '-- Select Tab --';
        tabTrigger.classList.remove('has-value');
        tabTrigger.classList.add('disabled');
      }
    } catch (err) {
      console.error('[OwnerDev] Save error:', err);
      this._toast('Save failed. Please try again.', 'error');
    } finally {
      this.state._savingCells.delete(cellKey);
    }
  },

  /**
   * Handle NLR Tab save (searchable dropdown version)
   */
  async _onNlrTabChangeSearchable(campaign, ownerName, value) {
    const cellKey = `nlr-tab-${campaign}-${ownerName}`;
    if (this.state._savingCells.has(cellKey)) return;
    this.state._savingCells.add(cellKey);

    try {
      const res = await this._post('odSaveMapping', {
        campaign,
        ownerName,
        field: 'nlrTab',
        value,
        updatedBy: this.state.session.email
      });

      if (res.success) {
        this._upsertMapping(campaign, ownerName, { nlrTab: value });
        this._renderStats();
        this._toast('Saved', 'success');
        this._updateRowStatus(campaign, ownerName);
      } else {
        this._toast(res.message || 'Save failed', 'error');
      }
    } catch (err) {
      console.error('[OwnerDev] Save error:', err);
      this._toast('Save failed. Please try again.', 'error');
    } finally {
      this.state._savingCells.delete(cellKey);
    }
  },

  // ══════════════════════════════════════════════════════
  // RENDERING — Team Management
  // ══════════════════════════════════════════════════════

  renderTeam() {
    const team = this._getEffectiveTeam();
    if (!team) return;

    const teamCfg = OD_CONFIG.teams[team];
    const members = this.state.users.filter(u => u.team === team);

    // Title
    document.getElementById('team-view-title').textContent = (teamCfg?.label || 'Team') + ' Members';
    document.getElementById('team-count').textContent = members.length + ' member' + (members.length !== 1 ? 's' : '');

    // Member cards
    const grid = document.getElementById('team-members-grid');
    if (members.length === 0) {
      grid.innerHTML = '<div class="empty-state"><div class="empty-state-text">No team members yet</div></div>';
      return;
    }

    let html = '';
    for (const m of members) {
      const initials = (m.name || m.email).split(/\s+/).map(w => w[0]).join('').toUpperCase().slice(0, 2);
      const roleLabel = m.role === 'manager' ? 'Manager' : 'Member';
      const isManager = this.state.session.role === 'manager';
      const isSelf = m.email.toLowerCase() === this.state.session.email.toLowerCase();

      html += `<div class="team-member-card">
        <div class="team-member-avatar" style="background:${teamCfg?.color || 'var(--gray-400)'}">${initials}</div>
        <div class="team-member-info">
          <div class="team-member-name">${this._esc(m.name || m.email)}</div>
          <div class="team-member-email">${this._esc(m.email)}</div>
          <span class="team-member-role">${roleLabel}</span>
        </div>
        ${isManager && !isSelf ? `<div class="team-member-actions"><button class="btn-remove-member" onclick="OwnerDev.removeMember('${this._esc(m.email)}')">Remove</button></div>` : ''}
      </div>`;
    }

    grid.innerHTML = html;
  },

  /**
   * Add a new team member
   */
  async addMember() {
    const emailInput = document.getElementById('add-email');
    const nameInput = document.getElementById('add-name');
    const error = document.getElementById('add-member-error');
    error.textContent = '';

    const email = emailInput.value.trim().toLowerCase();
    const name = nameInput.value.trim();

    if (!email) { error.textContent = 'Email is required'; return; }
    if (!name) { error.textContent = 'Name is required'; return; }

    const btn = document.querySelector('.btn-add-member');
    btn.disabled = true;

    try {
      const res = await this._post('odSaveUser', {
        email,
        name,
        team: this._getEffectiveTeam(),
        role: 'member',
        addedBy: this.state.session.email
      });

      if (res.success) {
        this._toast('Member added', 'success');
        emailInput.value = '';
        nameInput.value = '';

        // Re-fetch users and re-render
        try {
          const usersRes = await this._api('odGetUsers');
          if (usersRes.success) this.state.users = usersRes.users || [];
        } catch {}
        this.renderTeam();
      } else {
        error.textContent = res.message || 'Failed to add member';
      }
    } catch (err) {
      console.error('[OwnerDev] Add member error:', err);
      error.textContent = 'Connection failed. Please try again.';
    } finally {
      btn.disabled = false;
    }
  },

  /**
   * Remove (soft delete) a team member
   */
  async removeMember(email) {
    if (!confirm(`Remove ${email} from the team?`)) return;

    try {
      const res = await this._post('odDeleteUser', {
        email,
        deletedBy: this.state.session.email
      });

      if (res.success) {
        this._toast('Member removed', 'success');
        // Re-fetch users and re-render
        try {
          const usersRes = await this._api('odGetUsers');
          if (usersRes.success) this.state.users = usersRes.users || [];
        } catch {}
        this.renderTeam();
      } else {
        this._toast(res.message || 'Failed to remove member', 'error');
      }
    } catch (err) {
      console.error('[OwnerDev] Remove member error:', err);
      this._toast('Connection failed. Please try again.', 'error');
    }
  },

  // ══════════════════════════════════════════════════════
  // DATA HELPERS
  // ══════════════════════════════════════════════════════

  /**
   * Build flat array of { campaign, ownerName } rows from campaigns state.
   * Applies campaign filter and search query.
   * @param {boolean} ignoreFilter - If true, ignore campaign filter (for total counts)
   */
  _getFilteredRows(ignoreFilter = false) {
    const rows = [];
    for (const [campaignKey, campaign] of Object.entries(this.state.campaigns)) {
      if (!ignoreFilter && this.state.activeCampaign !== 'all' && campaignKey !== this.state.activeCampaign) continue;
      for (const ownerName of (campaign.owners || [])) {
        if (this.state.searchQuery && !ownerName.toLowerCase().includes(this.state.searchQuery)) continue;
        rows.push({ campaign: campaignKey, ownerName });
      }
    }
    return rows;
  },

  /**
   * Find an existing mapping for a campaign+owner
   */
  _findMapping(campaign, ownerName) {
    return this.state.mappings.find(
      m => m.campaign === campaign && m.ownerName === ownerName
    ) || null;
  },

  /**
   * Upsert mapping in local state
   */
  _upsertMapping(campaign, ownerName, updates) {
    let mapping = this._findMapping(campaign, ownerName);
    if (mapping) {
      Object.assign(mapping, updates);
    } else {
      mapping = { campaign, ownerName, camCompany: '', nlrWorkbookId: '', nlrWorkbookName: '', nlrTab: '', ...updates };
      this.state.mappings.push(mapping);
    }
  },

  /**
   * Determine row status: 'mapped' | 'partial' | 'unmapped'
   */
  _getRowStatus(row, mapping) {
    if (!mapping) mapping = this._findMapping(row.campaign, row.ownerName);
    const hasCam = !!(mapping?.camCompany);
    const hasNlr = !!(mapping?.nlrTab);

    if (hasCam && hasNlr) return 'mapped';
    if (hasCam || hasNlr) return 'partial';
    return 'unmapped';
  },

  /**
   * Update just the status badge in an existing row without full re-render
   */
  _updateRowStatus(campaign, ownerName) {
    const rowId = this._rowId(campaign, ownerName);
    const tr = document.querySelector(`tr[data-row-id="${rowId}"]`);
    if (!tr) return;

    const row = { campaign, ownerName };
    const status = this._getRowStatus(row);
    const statusLabel = status === 'mapped' ? 'Mapped' : status === 'partial' ? 'Partial' : 'Unmapped';
    const statusIcon = status === 'mapped' ? '\u2705' : status === 'partial' ? '\u26A0\uFE0F' : '\u274C';

    const statusTd = tr.querySelector('td:last-child');
    if (statusTd) {
      statusTd.innerHTML = `<span class="status-badge ${status}">${statusIcon} ${statusLabel}</span>`;
    }
  },

  /**
   * Sort rows array by current sort column
   */
  _sortRows(rows) {
    const col = this.state.sortCol;
    const asc = this.state.sortAsc ? 1 : -1;

    return [...rows].sort((a, b) => {
      let va, vb;
      if (col === 'campaign') {
        va = (OD_CONFIG.campaignSources[a.campaign]?.label || a.campaign).toLowerCase();
        vb = (OD_CONFIG.campaignSources[b.campaign]?.label || b.campaign).toLowerCase();
      } else if (col === 'ownerName') {
        va = a.ownerName.toLowerCase();
        vb = b.ownerName.toLowerCase();
      } else if (col === 'status') {
        const order = { unmapped: 0, partial: 1, mapped: 2 };
        va = order[this._getRowStatus(a)] ?? 0;
        vb = order[this._getRowStatus(b)] ?? 0;
      } else {
        va = a[col] || '';
        vb = b[col] || '';
      }
      if (va < vb) return -1 * asc;
      if (va > vb) return 1 * asc;
      return 0;
    });
  },

  /**
   * Update sort arrows in table header
   */
  _updateSortArrows() {
    document.querySelectorAll('.mapping-table thead th').forEach(th => {
      const col = th.dataset.col;
      const arrow = th.querySelector('.sort-arrow');
      if (!arrow) return;

      if (col === this.state.sortCol) {
        th.classList.add('sorted');
        arrow.textContent = this.state.sortAsc ? '\u25B2' : '\u25BC';
      } else {
        th.classList.remove('sorted');
        arrow.textContent = '\u25B2';
      }
    });
  },

  /**
   * Generate a unique row ID for data attributes
   */
  _rowId(campaign, ownerName) {
    return btoa(campaign + '|' + ownerName).replace(/[^a-zA-Z0-9]/g, '');
  },

  // ══════════════════════════════════════════════════════
  // UI HELPERS
  // ══════════════════════════════════════════════════════

  _showLoading(text) {
    document.getElementById('loading-text').textContent = text || 'Loading...';
    document.getElementById('loading-screen').style.display = 'flex';
  },

  _hideLoading() {
    document.getElementById('loading-screen').style.display = 'none';
  },

  /**
   * Show a brief toast notification
   */
  _toast(message, type = 'success') {
    const el = document.getElementById('toast');
    el.textContent = type === 'success' ? message + ' \u2713' : message;
    el.className = 'toast ' + type;

    // Force reflow for animation
    void el.offsetWidth;
    el.classList.add('show');

    clearTimeout(this._toastTimer);
    this._toastTimer = setTimeout(() => {
      el.classList.remove('show');
    }, 2000);
  },

  /**
   * HTML-escape a string
   */
  _esc(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  },

  // ══════════════════════════════════════════════════════
  // VIEW-AS (Superadmin team switcher)
  // ══════════════════════════════════════════════════════

  /**
   * Get the effective team for rendering (view-as or real)
   */
  _getEffectiveTeam() {
    return this.state.viewAsTeam || this.state.session?.team || 'maddie';
  },

  /**
   * Build and inject the View-As bar at the top of the page
   */
  _buildViewAsBar() {
    if (document.getElementById('view-as-bar')) return; // already built

    const bar = document.createElement('div');
    bar.id = 'view-as-bar';
    bar.style.cssText = `
      position:fixed; top:0; left:0; right:0; z-index:10000;
      background:#1a1a2e; color:#fff; padding:0 16px;
      height:40px; display:flex; align-items:center; gap:12px;
      font-size:13px; font-family:var(--font,Inter,sans-serif);
      border-bottom:2px solid #00c8ff; box-shadow:0 2px 8px rgba(0,0,0,0.3);
    `;

    // Build team pills
    let pills = '';
    for (const [key, cfg] of Object.entries(OD_CONFIG.teams)) {
      pills += `<button class="va-pill" data-team="${key}" onclick="OwnerDev.viewAsTeam('${key}')" style="
        background:transparent; border:1px solid rgba(255,255,255,0.3); color:#fff;
        padding:4px 12px; border-radius:20px; font-size:12px; font-weight:600;
        cursor:pointer; transition:all 0.15s;
      ">${cfg.icon} ${cfg.label}</button>`;
    }

    bar.innerHTML = `
      <span style="font-weight:700;color:#00c8ff;white-space:nowrap;">View As</span>
      <span style="color:rgba(255,255,255,0.4)">|</span>
      ${pills}
      <button class="va-pill va-pill-reset" data-team="off" onclick="OwnerDev.viewAsTeam(null)" style="
        background:transparent; border:1px solid rgba(255,255,255,0.15); color:rgba(255,255,255,0.5);
        padding:4px 12px; border-radius:20px; font-size:12px; font-weight:500;
        cursor:pointer; margin-left:auto; transition:all 0.15s;
      ">Reset (${this.state.session?.team || 'maddie'})</button>
    `;

    document.body.prepend(bar);

    // Push dashboard content down to make room for the bar
    document.body.style.paddingTop = '40px';

    // Highlight current team pill
    this._updateViewAsPills();
  },

  /**
   * Switch view to a different team
   */
  viewAsTeam(teamKey) {
    this.state.viewAsTeam = teamKey;

    // Update badge in top bar
    this._renderUserInfo();

    // Update pill highlights
    this._updateViewAsPills();

    // Re-render the mapping table (changes editable columns)
    this._renderStats();
    this.renderMapping();

    // Re-render team view if active
    if (this.state.activeTab === 'team') {
      this.renderTeam();
    }

    // Toast
    if (teamKey) {
      const cfg = OD_CONFIG.teams[teamKey];
      this._toast(`Viewing as ${cfg?.label || teamKey}`, 'success');
    } else {
      this._toast('Reset to own view', 'success');
    }
  },

  /**
   * Update View-As pill active states
   */
  _updateViewAsPills() {
    const active = this.state.viewAsTeam;
    document.querySelectorAll('.va-pill').forEach(pill => {
      const team = pill.dataset.team;
      if (team === 'off') {
        pill.style.borderColor = !active ? '#00c8ff' : 'rgba(255,255,255,0.15)';
        pill.style.color = !active ? '#00c8ff' : 'rgba(255,255,255,0.5)';
      } else {
        const isActive = team === active;
        const cfg = OD_CONFIG.teams[team];
        pill.style.borderColor = isActive ? (cfg?.color || '#00c8ff') : 'rgba(255,255,255,0.3)';
        pill.style.background = isActive ? (cfg?.color || '#00c8ff') : 'transparent';
        pill.style.color = isActive ? '#fff' : 'rgba(255,255,255,0.7)';
      }
    });
  },

  // ══════════════════════════════════════════════════════
  // PLANNING TAB
  // ══════════════════════════════════════════════════════

  _PLAN_DAY_NAMES: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],

  async _initPlanning() {
    if (!this.state.planningLoaded) {
      try {
        const res = await this._api('odGetPlanning');
        if (res.success && res.planning) {
          this.state.planningData = res.planning;
        }
      } catch (err) {
        console.warn('[OwnerDev] Failed to load planning:', err.message);
      }
      this.state.planningLoaded = true;
    }
    this._renderPlanningGrid();
  },

  _renderPlanningGrid() {
    const grid = document.getElementById('planning-grid');
    const poolCards = document.getElementById('planning-pool-cards');
    if (!grid || !poolCards) return;

    // All campaigns with owners
    const allCampaigns = {};
    for (const [key, camp] of Object.entries(this.state.campaigns)) {
      if (camp.owners && camp.owners.length > 0) {
        allCampaigns[key] = camp;
      }
    }

    // Determine assigned campaigns (exist in planningData)
    const assignedKeys = new Set(this.state.planningData.map(p => p.campaignKey));
    const unassignedKeys = Object.keys(allCampaigns).filter(k => !assignedKeys.has(k));

    // Render unassigned pool
    poolCards.innerHTML = unassignedKeys.map(key =>
      this._planningCardHtml(key, allCampaigns[key])
    ).join('');

    // Today index: Mon=0, Sun=6
    const todayIdx = (new Date().getDay() + 6) % 7;

    // Render 7 day columns
    grid.innerHTML = this._PLAN_DAY_NAMES.map((name, dayIdx) => {
      const dayCampaigns = this.state.planningData
        .filter(p => p.day === dayIdx)
        .sort((a, b) => a.sortOrder - b.sortOrder);

      const todayCls = dayIdx === todayIdx ? ' today' : '';
      const cardsHtml = dayCampaigns.map(p => {
        const camp = allCampaigns[p.campaignKey];
        if (!camp) return '';
        return this._planningCardHtml(p.campaignKey, camp, dayIdx);
      }).join('');

      return `
        <div class="planning-day-col">
          <div class="planning-day-header${todayCls}">${name}</div>
          <div class="planning-drop-zone" data-day="${dayIdx}">${cardsHtml}</div>
        </div>`;
    }).join('');

    this._setupPlanningDnD();
  },

  _planningCardHtml(campaignKey, camp, day) {
    const label = camp?.label || campaignKey;
    const ownerCount = camp?.owners?.length || 0;
    const logos = (typeof NationalApp !== 'undefined' && NationalApp.CAMPAIGN_LOGOS)
      ? NationalApp.CAMPAIGN_LOGOS[campaignKey] : null;

    let logoHtml;
    if (Array.isArray(logos)) {
      logoHtml = `<span class="planning-card-logo-dual">${logos.map(l => `<img src="${l}" class="planning-card-logo">`).join('')}</span>`;
    } else if (logos) {
      logoHtml = `<img src="${logos}" class="planning-card-logo">`;
    } else {
      logoHtml = `<span class="planning-card-logo-placeholder">&#x1F4CA;</span>`;
    }

    const reorderBtn = day !== undefined
      ? `<button class="planning-card-reorder" onclick="event.stopPropagation();OwnerDev._openReorderModal('${campaignKey}',${day})" title="Reorder owners">&#x1F465;</button>`
      : '';

    return `
      <div class="planning-card" draggable="true" data-campaign="${campaignKey}">
        ${logoHtml}
        <div class="planning-card-info">
          <div class="planning-card-label">${this._esc(label)}</div>
          <div class="planning-card-owners">${ownerCount} owner${ownerCount !== 1 ? 's' : ''}</div>
        </div>
        ${reorderBtn}
      </div>`;
  },

  _setupPlanningDnD() {
    let draggedKey = null;

    // Cards — dragstart/dragend
    document.querySelectorAll('#view-planning .planning-card').forEach(card => {
      card.addEventListener('dragstart', e => {
        draggedKey = card.dataset.campaign;
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', draggedKey);
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        draggedKey = null;
        // Clean up all insertion indicators
        document.querySelectorAll('.planning-insert-indicator').forEach(el => el.remove());
        document.querySelectorAll('.planning-drop-zone').forEach(z => z.classList.remove('drag-over'));
      });
    });

    // Drop zones — track insertion position via indicator line
    document.querySelectorAll('.planning-drop-zone').forEach(zone => {
      zone.addEventListener('dragover', e => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        this._updateInsertIndicator(zone, e.clientY, draggedKey);
      });
      zone.addEventListener('dragenter', e => {
        e.preventDefault();
        zone.classList.add('drag-over');
      });
      zone.addEventListener('dragleave', e => {
        if (!zone.contains(e.relatedTarget)) {
          zone.classList.remove('drag-over');
          this._removeInsertIndicator(zone);
        }
      });
      zone.addEventListener('drop', e => {
        e.preventDefault();
        zone.classList.remove('drag-over');
        const key = e.dataTransfer.getData('text/plain') || draggedKey;
        if (!key) return;
        const day = parseInt(zone.dataset.day);
        const insertIdx = this._getInsertIndex(zone, e.clientY, key);
        this._removeInsertIndicator(zone);
        this._moveCampaignToDay(key, day, insertIdx);
      });
    });

    // Pool drop zone
    const pool = document.getElementById('planning-pool');
    if (pool) {
      pool.addEventListener('dragover', e => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; });
      pool.addEventListener('dragenter', e => { e.preventDefault(); pool.classList.add('drag-over'); });
      pool.addEventListener('dragleave', e => {
        if (!pool.contains(e.relatedTarget)) pool.classList.remove('drag-over');
      });
      pool.addEventListener('drop', e => {
        e.preventDefault();
        pool.classList.remove('drag-over');
        const key = e.dataTransfer.getData('text/plain') || draggedKey;
        if (!key) return;
        this._removeCampaignFromPlanning(key);
      });
    }
  },

  /** Get the insert index based on cursor Y within a drop zone */
  _getInsertIndex(zone, clientY, draggedKey) {
    const cards = [...zone.querySelectorAll('.planning-card')].filter(
      c => c.dataset.campaign !== draggedKey
    );
    for (let i = 0; i < cards.length; i++) {
      const rect = cards[i].getBoundingClientRect();
      if (clientY < rect.top + rect.height / 2) return i;
    }
    return cards.length;
  },

  /** Show/move a thin indicator line at the insertion point */
  _updateInsertIndicator(zone, clientY, draggedKey) {
    let indicator = zone.querySelector('.planning-insert-indicator');
    if (!indicator) {
      indicator = document.createElement('div');
      indicator.className = 'planning-insert-indicator';
      zone.appendChild(indicator);
    }

    const cards = [...zone.querySelectorAll('.planning-card')].filter(
      c => c.dataset.campaign !== draggedKey
    );

    if (cards.length === 0) {
      indicator.style.display = 'none';
      return;
    }

    indicator.style.display = '';
    const idx = this._getInsertIndex(zone, clientY, draggedKey);
    if (idx < cards.length) {
      zone.insertBefore(indicator, cards[idx]);
    } else {
      // After last card
      const lastCard = cards[cards.length - 1];
      if (lastCard.nextSibling) {
        zone.insertBefore(indicator, lastCard.nextSibling);
      } else {
        zone.appendChild(indicator);
      }
    }
  },

  _removeInsertIndicator(zone) {
    zone.querySelectorAll('.planning-insert-indicator').forEach(el => el.remove());
  },

  _moveCampaignToDay(campaignKey, day, insertIdx) {
    // Preserve existing ownerOrder if moving within/between days
    const existing = this.state.planningData.find(p => p.campaignKey === campaignKey);
    const ownerOrder = existing ? (existing.ownerOrder || []) : [];

    // Remove from current position
    this.state.planningData = this.state.planningData.filter(p => p.campaignKey !== campaignKey);

    // Get ordered list of campaigns already in this day
    const dayItems = this.state.planningData
      .filter(p => p.day === day)
      .sort((a, b) => a.sortOrder - b.sortOrder);

    // Insert at the right position
    const idx = (insertIdx !== undefined) ? insertIdx : dayItems.length;
    dayItems.splice(idx, 0, { day, sortOrder: 0, campaignKey, ownerOrder });

    // Recalculate sort orders
    dayItems.forEach((item, i) => { item.sortOrder = i; });

    // Replace this day's entries in planningData
    this.state.planningData = [
      ...this.state.planningData.filter(p => p.day !== day),
      ...dayItems
    ];
    this.state._planningDirty = true;
    this._renderPlanningGrid();
  },

  _removeCampaignFromPlanning(campaignKey) {
    this.state.planningData = this.state.planningData.filter(p => p.campaignKey !== campaignKey);
    this.state._planningDirty = true;
    this._renderPlanningGrid();
  },

  // ── Owner Reorder Modal ──

  _openReorderModal(campaignKey, day) {
    const camp = this.state.campaigns[campaignKey];
    if (!camp || !camp.owners || !camp.owners.length) return;

    const planEntry = this.state.planningData.find(p => p.campaignKey === campaignKey && p.day === day);
    let orderedOwners;
    if (planEntry && planEntry.ownerOrder && planEntry.ownerOrder.length > 0) {
      // Use saved order, append any new owners not in the saved list
      const known = new Set(planEntry.ownerOrder.map(n => n.toLowerCase()));
      orderedOwners = [...planEntry.ownerOrder];
      for (const o of camp.owners) {
        if (!known.has(o.toLowerCase())) orderedOwners.push(o);
      }
    } else {
      orderedOwners = [...camp.owners];
    }

    this._reorderState = { campaignKey, day, owners: orderedOwners };

    document.getElementById('reorder-modal-title').textContent =
      `Reorder — ${camp.label || campaignKey}`;
    this._renderReorderList();
    document.getElementById('owner-reorder-modal').style.display = '';
  },

  _renderReorderList() {
    const list = document.getElementById('reorder-owner-list');
    if (!list || !this._reorderState) return;

    list.innerHTML = this._reorderState.owners.map((name, idx) =>
      `<div class="reorder-item" draggable="true" data-idx="${idx}">
        <span class="reorder-drag-handle">&#x2630;</span>
        <span class="reorder-item-rank">${idx + 1}</span>
        <span class="reorder-item-name">${this._esc(name)}</span>
      </div>`
    ).join('');

    this._setupOwnerReorderDnD();
  },

  _setupOwnerReorderDnD() {
    const list = document.getElementById('reorder-owner-list');
    if (!list) return;
    let dragIdx = null;

    list.querySelectorAll('.reorder-item').forEach(item => {
      item.addEventListener('dragstart', e => {
        dragIdx = parseInt(item.dataset.idx);
        item.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', String(dragIdx));
      });
      item.addEventListener('dragend', () => {
        item.classList.remove('dragging');
        list.querySelectorAll('.reorder-item').forEach(i => i.classList.remove('drag-over-item'));
      });
      item.addEventListener('dragover', e => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
      });
      item.addEventListener('dragenter', e => {
        e.preventDefault();
        item.classList.add('drag-over-item');
      });
      item.addEventListener('dragleave', () => {
        item.classList.remove('drag-over-item');
      });
      item.addEventListener('drop', e => {
        e.preventDefault();
        item.classList.remove('drag-over-item');
        const fromIdx = parseInt(e.dataTransfer.getData('text/plain'));
        const toIdx = parseInt(item.dataset.idx);
        if (isNaN(fromIdx) || isNaN(toIdx) || fromIdx === toIdx) return;

        // Reorder
        const owners = this._reorderState.owners;
        const [moved] = owners.splice(fromIdx, 1);
        owners.splice(toIdx, 0, moved);
        this._reorderState.owners = owners;
        this.state._planningDirty = true;

        // Update planning data
        const entry = this.state.planningData.find(
          p => p.campaignKey === this._reorderState.campaignKey && p.day === this._reorderState.day
        );
        if (entry) entry.ownerOrder = [...owners];

        this._renderReorderList();
      });
    });
  },

  closeReorderModal() {
    document.getElementById('owner-reorder-modal').style.display = 'none';
    this._reorderState = null;
  },

  // ── Save Planning ──

  async savePlanning() {
    const status = document.getElementById('planning-save-status');
    if (status) status.textContent = 'Saving...';
    try {
      const res = await this._post('odSavePlanning', {
        planning: this.state.planningData,
        email: this.state.session.email
      });
      if (res.success) {
        this.state._planningDirty = false;
        this._toast('Schedule saved');
        if (status) status.textContent = 'Saved';
        setTimeout(() => { if (status) status.textContent = ''; }, 2000);
        // Update NationalApp's cache so Coach view is instant
        if (typeof NationalApp !== 'undefined') {
          NationalApp._planningSchedule = [...this.state.planningData];
          NationalApp._savePlanningToCache(this.state.planningData);
          NationalApp._coachInitDone = false; // force re-render on next Coach tab switch
        }
      } else {
        this._toast(res.message || 'Failed to save', 'error');
        if (status) status.textContent = 'Error';
      }
    } catch (err) {
      console.error('[OwnerDev] savePlanning error:', err);
      this._toast('Failed to save schedule', 'error');
      if (status) status.textContent = 'Error';
    }
  }
};

// ── Boot ──
document.addEventListener('DOMContentLoaded', () => OwnerDev.init());
