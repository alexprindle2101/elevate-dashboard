// ═══════════════════════════════════════════════════════
// National Consultant Dashboard — App Controller
// Pulls from 4 external Google Sheets, renders owner reviews
// ═══════════════════════════════════════════════════════

const NationalApp = {

  // ── Status code definitions (matching Campaign Tracker) ──
  STATUS_CODES: {
    22: { label: 'Bored Leaders', css: 'sc-22' },
    33: { label: 'Top Leaders Interviewing Only', css: 'sc-33' },
    44: { label: 'Maintaining Not Growing', css: 'sc-44' },
    55: { label: 'Leaders Busy', css: 'sc-55' },
    66: { label: 'Promotion Factory', css: 'sc-66' }
  },

  // ── Recruiting row labels (always the same) ──
  RECRUITING_LABELS: [
    { label: 'Applies Received', isRate: false },
    { label: 'Sent to List', isRate: false },
    { label: '1st Rounds Booked', isRate: false },
    { label: '1st Rounds Showed', isRate: false },
    { label: 'Retention', isRate: true },
    { label: '% Call List Booked', isRate: true },
    { label: '2nd Rounds Booked', isRate: false },
    { label: '2nd Rounds Showed', isRate: false },
    { label: 'Retention', isRate: true },
    { label: 'New Starts Booked', isRate: false },
    { label: 'New Starts Showed', isRate: false },
    { label: 'New Start Retention', isRate: true }
  ],

  // ── Campaign logo map (key → logo path) ──
  CAMPAIGN_LOGOS: {
    'att-b2b':         'references/logos/logo-att.png',
    'att-nds':         ['references/logos/logo-att.png', 'references/logos/logo-verizon.png'],
    'att-res':         'references/logos/logo-att.png',
    'frontier':        'references/logos/logo-frontier.png',
    'leafguard':       'references/logos/logo-leafguard.png',
    'lumen':           'references/logos/logo-lumen.png',
    'rogers':          'references/logos/logo-rogers.png',
    'truconnect':      'references/logos/logo-truconnect.png',
    'verizon':         'references/logos/logo-verizon.png',
    'verizon-fios':    'references/logos/logo-verizon.png'
  },

  state: {
    campaign: null,
    owners: [],
    selectedOwner: null,
    currentTab: 'health',
    session: null,
    campaignTotals: {},
    campaignRecruiting: null,
    camMapping: {},
    allCompanyNames: []
  },

  // ══════════════════════════════════════════════════
  // INIT
  // ══════════════════════════════════════════════════

  async init() {
    console.log('[NationalApp] init');
    this.state.session = this._getSession();
    if (!this.state.session) {
      this._showLogin();
      return;
    }
    document.getElementById('user-name').textContent = this.state.session.name || this.state.session.email;

    // Load planning from cache, then fetch everything in parallel
    this._loadPlanningFromCache();

    this._showLoading('Loading campaigns...');
    try {
      await Promise.all([
        this._fetchPlanningSchedule(),
        this.loadCampaignData('att-b2b')
      ]);
    } catch (err) {
      console.warn('[NationalApp] Initial fetch failed:', err.message);
    }
    this._hideLoading();
    document.getElementById('dashboard').style.display = 'block';
    this._showLandingPage();
  },

  /**
   * Embedded coach view init — called from OwnerDev.switchTab('coach').
   * Skips auth (uses OwnerDev's session) and doesn't toggle dashboard visibility.
   * @param {Object} session - { email, name } from OwnerDev
   */
  async initCoachView(session) {
    const campaign = session.campaign || null;
    if (this._coachInitDone && this._coachCampaign === campaign) {
      return;
    }
    console.log('[NationalApp] initCoachView (embedded in Owner Dev), campaign:', campaign);
    this._embedded = true;
    this._coachCampaign = campaign;
    this.state.session = { email: session.email, name: session.name, loginTime: Date.now() };

    // Load planning schedule from cache, then fetch fresh in parallel with data
    if (!this._planningSchedule) {
      this._loadPlanningFromCache();
    }
    // Fetch flagged reps + notes (non-blocking)
    this._fetchFlaggedReps();
    this._fetchNotes();

    if (campaign) {
      // Direct campaign selection — wait for both planning + campaign data
      this._showLoading('Loading campaign data...');
      try {
        await Promise.all([
          this._fetchPlanningSchedule(),
          this.selectCampaign(campaign)
        ]);
      } catch (err) {
        console.warn('[NationalApp] Coach view campaign fetch failed:', err.message);
      }
      this._hideLoading();
    } else {
      // Landing page — fetch planning + campaign data together, show loading until both ready
      this._showLoading('Loading coaching data...');
      const planningPromise = this._fetchPlanningSchedule();
      const dataPromise = this._tryRenderLandingFromCache()
        ? Promise.resolve(true)
        : this.loadCampaignData(Object.keys(NATIONAL_CONFIG.campaigns)[0] || 'frontier').catch(err => {
            console.warn('[NationalApp] Coach view campaign fetch failed:', err.message);
          });
      await Promise.all([planningPromise, dataPromise]);
      this._hideLoading();
      this._showLandingPage();
    }

    this._coachInitDone = true;
    this._coachCampaign = campaign;
  },

  /**
   * Try to render the landing page instantly from OwnerDev's od_data_cache.
   * Builds a lightweight _allCampaignsData with just labels + owner counts.
   * Returns true if successful, false to fall back to full fetch.
   */
  _tryRenderLandingFromCache() {
    try {
      const raw = localStorage.getItem('od_data_cache');
      if (!raw) return false;
      const cache = JSON.parse(raw);
      if (!cache.campaigns) return false;

      const lightweight = {};
      let hasCampaigns = false;
      for (const [key, camp] of Object.entries(cache.campaigns)) {
        if (!camp.owners || !camp.owners.length) continue;
        lightweight[key] = {
          label: camp.label || key,
          owners: camp.owners,  // string array — .length gives owner count
          weeks: []             // empty — landing page doesn't need weeks
        };
        hasCampaigns = true;

        // Ensure NATIONAL_CONFIG has entries for dynamically discovered campaigns
        if (!NATIONAL_CONFIG.campaigns[key]) {
          NATIONAL_CONFIG.campaigns[key] = { label: camp.label || key, weeksToPull: 6 };
        }
      }

      if (!hasCampaigns) return false;

      this._allCampaignsData = lightweight;
      this._populateCampaignSelector(lightweight);
      return true;
    } catch (err) {
      console.warn('[NationalApp] Landing cache read failed:', err.message);
      return false;
    }
  },

  // ══════════════════════════════════════════════════
  // SESSION (simple localStorage — same pattern as admin)
  // ══════════════════════════════════════════════════

  _getSession() {
    try {
      const raw = localStorage.getItem(NATIONAL_CONFIG.sessionKey);
      if (!raw) return null;
      const s = JSON.parse(raw);
      if (Date.now() - s.loginTime > NATIONAL_CONFIG.sessionDuration) {
        localStorage.removeItem(NATIONAL_CONFIG.sessionKey);
        return null;
      }
      return s;
    } catch { return null; }
  },

  _saveSession(email, name) {
    const s = { email, name, loginTime: Date.now() };
    localStorage.setItem(NATIONAL_CONFIG.sessionKey, JSON.stringify(s));
    return s;
  },

  logout() {
    localStorage.removeItem(NATIONAL_CONFIG.sessionKey);
    window.location.reload();
  },

  // ══════════════════════════════════════════════════
  // LOGIN (simple email-only for now, no PIN)
  // ══════════════════════════════════════════════════

  _showLogin() {
    const screen = document.getElementById('login-screen');
    screen.style.display = 'flex';
    const btn = document.getElementById('login-btn');
    const input = document.getElementById('login-email');
    const error = document.getElementById('login-error');

    const doLogin = () => {
      let email = input.value.trim().toLowerCase();
      if (!email) { error.textContent = 'Please enter your email'; return; }
      if (NATIONAL_CONFIG.loginAliases[email]) email = NATIONAL_CONFIG.loginAliases[email];
      const name = email.split('@')[0].replace(/[._]/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
      this.state.session = this._saveSession(email, name);
      screen.style.display = 'none';
      this.init();
    };

    btn.addEventListener('click', doLogin);
    input.addEventListener('keydown', e => { if (e.key === 'Enter') doLogin(); });
    setTimeout(() => input.focus(), 100);
  },

  // ══════════════════════════════════════════════════
  // DATA LOADING
  // ══════════════════════════════════════════════════

  // ── Fetch with timeout (default 20s) ──
  _fetchWithTimeout(promise, ms = 20000) {
    return Promise.race([
      promise,
      new Promise((_, reject) => setTimeout(() => reject(new Error('timeout')), ms))
    ]);
  },

  async loadCampaignData(campaignKey) {
    // Config entry may be dynamically created by _populateCampaignSelector
    if (!NATIONAL_CONFIG.campaigns[campaignKey]) {
      NATIONAL_CONFIG.campaigns[campaignKey] = { label: campaignKey, weeksToPull: 6 };
    }

    const hasApi = !!NATIONAL_CONFIG.appsScriptUrl;
    const hasNational = hasApi && NATIONAL_CONFIG.sheets.national && NATIONAL_CONFIG.sheets.national.id;

    // ── Fire ALL independent fetches in parallel (each with 20s timeout) ──
    const fetchPromises = {};
    if (hasNational) fetchPromises.recruiting = this._fetchWithTimeout(this._fetchRecruitingFromSheet(campaignKey));
    if (hasApi)      fetchPromises.audit      = this._fetchWithTimeout(this._fetchOnlinePresence());
    if (hasApi)      fetchPromises.camMapping  = this._fetchWithTimeout(this._fetchOwnerCamMapping());
    // Legacy B2B/NDS enrichment (still needed for campaigns using local tabs)
    const isB2B = campaignKey === 'att-b2b';
    const isNDS = campaignKey.indexOf('nds') >= 0 || campaignKey.indexOf('NDS') >= 0;
    const isRes = campaignKey === 'att-res';
    if (isB2B && hasApi) fetchPromises.headcount  = this._fetchWithTimeout(this._fetchB2BHeadcount());
    if (isB2B && hasApi) fetchPromises.production = this._fetchWithTimeout(this._fetchB2BProduction());
    if (isNDS && hasApi) fetchPromises.ndsHeadcount  = this._fetchWithTimeout(this._fetchNDSHeadcount());
    if (isNDS && hasApi) fetchPromises.ndsProduction = this._fetchWithTimeout(this._fetchNDSProduction());
    if (isRes && hasApi) fetchPromises.d2dResRanking = this._fetchWithTimeout(this._fetchD2DResRanking());
    // Indeed/recruiting costs excluded from initial load — fetched only via Import button

    const keys = Object.keys(fetchPromises);
    const settled = await Promise.allSettled(keys.map(k => fetchPromises[k]));
    const results = {};
    keys.forEach((k, i) => {
      if (settled[i].status === 'fulfilled') {
        results[k] = settled[i].value;
      } else {
        console.warn(`[NationalApp] ${k} fetch failed:`, settled[i].reason?.message);
      }
    });

    // ── Build owners from recruiting data ──
    const sheetData = results.recruiting || null;
    console.log('[NationalApp] recruiting result for', campaignKey, ':',
      sheetData ? { owners: sheetData.owners?.length, weeks: sheetData.weeks?.length, products: sheetData.products } : 'null');
    if (sheetData && sheetData.owners && sheetData.owners.length) {
      this._buildOwnersFromSheet(campaignKey, sheetData);
      console.log('[NationalApp] Built owners for', campaignKey, ':', this.state.owners.length);
    } else {
      console.log('[NationalApp] No recruiting data for', campaignKey, '— showing empty state');
      this.state.owners = [];
      this.state.campaignTotals = {};
    }

    // ── Apply enrichments from parallel results ──
    if (results.camMapping) {
      this.state.camMapping = results.camMapping.mapping || results.camMapping;
    }

    if (results.audit && results.audit.businesses && results.audit.businesses.length) {
      this.state.allCompanyNames = results.audit.allCompanyNames || [];
      this._cachedAuditBusinesses = results.audit.businesses;
      this._mapAuditToOwners(results.audit.businesses, this.state.camMapping || null);
    }

    // Legacy B2B/NDS headcount enrichment
    if (results.headcount && results.headcount.owners && Object.keys(results.headcount.owners).length) {
      this._enrichOwnersWithNLR(results.headcount.owners);
    }

    if (results.ndsHeadcount && results.ndsHeadcount.owners && Object.keys(results.ndsHeadcount.owners).length) {
      this._enrichOwnersWithNLR(results.ndsHeadcount.owners);
    }

    if (results.ndsProduction && results.ndsProduction.owners && Object.keys(results.ndsProduction.owners).length) {
      this._enrichOwnersWithProduction(results.ndsProduction.owners);
    }

    if (results.production && results.production.owners && Object.keys(results.production.owners).length) {
      this._enrichOwnersWithProduction(results.production.owners);
    }

    // D2D Res ranking enrichment (att-res campaign only)
    if (results.d2dResRanking && results.d2dResRanking.ranking && results.d2dResRanking.ranking.length) {
      this._enrichOwnersWithD2DRanking(results.d2dResRanking.ranking);
    }

    // Pre-fetch Indeed Tracking data for configured owners (non-blocking)
    this._prefetchIndeedTracking();
  },

  // ── Pre-fetch Indeed Tracking for all owners in background ──
  async _prefetchIndeedTracking() {
    for (const owner of this.state.owners) {
      if (owner.indeedTracking || owner._trackingFetched || owner._trackingFetching) continue;
      owner._trackingFetching = true;
      try {
        const url = NATIONAL_CONFIG.appsScriptUrl +
          '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
          '&action=indeedTracking' +
          '&owner=' + encodeURIComponent(owner.name) +
          '&_t=' + Date.now();
        const resp = await this._fetchWithTimeout(fetch(url), 45000);
        const result = await resp.json();
        if (!result.error && result.weeks && result.weeks.length) {
          owner.indeedTracking = result;
          owner._trackingFetched = true;
          console.log('[NationalApp] Pre-fetched Indeed Tracking for', owner.name, ':', result.weeks.length, 'weeks');
        }
      } catch (err) {
        console.warn('[NationalApp] Pre-fetch Indeed Tracking for', owner.name, ':', err.message);
      }
      owner._trackingFetching = false;
    }
  },

  // ── Planning schedule: localStorage cache + background refresh ──
  _PLANNING_CACHE_KEY: 'od_planning_cache',

  _loadPlanningFromCache() {
    try {
      const raw = localStorage.getItem(this._PLANNING_CACHE_KEY);
      if (!raw) return false;
      const cache = JSON.parse(raw);
      if (cache && Array.isArray(cache.planning)) {
        this._planningSchedule = cache.planning;
        return true;
      }
    } catch { /* ignore */ }
    return false;
  },

  _savePlanningToCache(planning) {
    try {
      localStorage.setItem(this._PLANNING_CACHE_KEY, JSON.stringify({ planning, _ts: Date.now() }));
    } catch { /* ignore */ }
  },

  // ── NEW data badge logic (NLR weekly / Cam monthly) ──
  _TAB_VIEWED_KEY: 'od_tab_viewed',

  /**
   * Get the stored "last viewed" timestamps { recruiting: ts, audit: ts }
   */
  _getTabViewed(ownerName, campaign) {
    try {
      const all = JSON.parse(localStorage.getItem(this._TAB_VIEWED_KEY) || '{}');
      const key = (campaign + '::' + ownerName).toLowerCase();
      return all[key] || {};
    } catch { return {}; }
  },

  /**
   * Save a "viewed" timestamp for an owner+campaign+tab
   */
  _markTabViewed(ownerName, campaign, tab) {
    try {
      const all = JSON.parse(localStorage.getItem(this._TAB_VIEWED_KEY) || '{}');
      const key = (campaign + '::' + ownerName).toLowerCase();
      if (!all[key]) all[key] = {};
      all[key][tab] = Date.now();
      localStorage.setItem(this._TAB_VIEWED_KEY, JSON.stringify(all));
    } catch { /* ignore */ }
  },

  /**
   * Get the scheduled day (0=Mon … 6=Sun) for a campaign from planning data.
   * Returns -1 if not scheduled.
   */
  _getCampaignScheduledDay(campaignKey) {
    const sched = this._planningSchedule || [];
    const entry = sched.find(p => p.campaignKey === campaignKey);
    return entry ? entry.day : -1;
  },

  /**
   * Check whether the NLR "NEW" badge should show for this owner.
   * Shows on the campaign's scheduled day each week, clears once the coach views recruiting tab.
   */
  _shouldShowNlrBadge(ownerName, campaign) {
    // Owner must have NLR mapping
    if (this._isNonPartner(ownerName, 'nlrWorkbookId') || this._isUnmapped(ownerName, 'nlrTab')) return false;

    const scheduledDay = this._getCampaignScheduledDay(campaign);
    if (scheduledDay < 0) return false;

    const now = new Date();
    const todayIdx = (now.getDay() + 6) % 7; // Mon=0
    if (todayIdx !== scheduledDay) return false;

    // Check if coach already viewed recruiting tab today
    const viewed = this._getTabViewed(ownerName, campaign);
    if (viewed.recruiting) {
      const viewedDate = new Date(viewed.recruiting);
      if (viewedDate.toDateString() === now.toDateString()) return false; // already viewed today
    }
    return true;
  },

  /**
   * Check whether the Cam/BIS "NEW" badge should show for this owner.
   * Shows on the FIRST scheduled day of each calendar month, clears once coach views audit tab.
   */
  _shouldShowCamBadge(ownerName, campaign) {
    // Owner must have Cam mapping
    if (this._isNonPartner(ownerName, 'camCompany') || this._isUnmapped(ownerName, 'camCompany')) return false;

    const scheduledDay = this._getCampaignScheduledDay(campaign);
    if (scheduledDay < 0) return false;

    const now = new Date();
    const todayIdx = (now.getDay() + 6) % 7;
    if (todayIdx !== scheduledDay) return false;

    // Is this the FIRST occurrence of scheduledDay in this calendar month?
    const dayOfMonth = now.getDate();
    if (dayOfMonth > 7) return false; // first occurrence is always in days 1-7

    // Check if coach already viewed audit tab this month
    const viewed = this._getTabViewed(ownerName, campaign);
    if (viewed.audit) {
      const viewedDate = new Date(viewed.audit);
      if (viewedDate.getMonth() === now.getMonth() && viewedDate.getFullYear() === now.getFullYear()) return false;
    }
    return true;
  },

  /**
   * Update the NLR/Cam badges on the detail tab buttons for the current owner.
   */
  _updateDetailTabBadges(ownerName, campaign) {
    const nlrBadge = document.getElementById('tab-badge-recruiting');
    const camBadge = document.getElementById('tab-badge-audit');
    if (nlrBadge) nlrBadge.style.display = this._shouldShowNlrBadge(ownerName, campaign) ? '' : 'none';
    if (camBadge) camBadge.style.display = this._shouldShowCamBadge(ownerName, campaign) ? '' : 'none';
  },

  async _fetchPlanningSchedule() {
    try {
      const url = new URL(NATIONAL_CONFIG.appsScriptUrl || OD_CONFIG.appsScriptUrl);
      url.searchParams.set('key', NATIONAL_CONFIG.apiKey || OD_CONFIG.apiKey);
      url.searchParams.set('action', 'odGetPlanning');
      const res = await fetch(url.toString()).then(r => r.json());
      if (res.success && res.planning) {
        this._planningSchedule = res.planning;
        this._savePlanningToCache(res.planning);
        // Re-render landing if visible
        const landing = document.getElementById('campaign-landing');
        if (landing && landing.style.display !== 'none') {
          this._showLandingPage();
        }
      }
    } catch (err) {
      console.warn('[NationalApp] Failed to fetch planning schedule:', err.message);
    }
  },

  // ── Fetch ALL recruiting data from Ken's national sheet via NationalCode.gs ──
  // Returns the full campaigns dict. Caches in this._allCampaignsData.
  async _fetchRecruitingFromSheet(campaignKey) {
    // If we have cached data WITH actual week data, return it.
    // (Lightweight landing-page-only entries have empty weeks — must not short-circuit here)
    if (this._allCampaignsData && this._allCampaignsData[campaignKey]
        && this._allCampaignsData[campaignKey].weeks
        && this._allCampaignsData[campaignKey].weeks.length > 0) {
      const cd = this._allCampaignsData[campaignKey];
      console.log('[NationalApp] Cache HIT for', campaignKey, '— owners:', (cd.owners||[]).length, 'weeks:', (cd.weeks||[]).length);
      return { owners: cd.owners || [], weeks: cd.weeks || [], label: cd.label || '', products: cd.products || ['Total'] };
    }

    const url = NATIONAL_CONFIG.appsScriptUrl +
      '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
      '&action=recruiting&weeks=0' +
      '&_t=' + Date.now();
    const resp = await fetch(url);
    const result = await resp.json();
    if (result.error) throw new Error(result.error);

    // Cache ALL campaigns data
    if (result.campaigns) {
      this._allCampaignsData = result.campaigns;
      console.log('[NationalApp] Cached campaigns:', Object.keys(result.campaigns), 'looking for:', campaignKey);
      // Dynamically populate campaign selector and config
      this._populateCampaignSelector(result.campaigns);
    }

    // Extract the campaign-specific data
    const campaignData = result.campaigns && result.campaigns[campaignKey];
    if (!campaignData) {
      console.warn('[NationalApp] Campaign not found in response:', campaignKey, 'available:', Object.keys(result.campaigns || {}));
      return null;
    }

    return {
      owners: campaignData.owners || [],
      weeks: campaignData.weeks || [],
      label: campaignData.label || '',
      products: campaignData.products || ['Total']
    };
  },

  // ── Dynamically populate campaign selector dropdown from backend data ──
  _populateCampaignSelector(campaigns) {
    const select = document.getElementById('campaign-select');
    if (!select) return;

    // Sort campaign keys alphabetically by display label
    const keys = Object.keys(campaigns).sort((a, b) => {
      const labelA = (campaigns[a].label || a).toLowerCase();
      const labelB = (campaigns[b].label || b).toLowerCase();
      return labelA.localeCompare(labelB);
    });

    select.innerHTML = '';
    keys.forEach(key => {
      const opt = document.createElement('option');
      opt.value = key;
      opt.textContent = campaigns[key].label || key;
      select.appendChild(opt);

      // Ensure config has an entry for this campaign (so loadCampaignData doesn't throw)
      if (!NATIONAL_CONFIG.campaigns[key]) {
        NATIONAL_CONFIG.campaigns[key] = {
          label: campaigns[key].label || key,
          sectionHeader: campaigns[key].label || key,
          weeksToPull: 6
        };
      }
    });
    select.value = this.state.campaign;
  },

  // ── Fetch owner → Cam company mapping + cost sheet assignments from _OwnerCamMapping tab ──
  async _fetchOwnerCamMapping() {
    const url = NATIONAL_CONFIG.appsScriptUrl +
      '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
      '&action=ownerCamMapping' +
      '&_t=' + Date.now();
    const resp = await fetch(url);
    const result = await resp.json();
    if (result.error) throw new Error(result.error);
    return { mapping: result.mapping || {} };
  },

  // ── Fetch online presence data from Cam's Performance Audit sheet ──
  async _fetchOnlinePresence() {
    const url = NATIONAL_CONFIG.appsScriptUrl +
      '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
      '&action=onlinePresence' +
      '&_t=' + Date.now();
    const resp = await fetch(url);
    const result = await resp.json();
    if (result.error) throw new Error(result.error);
    return result;
  },

  // ── Lazy-load single owner's NLR data from mapped workbook ──
  async _fetchOwnerNlrData(owner) {
    owner._nlrFetched = true; // prevent duplicate fetches
    try {
      const url = NATIONAL_CONFIG.appsScriptUrl +
        '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
        '&action=ownerNlrData' +
        '&owner=' + encodeURIComponent(owner.name) +
        '&campaign=' + encodeURIComponent(this.state.campaign || '') +
        '&_t=' + Date.now();
      const resp = await fetch(url);
      const result = await resp.json();

      if (result.error) {
        console.warn('[NationalApp] NLR data error for', owner.name, ':', result.error);
        return;
      }
      if (!result.mapped) {
        console.log('[NationalApp] No NLR mapping for', owner.name);
        return;
      }
      if (!result.trend || result.trend.length === 0) {
        console.log('[NationalApp] NLR mapped but no data for', owner.name);
        return;
      }

      // Store NLR data on the owner object
      console.log('[NationalApp] NLR data loaded for', owner.name, ':', result.trend.length, 'weeks');
      owner.nlrData = result.trend;

      // Re-render if this owner is still selected
      if (this.state.selectedOwner === owner) {
        if (this.state.currentTab === 'recruiting') {
          this.renderRecruitingTab(owner);
        }
      }
    } catch (err) {
      console.warn('[NationalApp] NLR fetch failed for', owner.name, ':', err.message);
    }
  },

  // ── Fetch B2B headcount/production from local _B2B_Headcount tab ──
  async _fetchB2BHeadcount() {
    const url = NATIONAL_CONFIG.appsScriptUrl +
      '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
      '&action=b2bHeadcount' +
      '&_t=' + Date.now();
    const resp = await fetch(url);
    const result = await resp.json();
    if (result.error) throw new Error(result.error);
    return result;
  },

  // ── Fetch NDS headcount from local _NDS_Headcount tab ──
  async _fetchNDSHeadcount() {
    const url = NATIONAL_CONFIG.appsScriptUrl +
      '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
      '&action=ndsHeadcount' +
      '&_t=' + Date.now();
    const resp = await fetch(url);
    const result = await resp.json();
    if (result.error) throw new Error(result.error);
    return result;
  },

  // ── Fetch NDS production/sales from NDS One-on-Ones sheet ──
  async _fetchNDSProduction() {
    const url = NATIONAL_CONFIG.appsScriptUrl +
      '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
      '&action=ndsProduction' +
      '&_t=' + Date.now();
    const resp = await fetch(url);
    const result = await resp.json();
    if (result.error) throw new Error(result.error);
    return result;
  },

  // ── Fetch AT&T Res sales data for a single owner on demand ──
  async _fetchResOwnerSales(owner) {
    owner._resSalesFetching = true;
    try {
      const url = NATIONAL_CONFIG.appsScriptUrl +
        '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
        '&action=resOwnerSales' +
        '&owner=' + encodeURIComponent(owner.name) +
        '&_t=' + Date.now();
      const resp = await this._fetchWithTimeout(fetch(url), 30000);
      const result = await resp.json();
      // Clear fetching flag BEFORE re-render so renderSalesTab doesn't show spinner again
      owner._resSalesFetching = false;
      if (!result.error && (result.summary || result.reps)) {
        owner.sales = { summary: result.summary, reps: result.reps || [] };
        owner._resSalesFetched = true;
        console.log('[NationalApp] Res sales loaded for', owner.name, ':', result.reps?.length, 'reps, tab:', result.tab);
      } else {
        console.warn('[NationalApp] Res sales empty for', owner.name, ':', result.error || 'no data');
      }
    } catch (err) {
      console.warn('[NationalApp] Res sales fetch failed for', owner.name, ':', err.message);
      owner._resSalesFetching = false;
    }
    // Always re-render if user is still on sales tab (handles both success + error/empty)
    if (this.state.selectedOwner === owner && this.state.currentTab === 'sales') {
      this.renderSalesTab(owner);
    }
  },

  // ── Fetch NDS sales data for a single owner on demand ──
  async _fetchNDSOwnerSales(owner) {
    owner._ndsSalesFetching = true;
    try {
      const url = NATIONAL_CONFIG.appsScriptUrl +
        '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
        '&action=ndsOwnerSales' +
        '&owner=' + encodeURIComponent(owner.name) +
        '&_t=' + Date.now();
      const resp = await this._fetchWithTimeout(fetch(url), 30000);
      const result = await resp.json();
      // Clear fetching flag BEFORE re-render so renderSalesTab doesn't show spinner again
      owner._ndsSalesFetching = false;
      if (!result.error && (result.summary || result.reps)) {
        owner.sales = { summary: result.summary, reps: result.reps || [] };
        owner._ndsSalesFetched = true;
        console.log('[NationalApp] NDS sales loaded for', owner.name, ':', result.reps?.length, 'reps, tab:', result.tab);
      } else {
        console.warn('[NationalApp] NDS sales empty for', owner.name, ':', result.error || 'no data');
      }
    } catch (err) {
      console.warn('[NationalApp] NDS sales fetch failed for', owner.name, ':', err.message);
      owner._ndsSalesFetching = false;
    }
    // Always re-render if user is still on sales tab (handles both success + error/empty)
    if (this.state.selectedOwner === owner && this.state.currentTab === 'sales') {
      this.renderSalesTab(owner);
    }
  },

  // ── Fetch B2B production/sales data from local _B2B_Production tab ──
  async _fetchB2BProduction() {
    const url = NATIONAL_CONFIG.appsScriptUrl +
      '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
      '&action=b2bProduction' +
      '&_t=' + Date.now();
    const resp = await fetch(url);
    const result = await resp.json();
    if (result.error) throw new Error(result.error);
    return result;
  },

  // ── Enrich B2B owners with NLR production/sales data ──
  _enrichOwnersWithProduction(prodOwners) {
    // Build lowercase lookup of production tab names
    const prodKeys = Object.keys(prodOwners);
    const prodLower = {};
    for (const key of prodKeys) {
      prodLower[key.toLowerCase()] = key;
    }

    let matched = 0;
    for (const owner of this.state.owners) {
      const ownerLc = owner.name.toLowerCase().trim();

      // Try exact match first
      let prodKey = prodLower[ownerLc];

      // Try starts-with
      if (!prodKey) {
        for (const lc in prodLower) {
          if (lc.startsWith(ownerLc) || ownerLc.startsWith(lc)) {
            prodKey = prodLower[lc];
            break;
          }
        }
      }

      // Try contains
      if (!prodKey) {
        for (const lc in prodLower) {
          if (lc.indexOf(ownerLc) >= 0 || ownerLc.indexOf(lc) >= 0) {
            prodKey = prodLower[lc];
            break;
          }
        }
      }

      if (!prodKey) continue;

      const prod = prodOwners[prodKey];
      if (prod.summary) {
        owner.sales.summary = prod.summary;
      }
      if (prod.reps && prod.reps.length) {
        owner.sales.reps = prod.reps;
      }
      matched++;
    }
    console.log('[NationalApp] Production enrichment: matched', matched, 'of', this.state.owners.length, 'owners');
  },

  // ── Fetch D2D Residential ranking from _D2D_Res_Ranking tab ──
  async _fetchD2DResRanking() {
    const url = NATIONAL_CONFIG.appsScriptUrl +
      '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
      '&action=d2dResRanking' +
      '&_t=' + Date.now();
    const resp = await fetch(url);
    const result = await resp.json();
    if (result.error) throw new Error(result.error);
    return result;
  },

  // ── Enrich owners with D2D Res ranking data ──
  // Fuzzy-matches ranking owner names to campaign owner names and sets o.d2dRank
  // Aliases handle cases where Tableau uses a different name than the campaign sheet
  _D2D_RANK_ALIASES: {
    'wayne rude': 'flloyd (wayne) rude',
    'tre mitchell': 'lamar (tre) mitchell iii',
    'john richard young': 'jr (john richard) young'
  },

  _enrichOwnersWithD2DRanking(ranking) {
    // Build lowercase lookup: ranking owner name → ranking entry
    const rankLower = {};
    for (const entry of ranking) {
      const key = entry.owner.toLowerCase().trim();
      rankLower[key] = entry;
      // Also register under alias if one exists
      const alias = this._D2D_RANK_ALIASES[key];
      if (alias) rankLower[alias] = entry;
    }

    let matched = 0;
    for (const owner of this.state.owners) {
      const ownerLc = owner.name.toLowerCase().trim();

      // Exact match (includes aliases)
      let entry = rankLower[ownerLc];

      // Starts-with (full string)
      if (!entry) {
        for (const lc in rankLower) {
          if (lc.startsWith(ownerLc) || ownerLc.startsWith(lc)) { entry = rankLower[lc]; break; }
        }
      }

      // Last name exact + first name prefix (handles Jess↔Jessica, Mike↔Michael, etc.)
      if (!entry) {
        const ownerParts = ownerLc.split(/\s+/);
        if (ownerParts.length >= 2) {
          const ownerFirst = ownerParts[0];
          const ownerLast = ownerParts[ownerParts.length - 1];
          for (const lc in rankLower) {
            const parts = lc.split(/\s+/);
            if (parts.length < 2) continue;
            const rFirst = parts[0];
            const rLast = parts[parts.length - 1];
            if (rLast === ownerLast && (rFirst.startsWith(ownerFirst) || ownerFirst.startsWith(rFirst))) {
              entry = rankLower[lc]; break;
            }
          }
        }
      }

      // Contains (full string)
      if (!entry) {
        for (const lc in rankLower) {
          if (lc.indexOf(ownerLc) >= 0 || ownerLc.indexOf(lc) >= 0) { entry = rankLower[lc]; break; }
        }
      }

      if (entry) {
        owner.d2dRank = entry.rank;
        owner.d2dTotalUnits = entry.totalUnits;
        owner.d2dProducts = entry.products || {};
        matched++;
      }
    }
    console.log('[NationalApp] D2D Res ranking enrichment: matched', matched, 'of', this.state.owners.length, 'owners');
  },

  // ── Fetch Indeed ad cost data from per-owner Drive spreadsheets ──
  // Passes current campaign's owner names so server only opens matching files
  // (_fetchIndeedCosts removed — costs now fetched per-owner in _loadAndRenderCosts)

  // ── Enrich owners with Indeed ad cost data ──
  _enrichOwnersWithIndeedCosts(indeedOwners) {
    const keys = Object.keys(indeedOwners);
    const lower = {};
    for (const key of keys) lower[key.toLowerCase().trim()] = key;

    let matched = 0;
    for (const owner of this.state.owners) {
      const ownerLc = owner.name.toLowerCase().trim();

      // Exact match
      let matchKey = lower[ownerLc];

      // Starts-with
      if (!matchKey) {
        for (const lc in lower) {
          if (lc.startsWith(ownerLc) || ownerLc.startsWith(lc)) { matchKey = lower[lc]; break; }
        }
      }

      // Contains
      if (!matchKey) {
        for (const lc in lower) {
          if (lc.indexOf(ownerLc) >= 0 || ownerLc.indexOf(lc) >= 0) { matchKey = lower[lc]; break; }
        }
      }

      if (!matchKey) continue;
      matched++;
    }
    console.log('[NationalApp] Indeed costs enrichment: matched', matched, 'of', this.state.owners.length, 'owners');
  },

  // ── Import NLR headcount data into local sheet (one-time sync) ──
  async importNLRHeadcount() {
    const btn = document.getElementById('btn-import-nlr');
    if (btn) { btn.disabled = true; btn.textContent = 'Importing...'; }

    try {
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'importNLRHeadcount'
        })
      });
      const result = await resp.json();
      if (result.error) throw new Error(result.error);

      console.log('[NationalApp] NLR import complete:', result);
      if (btn) { btn.textContent = `Imported ${result.ownersImported} owners (${result.rowsWritten} rows)`; }

      // Reload to pick up the freshly imported data
      await this.loadCampaignData(this.state.campaign);
      this.renderDashboard();
    } catch (err) {
      console.error('[NationalApp] NLR import failed:', err);
      if (btn) { btn.disabled = false; btn.textContent = 'Import Failed — Retry'; }
    }
  },

  // ── Enrich B2B owners with NLR headcount/production data ──
  // Uses fuzzy name matching: NLR tabs have full names (e.g., "Alex Badawi"),
  // while recruiting data may have short names (e.g., "Jay T", "Mason").
  // Match by: exact match → starts-with → contains (case-insensitive).
  _enrichOwnersWithNLR(nlrOwners) {
    // Build lowercase lookup of NLR tab names
    const nlrKeys = Object.keys(nlrOwners);
    const nlrLower = {};
    for (const key of nlrKeys) {
      nlrLower[key.toLowerCase()] = key;
    }

    let matched = 0;
    for (const owner of this.state.owners) {
      const ownerLc = owner.name.toLowerCase().trim();

      // Try exact match first
      let nlrKey = nlrLower[ownerLc];

      // Try starts-with: NLR "Justin Wood" starts with recruiting "Justin"
      if (!nlrKey) {
        for (const lc in nlrLower) {
          if (lc.startsWith(ownerLc) || ownerLc.startsWith(lc)) {
            nlrKey = nlrLower[lc];
            break;
          }
        }
      }

      // Try contains: recruiting "Nigel Gil" contained in NLR "Nigel Gilbert"
      if (!nlrKey) {
        for (const lc in nlrLower) {
          if (lc.indexOf(ownerLc) >= 0 || ownerLc.indexOf(lc) >= 0) {
            nlrKey = nlrLower[lc];
            break;
          }
        }
      }

      if (!nlrKey) continue;

      // Store the matched sheet name for write-back
      owner._sheetName = nlrKey;

      const nlr = nlrOwners[nlrKey];

      // Set current headcount from latest row
      owner.headcount.active = nlr.current.active || 0;
      owner.headcount.leaders = nlr.current.leaders || 0;
      owner.headcount.training = nlr.current.training || 0;

      // Tie leader headcount to recruiting tab leader count
      if (owner.headcount.leaders && owner.recruiting) {
        owner.recruiting.leaders = owner.headcount.leaders;
        if (owner.recruiting.rows.length) {
          const actuals = owner.recruiting.rows.map(r => r.values);
          owner.recruiting.rows = this._buildRows(owner.headcount.leaders, actuals);
        }
      }

      // Set current production from latest row (per-category breakdown)
      // BUT skip if the newest consolidated week has no production — that means
      // the coach needs to manually enter it, and we shouldn't override with NLR data
      if (!owner._newestWeekMissingProd) {
        const prod = nlr.current.production;
        if (prod && typeof prod === 'object') {
          let totalP = 0, totalG = 0;
          for (const pName in prod) {
            owner.production.products[pName] = {
              actual: prod[pName].production || 0,
              goal: prod[pName].goals || 0
            };
            totalP += prod[pName].production || 0;
            totalG += prod[pName].goals || 0;
          }
          owner.production.totalActual = totalP;
          owner.production.totalGoal = totalG;
        } else {
          // Fallback: single total (backward compat)
          owner.production.totalActual = nlr.current.productionLW || 0;
          owner.production.totalGoal = nlr.current.productionGoals || 0;
        }
      }

      // Build headcount history from ALL trend rows
      owner.headcountHistory = (nlr.trend || []).map(row => ({
        date: row.date,
        active: row.active,
        leaders: row.leaders,
        training: row.training
      }));

      // Build production history from ALL trend rows with per-product data
      owner.productionHistory = (nlr.trend || []).map(row => {
        const entry = {
          date: row.date,
          tA: row.productionLW || 0,
          tG: row.productionGoals || 0,
          products: {}
        };
        if (row.production && typeof row.production === 'object') {
          for (const pName in row.production) {
            entry.products[pName] = {
              actual: row.production[pName].production || 0,
              goal: row.production[pName].goals || 0
            };
          }
        }
        return entry;
      });

      matched++;
    }
    console.log('[NationalApp] NLR headcount enrichment: matched', matched, 'of', this.state.owners.length, 'owners');
  },

  // ── Map online presence businesses to owners ──
  // Uses _OwnerCamMapping tab (Owner Name → Cam Company Name) for exact matching.
  // Falls back to old alias/partial matching if no mapping is available.
  _mapAuditToOwners(businesses, camMapping) {
    // Build owner lookup: lowercase name → owner object
    const ownerMap = {};
    for (const owner of this.state.owners) {
      ownerMap[owner.name.toLowerCase()] = owner;
      if (owner.tab && owner.tab.toLowerCase() !== owner.name.toLowerCase()) {
        ownerMap[owner.tab.toLowerCase()] = owner;
      }
    }

    // Build reverse lookup from mapping: lowercase Cam company name → owner object
    const companyToOwner = {};
    if (camMapping && typeof camMapping === 'object') {
      for (const [ownerName, companies] of Object.entries(camMapping)) {
        const owner = ownerMap[ownerName.toLowerCase()];
        if (!owner) continue;
        for (const company of companies) {
          companyToOwner[company.toLowerCase().trim()] = owner;
        }
      }
      console.log('[NationalApp] Cam mapping loaded:', Object.keys(companyToOwner).length, 'companies →', Object.keys(camMapping).length, 'owners');
    }

    const hasMapping = Object.keys(companyToOwner).length > 0;
    const unmatched = [];

    for (const biz of businesses) {
      const bizName = (biz.businessName || biz.clientName || '').toLowerCase().trim();
      const clientKey = (biz.clientName || '').toLowerCase().trim();
      let owner = null;

      if (hasMapping) {
        // Use mapping: match on business name or client name
        owner = companyToOwner[bizName] || companyToOwner[clientKey] || null;
      } else {
        // Fallback: alias + partial matching (legacy behavior)
        const aliases = NATIONAL_CONFIG.ownerAliases || {};
        owner = ownerMap[clientKey]
          || (aliases[clientKey] && ownerMap[aliases[clientKey].toLowerCase()])
          || null;
        if (!owner) {
          for (const o of this.state.owners) {
            const oLower = o.name.toLowerCase();
            if (clientKey.indexOf(oLower) >= 0 || oLower.indexOf(clientKey) >= 0) {
              owner = o;
              break;
            }
          }
        }
      }

      if (owner) {
        if (!owner.audit.businesses) owner.audit.businesses = [];
        owner.audit.businesses.push(biz);
      } else {
        unmatched.push(biz.clientName + (biz.businessName ? ' / ' + biz.businessName : ''));
      }
    }

    // Calculate grades for each owner from their businesses
    for (const owner of this.state.owners) {
      if (owner.audit.businesses && owner.audit.businesses.length) {
        this._calculateAuditGrades(owner);
      }
    }
  },

  // ── Calculate audit grades from business data ──
  _calculateAuditGrades(owner) {
    const bizList = owner.audit.businesses;
    const total = bizList.length;

    // Reviews: average GBL rating + total review count (for A+ threshold)
    let ratingSum = 0, ratingCount = 0, totalReviews = 0;
    for (const b of bizList) {
      if (b.gbl && b.gbl.rating != null && b.gbl.rating > 0) {
        ratingSum += b.gbl.rating;
        ratingCount++;
        totalReviews += (b.gbl.reviews || 0);
      }
    }
    if (ratingCount > 0) {
      const avg = ratingSum / ratingCount;
      owner.audit.grades.reviews = this._ratingToGrade(avg, totalReviews);
      owner.audit.reviewsAvg = Math.round(avg * 10) / 10;
      owner.audit.reviewsCount = ratingCount;
    }

    // Website: aggregate per-business grades → modal (most common) grade
    const websiteGrades = [];
    for (const b of bizList) {
      const g = this._bizWebsiteGrade(b);
      if (g !== '—') websiteGrades.push(g);
    }
    owner.audit.grades.website = this._modalGrade(websiteGrades);
    owner.audit.websiteUpdated = websiteGrades.length;

    // Social: aggregate per-business grades → modal grade
    const socialGrades = [];
    for (const b of bizList) {
      const g = this._bizSocialGrade(b);
      if (g !== '—') socialGrades.push(g);
    }
    owner.audit.grades.social = this._modalGrade(socialGrades);
    owner.audit.igCount = socialGrades.length;

    // SEO: % passing
    let seoPassing = 0;
    for (const b of bizList) {
      const val = (b.seo?.check || '').toLowerCase();
      if (val === 'pass' || val === 'yes' || val === 'y' || val === '✓' || val === 'true' || val === 'x' || val === 'good') {
        seoPassing++;
      }
    }
    owner.audit.grades.seo = this._pctToGrade(seoPassing, total);
    owner.audit.seoPassing = seoPassing;
  },

  // Matches Cam's Performance Audit grading formula
  _ratingToGrade(rating, reviewCount) {
    if (rating >= 4.9 && reviewCount > 49) return 'A+';
    if (rating >= 4.4) return 'A';
    if (rating >= 4.0) return 'A-';
    if (rating >= 3.5) return 'B';
    if (rating >= 3.2) return 'B-';
    if (rating >= 2.8) return 'C';
    if (rating >= 2.5) return 'C-';
    return 'D';
  },

  _pctToGrade(count, total) {
    if (!total) return '—';
    const pct = count / total;
    if (pct >= 0.9) return 'A';
    if (pct >= 0.7) return 'B';
    if (pct >= 0.5) return 'C';
    if (pct >= 0.3) return 'D';
    return 'F';
  },

  // Aggregate an array of letter grades into a single owner-level grade
  // Uses grade-point average: A+=4.3, A=4, A-=3.7, B=3, B-=2.7, C=2.5, C-=2.3, D=1
  _modalGrade(grades) {
    if (!grades.length) return '—';
    const gpa = { 'A+': 4.3, 'A': 4, 'A-': 3.7, 'B': 3, 'B-': 2.7, 'C': 2.5, 'C-': 2.3, 'D': 1, 'F': 0 };
    let sum = 0;
    for (const g of grades) sum += (gpa[g] ?? 0);
    const avg = sum / grades.length;
    if (avg >= 4.15) return 'A+';
    if (avg >= 3.85) return 'A';
    if (avg >= 3.35) return 'A-';
    if (avg >= 2.85) return 'B';
    if (avg >= 2.6)  return 'B-';
    if (avg >= 2.4)  return 'C';
    if (avg >= 2.15) return 'C-';
    return 'D';
  },

  // ── Build owner objects from national sheet data ──
  _buildOwnersFromSheet(campaignKey, sheetData) {
    const ownerNames = sheetData.owners;
    const allWeeks = sheetData.weeks || [];

    // Store campaign product list for rendering
    this.state.campaignProducts = sheetData.products || ['Total'];

    // Campaign table: 4 most recent weeks where the week date is not in the future
    const now = new Date();
    const todayEnd = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
    const pastWeeks = allWeeks.filter(w => {
      const parts = w.tabName.split('/');
      if (parts.length !== 3) return true;
      const d = new Date(+parts[2], +parts[0] - 1, +parts[1]);
      return d <= todayEnd;
    });
    const campaignWeeks = pastWeeks.slice(0, 4).reverse();
    const campaignLabels = campaignWeeks.map(w => w.tabName);

    // Owner detail: ALL weeks, left-to-right chronological
    const allWeeksChron = [...allWeeks].reverse();
    const allLabels = allWeeksChron.map(w => w.tabName);


    this.state.owners = ownerNames.map(name => {
      // Helper: extract metrics array from owner data
      // Handles both new format {metrics:[...], health:{...}} and legacy array format
      const _getMetrics = (ownerData) => {
        if (!ownerData) return new Array(12).fill(0);
        if (Array.isArray(ownerData)) return ownerData; // legacy array format
        if (ownerData.metrics) return ownerData.metrics; // new object format
        return new Array(12).fill(0);
      };

      // Campaign-level actuals (4 weeks) for aggregation
      const actuals4 = Array.from({ length: 12 }, () => []);
      for (let wi = 0; wi < campaignWeeks.length; wi++) {
        const weekData = campaignWeeks[wi].data || {};
        const metrics = _getMetrics(weekData[name]);
        for (let ri = 0; ri < 12; ri++) {
          actuals4[ri].push(metrics[ri] || 0);
        }
      }

      // Full history actuals (all weeks) for owner detail tab
      const actualsFull = Array.from({ length: 12 }, () => []);
      for (let wi = 0; wi < allWeeksChron.length; wi++) {
        const weekData = allWeeksChron[wi].data || {};
        const metrics = _getMetrics(weekData[name]);
        for (let ri = 0; ri < 12; ri++) {
          actualsFull[ri].push(metrics[ri] || 0);
        }
      }

      // ── Extract health data from week data ──
      // New format: ownerData is {metrics:[...], health:{...}}
      // Legacy format: ownerData is array with .health property (lost in JSON)
      // Headcount and production are tracked separately:
      //   - Headcount: use newest week (even if all zeros = fresh week, inputs should be blank)
      //   - Production: use most recent week with actual data (for rankings/display)
      let latestHeadcount = {};  // headcount from newest week (may be zeros)
      let latestProdHealth = {}; // health from most recent week with production data
      let lastNonZeroLeaders = 0; // Track last non-zero leader count for recruiting table
      const hcHistory = [];
      const prodHistory = [];
      // allWeeks is newest-first; allWeeksChron is oldest-first
      // Check newest week first for headcount — if it has a health object (even all zeros),
      // that's the current headcount state (fresh week = blank inputs, not prefilled)
      const newestWeek = allWeeks[0];
      let headcountSetFromNewest = false;
      if (newestWeek) {
        const nwData = (newestWeek.data || {})[name];
        const nwHealth = nwData && (nwData.health || null);
        if (nwHealth) {
          latestHeadcount = nwHealth;
          headcountSetFromNewest = true;
        }
      }
      for (let wi = 0; wi < allWeeksChron.length; wi++) {
        const weekData = allWeeksChron[wi].data || {};
        const ownerData = weekData[name];
        const health = ownerData && (ownerData.health || null);
        if (health) {
          const h = health;
          // Headcount fallback: only if newest week didn't have this owner's data
          if (!headcountSetFromNewest && ((h.active || 0) > 0 || (h.leaders || 0) > 0 || (h.training || 0) > 0)) {
            latestHeadcount = h;
          }
          // Production: track the most recent week with actual production data
          // (allWeeksChron is oldest-first, so last match wins = newest with data)
          // Production should show last week's values even when headcount resets on a fresh week
          const hProd = h.production;
          if (hProd && typeof hProd === 'object') {
            const hasProd = Object.keys(hProd).some(k => k !== 'Total' && (hProd[k]?.production > 0 || hProd[k]?.goals > 0));
            if (hasProd) latestProdHealth = h;
          }
          if ((h.leaders || 0) > 0) lastNonZeroLeaders = h.leaders;
          const hcActive = h.active || 0, hcLeaders = h.leaders || 0, hcTraining = h.training || 0;
          const hcClosers = h.closers || 0, hcLeadGen = h.leadGen || 0;
          if (hcActive > 0 || hcLeaders > 0 || hcTraining > 0) {
            hcHistory.push({ date: allWeeksChron[wi].tabName, active: hcActive, leaders: hcLeaders, training: hcTraining, closers: hcClosers, leadGen: hcLeadGen });
          }
          // Production history — handle both old format (single numbers) and new format (per-product objects)
          // Always include the most recent week (even if all zeros) so coaches can fill it in
          const isNewestWeek = (wi === allWeeksChron.length - 1);
          const prod = h.production;
          if (prod && typeof prod === 'object' && !Array.isArray(prod)) {
            // New per-product format: { Frontier: {production, goals}, Cell: {production, goals}, ... }
            let totalProd = 0, totalGoal = 0;
            const entry = { date: allWeeksChron[wi].tabName, products: {} };
            for (const pName in prod) {
              if (pName === 'Total') continue;
              const pv = prod[pName];
              entry.products[pName] = { actual: pv.production || 0, goal: pv.goals || 0 };
              totalProd += pv.production || 0;
              totalGoal += pv.goals || 0;
            }
            // Sum across products (no separate Total column)
            entry.tA = totalProd;
            entry.tG = totalGoal;
            // Include newest week even if zeros (editable), skip older empty weeks
            if (totalProd > 0 || totalGoal > 0 || isNewestWeek) prodHistory.push(entry);
          } else if (isNewestWeek) {
            // Newest week with no production object at all — add empty entry so it's editable
            // Inherit product names from previous weeks so the cards know what columns to show
            const inheritProducts = {};
            for (let pi = prodHistory.length - 1; pi >= 0; pi--) {
              const prev = prodHistory[pi];
              if (prev.products && Object.keys(prev.products).length > 0) {
                for (const pn of Object.keys(prev.products)) inheritProducts[pn] = { actual: 0, goal: 0 };
                break;
              }
            }
            prodHistory.push({ date: allWeeksChron[wi].tabName, tA: 0, tG: 0, products: inheritProducts });
          } else {
            // Legacy single-value format (older weeks)
            const legA = (typeof prod === 'number') ? prod : 0;
            const legG = (typeof h.goals === 'number') ? h.goals : 0;
            if (legA > 0 || legG > 0) {
              prodHistory.push({ date: allWeeksChron[wi].tabName, tA: legA, tG: legG, products: {} });
            }
          }
        } else if (wi === allWeeksChron.length - 1) {
          // Newest week has no health data at all — add empty production entry so it's editable
          const inheritProducts = {};
          for (let pi = prodHistory.length - 1; pi >= 0; pi--) {
            const prev = prodHistory[pi];
            if (prev.products && Object.keys(prev.products).length > 0) {
              for (const pn of Object.keys(prev.products)) inheritProducts[pn] = { actual: 0, goal: 0 };
              break;
            }
          }
          prodHistory.push({ date: allWeeksChron[wi].tabName, tA: 0, tG: 0, products: inheritProducts });
        }
      }

      // Ensure the newest week is in prodHistory (may have been skipped if health
      // existed but production was null/empty, or if week index didn't match)
      const newestWeekDate = allWeeksChron.length > 0 ? allWeeksChron[allWeeksChron.length - 1].tabName : null;
      if (newestWeekDate && !prodHistory.some(p => p.date === newestWeekDate)) {
        const inheritProducts = {};
        for (let pi = prodHistory.length - 1; pi >= 0; pi--) {
          const prev = prodHistory[pi];
          if (prev.products && Object.keys(prev.products).length > 0) {
            for (const pn of Object.keys(prev.products)) inheritProducts[pn] = { actual: 0, goal: 0 };
            break;
          }
        }
        prodHistory.push({ date: newestWeekDate, tA: 0, tG: 0, products: inheritProducts });
      }

      // Build current production from most recent week with actual production data
      // (separate from headcount which uses the newest week even if all zeros)
      let currentProd = { totalGoal: 0, totalActual: 0, wirelessGoal: 0, wirelessActual: 0, products: {} };
      const lp = latestProdHealth.production;
      if (lp && typeof lp === 'object' && !Array.isArray(lp)) {
        let totalP = 0, totalG = 0;
        for (const pName in lp) {
          if (pName === 'Total') continue;
          currentProd.products[pName] = { actual: lp[pName].production || 0, goal: lp[pName].goals || 0 };
          totalP += lp[pName].production || 0;
          totalG += lp[pName].goals || 0;
        }
        currentProd.totalActual = totalP;
        currentProd.totalGoal = totalG;
      } else if (typeof lp === 'number') {
        currentProd.totalActual = lp;
        currentProd.totalGoal = (typeof latestProdHealth.goals === 'number') ? latestProdHealth.goals : 0;
      }
      // Check if the newest week is missing production data
      // If so, override currentProd to show empty/editable state instead of last week's values
      const newestProdEntry = prodHistory.length > 0 ? prodHistory[prodHistory.length - 1] : null;
      const newestWeekMissingProd = newestProdEntry && newestProdEntry.tA === 0 && newestProdEntry.tG === 0;
      if (newestWeekMissingProd) {
        currentProd = { totalGoal: 0, totalActual: 0, wirelessGoal: 0, wirelessActual: 0, products: {} };
        // Inherit product names from the entry (which inherited from prior weeks)
        if (newestProdEntry.products && Object.keys(newestProdEntry.products).length > 0) {
          for (const pn of Object.keys(newestProdEntry.products)) {
            currentProd.products[pn] = { actual: 0, goal: 0 };
          }
        }
      }
      // If currentProd has no products but prodHistory does, inherit product names
      // so editable cards show the right columns for the newest empty week
      if (!Object.keys(currentProd.products).length && prodHistory.length > 0) {
        const fallbackPH = prodHistory.find(p => p.products && Object.keys(p.products).length > 0);
        if (fallbackPH) {
          for (const pn of Object.keys(fallbackPH.products)) {
            currentProd.products[pn] = { actual: 0, goal: 0 };
          }
        }
      }

      return {
        name: name,
        tab: name,
        statusCode: null,
        headcount: {
          active: latestHeadcount.active || 0,
          leaders: latestHeadcount.leaders || 0,
          closers: latestHeadcount.closers || 0,
          leadGen: latestHeadcount.leadGen || 0,
          training: latestHeadcount.training || 0
        },
        headcountHistory: hcHistory,
        production: currentProd,
        productionHistory: prodHistory,
        _newestWeekMissingProd: !!newestWeekMissingProd,
        nextGoals: { totalUnits: 0, wirelessUnits: 0 },
        recruiting: {
          leaders: lastNonZeroLeaders || latestHeadcount.leaders || 0,
          weeks: campaignLabels,
          rows: this._buildRows(lastNonZeroLeaders || latestHeadcount.leaders || 0, actuals4)
        },
        recruitingFull: {
          leaders: lastNonZeroLeaders || latestHeadcount.leaders || 0,
          weeks: allLabels,
          rows: this._buildRows(lastNonZeroLeaders || latestHeadcount.leaders || 0, actualsFull)
        },
        sales: {
          summary: null,
          reps: []
        },
        audit: {
          grades: { reviews: '—', website: '—', social: '—', seo: '—' },
          details: {}
        },
        notes: ''
      };
    });

    // Store latest week date for overview display
    this._latestWeekDate = pastWeeks.length ? pastWeeks[0].tabName : null;

    // Campaign-level totals + aggregate recruiting (4 weeks only)
    this._buildCampaignAggregates(campaignLabels);

    // For non-att-res campaigns, rank owners by production (att-res uses Tableau ranking)
    if (this.state.campaign !== 'att-res') {
      this._rankOwnersByProduction();
    }
  },

  // ── Build campaign-level totals and aggregate recruiting table ──
  _buildCampaignAggregates(weekLabels) {
    const totals = this.state.owners.reduce((acc, o) => {
      acc.headcount += o.headcount.active;
      acc.leaders += o.headcount.leaders;
      acc.production += o.production.totalActual;
      return acc;
    }, { headcount: 0, leaders: 0, production: 0 });

    // Aggregate actuals across all owners
    const numWeeks = weekLabels.length || 1;
    const aggA = Array.from({ length: 12 }, () => new Array(numWeeks).fill(0));

    this.state.owners.forEach(o => {
      if (!o.recruiting.rows.length) return;
      o.recruiting.rows.forEach((row, ri) => {
        row.values.forEach((v, wi) => { aggA[ri][wi] += v; });
      });
    });

    // For rate rows, compute average instead of sum
    this.RECRUITING_LABELS.forEach((def, i) => {
      if (def.isRate) {
        const cnt = this.state.owners.filter(o => o.recruiting.rows.length).length || 1;
        aggA[i] = aggA[i].map(v => Math.round(v / cnt));
      }
    });

    this.state.campaignRecruiting = {
      leaders: totals.leaders,
      weeks: weekLabels,
      rows: this._buildRows(totals.leaders, aggA),
      showLegend: true
    };

    // KPI totals — enriched with breakdowns for campaign overview cards
    const crRows = this.state.campaignRecruiting.rows;
    const _t = (idx) => crRows[idx] ? crRows[idx].total : 0;

    // Aggregate headcount breakdown across all owners
    const hcBreakdown = this.state.owners.reduce((acc, o) => {
      acc.leaders += o.headcount.leaders || 0;
      acc.closers += o.headcount.closers || 0;
      acc.leadGen += o.headcount.leadGen || 0;
      acc.training += o.headcount.training || 0;
      acc.active += o.headcount.active || 0;
      return acc;
    }, { leaders: 0, closers: 0, leadGen: 0, training: 0, active: 0 });
    hcBreakdown.dist = Math.max(hcBreakdown.active - hcBreakdown.leaders, 0);

    // Aggregate production breakdown across all owners
    const prodBreakdown = {};
    this.state.owners.forEach(o => {
      if (!o.production || !o.production.products) return;
      for (const [pName, pData] of Object.entries(o.production.products)) {
        if (!prodBreakdown[pName]) prodBreakdown[pName] = 0;
        prodBreakdown[pName] += pData.actual || 0;
      }
    });

    this.state.campaignTotals = {
      latestWeek: this._latestWeekDate || null,
      headcount: totals.headcount,
      hcBreakdown,
      // 1st Rounds
      firstBooked: _t(2),
      firstShowed: _t(3),
      firstRetention: _t(2) ? Math.round((_t(3) / _t(2)) * 100) : 0,
      // 2nd Rounds
      secondBooked: _t(6),
      secondShowed: _t(7),
      secondRetention: _t(6) ? Math.round((_t(7) / _t(6)) * 100) : 0,
      // New Starts
      newStartsBooked: _t(9),
      newStartsShowed: _t(10),
      newStartRetention: _t(9) ? Math.round((_t(10) / _t(9)) * 100) : 0,
      // Production
      production: totals.production,
      prodBreakdown
    };
  },

  // ── Rank owners by most recent week's production (for non-att-res campaigns) ──
  _rankOwnersByProduction() {
    const sorted = [...this.state.owners]
      .map((o, idx) => ({ idx, prod: o.production?.totalActual || 0 }))
      .sort((a, b) => b.prod - a.prod);

    sorted.forEach((entry, rank) => {
      const owner = this.state.owners[entry.idx];
      owner.d2dRank = rank + 1;
      owner.d2dTotalUnits = entry.prod;
    });

    console.log('[NationalApp] Production ranking: top 3 —',
      sorted.slice(0, 3).map((e, i) => '#' + (i+1) + ' ' + this.state.owners[e.idx].name + ' (' + e.prod + ')').join(', '));
  },

  // ── Calculate projected weekly numbers from leader count ──
  _calcProjected(leaders) {
    const RET_1 = 0.50;        // 1st Retention
    const RET_2 = 0.60;        // 2nd Retention
    const RET_NS = 0.60;       // New Start Retention
    const CALL_LIST_BOOKED = 0.45;

    const secondsBooked = leaders * 5;
    const firstShowed   = secondsBooked / RET_1;
    const firstBooked   = firstShowed / RET_1;
    const sentToList    = Math.round(firstBooked / CALL_LIST_BOOKED);
    const secondsShowed = secondsBooked * RET_2;
    const newStartsBook = Math.round(secondsShowed * RET_2);
    const newStartsShow = Math.round(newStartsBook * RET_NS);

    // Row order matches RECRUITING_LABELS
    return [
      sentToList,                           // 0  Applies Received
      sentToList,                           // 1  Sent to List
      Math.round(firstBooked),              // 2  1st Rounds Booked
      Math.round(firstShowed),              // 3  1st Rounds Showed
      Math.round(RET_1 * 100),             // 4  Retention (%)
      Math.round(CALL_LIST_BOOKED * 100),  // 5  % Call List Booked (%)
      secondsBooked,                        // 6  2nd Rounds Booked
      Math.round(secondsShowed),            // 7  2nd Rounds Showed
      Math.round(RET_2 * 100),             // 8  Retention (%)
      newStartsBook,                        // 9  New Starts Booked
      newStartsShow,                        // 10 New Starts Showed
      Math.round(RET_NS * 100)             // 11 New Start Retention (%)
    ];
  },

  // ── Build recruiting rows from leader count + actuals arrays ──
  _buildRows(leaders, actuals) {
    const projected = this._calcProjected(leaders);
    return this.RECRUITING_LABELS.map((def, i) => {
      const vals = actuals[i] || [];
      const p = projected[i];
      let total;
      if (def.isRate) {
        const nums = vals.filter(v => typeof v === 'number');
        total = nums.length ? Math.round(nums.reduce((a, b) => a + b, 0) / nums.length) : 0;
      } else {
        total = vals.reduce((a, b) => a + (typeof b === 'number' ? b : 0), 0);
      }
      return { label: def.label, projected: p, values: vals, total, isRate: def.isRate };
    });
  },

  // ── [REMOVED] _loadScaffoldData — hardcoded demo data for 7 test owners removed ──
  // Data now comes exclusively from consolidated campaign tabs via readConsolidatedRecruiting().
  // If no data exists, the UI shows an empty state instead of fake numbers.

  // ══════════════════════════════════════════════════
  // CAMPAIGN LANDING PAGE
  // ══════════════════════════════════════════════════

  _showLandingPage() {
    // Build campaign cards only from campaigns that have actual data
    const campaigns = this._allCampaignsData || {};
    const configCampaigns = NATIONAL_CONFIG.campaigns || {};

    const container = document.getElementById('campaign-cards');
    if (!container) return;

    // Only include campaigns that have owners
    const allKeys = new Set([...Object.keys(campaigns), ...Object.keys(configCampaigns)]);
    const activeKeys = [...allKeys].filter(key => {
      const cd = campaigns[key];
      if (!cd) return false;
      return (cd.owners || []).length > 0;
    });
    // Debug: log all campaigns and their owner counts
    console.log('[NationalApp] _showLandingPage — campaigns:', [...allKeys].map(k => k + ':' + (campaigns[k]?.owners?.length || 0)).join(', '));

    // Sort: planning schedule (today first, then tomorrow, etc.) or alphabetical fallback
    const schedule = this._planningSchedule || null;
    let sorted;
    if (schedule && schedule.length > 0) {
      const todayIdx = (new Date().getDay() + 6) % 7; // Mon=0, Sun=6
      sorted = activeKeys.sort((a, b) => {
        const aEntry = schedule.find(p => p.campaignKey === a);
        const bEntry = schedule.find(p => p.campaignKey === b);
        if (!aEntry && !bEntry) {
          const la = (campaigns[a]?.label || configCampaigns[a]?.label || a).toLowerCase();
          const lb = (campaigns[b]?.label || configCampaigns[b]?.label || b).toLowerCase();
          return la.localeCompare(lb);
        }
        if (!aEntry) return 1;
        if (!bEntry) return -1;
        const aDist = (aEntry.day - todayIdx + 7) % 7;
        const bDist = (bEntry.day - todayIdx + 7) % 7;
        if (aDist !== bDist) return aDist - bDist;
        return aEntry.sortOrder - bEntry.sortOrder;
      });
    } else {
      sorted = activeKeys.sort((a, b) => {
        const la = (campaigns[a]?.label || configCampaigns[a]?.label || a).toLowerCase();
        const lb = (campaigns[b]?.label || configCampaigns[b]?.label || b).toLowerCase();
        return la.localeCompare(lb);
      });
    }

    container.innerHTML = sorted.map(key => {
      const label = campaigns[key]?.label || configCampaigns[key]?.label || key;
      const logo = this.CAMPAIGN_LOGOS[key];
      const ownerCount = campaigns[key]?.owners?.length || 0;

      let logoHtml;
      if (Array.isArray(logo)) {
        logoHtml = `<div class="campaign-card-logo-dual">${logo.map(l =>
          `<img src="${l}" alt="" class="campaign-card-logo-half">`
        ).join('<span class="logo-divider">×</span>')}</div>`;
      } else if (logo) {
        logoHtml = `<img src="${logo}" alt="${this._esc(label)}" class="campaign-card-logo">`;
      } else {
        logoHtml = `<div class="campaign-card-logo placeholder">📊</div>`;
      }

      // Extract variant tag if label has "Brand: Variant" format
      const colonIdx = label.indexOf(':');
      const variant = colonIdx > 0 ? label.substring(colonIdx + 1).trim() : '';
      const variantHtml = variant
        ? `<div class="campaign-card-variant">${this._esc(variant)}</div>`
        : '';

      // Day pill from planning schedule
      const dayNames = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
      const planEntry = (schedule || []).find(p => p.campaignKey === key);
      const dayPill = planEntry !== undefined
        ? `<span class="campaign-card-day-pill">${dayNames[planEntry.day] || ''}</span>`
        : '';

      return `
        <div class="campaign-card" onclick="NationalApp.selectCampaign('${key}')">
          ${dayPill}
          ${logoHtml}
          <div class="campaign-card-label">${this._esc(label)}</div>
          ${variantHtml}
          ${ownerCount ? `<div class="campaign-card-owners">${ownerCount} owner${ownerCount !== 1 ? 's' : ''}</div>` : ''}
        </div>`;
    }).join('');

    // Show empty state if no campaigns have data
    if (sorted.length === 0) {
      container.innerHTML = `<div style="grid-column:1/-1;text-align:center;padding:40px;color:#708090;">
        <div style="font-size:40px;margin-bottom:12px;">📊</div>
        <div style="font-weight:600;">No campaign data yet</div>
        <div style="font-size:13px;margin-top:6px;">Run a data refresh to pull from campaign spreadsheets</div>
      </div>`;
    }

    // Show landing, hide campaign detail views
    document.getElementById('campaign-landing').style.display = '';
    document.querySelector('.campaign-overview').style.display = 'none';
    document.querySelector('.owners-section').style.display = 'none';
    document.getElementById('owner-detail').style.display = 'none';

    // Hide refresh buttons for non-superadmins
    const isSA = this._isSuperadmin();
    const refreshAll = document.getElementById('btn-refresh-all');
    if (refreshAll) refreshAll.style.display = isSA ? '' : 'none';
    const refreshStatus = document.getElementById('landing-refresh-status');
    if (refreshStatus) refreshStatus.style.display = isSA ? '' : 'none';

    // Prefetch today's scheduled campaigns in background
    this._prefetchTodaysCampaigns();
  },

  /**
   * Prefetch full campaign data for today's scheduled campaigns.
   * Runs in background so clicking into a campaign is near-instant.
   */
  async _prefetchTodaysCampaigns() {
    const sched = this._planningSchedule || [];
    if (!sched.length) return;

    const todayIdx = (new Date().getDay() + 6) % 7;
    const todayCampaigns = sched
      .filter(p => p.day === todayIdx)
      .sort((a, b) => a.sortOrder - b.sortOrder);

    if (!todayCampaigns.length) return;

    for (const entry of todayCampaigns) {
      const key = entry.campaignKey;
      // Skip if already cached and fresh
      if (this._readCoachCampaignCache(key)) {
        console.log('[NationalApp] Prefetch skip (cached):', key);
        continue;
      }
      // Skip if already prefetching
      if (this._prefetching && this._prefetching[key]) continue;
      if (!this._prefetching) this._prefetching = {};
      this._prefetching[key] = true;

      console.log('[NationalApp] Prefetching campaign data:', key);
      try {
        // Save current state, load campaign, cache it, restore state
        const savedOwners = this.state.owners;
        const savedTotals = this.state.campaignTotals;
        const savedCampaign = this.state.campaign;

        this.state.campaign = key;
        await this.loadCampaignData(key);
        this._writeCoachCampaignCache(key);

        // Restore previous state
        this.state.owners = savedOwners;
        this.state.campaignTotals = savedTotals;
        this.state.campaign = savedCampaign;

        console.log('[NationalApp] Prefetched + cached:', key);
      } catch (err) {
        console.warn('[NationalApp] Prefetch failed for', key, ':', err.message);
      }
      this._prefetching[key] = false;
    }
  },

  _COACH_CACHE_MAX_AGE: 15 * 60 * 1000, // 15 min per-campaign cache

  async selectCampaign(campaignKey) {
    this.state.campaign = campaignKey;
    this.state.selectedOwner = null;

    // Update the dropdown to match
    const select = document.getElementById('campaign-select');
    if (select) select.value = campaignKey;

    // Hide landing, show campaign view
    document.getElementById('campaign-landing').style.display = 'none';
    document.querySelector('.campaign-overview').style.display = '';
    document.querySelector('.owners-section').style.display = '';

    // Check per-campaign localStorage cache first
    const cached = this._readCoachCampaignCache(campaignKey);
    if (cached) {
      console.log('[NationalApp] Rendering campaign from cache:', campaignKey);
      this.state.owners = cached.owners;
      this.state.campaignTotals = cached.campaignTotals || {};

      // Ranking: att-res uses Tableau data (async fetch), all others use production (instant)
      const isRes = campaignKey === 'att-res';
      if (isRes) {
        this._showLoading('Loading rankings...');
        try {
          const rankData = await this._fetchWithTimeout(this._fetchD2DResRanking(), 10000);
          if (rankData?.ranking?.length) {
            this._enrichOwnersWithD2DRanking(rankData.ranking);
          }
        } catch (err) {
          console.warn('[NationalApp] D2D ranking fetch failed:', err.message);
        }
        this._hideLoading();
      } else {
        this._rankOwnersByProduction();
      }
      this.renderCampaignOverview();
      this.renderOwnersList();

      // Fetch camMapping + audit in background (not cached)
      Promise.allSettled([
        this._fetchOwnerCamMapping(),
        this._fetchOnlinePresence()
      ]).then(([camRes, auditRes]) => {
        if (camRes.status === 'fulfilled' && camRes.value) {
          this.state.camMapping = camRes.value.mapping || camRes.value;
        }
        if (auditRes.status === 'fulfilled' && auditRes.value?.businesses?.length) {
          this.state.allCompanyNames = auditRes.value.allCompanyNames || [];
          this._cachedAuditBusinesses = auditRes.value.businesses;
          this._mapAuditToOwners(auditRes.value.businesses, this.state.camMapping || null);
        }
      });
      return;
    }

    // No cache — full fetch
    this._showLoading('Loading campaign data...');
    try {
      await this.loadCampaignData(campaignKey);
      // Cache post-enrichment state for next time
      this._writeCoachCampaignCache(campaignKey);
    } catch (err) {
      console.error('Failed to load campaign:', err);
    }
    this._hideLoading();
    console.log('[NationalApp] selectCampaign render:', campaignKey, 'owners:', this.state.owners.length);
    try { this.renderCampaignOverview(); } catch (e) { console.error('[NationalApp] renderCampaignOverview error:', e); }
    try { this.renderOwnersList(); } catch (e) { console.error('[NationalApp] renderOwnersList error:', e); }
  },

  // ── Back to landing page ──
  backToLanding() {
    console.log('[NationalApp] backToLanding — _allCampaignsData keys:', Object.keys(this._allCampaignsData || {}));
    this.state.campaign = null;
    this.state.selectedOwner = null;
    this._showLandingPage();
  },

  // ── Legacy dropdown switcher (still works from top bar) ──
  async switchCampaign(campaignKey) {
    await this.selectCampaign(campaignKey);
  },

  // ══════════════════════════════════════════════════
  // COACH CAMPAIGN CACHE (per-campaign localStorage)
  // ══════════════════════════════════════════════════

  _readCoachCampaignCache(campaignKey) {
    try {
      const raw = localStorage.getItem('coach_cache_' + campaignKey);
      if (!raw) return null;
      const data = JSON.parse(raw);
      if (Date.now() - (data._ts || 0) > this._COACH_CACHE_MAX_AGE) {
        localStorage.removeItem('coach_cache_' + campaignKey);
        return null;
      }
      if (!data.owners || !data.owners.length) return null;
      // Reject stale caches where owners have no real health data
      const hasRealData = data.owners.some(o =>
        (o.headcountHistory && o.headcountHistory.length > 0) ||
        (o.headcount && o.headcount.active > 0)
      );
      if (!hasRealData) {
        console.warn('[NationalApp] Rejecting stale coach cache (no health data):', campaignKey);
        localStorage.removeItem('coach_cache_' + campaignKey);
        return null;
      }
      return data;
    } catch { return null; }
  },

  _writeCoachCampaignCache(campaignKey) {
    try {
      // Only cache if owners have real data (not empty-weeks placeholders)
      const hasRealData = this.state.owners.some(o =>
        (o.headcountHistory && o.headcountHistory.length > 0) ||
        (o.headcount && o.headcount.active > 0)
      );
      if (!hasRealData) {
        console.warn('[NationalApp] Skipping coach cache write (no health data):', campaignKey);
        return;
      }
      localStorage.setItem('coach_cache_' + campaignKey, JSON.stringify({
        _ts: Date.now(),
        owners: this.state.owners,
        campaignTotals: this.state.campaignTotals
      }));
    } catch (err) {
      console.warn('[NationalApp] Coach cache write failed:', err.message);
    }
  },

  _clearCoachCampaignCache(campaignKey) {
    try { localStorage.removeItem('coach_cache_' + campaignKey); } catch (e) {}
  },

  _clearAllCoachCaches() {
    const toRemove = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k && k.startsWith('coach_cache_')) toRemove.push(k);
    }
    toRemove.forEach(k => localStorage.removeItem(k));
  },

  // ══════════════════════════════════════════════════
  // REFRESH CAMPAIGN DATA
  // Pulls latest data from per-campaign source spreadsheets
  // into consolidated tabs, then reloads the view.
  // ══════════════════════════════════════════════════

  async importLatestRecruiting() {
    if (!NATIONAL_CONFIG.appsScriptUrl) {
      alert('Apps Script URL not configured. Deploy NationalCode.gs first.');
      return;
    }

    const campaignKey = this.state.campaign;
    if (!campaignKey) { alert('No campaign selected.'); return; }

    const btn = document.getElementById('btn-import-recruiting');
    const status = document.getElementById('import-status');

    if (btn) { btn.disabled = true; btn.textContent = 'Refreshing...'; }
    if (status) { status.textContent = ''; status.className = 'import-status'; }

    try {
      const result = await this._refreshCampaign(campaignKey, 60000);

      // Clear cache for this campaign only, then reload
      this._clearCoachCampaignCache(campaignKey);
      if (this._allCampaignsData) delete this._allCampaignsData[campaignKey];
      await this.loadCampaignData(campaignKey);

      // Re-render dashboard
      this.renderCampaignOverview();
      this.renderOwnersList();

      const msg = `Refreshed ${campaignKey} (${result.rows || 0} rows)`;
      if (status) {
        status.textContent = msg;
        status.className = 'import-status import-success';
        setTimeout(() => { status.textContent = ''; status.className = 'import-status'; }, 6000);
      }
      console.log('[NationalApp] Single campaign refresh:', result);

    } catch (err) {
      console.error('[NationalApp] Refresh failed:', err);
      if (status) {
        status.textContent = 'Refresh failed: ' + err.message;
        status.className = 'import-status import-error';
      }
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = 'Refresh Data'; }
    }
  },

  /**
   * Refresh a single campaign via the backend.
   * @param {string} campaignKey - e.g. 'lumen', 'frontier'
   * @param {number} [timeout=60000] - Fetch timeout in ms
   * @returns {Promise<Object>} Backend response
   */
  async _refreshCampaign(campaignKey, timeout) {
    const resp = await this._fetchWithTimeout(
      fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'refreshCampaign',
          campaign: campaignKey
        })
      }),
      timeout || 60000
    );
    const result = await resp.json();
    if (result.error) throw new Error(result.error);
    if (!result.ok) throw new Error(result.error || 'Refresh failed');
    return result;
  },

  /**
   * Refresh all campaign data from the landing page.
   * Same as importLatestRecruiting but targets landing page UI elements
   * and reloads the landing page when done.
   */
  /**
   * Refresh a single campaign from the landing page (per-card refresh button).
   */
  async refreshSingleFromLanding(campaignKey, btnEl) {
    if (!NATIONAL_CONFIG.appsScriptUrl) return;

    const originalText = btnEl.textContent;
    btnEl.disabled = true;
    btnEl.textContent = '...';
    btnEl.classList.add('spinning');

    try {
      await this._refreshCampaign(campaignKey, 60000);
      this._clearCoachCampaignCache(campaignKey);
      if (this._allCampaignsData) delete this._allCampaignsData[campaignKey];

      // Reload all campaign data and re-render landing
      await this.loadCampaignData(campaignKey);
      this._showLandingPage();

      console.log('[NationalApp] Single refresh done:', campaignKey);
    } catch (err) {
      console.error('[NationalApp] Single refresh failed:', campaignKey, err);
      alert('Refresh failed for ' + campaignKey + ': ' + err.message);
    } finally {
      // Button may be gone after re-render, but safety first
      if (btnEl) { btnEl.disabled = false; btnEl.textContent = originalText; btnEl.classList.remove('spinning'); }
    }
  },

  /**
   * Refresh all campaigns sequentially from the landing page.
   * Each campaign is refreshed individually so a failure in one
   * doesn't block the others, and total execution time is bounded.
   */
  async refreshAllFromLanding() {
    if (!NATIONAL_CONFIG.appsScriptUrl) {
      alert('Apps Script URL not configured. Deploy NationalCode.gs first.');
      return;
    }

    const btn = document.getElementById('btn-refresh-all');
    const status = document.getElementById('landing-refresh-status');

    // Gather all campaign keys that have sheetIds
    const allKeys = Object.keys(this._allCampaignsData || {});
    const configKeys = Object.keys(NATIONAL_CONFIG.campaigns || {});
    const keys = [...new Set([...allKeys, ...configKeys])];

    if (btn) { btn.disabled = true; btn.textContent = 'Refreshing...'; }
    if (status) { status.textContent = `Refreshing 0/${keys.length} campaigns...`; status.className = 'import-status'; }

    let successCount = 0;
    let totalRows = 0;
    const errors = [];

    for (let i = 0; i < keys.length; i++) {
      const key = keys[i];
      if (status) status.textContent = `Refreshing ${i + 1}/${keys.length}: ${key}...`;

      try {
        const result = await this._refreshCampaign(key, 60000);
        successCount++;
        totalRows += result.rows || 0;
      } catch (err) {
        console.warn('[NationalApp] Refresh failed for', key, err.message);
        errors.push(key);
      }
    }

    // Clear all caches and reload
    this._allCampaignsData = null;
    this._clearAllCoachCaches();
    this._coachInitDone = false;

    try {
      await this.loadCampaignData('att-b2b');
    } catch (e) { /* landing page will show whatever loaded */ }

    this._showLandingPage();

    if (status) {
      const errMsg = errors.length ? ` (${errors.length} failed: ${errors.join(', ')})` : '';
      status.textContent = `Refreshed ${successCount}/${keys.length} campaigns (${totalRows} rows)${errMsg}`;
      status.className = errors.length ? 'import-status import-error' : 'import-status import-success';
      setTimeout(() => { status.textContent = ''; status.className = 'import-status'; }, 8000);
    }

    console.log('[NationalApp] Full refresh from landing:', { successCount, totalRows, errors });

    if (btn) { btn.disabled = false; btn.textContent = 'Refresh All Data'; }
  },

  // (Bulk importCosts removed — costs now lazy-load per-owner when Recruiting tab opens)

  // ══════════════════════════════════════════════════
  // RENDER: Campaign Overview (KPIs + Recruiting Table + Status Codes)
  // ══════════════════════════════════════════════════

  renderCampaignOverview() {
    const t = this.state.campaignTotals || {};
    const cfg = NATIONAL_CONFIG.campaigns[this.state.campaign];
    document.getElementById('campaign-title').textContent = (cfg?.label || 'Campaign') + ' Campaign';
    const weekLabel = t.latestWeek ? this._formatWeekDate(t.latestWeek) : this._formatCurrentWeek();
    document.getElementById('overview-date').textContent = 'Week of ' + weekLabel;

    const hc = t.hcBreakdown || {};
    const pb = t.prodBreakdown || {};

    // Build production breakdown items (skip if only one product, skip zero values, sort desc)
    const prodEntries = Object.entries(pb).filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]);
    const prodItems = prodEntries.length > 1
      ? prodEntries.map(([name, val]) => `<div class="kpi-breakdown-item"><span class="kpi-bd-label">${this._esc(name)}</span><span class="kpi-bd-value">${val.toLocaleString()}</span></div>`).join('')
      : '';

    // Build headcount breakdown items
    const hcItems = [
      hc.leaders ? ['Leaders', hc.leaders] : null,
      hc.dist ? ['Distributors', hc.dist] : null,
      hc.closers ? ['Closers', hc.closers] : null,
      hc.leadGen ? ['Lead Gen', hc.leadGen] : null,
      hc.training ? ['Training', hc.training] : null
    ].filter(Boolean)
      .map(([label, val]) => `<div class="kpi-breakdown-item"><span class="kpi-bd-label">${label}</span><span class="kpi-bd-value">${val.toLocaleString()}</span></div>`)
      .join('');

    const container = document.getElementById('campaign-kpis');
    container.innerHTML = `
      <div class="kpi-card">
        <div class="kpi-label">Active Headcount</div>
        <div class="kpi-value">${t.headcount?.toLocaleString() || '—'}</div>
        <div class="kpi-breakdown">${hcItems || '<div class="kpi-bd-empty">No breakdown</div>'}</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">1st Rounds</div>
        <div class="kpi-value">${(t.firstShowed || 0).toLocaleString()}<span class="kpi-sub"> showed</span></div>
        <div class="kpi-context">out of ${(t.firstBooked || 0).toLocaleString()} booked</div>
        <div class="kpi-retention">${t.firstRetention || 0}% retention</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">2nd Rounds</div>
        <div class="kpi-value">${(t.secondShowed || 0).toLocaleString()}<span class="kpi-sub"> showed</span></div>
        <div class="kpi-context">out of ${(t.secondBooked || 0).toLocaleString()} booked</div>
        <div class="kpi-retention">${t.secondRetention || 0}% retention</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">New Starts</div>
        <div class="kpi-value">${(t.newStartsShowed || 0).toLocaleString()}<span class="kpi-sub"> showed</span></div>
        <div class="kpi-context">out of ${(t.newStartsBooked || 0).toLocaleString()} booked</div>
        <div class="kpi-retention">${t.newStartRetention || 0}% retention</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Total Production</div>
        <div class="kpi-value">${t.production?.toLocaleString() || '—'}</div>
        <div class="kpi-breakdown">${prodItems || '<div class="kpi-bd-empty">No breakdown</div>'}</div>
      </div>`;

    // Hide campaign-level refresh for non-superadmins
    const isSA = this._isSuperadmin();
    const refreshBtn = document.getElementById('btn-import-recruiting');
    if (refreshBtn) refreshBtn.style.display = isSA ? '' : 'none';
    const importWeeks = document.getElementById('import-weeks');
    if (importWeeks) importWeeks.style.display = isSA ? '' : 'none';
    const importStatus = document.getElementById('import-status');
    if (importStatus) importStatus.style.display = isSA ? '' : 'none';
  },

  // ══════════════════════════════════════════════════
  // RENDER: Owners List (directory-style cards)
  // ══════════════════════════════════════════════════

  renderOwnersList() {
    const container = document.getElementById('owners-list');
    const owners = this.state.owners;

    // Apply planning owner order if available for current campaign + today
    const schedule = this._planningSchedule || [];
    const todayIdx = (new Date().getDay() + 6) % 7;
    const planEntry = schedule.find(
      p => p.campaignKey === this.state.campaign && p.day === todayIdx
    );
    if (planEntry && planEntry.ownerOrder && planEntry.ownerOrder.length > 0) {
      const orderMap = {};
      planEntry.ownerOrder.forEach((name, i) => { orderMap[name.toLowerCase()] = i; });
      owners.sort((a, b) => {
        const ai = orderMap[a.name.toLowerCase()] ?? 9999;
        const bi = orderMap[b.name.toLowerCase()] ?? 9999;
        return ai - bi;
      });
    }

    if (!owners.length) {
      container.innerHTML = `
        <div class="empty-state" style="grid-column:1/-1">
          <div class="empty-state-icon">📊</div>
          <div class="empty-state-text">No owner data available yet.<br>Click <strong>Refresh Data</strong> to pull from campaign spreadsheets.</div>
        </div>`;
      return;
    }

    container.innerHTML = owners.map((o, idx) => {
      const rankCls = o.d2dRank === 1 ? ' rank-gold' : o.d2dRank === 2 ? ' rank-silver' : o.d2dRank === 3 ? ' rank-bronze' : '';
      const cardCls = o.d2dRank <= 3 ? rankCls : '';
      const rankBadge = o.d2dRank
        ? `<span class="owner-rank-badge${rankCls}" title="#${o.d2dRank} — ${o.d2dTotalUnits || 0} units LW">#${o.d2dRank}</span>`
        : '';
      return `
        <div class="owner-card${cardCls}" onclick="NationalApp.openOwnerDetail(${idx})">
          ${rankBadge}
          <span class="owner-card-name">${this._esc(o.name)}</span>
          <span class="owner-card-arrow">→</span>
        </div>`;
    }).join('');
  },

  filterOwners(query) {
    const q = query.toLowerCase();
    const cards = document.querySelectorAll('.owner-card');
    cards.forEach((card, idx) => {
      const owner = this.state.owners[idx];
      if (!owner) return;
      card.style.display = owner.name.toLowerCase().includes(q) ? '' : 'none';
    });
  },

  // ══════════════════════════════════════════════════
  // RENDER: Owner Detail
  // ══════════════════════════════════════════════════

  openOwnerDetail(idx) {
    const owner = this.state.owners[idx];
    if (!owner) return;
    this.state.selectedOwner = owner;
    this.state.currentTab = 'health';

    document.querySelector('.campaign-overview').style.display = 'none';
    document.querySelector('.owners-section').style.display = 'none';
    const detail = document.getElementById('owner-detail');
    detail.style.display = 'block';

    document.getElementById('detail-owner-name').textContent = owner.name;
    const badge = document.getElementById('detail-owner-badge');
    badge.textContent = NATIONAL_CONFIG.campaigns[this.state.campaign]?.label || '';
    badge.className = 'detail-badge';

    document.querySelectorAll('.detail-tab').forEach(t => t.classList.remove('active'));
    document.querySelector('.detail-tab[data-tab="health"]').classList.add('active');

    // Hide Sales tab for specific campaigns
    const hideSales = ['frontier', 'verizon-fios', 'leafguard', 'lumen'];
    const salesTab = document.querySelector('.detail-tab[data-tab="sales"]');
    if (salesTab) salesTab.style.display = hideSales.includes(this.state.campaign) ? 'none' : '';

    this.renderHealthTab(owner);
    this._showTab('health');

    // Update NLR / Cam "NEW" badges on detail tabs
    this._updateDetailTabBadges(owner.name, this.state.campaign);

    // Lazy-load NLR data for this specific owner (non-blocking)
    if (!owner._nlrFetched) {
      this._fetchOwnerNlrData(owner);
    }

    // Lazy-load sales data for this owner (non-blocking)
    const isNDS = this.state.campaign && this.state.campaign.indexOf('nds') >= 0;
    const isRes = this.state.campaign === 'att-res';
    if (isNDS && !owner._ndsSalesFetched) {
      this._fetchNDSOwnerSales(owner);
    }
    if (isRes && !owner._resSalesFetched) {
      this._fetchResOwnerSales(owner);
    }
  },

  closeOwnerDetail() {
    this.state.selectedOwner = null;
    document.getElementById('owner-detail').style.display = 'none';
    document.querySelector('.campaign-overview').style.display = '';
    document.querySelector('.owners-section').style.display = '';
  },

  switchDetailTab(tab) {
    this.state.currentTab = tab;
    document.querySelectorAll('.detail-tab').forEach(t => t.classList.remove('active'));
    document.querySelector(`.detail-tab[data-tab="${tab}"]`).classList.add('active');
    this._showTab(tab);

    const owner = this.state.selectedOwner;
    if (!owner) return;

    // Clear "NEW" badge when coach views the tab
    if (tab === 'recruiting' || tab === 'audit') {
      this._markTabViewed(owner.name, this.state.campaign, tab);
      const badge = document.getElementById('tab-badge-' + tab);
      if (badge) badge.style.display = 'none';
    }

    switch (tab) {
      case 'health': this.renderHealthTab(owner); break;
      case 'recruiting': this.renderRecruitingTab(owner); break;
      case 'sales': this.renderSalesTab(owner); break;
      case 'audit': this.renderAuditTab(owner); break;
    }
  },

  _showTab(tab) {
    ['health', 'recruiting', 'sales', 'audit'].forEach(t => {
      const el = document.getElementById('tab-' + t);
      if (el) el.style.display = t === tab ? 'block' : 'none';
    });
  },

  // ══════════════════════════════════════════════════
  // RENDER: Health Tab (1-on-1 Coaching Flow)
  // ══════════════════════════════════════════════════

  renderHealthTab(owner) {
    const hc = owner.headcount;
    const prod = owner.production;
    const goals = owner.nextGoals;
    const ownerIdx = this.state.owners.indexOf(owner);

    // ── Section 1: Headcount Check (pre-fill with current values so coaches can update all week) ──
    const headcountEl = document.getElementById('health-headcount');
    const isLG = this.state.campaign === 'leafguard';
    const _hcVal = (field) => hc[field] ? hc[field] : '';
    const _hcDist = isLG
      ? (hc.active && hc.closers ? Math.max(hc.active - hc.closers, 0) : '—')
      : (hc.active && hc.leaders ? Math.max(hc.active - hc.leaders, 0) : '—');
    const hcFieldsHtml = isLG ? `
        <div class="hc-field">
          <label class="hc-field-label">Active Reps</label>
          <input type="number" class="hc-input" id="hc-active-${ownerIdx}" value="${_hcVal('active')}" min="0" placeholder="—"
            onchange="NationalApp._updateHeadcount(${ownerIdx}, 'active', this.value)">
        </div>
        <div class="hc-field">
          <label class="hc-field-label">Closers</label>
          <input type="number" class="hc-input" id="hc-closers-${ownerIdx}" value="${_hcVal('closers')}" min="0" placeholder="—"
            onchange="NationalApp._updateHeadcount(${ownerIdx}, 'closers', this.value)">
        </div>
        <div class="hc-field hc-field-calc">
          <label class="hc-field-label">Lead Gen</label>
          <div class="hc-value" id="hc-dist-${ownerIdx}">${_hcDist}</div>
          <div class="hc-calc-note">Active − Closers</div>
        </div>
        <div class="hc-field">
          <label class="hc-field-label">Leaders</label>
          <input type="number" class="hc-input" id="hc-leaders-${ownerIdx}" value="${_hcVal('leaders')}" min="0" placeholder="—"
            onchange="NationalApp._updateHeadcount(${ownerIdx}, 'leaders', this.value)">
        </div>
        <div class="hc-field">
          <label class="hc-field-label">In Training</label>
          <input type="number" class="hc-input" id="hc-training-${ownerIdx}" value="${_hcVal('training')}" min="0" placeholder="—"
            onchange="NationalApp._updateHeadcount(${ownerIdx}, 'training', this.value)">
        </div>` : `
        <div class="hc-field">
          <label class="hc-field-label">Active Reps</label>
          <input type="number" class="hc-input" id="hc-active-${ownerIdx}" value="${_hcVal('active')}" min="0" placeholder="—"
            onchange="NationalApp._updateHeadcount(${ownerIdx}, 'active', this.value)">
        </div>
        <div class="hc-field">
          <label class="hc-field-label">Leaders</label>
          <input type="number" class="hc-input" id="hc-leaders-${ownerIdx}" value="${_hcVal('leaders')}" min="0" placeholder="—"
            onchange="NationalApp._updateHeadcount(${ownerIdx}, 'leaders', this.value)">
        </div>
        <div class="hc-field hc-field-calc">
          <label class="hc-field-label">Distributors</label>
          <div class="hc-value" id="hc-dist-${ownerIdx}">${_hcDist}</div>
          <div class="hc-calc-note">Active − Leaders</div>
        </div>
        <div class="hc-field">
          <label class="hc-field-label">In Training</label>
          <input type="number" class="hc-input" id="hc-training-${ownerIdx}" value="${_hcVal('training')}" min="0" placeholder="—"
            onchange="NationalApp._updateHeadcount(${ownerIdx}, 'training', this.value)">
        </div>`;
    headcountEl.innerHTML = `
      <div class="coaching-label">Headcount</div>
      <div class="hc-grid">${hcFieldsHtml}</div>
      <div class="hc-submit-row">
        <button class="hc-submit-btn" id="hc-submit-${ownerIdx}"
          onclick="NationalApp._submitHeadcount(${ownerIdx})">Update Headcount</button>
        <span class="hc-submit-note" id="hc-submit-note-${ownerIdx}"></span>
      </div>`;

    // ── Headcount Trend Table ──
    this._renderHeadcountTrend(owner, ownerIdx);

    // ── Section 2: Production Review (per-product cards) ──
    const prodEl = document.getElementById('health-production');
    let prodCardsHtml = '';
    const productEntries = prod.products || {};
    const productNames = Object.keys(productEntries);
    if (this.state.campaign === 'leafguard' && productNames.length >= 4) {
      // LeafGuard: 2 combined cards
      // Card 1: Gross Sales (main) + Personal Prod (sub)
      const gs = productEntries['Gross Sales'] || { actual: 0, goal: 0 };
      const pp = productEntries['Personal Prod'] || { actual: 0, goal: 0 };
      // Card 2: Gross Leads (main) + Number of Sales (sub)
      const gl = productEntries['Gross Leads'] || { actual: 0, goal: 0 };
      const ns = productEntries['Number of Sales'] || { actual: 0, goal: 0 };
      prodCardsHtml = this._prodCardCombined('Gross Sales', gs.actual, gs.goal, 'Personal Prod', pp.actual, true)
                    + this._prodCardCombined('Gross Leads', gl.actual, gl.goal, 'Number of Sales', ns.actual, false);
    } else if (productNames.length > 0) {
      // Show per-product cards only — hide products with no goal, no current units, and no history
      const prodHist = owner.productionHistory || [];
      // Determine if production is missing (all zeros) → make cards editable
      const prodMissing = !prod.totalActual && !prod.totalGoal;
      for (const pName of productNames) {
        const pData = productEntries[pName];
        if (!pData.actual && !pData.goal) {
          // Check if this product has ever had production historically
          const everSold = prodHist.some(w => w.products && w.products[pName] && (w.products[pName].actual > 0));
          if (!everSold && !prodMissing) continue;
        }
        prodCardsHtml += prodMissing
          ? this._prodCardEditable(pName, pData.actual, pData.goal, ownerIdx)
          : this._prodCard(pName, pData.actual, pData.goal);
      }
    } else {
      // Fallback: single total card
      const prodMissing = !prod.totalActual && !prod.totalGoal;
      prodCardsHtml = prodMissing
        ? this._prodCardEditable('Total Units', prod.totalActual, prod.totalGoal, ownerIdx)
        : this._prodCard('Total Units', prod.totalActual, prod.totalGoal);
    }
    const prodMissing = !prod.totalActual && !prod.totalGoal;
    prodEl.innerHTML = `
      <div class="coaching-label">Production Review <span class="coaching-sublabel">Last Week</span>${prodMissing ? ' <span style="color:var(--orange);font-size:12px;font-weight:500;">— No data entered</span>' : ''}</div>
      <div class="prod-cards">${prodCardsHtml}</div>
      ${prodMissing ? `<div class="hc-submit-row">
        <button class="hc-submit-btn" onclick="NationalApp._submitProductionCards(${ownerIdx})">Save Production</button>
        <span class="hc-submit-note" id="prod-submit-note-${ownerIdx}"></span>
      </div>` : ''}`;

    // ── Production Trend Table ──
    this._renderProductionTrend(owner, ownerIdx);

    // ── Section 3: Set Goals (Next Week) ──
    const goalsEl = document.getElementById('health-goals');
    let goalFieldsHtml = '';
    if (productNames.length > 0) {
      // LeafGuard: only show Gross Sales + Gross Leads goals (skip Personal Prod + Number of Sales)
      // All other campaigns: show ALL products so coaches can set goals for new product lines
      // (Production Review cards still hide unsold products to avoid permanent red 0s)
      const goalProducts = this.state.campaign === 'leafguard'
        ? productNames.filter(p => p === 'Gross Sales' || p === 'Gross Leads')
        : [...productNames];
      for (const pName of goalProducts) {
        goalFieldsHtml += `
          <div class="goal-field">
            <label class="goal-field-label">${this._esc(pName)} Goal</label>
            <input type="number" class="goal-input" id="goal-${pName.toLowerCase().replace(/\s+/g,'-')}-${ownerIdx}"
              value="" min="0" placeholder="—"
              onchange="NationalApp._updateGoal(${ownerIdx}, '${pName}', this.value)">
          </div>`;
      }
    } else {
      goalFieldsHtml = `
        <div class="goal-field">
          <label class="goal-field-label">Total Units</label>
          <input type="number" class="goal-input" id="goal-total-${ownerIdx}" value="" min="0"
            placeholder="—"
            onchange="NationalApp._updateGoal(${ownerIdx}, 'totalUnits', this.value)">
        </div>`;
    }
    goalsEl.innerHTML = `
      <div class="coaching-label">Set Goals <span class="coaching-sublabel">Next Week</span></div>
      <div class="goals-grid">${goalFieldsHtml}</div>
      <div class="hc-submit-row">
        <button class="hc-submit-btn" onclick="NationalApp._submitGoals(${ownerIdx})">Submit Goals</button>
        <span class="hc-submit-note" id="goal-submit-note-${ownerIdx}"></span>
      </div>`;

    // ── Section 4: Notes ──
    const notesEl = document.getElementById('health-notes');
    if (notesEl) {
      const ownerNotes = (this.state.campaignNotes || [])
        .filter(n => n.ownerName === owner.name)
        .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
      const notesHtml = ownerNotes.length
        ? ownerNotes.map(n => `
          <div class="note-entry">
            <div class="note-entry-header">
              <span class="note-author">${this._esc(n.coachName || 'Unknown')}</span>
              <span class="note-time">${this._relativeTime(n.timestamp)}</span>
              <button class="note-delete-btn" onclick="NationalApp._deleteNote('${n.noteId}', ${ownerIdx})" title="Delete note">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M3 6h18"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/></svg>
              </button>
            </div>
            <div class="note-text">${this._esc(n.text)}</div>
          </div>`).join('')
        : '<div class="notes-empty">No notes yet</div>';

      notesEl.innerHTML = `
        <div class="coaching-label">Notes</div>
        <div class="note-input-row">
          <textarea class="note-input" id="note-input-${ownerIdx}" placeholder="Add a note..." rows="2"></textarea>
          <button class="note-add-btn" onclick="NationalApp._addNote(${ownerIdx})">Add</button>
        </div>
        <div class="notes-log" id="notes-log-${ownerIdx}">${notesHtml}</div>`;
    }
  },

  // ── Headcount input handler ──
  _updateHeadcount(ownerIdx, field, value) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    owner.headcount[field] = parseInt(value) || 0;
    // Recalc distributors display
    const activeEl = document.getElementById('hc-active-' + ownerIdx);
    const leadersEl = document.getElementById('hc-leaders-' + ownerIdx);
    const activeVal = parseInt(activeEl?.value) || 0;
    const leadersVal = parseInt(leadersEl?.value) || 0;
    const dist = activeVal - leadersVal;
    const distEl = document.getElementById('hc-dist-' + ownerIdx);
    if (distEl) distEl.textContent = (activeEl?.value && leadersEl?.value) ? dist : '—';
  },

  // ── Submit headcount — updates the most recent row in the consolidated sheet ──
  async _submitHeadcount(ownerIdx) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;

    // Read current values from inputs
    const active = parseInt(document.getElementById('hc-active-' + ownerIdx)?.value) || 0;
    const leaders = parseInt(document.getElementById('hc-leaders-' + ownerIdx)?.value) || 0;
    const training = parseInt(document.getElementById('hc-training-' + ownerIdx)?.value) || 0;
    const closers = parseInt(document.getElementById('hc-closers-' + ownerIdx)?.value) || 0;

    if (!active && !leaders && !training) return; // Don't submit empty

    // Update local state
    owner.headcount.active = active;
    owner.headcount.leaders = leaders;
    owner.headcount.training = training;
    if (closers) owner.headcount.closers = closers;

    // Update headcountHistory: find or create the current (newest) week entry
    const hist = owner.headcountHistory || [];
    // The newest week date — use _latestWeekDate or derive from current week
    const newestWeekDate = this._latestWeekDate || (() => {
      const now = new Date();
      const day = now.getDay(); // 0=Sun
      const sun = new Date(now.getFullYear(), now.getMonth(), now.getDate() - day);
      return (sun.getMonth() + 1) + '/' + sun.getDate() + '/' + sun.getFullYear();
    })();
    // Check if the newest week already has an entry in history
    const existingIdx = hist.findIndex(h => h.date === newestWeekDate);
    const hcEntry = { date: newestWeekDate, active, leaders, training, closers: closers || 0, leadGen: 0 };
    if (existingIdx >= 0) {
      // Update existing entry for this week
      hist[existingIdx] = hcEntry;
    } else {
      // New week — push to end (history is oldest-first)
      hist.push(hcEntry);
    }
    owner.headcountHistory = hist;

    // Re-render the trend chart
    this._renderHeadcountTrend(owner, ownerIdx);

    // Recalculate recruiting projected values based on new leader count
    if (owner.recruiting && owner.recruiting.rows.length) {
      owner.recruiting.leaders = leaders;
      const actuals = owner.recruiting.rows.map(r => r.values);
      owner.recruiting.rows = this._buildRows(leaders, actuals);
      if (this.state.currentTab === 'recruiting') {
        this.renderRecruitingTab(owner);
      }
    }

    // ── Write to spreadsheet via Apps Script ──
    const campaignLabel = NATIONAL_CONFIG.campaigns[this.state.campaign]?.label || '';
    const dist = Math.max(active - leaders, 0);
    const note = document.getElementById('hc-submit-note-' + ownerIdx);
    try {
      if (note) { note.textContent = 'Saving...'; note.classList.add('show'); }
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'updateHeadcount',
          ownerName: owner.name,
          date: newestWeekDate,
          active, leaders, dist, training,
          campaignLabel
        })
      });
      const result = await resp.json();
      if (result.error) throw new Error(result.error);
      console.log('[NationalApp] Headcount saved:', result);
      this._invalidateOdCache();
      if (note) {
        note.textContent = `Saved ✓ (row ${result.row})`;
        setTimeout(() => note.classList.remove('show'), 3000);
      }
    } catch (err) {
      console.error('[NationalApp] Headcount save failed:', err);
      if (note) {
        note.textContent = 'Save failed — ' + err.message;
        note.style.color = '#e53535';
        setTimeout(() => { note.classList.remove('show'); note.style.color = ''; }, 5000);
      }
    }
  },

  // ── Render headcount week-over-week bar chart (newest-first, dynamic Y-axis) ──
  _renderHeadcountTrend(owner, ownerIdx) {
    const trendEl = document.getElementById('health-hc-trend');
    if (!trendEl) return;

    const rawHist = owner.headcountHistory || [];
    if (!rawHist.length) { trendEl.style.display = 'none'; return; }
    trendEl.style.display = '';

    const isLeafGuard = this.state.campaign === 'leafguard';

    // Reverse so newest week is on the LEFT, skip all-zero weeks
    const hist = [...rawHist].filter(r => (r.active || 0) > 0 || (r.leaders || 0) > 0 || (r.training || 0) > 0).reverse();
    const n = hist.length;
    if (!n) { trendEl.style.display = 'none'; return; }

    const shortDate = (d) => {
      if (!d) return '';
      const s = String(d);
      // If it's a full Date toString(), parse it
      if (s.length > 10 && s.indexOf('GMT') >= 0) {
        const dt = new Date(s);
        if (!isNaN(dt)) return (dt.getMonth()+1) + '/' + dt.getDate();
      }
      const parts = s.split('/');
      return parts.length >= 2 ? parts[0] + '/' + parts[1] : s;
    };

    // ── Layout constants ──
    const VISIBLE = 5;
    const YAXIS_W = 36;
    const PAD_R = 10, PAD_T = 14, PAD_B = 30;
    const BAR_R = 5;
    const GAP = 0.14;
    const MIN_LABEL_H = 14;
    const svgH = 222;
    const plotH = svgH - PAD_T - PAD_B;
    const REF_W = 500;
    const barAreaVisibleW = REF_W - YAXIS_W;
    const slotW = barAreaVisibleW / VISIBLE;
    const barAreaW = n * slotW + PAD_R;
    const needsScroll = n > VISIBLE;
    const barW = slotW * (1 - GAP);
    const barOff = (slotW - barW) / 2;
    const baseY = PAD_T + plotH;

    // Store chart state for scroll-driven re-renders
    this._hcData = { hist, ownerIdx, n, slotW, barW, barOff, barAreaW, barAreaVisibleW, plotH, PAD_T, PAD_R, BAR_R, GAP, MIN_LABEL_H, svgH, baseY, YAXIS_W, VISIBLE, shortDate, isLeafGuard };

    // Compute initial yMax from visible bars (multiples of 10)
    const visibleMax = this._getHcVisibleMax(0);
    const yMax = Math.ceil(visibleMax / 10) * 10 || 10;
    this._hcCurrentYMax = yMax;

    const yAxisSvg = this._buildHcYAxisSvg(yMax);
    const barsSvg = this._buildHcBarsSvg(yMax);

    const displayW = needsScroll ? REF_W : (YAXIS_W + barAreaW);

    // ── Build table (back side) ──
    const tableRows = hist.map((r, i) => {
      const origIdx = n - 1 - i;
      const prev = i < n - 1 ? hist[i + 1] : null;
      const arrow = this._trendArrow(r.active, prev?.active);
      if (isLeafGuard) {
        const leadGen = Math.max((r.active || 0) - (r.closers || 0), 0);
        return `<tr>
          <td class="bold">${this._esc(r.date)}</td>
          <td class="num"><input type="number" class="hc-edit-input" value="${r.active || 0}" min="0"
            onchange="NationalApp._onHcTableEdit(${ownerIdx},${origIdx},'active',this.value)">${arrow}</td>
          <td class="num"><input type="number" class="hc-edit-input" value="${r.closers || 0}" min="0"
            onchange="NationalApp._onHcTableEdit(${ownerIdx},${origIdx},'closers',this.value)"></td>
          <td class="num hc-dist-cell">${leadGen}</td>
          <td class="num"><input type="number" class="hc-edit-input" value="${r.leaders || 0}" min="0"
            onchange="NationalApp._onHcTableEdit(${ownerIdx},${origIdx},'leaders',this.value)"></td>
          <td class="num"><input type="number" class="hc-edit-input" value="${r.training || 0}" min="0"
            onchange="NationalApp._onHcTableEdit(${ownerIdx},${origIdx},'training',this.value)"></td>
        </tr>`;
      }
      const dist = Math.max((r.active || 0) - (r.leaders || 0), 0);
      return `<tr>
        <td class="bold">${this._esc(r.date)}</td>
        <td class="num"><input type="number" class="hc-edit-input" value="${r.active || 0}" min="0"
          onchange="NationalApp._onHcTableEdit(${ownerIdx},${origIdx},'active',this.value)"
          oninput="NationalApp._updateHcDist(this)">${arrow}</td>
        <td class="num"><input type="number" class="hc-edit-input" value="${r.leaders || 0}" min="0"
          onchange="NationalApp._onHcTableEdit(${ownerIdx},${origIdx},'leaders',this.value)"
          oninput="NationalApp._updateHcDist(this)"></td>
        <td class="num hc-dist-cell">${dist}</td>
        <td class="num"><input type="number" class="hc-edit-input" value="${r.training || 0}" min="0"
          onchange="NationalApp._onHcTableEdit(${ownerIdx},${origIdx},'training',this.value)"></td>
      </tr>`;
    }).join('');

    const tableHeaders = isLeafGuard
      ? '<th>Week</th><th class="num">Active</th><th class="num">Closers</th><th class="num">Lead Gen</th><th class="num">Leaders</th><th class="num">Training</th>'
      : '<th>Week</th><th class="num">Active</th><th class="num">Leaders</th><th class="num">Dist</th><th class="num">Training</th>';

    const tableHtml = `
      <div class="data-table-wrap trend-scroll">
        <table class="data-table">
          <thead><tr>
            ${tableHeaders}
          </tr></thead>
          <tbody>${tableRows}</tbody>
        </table>
      </div>`;

    // ── Assemble flip card (split Y-axis + scrollable bars) ──
    trendEl.innerHTML = `
      <div class="coaching-label">
        Week-over-Week Headcount
        <button class="flip-btn" onclick="NationalApp._flipHcCard()" title="Flip to ${this._hcFlipped ? 'chart' : 'table'} view">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
        </button>
      </div>
      <div class="flip-card${this._hcFlipped ? ' flipped' : ''}" id="hc-flip-card">
        <div class="flip-card-inner">
          <div class="flip-card-front">
            <div class="hc-chart-outer" style="position:relative; max-width:${displayW}px">
              <div style="display:flex">
                <svg id="hc-yaxis-svg" width="${YAXIS_W}" height="${svgH}" style="flex-shrink:0; background:var(--card-bg,#ddeaf5)">${yAxisSvg}</svg>
                <div class="hc-chart-wrap" id="hc-chart-scroll" style="flex:1; min-width:0">
                  <svg id="hc-bars-svg" width="${barAreaW}" height="${svgH}" overflow="hidden">${barsSvg}</svg>
                </div>
              </div>
              <div class="hc-chart-tooltip" id="hc-chart-tt"></div>
            </div>
            <div class="hc-chart-legend">
              <span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch swatch-leaders"></span>Leaders</span>
              <span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch swatch-dist"></span>Distributors</span>
              <span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch swatch-training"></span>Training</span>
              ${isLeafGuard ? '<span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch" style="background:#f59e0b"></span>Closers</span><span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch" style="background:#22c55e"></span>Lead Gen</span>' : ''}
            </div>
          </div>
          <div class="flip-card-back">
            ${tableHtml}
          </div>
        </div>
      </div>`;

    // Attach scroll listener for dynamic Y-axis rescaling
    const scrollEl = document.getElementById('hc-chart-scroll');
    if (scrollEl && needsScroll) {
      this._hcScrollRaf = null;
      scrollEl.addEventListener('scroll', () => {
        if (this._hcScrollRaf) return;
        this._hcScrollRaf = requestAnimationFrame(() => {
          this._hcScrollRaf = null;
          this._onHcScroll();
        });
      });
    }
  },

  // ── Dynamic Y-axis helpers ──
  _hcCurrentYMax: 10,
  _hcScrollRaf: null,
  _hcData: null,

  _getHcVisibleMax(scrollLeft) {
    const d = this._hcData;
    if (!d) return 10;
    const firstVisible = Math.max(0, Math.floor(scrollLeft / d.slotW) - 1);
    const lastVisible = Math.min(d.n - 1, Math.ceil((scrollLeft + d.barAreaVisibleW) / d.slotW));
    let maxVal = 1;
    for (let i = firstVisible; i <= lastVisible; i++) {
      const r = d.hist[i];
      const total = (r.active || 0) + (r.training || 0);
      if (total > maxVal) maxVal = total;
    }
    return maxVal;
  },

  _onHcScroll() {
    const scrollEl = document.getElementById('hc-chart-scroll');
    if (!scrollEl || !this._hcData) return;
    const visibleMax = this._getHcVisibleMax(scrollEl.scrollLeft);
    const yMax = Math.ceil(visibleMax / 10) * 10 || 10;
    if (yMax === this._hcCurrentYMax) return;
    this._hcCurrentYMax = yMax;
    const yAxisEl = document.getElementById('hc-yaxis-svg');
    if (yAxisEl) yAxisEl.innerHTML = this._buildHcYAxisSvg(yMax);
    const barsEl = document.getElementById('hc-bars-svg');
    if (barsEl) barsEl.innerHTML = this._buildHcBarsSvg(yMax);
  },

  _buildHcYAxisSvg(yMax) {
    const d = this._hcData;
    if (!d) return '';
    const yScale = d.plotH / yMax;
    let svg = '';
    for (let val = 0; val <= yMax; val += 10) {
      const y = d.baseY - val * yScale;
      svg += `<text x="${d.YAXIS_W - 6}" y="${y + 3.5}" text-anchor="end" fill="#b0b8c4" font-size="10" font-family="Inter,sans-serif">${val}</text>`;
    }
    return svg;
  },

  _buildHcBarsSvg(yMax) {
    const d = this._hcData;
    if (!d) return '';
    const yScale = d.plotH / yMax;
    let svg = '';

    // Gridlines at multiples of 10
    for (let val = 0; val <= yMax; val += 10) {
      const y = d.baseY - val * yScale;
      svg += `<line x1="0" y1="${y}" x2="${d.barAreaW}" y2="${y}" stroke="#e8ecf1" stroke-width="0.7"/>`;
    }

    const roundTop = (x, y, w, h, r) => {
      if (h <= 0) return '';
      const cr = Math.min(r, h / 2, w / 2);
      return `M${x},${y + h}L${x},${y + cr}Q${x},${y} ${x + cr},${y}L${x + w - cr},${y}Q${x + w},${y} ${x + w},${y + cr}L${x + w},${y + h}Z`;
    };

    const segLabel = (cx, segTop, segH, val, color) => {
      if (segH < d.MIN_LABEL_H || !val) return '';
      const ty = segTop + segH / 2 + 4;
      return `<text x="${cx}" y="${ty}" text-anchor="middle" fill="${color}" font-size="11" font-weight="700" font-family="Inter,sans-serif">${val}</text>`;
    };

    const isLG = this._hcData.isLeafGuard;

    d.hist.forEach((r, i) => {
      const origIdx = d.n - 1 - i;
      const active = r.active || 0;
      const leaders = r.leaders || 0;
      const training = r.training || 0;
      const x = i * d.slotW + d.barOff;
      const cx = x + d.barW / 2;

      if (isLG) {
        // LeafGuard: TWO bars per week side by side
        // Bar 1 (left): Standard — Leaders (blue) + Dist (teal) + Training (dashed)
        // Bar 2 (right): Closers (orange) + Lead Gen (green)
        const halfW = d.barW * 0.46;
        const gap = d.barW * 0.08;
        const x1 = x, x2 = x + halfW + gap;
        const cx1 = x1 + halfW / 2, cx2 = x2 + halfW / 2;

        // Bar 1: Standard headcount
        const dist = Math.max(active - leaders, 0);
        const leaderH = leaders * yScale;
        const distH = dist * yScale;
        const solidH = active * yScale;
        const trainH = training * yScale;
        const solidTop = d.baseY - solidH;
        const trainTop = solidTop - trainH;
        const topmost1 = trainH > 0 ? 'training' : (distH > 0 ? 'dist' : 'leader');

        if (leaderH > 0) {
          svg += topmost1 === 'leader'
            ? `<path d="${roundTop(x1, d.baseY - leaderH, halfW, leaderH, d.BAR_R)}" fill="#5b9cf6" opacity="0.9"/>`
            : `<rect x="${x1}" y="${d.baseY - leaderH}" width="${halfW}" height="${leaderH}" fill="#5b9cf6" opacity="0.9"/>`;
          svg += segLabel(cx1, d.baseY - leaderH, leaderH, leaders, '#fff');
        }
        if (distH > 0) {
          svg += topmost1 === 'dist'
            ? `<path d="${roundTop(x1, solidTop, halfW, distH, d.BAR_R)}" fill="#0ea5a0" opacity="0.85"/>`
            : `<rect x="${x1}" y="${solidTop}" width="${halfW}" height="${distH}" fill="#0ea5a0" opacity="0.85"/>`;
          svg += segLabel(cx1, solidTop, distH, dist, '#fff');
        }
        if (trainH > 0) {
          svg += `<path d="${roundTop(x1 + 0.5, trainTop + 0.5, halfW - 1, trainH - 1, d.BAR_R)}" fill="rgba(139,92,246,0.08)" stroke="#a78bfa" stroke-width="1" stroke-dasharray="3 2"/>`;
          svg += segLabel(cx1, trainTop, trainH, training, '#7c3aed');
        }

        // Bar 2: Closers + Lead Gen
        const closers = r.closers || 0;
        const leadGen = Math.max(active - closers, 0);
        const closerH = closers * yScale;
        const lgH = leadGen * yScale;
        const bar2Top = d.baseY - solidH; // same total height as active
        const topmost2 = lgH > 0 ? 'lg' : 'closer';

        if (closerH > 0) {
          svg += topmost2 === 'closer'
            ? `<path d="${roundTop(x2, d.baseY - closerH, halfW, closerH, d.BAR_R)}" fill="#f59e0b" opacity="0.9"/>`
            : `<rect x="${x2}" y="${d.baseY - closerH}" width="${halfW}" height="${closerH}" fill="#f59e0b" opacity="0.9"/>`;
          svg += segLabel(cx2, d.baseY - closerH, closerH, closers, '#fff');
        }
        if (lgH > 0) {
          svg += topmost2 === 'lg'
            ? `<path d="${roundTop(x2, bar2Top, halfW, lgH, d.BAR_R)}" fill="#22c55e" opacity="0.85"/>`
            : `<rect x="${x2}" y="${bar2Top}" width="${halfW}" height="${lgH}" fill="#22c55e" opacity="0.85"/>`;
          svg += segLabel(cx2, bar2Top, lgH, leadGen, '#fff');
        }

        // Hover target spans both bars
        const totalH = solidH + trainH;
        const topY = trainH > 0 ? trainTop : solidTop;
        svg += `<rect x="${x}" y="${Math.min(topY, d.baseY - 1)}" width="${d.barW}" height="${Math.max(totalH, 4)}" fill="transparent" style="cursor:pointer" onmouseenter="NationalApp._showHcTooltip(event,${origIdx},${d.ownerIdx})" onmouseleave="NationalApp._hideHcTooltip()"/>`;
      } else {
        // Standard: Leaders (blue, bottom) + Dist (teal, middle) + Training (dashed)
        const dist = Math.max(active - leaders, 0);
        const leaderH = leaders * yScale;
        const distH = dist * yScale;
        const solidH = active * yScale;
        const trainH = training * yScale;
        const solidTop = d.baseY - solidH;
        const trainTop = solidTop - trainH;
        const topmost = trainH > 0 ? 'training' : (distH > 0 ? 'dist' : 'leader');

        if (leaderH > 0) {
          if (topmost === 'leader') {
            svg += `<path d="${roundTop(x, d.baseY - leaderH, d.barW, leaderH, d.BAR_R)}" fill="#5b9cf6" opacity="0.9"/>`;
          } else {
            svg += `<rect x="${x}" y="${d.baseY - leaderH}" width="${d.barW}" height="${leaderH}" fill="#5b9cf6" opacity="0.9"/>`;
          }
          svg += segLabel(cx, d.baseY - leaderH, leaderH, leaders, '#fff');
        }
        if (distH > 0) {
          if (topmost === 'dist') {
            svg += `<path d="${roundTop(x, solidTop, d.barW, distH, d.BAR_R)}" fill="#0ea5a0" opacity="0.85"/>`;
          } else {
            svg += `<rect x="${x}" y="${solidTop}" width="${d.barW}" height="${distH}" fill="#0ea5a0" opacity="0.85"/>`;
          }
          svg += segLabel(cx, solidTop, distH, dist, '#fff');
        }
        if (trainH > 0) {
          svg += `<path d="${roundTop(x + 0.5, trainTop + 0.5, d.barW - 1, trainH - 1, d.BAR_R)}" fill="rgba(139,92,246,0.08)" stroke="#a78bfa" stroke-width="1" stroke-dasharray="3 2"/>`;
          svg += segLabel(cx, trainTop, trainH, training, '#7c3aed');
        }
        const totalH = solidH + trainH;
        const topY = trainH > 0 ? trainTop : solidTop;
        svg += `<rect x="${x}" y="${Math.min(topY, d.baseY - 1)}" width="${d.barW}" height="${Math.max(totalH, 4)}" fill="transparent" style="cursor:pointer" onmouseenter="NationalApp._showHcTooltip(event,${origIdx},${d.ownerIdx})" onmouseleave="NationalApp._hideHcTooltip()"/>`;
      }

      // X-axis date label
      svg += `<text x="${cx}" y="${d.baseY + 16}" text-anchor="middle" fill="#8a95a5" font-size="10" font-weight="600" font-family="Inter,sans-serif">${d.shortDate(r.date)}</text>`;
    });

    return svg;
  },

  _hcFlipped: false,

  _flipHcCard() {
    this._hcFlipped = !this._hcFlipped;

    // If flipping back to chart and table data was edited, re-render everything
    if (!this._hcFlipped && this._hcDirty) {
      this._hcDirty = false;
      const ownerIdx = this._hcData?.ownerIdx;
      if (ownerIdx !== undefined) {
        const owner = this.state.owners[ownerIdx];
        if (owner) {
          this._renderHeadcountTrend(owner, ownerIdx);
          return;
        }
      }
    }

    const card = document.getElementById('hc-flip-card');
    if (card) card.classList.toggle('flipped', this._hcFlipped);
    const btn = card?.parentElement?.querySelector('.flip-btn');
    if (btn) btn.title = 'Flip to ' + (this._hcFlipped ? 'chart' : 'table') + ' view';
  },

  // ── Headcount table inline-edit helpers ──
  _hcDirty: false,

  _onHcTableEdit(ownerIdx, histIdx, field, value) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    const entry = owner.headcountHistory[histIdx];
    if (!entry) return;
    entry[field] = parseInt(value) || 0;
    // If editing the most recent entry, keep current headcount in sync
    if (histIdx === owner.headcountHistory.length - 1) {
      if (field in owner.headcount) owner.headcount[field] = entry[field];
    }
    this._hcDirty = true;

    // Debounced save to spreadsheet
    const saveKey = `${ownerIdx}_${histIdx}`;
    if (this._hcSaveTimers?.[saveKey]) clearTimeout(this._hcSaveTimers[saveKey]);
    if (!this._hcSaveTimers) this._hcSaveTimers = {};
    this._hcSaveTimers[saveKey] = setTimeout(() => {
      this._saveHcRow(owner, entry);
      delete this._hcSaveTimers[saveKey];
    }, 1200);
  },

  _hcSaveTimers: null,

  async _saveHcRow(owner, entry) {
    const sheetName = owner._sheetName || owner.tab || owner.name;
    const dist = Math.max((entry.active || 0) - (entry.leaders || 0), 0);
    // Get campaign label for writing to the consolidated tab
    const campaignCfg = NATIONAL_CONFIG.campaigns[this.state.campaign];
    const campaignLabel = campaignCfg?.label || '';
    try {
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'updateHeadcount',
          ownerName: owner.name,
          date: entry.date,
          active: entry.active || 0,
          leaders: entry.leaders || 0,
          dist: dist,
          training: entry.training || 0,
          campaignLabel: campaignLabel
        })
      });
      const result = await resp.json();
      if (result.error) console.warn('[HC Save] Error:', result.error);
      else console.log('[HC Save] Saved', owner.name, entry.date, '→', campaignLabel);
    } catch (err) {
      console.warn('[HC Save] Network error:', err.message);
    }
  },

  _updateHcDist(input) {
    const row = input.closest('tr');
    if (!row) return;
    const inputs = row.querySelectorAll('.hc-edit-input');
    const active = parseInt(inputs[0]?.value) || 0;
    const leaders = parseInt(inputs[1]?.value) || 0;
    const distCell = row.querySelector('.hc-dist-cell');
    if (distCell) distCell.textContent = Math.max(active - leaders, 0);
  },

  // ── Headcount chart tooltip helpers ──
  _showHcTooltip(event, origIdx, ownerIdx) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    const r = owner.headcountHistory[origIdx];
    if (!r) return;
    const dist = r.active - r.leaders;
    const tt = document.getElementById('hc-chart-tt');
    if (!tt) return;

    tt.innerHTML = `
      <div style="font-weight:700;margin-bottom:4px">${this._esc(r.date)}</div>
      <div><span class="tt-swatch" style="background:#0ea5a0"></span>Distributors: <strong>${dist}</strong></div>
      <div><span class="tt-swatch" style="background:#3b82f6"></span>Leaders: <strong>${r.leaders}</strong></div>
      <div style="border-top:1px solid rgba(255,255,255,0.2);margin:4px 0;padding-top:4px">Active: <strong>${r.active}</strong></div>
      ${r.training ? `<div><span class="tt-swatch" style="background:rgba(139,92,246,0.3);border:1px dashed #8b5cf6"></span>Training: <strong>${r.training}</strong></div>` : ''}`;

    // Position above the hovered bar — tooltip is in .hc-chart-outer (relative)
    const outer = tt.parentElement;
    const rect = event.target.getBoundingClientRect();
    const outerRect = outer.getBoundingClientRect();

    let left = rect.left + rect.width / 2 - outerRect.left;
    tt.style.left = left + 'px';
    tt.style.top = (rect.top - outerRect.top - 8) + 'px';
    tt.style.transform = 'translateX(-50%) translateY(-100%)';
    tt.classList.add('visible');

    // Clamp so tooltip doesn't overflow the outer container
    requestAnimationFrame(() => {
      const ttRect = tt.getBoundingClientRect();
      if (ttRect.left < outerRect.left + 4) {
        tt.style.left = '4px';
        tt.style.transform = 'translateY(-100%)';
      }
      if (ttRect.right > outerRect.right - 4) {
        tt.style.transform = 'translateX(-100%) translateY(-100%)';
      }
    });
  },

  _hideHcTooltip() {
    const tt = document.getElementById('hc-chart-tt');
    if (tt) tt.classList.remove('visible');
  },

  // ── Trend arrow helper ──
  _trendArrow(current, previous) {
    if (previous === null || previous === undefined) return '';
    const diff = current - previous;
    if (diff > 0) return `<span class="trend-up">+${diff}</span>`;
    if (diff < 0) return `<span class="trend-down">${diff}</span>`;
    return '';
  },

  // ── Render production week-over-week bar chart (newest-first, dynamic Y-axis) ──
  _renderProductionTrend(owner, ownerIdx) {
    const trendEl = document.getElementById('health-prod-trend');
    if (!trendEl) return;

    const rawHist = owner.productionHistory || [];
    if (!rawHist.length) { trendEl.style.display = 'none'; return; }
    trendEl.style.display = '';

    // Reverse so newest week is on the LEFT
    // Keep zero-production weeks visible so coaches can fill them in during meetings
    // Track original indices for tooltip accuracy
    const hist = rawHist
      .map((r, idx) => ({ ...r, _origIdx: idx }))
      .reverse();
    const n = hist.length;
    if (!n) { trendEl.style.display = 'none'; return; }

    const shortDate = (d) => {
      if (!d) return '';
      const s = String(d);
      // If it's a full Date toString(), parse it
      if (s.length > 10 && s.indexOf('GMT') >= 0) {
        const dt = new Date(s);
        if (!isNaN(dt)) return (dt.getMonth()+1) + '/' + dt.getDate();
      }
      const parts = s.split('/');
      return parts.length >= 2 ? parts[0] + '/' + parts[1] : s;
    };

    // ── Layout constants (match headcount chart) ──
    const VISIBLE = 5;
    const YAXIS_W = 36;
    const PAD_R = 10, PAD_T = 14, PAD_B = 30;
    const BAR_R = 5;
    const GAP = 0.14;
    const MIN_LABEL_H = 14;
    const svgH = 222;
    const plotH = svgH - PAD_T - PAD_B;
    const REF_W = 500;
    const barAreaVisibleW = REF_W - YAXIS_W;
    const slotW = barAreaVisibleW / VISIBLE;
    const barAreaW = n * slotW + PAD_R;
    const needsScroll = n > VISIBLE;
    const barW = slotW * (1 - GAP);
    const barOff = (slotW - barW) / 2;
    const baseY = PAD_T + plotH;

    // Store chart state for scroll-driven re-renders
    this._prodData = { hist, ownerIdx, n, slotW, barW, barOff, barAreaW, barAreaVisibleW, plotH, PAD_T, PAD_R, BAR_R, GAP, MIN_LABEL_H, svgH, baseY, YAXIS_W, VISIBLE, shortDate };

    // Compute initial yMax from visible bars with dynamic step
    const visibleMax = this._getProdVisibleMax(0);
    const prodStep = this._prodStepForMax(visibleMax);
    const yMax = Math.ceil(visibleMax / prodStep) * prodStep || prodStep;
    this._prodCurrentYMax = yMax;

    const yAxisSvg = this._buildProdYAxisSvg(yMax);
    const barsSvg = this._buildProdBarsSvg(yMax);

    const displayW = needsScroll ? REF_W : (YAXIS_W + barAreaW);

    // ── Build editable table (back side) — per-product columns ──
    // Determine product names from the data (use first entry that has products)
    let tableProductNames = [];
    for (const r of hist) {
      if (r.products && Object.keys(r.products).length) {
        tableProductNames = Object.keys(r.products);
        break;
      }
    }
    const hasProducts = tableProductNames.length > 0;

    const tableRows = hist.map((r, i, arr) => {
      const realOrigIdx = r._origIdx !== undefined ? r._origIdx : (n - 1 - hist.indexOf(r));
      const prev = i < arr.length - 1 ? arr[i + 1] : null;
      const arrow = this._trendArrow(r.tA, prev?.tA);
      const pct = r.tG > 0 ? Math.round((r.tA / r.tG) * 100) : 0;
      const pctClass = pct >= 100 ? 'pct-green' : pct >= 80 ? 'pct-yellow' : pct >= 60 ? 'pct-orange' : 'pct-red';

      let productCells = '';
      if (hasProducts) {
        for (const pName of tableProductNames) {
          const pData = (r.products || {})[pName] || { actual: 0, goal: 0 };
          const escapedPName = this._esc(pName).replace(/'/g, "\\'");
          productCells += `<td class="num"><input type="number" class="hc-edit-input" value="${pData.actual || ''}" min="0" placeholder="—"
            onchange="NationalApp._onProdProductEdit(${ownerIdx},${realOrigIdx},'${escapedPName}','actual',this.value)"></td>
          <td class="num"><input type="number" class="hc-edit-input" value="${pData.goal || ''}" min="0" placeholder="—"
            onchange="NationalApp._onProdProductEdit(${ownerIdx},${realOrigIdx},'${escapedPName}','goal',this.value)"></td>`;
        }
      }

      return `<tr>
        <td class="bold">${this._esc(r.date)}</td>
        ${productCells}
        ${!hasProducts ? `<td class="num"><input type="number" class="hc-edit-input" value="${r.tA || ''}" min="0" placeholder="—"
          onchange="NationalApp._onProdTableEdit(${ownerIdx},${realOrigIdx},'tA',this.value)">${arrow}</td>
        <td class="num"><input type="number" class="hc-edit-input" value="${r.tG || ''}" min="0" placeholder="—"
          onchange="NationalApp._onProdTableEdit(${ownerIdx},${realOrigIdx},'tG',this.value)"></td>` : ''}
        <td class="num"><span class="prod-pct-badge ${pctClass}">${pct}%</span></td>
      </tr>`;
    }).join('');

    // Build per-product header columns
    let productHeaders = '';
    if (hasProducts) {
      for (const pName of tableProductNames) {
        productHeaders += `<th class="num">${this._esc(pName)}</th><th class="num">Goal</th>`;
      }
    }

    const tableHtml = `
      <div class="data-table-wrap trend-scroll">
        <table class="data-table">
          <thead><tr>
            <th>Week</th>${productHeaders}${!hasProducts ? '<th class="num">Actual</th><th class="num">Goal</th>' : ''}<th class="num">%</th>
          </tr></thead>
          <tbody>${tableRows}</tbody>
        </table>
      </div>`;

    // ── Assemble flip card (split Y-axis + scrollable bars) ──
    trendEl.innerHTML = `
      <div class="coaching-label">
        Week-over-Week Production
        <button class="flip-btn" onclick="NationalApp._flipProdCard()" title="Flip to ${this._prodFlipped ? 'chart' : 'table'} view">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
        </button>
      </div>
      <div class="flip-card${this._prodFlipped ? ' flipped' : ''}" id="prod-flip-card">
        <div class="flip-card-inner">
          <div class="flip-card-front">
            <div class="prod-chart-outer" style="position:relative; max-width:${displayW}px">
              <div style="display:flex">
                <svg id="prod-yaxis-svg" width="${YAXIS_W}" height="${svgH}" style="flex-shrink:0; background:var(--card-bg,#ddeaf5)">${yAxisSvg}</svg>
                <div class="hc-chart-wrap" id="prod-chart-scroll" style="flex:1; min-width:0">
                  <svg id="prod-bars-svg" width="${barAreaW}" height="${svgH}" overflow="hidden">${barsSvg}</svg>
                </div>
              </div>
              <div class="hc-chart-tooltip" id="prod-chart-tt"></div>
            </div>
            <div class="hc-chart-legend">
              ${tableProductNames.length > 1 && this.state.campaign !== 'leafguard'
                ? tableProductNames.map((pName, pi) => {
                    let style;
                    if (pi === 0) style = 'background:#888';
                    else if (pi === 1) style = 'background:repeating-conic-gradient(#888 0% 25%, rgba(255,255,255,0.3) 0% 50%) 0 0/6px 6px';
                    else style = 'background:repeating-linear-gradient(45deg,#888,#888 2px,rgba(255,255,255,0.3) 2px,rgba(255,255,255,0.3) 4px)';
                    return `<span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch" style="${style}"></span>${this._esc(pName)}</span>`;
                  }).join('')
                  + '<span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch" style="background:none;border-top:2px dashed #888;height:0;width:10px;border-radius:0"></span>Goal</span>'
                  + '<span class="hc-chart-legend-item" style="margin-left:8px;font-size:10px;color:#8a95a5">Colors = goal %</span>'
                : this.state.campaign === 'leafguard'
                  ? '<span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch swatch-prod-actual"></span>Gross Sales</span><span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch swatch-prod-goal"></span>Goal</span>'
                  : '<span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch swatch-prod-actual"></span>Actual</span><span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch swatch-prod-goal"></span>Goal</span>'
              }
            </div>
          </div>
          <div class="flip-card-back">
            ${tableHtml}
          </div>
        </div>
      </div>`;

    // Attach scroll listener for dynamic Y-axis rescaling
    const scrollEl = document.getElementById('prod-chart-scroll');
    if (scrollEl && needsScroll) {
      this._prodScrollRaf = null;
      scrollEl.addEventListener('scroll', () => {
        if (this._prodScrollRaf) return;
        this._prodScrollRaf = requestAnimationFrame(() => {
          this._prodScrollRaf = null;
          this._onProdScroll();
        });
      });
    }
  },

  // ── Production dynamic Y-axis helpers ──
  _prodCurrentYMax: 5,
  _prodScrollRaf: null,
  _prodData: null,

  _getProdVisibleMax(scrollLeft) {
    const d = this._prodData;
    if (!d) return 5;
    const firstVisible = Math.max(0, Math.floor(scrollLeft / d.slotW) - 1);
    const lastVisible = Math.min(d.n - 1, Math.ceil((scrollLeft + d.barAreaVisibleW) / d.slotW));
    const isLeafGuard = this.state.campaign === 'leafguard';
    let maxVal = 1;
    for (let i = firstVisible; i <= lastVisible; i++) {
      const r = d.hist[i];
      if (r.products && Object.keys(r.products).length > 1) {
        if (isLeafGuard) {
          // LeafGuard chart: only Gross Sales
          const gs = r.products['Gross Sales'] || { actual: 0, goal: 0 };
          const v = Math.max(gs.actual || 0, gs.goal || 0);
          if (v > maxVal) maxVal = v;
        } else {
          for (const pName in r.products) {
            const pv = r.products[pName];
            const v = Math.max(pv.actual || 0, pv.goal || 0);
            if (v > maxVal) maxVal = v;
          }
        }
      } else {
        const barTop = Math.max(r.tA || 0, r.tG || 0);
        if (barTop > maxVal) maxVal = barTop;
      }
    }
    return maxVal;
  },

  // Dynamic step size based on value magnitude
  _prodStepForMax(maxVal) {
    if (maxVal <= 50) return 10;
    if (maxVal <= 200) return 25;
    if (maxVal <= 500) return 50;
    if (maxVal <= 2000) return 250;
    if (maxVal <= 10000) return 1000;
    if (maxVal <= 50000) return 5000;
    if (maxVal <= 200000) return 25000;
    return 50000;
  },

  _onProdScroll() {
    const scrollEl = document.getElementById('prod-chart-scroll');
    if (!scrollEl || !this._prodData) return;
    const visibleMax = this._getProdVisibleMax(scrollEl.scrollLeft);
    const prodStep = this._prodStepForMax(visibleMax);
    const yMax = Math.ceil(visibleMax / prodStep) * prodStep || prodStep;
    if (yMax === this._prodCurrentYMax) return;
    this._prodCurrentYMax = yMax;
    const yAxisEl = document.getElementById('prod-yaxis-svg');
    if (yAxisEl) yAxisEl.innerHTML = this._buildProdYAxisSvg(yMax);
    const barsEl = document.getElementById('prod-bars-svg');
    if (barsEl) barsEl.innerHTML = this._buildProdBarsSvg(yMax);
  },

  _buildProdYAxisSvg(yMax) {
    const d = this._prodData;
    if (!d) return '';
    const yScale = d.plotH / yMax;
    const step = this._prodStepForMax(yMax);
    const formatTick = (v) => v >= 1000 ? (v / 1000) + 'k' : String(v);
    let svg = '';
    for (let val = 0; val <= yMax; val += step) {
      const y = d.baseY - val * yScale;
      svg += `<text x="${d.YAXIS_W - 6}" y="${y + 3.5}" text-anchor="end" fill="#b0b8c4" font-size="10" font-family="Inter,sans-serif">${formatTick(val)}</text>`;
    }
    return svg;
  },

  // Product color palette for grouped bars
  _PROD_COLORS: ['#3b82f6', '#06b6d4', '#8b5cf6', '#f59e0b', '#ec4899'],

  _buildProdBarsSvg(yMax) {
    const d = this._prodData;
    if (!d) return '';
    const yScale = d.plotH / yMax;
    const step = this._prodStepForMax(yMax);
    let svg = '';

    // Gridlines
    for (let val = 0; val <= yMax; val += step) {
      const y = d.baseY - val * yScale;
      svg += `<line x1="0" y1="${y}" x2="${d.barAreaW}" y2="${y}" stroke="#e8ecf1" stroke-width="0.7"/>`;
    }

    const roundTop = (x, y, w, h, r) => {
      if (h <= 0) return '';
      const cr = Math.min(r, h / 2, w / 2);
      return `M${x},${y + h}L${x},${y + cr}Q${x},${y} ${x + cr},${y}L${x + w - cr},${y}Q${x + w},${y} ${x + w},${y + cr}L${x + w},${y + h}Z`;
    };

    const segLabel = (cx, segTop, segH, val, color) => {
      if (segH < d.MIN_LABEL_H || !val) return '';
      const ty = segTop + segH / 2 + 4;
      return `<text x="${cx}" y="${ty}" text-anchor="middle" fill="${color}" font-size="11" font-weight="700" font-family="Inter,sans-serif">${val}</text>`;
    };

    // Detect if multi-product — use first entry with products
    let productNames = [];
    for (const r of d.hist) {
      if (r.products && Object.keys(r.products).length > 1) {
        productNames = Object.keys(r.products);
        break;
      }
    }
    const isLeafGuard = this.state?.campaign === 'leafguard';
    const isMulti = productNames.length > 1 && !isLeafGuard;
    const formatLabel = (v) => v >= 1000 ? (v / 1000) + 'k' : String(v);

    d.hist.forEach((r, i) => {
      const origIdx = r._origIdx !== undefined ? r._origIdx : (d.n - 1 - i);
      const slotX = i * d.slotW + d.barOff;
      const slotCX = slotX + d.barW / 2;

      if (isLeafGuard && r.products) {
        // LeafGuard: single Gross Sales bar
        const gs = r.products['Gross Sales'] || { actual: 0, goal: 0 };
        const actual = gs.actual || 0;
        const goal = gs.goal || 0;
        const actualH = actual * yScale;
        const goalH = goal * yScale;
        const actualTop = d.baseY - actualH;
        const pct = goal > 0 ? (actual / goal) : 0;
        const barColor = pct >= 1 ? '#22c55e' : pct >= 0.8 ? '#f0b429' : pct >= 0.6 ? '#f97316' : '#e53535';

        if (actualH > 0) {
          svg += `<path d="${roundTop(slotX, actualTop, d.barW, actualH, d.BAR_R)}" fill="${barColor}" opacity="0.85"/>`;
          svg += segLabel(slotCX, actualTop, actualH, formatLabel(actual), '#fff');
        }
        if (goal > 0) {
          const goalY = d.baseY - goalH;
          svg += `<line x1="${slotX - 2}" y1="${goalY}" x2="${slotX + d.barW + 2}" y2="${goalY}" stroke="#6366f1" stroke-width="2" stroke-dasharray="4 2" opacity="0.7"/>`;
        }
        const topY = Math.min(actualH > 0 ? actualTop : d.baseY, goal > 0 ? d.baseY - goalH : d.baseY);
        const totalH = d.baseY - topY;
        svg += `<rect x="${slotX}" y="${Math.min(topY, d.baseY - 1)}" width="${d.barW}" height="${Math.max(totalH, 4)}" fill="transparent" style="cursor:pointer" onmouseenter="NationalApp._showProdTooltip(event,${origIdx},${d.ownerIdx},'Gross Sales')" onmouseleave="NationalApp._hideProdTooltip()"/>`;
      } else if (isMulti) {
        // ── Grouped bars: goal-attainment colors, diamond pattern for product 2+ ──
        // Filter to products that have data or a goal this week
        const activeProducts = productNames.filter(pName => {
          const pv = (r.products && r.products[pName]) || { actual: 0, goal: 0 };
          return (pv.actual || 0) > 0 || (pv.goal || 0) > 0;
        });
        const visibleCount = activeProducts.length || 1;
        const subW = d.barW / visibleCount;
        activeProducts.forEach((pName, vi) => {
          const pi = productNames.indexOf(pName); // original product index for pattern
          const pv = (r.products && r.products[pName]) || { actual: 0, goal: 0 };
          const actual = pv.actual || 0;
          const goal = pv.goal || 0;
          const bx = slotX + vi * subW;
          const bcx = bx + subW / 2;
          const actualH = actual * yScale;
          const goalH = goal * yScale;
          const actualTop = d.baseY - actualH;

          // Goal-attainment color for all bars
          const pct = goal > 0 ? (actual / goal) : 0;
          const barColor = pct >= 1 ? '#22c55e' : pct >= 0.8 ? '#f0b429' : pct >= 0.6 ? '#f97316' : '#e53535';

          // Corners: round outer edges of the group
          const rLeft = vi === 0 ? d.BAR_R : 0;
          const rRight = vi === visibleCount - 1 ? d.BAR_R : 0;

          if (actualH > 0) {
            const barPath = roundTop(bx, actualTop, subW, actualH, Math.min(rLeft, rRight) || 2);
            if (pi === 0) {
              // Product 1: solid fill
              svg += `<path d="${barPath}" fill="${barColor}" opacity="0.85"/>`;
            } else if (pi === 1) {
              // Product 2: diamond checkerboard
              const patId = 'pat-diamond-' + i + '-' + pi;
              svg += `<defs><pattern id="${patId}" width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)">
                <rect width="8" height="8" fill="${barColor}" opacity="0.85"/>
                <rect x="0" y="0" width="4" height="4" fill="rgba(255,255,255,0.22)"/>
                <rect x="4" y="4" width="4" height="4" fill="rgba(255,255,255,0.22)"/>
              </pattern></defs>`;
              svg += `<path d="${barPath}" fill="url(#${patId})"/>`;
            } else {
              // Product 3+: diagonal stripes
              const patId = 'pat-stripe-' + i + '-' + pi;
              svg += `<defs><pattern id="${patId}" width="6" height="6" patternUnits="userSpaceOnUse" patternTransform="rotate(45)">
                <rect width="6" height="6" fill="${barColor}" opacity="0.85"/>
                <rect x="0" y="0" width="3" height="6" fill="rgba(255,255,255,0.22)"/>
              </pattern></defs>`;
              svg += `<path d="${barPath}" fill="url(#${patId})"/>`;
            }
            svg += segLabel(bcx, actualTop, actualH, actual, '#fff');
          }

          // Goal marker line
          if (goal > 0) {
            const goalY = d.baseY - goalH;
            svg += `<line x1="${bx}" y1="${goalY}" x2="${bx + subW}" y2="${goalY}" stroke="${barColor}" stroke-width="2" stroke-dasharray="3 2" opacity="0.7"/>`;
          }

          // Hover target per sub-bar
          const topY = Math.min(actualH > 0 ? actualTop : d.baseY, goal > 0 ? d.baseY - goalH : d.baseY);
          const totalH = d.baseY - topY;
          svg += `<rect x="${bx}" y="${Math.min(topY, d.baseY - 1)}" width="${subW}" height="${Math.max(totalH, 4)}" fill="transparent" style="cursor:pointer" onmouseenter="NationalApp._showProdTooltip(event,${origIdx},${d.ownerIdx},'${pName}')" onmouseleave="NationalApp._hideProdTooltip()"/>`;
        });
      } else {
        // ── Single-product: original behavior ──
        const actual = r.tA || 0;
        const goal = r.tG || 0;
        const actualH = actual * yScale;
        const goalH = goal * yScale;
        const actualTop = d.baseY - actualH;
        const pct = goal > 0 ? (actual / goal) : 0;
        const barColor = pct >= 1 ? '#22c55e' : pct >= 0.8 ? '#f0b429' : pct >= 0.6 ? '#f97316' : '#e53535';

        if (actualH > 0) {
          svg += `<path d="${roundTop(slotX, actualTop, d.barW, actualH, d.BAR_R)}" fill="${barColor}" opacity="0.85"/>`;
          svg += segLabel(slotCX, actualTop, actualH, actual, '#fff');
        }

        if (goal > 0) {
          const goalY = d.baseY - goalH;
          svg += `<line x1="${slotX - 2}" y1="${goalY}" x2="${slotX + d.barW + 2}" y2="${goalY}" stroke="#6366f1" stroke-width="2" stroke-dasharray="4 2" opacity="0.7"/>`;
        }

        const topY = Math.min(actualH > 0 ? actualTop : d.baseY, goal > 0 ? d.baseY - goalH : d.baseY);
        const totalH = d.baseY - topY;
        svg += `<rect x="${slotX}" y="${Math.min(topY, d.baseY - 1)}" width="${d.barW}" height="${Math.max(totalH, 4)}" fill="transparent" style="cursor:pointer" onmouseenter="NationalApp._showProdTooltip(event,${origIdx},${d.ownerIdx})" onmouseleave="NationalApp._hideProdTooltip()"/>`;
      }

      // X-axis date label
      svg += `<text x="${slotCX}" y="${d.baseY + 16}" text-anchor="middle" fill="#8a95a5" font-size="10" font-weight="600" font-family="Inter,sans-serif">${d.shortDate(r.date)}</text>`;
    });

    return svg;
  },

  _prodFlipped: false,

  _flipProdCard() {
    this._prodFlipped = !this._prodFlipped;

    // If flipping back to chart and table data was edited, re-render everything
    if (!this._prodFlipped && this._prodDirty) {
      this._prodDirty = false;
      const ownerIdx = this._prodData?.ownerIdx;
      if (ownerIdx !== undefined) {
        const owner = this.state.owners[ownerIdx];
        if (owner) {
          this._renderProductionTrend(owner, ownerIdx);
          return;
        }
      }
    }

    const card = document.getElementById('prod-flip-card');
    if (card) card.classList.toggle('flipped', this._prodFlipped);
    const btn = card?.parentElement?.querySelector('.flip-btn');
    if (btn) btn.title = 'Flip to ' + (this._prodFlipped ? 'chart' : 'table') + ' view';
  },

  // ── Production table inline-edit helpers ──
  _prodDirty: false,

  _onProdTableEdit(ownerIdx, histIdx, field, value) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    const entry = owner.productionHistory[histIdx];
    if (!entry) return;
    entry[field] = parseInt(value) || 0;
    // If editing the most recent entry, keep current production in sync
    if (histIdx === owner.productionHistory.length - 1) {
      if (field === 'tA') owner.production.totalActual = entry.tA;
      if (field === 'tG') owner.production.totalGoal = entry.tG;
    }
    this._prodDirty = true;

    // Update % badge live
    const row = event?.target?.closest('tr');
    if (row) {
      const pctCell = row.querySelector('.prod-pct-badge');
      if (pctCell) {
        const pct = entry.tG > 0 ? Math.round((entry.tA / entry.tG) * 100) : 0;
        const pctClass = pct >= 100 ? 'pct-green' : pct >= 80 ? 'pct-yellow' : pct >= 60 ? 'pct-orange' : 'pct-red';
        pctCell.textContent = pct + '%';
        pctCell.className = 'prod-pct-badge ' + pctClass;
      }
    }

    // Debounced save to spreadsheet
    const saveKey = `prod_${ownerIdx}_${histIdx}`;
    if (this._hcSaveTimers?.[saveKey]) clearTimeout(this._hcSaveTimers[saveKey]);
    if (!this._hcSaveTimers) this._hcSaveTimers = {};
    this._hcSaveTimers[saveKey] = setTimeout(() => {
      this._saveProdRow(owner, entry);
      delete this._hcSaveTimers[saveKey];
    }, 1200);
  },

  // ── Per-product inline edit handler (for campaigns with product breakdown) ──
  _onProdProductEdit(ownerIdx, histIdx, productName, field, value) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    const entry = owner.productionHistory[histIdx];
    if (!entry) return;
    if (!entry.products) entry.products = {};
    if (!entry.products[productName]) entry.products[productName] = { actual: 0, goal: 0 };
    entry.products[productName][field] = parseInt(value) || 0;

    // Recompute totals from all products
    let totalProd = 0, totalGoal = 0;
    for (const pName in entry.products) {
      totalProd += entry.products[pName].actual || 0;
      totalGoal += entry.products[pName].goal || 0;
    }
    entry.tA = totalProd;
    entry.tG = totalGoal;

    // If editing the most recent entry, keep current production in sync
    if (histIdx === owner.productionHistory.length - 1) {
      owner.production.totalActual = totalProd;
      owner.production.totalGoal = totalGoal;
      if (owner.production.products?.[productName]) {
        owner.production.products[productName].actual = entry.products[productName].actual;
        owner.production.products[productName].goal = entry.products[productName].goal;
      }
    }

    // Update % badge live
    const row = event?.target?.closest('tr');
    if (row) {
      const pctCell = row.querySelector('.prod-pct-badge');
      if (pctCell) {
        const pct = totalGoal > 0 ? Math.round((totalProd / totalGoal) * 100) : 0;
        const pctClass = pct >= 100 ? 'pct-green' : pct >= 80 ? 'pct-yellow' : pct >= 60 ? 'pct-orange' : 'pct-red';
        pctCell.textContent = pct + '%';
        pctCell.className = 'prod-pct-badge ' + pctClass;
      }
    }

    // Debounced save to spreadsheet
    const saveKey = `prod_${ownerIdx}_${histIdx}`;
    if (this._hcSaveTimers?.[saveKey]) clearTimeout(this._hcSaveTimers[saveKey]);
    if (!this._hcSaveTimers) this._hcSaveTimers = {};
    this._hcSaveTimers[saveKey] = setTimeout(() => {
      this._saveProdRow(owner, entry);
      delete this._hcSaveTimers[saveKey];
    }, 1200);
  },

  async _saveProdRow(owner, entry) {
    const sheetName = owner._sheetName || owner.tab || owner.name;
    const campaignLabel = this._getCampaignLabel();
    // Build per-product payload for backend
    const products = entry.products || {};
    const productKeys = Object.keys(products);
    try {
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'updateProduction',
          ownerName: sheetName,
          date: entry.date,
          campaignLabel: campaignLabel,
          // Send per-product data for campaigns that have it
          products: productKeys.length > 0 ? products : null,
          // Legacy fallback params
          internet: entry.tA || 0,
          wireless: 0,
          dtv: 0,
          goals: entry.tG ? String(entry.tG) : ''
        })
      });
      const result = await resp.json();
      if (result.error) console.warn('[Prod Save] Error:', result.error);
      else {
        console.log('[Prod Save] Saved', sheetName, entry.date, productKeys.length ? productKeys : 'legacy');
        this._invalidateOdCache();
      }
    } catch (err) {
      console.warn('[Prod Save] Network error:', err.message);
    }
  },

  // ── Production chart tooltip helpers ──
  _showProdTooltip(event, origIdx, ownerIdx, productName) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    const r = owner.productionHistory[origIdx];
    if (!r) return;
    const tt = document.getElementById('prod-chart-tt');
    if (!tt) return;

    if (productName && r.products && r.products[productName]) {
      // Multi-product tooltip — show specific product
      const pv = r.products[productName];
      const actual = pv.actual || 0;
      const goal = pv.goal || 0;
      const pct = goal > 0 ? Math.round((actual / goal) * 100) : 0;
      const pi = Object.keys(r.products).indexOf(productName);
      const color = this._PROD_COLORS[pi % this._PROD_COLORS.length];
      tt.innerHTML = `
        <div style="font-weight:700;margin-bottom:4px">${this._esc(r.date)}</div>
        <div style="font-weight:600;color:${color};margin-bottom:2px">${this._esc(productName)}</div>
        <div>Actual: <strong>${actual}</strong></div>
        <div>Goal: <strong>${goal}</strong></div>
        <div style="border-top:1px solid rgba(255,255,255,0.2);margin:4px 0;padding-top:4px">${pct}% of goal</div>`;
    } else {
      // Single-product tooltip
      const pct = r.tG > 0 ? Math.round((r.tA / r.tG) * 100) : 0;
      tt.innerHTML = `
        <div style="font-weight:700;margin-bottom:4px">${this._esc(r.date)}</div>
        <div><span class="tt-swatch" style="background:#22c55e"></span>Actual: <strong>${r.tA}</strong></div>
        <div><span class="tt-swatch" style="background:#6366f1;border-radius:0;height:2px;width:10px;border-top:2px dashed #6366f1;background:none"></span>Goal: <strong>${r.tG}</strong></div>
        <div style="border-top:1px solid rgba(255,255,255,0.2);margin:4px 0;padding-top:4px">${pct}% of goal</div>`;
    }

    tt.classList.add('visible');
    const outer = tt.closest('.prod-chart-outer');
    if (!outer) return;
    const rect = outer.getBoundingClientRect();
    const bar = event.target.getBoundingClientRect();
    let left = bar.left - rect.left + bar.width / 2 - tt.offsetWidth / 2;
    left = Math.max(4, Math.min(left, rect.width - tt.offsetWidth - 4));
    tt.style.left = left + 'px';
    tt.style.top = (bar.top - rect.top - tt.offsetHeight - 6) + 'px';
  },

  _hideProdTooltip() {
    const tt = document.getElementById('prod-chart-tt');
    if (tt) tt.classList.remove('visible');
  },

  // ── Goal input handler ──
  _updateGoal(ownerIdx, field, value) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    owner.nextGoals[field] = parseInt(value) || 0;
  },

  // ── Submit goals — writes to consolidated tab for next week ──
  async _submitGoals(ownerIdx) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    const cfg = NATIONAL_CONFIG.campaigns[this.state.campaign];
    if (!cfg) return;

    // Collect goal values from all goal inputs in the DOM (not just production.products,
    // since Set Goals shows ALL products including ones not yet sold)
    const goals = {};
    let anyGoal = false;

    const goalInputs = document.querySelectorAll('#health-goals .goal-input');
    if (goalInputs.length > 0) {
      goalInputs.forEach(el => {
        if (el.value) {
          // Extract product name from the id: "goal-{product}-{ownerIdx}"
          const idParts = el.id.replace('goal-', '').replace('-' + ownerIdx, '');
          // Map back to original product name from the label
          const label = el.closest('.goal-field')?.querySelector('.goal-field-label')?.textContent?.replace(' Goal', '') || idParts;
          goals[label] = parseInt(el.value) || 0;
          if (goals[label]) anyGoal = true;
        }
      });
    } else {
      const totalEl = document.getElementById('goal-total-' + ownerIdx);
      if (totalEl && totalEl.value) {
        goals['Total'] = parseInt(totalEl.value) || 0;
        if (goals['Total']) anyGoal = true;
      }
    }

    if (!anyGoal) return;

    const note = document.getElementById('goal-submit-note-' + ownerIdx);
    const btn = document.querySelector(`#health-goals .hc-submit-btn`);
    if (btn) { btn.disabled = true; btn.textContent = 'Saving...'; }

    try {
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'saveGoals',
          ownerName: owner._sheetName || owner.tab || owner.name,
          campaignLabel: cfg.label,
          campaignKey: this.state.campaign,
          goals: goals
        })
      });
      const result = await resp.json();
      if (result.error) {
        console.warn('[Goals] Error:', result.error);
        if (note) { note.textContent = 'Error: ' + result.error; note.classList.add('show'); }
      } else {
        console.log('[Goals] Saved:', result);
        if (note) { note.textContent = 'Goals saved for ' + (result.date || 'next week'); note.classList.add('show'); }
      }
    } catch (err) {
      console.warn('[Goals] Network error:', err.message);
      if (note) { note.textContent = 'Network error'; note.classList.add('show'); }
    }

    if (btn) { btn.disabled = false; btn.textContent = 'Submit Goals'; }
    if (note) setTimeout(() => note.classList.remove('show'), 3000);
  },

  // ── Notes log: add / delete ──
  async _addNote(ownerIdx) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    const ta = document.getElementById('note-input-' + ownerIdx);
    const text = ta?.value?.trim();
    if (!text) return;

    const coachName = this.state.session?.name || 'Unknown';
    const btn = ta.nextElementSibling;
    if (btn) { btn.disabled = true; btn.textContent = '...'; }

    try {
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'addOwnerNote',
          campaign: this.state.campaign,
          ownerName: owner.name,
          coachName: coachName,
          text: text
        })
      });
      const result = await resp.json();
      if (result.error) throw new Error(result.error);

      // Add to local cache and re-render
      if (!this.state.campaignNotes) this.state.campaignNotes = [];
      this.state.campaignNotes.push({
        noteId: result.noteId,
        campaign: this.state.campaign,
        ownerName: owner.name,
        coachName: coachName,
        text: text,
        timestamp: result.timestamp || new Date().toISOString()
      });
      ta.value = '';
      this.renderHealthTab(owner);
    } catch (err) {
      console.warn('[Notes] Add failed:', err);
    }
    if (btn) { btn.disabled = false; btn.textContent = 'Add'; }
  },

  async _deleteNote(noteId, ownerIdx) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;

    try {
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'deleteOwnerNote',
          noteId: noteId
        })
      });
      const result = await resp.json();
      if (result.error) throw new Error(result.error);

      // Remove from local cache and re-render
      this.state.campaignNotes = (this.state.campaignNotes || []).filter(n => n.noteId !== noteId);
      this.renderHealthTab(owner);
    } catch (err) {
      console.warn('[Notes] Delete failed:', err);
    }
  },

  _relativeTime(isoStr) {
    if (!isoStr) return '';
    const diff = Date.now() - new Date(isoStr).getTime();
    const mins = Math.floor(diff / 60000);
    if (mins < 1) return 'just now';
    if (mins < 60) return mins + 'm ago';
    const hrs = Math.floor(mins / 60);
    if (hrs < 24) return hrs + 'h ago';
    const days = Math.floor(hrs / 24);
    if (days < 7) return days + 'd ago';
    const weeks = Math.floor(days / 7);
    return weeks + 'w ago';
  },

  // ── Production card (colored card, big actual / small goal) ──
  _prodCard(label, actual, goal) {
    const pct = goal ? Math.round((actual / goal) * 100) : 0;
    const goalLine = goal ? `<div class="prod-card-goal">goal ${Number(goal).toLocaleString()}</div>` : '';
    return `
      <div class="prod-card ${this._pctClass(pct)}">
        <div class="prod-card-label">${label}</div>
        <div class="prod-card-actual">${Number(actual).toLocaleString()}</div>
        ${goalLine}
      </div>`;
  },

  // ── Editable production card (for weeks with missing data) ──
  _prodCardEditable(label, actual, goal, ownerIdx) {
    const safeLabel = label.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase();
    return `
      <div class="prod-card pct-none" style="border:2px dashed var(--orange);">
        <div class="prod-card-label">${label}</div>
        <input type="number" class="hc-input prod-card-edit-input" id="prod-edit-${safeLabel}-${ownerIdx}"
          value="${actual || ''}" min="0" placeholder="—"
          style="font-size:28px;font-weight:700;text-align:center;width:80%;margin:4px auto;">
        <div class="prod-card-goal" style="margin-top:4px;">
          goal: <input type="number" class="hc-input prod-card-edit-input" id="prod-edit-goal-${safeLabel}-${ownerIdx}"
            value="${goal || ''}" min="0" placeholder="—"
            style="font-size:13px;width:60px;text-align:center;display:inline-block;">
        </div>
      </div>`;
  },

  // ── Submit production from editable cards ──
  async _submitProductionCards(ownerIdx) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    const prod = owner.production || {};
    const productEntries = prod.products || {};
    const productNames = Object.keys(productEntries);
    const note = document.getElementById('prod-submit-note-' + ownerIdx);

    // Read values from editable card inputs
    const updates = {};
    let totalActual = 0, totalGoal = 0;
    if (productNames.length > 0) {
      for (const pName of productNames) {
        const safeLabel = pName.replace(/[^a-zA-Z0-9]/g, '-').toLowerCase();
        const actualVal = parseInt(document.getElementById('prod-edit-' + safeLabel + '-' + ownerIdx)?.value) || 0;
        const goalVal = parseInt(document.getElementById('prod-edit-goal-' + safeLabel + '-' + ownerIdx)?.value) || 0;
        updates[pName] = { actual: actualVal, goal: goalVal };
        totalActual += actualVal;
        totalGoal += goalVal;
      }
    } else {
      totalActual = parseInt(document.getElementById('prod-edit-total-units-' + ownerIdx)?.value) || 0;
      totalGoal = parseInt(document.getElementById('prod-edit-goal-total-units-' + ownerIdx)?.value) || 0;
    }

    if (!totalActual && !totalGoal) return;

    // Update in-memory state
    prod.totalActual = totalActual;
    prod.totalGoal = totalGoal;
    for (const pName in updates) {
      if (prod.products[pName]) {
        prod.products[pName].actual = updates[pName].actual;
        prod.products[pName].goal = updates[pName].goal;
      }
    }

    // Also update the newest production history entry
    const prodHist = owner.productionHistory || [];
    if (prodHist.length > 0) {
      const newest = prodHist[prodHist.length - 1];
      newest.tA = totalActual;
      newest.tG = totalGoal;
      for (const pName in updates) {
        if (!newest.products) newest.products = {};
        newest.products[pName] = { actual: updates[pName].actual, goal: updates[pName].goal };
      }
    }

    // Save to backend
    try {
      if (note) { note.textContent = 'Saving...'; note.classList.add('show'); }
      const sheetName = owner._sheetName || owner.tab || owner.name;
      const campaignLabel = this._getCampaignLabel();
      // Determine the date for the production row (newest week)
      const prodDate = prodHist.length > 0 ? prodHist[prodHist.length - 1].date : this._latestWeekDate;
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'updateProduction',
          ownerName: sheetName,
          date: prodDate,
          campaignLabel: campaignLabel,
          products: productNames.length > 0 ? updates : null,
          internet: totalActual,
          wireless: 0,
          dtv: 0,
          goals: totalGoal ? String(totalGoal) : ''
        })
      });
      const result = await resp.json();
      if (result.error) throw new Error(result.error);
      console.log('[Prod Cards] Saved:', sheetName, prodDate, updates);
      this._invalidateOdCache();
      if (note) {
        note.textContent = 'Saved ✓';
        setTimeout(() => note.classList.remove('show'), 3000);
      }
      // Re-render health tab to swap cards from editable to display
      this.renderHealthTab(owner);
    } catch (err) {
      console.error('[Prod Cards] Save failed:', err);
      if (note) {
        note.textContent = 'Save failed — ' + err.message;
        note.style.color = '#e53535';
        setTimeout(() => { note.classList.remove('show'); note.style.color = ''; }, 5000);
      }
    }
  },

  // ── Combined production card (main metric + sub metric underneath) ──
  _prodCardCombined(mainLabel, mainActual, mainGoal, subLabel, subActual, isCurrency) {
    const pct = mainGoal ? Math.round((mainActual / mainGoal) * 100) : 0;
    const fmt = (v) => isCurrency ? '$' + Number(v).toLocaleString() : Number(v).toLocaleString();
    const goalLine = mainGoal ? `<div class="prod-card-goal">goal ${fmt(mainGoal)}</div>` : '';
    return `
      <div class="prod-card ${this._pctClass(pct)}">
        <div class="prod-card-label">${mainLabel}</div>
        <div class="prod-card-actual">${fmt(mainActual)}</div>
        ${goalLine}
        <div class="prod-card-sub">
          <span class="prod-card-sub-label">${subLabel}</span>
          <span class="prod-card-sub-value">${isCurrency ? '$' + Number(subActual).toLocaleString() : Number(subActual).toLocaleString()}</span>
        </div>
      </div>`;
  },

  // ── Production percentage → class ──
  _pctClass(pct) {
    if (pct >= 100) return 'pct-green';
    if (pct >= 80) return 'pct-yellow';
    if (pct >= 60) return 'pct-orange';
    return 'pct-red';
  },

  // ══════════════════════════════════════════════════
  // RENDER: Recruiting Tab (spreadsheet-style table)
  // ══════════════════════════════════════════════════

  renderRecruitingTab(owner) {
    // Clear NLR banner (will be re-rendered by cost section if data exists)
    const bannerEl = document.getElementById('recruiting-banner');
    if (bannerEl) bannerEl.innerHTML = '';

    const r = owner.recruiting;
    if (!r || !r.rows || !r.rows.length) {
      const el = document.getElementById('owner-recruiting-table');
      if (el) el.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-text">Recruiting data will populate from Campaign Tracker Section 2.</div>
        </div>`;
      const wowEl = document.getElementById('owner-recruiting-wow');
      if (wowEl) wowEl.innerHTML = '';
    } else {
      // Projected table (4 most recent weeks)
      this._renderRecruitingTable(r, 'owner-recruiting-table');
      // Week-over-week raw data (all weeks)
      this._renderRecruitingWoW(owner);
    }
    // Recruiting costs — lazy load per-owner on first view
    this._loadAndRenderCosts(owner);
  },

  // ── NLR banner (renders above the cost/platform breakdown section) ──
  _renderNlrBanner() {
    const bannerEl = document.getElementById('recruiting-banner');
    if (!bannerEl) return;
    const now = new Date();
    const monthYear = now.toLocaleString('en-US', { month: 'long', year: 'numeric' });
    bannerEl.innerHTML = `
      <div class="recruit-banner">
        <div class="recruit-banner-left">
          <img src="references/logos/nlr-logo-symbol.jpg" alt="NLR" class="recruit-banner-logo">
          <div class="recruit-banner-title">Recruiting Ad Spend Report</div>
        </div>
        <div class="recruit-banner-right">Powered by Next Level Recruiting · ${this._esc(monthYear)}</div>
      </div>`;
  },

  // ── Load and render Indeed Tracking for a single owner ──
  async _loadAndRenderCosts(owner) {
    const el = document.getElementById('owner-indeed-costs');
    if (!el) return;

    // Check if owner is Non-Partner or unmapped for NLR
    if (this._isNonPartner(owner.name, 'nlrWorkbookId') || this._isUnmapped(owner.name, 'nlrWorkbookId')) {
      const bannerEl = document.getElementById('recruiting-banner');
      if (bannerEl) bannerEl.innerHTML = '';
      el.innerHTML = `
        <div class="empty-state" style="padding:40px 20px;">
          <div class="empty-state-icon">🚫</div>
          <div class="empty-state-text" style="font-size:1.1rem;color:#4a7090;">
            Not Utilizing NLR Services
          </div>
        </div>`;
      return;
    }

    // Render NLR banner above cost section
    this._renderNlrBanner();

    // Map NLR data to renderer format and compute WoW deltas
    const _mapNlrWeeks = (nlr) => {
      const weeks = nlr.map(w => ({
        weekOf: w.date || '',
        totalSpend: w.totalSpend || 0,
        totalApplies: w.applies || 0,
        total2nds: w['2nds'] || 0,
        totalNewStarts: w.newStarts || 0,
        cpa: w.cpa || 0,
        cpns: w.cpns || 0,
        numAds: w.numAds || w.ads?.length || 0,
        ads: w.ads || []
      }));
      // Compute week-over-week deltas (mirrors readIndeedTracking server logic)
      for (let i = 0; i < weeks.length; i++) {
        if (i === 0) { weeks[i].delta = null; continue; }
        const prev = weeks[i - 1], curr = weeks[i];
        weeks[i].delta = {
          spend: +(curr.totalSpend - prev.totalSpend).toFixed(2),
          applies: curr.totalApplies - prev.totalApplies,
          seconds: curr.total2nds - prev.total2nds,
          newStarts: curr.totalNewStarts - prev.totalNewStarts,
          cpa: +(curr.cpa - prev.cpa).toFixed(2),
          cpns: +(curr.cpns - prev.cpns).toFixed(2),
          spendPct: prev.totalSpend > 0 ? +((curr.totalSpend - prev.totalSpend) / prev.totalSpend * 100).toFixed(1) : null,
          appliesPct: prev.totalApplies > 0 ? +((curr.totalApplies - prev.totalApplies) / prev.totalApplies * 100).toFixed(1) : null,
          cpaPct: prev.cpa > 0 ? +((curr.cpa - prev.cpa) / prev.cpa * 100).toFixed(1) : null,
          cpnsPct: prev.cpns > 0 ? +((curr.cpns - prev.cpns) / prev.cpns * 100).toFixed(1) : null
        };
      }
      return weeks;
    };

    // Use NLR data already fetched by _fetchOwnerNlrData
    if (owner.nlrData && owner.nlrData.length > 0) {
      owner.indeedTracking = { weeks: _mapNlrWeeks(owner.nlrData) };
      this._renderIndeedTracking(owner);
      return;
    }

    // NLR pre-fetch may still be in progress — wait for it
    el.innerHTML = `<div class="coaching-label">Weekly Ad Spend</div>
      <div class="empty-state"><div class="loading-spinner" style="width:24px;height:24px;border-width:3px;margin:0 auto"></div>
      <div class="empty-state-text" style="margin-top:8px">Loading weekly ad data...</div></div>`;

    for (let i = 0; i < 60; i++) {
      await new Promise(r => setTimeout(r, 500));
      if (owner.nlrData) break;
    }

    if (owner.nlrData && owner.nlrData.length > 0) {
      owner.indeedTracking = { weeks: _mapNlrWeeks(owner.nlrData) };
      if (this.state.selectedOwner === owner && this.state.currentTab === 'recruiting') {
        this._renderIndeedTracking(owner);
      }
      return;
    }

    if (this.state.selectedOwner === owner && this.state.currentTab === 'recruiting') {
      el.innerHTML = `<div class="coaching-label">Weekly Ad Spend</div>
        <div class="empty-state"><div class="empty-state-text">No ad spend data available.</div></div>`;
    }
  },

  // ── Week-over-week recruiting: flip card (charts front, table back) ──
  _renderRecruitingWoW(owner) {
    const el = document.getElementById('owner-recruiting-wow');
    if (!el) return;

    const full = owner.recruitingFull || owner.recruiting;
    if (!full || !full.weeks || !full.weeks.length) {
      el.innerHTML = '';
      return;
    }

    const weeks = full.weeks;
    const rows = full.rows || [];
    const labels = this.RECRUITING_LABELS;

    // Filter to weeks that have at least one non-zero value
    const activeWeekIdxs = [];
    for (let wi = 0; wi < weeks.length; wi++) {
      const hasData = labels.some((_, ri) => (rows[ri]?.values?.[wi] ?? 0) !== 0);
      if (hasData) activeWeekIdxs.push(wi);
    }

    if (!activeWeekIdxs.length) {
      el.innerHTML = '';
      return;
    }

    // Reversed (newest-first) active week indices
    const revIdxs = [...activeWeekIdxs].reverse();
    const n = revIdxs.length;

    const shortDate = (d) => {
      if (!d) return '';
      const parts = String(d).split(/[-\/]/);
      return parts.length >= 2 ? parts[0] + '/' + parts[1] : d;
    };

    // Extract series data (newest-first)
    const series = (ri) => revIdxs.map(wi => rows[ri]?.values?.[wi] ?? 0);

    const newStartsShowed = series(10);
    const newStartsBooked = series(9);
    const r2Booked = series(6);
    const r2Showed = series(7);
    const r1Booked = series(2);
    const r1Showed = series(3);
    const weekLabels = revIdxs.map(wi => shortDate(weeks[wi]));

    // ── Build 3 retention % trend cards ──
    // Pass row indices + projected values for color matching with recruiting table
    const projNS = rows[10]?.projected ?? 0;
    const projR2 = rows[7]?.projected ?? 0;
    const projR1 = rows[3]?.projected ?? 0;
    const chart1 = this._buildRetentionCard('New Starts', weekLabels, newStartsBooked, newStartsShowed, 'Booked new starts that actually show up', 11, 10, projNS);
    const chart2 = this._buildRetentionCard('2nd Round Interviews', weekLabels, r2Booked, r2Showed, '1st round interview quality', 8, 7, projR2);
    const chart3 = this._buildRetentionCard('1st Round Interviews', weekLabels, r1Booked, r1Showed, 'Recruiting hub / phone booker performance', 4, 3, projR1);

    // ── Build table (back side — existing stacked cards) ──
    let tableHtml = '<div class="wow-cards">';
    [...activeWeekIdxs].reverse().forEach(wi => {
      const prevWi = activeWeekIdxs[activeWeekIdxs.indexOf(wi) - 1] ?? null;
      tableHtml += `<div class="wow-card">
        <div class="wow-card-header">${this._esc(weeks[wi])}</div>
        <div class="data-table-wrap"><table class="data-table"><tbody>`;
      labels.forEach((def, ri) => {
        const val = rows[ri]?.values?.[wi] ?? 0;
        const prev = prevWi !== null ? (rows[ri]?.values?.[prevWi] ?? null) : null;
        const display = def.isRate ? val + '%' : val;
        const arrow = prev !== null ? this._trendArrow(val, prev) : '';
        tableHtml += `<tr><td>${this._esc(def.label)}</td><td class="num">${display} ${arrow}</td></tr>`;
      });
      tableHtml += `</tbody></table></div></div>`;
    });
    tableHtml += '</div>';

    // ── Assemble flip card ──
    el.innerHTML = `
      <div class="coaching-label">
        Week-over-Week Recruiting
        <button class="flip-btn" onclick="NationalApp._flipRecruitCard()" title="Flip to ${this._recruitFlipped ? 'charts' : 'table'} view">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
        </button>
      </div>
      <div class="flip-card${this._recruitFlipped ? ' flipped' : ''}" id="recruit-flip-card">
        <div class="flip-card-inner">
          <div class="flip-card-front">
            <div class="recruit-charts-grid">
              ${chart1}${chart2}${chart3}
            </div>
          </div>
          <div class="flip-card-back">
            ${tableHtml}
          </div>
        </div>
      </div>`;
  },

  _recruitFlipped: false,

  _flipRecruitCard() {
    this._recruitFlipped = !this._recruitFlipped;
    const card = document.getElementById('recruit-flip-card');
    if (card) card.classList.toggle('flipped', this._recruitFlipped);
    const btn = card?.parentElement?.querySelector('.flip-btn');
    if (btn) btn.title = 'Flip to ' + (this._recruitFlipped ? 'charts' : 'table') + ' view';
  },

  // ── Build a single retention bar card (showed vs booked bars + hero number) ──
  // Map cell-* class to hex color for SVG
  _cellColorHex(cls) {
    if (cls === 'cell-blue') return '#22d3ee';
    if (cls === 'cell-green') return '#84cc16';
    if (cls === 'cell-yellow') return '#ffeb3b';
    if (cls === 'cell-orange') return '#f9a825';
    if (cls === 'cell-red') return '#e53935';
    return '#8a95a5';
  },

  _showRcTooltip(event) {
    const el = event.target;
    const card = el.closest('.recruit-chart-card');
    if (!card) return;
    const tt = card.querySelector('.rc-tooltip');
    if (!tt) return;
    const wk = el.getAttribute('data-wk') || '';
    const bk = el.getAttribute('data-bk') || '0';
    const sh = el.getAttribute('data-sh') || '0';
    const pct = el.getAttribute('data-pct');
    tt.innerHTML = `<div style="font-weight:700;margin-bottom:2px">${wk}</div><div>Booked: <strong>${bk}</strong></div><div>Showed: <strong>${sh}</strong></div>${pct ? '<div style="border-top:1px solid rgba(255,255,255,0.2);margin-top:3px;padding-top:3px">Retention: <strong>' + pct + '%</strong></div>' : ''}`;
    tt.classList.add('visible');
    const rect = card.getBoundingClientRect();
    const bar = el.getBoundingClientRect();
    let left = bar.left - rect.left + bar.width / 2 - tt.offsetWidth / 2;
    left = Math.max(4, Math.min(left, rect.width - tt.offsetWidth - 4));
    tt.style.left = left + 'px';
    tt.style.top = (bar.top - rect.top - tt.offsetHeight - 6) + 'px';
  },

  _hideRcTooltip() {
    document.querySelectorAll('.rc-tooltip.visible').forEach(t => t.classList.remove('visible'));
  },

  _buildRetentionCard(title, weekLabels, booked, showed, subtitle, retentionRowIdx, showedRowIdx, projected) {
    const n = weekLabels.length;
    const VISIBLE = 4;

    // Current week hero values
    const curShowed = showed[0] || 0;
    const curBooked = booked[0] || 0;
    const curPct = curBooked > 0 ? Math.round((curShowed / curBooked) * 100) : null;

    // Hero number color — use showed value vs projected (same as table's ratio thresholds)
    const heroColorCls = this._cellColor(curShowed, projected, false, showedRowIdx);
    const heroColor = this._cellColorHex(heroColorCls);

    // SVG dimensions — viewBox scales to fill card, scroll if > VISIBLE weeks
    const svgH = 130;
    const PAD_T = 22, PAD_B = 18;
    const plotH = svgH - PAD_T - PAD_B;
    const baseY = PAD_T + plotH;
    const BAR_R = 3;
    const SLOT_W = 46;
    const vbW = n * SLOT_W + 6;
    const svgWidthPct = n <= VISIBLE ? '100%' : `${(n / VISIBLE) * 100}%`;

    // Y-axis max
    let maxVal = 1;
    for (let i = 0; i < n; i++) {
      if (booked[i] > maxVal) maxVal = booked[i];
      if (showed[i] > maxVal) maxVal = showed[i];
    }
    const step = maxVal > 40 ? 10 : maxVal > 15 ? 5 : maxVal > 5 ? 2 : 1;
    const yMax = Math.ceil(maxVal / step) * step || step;
    const yScale = plotH / yMax;

    // Bar sizing
    const GAP_FRAC = 0.22;
    const barW = SLOT_W * (1 - GAP_FRAC);
    const barOff = (SLOT_W - barW) / 2;

    const roundTop = (x, y, w, h, r) => {
      if (h <= 0) return '';
      const cr = Math.min(r, h / 2, w / 2);
      return `M${x},${y+h}L${x},${y+cr}Q${x},${y} ${x+cr},${y}L${x+w-cr},${y}Q${x+w},${y} ${x+w},${y+cr}L${x+w},${y+h}Z`;
    };

    let svg = '';

    // Shadow filter for yellow text readability
    svg += `<defs><filter id="txtShadow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="0" stdDeviation="1.5" flood-color="#000" flood-opacity="0.4"/>
    </filter></defs>`;
    const _isLight = (c) => c === '#ffeb3b' || c === '#f9a825' || c === '#84cc16';

    // Light gridlines
    for (let v = 0; v <= yMax; v += step) {
      const y = baseY - v * yScale;
      svg += `<line x1="0" y1="${y}" x2="${vbW}" y2="${y}" stroke="#e8ecf1" stroke-width="0.5"/>`;
    }

    // Bars per week — color based on retention % using table thresholds
    for (let i = 0; i < n; i++) {
      const bk = booked[i] || 0;
      const sh = showed[i] || 0;
      const pct = bk > 0 ? Math.round((sh / bk) * 100) : null;

      // Bar color from retention thresholds (same as table)
      const barCls = pct !== null ? this._cellColor(pct, null, true, retentionRowIdx) : '';
      const barColor = this._cellColorHex(barCls);

      const gx = i * SLOT_W + barOff;
      const cx = i * SLOT_W + SLOT_W / 2;

      // Booked bar (ghost — the potential)
      const bkH = bk * yScale;
      if (bkH > 0) {
        svg += `<path d="${roundTop(gx, baseY - bkH, barW, bkH, BAR_R)}" fill="${barColor}" opacity="0.13"/>`;
        svg += `<path d="${roundTop(gx, baseY - bkH, barW, bkH, BAR_R)}" fill="none" stroke="${barColor}" stroke-width="1" opacity="0.35" stroke-dasharray="3,2"/>`;
      }

      // Showed bar (solid — actual output)
      const shH = sh * yScale;
      if (shH > 0) {
        svg += `<path d="${roundTop(gx, baseY - shH, barW, shH, BAR_R)}" fill="${barColor}" opacity="0.88"/>`;
      }

      // Showed count inside solid bar (if tall enough)
      if (sh > 0 && shH > 14) {
        const _shadow = _isLight(barColor) ? ' filter="url(#txtShadow)"' : '';
        svg += `<text x="${cx}" y="${baseY - shH / 2 + 4}" text-anchor="middle" fill="#fff" font-size="9" font-weight="700" font-family="Inter,sans-serif"${_shadow}>${sh}</text>`;
      }

      // Hover target over entire bar area
      const hoverTop = Math.min(bkH > 0 ? baseY - bkH : baseY, shH > 0 ? baseY - shH : baseY) - 16;
      const hoverH = baseY - hoverTop;
      svg += `<rect x="${gx}" y="${hoverTop}" width="${barW}" height="${Math.max(hoverH, 4)}" fill="transparent" style="cursor:pointer" data-wk="${weekLabels[i]}" data-bk="${bk}" data-sh="${sh}"${pct !== null ? ` data-pct="${pct}"` : ''} onmouseenter="NationalApp._showRcTooltip(event)" onmouseleave="NationalApp._hideRcTooltip()"/>`;

      // Retention % pill above bar
      if (pct !== null) {
        const topY = Math.min(baseY - bkH, baseY - shH);
        const pillY = topY - 14;
        const pillW = pct >= 100 ? 30 : 26;
        const pillShadow = _isLight(barColor) ? ' filter="url(#txtShadow)"' : '';
        svg += `<rect x="${cx - pillW/2}" y="${pillY}" width="${pillW}" height="12" rx="6" fill="${barColor}" opacity="0.9"/>`;
        svg += `<text x="${cx}" y="${pillY + 9}" text-anchor="middle" fill="#fff" font-size="7.5" font-weight="800" font-family="Inter,sans-serif"${pillShadow}>${pct}%</text>`;
      }

      // Week label below
      svg += `<text x="${cx}" y="${svgH - 3}" text-anchor="middle" fill="#8a95a5" font-size="7" font-weight="600" font-family="Inter,sans-serif">${weekLabels[i]}</text>`;
    }

    // Legend
    const legend = `<div class="hc-chart-legend" style="margin-top:4px">
      <span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch" style="background:#8a95a5"></span>Showed</span>
      <span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch" style="background:#8a95a5;opacity:0.18;border:1.5px dashed #8a95a5"></span>Booked</span>
      <span class="hc-chart-legend-item" style="font-size:10px;color:#8a95a5">Colors = retention %</span>
    </div>`;

    return `
      <div class="recruit-chart-card">
        <div class="rc-card-header">
          <div>
            <div class="recruit-chart-title">${title}</div>
            <div class="rc-card-subtitle">${subtitle}</div>
          </div>
          <div class="rc-card-hero">
            <div class="rc-card-hero-num" style="color:${heroColor}">${curShowed}</div>
            <div class="rc-card-hero-label">showed</div>
          </div>
        </div>
        <div class="rc-bar-wrap">
          <svg viewBox="0 0 ${vbW} ${svgH}" width="${svgWidthPct}" preserveAspectRatio="xMinYMid meet" style="display:block">${svg}</svg>
        </div>
        <div class="hc-chart-tooltip rc-tooltip"></div>
        ${legend}
      </div>`;
  },



  _fmtDollar(v) {
    if (!v && v !== 0) return '—';
    if (Number(v) === 0) return '—';
    return '$' + Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  },

  _fmtNum(v) {
    if (!v && v !== 0) return '—';
    if (Number(v) === 0) return '—';
    return Number(v).toLocaleString('en-US');
  },

  // Trend arrow helper (lower = good for cost rows, higher = good for count rows)
  _costTrendArrow(current, previous, lowerIsBetter) {
    if (previous === null || previous === undefined) return '';
    const diff = current - previous;
    if (Math.abs(diff) < 0.01) return '<span class="trend-flat">—</span>';
    if (lowerIsBetter) {
      return diff < 0
        ? '<span class="trend-up">&#9650;</span>'
        : '<span class="trend-down">&#9660;</span>';
    }
    return diff > 0
      ? '<span class="trend-up">&#9650;</span>'
      : '<span class="trend-down">&#9660;</span>';
  },

  // ══════════════════════════════════════════════════
  // RENDER: Weekly Indeed Tracking (ad spend view)
  // ══════════════════════════════════════════════════

  _renderIndeedTracking(owner) {
    const el = document.getElementById('owner-indeed-costs');
    if (!el) return;

    const data = owner.indeedTracking;
    if (!data || !data.weeks || !data.weeks.length) {
      el.innerHTML = '';
      return;
    }

    const weeks = data.weeks; // oldest-first
    // Default to prior week (current week is still in progress)
    const defaultIdx = weeks.length > 1 ? weeks.length - 2 : weeks.length - 1;
    const defaultWeek = weeks[defaultIdx];
    const defaultPrev = defaultIdx > 0 ? weeks[defaultIdx - 1] : null;

    let html = '';

    // ── Part 1: Weekly Totals KPI Cards (show prior week) ──
    html += `<div class="coaching-label">Weekly Ad Spend Overview</div>`;
    html += `<div class="it-kpi-row">`;
    html += this._itKpiCard('Total Spend', this._fmtDollar(defaultWeek.totalSpend),
      defaultWeek.delta ? defaultWeek.delta.spendPct : null, true);
    html += this._itKpiCard('Applies', this._fmtNum(defaultWeek.totalApplies),
      defaultWeek.delta ? defaultWeek.delta.appliesPct : null, false);
    html += this._itKpiCard('2nds', this._fmtNum(defaultWeek.total2nds), null, false);
    html += this._itKpiCard('New Starts', this._fmtNum(defaultWeek.totalNewStarts), null, false);
    html += this._itKpiCard('CPA', this._fmtDollar(defaultWeek.cpa),
      defaultWeek.delta ? defaultWeek.delta.cpaPct : null, true);
    html += this._itKpiCard('CPNS', this._fmtDollar(defaultWeek.cpns),
      defaultWeek.delta ? defaultWeek.delta.cpnsPct : null, true);
    html += `</div>`;
    html += `<div class="it-kpi-week-label">Week of ${this._esc(defaultWeek.weekOf)} · ${defaultWeek.numAds} ads</div>`;

    // ── Part 2: Week-over-Week Trend Table ──
    html += `<div class="coaching-label it-section-label">Week-over-Week Trends</div>`;
    html += this._buildTrackingTrend(weeks);

    // ── Part 3: Ad Breakdown (defaults to prior week) ──
    html += `<div class="coaching-label it-section-label">Ad Breakdown
      <select class="it-week-select" onchange="NationalApp._switchTrackingWeek(this.value)">
        ${weeks.map((w, i) => `<option value="${i}"${i === defaultIdx ? ' selected' : ''}>${this._esc(w.weekOf)}</option>`).join('')}
      </select>
    </div>`;
    html += `<div id="it-ad-breakdown">`;
    html += this._buildAdBreakdown(defaultWeek, defaultPrev);
    html += `</div>`;

    el.innerHTML = html;
  },

  // ── KPI card for weekly totals ──
  _itKpiCard(label, value, pctChange, lowerIsBetter) {
    let arrow = '';
    if (pctChange !== null && pctChange !== undefined) {
      const abs = Math.abs(pctChange);
      if (abs < 0.5) {
        arrow = `<span class="it-kpi-delta it-flat">—</span>`;
      } else {
        const isGood = lowerIsBetter ? (pctChange < 0) : (pctChange > 0);
        const cls = isGood ? 'it-good' : 'it-bad';
        const sym = pctChange > 0 ? '▲' : '▼';
        arrow = `<span class="it-kpi-delta ${cls}">${sym} ${abs.toFixed(1)}%</span>`;
      }
    }
    return `<div class="it-kpi-card">
      <div class="it-kpi-value">${value}${arrow}</div>
      <div class="it-kpi-label">${label}</div>
    </div>`;
  },

  // ── WoW trend table: weeks as columns, key metrics as rows ──
  _buildTrackingTrend(weeks) {
    const TREND_ROWS = [
      { label: 'Total Spend',  key: 'totalSpend',     fmt: 'dollar', lowerBetter: true },
      { label: '# Applies',    key: 'totalApplies',   fmt: 'num',    lowerBetter: false },
      { label: '# 2nds',       key: 'total2nds',      fmt: 'num',    lowerBetter: false },
      { label: '# New Starts', key: 'totalNewStarts', fmt: 'num',    lowerBetter: false },
      { label: 'CPA',          key: 'cpa',            fmt: 'dollar', lowerBetter: true },
      { label: 'CPNS',         key: 'cpns',           fmt: 'dollar', lowerBetter: true },
      { label: '# Ads',        key: 'numAds',         fmt: 'num',    lowerBetter: false }
    ];

    // Show last 6 weeks max, newest on left
    const shown = weeks.slice(-6).reverse();

    let h = `<div class="it-trend-wrap"><div class="data-table-wrap"><table class="data-table it-trend-table">
      <thead><tr><th></th>
        ${shown.map(w => `<th class="num">${this._esc(w.weekOf)}</th>`).join('')}
      </tr></thead><tbody>`;

    for (const r of TREND_ROWS) {
      const fmt = r.fmt === 'dollar' ? this._fmtDollar : this._fmtNum;
      h += `<tr><td class="rc-label">${r.label}</td>`;
      for (let i = 0; i < shown.length; i++) {
        const val = shown[i][r.key] ?? 0;
        // Compare against next column (older week) since order is reversed
        const older = i < shown.length - 1 ? (shown[i + 1][r.key] ?? null) : null;
        const arrow = older !== null ? this._costTrendArrow(val, older, r.lowerBetter) : '';
        h += `<td class="num">${fmt(val)} ${arrow}</td>`;
      }
      h += `</tr>`;
    }

    h += `</tbody></table></div></div>`;
    return h;
  },

  // ── Ad breakdown table for a single week ──
  _buildAdBreakdown(week, prevWeek) {
    if (!week || !week.ads || !week.ads.length) {
      return `<div class="empty-state"><div class="empty-state-text">No ads this week</div></div>`;
    }

    // Group ads by account
    const byAccount = {};
    const accountOrder = [];
    for (const ad of week.ads) {
      const acct = ad.account || 'Unknown';
      if (!byAccount[acct]) {
        byAccount[acct] = [];
        accountOrder.push(acct);
      }
      byAccount[acct].push(ad);
    }

    // Build prev-week lookup by adTitle for effectiveness comparison
    const prevByTitle = {};
    if (prevWeek && prevWeek.ads) {
      for (const ad of prevWeek.ads) {
        if (ad.adTitle) prevByTitle[ad.adTitle] = ad;
      }
    }

    let h = `<div class="it-breakdown-wrap"><div class="data-table-wrap"><table class="data-table it-breakdown-table">
      <thead><tr>
        <th>Indeed Account</th>
        <th>Ad Title</th>
        <th>Location</th>
        <th class="num">Spend</th>
        <th class="num">Applies</th>
        <th class="num">2nds</th>
        <th class="num">New Starts</th>
        <th class="num">CPA</th>
        <th class="num">CPNS</th>
        <th>Plan</th>
      </tr></thead><tbody>`;

    for (const acct of accountOrder) {
      const ads = byAccount[acct];
      // Account group header
      const acctSpend = ads.reduce((s, a) => s + a.spend, 0);
      const acctApplies = ads.reduce((s, a) => s + a.applies, 0);
      const acct2nds = ads.reduce((s, a) => s + a.seconds, 0);
      const acctNS = ads.reduce((s, a) => s + a.newStarts, 0);
      const acctCPA = acctApplies > 0 ? acctSpend / acctApplies : 0;
      const acctCPNS = acctNS > 0 ? acctSpend / acctNS : 0;

      h += `<tr class="it-acct-row">
        <td colspan="3"><strong>${this._esc(acct)}</strong> <span class="it-ad-count">${ads.length} ad${ads.length !== 1 ? 's' : ''}</span></td>
        <td class="num"><strong>${this._fmtDollar(acctSpend)}</strong></td>
        <td class="num"><strong>${this._fmtNum(acctApplies)}</strong></td>
        <td class="num"><strong>${this._fmtNum(acct2nds)}</strong></td>
        <td class="num"><strong>${this._fmtNum(acctNS)}</strong></td>
        <td class="num"><strong>${this._fmtDollar(acctCPA)}</strong></td>
        <td class="num"><strong>${this._fmtDollar(acctCPNS)}</strong></td>
        <td></td>
      </tr>`;

      for (const ad of ads) {
        // WoW effectiveness arrow for CPA
        const prevAd = prevByTitle[ad.adTitle];
        const cpaArrow = prevAd && prevAd.cpa > 0 && ad.cpa > 0
          ? this._costTrendArrow(ad.cpa, prevAd.cpa, true) : '';
        const cpnsArrow = prevAd && prevAd.cpns > 0 && ad.cpns > 0
          ? this._costTrendArrow(ad.cpns, prevAd.cpns, true) : '';

        // Plan column styling
        const planLc = (ad.plan || '').toLowerCase();
        const planCls = planLc === 'skip' ? 'it-plan-skip'
          : planLc === 'repost' || planLc === 'repost' ? 'it-plan-repost'
          : planLc.includes('paused') ? 'it-plan-paused'
          : '';

        h += `<tr class="it-ad-row">
          <td class="it-indent"></td>
          <td class="it-ad-title">${this._esc(ad.adTitle)}</td>
          <td>${this._esc(ad.location)}</td>
          <td class="num">${this._fmtDollar(ad.spend)}</td>
          <td class="num">${this._fmtNum(ad.applies)}</td>
          <td class="num">${this._fmtNum(ad.seconds)}</td>
          <td class="num">${this._fmtNum(ad.newStarts)}</td>
          <td class="num">${this._fmtDollar(ad.cpa)} ${cpaArrow}</td>
          <td class="num">${this._fmtDollar(ad.cpns)} ${cpnsArrow}</td>
          <td class="${planCls}">${this._esc(ad.plan)}</td>
        </tr>`;
      }
    }

    // Total row
    h += `<tr class="it-total-row">
      <td colspan="3"><strong>TOTAL</strong></td>
      <td class="num"><strong>${this._fmtDollar(week.totalSpend)}</strong></td>
      <td class="num"><strong>${this._fmtNum(week.totalApplies)}</strong></td>
      <td class="num"><strong>${this._fmtNum(week.total2nds)}</strong></td>
      <td class="num"><strong>${this._fmtNum(week.totalNewStarts)}</strong></td>
      <td class="num"><strong>${this._fmtDollar(week.cpa)}</strong></td>
      <td class="num"><strong>${this._fmtDollar(week.cpns)}</strong></td>
      <td></td>
    </tr>`;

    h += `</tbody></table></div></div>`;
    return h;
  },

  // ── Switch ad breakdown week via dropdown ──
  _switchTrackingWeek(idx) {
    const owner = this.state.selectedOwner;
    if (!owner || !owner.indeedTracking) return;
    const weeks = owner.indeedTracking.weeks;
    const i = parseInt(idx);
    const week = weeks[i];
    const prev = i > 0 ? weeks[i - 1] : null;
    if (!week) return;

    const wrap = document.getElementById('it-ad-breakdown');
    if (wrap) {
      wrap.innerHTML = this._buildAdBreakdown(week, prev);
    }
  },

  // ══════════════════════════════════════════════════
  // RENDER: Sales Tab
  // ══════════════════════════════════════════════════

  _pct(v) {
    if (v === null || v === undefined || v === '' || v === 0) return '—';
    // Already a decimal like 0.42 → 42%
    if (typeof v === 'number' && v > 0 && v <= 1) return Math.round(v * 100) + '%';
    if (typeof v === 'number' && v > 1) return Math.round(v) + '%';
    return String(v);
  },

  renderSalesTab(owner) {
    // Sales data comes from Tableau — no Non-Partner gate needed here

    // Show loading if sales data is still being fetched
    if (owner._ndsSalesFetching || owner._resSalesFetching) {
      const summaryEl = document.getElementById('sales-summary');
      const repsEl = document.getElementById('sales-reps-table');
      if (summaryEl) summaryEl.innerHTML = '<div class="coaching-section" style="text-align:center;padding:40px;"><div class="loading-spinner"></div><div style="margin-top:12px;font-size:13px;color:#708090;">Loading sales data...</div></div>';
      if (repsEl) repsEl.innerHTML = '';
      return;
    }

    const s = owner.sales;
    const sm = s.summary;
    const isNDS = this.state.campaign && (this.state.campaign.indexOf('nds') >= 0 || this.state.campaign.indexOf('NDS') >= 0);
    const isRes = this.state.campaign === 'att-res';
    // Store reps for checkbox flag/unflag reference
    this._currentSalesReps = s.reps || [];

    // ── Card 1: Owner Summary ──
    const summaryEl = document.getElementById('sales-summary');
    if (sm) {
      const kpis = isNDS
        ? [
            { label: 'New/Ports LW', value: sm.newPorts || sm.totalVolume, cls: 'big' },
            { label: 'Rep Count', value: sm.repCount },
            { label: 'Order Count', value: sm.orderCount }
          ]
        : isRes
        ? [
            { label: 'Total Volume', value: sm.totalVolume ?? '—', cls: 'big' },
            { label: 'Rep Count', value: sm.repCount ?? '—' }
          ]
        : [
            { label: 'Total Volume', value: sm.totalVolume ?? '—', cls: 'big' },
            { label: 'Rep Count', value: sm.repCount ?? '—' },
            { label: 'Sales / Rep', value: sm.salesPerRep ?? '—' },
            { label: 'Order Count', value: sm.orderCount ?? '—' }
          ];

      const metricsHtml = isNDS
        ? `<div class="sales-metrics-grid">
            <div class="sales-metric-group">
              <div class="sales-metric-group-label">Quality Metrics</div>
              <div class="sales-metric-row"><span>Cancel/Fraud Review</span><span class="num">${this._pct(sm.cancelFraudPct)}</span></div>
              <div class="sales-metric-row"><span>Extra & Premium</span><span class="num">${this._pct(sm.extraPremiumPct)}</span></div>
              <div class="sales-metric-row"><span>Next Up</span><span class="num">${this._pct(sm.nextUpPct)}</span></div>
              <div class="sales-metric-row"><span>Insurance</span><span class="num">${this._pct(sm.insurancePct)}</span></div>
            </div>
            <div class="sales-metric-group">
              <div class="sales-metric-group-label">Performance %</div>
              <div class="sales-metric-row"><span>ABP</span><span class="num">${this._pct(sm.abpPct)}</span></div>
              <div class="sales-metric-row"><span>BYOD</span><span class="num">${this._pct(sm.byodPct)}</span></div>
              <div class="sales-metric-row"><span>New % of New/Ports</span><span class="num">${this._pct(sm.newOfNewPortsPct)}</span></div>
              <div class="sales-metric-row"><span>High/Med Credit</span><span class="num">${this._pct(sm.highMedCreditPct)}</span></div>
            </div>
            <div class="sales-metric-group">
              <div class="sales-metric-group-label">Order Timing</div>
              <div class="sales-metric-row"><span>Away from Doors</span><span class="num">${this._pct(sm.awayFromDoorsPct)}</span></div>
              <div class="sales-metric-row"><span>Before 3:00 PM</span><span class="num">${this._pct(sm.before3pmPct)}</span></div>
              <div class="sales-metric-row"><span>After 7:30 PM</span><span class="num">${this._pct(sm.after730pmPct)}</span></div>
            </div>
          </div>`
        : isRes
          ? `<div class="sales-metrics-grid">
            <div class="sales-metric-group">
              <div class="sales-metric-group-label">Sales Breakdown</div>
              <div class="sales-metric-row"><span>New Internet</span><span class="num">${sm.newInternet ?? '—'}</span></div>
              <div class="sales-metric-row"><span>Upgrade Internet</span><span class="num">${sm.upgradeInternet ?? '—'}</span></div>
              <div class="sales-metric-row"><span>Video</span><span class="num">${sm.videoSales ?? '—'}</span></div>
              <div class="sales-metric-row"><span>Wireless</span><span class="num">${sm.wirelessSales ?? '—'}</span></div>
              <div class="sales-metric-row"><span>Voice</span><span class="num">${sm.voiceSales ?? '—'}</span></div>
            </div>
            <div class="sales-metric-group">
              <div class="sales-metric-group-label">Performance %</div>
              <div class="sales-metric-row"><span>ABP Mix</span><span class="num">${this._pct(sm.abpMix)}</span></div>
              <div class="sales-metric-row"><span>1Gig+ Mix</span><span class="num">${this._pct(sm.gigMix)}</span></div>
              <div class="sales-metric-row"><span>Tech Install</span><span class="num">${this._pct(sm.techInstall)}</span></div>
            </div>
          </div>`
          : `<div class="sales-metrics-grid">
            <div class="sales-metric-group">
              <div class="sales-metric-group-label">Sales Breakdown</div>
              <div class="sales-metric-row"><span>Internet</span><span class="num">${sm.internet ?? '—'}</span></div>
              <div class="sales-metric-row"><span>VOIP</span><span class="num">${sm.voip ?? '—'}</span></div>
              <div class="sales-metric-row"><span>Wireless</span><span class="num">${sm.wireless ?? '—'}</span></div>
              <div class="sales-metric-row"><span>AIR/AWB</span><span class="num">${sm.airAwb ?? '—'}</span></div>
            </div>
            <div class="sales-metric-group">
              <div class="sales-metric-group-label">Order Timing</div>
              <div class="sales-metric-row"><span>Before 12 PM</span><span class="num">${sm.ordersBefore ?? '—'} <small>(${this._pct(sm.earlyPct)})</small></span></div>
              <div class="sales-metric-row"><span>After 5 PM</span><span class="num">${sm.ordersAfter ?? '—'} <small>(${this._pct(sm.latePct)})</small></span></div>
            </div>
            <div class="sales-metric-group">
              <div class="sales-metric-group-label">Performance %</div>
              <div class="sales-metric-row"><span>Weekend Selling</span><span class="num">${this._pct(sm.weekendPct)}</span></div>
              <div class="sales-metric-row"><span>Rep Tier Attainment</span><span class="num">${this._pct(sm.tierPct)}</span></div>
              <div class="sales-metric-row"><span>ABP</span><span class="num">${this._pct(sm.abpPct)}</span></div>
              <div class="sales-metric-row"><span>CRU</span><span class="num">${this._pct(sm.cruPct)}</span></div>
              <div class="sales-metric-row"><span>New Wireless</span><span class="num">${this._pct(sm.newWrlsPct)}</span></div>
              <div class="sales-metric-row"><span>BYOD</span><span class="num">${this._pct(sm.byodPct)}</span></div>
            </div>
          </div>`;

      summaryEl.innerHTML = `
        <div class="coaching-section">
          <div class="coaching-label">Owner Overview</div>
          <div class="sales-kpi-grid">
            ${kpis.map(k => `
              <div class="health-kpi${k.cls ? ' ' + k.cls : ''}">
                <div class="health-kpi-value">${k.value}</div>
                <div class="health-kpi-label">${k.label}</div>
              </div>
            `).join('')}
          </div>
          ${metricsHtml}
        </div>`;
    } else {
      summaryEl.innerHTML = `
        <div class="coaching-section">
          <div class="coaching-label">Owner Overview</div>
          <div class="empty-state">
            <div class="empty-state-text">No production data yet. Click "Import Recruiting" to pull sales data.</div>
          </div>
        </div>`;
    }

    // ── Card 2: Rep Breakdown Table ──
    const repsEl = document.getElementById('sales-reps-table');
    if (s.reps.length) {
      const _stickyTh0 = 'position:sticky;left:0;z-index:2;background:var(--card-bg,#ddeaf5);width:30px;';
      const _stickyTh1 = 'position:sticky;left:30px;z-index:2;background:var(--card-bg,#ddeaf5);cursor:pointer;user-select:none;';
      const _sh = (label, col) => `<th class="num sortable-th" onclick="NationalApp._sortSalesReps('${col}')" style="cursor:pointer;user-select:none;">${label} <span style="font-size:10px;opacity:0.5;">&#x25B2;&#x25BC;</span></th>`;
      const repHeaders = isNDS
        ? `<th style="${_stickyTh0}"></th><th class="sortable-th" onclick="NationalApp._sortSalesReps('name')" style="${_stickyTh1}">Rep Name <span style="font-size:10px;opacity:0.5;">&#x25B2;&#x25BC;</span></th>
           ${_sh('New/Ports','newPorts')}
           ${_sh('Orders','orderCount')}
           ${_sh('Cancel %','cancelFraudPct')}
           ${_sh('Extra %','extraPremiumPct')}
           ${_sh('Next Up %','nextUpPct')}
           ${_sh('ABP %','abpPct')}
           ${_sh('BYOD %','byodPct')}
           ${_sh('New %','newOfNewPortsPct')}
           ${_sh('Insurance %','insurancePct')}
           ${_sh('Credit %','highMedCreditPct')}
           ${_sh('Away %','awayFromDoorsPct')}
           ${_sh('Before 3 %','before3pmPct')}
           ${_sh('After 7:30 %','after730pmPct')}`
        : isRes
        ? `<th style="${_stickyTh0}"></th><th class="sortable-th" onclick="NationalApp._sortSalesReps('name')" style="${_stickyTh1}">Rep Name <span style="font-size:10px;opacity:0.5;">&#x25B2;&#x25BC;</span></th>
           ${_sh('Volume','totalVolume')}
           ${_sh('New Internet','newInternet')}
           ${_sh('Upgrade Internet','upgradeInternet')}
           ${_sh('Video','videoSales')}
           ${_sh('Wireless','wirelessSales')}
           ${_sh('Voice','voiceSales')}
           ${_sh('ABP Mix','abpMix')}
           ${_sh('1Gig+ Mix','gigMix')}
           ${_sh('Tech Install','techInstall')}`
        : `<th style="${_stickyTh0}"></th><th class="sortable-th" onclick="NationalApp._sortSalesReps('name')" style="${_stickyTh1}">Rep Name <span style="font-size:10px;opacity:0.5;">&#x25B2;&#x25BC;</span></th>
           ${_sh('Volume','totalVolume')}
           ${_sh('Orders','orderCount')}
           ${_sh('Sales/Rep','salesPerRep')}
           ${_sh('Internet','internet')}
           ${_sh('VOIP','voip')}
           ${_sh('Wireless','wireless')}
           ${_sh('AIR/AWB','airAwb')}
           ${_sh('Early %','earlyPct')}
           ${_sh('Late %','latePct')}
           ${_sh('ABP %','abpPct')}
           ${_sh('CRU %','cruPct')}
           ${_sh('New Wrls %','newWrlsPct')}
           ${_sh('BYOD %','byodPct')}`;

      const _stickyTd0 = 'position:sticky;left:0;z-index:1;background:inherit;';
      const _stickyTd1 = 'position:sticky;left:30px;z-index:1;background:inherit;';
      const _ownerName = owner.name;
      const _campaign = this.state.campaign;
      const _repRow = (rep, ri, cells) => {
        const _flagged = this._isRepFlagged(rep.name, _ownerName, _campaign);
        return `<tr id="sales-rep-row-${ri}">
          <td style="${_stickyTd0}"><input type="checkbox" class="rep-highlight-cb" ${_flagged ? 'checked' : ''} onchange="NationalApp._toggleRepHighlight(${ri}, this.checked)"></td>
          <td class="bold" style="${_stickyTd1}">${this._esc(rep.name)}</td>
          ${cells}
        </tr>`;
      };
      const repRows = isNDS
        ? s.reps.map((rep, ri) => _repRow(rep, ri, `
              <td class="num">${rep.newPorts || rep.totalVolume}</td>
              <td class="num">${rep.orderCount}</td>
              <td class="num">${this._pct(rep.cancelFraudPct)}</td>
              <td class="num">${this._pct(rep.extraPremiumPct)}</td>
              <td class="num">${this._pct(rep.nextUpPct)}</td>
              <td class="num">${this._pct(rep.abpPct)}</td>
              <td class="num">${this._pct(rep.byodPct)}</td>
              <td class="num">${this._pct(rep.newOfNewPortsPct)}</td>
              <td class="num">${this._pct(rep.insurancePct)}</td>
              <td class="num">${this._pct(rep.highMedCreditPct)}</td>
              <td class="num">${this._pct(rep.awayFromDoorsPct)}</td>
              <td class="num">${this._pct(rep.before3pmPct)}</td>
              <td class="num">${this._pct(rep.after730pmPct)}</td>`)).join('')
        : isRes
        ? s.reps.map((rep, ri) => _repRow(rep, ri, `
              <td class="num">${rep.totalVolume}</td>
              <td class="num">${rep.newInternet ?? '—'}</td>
              <td class="num">${rep.upgradeInternet ?? '—'}</td>
              <td class="num">${rep.videoSales ?? '—'}</td>
              <td class="num">${rep.wirelessSales ?? '—'}</td>
              <td class="num">${rep.voiceSales ?? '—'}</td>
              <td class="num">${this._pct(rep.abpMix)}</td>
              <td class="num">${this._pct(rep.gigMix)}</td>
              <td class="num">${this._pct(rep.techInstall)}</td>`)).join('')
        : s.reps.map((rep, ri) => _repRow(rep, ri, `
              <td class="num">${rep.totalVolume}</td>
              <td class="num">${rep.orderCount ?? '—'}</td>
              <td class="num">${rep.salesPerRep ?? '—'}</td>
              <td class="num">${rep.internet ?? '—'}</td>
              <td class="num">${rep.voip ?? '—'}</td>
              <td class="num">${rep.wireless ?? '—'}</td>
              <td class="num">${rep.airAwb ?? '—'}</td>
              <td class="num">${this._pct(rep.earlyPct)}</td>
              <td class="num">${this._pct(rep.latePct)}</td>
              <td class="num">${this._pct(rep.abpPct)}</td>
              <td class="num">${this._pct(rep.cruPct)}</td>
              <td class="num">${this._pct(rep.newWrlsPct)}</td>
              <td class="num">${this._pct(rep.byodPct)}</td>`)).join('');

      repsEl.innerHTML = `
        <div class="coaching-section">
          <div class="coaching-label">Rep Breakdown <span class="coaching-sublabel">${s.reps.length} reps</span></div>
          <div class="data-table-wrap" style="overflow-x:auto;-webkit-overflow-scrolling:touch;">
            <table class="data-table" style="min-width:900px;">
              <thead>
                <tr>${repHeaders}</tr>
              </thead>
              <tbody>${repRows}</tbody>
            </table>
          </div>
        </div>`;
    } else {
      repsEl.innerHTML = `
        <div class="coaching-section">
          <div class="coaching-label">Rep Breakdown</div>
          <div class="empty-state">
            <div class="empty-state-text">No rep data yet. Click "Refresh Data" to pull sales data from Tableau.</div>
          </div>
        </div>`;
    }
  },

  // ── Sort sales reps table by column ──
  _salesSortCol: null,
  _salesSortAsc: true,

  _sortSalesReps(col) {
    const owner = this.state.selectedOwner;
    if (!owner || !owner.sales || !owner.sales.reps) return;

    // Toggle direction if same column, otherwise default descending (except name = ascending)
    if (this._salesSortCol === col) {
      this._salesSortAsc = !this._salesSortAsc;
    } else {
      this._salesSortCol = col;
      this._salesSortAsc = col === 'name';
    }

    const dir = this._salesSortAsc ? 1 : -1;
    owner.sales.reps.sort((a, b) => {
      const va = a[col], vb = b[col];
      if (col === 'name') return dir * String(va || '').localeCompare(String(vb || ''));
      return dir * ((va || 0) - (vb || 0));
    });

    this.renderSalesTab(owner);
  },

  // ── Flag/unflag rep for one-on-one (checkbox in sales table) ──
  _toggleRepHighlight(rowIdx, checked) {
    const owner = this.state.selectedOwner;
    const reps = this._currentSalesReps || [];
    const rep = reps[rowIdx];
    if (!rep || !owner) return;

    const repName = rep.name;
    const ownerName = owner.name;
    const campaign = this.state.campaign;

    if (checked) {
      // Flag the rep
      this._postFlag('odFlagRep', { repName, ownerName, campaign, flaggedBy: this.state.session?.email || '' });
      if (!this._flaggedReps) this._flaggedReps = [];
      this._flaggedReps.push({ repName, ownerName, campaign });
    } else {
      // Unflag the rep
      this._postFlag('odUnflagRep', { repName, ownerName, campaign });
      if (this._flaggedReps) {
        this._flaggedReps = this._flaggedReps.filter(f =>
          !(f.repName.toLowerCase() === repName.toLowerCase() &&
            f.ownerName.toLowerCase() === ownerName.toLowerCase() &&
            f.campaign.toLowerCase() === campaign.toLowerCase())
        );
      }
    }
    // Update badge count
    this._updateFlaggedBadge();
  },

  async _postFlag(action, data) {
    try {
      const apiUrl = NATIONAL_CONFIG.appsScriptUrl || (typeof OD_CONFIG !== 'undefined' ? OD_CONFIG.appsScriptUrl : '');
      const apiKey = NATIONAL_CONFIG.apiKey || (typeof OD_CONFIG !== 'undefined' ? OD_CONFIG.apiKey : '');
      await fetch(apiUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({ action, key: apiKey, ...data })
      });
    } catch (err) {
      console.warn('[NationalApp] Flag POST error:', err.message);
    }
  },

  _FLAGGED_CACHE_KEY: 'od_flagged_reps_cache',

  async _fetchFlaggedReps() {
    // Load from cache first
    try {
      const raw = localStorage.getItem(this._FLAGGED_CACHE_KEY);
      if (raw) {
        const cached = JSON.parse(raw);
        if (cached.reps) this._flaggedReps = cached.reps;
      }
    } catch { /* ignore */ }

    // Fetch fresh
    try {
      const apiUrl = NATIONAL_CONFIG.appsScriptUrl || (typeof OD_CONFIG !== 'undefined' ? OD_CONFIG.appsScriptUrl : '');
      const apiKey = NATIONAL_CONFIG.apiKey || (typeof OD_CONFIG !== 'undefined' ? OD_CONFIG.apiKey : '');
      const url = new URL(apiUrl);
      url.searchParams.set('key', apiKey);
      url.searchParams.set('action', 'odGetFlaggedReps');
      const res = await fetch(url.toString()).then(r => r.json());
      if (res.success && res.reps) {
        this._flaggedReps = res.reps;
        localStorage.setItem(this._FLAGGED_CACHE_KEY, JSON.stringify({ reps: res.reps, _ts: Date.now() }));
      }
    } catch (err) {
      console.warn('[NationalApp] Fetch flagged reps error:', err.message);
    }
    this._updateFlaggedBadge();
  },

  // ── Fetch owner notes ──
  async _fetchNotes() {
    try {
      const apiUrl = NATIONAL_CONFIG.appsScriptUrl || (typeof OD_CONFIG !== 'undefined' ? OD_CONFIG.appsScriptUrl : '');
      const apiKey = NATIONAL_CONFIG.apiKey || (typeof OD_CONFIG !== 'undefined' ? OD_CONFIG.apiKey : '');
      const url = new URL(apiUrl);
      url.searchParams.set('key', apiKey);
      url.searchParams.set('action', 'odGetNotes');
      const res = await fetch(url.toString()).then(r => r.json());
      if (res.success && res.notes) {
        this.state.campaignNotes = res.notes;
      }
    } catch (err) {
      console.warn('[NationalApp] Fetch notes error:', err.message);
    }
  },

  _isRepFlagged(repName, ownerName, campaign) {
    if (!this._flaggedReps) return false;
    return this._flaggedReps.some(f =>
      f.repName.toLowerCase() === repName.toLowerCase() &&
      f.ownerName.toLowerCase() === ownerName.toLowerCase() &&
      f.campaign.toLowerCase() === campaign.toLowerCase()
    );
  },

  _updateFlaggedBadge() {
    const count = (this._flaggedReps || []).length;
    const badge = document.getElementById('planning-notif-badge');
    if (badge) {
      badge.textContent = count;
      badge.style.display = count > 0 ? '' : 'none';
    }
  },

  // ══════════════════════════════════════════════════
  // RENDER: Audit Tab (Online Presence)
  // ══════════════════════════════════════════════════

  /**
   * Check if an owner is marked as Non-Partner for a specific column.
   * Reads from OwnerDev.state.mappings if available (embedded mode).
   * @param {string} ownerName
   * @param {string} field - 'camCompany', 'nlrWorkbookId', or 'nlrTab'
   * @returns {boolean}
   */
  _isNonPartner(ownerName, field) {
    if (typeof OwnerDev === 'undefined' || !OwnerDev.state?.mappings) return false;
    const campaign = this.state.campaign || '';
    const mapping = OwnerDev.state.mappings.find(
      m => m.campaign === campaign && (m.ownerName || '').toLowerCase() === (ownerName || '').toLowerCase()
    );
    if (!mapping) return false;
    return (mapping[field] || '').toLowerCase() === 'non-partner';
  },

  /**
   * Check if an owner has NO mapping value for a specific column (not yet filled in).
   * @param {string} ownerName
   * @param {string} field - 'camCompany', 'nlrWorkbookId', or 'nlrTab'
   * @returns {boolean}
   */
  _isUnmapped(ownerName, field) {
    if (typeof OwnerDev === 'undefined' || !OwnerDev.state?.mappings) return true;
    const campaign = this.state.campaign || '';
    const mapping = OwnerDev.state.mappings.find(
      m => m.campaign === campaign && (m.ownerName || '').toLowerCase() === (ownerName || '').toLowerCase()
    );
    if (!mapping) return true;
    return !(mapping[field] || '').trim();
  },

  renderAuditTab(owner) {
    const a = owner.audit;
    const bizList = a.businesses || [];
    const total = bizList.length;

    // ── Check if owner is Non-Partner or unmapped for BIS ──
    const bisNonPartner = this._isNonPartner(owner.name, 'camCompany');
    const bisUnmapped = this._isUnmapped(owner.name, 'camCompany');
    if (bisNonPartner || (bisUnmapped && !bizList.length)) {
      const grades = document.getElementById('audit-grades');
      grades.innerHTML = `
        <div class="bis-report-header">
          <img src="https://betterimagesolutions.com/wp-content/uploads/2025/12/cropped-BIS_Standard-scaled-e1766050595589.png"
               alt="Better Image Solutions" class="bis-logo" onerror="this.style.display='none'">
          <div class="bis-report-title">Monthly Performance Audit Report</div>
        </div>`;
      const details = document.getElementById('audit-details');
      details.innerHTML = `
        <div class="empty-state" style="padding:40px 20px;">
          <div class="empty-state-icon">🚫</div>
          <div class="empty-state-text" style="font-size:1.1rem;color:#4a7090;">
            Not Utilizing BIS Services
          </div>
        </div>`;
      return;
    }

    // ── BIS-style report header ──
    const grades = document.getElementById('audit-grades');
    const now = new Date();
    const auditMonth = now.toLocaleString('default', { month: 'long', year: 'numeric' });

    grades.innerHTML = `
      <div class="bis-report-header">
        <img src="https://betterimagesolutions.com/wp-content/uploads/2025/12/cropped-BIS_Standard-scaled-e1766050595589.png"
             alt="Better Image Solutions" class="bis-logo" onerror="this.style.display='none'">
        <div class="bis-report-title">Monthly Performance Audit Report</div>
        <div class="bis-report-sub">Powered by Better Image Solutions · ${auditMonth}</div>
      </div>`;

    const details = document.getElementById('audit-details');

    if (!bizList.length) {
      details.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-icon">🏢</div>
          <div class="empty-state-text">No companies mapped for this owner.</div>
        </div>`;
      return;
    }

    // If multiple companies → show a dropdown selector, one report at a time
    if (bizList.length > 1) {
      const options = bizList.map((b, i) =>
        `<option value="${i}">${this._esc(b.businessName || b.clientName)}</option>`
      ).join('');

      details.innerHTML = `
        <div class="bis-company-selector">
          <label class="bis-selector-label">Company</label>
          <select class="bis-selector-dropdown" id="bis-company-select"
            onchange="NationalApp._switchBizReport()">
            ${options}
          </select>
        </div>
        <div class="bis-reports" id="bis-reports-container">
          ${this._renderBizReport(bizList[0], auditMonth)}
        </div>`;
      this._currentBizList = bizList;
      this._currentAuditMonth = auditMonth;
    } else {
      details.innerHTML = `
        <div class="bis-reports">
          ${this._renderBizReport(bizList[0], auditMonth)}
        </div>`;
    }
  },

  // ── Render the Claim Companies section ──
  _renderClaimSection(owner) {
    const mapping = this.state.camMapping || {};
    const claimed = mapping[owner.name] || [];
    const allNames = this.state.allCompanyNames || [];

    // Build set of all claimed company names (across all owners) for exclusion
    const allClaimed = new Set();
    for (const companies of Object.values(mapping)) {
      for (const c of companies) allClaimed.add(c.toLowerCase());
    }

    // Unclaimed = all company names not yet claimed by anyone
    const unclaimed = allNames.filter(n => !allClaimed.has(n.toLowerCase()));

    const chipsHTML = claimed.map(c =>
      `<span class="claimed-chip">
        ${this._esc(c)}
        <button class="claimed-chip-x" onclick="NationalApp.unclaimCompany('${this._esc(c.replace(/'/g, "\\'"))}')" title="Remove">&times;</button>
      </span>`
    ).join('');

    const optionsHTML = unclaimed.map(c =>
      `<div class="claim-dropdown-item" onclick="NationalApp.claimCompany('${this._esc(c.replace(/'/g, "\\'"))}')">${this._esc(c)}</div>`
    ).join('');

    return `
      <div class="claim-section">
        <div class="claim-header">
          <div class="section-label" style="margin:0">Claimed Companies</div>
          <span class="claim-count">${claimed.length} claimed · ${unclaimed.length} available</span>
        </div>
        ${claimed.length ? `<div class="claimed-chips">${chipsHTML}</div>` : ''}
        <div class="claim-search-wrap">
          <input type="text" class="claim-search-input" id="claim-search"
            placeholder="Search companies to claim..."
            oninput="NationalApp._filterClaimDropdown(this.value)"
            onfocus="NationalApp._showClaimDropdown()"
            autocomplete="off">
          <div class="claim-dropdown" id="claim-dropdown" style="display:none">
            ${optionsHTML || '<div class="claim-dropdown-empty">No unclaimed companies available</div>'}
          </div>
        </div>
        <div class="claim-status" id="claim-status"></div>
      </div>`;
  },

  _showClaimDropdown() {
    const dd = document.getElementById('claim-dropdown');
    if (dd) dd.style.display = 'block';
    // Close on click outside
    setTimeout(() => {
      const handler = (e) => {
        const wrap = e.target.closest('.claim-search-wrap');
        if (!wrap) {
          dd.style.display = 'none';
          document.removeEventListener('click', handler);
        }
      };
      document.addEventListener('click', handler);
    }, 10);
  },

  _filterClaimDropdown(query) {
    const dd = document.getElementById('claim-dropdown');
    if (!dd) return;
    dd.style.display = 'block';
    const q = query.toLowerCase().trim();
    const items = dd.querySelectorAll('.claim-dropdown-item');
    let visible = 0;
    items.forEach(item => {
      const match = !q || item.textContent.toLowerCase().includes(q);
      item.style.display = match ? 'block' : 'none';
      if (match) visible++;
    });
    // Show/hide empty state
    const empty = dd.querySelector('.claim-dropdown-empty');
    if (empty) empty.style.display = visible === 0 ? 'block' : 'none';
    // If all items are filtered out, show empty message if it doesn't exist
    if (visible === 0 && !empty) {
      dd.insertAdjacentHTML('beforeend', '<div class="claim-dropdown-empty">No matches found</div>');
    }
  },

  _switchBizReport() {
    const select = document.getElementById('bis-company-select');
    const container = document.getElementById('bis-reports-container');
    if (!select || !container || !this._currentBizList) return;
    const idx = parseInt(select.value, 10);
    const biz = this._currentBizList[idx];
    if (biz) {
      container.innerHTML = this._renderBizReport(biz, this._currentAuditMonth);
    }
  },

  async claimCompany(companyName) {
    const owner = this.state.selectedOwner;
    if (!owner) return;

    const status = document.getElementById('claim-status');
    if (status) { status.textContent = 'Claiming...'; status.className = 'claim-status'; }

    try {
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'claimCompany',
          ownerName: owner.name,
          companyName: companyName
        })
      });
      const result = await resp.json();
      if (result.error) throw new Error(result.error);

      // Update mapping in state
      this.state.camMapping = result.mapping || {};

      // Re-map audit data to owners with new mapping and re-render
      this._remapAndRenderAudit();

      if (status) {
        status.textContent = 'Claimed: ' + companyName;
        status.className = 'claim-status claim-success';
        setTimeout(() => { status.textContent = ''; }, 4000);
      }
    } catch (err) {
      console.error('[NationalApp] Claim failed:', err);
      if (status) {
        status.textContent = 'Error: ' + err.message;
        status.className = 'claim-status claim-error';
      }
    }
  },

  async unclaimCompany(companyName) {
    const owner = this.state.selectedOwner;
    if (!owner) return;

    const status = document.getElementById('claim-status');
    if (status) { status.textContent = 'Removing...'; status.className = 'claim-status'; }

    try {
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'unclaimCompany',
          ownerName: owner.name,
          companyName: companyName
        })
      });
      const result = await resp.json();
      if (result.error) throw new Error(result.error);

      // Update mapping in state
      this.state.camMapping = result.mapping || {};

      // Re-map audit data to owners with new mapping and re-render
      this._remapAndRenderAudit();

      if (status) {
        status.textContent = 'Removed: ' + companyName;
        status.className = 'claim-status claim-success';
        setTimeout(() => { status.textContent = ''; }, 4000);
      }
    } catch (err) {
      console.error('[NationalApp] Unclaim failed:', err);
      if (status) {
        status.textContent = 'Error: ' + err.message;
        status.className = 'claim-status claim-error';
      }
    }
  },

  // Re-map all audit businesses to owners with current mapping, then re-render
  _remapAndRenderAudit() {
    // Clear existing audit businesses for all owners
    for (const o of this.state.owners) {
      o.audit.businesses = [];
      o.audit.grades = { reviews: '—', website: '—', social: '—', seo: '—' };
    }

    // Re-run mapping with stored audit data
    if (this._cachedAuditBusinesses && this._cachedAuditBusinesses.length) {
      this._mapAuditToOwners(this._cachedAuditBusinesses, this.state.camMapping);
    }

    // Re-render the current owner's audit tab
    const owner = this.state.selectedOwner;
    if (owner) this.renderAuditTab(owner);
  },

  // ── Render a single business card ──
  // ═══════════════════════════════════════════════════════
  // BIS-STYLE PERFORMANCE AUDIT REPORT RENDERING
  // ═══════════════════════════════════════════════════════

  _renderBizReport(b, auditMonth) {
    const e = s => this._esc(s || '');
    const extUrl = s => { if (!s) return ''; s = s.trim(); return /^https?:\/\//i.test(s) ? s : 'https://' + s; };
    const bizName = e(b.businessName || b.clientName);

    // ── Compute section grades ──
    const reviewGrade = this._bizReviewGrade(b);
    const websiteGrade = this._bizWebsiteGrade(b);
    const socialGrade = this._bizSocialGrade(b);

    // ── Status / Service ──
    const st = b.serviceStatus;
    const rawStatus = st?.status || (st?.full ? 'Full' : st?.lite ? 'Lite' : '');
    const cleanStatus = rawStatus.startsWith('#') ? '' : rawStatus;

    // ── Reviews Section ──
    const platforms = [
      { name: 'Google', data: b.gbl },
      { name: 'Glassdoor', data: b.glassdoor },
      { name: 'Indeed', data: b.indeed },
      { name: b.other?.platform || 'Other', data: b.other }
    ];

    const reviewsHTML = `
      <div class="bis-section">
        <div class="bis-grade-box ${this._gradeClass(reviewGrade)}">
          <div class="bis-grade-letter">${reviewGrade}</div>
        </div>
        <div class="bis-section-content">
          <div class="bis-section-banner">Reviews</div>
          <table class="bis-table">
            <thead>
              <tr>
                <th></th>
                ${platforms.map(p => {
                  const link = p.data?.link;
                  return link
                    ? `<th class="bis-platform-header"><a href="${extUrl(link)}" target="_blank" rel="noopener" class="bis-link">${e(p.name)} ↗</a></th>`
                    : `<th class="bis-platform-header">${e(p.name)}</th>`;
                }).join('')}
              </tr>
            </thead>
            <tbody>
              <tr>
                <td class="bis-row-label">Rating</td>
                ${platforms.map(p => {
                  const rating = p.data?.rating;
                  const cls = this._ratingColor(rating);
                  return `<td class="bis-big-value ${cls}">${rating != null && rating > 0 ? rating.toFixed(1) : '—'}</td>`;
                }).join('')}
              </tr>
              <tr>
                <td class="bis-row-label"># of Reviews</td>
                ${platforms.map(p => {
                  const rev = p.data?.reviews;
                  return `<td class="bis-big-value">${rev != null && rev > 0 ? rev : '—'}</td>`;
                }).join('')}
              </tr>
              <tr class="bis-notes-row">
                <td class="bis-row-label">Notes</td>
                ${platforms.map(p => {
                  const notes = p.data?.notes || '';
                  return `<td class="bis-notes-cell">${e(notes)}</td>`;
                }).join('')}
              </tr>
            </tbody>
          </table>
        </div>
      </div>`;

    // ── Website Section ──
    const ws = b.website;
    const blog = b.blog;
    const blogCount3 = (blog?.threeMonthCount > 10000) ? 0 : (blog?.threeMonthCount || 0);

    const websiteHTML = `
      <div class="bis-section">
        <div class="bis-grade-box ${this._gradeClass(websiteGrade)}">
          <div class="bis-grade-letter">${websiteGrade}</div>
        </div>
        <div class="bis-section-content">
          <div class="bis-section-banner">${ws?.url ? `<a href="${extUrl(ws.url)}" target="_blank" rel="noopener" class="bis-link">Website ↗</a>` : 'Website'}</div>
          <table class="bis-table bis-table-website">
            <thead>
              <tr>
                <th colspan="2" class="bis-sub-header">Site Update</th>
                <th colspan="2" class="bis-sub-header">Blog</th>
              </tr>
              <tr>
                <th>Site Photos</th>
                <th>Last Updated / Reviewed</th>
                <th>Last Updated</th>
                <th>In Queue</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td class="bis-big-value">${e(ws?.sitePhotos) || '—'}</td>
                <td class="bis-big-value">${e(ws?.lastUpdated) || '—'}</td>
                <td class="bis-big-value">${e(blog?.lastBlogPost) || '—'}</td>
                <td class="bis-big-value">${blog?.onQueue != null ? blog.onQueue : '—'}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>`;

    // ── Social Media Section ──
    const ig = b.instagram || {};
    const socialHTML = `
      <div class="bis-section">
        <div class="bis-grade-box ${this._gradeClass(socialGrade)}">
          <div class="bis-grade-letter">${socialGrade}</div>
        </div>
        <div class="bis-section-content">
          <div class="bis-section-banner">${ig.link ? `<a href="${extUrl(ig.link)}" target="_blank" rel="noopener" class="bis-link">Social Media ↗</a>` : 'Social Media'}</div>
          <table class="bis-table">
            <thead>
              <tr>
                <th>Shared</th>
                <th>Generated</th>
                <th>Followers</th>
                <th>Following</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td class="bis-big-value">${ig.shared != null ? ig.shared : '—'}</td>
                <td class="bis-big-value">${ig.generated != null ? ig.generated : '—'}</td>
                <td class="bis-big-value">${ig.followers != null ? this._fmtNum(ig.followers) : '—'}</td>
                <td class="bis-big-value">${ig.following != null ? this._fmtNum(ig.following) : '—'}</td>
              </tr>
              ${b.otherNotes ? `<tr class="bis-notes-row"><td colspan="4" class="bis-notes-cell">${e(b.otherNotes)}</td></tr>` : ''}
            </tbody>
          </table>
        </div>
      </div>`;

    // ── SEO row (compact, if available) ──
    const seoVal = (b.seo?.check || '').toLowerCase();
    const seoPass = seoVal === 'pass' || seoVal === 'yes' || seoVal === 'y' || seoVal === '✓' || seoVal === 'true' || seoVal === 'x' || seoVal === 'good';
    const seoHTML = b.seo?.check ? `
      <div class="bis-seo-row">
        <span class="bis-seo-label">SEO</span>
        <span class="bis-seo-badge ${seoPass ? 'seo-pass' : 'seo-fail'}">${seoPass ? '✓ Passing' : '✗ Needs Work'}</span>
      </div>` : '';

    return `
      <div class="bis-report-card">
        <div class="bis-company-bar">
          <div class="bis-company-name">${bizName}</div>
          ${cleanStatus ? `<span class="bis-company-status">${e(cleanStatus)}</span>` : ''}
        </div>
        ${b.accountManager ? `<div class="bis-account-manager">${e(b.accountManager)}${b.services ? ' · ' + e(b.services) : ''}</div>` : ''}
        ${reviewsHTML}
        ${websiteHTML}
        ${socialHTML}
        ${seoHTML}
      </div>`;
  },

  // ── Per-business grade helpers ──
  // Determine service tier: 'full' or 'lite'
  _bizTier(b) {
    const st = b.serviceStatus;
    if (st?.lite && !st?.full) return 'lite';
    return 'full'; // default to full if both set or neither set
  },

  _bizReviewGrade(b) {
    const rating = b.gbl?.rating;
    if (rating != null && rating > 0) return this._ratingToGrade(rating, b.gbl?.reviews || 0);
    return '—';
  },

  // Website grade — matches Cam's Performance Audit rubric
  // Full: Team Page photos + Blog (last post recency + queue depth)
  // Lite: Team Page photos + Updated on current month (no blog)
  _bizWebsiteGrade(b) {
    const ws = b.website;
    const blog = b.blog;
    if (!ws?.url && !ws?.sitePhotos && !blog?.url) return '—';

    const photos = (ws?.sitePhotos || '').toLowerCase();
    const isLite = this._bizTier(b) === 'lite';

    // Classify photo quality
    const noStock = photos.includes('no stock');
    const someStock = photos.includes('some stock');
    const mostlyStock = photos.includes('mostly stock');
    const allStock = photos.includes('all stock');
    // If none of the keywords match, infer from content
    const photoTier = noStock ? 4 : someStock ? 3 : mostlyStock ? 2 : allStock ? 1 : (photos ? 3 : 0);
    // 4=no stock, 3=some stock, 2=mostly stock, 1=all stock, 0=unknown

    if (isLite) {
      // ── Website (Lite) — photos + updated this month ──
      const updated = (ws?.updatedMonth || '').toLowerCase();
      const isUpdated = ['yes','y','✓','true','x'].includes(updated);
      if (photoTier >= 4 && isUpdated) return 'A+';  // no stock + updated
      if (photoTier >= 4) return 'A';                  // no stock, not updated
      if (photoTier >= 3 && isUpdated) return 'B';     // some stock + updated
      if (photoTier >= 3) return 'B-';                 // some stock, not updated
      if (photoTier >= 2) return 'C';                  // mostly stock
      return 'D';                                       // all stock
    }

    // ── Website (Full) — photos + blog recency + queue ──
    const blogCurrent = (blog?.currentMonth || 0) > 0;
    const blogRecent = (blog?.threeMonthCount || 0) > 0;
    const queue = blog?.onQueue || 0;

    // Blog recency: current month > previous month > 1-of-3 > none
    // "Current Month" = blogCurrent, "Previous Month" = !blogCurrent && blogRecent
    // "1 out of 3 months" = threeMonthCount > 0 (sparse), "Blog not updated" = neither

    if (photoTier >= 4) {
      // No stock photos
      if (blogCurrent && queue >= 3) return 'A+';
      if (blogCurrent && queue >= 1) return 'A';
      if (blogRecent && queue >= 1) return 'A-';
      if (blogCurrent || blogRecent) return 'A-';
      return 'B'; // no stock but no blog activity
    }
    if (photoTier >= 3) {
      // Some stock photos
      if (blogCurrent && queue >= 1) return 'B';
      if (blogRecent) return 'B-';
      return 'C';
    }
    if (photoTier >= 2) {
      // Mostly stock photos
      if (blogCurrent && queue >= 1) return 'C';
      if (queue >= 1) return 'C-';
      return 'D';
    }
    // All stock photos
    return 'D';
  },

  // Social Media grade — matches Cam's Performance Audit rubric
  // Full: Shared Posts per week + Followers Count
  // Lite: lower thresholds (>=1/week for A+)
  _bizSocialGrade(b) {
    const ig = b.instagram;
    if (!ig || (!ig.link && !ig.followers && !ig.shared && !ig.generated)) return '—';
    const shared = ig.shared || 0;
    const followers = ig.followers || 0;
    const isLite = this._bizTier(b) === 'lite';

    if (isLite) {
      // ── Social Media (Lite) ──
      if (shared >= 1 && followers >= 500) return 'A+';
      if (shared >= 1 && followers >= 100) return 'A';
      if (shared >= 1) return 'B';
      if (shared > 0) return 'C';  // <1 per week but some activity
      return 'D';
    }

    // ── Social Media (Full) ──
    if (shared >= 3 && followers >= 500) return 'A+';
    if (shared >= 3) return 'A';
    if (shared >= 2) return 'B';
    if (shared >= 1) return 'C';
    return 'D';
  },

  _ratingColor(rating) {
    if (rating == null) return '';
    if (rating >= 4.5) return 'rating-great';
    if (rating >= 4.0) return 'rating-good';
    if (rating >= 3.0) return 'rating-ok';
    return 'rating-bad';
  },

  _fmtNum(n) {
    if (n == null || n === 0) return '0';
    if (n >= 1000) return (n / 1000).toFixed(1) + 'K';
    return String(n);
  },

  // ══════════════════════════════════════════════════
  // REUSABLE: Recruiting Table (spreadsheet format)
  // ══════════════════════════════════════════════════

  _renderRecruitingTable(data, containerId) {
    const el = document.getElementById(containerId);
    if (!el) return;
    if (!data || !data.rows || !data.rows.length) {
      el.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-text">Recruiting table data will populate once connected.</div>
        </div>`;
      return;
    }

    const weeks = [...(data.weeks || [])].reverse();   // newest-first
    const rows = data.rows || [];
    const leaders = data.leaders || 0;

    let html = '';

    // 2nd Rounds Required gauge — scales with leader count
    html += this._build2ndRoundsGauge(leaders);

    // Table
    html += `<div class="rt-wrap"><table class="rt-table"><thead>`;

    // Group header row
    html += `<tr class="rt-group-row">
      <th></th>
      <th class="rt-group-projected">Projected Weekly<br>Numbers Needed</th>
      <th class="rt-group-actual" colspan="${weeks.length}"></th>
      <th class="rt-group-total">Total / Month<br>Overview</th>
    </tr>`;

    // Date header row (already reversed)
    html += `<tr class="rt-date-row">
      <th></th>
      <th></th>
      ${weeks.map(w => `<th>${this._esc(w)}</th>`).join('')}
      <th></th>
    </tr>`;

    html += `</thead><tbody>`;

    // Detect all-zero week columns (data never entered — show dashes)
    const numWeeks = weeks.length;
    const zeroWeeks = new Array(numWeeks).fill(true);
    rows.forEach(row => {
      const vals = [...row.values].reverse();
      vals.forEach((v, wi) => { if (v !== 0 && v !== '0%') zeroWeeks[wi] = false; });
    });

    // Data rows
    rows.forEach((row, ri) => {
      html += `<tr>`;
      html += `<td>${this._esc(row.label)}</td>`;
      html += `<td class="rt-projected">${this._fmtCell(row.projected, row.isRate)}</td>`;

      // Weekly values with conditional coloring (reversed: newest-first)
      const vals = [...row.values].reverse();
      vals.forEach((val, wi) => {
        if (zeroWeeks[wi]) {
          html += `<td class="rt-dash">—</td>`;
        } else {
          const color = this._cellColor(val, row.projected, row.isRate, ri);
          html += `<td class="${color}">${this._fmtCell(val, row.isRate)}</td>`;
        }
      });

      // Total column
      const totalColor = this._cellColor(
        row.total,
        row.isRate ? row.projected : row.projected * weeks.length,
        row.isRate,
        ri
      );
      html += `<td class="rt-total ${totalColor}">${this._fmtCell(row.total, row.isRate)}</td>`;

      html += `</tr>`;
    });

    html += `</tbody></table></div>`;
    el.innerHTML = html;
  },

  // ── 2nd Rounds Required gauge (scales with leader count) ──
  _build2ndRoundsGauge(leaders) {
    const tiers = [
      { label: 'Bored Leaders',                mult: 2, color: '#e53935' },
      { label: 'Top Leaders Interviewing Only', mult: 3, color: '#f9a825' },
      { label: 'Maintaining, Not Growing',      mult: 4, color: '#ffeb3b' },
      { label: 'Leaders Busy',                  mult: 5, color: '#84cc16' },
      { label: 'Promotion Factory',             mult: 6, color: '#22d3ee' }
    ];
    let h = `<table class="rt-gauge-table">`;
    h += `<thead><tr><th colspan="2"># of Leaders: ${leaders}</th></tr></thead>`;
    h += `<tbody>`;
    h += `<tr class="rt-gauge-header"><td colspan="2">2ND ROUNDS REQUIRED</td></tr>`;
    tiers.forEach(t => {
      const val = leaders * t.mult;
      const isLight = t.color === '#ffeb3b' || t.color === '#f9a825' || t.color === '#84cc16';
      const shadow = isLight ? '0 1px 3px rgba(0,0,0,0.55), 0 0 8px rgba(0,0,0,0.3)' : '0 1px 3px rgba(0,0,0,0.45), 0 0 6px rgba(0,0,0,0.2)';
      h += `<tr><td>${this._esc(t.label)}</td><td class="rt-gauge-val" style="background:${t.color};color:#fff;text-shadow:${shadow}">${val}</td></tr>`;
    });
    h += `</tbody></table>`;
    return h;
  },

  // ── Status codes legend ──
  _renderStatusLegend(containerId) {
    const el = document.getElementById(containerId);
    if (!el) return;

    // Count owners per status code
    const counts = {};
    this.state.owners.forEach(o => {
      if (o.statusCode) counts[o.statusCode] = (counts[o.statusCode] || 0) + 1;
    });

    el.innerHTML = `
      <div class="section-label">Leader Status Codes</div>
      <div class="status-legend">
        ${Object.entries(this.STATUS_CODES).map(([code, def]) => {
          const cnt = counts[code] || 0;
          return `<span class="status-legend-item ${def.css}">${code} — ${def.label}${cnt ? ' (' + cnt + ')' : ''}</span>`;
        }).join('')}
      </div>`;
  },

  // ══════════════════════════════════════════════════
  // HELPERS
  // ══════════════════════════════════════════════════

  _isSuperadmin() {
    return typeof OwnerDev !== 'undefined' && OwnerDev.state && OwnerDev.state.isSuperadmin;
  },

  _esc(s) {
    const d = document.createElement('div');
    d.textContent = s || '';
    return d.innerHTML;
  },

  // Get the display label for the current campaign (used as the sheet tab name).
  // Checks multiple sources since config may not be populated on all code paths.
  _getCampaignLabel() {
    const key = this.state.campaign;
    if (!key) return '';
    // 1. NATIONAL_CONFIG (set by _populateCampaignSelector or loadCampaignData)
    const cfg = NATIONAL_CONFIG.campaigns[key];
    if (cfg?.label) return cfg.label;
    // 2. _allCampaignsData (in-memory cache from server)
    const acd = this._allCampaignsData?.[key];
    if (acd?.label) return acd.label;
    // 3. Title-case the key as last resort (frontier → Frontier)
    return key.charAt(0).toUpperCase() + key.slice(1).replace(/-./g, m => ' ' + m[1].toUpperCase());
  },

  // Invalidate ALL caches so next load fetches fresh data from server.
  // Called after headcount/production saves so manual edits aren't lost on reload.
  _invalidateOdCache() {
    try { localStorage.removeItem('od_data_cache'); } catch {}
    // Clear per-campaign coach cache (coach_cache_frontier, etc.)
    try {
      const keys = Object.keys(localStorage);
      for (const k of keys) {
        if (k.startsWith('coach_cache_')) localStorage.removeItem(k);
      }
    } catch {}
    // Also bust the in-memory campaign cache so re-entering triggers fresh server fetch
    if (this._allCampaignsData && this.state.campaign) {
      delete this._allCampaignsData[this.state.campaign];
    }
  },

  _formatCurrentWeek() {
    const d = new Date();
    const mon = d.toLocaleString('en-US', { month: 'short' });
    const day = d.getDate();
    return `${mon} ${day}, ${d.getFullYear()}`;
  },

  _formatWeekDate(tabName) {
    const parts = String(tabName).split('/');
    if (parts.length < 3) return tabName;
    const d = new Date(+parts[2], +parts[0] - 1, +parts[1]);
    if (isNaN(d)) return tabName;
    return d.toLocaleString('en-US', { month: 'short' }) + ' ' + d.getDate() + ', ' + d.getFullYear();
  },

  // Format cell value (add % suffix for rate rows)
  _fmtCell(val, isRate) {
    if (val === null || val === undefined || val === '—') return '—';
    return isRate ? val + '%' : val;
  },

  // Per-row color thresholds: [blue≥, green≥, yellow≥, orange≥, else red]
  // Rows 0-3 (Applies, Sent to List, 1st Booked, 1st Showed): ratio-based
  // Row 4 (1st Retention): absolute %
  // Row 5 (% Call List Booked): absolute %
  // Row 6 (2nds Booked): ratio-based
  // Row 7 (2nds Showed): ratio-based
  // Row 8 (2nd Retention): absolute %
  // Row 9 (New Starts Booked): ratio-based
  // Row 10 (New Starts Showed): ratio-based
  // Row 11 (New Start Retention): absolute %
  _COLOR_THRESHOLDS: {
    // ratio thresholds: value / projected [blue≥, green≥, yellow≥, orange≥]
    ratio: {
      0:  [1.20, 1.00, 0.85, 0.70],  // Applies Received
      1:  [1.20, 1.00, 0.85, 0.70],  // Sent to List
      2:  [1.20, 1.00, 0.85, 0.70],  // 1st Rounds Booked
      3:  [1.20, 1.00, 0.85, 0.70],  // 1st Rounds Showed
      6:  [1.20, 1.00, 0.80, 0.50],  // 2nd Rounds Booked
      7:  [1.20, 1.00, 0.50, 0.30],  // 2nd Rounds Showed
      9:  [1.20, 1.00, 0.85, 0.65],  // New Starts Booked
      10: [1.20, 1.00, 0.80, 0.60],  // New Starts Showed
    },
    // absolute % thresholds [blue≥, green≥, yellow≥, orange≥]
    absolute: {
      4:  [60, 50, 45, 35],  // 1st Retention
      5:  [50, 45, 38, 35],  // % Call List Booked
      8:  [60, 50, 40, 35],  // 2nd Retention
      11: [60, 50, 40, 35],  // New Start Retention
    }
  },

  // Conditional cell color based on actual vs projected + row-specific thresholds
  _cellColor(actual, projected, isRate, rowIdx) {
    if (actual === null || actual === undefined || actual === '—') return '';
    const a = parseFloat(actual);
    if (isNaN(a)) return '';

    const abs = this._COLOR_THRESHOLDS.absolute[rowIdx];
    if (abs) {
      // Absolute % comparison (rate rows)
      if (a >= abs[0]) return 'cell-blue';
      if (a >= abs[1]) return 'cell-green';
      if (a >= abs[2]) return 'cell-yellow';
      if (a >= abs[3]) return 'cell-orange';
      return 'cell-red';
    }

    const rat = this._COLOR_THRESHOLDS.ratio[rowIdx];
    if (rat) {
      if (projected === null || projected === undefined || projected === '—') return '';
      const p = parseFloat(projected);
      if (isNaN(p) || p === 0) return '';
      const ratio = a / p;
      if (ratio >= rat[0]) return 'cell-blue';
      if (ratio >= rat[1]) return 'cell-green';
      if (ratio >= rat[2]) return 'cell-yellow';
      if (ratio >= rat[3]) return 'cell-orange';
      return 'cell-red';
    }

    return '';
  },

  _statusColor(code) {
    const sc = this.STATUS_CODES[code];
    return sc ? sc.css : '';
  },

  _gradeClass(grade) {
    if (!grade || grade === '—') return '';
    const g = String(grade).toUpperCase().charAt(0);
    if (g === 'A') return 'grade-a';
    if (g === 'B') return 'grade-b';
    if (g === 'C') return 'grade-c';
    if (g === 'D') return 'grade-d';
    return 'grade-f';
  },

  _showLoading(msg) {
    const s = document.getElementById('loading-screen');
    const t = document.getElementById('loading-text');
    if (s) s.style.display = 'flex';
    if (t && msg) t.textContent = msg;
  },

  _hideLoading() {
    const s = document.getElementById('loading-screen');
    if (s) s.style.display = 'none';
  }
};

// ── Boot ──
// Only auto-init on national.html (standalone). When embedded in owner-dev.html,
// OwnerDev calls NationalApp.initCoachView() from the Coach tab.
document.addEventListener('DOMContentLoaded', () => {
  if (!document.getElementById('view-coach')) {
    NationalApp.init();
  }
});
