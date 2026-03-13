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

  state: {
    campaign: 'att-b2b',
    owners: [],
    selectedOwner: null,
    currentTab: 'health',
    session: null,
    campaignTotals: {},
    campaignRecruiting: null,
    camMapping: {},
    costSheets: {},
    costFilesList: null,
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
    this._showLoading('Loading campaign data...');
    document.getElementById('user-name').textContent = this.state.session.name || this.state.session.email;

    try {
      await this.loadCampaignData(this.state.campaign);
      this._hideLoading();
      document.getElementById('dashboard').style.display = 'block';
      this.renderCampaignOverview();
      this.renderOwnersList();
    } catch (err) {
      console.error('[NationalApp] init error:', err);
      this._hideLoading();
      document.getElementById('dashboard').style.display = 'block';
      this.renderCampaignOverview();
      this.renderOwnersList();
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
    const isB2B = campaignKey === 'att-b2b';

    // ── Fire ALL independent fetches in parallel (each with 20s timeout) ──
    const fetchPromises = {};
    if (hasNational) fetchPromises.recruiting = this._fetchWithTimeout(this._fetchRecruitingFromSheet(campaignKey));
    if (hasApi)      fetchPromises.audit      = this._fetchWithTimeout(this._fetchOnlinePresence());
    if (hasApi)      fetchPromises.camMapping  = this._fetchWithTimeout(this._fetchOwnerCamMapping());
    if (isB2B && hasApi) fetchPromises.headcount  = this._fetchWithTimeout(this._fetchB2BHeadcount());
    if (isB2B && hasApi) fetchPromises.production = this._fetchWithTimeout(this._fetchB2BProduction());
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

    // ── Build owners from recruiting data (or scaffold) ──
    const sheetData = results.recruiting || null;
    if (sheetData && sheetData.owners && sheetData.owners.length) {
      this._buildOwnersFromSheet(campaignKey, sheetData);
    } else {
      console.log('[NationalApp] No recruiting data, using scaffold');
      this._loadScaffoldData(campaignKey);
    }

    // ── Apply enrichments from parallel results ──
    if (results.camMapping) {
      this.state.camMapping = results.camMapping.mapping || results.camMapping;
      this.state.costSheets = results.camMapping.costSheets || {};
    }

    if (results.audit && results.audit.businesses && results.audit.businesses.length) {
      this.state.allCompanyNames = results.audit.allCompanyNames || [];
      this._cachedAuditBusinesses = results.audit.businesses;
      this._mapAuditToOwners(results.audit.businesses, this.state.camMapping || null);
    }

    if (results.headcount && results.headcount.owners && Object.keys(results.headcount.owners).length) {
      this._enrichOwnersWithNLR(results.headcount.owners);
    }

    if (results.production && results.production.owners && Object.keys(results.production.owners).length) {
      this._enrichOwnersWithProduction(results.production.owners);
    }

    // Recruiting costs are loaded per-owner when the Recruiting tab is opened (via _loadAndRenderCosts)
  },

  // ── Fetch ALL recruiting data from Ken's national sheet via NationalCode.gs ──
  // Returns the full campaigns dict. Caches in this._allCampaignsData.
  async _fetchRecruitingFromSheet(campaignKey) {
    // If we have cached data, return the specific campaign
    if (this._allCampaignsData && this._allCampaignsData[campaignKey]) {
      const cd = this._allCampaignsData[campaignKey];
      return { owners: cd.owners || [], weeks: cd.weeks || [], label: cd.label || '' };
    }

    const weeks = NATIONAL_CONFIG.campaigns[campaignKey]?.weeksToPull || 6;
    const url = NATIONAL_CONFIG.appsScriptUrl +
      '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
      '&action=recruiting&weeks=' + weeks +
      '&_t=' + Date.now();
    const resp = await fetch(url);
    const result = await resp.json();
    if (result.error) throw new Error(result.error);

    // Cache ALL campaigns data
    if (result.campaigns) {
      this._allCampaignsData = result.campaigns;
      // Dynamically populate campaign selector and config
      this._populateCampaignSelector(result.campaigns);
    }

    // Extract the campaign-specific data
    const campaignData = result.campaigns && result.campaigns[campaignKey];
    if (!campaignData) return null;

    return {
      owners: campaignData.owners || [],
      weeks: campaignData.weeks || [],
      label: campaignData.label || ''
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
    return { mapping: result.mapping || {}, costSheets: result.costSheets || {} };
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
      owner.indeedCosts = indeedOwners[matchKey];
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

      // Set current production from latest row
      owner.production.totalActual = nlr.current.productionLW || 0;
      owner.production.totalGoal = nlr.current.productionGoals || 0;

      // Build headcount history from ALL trend rows
      owner.headcountHistory = (nlr.trend || []).map(row => ({
        date: row.date,
        active: row.active,
        leaders: row.leaders,
        training: row.training
      }));

      // Build production history from ALL trend rows
      // Format: tA = total actual, tG = total goal, wA = wireless actual, wG = wireless goal
      // NLR only has total production (no wireless breakdown), so wireless = 0
      owner.productionHistory = (nlr.trend || []).map(row => ({
        date: row.date,
        tA: row.productionLW,
        tG: row.productionGoals,
        wA: 0,
        wG: 0
      }));

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

    // Campaign table: 4 most recent weeks, left-to-right chronological
    const campaignWeeks = allWeeks.slice(0, 4).reverse();
    const campaignLabels = campaignWeeks.map(w => w.tabName);

    // Owner detail: ALL weeks, left-to-right chronological
    const allWeeksChron = [...allWeeks].reverse();
    const allLabels = allWeeksChron.map(w => w.tabName);

    this.state.owners = ownerNames.map(name => {
      // Campaign-level actuals (4 weeks) for aggregation
      const actuals4 = Array.from({ length: 12 }, () => []);
      for (let wi = 0; wi < campaignWeeks.length; wi++) {
        const weekData = campaignWeeks[wi].data || {};
        const ownerVals = weekData[name] || new Array(12).fill(0);
        for (let ri = 0; ri < 12; ri++) {
          actuals4[ri].push(ownerVals[ri] || 0);
        }
      }

      // Full history actuals (all weeks) for owner detail tab
      const actualsFull = Array.from({ length: 12 }, () => []);
      for (let wi = 0; wi < allWeeksChron.length; wi++) {
        const weekData = allWeeksChron[wi].data || {};
        const ownerVals = weekData[name] || new Array(12).fill(0);
        for (let ri = 0; ri < 12; ri++) {
          actualsFull[ri].push(ownerVals[ri] || 0);
        }
      }

      return {
        name: name,
        tab: name,
        statusCode: null,
        headcount: { active: 0, leaders: 0, training: 0 },
        headcountHistory: [],
        production: { totalGoal: 0, totalActual: 0, wirelessGoal: 0, wirelessActual: 0 },
        productionHistory: [],
        nextGoals: { totalUnits: 0, wirelessUnits: 0 },
        recruiting: {
          leaders: 0,
          weeks: campaignLabels,
          rows: this._buildRows(0, actuals4)
        },
        recruitingFull: {
          leaders: 0,
          weeks: allLabels,
          rows: this._buildRows(0, actualsFull)
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

    // Campaign-level totals + aggregate recruiting (4 weeks only)
    this._buildCampaignAggregates(campaignLabels);
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

    // KPI totals
    const firstBookedIdx = 2;
    const newStartsIdx = 9;
    const startRetIdx = 11;

    const crRows = this.state.campaignRecruiting.rows;
    this.state.campaignTotals = {
      headcount: totals.headcount,
      firstBooked: crRows[firstBookedIdx] ? crRows[firstBookedIdx].total : 0,
      newStarts: crRows[newStartsIdx] ? crRows[newStartsIdx].total : 0,
      retention: crRows[startRetIdx] ? crRows[startRetIdx].total + '%' : '—',
      production: totals.production
    };
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

  // ── Scaffold data based on actual spreadsheet observations ──
  _loadScaffoldData(campaignKey) {
    const ownerDefs = NATIONAL_CONFIG.owners[campaignKey] || [];
    const weeks = ['Feb-9', 'Feb-16', 'Feb-23', 'Mar-2'];

    // Per-owner demo data (compact: health, status, recruiting projected/actuals, sales, audit)
    // h = headcount: active, leaders, training
    // p = production: totalGoal, totalActual, wirelessGoal, wirelessActual (goal = set last week, actual = from Tableau)
    // g = next week goals: totalUnits, wirelessUnits
    const demo = {
      'Jay T': {
        h: { active: 12, leaders: 3, training: 4 },
        hcHist: [
          { date: '2/17', active: 9, leaders: 2, training: 3 },
          { date: '2/24', active: 11, leaders: 3, training: 5 },
          { date: '3/3',  active: 12, leaders: 3, training: 4 }
        ],
        p: { totalGoal: 22, totalActual: 18, wirelessGoal: 15, wirelessActual: 12 },
        prodHist: [
          { date: '2/17', tA: 14, tG: 18, wA: 9, wG: 12 },
          { date: '2/24', tA: 16, tG: 20, wA: 11, wG: 14 },
          { date: '3/3',  tA: 18, tG: 22, wA: 12, wG: 15 }
        ],
        g: { totalUnits: 0, wirelessUnits: 0 },
        sc: 55,
        rLeaders: 3,

        rA: [[36,44,32,42],[28,35,25,32],[18,22,15,20],[14,18,12,17],[78,82,80,85],[47,52,42,50],[8,12,7,10],[7,10,5,8],[88,83,71,80],[4,6,3,5],[3,5,3,4],[75,83,100,80]],
        s: { totalSales: 42, newInternet: 18, upgrades: 14, videoSales: 10, abpMix: '72%', gigMix: '45%' },
        a: { reviews: 'A', website: 'B+', social: 'B', seo: 'A-' }
      },
      'Mason': {
        h: { active: 8, leaders: 2, training: 3 },
        hcHist: [
          { date: '2/17', active: 6, leaders: 1, training: 2 },
          { date: '2/24', active: 7, leaders: 2, training: 3 },
          { date: '3/3',  active: 8, leaders: 2, training: 3 }
        ],
        p: { totalGoal: 18, totalActual: 14, wirelessGoal: 12, wirelessActual: 9 },
        prodHist: [
          { date: '2/17', tA: 10, tG: 14, wA: 7, wG: 10 },
          { date: '2/24', tA: 12, tG: 16, wA: 8, wG: 11 },
          { date: '3/3',  tA: 14, tG: 18, wA: 9, wG: 12 }
        ],
        g: { totalUnits: 0, wirelessUnits: 0 },
        sc: 22,
        rLeaders: 2,

        rA: [[24,28,20,26],[18,22,15,20],[12,16,10,14],[9,12,7,11],[75,75,70,79],[40,53,33,47],[5,8,4,6],[4,6,3,5],[80,75,75,83],[3,4,2,3],[2,3,2,3],[67,75,100,100]],
        s: { totalSales: 31, newInternet: 14, upgrades: 10, videoSales: 7, abpMix: '68%', gigMix: '38%' },
        a: { reviews: 'B+', website: 'B', social: 'C+', seo: 'B' }
      },
      'Steven Sykes': {
        h: { active: 15, leaders: 4, training: 5 },
        hcHist: [
          { date: '2/17', active: 12, leaders: 3, training: 4 },
          { date: '2/24', active: 14, leaders: 4, training: 6 },
          { date: '3/3',  active: 15, leaders: 4, training: 5 }
        ],
        p: { totalGoal: 28, totalActual: 24, wirelessGoal: 20, wirelessActual: 18 },
        prodHist: [
          { date: '2/17', tA: 20, tG: 24, wA: 14, wG: 17 },
          { date: '2/24', tA: 22, tG: 26, wA: 16, wG: 19 },
          { date: '3/3',  tA: 24, tG: 28, wA: 18, wG: 20 }
        ],
        g: { totalUnits: 0, wirelessUnits: 0 },
        sc: 66,
        rLeaders: 4,

        rA: [[52,60,48,58],[40,48,36,44],[26,32,22,30],[22,26,18,24],[85,81,82,80],[47,53,46,52],[12,16,10,14],[10,14,8,12],[83,88,80,86],[6,8,5,7],[5,7,4,6],[83,88,80,86]],
        s: { totalSales: 56, newInternet: 24, upgrades: 18, videoSales: 14, abpMix: '75%', gigMix: '52%' },
        a: { reviews: 'A', website: 'A-', social: 'A', seo: 'A' }
      },
      'Olin Salter': {
        h: { active: 6, leaders: 1, training: 3 },
        hcHist: [
          { date: '2/17', active: 5, leaders: 1, training: 2 },
          { date: '2/24', active: 5, leaders: 1, training: 2 },
          { date: '3/3',  active: 6, leaders: 1, training: 3 }
        ],
        p: { totalGoal: 14, totalActual: 8, wirelessGoal: 10, wirelessActual: 5 },
        prodHist: [
          { date: '2/17', tA: 6, tG: 12, wA: 4, wG: 8 },
          { date: '2/24', tA: 7, tG: 13, wA: 4, wG: 9 },
          { date: '3/3',  tA: 8, tG: 14, wA: 5, wG: 10 }
        ],
        g: { totalUnits: 0, wirelessUnits: 0 },
        sc: 44,
        rLeaders: 1,

        rA: [[14,18,12,16],[10,14,8,12],[7,10,6,8],[5,7,4,6],[71,70,67,75],[35,50,30,40],[3,5,2,4],[2,4,2,3],[67,80,100,75],[1,3,1,2],[1,2,1,2],[100,67,100,100]],
        s: { totalSales: 18, newInternet: 8, upgrades: 6, videoSales: 4, abpMix: '62%', gigMix: '30%' },
        a: { reviews: 'C+', website: 'C', social: 'D+', seo: 'C' }
      },
      'Eric Martinez': {
        h: { active: 10, leaders: 2, training: 4 },
        hcHist: [
          { date: '2/17', active: 8, leaders: 2, training: 3 },
          { date: '2/24', active: 9, leaders: 2, training: 4 },
          { date: '3/3',  active: 10, leaders: 2, training: 4 }
        ],
        p: { totalGoal: 20, totalActual: 16, wirelessGoal: 14, wirelessActual: 11 },
        prodHist: [
          { date: '2/17', tA: 12, tG: 16, wA: 8, wG: 11 },
          { date: '2/24', tA: 14, tG: 18, wA: 10, wG: 13 },
          { date: '3/3',  tA: 16, tG: 20, wA: 11, wG: 14 }
        ],
        g: { totalUnits: 0, wirelessUnits: 0 },
        sc: 55,
        rLeaders: 2,

        rA: [[30,38,28,34],[24,30,20,26],[16,20,12,18],[12,16,10,14],[75,80,83,78],[46,53,43,50],[7,10,5,8],[6,8,4,7],[86,80,80,88],[4,5,3,4],[3,4,3,4],[75,80,100,100]],
        s: { totalSales: 36, newInternet: 16, upgrades: 12, videoSales: 8, abpMix: '70%', gigMix: '42%' },
        a: { reviews: 'B', website: 'B+', social: 'B-', seo: 'B+' }
      },
      'Natalia Gwarda': {
        h: { active: 9, leaders: 2, training: 4 },
        hcHist: [
          { date: '2/17', active: 7, leaders: 1, training: 3 },
          { date: '2/24', active: 8, leaders: 2, training: 4 },
          { date: '3/3',  active: 9, leaders: 2, training: 4 }
        ],
        p: { totalGoal: 16, totalActual: 12, wirelessGoal: 11, wirelessActual: 8 },
        prodHist: [
          { date: '2/17', tA: 9, tG: 13, wA: 6, wG: 9 },
          { date: '2/24', tA: 10, tG: 14, wA: 7, wG: 10 },
          { date: '3/3',  tA: 12, tG: 16, wA: 8, wG: 11 }
        ],
        g: { totalUnits: 0, wirelessUnits: 0 },
        sc: 33,
        rLeaders: 2,

        rA: [[26,34,24,30],[20,26,18,22],[14,18,10,16],[10,14,8,12],[71,78,80,75],[44,53,38,50],[6,9,4,7],[5,7,3,6],[83,78,75,86],[3,5,2,4],[2,4,2,3],[67,80,100,75]],
        s: { totalSales: 28, newInternet: 12, upgrades: 9, videoSales: 7, abpMix: '66%', gigMix: '36%' },
        a: { reviews: 'B-', website: 'C+', social: 'B', seo: 'C+' }
      },
      'Nigel Gilbert': {
        h: { active: 7, leaders: 1, training: 3 },
        hcHist: [
          { date: '2/17', active: 6, leaders: 1, training: 2 },
          { date: '2/24', active: 6, leaders: 1, training: 3 },
          { date: '3/3',  active: 7, leaders: 1, training: 3 }
        ],
        p: { totalGoal: 14, totalActual: 10, wirelessGoal: 10, wirelessActual: 6 },
        prodHist: [
          { date: '2/17', tA: 7, tG: 12, wA: 4, wG: 8 },
          { date: '2/24', tA: 8, tG: 13, wA: 5, wG: 9 },
          { date: '3/3',  tA: 10, tG: 14, wA: 6, wG: 10 }
        ],
        g: { totalUnits: 0, wirelessUnits: 0 },
        sc: 22,
        rLeaders: 1,

        rA: [[16,20,14,18],[12,16,10,14],[8,12,7,10],[6,9,5,8],[75,75,71,80],[36,50,32,42],[4,6,3,5],[3,5,2,4],[75,83,67,80],[2,3,1,3],[1,2,1,2],[50,67,100,67]],
        s: { totalSales: 22, newInternet: 10, upgrades: 7, videoSales: 5, abpMix: '64%', gigMix: '32%' },
        a: { reviews: 'C', website: 'C-', social: 'D', seo: 'C-' }
      }
    };

    // Build owner objects
    this.state.owners = ownerDefs.map(def => {
      const d = demo[def.name] || {};
      const h = d.h || {};
      const p = d.p || {};
      const g = d.g || {};
      return {
        name: def.name,
        tab: def.tab,
        statusCode: d.sc || null,
        // 1-on-1 Coaching: Headcount (editable during call)
        headcount: {
          active: h.active || 0,
          leaders: h.leaders || 0,
          training: h.training || 0
        },
        // Headcount history (week-over-week log, newest last)
        headcountHistory: d.hcHist ? d.hcHist.map(e => ({ ...e })) : [],
        // 1-on-1 Coaching: Production Review (goal set last week vs actual from Tableau)
        production: {
          totalGoal: p.totalGoal || 0,
          totalActual: p.totalActual || 0,
          wirelessGoal: p.wirelessGoal || 0,
          wirelessActual: p.wirelessActual || 0
        },
        // Production history (week-over-week log)
        productionHistory: d.prodHist ? d.prodHist.map(e => ({ ...e })) : [],
        // 1-on-1 Coaching: Next week goals (set during call)
        nextGoals: {
          totalUnits: g.totalUnits || 0,
          wirelessUnits: g.wirelessUnits || 0
        },
        // Recruiting (spreadsheet format)
        recruiting: {
          leaders: d.rLeaders || 0,
          weeks: weeks,
          rows: d.rA ? this._buildRows(d.rLeaders || 0, d.rA) : []
        },
        // Sales
        sales: {
          summary: d.s || null,
          reps: []
        },
        // Audit
        audit: {
          grades: d.a || { reviews: '—', website: '—', social: '—', seo: '—' },
          details: {}
        }
      };
    });

    // Build campaign-level aggregates (shared helper)
    this._buildCampaignAggregates(weeks);
  },

  // ══════════════════════════════════════════════════
  // CAMPAIGN SWITCHING
  // ══════════════════════════════════════════════════

  async switchCampaign(campaignKey) {
    this.state.campaign = campaignKey;
    this.state.selectedOwner = null;
    this._showLoading('Switching campaign...');
    try {
      // loadCampaignData will use cached _allCampaignsData if available
      await this.loadCampaignData(campaignKey);
    } catch (err) {
      console.error('Failed to load campaign:', err);
    }
    this._hideLoading();
    this.renderCampaignOverview();
    this.renderOwnersList();
    document.getElementById('owner-detail').style.display = 'none';
    document.querySelector('.campaign-overview').style.display = '';
    document.querySelector('.owners-section').style.display = '';
  },

  // ══════════════════════════════════════════════════
  // IMPORT LATEST RECRUITING
  // Copies latest tab from source tracker → Ken's sheet
  // ══════════════════════════════════════════════════

  async importLatestRecruiting() {
    if (!NATIONAL_CONFIG.appsScriptUrl) {
      alert('Apps Script URL not configured. Deploy NationalCode.gs first.');
      return;
    }

    const btn = document.getElementById('btn-import-recruiting');
    const status = document.getElementById('import-status');
    const weeksSelect = document.getElementById('import-weeks');
    const weeks = weeksSelect ? parseInt(weeksSelect.value, 10) : 1;

    if (btn) { btn.disabled = true; btn.textContent = 'Importing...'; }
    if (weeksSelect) weeksSelect.disabled = true;
    if (status) { status.textContent = ''; status.className = 'import-status'; }

    try {
      const resp = await this._fetchWithTimeout(
        fetch(NATIONAL_CONFIG.appsScriptUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain' },
          body: JSON.stringify({
            key: NATIONAL_CONFIG.apiKey,
            action: 'importRecruiting',
            weeks: weeks
          })
        }),
        60000 // 60s — import does heavy server-side work
      );

      const result = await resp.json();
      if (result.error) throw new Error(result.error);

      // Clear cached data so next load fetches fresh
      this._allCampaignsData = null;

      // Update state with fresh data returned by the import
      if (result.recruiting && result.recruiting.campaigns) {
        this._allCampaignsData = result.recruiting.campaigns; // re-cache with fresh data
        this._populateCampaignSelector(result.recruiting.campaigns);
        const campaignKey = this.state.campaign;
        const campaignData = result.recruiting.campaigns[campaignKey];
        if (campaignData) {
          this._buildOwnersFromSheet(campaignKey, {
            owners: campaignData.owners || [],
            weeks: campaignData.weeks || [],
            label: campaignData.label || ''
          });
        }

        // Re-enrich B2B owners with weekly data (headcount + production)
        // CPA/Indeed costs are monthly and don't change on weekly import — skip
        if (campaignKey === 'att-b2b' && NATIONAL_CONFIG.appsScriptUrl) {
          const [hcRes, prodRes] = await Promise.allSettled([
            this._fetchWithTimeout(this._fetchB2BHeadcount()),
            this._fetchWithTimeout(this._fetchB2BProduction())
          ]);
          if (hcRes.status === 'fulfilled' && hcRes.value?.owners && Object.keys(hcRes.value.owners).length) {
            this._enrichOwnersWithNLR(hcRes.value.owners);
          }
          if (prodRes.status === 'fulfilled' && prodRes.value?.owners && Object.keys(prodRes.value.owners).length) {
            this._enrichOwnersWithProduction(prodRes.value.owners);
          }
        }

        // Re-map audit data to owners if cached
        if (this._cachedAuditBusinesses && this._cachedAuditBusinesses.length) {
          this._mapAuditToOwners(this._cachedAuditBusinesses, this.state.camMapping);
        }
      }

      // Re-render dashboard
      this.renderCampaignOverview();
      this.renderOwnersList();

      // Show success
      const tabCount = result.imported?.tabCount || 1;
      const tabNames = result.imported?.tabNames || [result.imported?.tabName || 'latest'];
      const msg = tabCount === 1
        ? 'Imported: ' + tabNames[0]
        : 'Imported ' + tabCount + ' weeks (' + tabNames[0] + ' → ' + tabNames[tabNames.length - 1] + ')';
      if (status) {
        status.textContent = msg;
        status.className = 'import-status import-success';
        setTimeout(() => { status.textContent = ''; status.className = 'import-status'; }, 6000);
      }
      console.log('[NationalApp] Import successful:', result.imported);

    } catch (err) {
      console.error('[NationalApp] Import failed:', err);
      if (status) {
        status.textContent = 'Import failed: ' + err.message;
        status.className = 'import-status import-error';
      }
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = 'Import Recruiting'; }
      if (weeksSelect) weeksSelect.disabled = false;
    }
  },

  // (Bulk importCosts removed — costs now lazy-load per-owner when Recruiting tab opens)

  // ══════════════════════════════════════════════════
  // RENDER: Campaign Overview (KPIs + Recruiting Table + Status Codes)
  // ══════════════════════════════════════════════════

  renderCampaignOverview() {
    const t = this.state.campaignTotals || {};
    const cfg = NATIONAL_CONFIG.campaigns[this.state.campaign];
    document.getElementById('campaign-title').textContent = (cfg?.label || 'Campaign') + ' Campaign';
    document.getElementById('overview-date').textContent = 'Week of ' + this._formatCurrentWeek();

    // KPI cards
    document.getElementById('kpi-headcount').textContent = t.headcount || '—';
    document.getElementById('kpi-1st-booked').textContent = t.firstBooked || '—';
    document.getElementById('kpi-starts').textContent = t.newStarts || '—';
    document.getElementById('kpi-retention').textContent = t.retention || '—';
    document.getElementById('kpi-production').textContent = t.production || '—';

  },

  // ══════════════════════════════════════════════════
  // RENDER: Owners List (directory-style cards)
  // ══════════════════════════════════════════════════

  renderOwnersList() {
    const container = document.getElementById('owners-list');
    const owners = this.state.owners;

    if (!owners.length) {
      container.innerHTML = `
        <div class="empty-state" style="grid-column:1/-1">
          <div class="empty-state-icon">📊</div>
          <div class="empty-state-text">No owner data available yet.<br>Configure NationalCode.gs to load data.</div>
        </div>`;
      return;
    }

    container.innerHTML = owners.map((o, idx) => {
      return `
        <div class="owner-card" onclick="NationalApp.openOwnerDetail(${idx})">
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

    this.renderHealthTab(owner);
    this._showTab('health');
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

    // ── Section 1: Headcount Check (inputs start blank for weekly entry) ──
    const headcountEl = document.getElementById('health-headcount');
    headcountEl.innerHTML = `
      <div class="coaching-label">Headcount</div>
      <div class="hc-grid">
        <div class="hc-field">
          <label class="hc-field-label">Active Reps</label>
          <input type="number" class="hc-input" id="hc-active-${ownerIdx}" value="" min="0" placeholder="—"
            onchange="NationalApp._updateHeadcount(${ownerIdx}, 'active', this.value)">
        </div>
        <div class="hc-field">
          <label class="hc-field-label">Leaders</label>
          <input type="number" class="hc-input" id="hc-leaders-${ownerIdx}" value="" min="0" placeholder="—"
            onchange="NationalApp._updateHeadcount(${ownerIdx}, 'leaders', this.value)">
        </div>
        <div class="hc-field hc-field-calc">
          <label class="hc-field-label">Distributors</label>
          <div class="hc-value" id="hc-dist-${ownerIdx}">—</div>
          <div class="hc-calc-note">Active − Leaders</div>
        </div>
        <div class="hc-field">
          <label class="hc-field-label">In Training</label>
          <input type="number" class="hc-input" id="hc-training-${ownerIdx}" value="" min="0" placeholder="—"
            onchange="NationalApp._updateHeadcount(${ownerIdx}, 'training', this.value)">
        </div>
      </div>
      <div class="hc-submit-row">
        <button class="hc-submit-btn" id="hc-submit-${ownerIdx}"
          onclick="NationalApp._submitHeadcount(${ownerIdx})">Submit Headcount</button>
        <span class="hc-submit-note" id="hc-submit-note-${ownerIdx}"></span>
      </div>`;

    // ── Headcount Trend Table ──
    this._renderHeadcountTrend(owner, ownerIdx);

    // ── Section 2: Production Review (two side-by-side cards) ──
    const prodEl = document.getElementById('health-production');
    const isB2B = this.state.campaign === 'att-b2b';
    prodEl.innerHTML = `
      <div class="coaching-label">Production Review <span class="coaching-sublabel">Last Week</span></div>
      <div class="prod-cards">
        ${this._prodCard('Total Units', prod.totalActual, prod.totalGoal)}
        ${isB2B ? '' : this._prodCard('Wireless Lines', prod.wirelessActual, prod.wirelessGoal)}
      </div>`;

    // ── Production Trend Table ──
    this._renderProductionTrend(owner, ownerIdx);

    // ── Section 3: Set Goals (Next Week) ──
    const goalsEl = document.getElementById('health-goals');
    goalsEl.innerHTML = `
      <div class="coaching-label">Set Goals <span class="coaching-sublabel">Next Week</span></div>
      <div class="goals-grid">
        <div class="goal-field">
          <label class="goal-field-label">Total Units</label>
          <input type="number" class="goal-input" id="goal-total-${ownerIdx}" value="${goals.totalUnits || ''}" min="0"
            placeholder="—"
            onchange="NationalApp._updateGoal(${ownerIdx}, 'totalUnits', this.value)">
        </div>
        ${isB2B ? '' : `<div class="goal-field">
          <label class="goal-field-label">Wireless Units</label>
          <input type="number" class="goal-input" id="goal-wireless-${ownerIdx}" value="${goals.wirelessUnits || ''}" min="0"
            placeholder="—"
            onchange="NationalApp._updateGoal(${ownerIdx}, 'wirelessUnits', this.value)">
        </div>`}
      </div>
      <div class="hc-submit-row">
        <button class="hc-submit-btn" onclick="NationalApp._submitGoals(${ownerIdx})">Submit Goals</button>
        <span class="hc-submit-note" id="goal-submit-note-${ownerIdx}"></span>
      </div>`;

    // ── Section 4: Notes ──
    const notesEl = document.getElementById('health-notes');
    if (notesEl) {
      notesEl.innerHTML = `
        <div class="coaching-label">Notes</div>
        <textarea class="owner-notes" id="owner-notes-${ownerIdx}"
          placeholder="Add notes for this owner..."
          oninput="NationalApp._onNotesInput(${ownerIdx})">${this._esc(owner.notes || '')}</textarea>
        <div class="notes-footer">
          <span class="hc-submit-note" id="notes-save-status-${ownerIdx}"></span>
        </div>`;
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

  // ── Submit headcount — logs a new dated row ──
  _submitHeadcount(ownerIdx) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;

    // Read current values from inputs
    const active = parseInt(document.getElementById('hc-active-' + ownerIdx)?.value) || 0;
    const leaders = parseInt(document.getElementById('hc-leaders-' + ownerIdx)?.value) || 0;
    const training = parseInt(document.getElementById('hc-training-' + ownerIdx)?.value) || 0;

    if (!active && !leaders && !training) return; // Don't submit empty

    // Update state
    owner.headcount.active = active;
    owner.headcount.leaders = leaders;
    owner.headcount.training = training;

    const now = new Date();
    const dateStr = (now.getMonth() + 1) + '/' + now.getDate();

    // Push new entry
    owner.headcountHistory.push({
      date: dateStr,
      active, leaders, training
    });

    // Re-render the trend table
    this._renderHeadcountTrend(owner, ownerIdx);

    // Recalculate recruiting projected values based on new leader count
    if (owner.recruiting && owner.recruiting.rows.length) {
      owner.recruiting.leaders = leaders;
      const actuals = owner.recruiting.rows.map(r => r.values);
      owner.recruiting.rows = this._buildRows(leaders, actuals);
      // Re-render recruiting tab if it's currently visible
      if (this.state.currentTab === 'recruiting') {
        this.renderRecruitingTab(owner);
      }
    }

    // Show confirmation
    const note = document.getElementById('hc-submit-note-' + ownerIdx);
    if (note) {
      note.textContent = 'Submitted ' + dateStr;
      note.classList.add('show');
      setTimeout(() => note.classList.remove('show'), 3000);
    }
  },

  // ── Render headcount week-over-week bar chart (newest-first, responsive) ──
  _renderHeadcountTrend(owner, ownerIdx) {
    const trendEl = document.getElementById('health-hc-trend');
    if (!trendEl) return;

    const rawHist = owner.headcountHistory || [];
    if (!rawHist.length) { trendEl.style.display = 'none'; return; }
    trendEl.style.display = '';

    // Reverse so newest week is on the LEFT
    const hist = [...rawHist].reverse();
    const n = hist.length;

    const shortDate = (d) => {
      if (!d) return '';
      const parts = String(d).split('/');
      return parts.length >= 2 ? parts[0] + '/' + parts[1] : d;
    };

    // ── Layout ──
    // Single SVG with viewBox. Fills container width naturally.
    // If many bars, becomes scrollable with a minimum bar width.
    const PAD_L = 36, PAD_R = 10, PAD_T = 14, PAD_B = 28;
    const BAR_R = 4;
    const GAP = 0.18;
    const MIN_SLOT = 56;        // bars never narrower than this
    const IDEAL_SLOT = 72;      // ideal bar slot width at comfortable density
    const MAX_FIT = 14;         // bars that fit before scrolling kicks in

    // Determine slot width: fill the space when ≤MAX_FIT bars, fixed when more
    // Use a virtual 700px reference width for the viewBox
    const REF_W = 700;
    const plotSpace = REF_W - PAD_L - PAD_R;
    let slotW;
    if (n <= MAX_FIT) {
      slotW = Math.min(IDEAL_SLOT, plotSpace / n);
    } else {
      slotW = MIN_SLOT;
    }

    const svgW = PAD_L + n * slotW + PAD_R;
    const svgH = 220;
    const plotH = svgH - PAD_T - PAD_B;
    const needsScroll = svgW > REF_W;

    // ── Y-axis scale ──
    const maxVal = Math.max(...hist.map(r => (r.active || 0) + (r.training || 0)), 1);
    const yMax = Math.ceil(maxVal / 5) * 5 || 5;
    const yScale = plotH / yMax;
    const baseY = PAD_T + plotH;

    // ── Bar geometry ──
    const barW = slotW * (1 - GAP);
    const barOff = (slotW - barW) / 2;

    const roundTop = (x, y, w, h, r) => {
      if (h <= 0) return '';
      const cr = Math.min(r, h / 2, w / 2);
      return `M${x},${y + h}L${x},${y + cr}Q${x},${y} ${x + cr},${y}` +
             `L${x + w - cr},${y}Q${x + w},${y} ${x + w},${y + cr}L${x + w},${y + h}Z`;
    };

    let svg = '';

    // ── Gridlines + Y-axis labels ──
    const ticks = 4;
    for (let t = 0; t <= ticks; t++) {
      const val = Math.round(yMax * t / ticks);
      const y = baseY - val * yScale;
      svg += `<line x1="${PAD_L}" y1="${y}" x2="${svgW - PAD_R}" y2="${y}" stroke="#e8ecf1" stroke-width="0.7"/>`;
      svg += `<text x="${PAD_L - 6}" y="${y + 3.5}" text-anchor="end" fill="#b0b8c4" font-size="10" font-family="Inter,sans-serif">${val}</text>`;
    }
    svg += `<text x="${PAD_L - 6}" y="${baseY + 3.5}" text-anchor="end" fill="#b0b8c4" font-size="10" font-family="Inter,sans-serif">0</text>`;

    // ── Bars ──
    hist.forEach((r, i) => {
      const origIdx = rawHist.length - 1 - i;
      const active = r.active || 0;
      const leaders = r.leaders || 0;
      const training = r.training || 0;
      const dist = active - leaders;
      const leaderH = leaders * yScale;
      const distH = Math.max(dist, 0) * yScale;
      const solidH = active * yScale;
      const trainH = training * yScale;
      const x = PAD_L + i * slotW + barOff;
      const solidTop = baseY - solidH;
      const trainTop = solidTop - trainH;

      const topmost = trainH > 0 ? 'training' : (distH > 0 ? 'dist' : 'leader');

      // Leaders (bottom — blue)
      if (leaderH > 0) {
        if (topmost === 'leader') {
          svg += `<path d="${roundTop(x, baseY - leaderH, barW, leaderH, BAR_R)}" fill="#5b9cf6" opacity="0.9"/>`;
        } else {
          svg += `<rect x="${x}" y="${baseY - leaderH}" width="${barW}" height="${leaderH}" fill="#5b9cf6" opacity="0.9"/>`;
        }
      }

      // Distributors (middle — teal)
      if (distH > 0) {
        if (topmost === 'dist') {
          svg += `<path d="${roundTop(x, solidTop, barW, distH, BAR_R)}" fill="#0ea5a0" opacity="0.85"/>`;
        } else {
          svg += `<rect x="${x}" y="${solidTop}" width="${barW}" height="${distH}" fill="#0ea5a0" opacity="0.85"/>`;
        }
      }

      // Training extension (dashed purple)
      if (trainH > 0) {
        svg += `<path d="${roundTop(x + 0.5, trainTop + 0.5, barW - 1, trainH - 1, BAR_R)}" fill="rgba(139,92,246,0.06)" stroke="#a78bfa" stroke-width="1" stroke-dasharray="3 2"/>`;
      }

      // Hover target
      const totalH = solidH + trainH;
      const topY = trainH > 0 ? trainTop : solidTop;
      svg += `<rect x="${x}" y="${Math.min(topY, baseY - 1)}" width="${barW}" height="${Math.max(totalH, 4)}" fill="transparent" style="cursor:pointer" onmouseenter="NationalApp._showHcTooltip(event,${origIdx},${ownerIdx})" onmouseleave="NationalApp._hideHcTooltip()"/>`;

      // X-axis date label
      svg += `<text x="${x + barW / 2}" y="${svgH - 8}" text-anchor="middle" fill="#8a95a5" font-size="10" font-weight="600" font-family="Inter,sans-serif">${shortDate(r.date)}</text>`;
    });

    // ── Assemble ──
    // When bars fit: SVG fills container via viewBox (responsive).
    // When too many: fixed-width SVG inside scrollable wrapper.
    const svgAttrs = needsScroll
      ? `viewBox="0 0 ${svgW} ${svgH}" width="${svgW}" height="${svgH}"`
      : `viewBox="0 0 ${svgW} ${svgH}" preserveAspectRatio="xMidYMid meet"`;

    trendEl.innerHTML = `
      <div class="coaching-label">Week-over-Week Headcount</div>
      <div class="hc-chart-wrap${needsScroll ? ' hc-chart-scrollable' : ''}">
        <svg ${svgAttrs}>${svg}</svg>
        <div class="hc-chart-tooltip" id="hc-chart-tt"></div>
      </div>
      <div class="hc-chart-legend">
        <span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch swatch-leaders"></span>Leaders</span>
        <span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch swatch-dist"></span>Distributors</span>
        <span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch swatch-training"></span>Training</span>
      </div>`;
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

  // ── Render production week-over-week trend table ──
  _renderProductionTrend(owner, ownerIdx) {
    const trendEl = document.getElementById('health-prod-trend');
    if (!trendEl) return;

    const hist = owner.productionHistory || [];
    if (!hist.length) {
      trendEl.style.display = 'none';
      return;
    }
    trendEl.style.display = '';
    const isB2B = this.state.campaign === 'att-b2b';

    trendEl.innerHTML = `
      <div class="coaching-label">Week-over-Week Production</div>
      <div class="data-table-wrap trend-scroll" id="prod-trend-scroll">
        <table class="data-table prod-trend-table">
          <thead>
            <tr>
              <th>Week</th>
              <th class="num">Total Units</th>
              ${isB2B ? '' : '<th class="num">Wireless Lines</th>'}
            </tr>
          </thead>
          <tbody>
            ${hist.map((r, i) => {
              const prev = i > 0 ? hist[i - 1] : null;
              return `
                <tr>
                  <td class="bold">${this._esc(r.date)}</td>
                  <td class="num"><span class="prod-trend-actual">${r.tA}</span><span class="prod-trend-suffix"><span class="prod-trend-goal"> of ${r.tG}</span> ${this._trendArrow(r.tA, prev?.tA)}</span></td>
                  ${isB2B ? '' : `<td class="num"><span class="prod-trend-actual">${r.wA}</span><span class="prod-trend-suffix"><span class="prod-trend-goal"> of ${r.wG}</span> ${this._trendArrow(r.wA, prev?.wA)}</span></td>`}
                </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>`;
    // Auto-scroll to bottom so most recent 5 weeks are visible
    requestAnimationFrame(() => {
      const el = document.getElementById('prod-trend-scroll');
      if (el) el.scrollTop = el.scrollHeight;
    });
  },

  // ── Goal input handler ──
  _updateGoal(ownerIdx, field, value) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    owner.nextGoals[field] = parseInt(value) || 0;
  },

  // ── Submit goals ──
  _submitGoals(ownerIdx) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;

    const total = parseInt(document.getElementById('goal-total-' + ownerIdx)?.value) || 0;
    const wirelessEl = document.getElementById('goal-wireless-' + ownerIdx);
    const wireless = wirelessEl ? (parseInt(wirelessEl.value) || 0) : 0;

    if (!total && !wireless) return;

    owner.nextGoals.totalUnits = total;
    owner.nextGoals.wirelessUnits = wireless;

    const note = document.getElementById('goal-submit-note-' + ownerIdx);
    if (note) {
      note.textContent = 'Goals saved ✓';
      note.classList.add('show');
      setTimeout(() => note.classList.remove('show'), 3000);
    }
  },

  // ── Notes: debounced auto-save ──
  _notesTimer: null,

  _onNotesInput(ownerIdx) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    const ta = document.getElementById('owner-notes-' + ownerIdx);
    if (!ta) return;
    owner.notes = ta.value;

    // Debounce: save 1.5s after last keystroke
    if (this._notesTimer) clearTimeout(this._notesTimer);
    this._notesTimer = setTimeout(() => this._saveNotes(ownerIdx), 1500);
  },

  async _saveNotes(ownerIdx) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    const status = document.getElementById('notes-save-status-' + ownerIdx);

    if (!NATIONAL_CONFIG.appsScriptUrl) {
      // No backend — just save locally
      if (status) { status.textContent = 'Saved locally'; status.classList.add('show'); setTimeout(() => status.classList.remove('show'), 2000); }
      return;
    }

    try {
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'saveOwnerNotes',
          ownerName: owner.name,
          campaign: this.state.campaign,
          notes: owner.notes
        })
      });
      const result = await resp.json();
      if (result.error) throw new Error(result.error);
      if (status) { status.textContent = 'Saved'; status.classList.add('show'); setTimeout(() => status.classList.remove('show'), 2000); }
    } catch (err) {
      console.warn('[NationalApp] Notes save failed:', err);
      if (status) { status.textContent = 'Save failed'; status.classList.add('show'); setTimeout(() => status.classList.remove('show'), 3000); }
    }
  },

  // ── Production card (colored card, big actual / small goal) ──
  _prodCard(label, actual, goal) {
    const pct = goal ? Math.round((actual / goal) * 100) : 0;
    return `
      <div class="prod-card ${this._pctClass(pct)}">
        <div class="prod-card-label">${label}</div>
        <div class="prod-card-actual">${actual}</div>
        <div class="prod-card-goal">goal ${goal}</div>
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

  // ── Lazy-load recruiting costs for a single owner via claimed costSheetId ──
  async _loadAndRenderCosts(owner) {
    const el = document.getElementById('owner-indeed-costs');
    if (!el) return;

    const costSheetId = this.state.costSheets[owner.name];

    // No cost sheet claimed — show claim UI
    if (!costSheetId) {
      this._renderCostClaimSection(owner);
      return;
    }

    // Already loaded — just render
    if (owner.indeedCosts) {
      this._renderIndeedCosts(owner);
      return;
    }

    // Already fetching — don't double-fetch
    if (owner._costsFetching) return;

    // Show loading indicator
    el.innerHTML = `<div class="coaching-label">Recruiting Costs</div>
      <div class="empty-state"><div class="loading-spinner" style="width:24px;height:24px;border-width:3px;margin:0 auto"></div>
      <div class="empty-state-text" style="margin-top:8px">Loading cost data...</div></div>`;

    owner._costsFetching = true;
    try {
      const url = NATIONAL_CONFIG.appsScriptUrl +
        '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
        '&action=readCostSheet' +
        '&sheetId=' + encodeURIComponent(costSheetId) +
        '&_t=' + Date.now();
      const resp = await this._fetchWithTimeout(fetch(url), 30000);
      const result = await resp.json();
      if (result.error) throw new Error(result.error);

      if (result.months && result.months.length) {
        owner.indeedCosts = result;
      }

      // Render (will show empty state if no data)
      if (this.state.selectedOwner === owner && this.state.currentTab === 'recruiting') {
        this._renderIndeedCosts(owner);
      }
    } catch (err) {
      console.warn('[NationalApp] Cost load for', owner.name, 'failed:', err.message);
      if (this.state.selectedOwner === owner && this.state.currentTab === 'recruiting') {
        el.innerHTML = `<div class="coaching-label">Recruiting Costs</div>
          <div class="empty-state"><div class="empty-state-text">Failed to load cost data</div></div>`;
      }
    } finally {
      owner._costsFetching = false;
    }
  },

  // ── Render cost sheet claim section (no spreadsheet assigned yet) ──
  async _renderCostClaimSection(owner) {
    const el = document.getElementById('owner-indeed-costs');
    if (!el) return;

    // Fetch the list of available files if not cached
    if (!this.state.costFilesList) {
      el.innerHTML = `<div class="coaching-label">Recruiting Costs</div>
        <div class="empty-state"><div class="loading-spinner" style="width:24px;height:24px;border-width:3px;margin:0 auto"></div>
        <div class="empty-state-text" style="margin-top:8px">Loading available spreadsheets...</div></div>`;
      try {
        const url = NATIONAL_CONFIG.appsScriptUrl +
          '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
          '&action=listCostFiles' +
          '&_t=' + Date.now();
        const resp = await this._fetchWithTimeout(fetch(url), 20000);
        const result = await resp.json();
        if (result.error) throw new Error(result.error);
        this.state.costFilesList = result.files || [];
      } catch (err) {
        console.warn('[NationalApp] Failed to list cost files:', err.message);
        el.innerHTML = '';
        return;
      }
    }

    // Build set of already-claimed sheet IDs
    const claimedIds = new Set(Object.values(this.state.costSheets));
    const available = this.state.costFilesList.filter(f => !claimedIds.has(f.id));

    const optionsHTML = available.map(f =>
      `<div class="claim-dropdown-item" onclick="NationalApp._claimCostSheet('${this._esc(f.id)}')">
        ${this._esc(f.name)}
      </div>`
    ).join('');

    el.innerHTML = `
      <div class="claim-section">
        <div class="claim-header">
          <div class="section-label" style="margin:0">Recruiting Costs</div>
          <span class="claim-count">${available.length} spreadsheet${available.length !== 1 ? 's' : ''} available</span>
        </div>
        <div class="claim-search-wrap">
          <input type="text" class="claim-search-input" id="cost-claim-search"
            placeholder="Search spreadsheets to claim..."
            oninput="NationalApp._filterCostClaimDropdown(this.value)"
            onfocus="NationalApp._showCostClaimDropdown()"
            autocomplete="off">
          <div class="claim-dropdown" id="cost-claim-dropdown" style="display:none">
            ${optionsHTML || '<div class="claim-dropdown-empty">No unclaimed spreadsheets available</div>'}
          </div>
        </div>
        <div class="claim-status" id="cost-claim-status"></div>
      </div>`;
  },

  _showCostClaimDropdown() {
    const dd = document.getElementById('cost-claim-dropdown');
    if (dd) dd.style.display = 'block';
    setTimeout(() => {
      const handler = (e) => {
        if (!e.target.closest('.claim-search-wrap')) {
          dd.style.display = 'none';
          document.removeEventListener('click', handler);
        }
      };
      document.addEventListener('click', handler);
    }, 10);
  },

  _filterCostClaimDropdown(query) {
    const dd = document.getElementById('cost-claim-dropdown');
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
    const empty = dd.querySelector('.claim-dropdown-empty');
    if (empty) empty.style.display = visible === 0 ? 'block' : 'none';
  },

  async _claimCostSheet(sheetId) {
    const owner = this.state.selectedOwner;
    if (!owner) return;

    const status = document.getElementById('cost-claim-status');
    if (status) { status.textContent = 'Claiming...'; status.className = 'claim-status'; }

    try {
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'claimCostSheet',
          ownerName: owner.name,
          sheetId: sheetId
        })
      });
      const result = await resp.json();
      if (result.error) throw new Error(result.error);

      // Update state with new mapping
      if (result.costSheets) this.state.costSheets = result.costSheets;
      if (result.mapping) this.state.camMapping = result.mapping;

      // Now load the cost data
      this._loadAndRenderCosts(owner);
    } catch (err) {
      console.error('[NationalApp] Claim cost sheet failed:', err);
      if (status) { status.textContent = 'Claim failed: ' + err.message; status.className = 'claim-status claim-error'; }
    }
  },

  async _unclaimCostSheet() {
    const owner = this.state.selectedOwner;
    if (!owner) return;

    try {
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'unclaimCostSheet',
          ownerName: owner.name
        })
      });
      const result = await resp.json();
      if (result.error) throw new Error(result.error);

      if (result.costSheets) this.state.costSheets = result.costSheets;
      if (result.mapping) this.state.camMapping = result.mapping;

      // Clear cached cost data and re-render
      owner.indeedCosts = null;
      this._renderCostClaimSection(owner);
    } catch (err) {
      console.error('[NationalApp] Unclaim cost sheet failed:', err);
    }
  },

  // ── Week-over-week raw recruiting data (stacked cards under projected table) ──
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

    // Stacked cards, newest first (reverse)
    let html = `<div class="coaching-label">Week-over-Week Recruiting</div>
      <div class="wow-cards">`;

    [...activeWeekIdxs].reverse().forEach(wi => {
      const prevWi = activeWeekIdxs[activeWeekIdxs.indexOf(wi) - 1] ?? null;
      html += `<div class="wow-card">
        <div class="wow-card-header">${this._esc(weeks[wi])}</div>
        <div class="data-table-wrap"><table class="data-table">
          <tbody>`;

      labels.forEach((def, ri) => {
        const val = rows[ri]?.values?.[wi] ?? 0;
        const prev = prevWi !== null ? (rows[ri]?.values?.[prevWi] ?? null) : null;
        const display = def.isRate ? val + '%' : val;
        const arrow = prev !== null ? this._trendArrow(val, prev) : '';
        html += `<tr>
          <td>${this._esc(def.label)}</td>
          <td class="num">${display} ${arrow}</td>
        </tr>`;
      });

      html += `</tbody></table></div></div>`;
    });

    html += `</div>`;
    el.innerHTML = html;
  },

  // ══════════════════════════════════════════════════
  // RENDER: Recruiting Costs (inside Recruiting tab)
  // ══════════════════════════════════════════════════

  // Row definitions for cost tables
  _COST_ROWS: [
    { label: '# of Ads',        key: 'numAds',         fmt: 'num' },
    { label: 'Total Cost',      key: 'totalCost',      fmt: 'dollar' },
    { label: '# of Applies',    key: 'numApplies',     fmt: 'num' },
    { label: 'Cost/Apply',      key: 'costPerApply',   fmt: 'dollar' },
    { label: '# of 2nds',       key: 'num2nds',        fmt: 'num' },
    { label: 'Cost/2nd',        key: 'costPer2nd',     fmt: 'dollar' },
    { label: '# of New Starts', key: 'numNewStarts',   fmt: 'num' },
    { label: 'Cost/New Start',  key: 'costPerNewStart', fmt: 'dollar' }
  ],

  _renderIndeedCosts(owner) {
    const el = document.getElementById('owner-indeed-costs');
    if (!el) return;

    const data = owner.indeedCosts;
    if (!data || !data.months || !data.months.length) {
      el.innerHTML = '';
      return;
    }

    const months = data.months;
    const hasMultiple = months.length > 1;
    let html = '';

    // ── Part 1: Platform Breakdown for selected month ──
    html += `<div class="coaching-label">
      ${hasMultiple ? 'Platform Breakdown' : 'Recruiting Costs'}
      <button class="btn-unclaim-cost" onclick="NationalApp._unclaimCostSheet()" title="Change spreadsheet">&times;</button>
      ${hasMultiple
        ? `<select class="rc-month-select" onchange="NationalApp._switchCostMonth(this.value)">
            ${months.map((m, i) => `<option value="${i}">${this._esc(m.month)}</option>`).join('')}
          </select>`
        : `<span class="coaching-sublabel">${this._esc(months[0].month)}</span>`
      }
    </div>`;
    html += this._buildCostTable(months[0]);

    // ── Part 2: Month-over-Month Trend Table (TOTAL values across months) ──
    if (hasMultiple) {
      html += `<div class="coaching-label rc-detail-label">Month-over-Month</div>`;
      html += this._buildCostTrend(months);
    }

    el.innerHTML = html;
  },

  _switchCostMonth(idx) {
    const owner = this.state.selectedOwner;
    if (!owner || !owner.indeedCosts || !owner.indeedCosts.months) return;
    const month = owner.indeedCosts.months[parseInt(idx)];
    if (!month) return;

    const wrap = document.querySelector('#owner-indeed-costs .rc-table-wrap');
    if (wrap) {
      wrap.outerHTML = this._buildCostTable(month);
    }
  },

  // ── Month-over-Month trend: columns = months, rows = key metrics (TOTAL only) ──
  _buildCostTrend(months) {
    // Key metrics for the trend view (subset of full rows)
    const trendRows = [
      { label: 'Total Cost',      key: 'totalCost',      fmt: 'dollar' },
      { label: '# of Applies',    key: 'numApplies',     fmt: 'num' },
      { label: 'Cost/Apply',      key: 'costPerApply',   fmt: 'dollar' },
      { label: '# of 2nds',       key: 'num2nds',        fmt: 'num' },
      { label: 'Cost/2nd',        key: 'costPer2nd',     fmt: 'dollar' },
      { label: '# of New Starts', key: 'numNewStarts',   fmt: 'num' },
      { label: 'Cost/New Start',  key: 'costPerNewStart', fmt: 'dollar' }
    ];

    let h = `<div class="rc-table-wrap rc-trend"><div class="data-table-wrap"><table class="data-table">
      <thead><tr>
        <th></th>
        ${months.map(m => `<th class="num">${this._esc(m.month)}</th>`).join('')}
      </tr></thead><tbody>`;

    for (const r of trendRows) {
      const fmt = r.fmt === 'dollar' ? this._fmtDollar : this._fmtNum;
      h += `<tr><td class="rc-label">${r.label}</td>`;

      for (let mi = 0; mi < months.length; mi++) {
        const val = (months[mi].total || {})[r.key] ?? 0;
        const prev = mi < months.length - 1 ? ((months[mi + 1].total || {})[r.key] ?? null) : null;
        // For cost metrics, lower is better (invert arrow)
        const isCost = r.key.startsWith('cost');
        const arrow = prev !== null ? this._costTrendArrow(val, prev, isCost) : '';
        h += `<td class="num">${fmt(val)} ${arrow}</td>`;
      }

      h += `</tr>`;
    }

    h += `</tbody></table></div></div>`;
    return h;
  },

  // Trend arrow for cost metrics (lower = good for cost rows, higher = good for count rows)
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

  // ── Platform breakdown table for a single month ──
  _buildCostTable(monthData) {
    const platforms = monthData.platforms || {};
    const total = monthData.total || {};
    const platformOrder = monthData.platformOrder || Object.keys(platforms);

    // Fallback for old data format (local/nlr/total only, no platforms key)
    if (!platformOrder.length && (monthData.local || monthData.nlr)) {
      const fallbackPlatforms = [];
      if (monthData.local && Object.keys(monthData.local).length) fallbackPlatforms.push('Local Indeed');
      if (monthData.nlr && Object.keys(monthData.nlr).length) fallbackPlatforms.push('NLR Indeed');
      const fallbackData = {};
      if (monthData.local) fallbackData['Local Indeed'] = monthData.local;
      if (monthData.nlr) fallbackData['NLR Indeed'] = monthData.nlr;
      return this._buildCostTableInner(fallbackPlatforms, fallbackData, total);
    }

    return this._buildCostTableInner(platformOrder, platforms, total);
  },

  _buildCostTableInner(platformOrder, platforms, total) {
    let h = `<div class="rc-table-wrap"><div class="data-table-wrap"><table class="data-table rc-platform-table">
      <thead><tr>
        <th></th>
        ${platformOrder.map(p => `<th class="num">${this._esc(p)}</th>`).join('')}
        <th class="num rc-total-col">TOTAL</th>
      </tr></thead><tbody>`;

    for (const r of this._COST_ROWS) {
      const fmt = r.fmt === 'dollar' ? this._fmtDollar : this._fmtNum;
      h += `<tr><td class="rc-label">${r.label}</td>`;

      for (const pName of platformOrder) {
        const val = (platforms[pName] || {})[r.key] ?? 0;
        h += `<td class="num">${fmt(val)}</td>`;
      }

      const tv = total[r.key] ?? 0;
      h += `<td class="num rc-total-val">${fmt(tv)}</td>`;
      h += `</tr>`;
    }

    h += `</tbody></table></div></div>`;
    return h;
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
    const s = owner.sales;
    const sm = s.summary;

    // ── Card 1: Owner Summary ──
    const summaryEl = document.getElementById('sales-summary');
    if (sm) {
      summaryEl.innerHTML = `
        <div class="coaching-section">
          <div class="coaching-label">Owner Overview</div>
          <div class="sales-kpi-grid">
            ${[
              { label: 'Total Volume', value: sm.totalVolume, cls: 'big' },
              { label: 'Rep Count', value: sm.repCount },
              { label: 'Sales / Rep', value: sm.salesPerRep },
              { label: 'Order Count', value: sm.orderCount }
            ].map(k => `
              <div class="health-kpi${k.cls ? ' ' + k.cls : ''}">
                <div class="health-kpi-value">${k.value}</div>
                <div class="health-kpi-label">${k.label}</div>
              </div>
            `).join('')}
          </div>
          <div class="sales-metrics-grid">
            <div class="sales-metric-group">
              <div class="sales-metric-group-label">Sales Breakdown</div>
              <div class="sales-metric-row"><span>Internet</span><span class="num">${sm.internet}</span></div>
              <div class="sales-metric-row"><span>VOIP</span><span class="num">${sm.voip}</span></div>
              <div class="sales-metric-row"><span>Wireless</span><span class="num">${sm.wireless}</span></div>
              <div class="sales-metric-row"><span>AIR/AWB</span><span class="num">${sm.airAwb}</span></div>
            </div>
            <div class="sales-metric-group">
              <div class="sales-metric-group-label">Order Timing</div>
              <div class="sales-metric-row"><span>Before 12 PM</span><span class="num">${sm.ordersBefore} <small>(${this._pct(sm.earlyPct)})</small></span></div>
              <div class="sales-metric-row"><span>After 5 PM</span><span class="num">${sm.ordersAfter} <small>(${this._pct(sm.latePct)})</small></span></div>
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
          </div>
        </div>`;
    } else {
      summaryEl.innerHTML = `
        <div class="coaching-section">
          <div class="coaching-label">Owner Overview</div>
          <div class="empty-state">
            <div class="empty-state-text">No production data yet. Click "Import Recruiting" to pull sales data from NLR.</div>
          </div>
        </div>`;
    }

    // ── Card 2: Rep Breakdown Table ──
    const repsEl = document.getElementById('sales-reps-table');
    if (s.reps.length) {
      repsEl.innerHTML = `
        <div class="coaching-section">
          <div class="coaching-label">Rep Breakdown <span class="coaching-sublabel">${s.reps.length} reps</span></div>
          <div class="data-table-wrap">
            <table class="data-table">
              <thead>
                <tr>
                  <th>Rep Name</th>
                  <th class="num">Volume</th>
                  <th class="num">Orders</th>
                  <th class="num">Sales/Rep</th>
                  <th class="num">Internet</th>
                  <th class="num">VOIP</th>
                  <th class="num">Wireless</th>
                  <th class="num">AIR/AWB</th>
                  <th class="num">Early %</th>
                  <th class="num">Late %</th>
                  <th class="num">ABP %</th>
                  <th class="num">CRU %</th>
                  <th class="num">New Wrls %</th>
                  <th class="num">BYOD %</th>
                </tr>
              </thead>
              <tbody>
                ${s.reps.map(rep => `
                  <tr>
                    <td class="bold">${this._esc(rep.name)}</td>
                    <td class="num">${rep.totalVolume}</td>
                    <td class="num">${rep.orderCount}</td>
                    <td class="num">${rep.salesPerRep}</td>
                    <td class="num">${rep.internet}</td>
                    <td class="num">${rep.voip}</td>
                    <td class="num">${rep.wireless}</td>
                    <td class="num">${rep.airAwb}</td>
                    <td class="num">${this._pct(rep.earlyPct)}</td>
                    <td class="num">${this._pct(rep.latePct)}</td>
                    <td class="num">${this._pct(rep.abpPct)}</td>
                    <td class="num">${this._pct(rep.cruPct)}</td>
                    <td class="num">${this._pct(rep.newWrlsPct)}</td>
                    <td class="num">${this._pct(rep.byodPct)}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>`;
    } else {
      repsEl.innerHTML = `
        <div class="coaching-section">
          <div class="coaching-label">Rep Breakdown</div>
          <div class="empty-state">
            <div class="empty-state-text">No rep data yet. Click "Import Recruiting" to pull sales data from NLR.</div>
          </div>
        </div>`;
    }
  },

  // ══════════════════════════════════════════════════
  // RENDER: Audit Tab (Online Presence)
  // ══════════════════════════════════════════════════

  renderAuditTab(owner) {
    const a = owner.audit;
    const bizList = a.businesses || [];
    const total = bizList.length;

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

    // ── Reports then Claim Section (below) ──
    const details = document.getElementById('audit-details');
    const claimHTML = this._renderClaimSection(owner);

    if (!bizList.length) {
      details.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-icon">🏢</div>
          <div class="empty-state-text">No companies claimed for this owner yet.<br>
          Use the search below to claim companies from Cam's report.</div>
        </div>` + claimHTML;
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
        </div>
        ${claimHTML}`;
      this._currentBizList = bizList;
      this._currentAuditMonth = auditMonth;
    } else {
      details.innerHTML = `
        <div class="bis-reports">
          ${this._renderBizReport(bizList[0], auditMonth)}
        </div>
        ${claimHTML}`;
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

      // Update mapping + costSheets in state
      this.state.camMapping = result.mapping || {};
      if (result.costSheets) this.state.costSheets = result.costSheets;

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

      // Update mapping + costSheets in state
      this.state.camMapping = result.mapping || {};
      if (result.costSheets) this.state.costSheets = result.costSheets;

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
          <div class="bis-section-banner">Website${ws?.url ? ` <a href="${extUrl(ws.url)}" target="_blank" rel="noopener" class="bis-link">↗</a>` : ''}</div>
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
          <div class="bis-section-banner">Social Media${ig.link ? ` <a href="${extUrl(ig.link)}" target="_blank" rel="noopener" class="bis-link">↗</a>` : ''}</div>
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

    const weeks = data.weeks || [];
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

    // Date header row
    html += `<tr class="rt-date-row">
      <th></th>
      <th></th>
      ${weeks.map(w => `<th>${this._esc(w)}</th>`).join('')}
      <th></th>
    </tr>`;

    html += `</thead><tbody>`;

    // Data rows
    rows.forEach((row, ri) => {
      html += `<tr>`;
      html += `<td>${this._esc(row.label)}</td>`;
      html += `<td class="rt-projected">${this._fmtCell(row.projected, row.isRate)}</td>`;

      // Weekly values with conditional coloring
      row.values.forEach(val => {
        const color = this._cellColor(val, row.projected, row.isRate, ri);
        html += `<td class="${color}">${this._fmtCell(val, row.isRate)}</td>`;
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
      { label: 'Bored Leaders',                mult: 2, color: '#fecaca' },
      { label: 'Top Leaders Interviewing Only', mult: 3, color: '#fed7aa' },
      { label: 'Maintaining, Not Growing',      mult: 4, color: '#fef08a' },
      { label: 'Leaders Busy',                  mult: 5, color: '#bbf7d0' },
      { label: 'Promotion Factory',             mult: 6, color: '#bfdbfe' }
    ];
    let h = `<table class="rt-gauge-table">`;
    h += `<thead><tr><th colspan="2"># of Leaders: ${leaders}</th></tr></thead>`;
    h += `<tbody>`;
    h += `<tr class="rt-gauge-header"><td colspan="2">2ND ROUNDS REQUIRED</td></tr>`;
    tiers.forEach(t => {
      const val = leaders * t.mult;
      h += `<tr><td>${this._esc(t.label)}</td><td class="rt-gauge-val" style="background:${t.color}">${val}</td></tr>`;
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

  _esc(s) {
    const d = document.createElement('div');
    d.textContent = s || '';
    return d.innerHTML;
  },

  _formatCurrentWeek() {
    const d = new Date();
    const mon = d.toLocaleString('en-US', { month: 'short' });
    const day = d.getDate();
    return `${mon} ${day}, ${d.getFullYear()}`;
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
document.addEventListener('DOMContentLoaded', () => NationalApp.init());
