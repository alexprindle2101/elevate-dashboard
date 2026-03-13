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

  // ── Render headcount week-over-week bar chart (newest-first, dynamic Y-axis) ──
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

    // ── Layout constants ──
    const VISIBLE = 5;
    const YAXIS_W = 36;
    const PAD_R = 10, PAD_T = 14, PAD_B = 28;
    const BAR_R = 5;
    const GAP = 0.14;
    const MIN_LABEL_H = 14;
    const svgH = 220;
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
    this._hcData = { hist, ownerIdx, n, slotW, barW, barOff, barAreaW, barAreaVisibleW, plotH, PAD_T, PAD_R, BAR_R, GAP, MIN_LABEL_H, svgH, baseY, YAXIS_W, VISIBLE, shortDate };

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
      const dist = Math.max((r.active || 0) - (r.leaders || 0), 0);
      const arrow = this._trendArrow(r.active, prev?.active);
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

    const tableHtml = `
      <div class="data-table-wrap trend-scroll">
        <table class="data-table">
          <thead><tr>
            <th>Week</th><th class="num">Active</th><th class="num">Leaders</th><th class="num">Dist</th><th class="num">Training</th>
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

    d.hist.forEach((r, i) => {
      const origIdx = d.n - 1 - i;
      const active = r.active || 0;
      const leaders = r.leaders || 0;
      const training = r.training || 0;
      const dist = Math.max(active - leaders, 0);
      const leaderH = leaders * yScale;
      const distH = dist * yScale;
      const solidH = active * yScale;
      const trainH = training * yScale;
      const x = i * d.slotW + d.barOff;
      const cx = x + d.barW / 2;
      const solidTop = d.baseY - solidH;
      const trainTop = solidTop - trainH;

      const topmost = trainH > 0 ? 'training' : (distH > 0 ? 'dist' : 'leader');

      // Leaders (bottom — blue)
      if (leaderH > 0) {
        if (topmost === 'leader') {
          svg += `<path d="${roundTop(x, d.baseY - leaderH, d.barW, leaderH, d.BAR_R)}" fill="#5b9cf6" opacity="0.9"/>`;
        } else {
          svg += `<rect x="${x}" y="${d.baseY - leaderH}" width="${d.barW}" height="${leaderH}" fill="#5b9cf6" opacity="0.9"/>`;
        }
        svg += segLabel(cx, d.baseY - leaderH, leaderH, leaders, '#fff');
      }

      // Distributors (middle — teal)
      if (distH > 0) {
        if (topmost === 'dist') {
          svg += `<path d="${roundTop(x, solidTop, d.barW, distH, d.BAR_R)}" fill="#0ea5a0" opacity="0.85"/>`;
        } else {
          svg += `<rect x="${x}" y="${solidTop}" width="${d.barW}" height="${distH}" fill="#0ea5a0" opacity="0.85"/>`;
        }
        svg += segLabel(cx, solidTop, distH, dist, '#fff');
      }

      // Training extension (dashed purple)
      if (trainH > 0) {
        svg += `<path d="${roundTop(x + 0.5, trainTop + 0.5, d.barW - 1, trainH - 1, d.BAR_R)}" fill="rgba(139,92,246,0.08)" stroke="#a78bfa" stroke-width="1" stroke-dasharray="3 2"/>`;
        svg += segLabel(cx, trainTop, trainH, training, '#7c3aed');
      }

      // Hover target
      const totalH = solidH + trainH;
      const topY = trainH > 0 ? trainTop : solidTop;
      svg += `<rect x="${x}" y="${Math.min(topY, d.baseY - 1)}" width="${d.barW}" height="${Math.max(totalH, 4)}" fill="transparent" style="cursor:pointer" onmouseenter="NationalApp._showHcTooltip(event,${origIdx},${d.ownerIdx})" onmouseleave="NationalApp._hideHcTooltip()"/>`;

      // X-axis date label
      svg += `<text x="${cx}" y="${d.svgH - 8}" text-anchor="middle" fill="#8a95a5" font-size="10" font-weight="600" font-family="Inter,sans-serif">${d.shortDate(r.date)}</text>`;
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
    try {
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'updateHeadcount',
          ownerName: sheetName,
          date: entry.date,
          active: entry.active || 0,
          leaders: entry.leaders || 0,
          dist: dist,
          training: entry.training || 0
        })
      });
      const result = await resp.json();
      if (result.error) console.warn('[HC Save] Error:', result.error);
      else console.log('[HC Save] Saved', sheetName, entry.date);
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
    const hist = [...rawHist].reverse();
    const n = hist.length;

    const shortDate = (d) => {
      if (!d) return '';
      const parts = String(d).split('/');
      return parts.length >= 2 ? parts[0] + '/' + parts[1] : d;
    };

    // ── Layout constants (match headcount chart) ──
    const VISIBLE = 5;
    const YAXIS_W = 36;
    const PAD_R = 10, PAD_T = 14, PAD_B = 28;
    const BAR_R = 5;
    const GAP = 0.14;
    const MIN_LABEL_H = 14;
    const svgH = 220;
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

    // Compute initial yMax from visible bars (multiples of 25 or 50)
    const visibleMax = this._getProdVisibleMax(0);
    const prodStep = visibleMax > 100 ? 50 : 25;
    const yMax = Math.ceil(visibleMax / prodStep) * prodStep || 25;
    this._prodCurrentYMax = yMax;

    const yAxisSvg = this._buildProdYAxisSvg(yMax);
    const barsSvg = this._buildProdBarsSvg(yMax);

    const displayW = needsScroll ? REF_W : (YAXIS_W + barAreaW);

    // ── Build editable table (back side) ──
    const tableRows = hist.map((r, i) => {
      const origIdx = n - 1 - i;
      const prev = i < n - 1 ? hist[i + 1] : null;
      const arrow = this._trendArrow(r.tA, prev?.tA);
      const pct = r.tG > 0 ? Math.round((r.tA / r.tG) * 100) : 0;
      const pctClass = pct >= 100 ? 'pct-green' : pct >= 80 ? 'pct-yellow' : pct >= 60 ? 'pct-orange' : 'pct-red';
      return `<tr>
        <td class="bold">${this._esc(r.date)}</td>
        <td class="num"><input type="number" class="hc-edit-input" value="${r.tA || 0}" min="0"
          onchange="NationalApp._onProdTableEdit(${ownerIdx},${origIdx},'tA',this.value)">${arrow}</td>
        <td class="num"><input type="number" class="hc-edit-input" value="${r.tG || 0}" min="0"
          onchange="NationalApp._onProdTableEdit(${ownerIdx},${origIdx},'tG',this.value)"></td>
        <td class="num"><span class="prod-pct-badge ${pctClass}">${pct}%</span></td>
      </tr>`;
    }).join('');

    const tableHtml = `
      <div class="data-table-wrap trend-scroll">
        <table class="data-table">
          <thead><tr>
            <th>Week</th><th class="num">Actual</th><th class="num">Goal</th><th class="num">%</th>
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
              <span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch swatch-prod-actual"></span>Actual</span>
              <span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch swatch-prod-goal"></span>Goal</span>
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
    let maxVal = 1;
    for (let i = firstVisible; i <= lastVisible; i++) {
      const r = d.hist[i];
      const barTop = Math.max(r.tA || 0, r.tG || 0);
      if (barTop > maxVal) maxVal = barTop;
    }
    return maxVal;
  },

  _onProdScroll() {
    const scrollEl = document.getElementById('prod-chart-scroll');
    if (!scrollEl || !this._prodData) return;
    const visibleMax = this._getProdVisibleMax(scrollEl.scrollLeft);
    const prodStep = visibleMax > 100 ? 50 : 25;
    const yMax = Math.ceil(visibleMax / prodStep) * prodStep || 25;
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
    const step = yMax > 100 ? 50 : 25;
    let svg = '';
    for (let val = 0; val <= yMax; val += step) {
      const y = d.baseY - val * yScale;
      svg += `<text x="${d.YAXIS_W - 6}" y="${y + 3.5}" text-anchor="end" fill="#b0b8c4" font-size="10" font-family="Inter,sans-serif">${val}</text>`;
    }
    return svg;
  },

  _buildProdBarsSvg(yMax) {
    const d = this._prodData;
    if (!d) return '';
    const yScale = d.plotH / yMax;
    const step = yMax > 100 ? 50 : 25;
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

    d.hist.forEach((r, i) => {
      const origIdx = d.n - 1 - i;
      const actual = r.tA || 0;
      const goal = r.tG || 0;
      const x = i * d.slotW + d.barOff;
      const cx = x + d.barW / 2;
      const actualH = actual * yScale;
      const goalH = goal * yScale;
      const actualTop = d.baseY - actualH;

      // Determine color based on goal attainment
      const pct = goal > 0 ? (actual / goal) : 0;
      const barColor = pct >= 1 ? '#22c55e' : pct >= 0.8 ? '#f0b429' : pct >= 0.6 ? '#f97316' : '#e53535';

      // Actual bar (solid, color-coded)
      if (actualH > 0) {
        svg += `<path d="${roundTop(x, actualTop, d.barW, actualH, d.BAR_R)}" fill="${barColor}" opacity="0.85"/>`;
        svg += segLabel(cx, actualTop, actualH, actual, '#fff');
      }

      // Goal marker line (dashed horizontal line at goal level)
      if (goal > 0) {
        const goalY = d.baseY - goalH;
        svg += `<line x1="${x - 2}" y1="${goalY}" x2="${x + d.barW + 2}" y2="${goalY}" stroke="#6366f1" stroke-width="2" stroke-dasharray="4 2" opacity="0.7"/>`;
      }

      // Hover target
      const topY = Math.min(actualTop, goal > 0 ? d.baseY - goalH : actualTop);
      const totalH = d.baseY - topY;
      svg += `<rect x="${x}" y="${Math.min(topY, d.baseY - 1)}" width="${d.barW}" height="${Math.max(totalH, 4)}" fill="transparent" style="cursor:pointer" onmouseenter="NationalApp._showProdTooltip(event,${origIdx},${d.ownerIdx})" onmouseleave="NationalApp._hideProdTooltip()"/>`;

      // X-axis date label
      svg += `<text x="${cx}" y="${d.svgH - 8}" text-anchor="middle" fill="#8a95a5" font-size="10" font-weight="600" font-family="Inter,sans-serif">${d.shortDate(r.date)}</text>`;
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

  async _saveProdRow(owner, entry) {
    const sheetName = owner._sheetName || owner.tab || owner.name;
    try {
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'updateProduction',
          ownerName: sheetName,
          date: entry.date,
          productionLW: entry.tA || 0,
          productionGoals: entry.tG || 0
        })
      });
      const result = await resp.json();
      if (result.error) console.warn('[Prod Save] Error:', result.error);
      else console.log('[Prod Save] Saved', sheetName, entry.date);
    } catch (err) {
      console.warn('[Prod Save] Network error:', err.message);
    }
  },

  // ── Production chart tooltip helpers ──
  _showProdTooltip(event, origIdx, ownerIdx) {
    const owner = this.state.owners[ownerIdx];
    if (!owner) return;
    const r = owner.productionHistory[origIdx];
    if (!r) return;
    const pct = r.tG > 0 ? Math.round((r.tA / r.tG) * 100) : 0;
    const tt = document.getElementById('prod-chart-tt');
    if (!tt) return;

    tt.innerHTML = `
      <div style="font-weight:700;margin-bottom:4px">${this._esc(r.date)}</div>
      <div><span class="tt-swatch" style="background:#22c55e"></span>Actual: <strong>${r.tA}</strong></div>
      <div><span class="tt-swatch" style="background:#6366f1;border-radius:0;height:2px;width:10px;border-top:2px dashed #6366f1;background:none"></span>Goal: <strong>${r.tG}</strong></div>
      <div style="border-top:1px solid rgba(255,255,255,0.2);margin:4px 0;padding-top:4px">${pct}% of goal</div>`;

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

  // ── Lazy-load recruiting costs for a single owner ──
  // Tries new weekly Indeed Tracking endpoint first, falls back to old cost sheet
  async _loadAndRenderCosts(owner) {
    const el = document.getElementById('owner-indeed-costs');
    if (!el) return;

    // Render NLR banner above cost section
    this._renderNlrBanner();

    // ── Try new weekly Indeed Tracking endpoint first ──
    if (owner.indeedTracking) {
      this._renderIndeedTracking(owner);
      return;
    }
    if (!owner._trackingFetched && !owner._trackingFetching) {
      owner._trackingFetching = true;
      el.innerHTML = `<div class="coaching-label">Weekly Ad Spend</div>
        <div class="empty-state"><div class="loading-spinner" style="width:24px;height:24px;border-width:3px;margin:0 auto"></div>
        <div class="empty-state-text" style="margin-top:8px">Loading weekly ad data...</div></div>`;
      const url = NATIONAL_CONFIG.appsScriptUrl +
        '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
        '&action=indeedTracking' +
        '&owner=' + encodeURIComponent(owner.name) +
        '&_t=' + Date.now();
      // Retry up to 2 times (Apps Script cold starts can timeout)
      let result = null;
      for (let attempt = 1; attempt <= 2; attempt++) {
        try {
          console.log('[NationalApp] Indeed Tracking fetch attempt', attempt, 'for', owner.name);
          const resp = await this._fetchWithTimeout(fetch(url), 45000);
          result = await resp.json();
          console.log('[NationalApp] Indeed Tracking response:', result.error || (result.weeks ? result.weeks.length + ' weeks' : 'no weeks'));
          if (!result.error && result.weeks && result.weeks.length) break;
          result = null; // didn't get valid data, retry
        } catch (err) {
          console.warn('[NationalApp] Indeed Tracking attempt', attempt, 'for', owner.name, ':', err.message);
          result = null;
        }
      }
      owner._trackingFetching = false;
      if (result && result.weeks && result.weeks.length) {
        owner.indeedTracking = result;
        owner._trackingFetched = true;
        if (this.state.selectedOwner === owner && this.state.currentTab === 'recruiting') {
          this._renderIndeedTracking(owner);
        }
        return;
      }
      // Don't cache failure — allow retry on next tab visit
      console.warn('[NationalApp] Indeed Tracking failed for', owner.name);
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
    const chart1 = this._buildRetentionCard('New Starts', '#22c55e', weekLabels, newStartsBooked, newStartsShowed, 'Booked new starts that actually show up');
    const chart2 = this._buildRetentionCard('2nd Round Interviews', '#3b82f6', weekLabels, r2Booked, r2Showed, '1st round interview quality');
    const chart3 = this._buildRetentionCard('1st Round Interviews', '#f59e0b', weekLabels, r1Booked, r1Showed, 'Recruiting hub / phone booker performance');

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
  _buildRetentionCard(title, color, weekLabels, booked, showed, subtitle) {
    const n = weekLabels.length;
    const VISIBLE = 4;

    // Current week hero values
    const curShowed = showed[0] || 0;
    const curBooked = booked[0] || 0;
    const curPct = curBooked > 0 ? Math.round((curShowed / curBooked) * 100) : null;

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

    // Light gridlines
    for (let v = 0; v <= yMax; v += step) {
      const y = baseY - v * yScale;
      svg += `<line x1="0" y1="${y}" x2="${vbW}" y2="${y}" stroke="#e8ecf1" stroke-width="0.5"/>`;
    }

    // Bars per week
    for (let i = 0; i < n; i++) {
      const bk = booked[i] || 0;
      const sh = showed[i] || 0;
      const pct = bk > 0 ? Math.round((sh / bk) * 100) : null;

      const gx = i * SLOT_W + barOff;
      const cx = i * SLOT_W + SLOT_W / 2;

      // Booked bar (ghost — the potential)
      const bkH = bk * yScale;
      if (bkH > 0) {
        svg += `<path d="${roundTop(gx, baseY - bkH, barW, bkH, BAR_R)}" fill="${color}" opacity="0.13"/>`;
        svg += `<path d="${roundTop(gx, baseY - bkH, barW, bkH, BAR_R)}" fill="none" stroke="${color}" stroke-width="1" opacity="0.35" stroke-dasharray="3,2"/>`;
      }

      // Showed bar (solid — actual output)
      const shH = sh * yScale;
      if (shH > 0) {
        svg += `<path d="${roundTop(gx, baseY - shH, barW, shH, BAR_R)}" fill="${color}" opacity="0.88"/>`;
      }

      // Booked count in ghost gap area (between showed top and booked top)
      const gapH = bkH - shH;
      if (bk > 0 && gapH > 12) {
        const gapMidY = baseY - shH - gapH / 2 + 3;
        svg += `<text x="${cx}" y="${gapMidY}" text-anchor="middle" fill="${color}" font-size="8" font-weight="700" font-family="Inter,sans-serif" opacity="0.55">${bk}</text>`;
      } else if (bk > 0 && bk === sh && shH > 28) {
        // Booked == showed: show booked near top of solid bar
        svg += `<text x="${cx}" y="${baseY - shH + 10}" text-anchor="middle" fill="rgba(255,255,255,0.6)" font-size="7" font-weight="600" font-family="Inter,sans-serif">${bk}</text>`;
      }

      // Showed count inside solid bar (if tall enough)
      if (sh > 0 && shH > 14) {
        svg += `<text x="${cx}" y="${baseY - shH / 2 + 4}" text-anchor="middle" fill="#fff" font-size="9" font-weight="700" font-family="Inter,sans-serif">${sh}</text>`;
      }

      // Retention % above bar
      if (pct !== null) {
        const topY = Math.min(baseY - bkH, baseY - shH);
        svg += `<text x="${cx}" y="${topY - 5}" text-anchor="middle" fill="${color}" font-size="8" font-weight="800" font-family="Inter,sans-serif">${pct}%</text>`;
      }

      // Week label below
      svg += `<text x="${cx}" y="${svgH - 3}" text-anchor="middle" fill="#8a95a5" font-size="7" font-weight="600" font-family="Inter,sans-serif">${weekLabels[i]}</text>`;
    }

    // Legend
    const legend = `<div class="hc-chart-legend" style="margin-top:4px">
      <span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch" style="background:${color}"></span>Showed</span>
      <span class="hc-chart-legend-item"><span class="hc-chart-legend-swatch" style="background:${color};opacity:0.18;border:1.5px dashed ${color}"></span>Booked</span>
    </div>`;

    return `
      <div class="recruit-chart-card">
        <div class="rc-card-header">
          <div>
            <div class="recruit-chart-title">${title}</div>
            <div class="rc-card-subtitle">${subtitle}</div>
          </div>
          <div class="rc-card-hero">
            <div class="rc-card-hero-num" style="color:${color}">${curShowed}</div>
            <div class="rc-card-hero-label">showed</div>
          </div>
        </div>
        <div class="rc-bar-wrap">
          <svg viewBox="0 0 ${vbW} ${svgH}" width="${svgWidthPct}" preserveAspectRatio="xMinYMid meet" style="display:block">${svg}</svg>
        </div>
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
    const latest = weeks[weeks.length - 1];
    const prev = weeks.length > 1 ? weeks[weeks.length - 2] : null;

    let html = '';

    // ── Part 1: Weekly Totals KPI Cards ──
    html += `<div class="coaching-label">Weekly Ad Spend Overview</div>`;
    html += `<div class="it-kpi-row">`;
    html += this._itKpiCard('Total Spend', this._fmtDollar(latest.totalSpend),
      latest.delta ? latest.delta.spendPct : null, true);
    html += this._itKpiCard('Applies', this._fmtNum(latest.totalApplies),
      latest.delta ? latest.delta.appliesPct : null, false);
    html += this._itKpiCard('2nds', this._fmtNum(latest.total2nds), null, false);
    html += this._itKpiCard('New Starts', this._fmtNum(latest.totalNewStarts), null, false);
    html += this._itKpiCard('CPA', this._fmtDollar(latest.cpa),
      latest.delta ? latest.delta.cpaPct : null, true);
    html += this._itKpiCard('CPNS', this._fmtDollar(latest.cpns),
      latest.delta ? latest.delta.cpnsPct : null, true);
    html += `</div>`;
    html += `<div class="it-kpi-week-label">Week of ${this._esc(latest.weekOf)} · ${latest.numAds} ads</div>`;

    // ── Part 2: Week-over-Week Trend Table ──
    html += `<div class="coaching-label it-section-label">Week-over-Week Trends</div>`;
    html += this._buildTrackingTrend(weeks);

    // ── Part 3: Ad Breakdown for Latest Week ──
    html += `<div class="coaching-label it-section-label">Ad Breakdown
      <select class="it-week-select" onchange="NationalApp._switchTrackingWeek(this.value)">
        ${weeks.map((w, i) => `<option value="${i}"${i === weeks.length - 1 ? ' selected' : ''}>${this._esc(w.weekOf)}</option>`).join('')}
      </select>
    </div>`;
    html += `<div id="it-ad-breakdown">`;
    html += this._buildAdBreakdown(latest, prev);
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

    // Show last 6 weeks max (newest on right)
    const shown = weeks.slice(-6);

    let h = `<div class="it-trend-wrap"><div class="data-table-wrap"><table class="data-table it-trend-table">
      <thead><tr><th></th>
        ${shown.map(w => `<th class="num">${this._esc(w.weekOf)}</th>`).join('')}
      </tr></thead><tbody>`;

    for (const r of TREND_ROWS) {
      const fmt = r.fmt === 'dollar' ? this._fmtDollar : this._fmtNum;
      h += `<tr><td class="rc-label">${r.label}</td>`;
      for (let i = 0; i < shown.length; i++) {
        const val = shown[i][r.key] ?? 0;
        const prev = i > 0 ? (shown[i - 1][r.key] ?? null) : null;
        const arrow = prev !== null ? this._costTrendArrow(val, prev, r.lowerBetter) : '';
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

    // Data rows
    rows.forEach((row, ri) => {
      html += `<tr>`;
      html += `<td>${this._esc(row.label)}</td>`;
      html += `<td class="rt-projected">${this._fmtCell(row.projected, row.isRate)}</td>`;

      // Weekly values with conditional coloring (reversed: newest-first)
      const vals = [...row.values].reverse();
      vals.forEach(val => {
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
