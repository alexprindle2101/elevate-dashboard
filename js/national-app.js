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
    campaignRecruiting: null
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

  async loadCampaignData(campaignKey) {
    const cfg = NATIONAL_CONFIG.campaigns[campaignKey];
    if (!cfg) throw new Error('Unknown campaign: ' + campaignKey);

    // Step 1: Try to load owner list + recruiting actuals from national sheet
    let sheetData = null;
    if (NATIONAL_CONFIG.appsScriptUrl && NATIONAL_CONFIG.sheets.national && NATIONAL_CONFIG.sheets.national.id) {
      try {
        sheetData = await this._fetchRecruitingFromSheet(campaignKey);
        console.log('[NationalApp] Loaded recruiting from national sheet:', sheetData);
      } catch (err) {
        console.warn('[NationalApp] National sheet fetch failed:', err.message);
      }
    }

    // Step 2: Build owners — from sheet if available, otherwise scaffold
    if (sheetData && sheetData.owners && sheetData.owners.length) {
      this._buildOwnersFromSheet(campaignKey, sheetData);
    } else if (NATIONAL_CONFIG.appsScriptUrl) {
      // Try original campaign overview API
      try {
        const url = NATIONAL_CONFIG.appsScriptUrl +
          '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
          '&campaign=' + encodeURIComponent(campaignKey);
        const resp = await fetch(url);
        const data = await resp.json();
        if (data.error) throw new Error(data.error);
        this.state.owners = data.owners || [];
        this.state.campaignTotals = data.totals || {};
        this.state.campaignRecruiting = data.campaignRecruiting || null;
        return;
      } catch (err) {
        console.warn('[NationalApp] API fetch failed, using scaffold data:', err.message);
        this._loadScaffoldData(campaignKey);
      }
    } else {
      console.log('[NationalApp] Using scaffold data for', campaignKey);
      this._loadScaffoldData(campaignKey);
    }
  },

  // ── Fetch recruiting data from Ken's national sheet via NationalCode.gs ──
  async _fetchRecruitingFromSheet(campaignKey) {
    const weeks = NATIONAL_CONFIG.campaigns[campaignKey]?.weeksToPull || 6;
    const url = NATIONAL_CONFIG.appsScriptUrl +
      '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
      '&action=recruiting&weeks=' + weeks;
    const resp = await fetch(url);
    const result = await resp.json();
    if (result.error) throw new Error(result.error);

    // Extract the campaign-specific data
    const campaignData = result.campaigns && result.campaigns[campaignKey];
    if (!campaignData) return null;

    return {
      owners: campaignData.owners || [],
      weeks: campaignData.weeks || [],
      label: campaignData.label || ''
    };
  },

  // ── Build owner objects from national sheet data ──
  _buildOwnersFromSheet(campaignKey, sheetData) {
    const ownerNames = sheetData.owners;
    const weeks = sheetData.weeks || [];

    // Week column labels (tab names from most recent to oldest)
    const weekLabels = weeks.map(w => w.tabName);

    this.state.owners = ownerNames.map(name => {
      // Build recruiting actuals: 12 rows × N weeks
      // Each row's values array = [week0val, week1val, ...]
      const actuals = Array.from({ length: 12 }, () => []);
      for (let wi = 0; wi < weeks.length; wi++) {
        const weekData = weeks[wi].data || {};
        const ownerVals = weekData[name] || new Array(12).fill(0);
        for (let ri = 0; ri < 12; ri++) {
          actuals[ri].push(ownerVals[ri] || 0);
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
          leaders: 0,          // Stays at 0 until Ken inputs during headcount step
          weeks: weekLabels,
          rows: this._buildRows(0, actuals)
        },
        sales: {
          summary: { totalSales: 0, newInternet: 0, upgrades: 0, videoSales: 0, abpMix: '—', gigMix: '—' },
          reps: []
        },
        audit: {
          grades: { reviews: '—', website: '—', social: '—', seo: '—' },
          details: {}
        }
      };
    });

    // Campaign-level totals + aggregate recruiting
    this._buildCampaignAggregates(weekLabels);
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
          summary: d.s || { totalSales: 0, newInternet: 0, upgrades: 0, videoSales: 0, abpMix: '—', gigMix: '—' },
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
      const resp = await fetch(NATIONAL_CONFIG.appsScriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          key: NATIONAL_CONFIG.apiKey,
          action: 'importRecruiting',
          weeks: weeks
        })
      });

      const result = await resp.json();
      if (result.error) throw new Error(result.error);

      // Update state with fresh data returned by the import
      if (result.recruiting && result.recruiting.campaigns) {
        const campaignKey = this.state.campaign;
        const campaignData = result.recruiting.campaigns[campaignKey];
        if (campaignData) {
          this._buildOwnersFromSheet(campaignKey, {
            owners: campaignData.owners || [],
            weeks: campaignData.weeks || [],
            label: campaignData.label || ''
          });
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

    // Campaign-level recruiting table
    this._renderRecruitingTable(this.state.campaignRecruiting, 'campaign-recruiting');
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

    // ── Section 1: Headcount Check (starts empty — filled during call) ──
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
    prodEl.innerHTML = `
      <div class="coaching-label">Production Review <span class="coaching-sublabel">Last Week</span></div>
      <div class="prod-cards">
        ${this._prodCard('Total Units', prod.totalActual, prod.totalGoal)}
        ${this._prodCard('Wireless Lines', prod.wirelessActual, prod.wirelessGoal)}
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
        <div class="goal-field">
          <label class="goal-field-label">Wireless Units</label>
          <input type="number" class="goal-input" id="goal-wireless-${ownerIdx}" value="${goals.wirelessUnits || ''}" min="0"
            placeholder="—"
            onchange="NationalApp._updateGoal(${ownerIdx}, 'wirelessUnits', this.value)">
        </div>
      </div>
      <div class="hc-submit-row">
        <button class="hc-submit-btn" onclick="NationalApp._submitGoals(${ownerIdx})">Submit Goals</button>
        <span class="hc-submit-note" id="goal-submit-note-${ownerIdx}"></span>
      </div>`;
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

  // ── Render headcount week-over-week trend table ──
  _renderHeadcountTrend(owner, ownerIdx) {
    const trendEl = document.getElementById('health-hc-trend');
    if (!trendEl) return;

    const hist = owner.headcountHistory || [];
    if (!hist.length) {
      trendEl.style.display = 'none';
      return;
    }
    trendEl.style.display = '';

    trendEl.innerHTML = `
      <div class="coaching-label">Week-over-Week Headcount</div>
      <div class="data-table-wrap">
        <table class="data-table">
          <thead>
            <tr>
              <th>Date</th>
              <th class="num">Active</th>
              <th class="num">Leaders</th>
              <th class="num">Dist</th>
              <th class="num">Training</th>
            </tr>
          </thead>
          <tbody>
            ${hist.map((r, i) => {
              const dist = r.active - r.leaders;
              const prev = i > 0 ? hist[i - 1] : null;
              return `
                <tr>
                  <td class="bold">${this._esc(r.date)}</td>
                  <td class="num">${r.active} ${this._trendArrow(r.active, prev?.active)}</td>
                  <td class="num">${r.leaders} ${this._trendArrow(r.leaders, prev?.leaders)}</td>
                  <td class="num">${dist} ${this._trendArrow(dist, prev ? prev.active - prev.leaders : null)}</td>
                  <td class="num">${r.training} ${this._trendArrow(r.training, prev?.training)}</td>
                </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>`;
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

    trendEl.innerHTML = `
      <div class="coaching-label">Week-over-Week Production</div>
      <div class="data-table-wrap">
        <table class="data-table prod-trend-table">
          <thead>
            <tr>
              <th>Week</th>
              <th class="num">Total Units</th>
              <th class="num">Wireless Lines</th>
            </tr>
          </thead>
          <tbody>
            ${hist.map((r, i) => {
              const prev = i > 0 ? hist[i - 1] : null;
              return `
                <tr>
                  <td class="bold">${this._esc(r.date)}</td>
                  <td class="num"><span class="prod-trend-actual">${r.tA}</span><span class="prod-trend-suffix"><span class="prod-trend-goal"> of ${r.tG}</span> ${this._trendArrow(r.tA, prev?.tA)}</span></td>
                  <td class="num"><span class="prod-trend-actual">${r.wA}</span><span class="prod-trend-suffix"><span class="prod-trend-goal"> of ${r.wG}</span> ${this._trendArrow(r.wA, prev?.wA)}</span></td>
                </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>`;
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
    const wireless = parseInt(document.getElementById('goal-wireless-' + ownerIdx)?.value) || 0;

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
      return;
    }
    this._renderRecruitingTable(r, 'owner-recruiting-table');
  },

  // ══════════════════════════════════════════════════
  // RENDER: Sales Tab
  // ══════════════════════════════════════════════════

  renderSalesTab(owner) {
    const s = owner.sales;

    const summary = document.getElementById('sales-summary');
    summary.innerHTML = [
      { label: 'Total Sales', value: s.summary.totalSales },
      { label: 'New Internet', value: s.summary.newInternet },
      { label: 'Upgrades', value: s.summary.upgrades },
      { label: 'Video Sales', value: s.summary.videoSales },
      { label: 'ABP Mix %', value: s.summary.abpMix },
      { label: '1Gig+ Mix %', value: s.summary.gigMix }
    ].map(k => `
      <div class="health-kpi">
        <div class="health-kpi-value">${k.value}</div>
        <div class="health-kpi-label">${k.label}</div>
      </div>
    `).join('');

    const repsEl = document.getElementById('sales-reps-table');
    if (s.reps.length) {
      repsEl.innerHTML = `
        <div class="section-label">Rep Sales Breakdown (Tableau)</div>
        <div class="data-table-wrap">
          <table class="data-table">
            <thead>
              <tr>
                <th>Name</th>
                <th class="num">New Internet</th>
                <th class="num">Upgrade</th>
                <th class="num">Video</th>
                <th class="num">Sales (All)</th>
                <th class="num">ABP Mix</th>
                <th class="num">1Gig+ Mix</th>
                <th class="num">Tech Install</th>
              </tr>
            </thead>
            <tbody>
              ${s.reps.map(rep => `
                <tr>
                  <td class="bold">${this._esc(rep.name)}</td>
                  <td class="num">${rep.newInternet || 0}</td>
                  <td class="num">${rep.upgrade || 0}</td>
                  <td class="num">${rep.video || 0}</td>
                  <td class="num">${rep.salesAll || 0}</td>
                  <td class="num">${rep.abpMix || '—'}</td>
                  <td class="num">${rep.gigMix || '—'}</td>
                  <td class="num">${rep.techInstall || '—'}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>`;
    } else {
      repsEl.innerHTML = `
        <div class="section-label">Rep Sales Breakdown</div>
        <div class="empty-state">
          <div class="empty-state-text">Sales data will populate from Campaign Tracker Section 3 (Tableau).</div>
        </div>`;
    }
  },

  // ══════════════════════════════════════════════════
  // RENDER: Audit Tab (Online Presence)
  // ══════════════════════════════════════════════════

  renderAuditTab(owner) {
    const a = owner.audit;

    const grades = document.getElementById('audit-grades');
    grades.innerHTML = [
      { title: 'Google Reviews', grade: a.grades.reviews, icon: '⭐' },
      { title: 'Website', grade: a.grades.website, icon: '🌐' },
      { title: 'Social Media', grade: a.grades.social, icon: '📱' },
      { title: 'SEO', grade: a.grades.seo, icon: '🔍' }
    ].map(g => `
      <div class="audit-grade-card">
        <div class="audit-grade-title">${g.icon} ${g.title}</div>
        <div class="audit-grade-value ${this._gradeClass(g.grade)}">${g.grade}</div>
      </div>
    `).join('');

    const details = document.getElementById('audit-details');
    if (a.details && Object.keys(a.details).length) {
      details.innerHTML = `
        <div class="section-label">Audit Details</div>
        <div class="data-table-wrap">
          <table class="data-table">
            <thead>
              <tr>
                <th>Metric</th>
                <th>Value</th>
                <th>Notes</th>
              </tr>
            </thead>
            <tbody>
              ${Object.entries(a.details).map(([key, val]) => `
                <tr>
                  <td class="bold">${this._esc(key)}</td>
                  <td>${this._esc(String(val.value || '—'))}</td>
                  <td>${this._esc(String(val.notes || ''))}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>`;
    } else {
      details.innerHTML = `
        <div class="section-label">Audit Details</div>
        <div class="empty-state">
          <div class="empty-state-text">Online presence audit data will populate from the Performance Audit sheet.</div>
        </div>`;
    }
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
