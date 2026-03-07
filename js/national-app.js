// ═══════════════════════════════════════════════════════
// National Consultant Dashboard — App Controller
// Pulls from 4 external Google Sheets, renders owner reviews
// ═══════════════════════════════════════════════════════

const NationalApp = {
  state: {
    campaign: 'att-b2b',
    owners: [],          // Array of owner objects with all aggregated data
    selectedOwner: null, // Currently expanded owner
    currentTab: 'health',
    session: null
  },

  // ══════════════════════════════════════════════════
  // INIT
  // ══════════════════════════════════════════════════

  async init() {
    console.log('[NationalApp] init');
    // Check session
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
      // Show with whatever data we have (possibly mock)
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
      // Check aliases
      if (NATIONAL_CONFIG.loginAliases[email]) email = NATIONAL_CONFIG.loginAliases[email];
      // For v1, accept any email — we'll add real auth later
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

    // If Apps Script is configured, fetch real data
    if (NATIONAL_CONFIG.appsScriptUrl) {
      try {
        const url = NATIONAL_CONFIG.appsScriptUrl +
          '?key=' + encodeURIComponent(NATIONAL_CONFIG.apiKey) +
          '&campaign=' + encodeURIComponent(campaignKey);
        const resp = await fetch(url);
        const data = await resp.json();
        if (data.error) throw new Error(data.error);
        this.state.owners = data.owners || [];
        this.state.campaignTotals = data.totals || {};
        return;
      } catch (err) {
        console.warn('[NationalApp] API fetch failed, using scaffold data:', err.message);
      }
    }

    // Scaffold data from what we observed in the spreadsheets
    console.log('[NationalApp] Using scaffold data for', campaignKey);
    this._loadScaffoldData(campaignKey);
  },

  // Scaffold data based on actual spreadsheet observations
  _loadScaffoldData(campaignKey) {
    const ownerDefs = NATIONAL_CONFIG.owners[campaignKey] || [];

    // Demo data per owner (based on real spreadsheet patterns)
    const demoData = {
      'Jay T':            { active: 12, leaders: 3, dist: 5,  training: 4, prodLW: 18, dtv: 3, goals: 22, status: 11, statusLabel: 'Growing', headcount: 12, fb: 28, fs: 22, t2: 15, r1: '79%', conv: '68%', sb: 10, ss: 8, r2: '80%', ns: 6, nss: 5, r3: '83%', sales: 42, newI: 18, upg: 14, vid: 10, abp: '72%', gig: '45%', revG: 'A', webG: 'B+', socG: 'B', seoG: 'A-' },
      'Mason':            { active: 8,  leaders: 2, dist: 3,  training: 3, prodLW: 14, dtv: 2, goals: 18, status: 22, statusLabel: 'All Leaders', headcount: 8, fb: 20, fs: 16, t2: 11, r1: '80%', conv: '69%', sb: 7, ss: 5, r2: '71%', ns: 4, nss: 3, r3: '75%', sales: 31, newI: 14, upg: 10, vid: 7, abp: '68%', gig: '38%', revG: 'B+', webG: 'B', socG: 'C+', seoG: 'B' },
      'Steven Sykes':     { active: 15, leaders: 4, dist: 6,  training: 5, prodLW: 24, dtv: 5, goals: 28, status: 11, statusLabel: 'Growing', headcount: 15, fb: 35, fs: 28, t2: 20, r1: '80%', conv: '71%', sb: 14, ss: 11, r2: '79%', ns: 8, nss: 7, r3: '88%', sales: 56, newI: 24, upg: 18, vid: 14, abp: '75%', gig: '52%', revG: 'A', webG: 'A-', socG: 'A', seoG: 'A' },
      'Olin Salter':      { active: 6,  leaders: 1, dist: 2,  training: 3, prodLW: 8,  dtv: 1, goals: 14, status: 44, statusLabel: 'Training Only', headcount: 6, fb: 14, fs: 10, t2: 6, r1: '71%', conv: '60%', sb: 4, ss: 3, r2: '75%', ns: 2, nss: 2, r3: '100%', sales: 18, newI: 8, upg: 6, vid: 4, abp: '62%', gig: '30%', revG: 'C+', webG: 'C', socG: 'D+', seoG: 'C' },
      'Eric Martinez':    { active: 10, leaders: 2, dist: 4,  training: 4, prodLW: 16, dtv: 3, goals: 20, status: 22, statusLabel: 'All Leaders', headcount: 10, fb: 24, fs: 18, t2: 13, r1: '75%', conv: '72%', sb: 8, ss: 6, r2: '75%', ns: 5, nss: 4, r3: '80%', sales: 36, newI: 16, upg: 12, vid: 8, abp: '70%', gig: '42%', revG: 'B', webG: 'B+', socG: 'B-', seoG: 'B+' },
      'Natalia Gwarda':   { active: 9,  leaders: 2, dist: 3,  training: 4, prodLW: 12, dtv: 2, goals: 16, status: 33, statusLabel: 'Interviewing Only', headcount: 9, fb: 22, fs: 17, t2: 12, r1: '77%', conv: '71%', sb: 7, ss: 5, r2: '71%', ns: 4, nss: 3, r3: '75%', sales: 28, newI: 12, upg: 9, vid: 7, abp: '66%', gig: '36%', revG: 'B-', webG: 'C+', socG: 'B', seoG: 'C+' },
      'Nigel Gilbert':    { active: 7,  leaders: 1, dist: 3,  training: 3, prodLW: 10, dtv: 1, goals: 14, status: 55, statusLabel: 'Leaders Busy', headcount: 7, fb: 16, fs: 12, t2: 8, r1: '75%', conv: '67%', sb: 5, ss: 4, r2: '80%', ns: 3, nss: 2, r3: '67%', sales: 22, newI: 10, upg: 7, vid: 5, abp: '64%', gig: '32%', revG: 'C', webG: 'C-', socG: 'D', seoG: 'C-' }
    };

    // Build scaffold data matching what we observed in the sheets
    this.state.owners = ownerDefs.map(def => {
      const d = demoData[def.name] || {};
      return {
        name: def.name,
        tab: def.tab,
        // Office Health (Section 1 of Campaign Tracker)
        health: {
          current: {
            active: d.active || '—', leaders: d.leaders || '—', dist: d.dist || '—',
            training: d.training || '—', productionLW: d.prodLW || '—',
            dtv: d.dtv || '—', goals: d.goals || '—'
          },
          trend: [
            { date: '2/17', active: (d.active||0)-3, leaders: d.leaders, dist: (d.dist||0)-1, training: (d.training||0)-1, productionLW: (d.prodLW||0)-4, dtv: (d.dtv||0)-1, goals: d.goals },
            { date: '2/24', active: (d.active||0)-1, leaders: d.leaders, dist: d.dist, training: d.training, productionLW: (d.prodLW||0)-2, dtv: d.dtv, goals: d.goals },
            { date: '3/3',  active: d.active, leaders: d.leaders, dist: d.dist, training: d.training, productionLW: d.prodLW, dtv: d.dtv, goals: d.goals }
          ]
        },
        statusCode: d.status || null,
        statusLabel: d.statusLabel || '',
        // Recruiting Pipeline
        recruiting: {
          funnel: {
            firstRoundsBooked: d.fb || '—', firstRoundsShowed: d.fs || '—',
            turnedTo2nd: d.t2 || '—', retention1: d.r1 || '—', conversion: d.conv || '—',
            secondRoundsBooked: d.sb || '—', secondRoundsShowed: d.ss || '—', retention2: d.r2 || '—',
            newStartScheduled: d.ns || '—', newStartsShowed: d.nss || '—', retention3: d.r3 || '—',
            activeHeadcount: d.headcount || '—'
          },
          weekly: [],
          reps: []
        },
        // Sales (Tableau data)
        sales: {
          summary: {
            totalSales: d.sales || '—', newInternet: d.newI || '—', upgrades: d.upg || '—',
            videoSales: d.vid || '—', abpMix: d.abp || '—', gigMix: d.gig || '—'
          },
          reps: []
        },
        // Online Presence (Performance Audit)
        audit: {
          grades: { reviews: d.revG || '—', website: d.webG || '—', social: d.socG || '—', seo: d.seoG || '—' },
          details: {}
        }
      };
    });

    // Campaign-level totals (aggregated)
    const totals = this.state.owners.reduce((acc, o) => {
      acc.headcount += (typeof o.health.current.active === 'number' ? o.health.current.active : 0);
      acc.firstBooked += (typeof o.recruiting.funnel.firstRoundsBooked === 'number' ? o.recruiting.funnel.firstRoundsBooked : 0);
      acc.newStarts += (typeof o.recruiting.funnel.newStartScheduled === 'number' ? o.recruiting.funnel.newStartScheduled : 0);
      acc.production += (typeof o.health.current.productionLW === 'number' ? o.health.current.productionLW : 0);
      return acc;
    }, { headcount: 0, firstBooked: 0, newStarts: 0, production: 0 });

    this.state.campaignTotals = {
      headcount: totals.headcount,
      firstBooked: totals.firstBooked,
      newStarts: totals.newStarts,
      retention: '78%',
      production: totals.production
    };
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
  // RENDER: Campaign Overview (KPI cards)
  // ══════════════════════════════════════════════════

  renderCampaignOverview() {
    const t = this.state.campaignTotals || {};
    const cfg = NATIONAL_CONFIG.campaigns[this.state.campaign];
    document.getElementById('campaign-title').textContent = (cfg?.label || 'Campaign') + ' Campaign';
    document.getElementById('overview-date').textContent = 'Week of ' + this._formatCurrentWeek();

    document.getElementById('kpi-headcount').textContent = t.headcount || '—';
    document.getElementById('kpi-1st-booked').textContent = t.firstBooked || '—';
    document.getElementById('kpi-starts').textContent = t.newStarts || '—';
    document.getElementById('kpi-retention').textContent = t.retention || '—';
    document.getElementById('kpi-production').textContent = t.production || '—';
  },

  // ══════════════════════════════════════════════════
  // RENDER: Owners List (cards)
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

    container.innerHTML = owners.map((o, idx) => `
      <div class="owner-card" onclick="NationalApp.openOwnerDetail(${idx})">
        <div class="owner-card-header">
          <span class="owner-card-name">${this._esc(o.name)}</span>
          <span class="owner-card-arrow">→</span>
        </div>
        <div class="owner-card-metrics">
          <div class="owner-metric">
            <div class="owner-metric-value">${o.health.current.active}</div>
            <div class="owner-metric-label">Active</div>
          </div>
          <div class="owner-metric">
            <div class="owner-metric-value">${o.recruiting.funnel.activeHeadcount}</div>
            <div class="owner-metric-label">Headcount</div>
          </div>
          <div class="owner-metric">
            <div class="owner-metric-value">${o.health.current.productionLW}</div>
            <div class="owner-metric-label">Production</div>
          </div>
        </div>
        ${o.statusLabel ? `
        <div class="owner-card-status">
          <span class="status-badge ${this._statusColor(o.statusCode)}">${this._esc(o.statusLabel)}</span>
        </div>` : ''}
      </div>
    `).join('');
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

    // Hide overview, show detail
    document.querySelector('.campaign-overview').style.display = 'none';
    document.querySelector('.owners-section').style.display = 'none';
    const detail = document.getElementById('owner-detail');
    detail.style.display = 'block';

    document.getElementById('detail-owner-name').textContent = owner.name;
    document.getElementById('detail-owner-badge').textContent = NATIONAL_CONFIG.campaigns[this.state.campaign]?.label || '';

    // Reset tab state
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
  // RENDER: Health Tab
  // ══════════════════════════════════════════════════

  renderHealthTab(owner) {
    const h = owner.health;

    // KPI cards
    const kpis = document.getElementById('health-kpis');
    kpis.innerHTML = [
      { label: 'Active Reps', value: h.current.active },
      { label: 'Leaders', value: h.current.leaders },
      { label: 'Distribution', value: h.current.dist },
      { label: 'In Training', value: h.current.training },
      { label: 'Production LW', value: h.current.productionLW },
      { label: 'DTV', value: h.current.dtv },
      { label: 'Goals', value: h.current.goals }
    ].map(k => `
      <div class="health-kpi">
        <div class="health-kpi-value">${k.value}</div>
        <div class="health-kpi-label">${k.label}</div>
      </div>
    `).join('');

    // Week-over-week trend table
    const trend = document.getElementById('health-trend');
    if (h.trend.length) {
      trend.innerHTML = `
        <div class="section-label">Week-over-Week Progression</div>
        <div class="data-table-wrap">
          <table class="data-table">
            <thead>
              <tr>
                <th>Date</th>
                <th class="num">Active</th>
                <th class="num">Leaders</th>
                <th class="num">Dist</th>
                <th class="num">Training</th>
                <th class="num">Production</th>
                <th class="num">DTV</th>
                <th>Goals</th>
              </tr>
            </thead>
            <tbody>
              ${h.trend.map(r => `
                <tr>
                  <td class="bold">${this._esc(r.date)}</td>
                  <td class="num">${r.active}</td>
                  <td class="num">${r.leaders}</td>
                  <td class="num">${r.dist}</td>
                  <td class="num">${r.training}</td>
                  <td class="num">${r.productionLW}</td>
                  <td class="num">${r.dtv}</td>
                  <td>${this._esc(r.goals)}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>`;
    } else {
      trend.innerHTML = `
        <div class="section-label">Week-over-Week Progression</div>
        <div class="empty-state">
          <div class="empty-state-text">Trend data will appear once NationalCode.gs is connected.</div>
        </div>`;
    }

    // Status codes
    const statusEl = document.getElementById('health-status');
    statusEl.innerHTML = `
      <div class="section-label">Leader Status Codes</div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <span class="status-badge green">11 — Leaders</span>
        <span class="status-badge yellow">22 — All Leaders</span>
        <span class="status-badge yellow">33 — Interviewing Only</span>
        <span class="status-badge" style="background:#fff3cd;color:#856404">44 — Training, Not Growing</span>
        <span class="status-badge red">55 — Leaders Busy</span>
        <span class="status-badge red">66 — Promotion Factory</span>
      </div>`;
  },

  // ══════════════════════════════════════════════════
  // RENDER: Recruiting Tab
  // ══════════════════════════════════════════════════

  renderRecruitingTab(owner) {
    const r = owner.recruiting;

    // Funnel visualization
    const funnel = document.getElementById('recruiting-funnel');
    const steps = [
      { label: '1st Booked', value: r.funnel.firstRoundsBooked },
      { label: '1st Showed', value: r.funnel.firstRoundsShowed, rate: r.funnel.retention1 },
      { label: 'Turned to 2nd', value: r.funnel.turnedTo2nd, rate: r.funnel.conversion },
      { label: '2nd Booked', value: r.funnel.secondRoundsBooked },
      { label: '2nd Showed', value: r.funnel.secondRoundsShowed, rate: r.funnel.retention2 },
      { label: 'Starts Sched', value: r.funnel.newStartScheduled },
      { label: 'Starts Showed', value: r.funnel.newStartsShowed, rate: r.funnel.retention3 },
      { label: 'Active', value: r.funnel.activeHeadcount }
    ];
    funnel.innerHTML = steps.map(s => `
      <div class="funnel-step">
        <div class="funnel-step-value">${s.value}</div>
        <div class="funnel-step-label">${s.label}</div>
        ${s.rate ? `<div class="funnel-step-rate ${this._rateColor(s.rate)}">${s.rate}</div>` : ''}
      </div>
    `).join('');

    // Weekly recruiting table
    const weeklyEl = document.getElementById('recruiting-weekly');
    if (r.weekly.length) {
      weeklyEl.innerHTML = `
        <div class="section-label" style="margin-top:20px">Weekly Recruiting Numbers</div>
        <div class="data-table-wrap">
          <table class="data-table">
            <thead>
              <tr>
                <th>Week</th>
                <th class="num">Calls Recv</th>
                <th class="num">No List</th>
                <th class="num">Booked</th>
                <th class="num">Showed</th>
                <th class="num">Retention</th>
                <th class="num">Starts Booked</th>
                <th class="num">Starts Showed</th>
                <th class="num">Start Retention</th>
              </tr>
            </thead>
            <tbody>
              ${r.weekly.map(w => `
                <tr>
                  <td class="bold">${this._esc(w.week)}</td>
                  <td class="num">${w.callsReceived || '—'}</td>
                  <td class="num">${w.noList || '—'}</td>
                  <td class="num">${w.booked || '—'}</td>
                  <td class="num">${w.showed || '—'}</td>
                  <td class="num">${w.retention || '—'}</td>
                  <td class="num">${w.startsBooked || '—'}</td>
                  <td class="num">${w.startsShowed || '—'}</td>
                  <td class="num">${w.startRetention || '—'}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>`;
    } else {
      weeklyEl.innerHTML = `
        <div class="section-label" style="margin-top:20px">Weekly Recruiting Numbers</div>
        <div class="empty-state">
          <div class="empty-state-text">Weekly data will populate from Campaign Tracker Section 2.</div>
        </div>`;
    }

    // Per-rep recruiting
    const repsEl = document.getElementById('recruiting-reps');
    if (r.reps.length) {
      repsEl.innerHTML = `
        <div class="section-label" style="margin-top:20px">Per-Rep Recruiting (Current Week)</div>
        <div class="data-table-wrap">
          <table class="data-table">
            <thead>
              <tr>
                <th>Name</th>
                <th class="num">1st Booked</th>
                <th class="num">1st Showed</th>
                <th class="num">Turned 2nd</th>
                <th class="num">Conversion</th>
                <th class="num">2nd Booked</th>
                <th class="num">New Starts</th>
              </tr>
            </thead>
            <tbody>
              ${r.reps.map(rep => `
                <tr>
                  <td class="bold">${this._esc(rep.name)}</td>
                  <td class="num">${rep.firstBooked || 0}</td>
                  <td class="num">${rep.firstShowed || 0}</td>
                  <td class="num">${rep.turned2nd || 0}</td>
                  <td class="num">${rep.conversion || '—'}</td>
                  <td class="num">${rep.secondBooked || 0}</td>
                  <td class="num">${rep.newStarts || 0}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>`;
    } else {
      repsEl.innerHTML = `
        <div class="section-label" style="margin-top:20px">Per-Rep Recruiting</div>
        <div class="empty-state">
          <div class="empty-state-text">Rep data will populate from All Campaigns Stats Tracker.</div>
        </div>`;
    }
  },

  // ══════════════════════════════════════════════════
  // RENDER: Sales Tab
  // ══════════════════════════════════════════════════

  renderSalesTab(owner) {
    const s = owner.sales;

    // Summary KPIs
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

    // Per-rep sales table
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

    // Grade cards
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

    // Details section
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

  _statusColor(code) {
    if (!code) return '';
    const c = parseInt(code);
    if (c <= 11) return 'green';
    if (c <= 33) return 'yellow';
    return 'red';
  },

  _rateColor(rate) {
    if (!rate || rate === '—') return '';
    const n = parseFloat(rate);
    if (isNaN(n)) return '';
    if (n >= 60) return 'trend-up';
    if (n >= 40) return 'trend-flat';
    return 'trend-down';
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
