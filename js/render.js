// ═══════════════════════════════════════════════════════
// ELEVATE — Rendering Functions
// Config-driven products, no hardcoded air/cell/fiber/voip
// ═══════════════════════════════════════════════════════

const Render = {
  activeCharts: [],
  expandedPeriods: {},
  currentView: 'days', // 'days' | 'weeks'

  // ── Helpers ──
  fmt(v) { return (v === null || v === undefined) ? '—' : v; },
  pct(n, d) { return d > 0 ? ((n / d) * 100).toFixed(1) : '0.0'; },

  dailyClass(u) {
    if (!u) return 'c-red';
    if (u <= (OFFICE_CONFIG.dailyYellowThreshold || 2)) return 'c-yellow';
    return 'c-green';
  },

  weeklyClass(u) {
    if (!u) return 'c-red';
    if (u <= (OFFICE_CONFIG.weeklyYellowThreshold || 9)) return 'c-yellow';
    return 'c-green';
  },

  rankLabel(i) {
    if (i === 0) return '<span class="rank-medal">🥇</span>';
    if (i === 1) return '<span class="rank-medal">🥈</span>';
    if (i === 2) return '<span class="rank-medal">🥉</span>';
    return i + 1;
  },

  visiblePeriods() {
    if (this.currentView === 'weeks') return [7, 8, 9, 10, 11];
    return [0, 1, 2, 3, 4, 5, 6];
  },

  // ── Period data access ──
  getPeriod(p, pi) {
    if (pi < 7) return p.days[pi];
    if (pi === 7) return DataPipeline.sumPeriods(p.days, OFFICE_CONFIG);
    if (pi === 8) return p.lw;
    if (pi === 9) return p.w2;
    if (pi === 10) return p.w3;
    // pi === 11: 4W total
    return DataPipeline.sumPeriods([...p.days, p.lw, p.w2, p.w3], OFFICE_CONFIG);
  },

  sumPeriod(people, pi) {
    const periods = people.map(p => this.getPeriod(p, pi));
    return DataPipeline.sumPeriods(periods, OFFICE_CONFIG);
  },

  twUnits(p) { return p.days.reduce((s, d) => s + d.units, 0); },
  twYeses(p) { return p.days.reduce((s, d) => s + d.y, 0); },

  // ── View Toggle ──
  setView(v) {
    this.currentView = v;
    document.querySelectorAll('.view-toggle').forEach(b => {
      b.classList.toggle('active', b.dataset.view === v);
    });
    this.renderMainTable(App.state.people);
    if (this._activeTeamLB) this._rebuildTeamLeaderboard();
  },

  // ── Toggle Period Expansion ──
  togglePeriod(pi) {
    this.expandedPeriods[pi] = !this.expandedPeriods[pi];
    this.renderMainTable(App.state.people);
    // If a team leaderboard is visible, rebuild it too
    if (this._activeTeamLB) this._rebuildTeamLeaderboard();
  },

  // ── Build Table Headers ──
  buildTableHeaders(tableId) {
    const table = document.getElementById(tableId);
    if (!table) return;
    const gr = table.querySelector('thead tr:nth-child(1)');
    const dr = table.querySelector('thead tr:nth-child(2)');
    if (!gr || !dr) return;

    const products = OFFICE_CONFIG.columns.products;
    gr.innerHTML = '';
    dr.innerHTML = '';

    gr.insertAdjacentHTML('beforeend', '<th colspan="2"></th>');
    dr.insertAdjacentHTML('beforeend', '<th style="width:36px">#</th><th class="name-col" style="min-width:120px">Name</th>');

    this.visiblePeriods().forEach(pi => {
      const isExp = !!this.expandedPeriods[pi];
      const isWeek = WEEK_PERIODS.has(pi);
      const bc = isWeek ? 'week-start' : 'day-start';

      // Group header
      const th = document.createElement('th');
      th.colSpan = isExp ? (products.length + 2) : 2; // yeses + products + units vs yeses + units
      th.className = `day-toggle ${bc}${isWeek ? ' week-col' : ''}${isExp ? ' expanded' : ''}`;
      th.dataset.pi = pi;
      th.innerHTML = `${PERIOD_LABELS[pi]} <span class="toggle-icon">▼</span>`;
      th.addEventListener('click', () => this.togglePeriod(pi));
      gr.appendChild(th);

      // Sub headers
      const yTh = document.createElement('th');
      yTh.textContent = 'Yeses';
      yTh.className = `${bc}${isWeek ? ' week-col' : ''}`;
      dr.appendChild(yTh);

      products.forEach(prod => {
        const dth = document.createElement('th');
        dth.textContent = prod.label;
        dth.className = `detail-col dp${pi}${isWeek ? ' week-col' : ''}${isExp ? '' : ' hidden'}`;
        dr.appendChild(dth);
      });

      const uTh = document.createElement('th');
      uTh.textContent = 'Units';
      uTh.className = isWeek ? 'week-col' : '';
      dr.appendChild(uTh);
    });
  },

  // ── Period Cells (config-driven) ──
  periodCells(d, pi) {
    const isExp = !!this.expandedPeriods[pi];
    const hc = isExp ? '' : 'hidden';
    const isWeek = WEEK_PERIODS.has(pi);
    const uc = isWeek ? this.weeklyClass(d.units) : this.dailyClass(d.units);
    const bc = isWeek ? 'week-start' : 'day-start';
    const wc = isWeek ? ' week-col' : '';

    let html = `<td class="val c-none ${bc}${wc}">${this.fmt(d.y)}</td>`;

    OFFICE_CONFIG.columns.products.forEach(prod => {
      const val = d.products ? (d.products[prod.key] || 0) : (d[prod.key] || 0);
      html += `<td class="val c-none dp${pi}${wc} ${hc}">${this.fmt(val)}</td>`;
    });

    html += `<td class="val ${uc}${wc}">${this.fmt(d.units)}</td>`;
    return html;
  },

  // ── Person Row ──
  personRowHTML(p, rankIdx) {
    const canView = Roster.canViewMetrics(p.name, App.state.currentRole, App.state.currentPersona, App.state.people);
    const isSelf = p.name === App.state.currentPersona;
    const nameCell = canView
      ? `<span class="name-text name-link" onclick="App.openPersonProfile('${p.name.replace(/'/g, "\\'")}')">${p.name}${isSelf ? ' <span style="font-size:9px;color:var(--sc-cyan);letter-spacing:1px">(YOU)</span>' : ''}</span><br><span style="font-size:10px;color:var(--silver-dim)">${p.role}</span>`
      : `<span class="name-text" style="cursor:default">${p.name}<span class="locked-badge">🔒</span></span><br><span style="font-size:10px;color:var(--silver-dim)">${p.role}</span>`;

    let cells = `<td class="rank ${rankIdx < 3 ? 'rank-' + (rankIdx + 1) : ''}">${this.rankLabel(rankIdx)}</td><td class="name-cell">${nameCell}</td>`;
    this.visiblePeriods().forEach(pi => {
      cells += this.periodCells(this.getPeriod(p, pi), pi);
    });
    return cells;
  },

  personRowClass(p) {
    if (p.name === App.state.currentPersona) return 'own-row';
    if (!Roster.canViewMetrics(p.name, App.state.currentRole, App.state.currentPersona, App.state.people)) return 'row-locked';
    return '';
  },

  groupTotalRowHTML(people, label, rowClass) {
    let cells = `<td></td><td class="name-cell" style="text-align:left;padding-left:8px">${label}</td>`;
    this.visiblePeriods().forEach(pi => {
      cells += this.periodCells(this.sumPeriod(people, pi), pi);
    });
    return `<tr class="group-total-row ${rowClass}">${cells}</tr>`;
  },

  // ── MAIN TABLE ──
  renderMainTable(people) {
    if (!people) return;
    this.buildTableHeaders('main-table');
    const tbody = document.getElementById('main-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    const active = people.filter(p => !Roster.deactivated.has(p.name) && !this.NON_SALES_ROLES.has(p._roleKey));

    // Split by role: leaders (manager, jd, l1) and reps
    const leaderRoles = new Set(['manager', 'jd', 'l1']);
    const leaders = active.filter(p => leaderRoles.has(p._roleKey));
    const reps = active.filter(p => !leaderRoles.has(p._roleKey));

    const sortedL = [...leaders].sort((a, b) => this.twUnits(b) - this.twUnits(a));
    sortedL.forEach((p, i) => {
      const tr = document.createElement('tr');
      tr.className = this.personRowClass(p);
      tr.innerHTML = this.personRowHTML(p, i);
      tbody.appendChild(tr);
    });
    if (sortedL.length > 0) {
      tbody.insertAdjacentHTML('beforeend', this.groupTotalRowHTML(leaders, '👑 Leader Total', 'leader-total-row'));
    }

    const sortedR = [...reps].sort((a, b) => this.twUnits(b) - this.twUnits(a));
    sortedR.forEach((p, i) => {
      const tr = document.createElement('tr');
      tr.className = this.personRowClass(p);
      tr.innerHTML = this.personRowHTML(p, i);
      tbody.appendChild(tr);
    });
    if (sortedR.length > 0) {
      tbody.insertAdjacentHTML('beforeend', this.groupTotalRowHTML(reps, '⚡ Rep Total', 'rep-total-row'));
    }

    tbody.insertAdjacentHTML('beforeend', this.groupTotalRowHTML(active, 'Office Total', 'office-total-row'));
  },

  // ── PODIUMS ──
  makePodium(id, items, getU, getY, getRole) {
    const wrap = document.getElementById(id);
    if (!wrap) return;
    wrap.innerHTML = '';
    const sorted = [...items].sort((a, b) => getU(b) - getU(a));
    const top3 = sorted.slice(0, 3);
    if (top3.length === 0) {
      wrap.innerHTML = '<div style="text-align:center;padding:24px;color:var(--silver-dim);font-size:13px">No data yet</div>';
      return;
    }
    const order = top3.length >= 3 ? [top3[1], top3[0], top3[2]] : top3;
    const classes = ['p2', 'p1', 'p3'], medals = ['🥈', '🥇', '🥉'];
    order.forEach((item, i) => {
      const card = document.createElement('div');
      card.className = `podium-card ${classes[i]}`;
      card.innerHTML = `<div class="podium-position">${medals[i]}</div><div class="podium-name">${item.name}</div><div class="podium-role">${getRole(item)}</div>` +
        `<div class="podium-stats"><div class="podium-stat"><div class="podium-stat-val units-val">${getU(item)}</div><div class="podium-stat-lbl">Units</div></div>` +
        `<div class="podium-stat"><div class="podium-stat-val yeses-val">${getY(item)}</div><div class="podium-stat-lbl">Yeses</div></div></div>`;
      wrap.appendChild(card);
    });
  },

  // ── TEAM GRID ──
  renderTeamGrid(teams) {
    const grid = document.getElementById('team-grid');
    if (!grid || !teams) return;
    grid.innerHTML = '';
    const sorted = [...teams].sort((a, b) => b.units - a.units);
    sorted.forEach((t, i) => {
      const canSee = Roster.canViewTeam(t.name, App.state.currentRole, App.state.currentPersona, App.state.people, App.state.teams);
      const isSubTeam = t.isSubTeam || false;
      const hasChildren = t.children && t.children.length > 0;
      const directCount = t.directMembers ? t.directMembers.length : (t.members ? t.members.length : 0);
      const totalCount = t.members ? t.members.length : 0;

      const card = document.createElement('div');
      card.className = `team-card${i === 0 ? ' rank-1-card' : ''}${!canSee ? ' team-locked' : ''}${isSubTeam ? ' sub-team-card' : ''}`;

      const badgeHtml = isSubTeam ? '<div class="team-parent-badge">Sub-team</div>' : '';
      const displayName = (t.emoji ? t.emoji + ' ' : '') + t.name;
      const memberLabel = hasChildren && totalCount !== directCount
        ? `${directCount} direct · ${totalCount} total`
        : `${totalCount} members`;

      card.innerHTML =
        badgeHtml +
        `<div class="team-rank">${i === 0 ? '🏆 ' : ''}#${i + 1}</div>` +
        `<div class="team-name${canSee ? ' name-link' : ''}" ${canSee ? `onclick="App.openTeamProfile('${t.name.replace(/'/g, "\\'")}')"` : ''}>${displayName}${!canSee ? ' <span class="locked-badge">🔒</span>' : ''}</div>` +
        `<div class="team-score">${t.units}</div>` +
        `<div class="team-label">Units this week</div>` +
        `<div class="team-label" style="font-size:10px;margin-top:-4px">${memberLabel}</div>`;
      grid.appendChild(card);
    });
  },

  // ── HERO STATS ──
  renderHeroStats(salesPeople) {
    const offU = salesPeople.reduce((s, p) => s + this.twUnits(p), 0);
    const offY = salesPeople.reduce((s, p) => s + this.twYeses(p), 0);
    const heroY = document.getElementById('hero-yeses');
    const heroU = document.getElementById('hero-units');
    if (heroY) heroY.textContent = offY;
    if (heroU) heroU.textContent = offU;

    // Member count (sales roles only) — leaders = anyone above Client Rep
    const repCount = salesPeople.filter(p => p._roleKey === 'rep').length;
    const leaderCount = salesPeople.length - repCount;
    const mc = document.getElementById('member-count');
    if (mc) mc.textContent = `${leaderCount} leaders · ${repCount} reps`;
  },

  // ── PROFILE PAGE ──
  openProfilePage(html) {
    this.activeCharts.forEach(c => { try { c.destroy(); } catch (e) {} });
    this.activeCharts = [];
    const page = document.getElementById('profile-page');
    const inner = document.getElementById('profile-page-inner');
    if (!page || !inner) return;
    inner.innerHTML = html;
    page.style.display = 'block';
    page.scrollTop = 0;
    document.body.style.overflow = 'hidden';
  },

  closeProfile() {
    this.activeCharts.forEach(c => { try { c.destroy(); } catch (e) {} });
    this.activeCharts = [];
    this._activeTeamLB = null;
    const page = document.getElementById('profile-page');
    if (page) page.style.display = 'none';
    document.body.style.overflow = '';
  },

  mkChart(canvasId, config) {
    const el = document.getElementById(canvasId);
    if (!el) return;
    const box = el.parentElement;
    if (box) { el.width = box.offsetWidth; el.height = box.offsetHeight; }
    const chart = new Chart(el, config);
    this.activeCharts.push(chart);
  },

  // ── PERSON PROFILE ──
  openPersonProfile(name) {
    if (!Roster.canViewMetrics(name, App.state.currentRole, App.state.currentPersona, App.state.people)) {
      this.openProfilePage(`
        <div class="profile-topbar">
          <button class="profile-back" onclick="Render.closeProfile()"><span>←</span> Back</button>
          <div class="profile-heading">
            <div class="profile-page-name">${name}</div>
            <div class="profile-page-sub">Profile Restricted</div>
          </div>
        </div>
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:80px 24px;text-align:center">
          <div style="font-size:48px;margin-bottom:16px">🔒</div>
          <div style="font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:800;color:var(--white);margin-bottom:8px">Access Restricted</div>
          <div style="color:var(--silver-dim);font-size:14px;max-width:360px;line-height:1.6">You don't have permission to view this person's full profile. Contact your Junior Director or Admin.</div>
        </div>`);
      return;
    }

    const p = App.state.people.find(x => x.name === name);
    if (!p) return;
    const m = p.metrics;
    const tw4 = this.getPeriod(p, 10);
    const twU = this.twUnits(p), twY = this.twYeses(p);
    const total = Math.max(m.totalActs, 1);
    const id = 'p' + name.replace(/[^a-zA-Z0-9]/g, '');
    const vsColor = m.vsPct >= 0 ? '#22c55e' : '#e53535';
    const vsArrow = m.vsPct >= 0 ? '↑' : '↓';

    const churnHTML = m.churnBuckets.map(b => `
      <div class="churn-bucket">
        <div class="churn-label">${b.label}</div>
        <div class="churn-fraction">${b.pct === 'N/A' ? 'N/A' : (b.disco > 0 ? `(${b.disco}/${b.activated})` : '---')}</div>
        <div class="churn-pct" style="color:${b.pct === 'N/A' ? 'var(--silver-dim)' : (b.disco > 0 ? '#f97316' : 'var(--silver-dim)')}">${b.pct === 'N/A' ? 'Pending' : (b.disco > 0 ? b.pct + '%' : '---')}</div>
      </div>`).join('');

    this.openProfilePage(`
      <div class="profile-topbar">
        <button class="profile-back" onclick="Render.closeProfile()"><span>←</span> Back</button>
        <div class="profile-heading">
          <div class="profile-page-name">${p.name}</div>
          <div class="profile-page-sub">${p.role} · Team: <span class="name-link" onclick="App.openTeamProfile('${(p.team || '').replace(/'/g, "\\'")}')" style="cursor:pointer">${p.team || 'Unassigned'}</span></div>
        </div>
      </div>

      <div class="profile-stats-bar">
        <div class="profile-stat"><div class="profile-stat-val units-val">${twU}</div><div class="profile-stat-lbl">This Wk Units</div></div>
        <div class="profile-stat"><div class="profile-stat-val">${twY}</div><div class="profile-stat-lbl">This Wk Yeses</div></div>
        <div class="profile-stat"><div class="profile-stat-val units-val">${tw4.units}</div><div class="profile-stat-lbl">4W Units</div></div>
        <div class="profile-stat"><div class="profile-stat-val">${tw4.y}</div><div class="profile-stat-lbl">4W Yeses</div></div>
        <div class="profile-stat"><div class="profile-stat-val" style="color:${vsColor}">${vsArrow}${Math.abs(m.vsPct)}%</div><div class="profile-stat-lbl">vs 4Wk Avg</div></div>
      </div>

      <div class="profile-section-title">Sales Breakdown</div>
      <div class="breakdown-grid">
        <div class="breakdown-remark" style="background:${m.remarkColor}22;border:1px solid ${m.remarkColor}55">
          <div style="font-size:10px;letter-spacing:2px;color:${m.remarkColor};text-transform:uppercase;font-family:'Barlow Condensed',sans-serif;font-weight:700;margin-bottom:6px">Remarks</div>
          <div style="font-size:18px;font-weight:700;color:${m.remarkColor}">${m.remark}</div>
          <div style="font-size:11px;color:var(--silver-dim);margin-top:6px">Rec Wk Avg: <b style="color:var(--silver)">${m.recentAvg}</b> &nbsp;·&nbsp; 4Wk Avg: <b style="color:var(--silver)">${m.fourWkAvg}</b></div>
        </div>
        <div class="breakdown-stat"><div class="breakdown-val">${m.recentAvg}</div><div class="breakdown-lbl">Rec. Wk Avg</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${vsColor}">${vsArrow}${Math.abs(m.vsPct)}%</div><div class="breakdown-lbl">vs 4Wk Avg</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${parseFloat(this.pct(m.active, total)) >= 75 ? '#22c55e' : '#f0b429'}">${this.pct(m.active, total)}%</div><div class="breakdown-lbl">Active</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:#f0b429">${this.pct(m.pending, total)}%</div><div class="breakdown-lbl">Pending</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${parseFloat(this.pct(m.cancel, total)) > 10 ? '#e53535' : 'var(--silver)'}">${this.pct(m.cancel, total)}%</div><div class="breakdown-lbl">Cancel</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${parseFloat(this.pct(m.projDisco, total)) > 5 ? '#f97316' : '#22c55e'}">${this.pct(m.projDisco, total)}%</div><div class="breakdown-lbl">Proj. Disco</div></div>
      </div>

      <div class="profile-section-title">Churn Buckets</div>
      <div class="churn-grid">${churnHTML}</div>

      <div class="profile-section-title">Sales Trends</div>
      <div class="charts-row">
        <div class="chart-box">
          <div class="chart-title">Sales Per Day</div>
          <div class="chart-wrap" style="height:160px"><canvas id="spd_${id}"></canvas></div>
          <div class="chart-title" style="margin-top:18px;font-size:11px">Pos / Neg Alpha</div>
          <div class="chart-wrap" style="height:100px"><canvas id="alpha_${id}"></canvas></div>
        </div>
        <div class="chart-box">
          <div class="chart-title">Time of Sales — Recent</div>
          <div class="chart-wrap" style="height:130px"><canvas id="tsr_${id}"></canvas></div>
          <div class="chart-title" style="margin-top:18px">Time of Sales — 4Wk Running</div>
          <div class="chart-wrap" style="height:130px"><canvas id="ts4_${id}"></canvas></div>
        </div>
      </div>

      <div class="profile-section-title">Products Sold</div>
      <div class="charts-row">
        <div class="chart-box">
          <div class="chart-title">Recent (2Wk)</div>
          <div class="chart-wrap" style="height:180px"><canvas id="pr_${id}"></canvas></div>
        </div>
        <div class="chart-box">
          <div class="chart-title">4Wk Running</div>
          <div class="chart-wrap" style="height:180px"><canvas id="p4_${id}"></canvas></div>
        </div>
      </div>
    `);

    setTimeout(() => this.drawCharts(id, m), 100);
  },

  // ── TEAM PROFILE ──
  openTeamProfile(teamName) {
    if (!Roster.canViewTeam(teamName, App.state.currentRole, App.state.currentPersona, App.state.people, App.state.teams)) {
      this.openProfilePage(`
        <div class="profile-topbar">
          <button class="profile-back" onclick="Render.closeProfile()"><span>←</span> Back</button>
          <div class="profile-heading">
            <div class="profile-page-name">${teamName}</div>
            <div class="profile-page-sub">Team Restricted</div>
          </div>
        </div>
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:80px 24px;text-align:center">
          <div style="font-size:48px;margin-bottom:16px">🔒</div>
          <div style="font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:800;color:var(--white);margin-bottom:8px">Access Restricted</div>
          <div style="color:var(--silver-dim);font-size:14px;max-width:360px;line-height:1.6">You can only view your own team's dashboard.</div>
        </div>`);
      return;
    }

    const team = App.state.teams.find(t => t.name === teamName);
    if (!team) return;
    const members = team.members || [];
    const directMembers = team.directMembers || members;
    const m = team.metrics;

    if (!m) {
      this.openProfilePage(`
        <div class="profile-topbar">
          <button class="profile-back" onclick="Render.closeProfile()"><span>←</span> Back</button>
          <div class="profile-heading">
            <div class="profile-page-name">${team.name}</div>
            <div class="profile-page-sub">Team · <span>0 members</span></div>
          </div>
        </div>
        <p style="color:var(--silver-dim);padding:40px;text-align:center">No members assigned to this team yet.</p>`);
      return;
    }

    const tw = { u: members.reduce((s, p) => s + this.twUnits(p), 0), y: members.reduce((s, p) => s + this.twYeses(p), 0) };
    const id = 't' + teamName.replace(/[^a-zA-Z0-9]/g, '');
    const tableId = 'team-lb-' + id;
    const vsColor = m.vsPct >= 0 ? '#22c55e' : '#e53535';
    const vsArrow = m.vsPct >= 0 ? '↑' : '↓';
    const total = Math.max(m.totalActs, 1);

    const churnHTML = m.churnBuckets.map(b => `
      <div class="churn-bucket">
        <div class="churn-label">${b.label}</div>
        <div class="churn-fraction">${b.pct === 'N/A' ? 'N/A' : (b.disco > 0 ? `(${b.disco}/${b.activated})` : '---')}</div>
        <div class="churn-pct" style="color:${b.pct === 'N/A' ? 'var(--silver-dim)' : (b.disco > 0 ? '#f97316' : 'var(--silver-dim)')}">${b.pct === 'N/A' ? 'Pending' : (b.disco > 0 ? b.pct + '%' : '---')}</div>
      </div>`).join('');

    // Determine if current user can manage this team
    const curEmail = (App.state.currentEmail || '').toLowerCase();
    const curRole = App.state.currentRole;
    const teamData = Object.values(App.state.teamsData || {}).find(t => t.name === teamName);
    const teamLeaderId = teamData ? (teamData.leaderId || '').toLowerCase() : '';
    const canManage = curRole === 'superadmin' || curRole === 'owner' || curRole === 'manager' || curRole === 'admin'
      || (teamLeaderId && curEmail === teamLeaderId);

    const tabBarHTML = canManage ? `
      <div class="view-toggle-group" style="margin-top:8px">
        <button class="view-toggle active" onclick="Render.switchTeamTab('overview')" id="team-tab-btn-overview">Overview</button>
        <button class="view-toggle" onclick="Render.switchTeamTab('manage')" id="team-tab-btn-manage">Manage Team</button>
      </div>` : '';

    const safeTeamName = teamName.replace(/'/g, "\\'");

    this.openProfilePage(`
      <div class="profile-topbar">
        <button class="profile-back" onclick="Render.closeProfile()"><span>←</span> Back</button>
        <div class="profile-heading">
          <div class="profile-page-name">${team.name}</div>
          <div class="profile-page-sub">Team · <span>${members.length} members</span></div>
          ${tabBarHTML}
        </div>
      </div>

      <div id="team-tab-overview">
      <div class="profile-stats-bar">
        <div class="profile-stat"><div class="profile-stat-val units-val">${tw.u}</div><div class="profile-stat-lbl">This Wk Units</div></div>
        <div class="profile-stat"><div class="profile-stat-val">${tw.y}</div><div class="profile-stat-lbl">This Wk Yeses</div></div>
        <div class="profile-stat"><div class="profile-stat-val units-val">${Math.round(m.fourWkAvg * 4)}</div><div class="profile-stat-lbl">4W Units</div></div>
        <div class="profile-stat"><div class="profile-stat-val">${members.length}</div><div class="profile-stat-lbl">Members</div></div>
        <div class="profile-stat"><div class="profile-stat-val" style="color:${vsColor}">${vsArrow}${Math.abs(m.vsPct)}%</div><div class="profile-stat-lbl">vs 4Wk Avg</div></div>
      </div>

      <div class="profile-section-title">Team Breakdown</div>
      <div class="breakdown-grid">
        <div class="breakdown-remark" style="background:${m.remarkColor}22;border:1px solid ${m.remarkColor}55">
          <div style="font-size:10px;letter-spacing:2px;color:${m.remarkColor};text-transform:uppercase;font-family:'Barlow Condensed',sans-serif;font-weight:700;margin-bottom:6px">Remarks</div>
          <div style="font-size:18px;font-weight:700;color:${m.remarkColor}">${m.remark}</div>
          <div style="font-size:11px;color:var(--silver-dim);margin-top:6px">Rec Wk Avg: <b style="color:var(--silver)">${m.recentAvg}</b> &nbsp;·&nbsp; 4Wk Avg: <b style="color:var(--silver)">${m.fourWkAvg}</b></div>
        </div>
        <div class="breakdown-stat"><div class="breakdown-val">${m.recentAvg}</div><div class="breakdown-lbl">Rec. Wk Avg</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${vsColor}">${vsArrow}${Math.abs(m.vsPct)}%</div><div class="breakdown-lbl">vs 4Wk Avg</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${parseFloat(this.pct(m.active, total)) >= 75 ? '#22c55e' : '#f0b429'}">${this.pct(m.active, total)}%</div><div class="breakdown-lbl">Active</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:#f0b429">${this.pct(m.pending, total)}%</div><div class="breakdown-lbl">Pending</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${parseFloat(this.pct(m.cancel, total)) > 10 ? '#e53535' : 'var(--silver)'}">${this.pct(m.cancel, total)}%</div><div class="breakdown-lbl">Cancel</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${parseFloat(this.pct(m.projDisco, total)) > 5 ? '#f97316' : '#22c55e'}">${this.pct(m.projDisco, total)}%</div><div class="breakdown-lbl">Proj. Disco</div></div>
      </div>

      <div class="profile-section-title">Churn Buckets</div>
      <div class="churn-grid">${churnHTML}</div>

      ${this._subTeamBreakdownHTML(team)}

      <div style="display:flex;align-items:center;gap:12px;margin-bottom:8px;margin-top:24px">
        <div class="profile-section-title" style="margin:0">Team Leaderboard</div>
        <div class="view-toggle-group" style="margin-left:auto">
          <button class="view-toggle${this.currentView === 'days' ? ' active' : ''}" onclick="Render.setView('days')" data-view="days">Days</button>
          <button class="view-toggle${this.currentView === 'weeks' ? ' active' : ''}" onclick="Render.setView('weeks')" data-view="weeks">Weeks</button>
        </div>
      </div>
      <div class="table-scroll">
        <div class="table-wrap">
          <table id="${tableId}">
            <thead>
              <tr class="day-group"></tr>
              <tr></tr>
            </thead>
            <tbody id="${tableId}-body"></tbody>
          </table>
        </div>
      </div>

      <div class="profile-section-title">Sales Trends</div>
      <div class="charts-row">
        <div class="chart-box">
          <div class="chart-title">Sales Per Day</div>
          <div class="chart-wrap" style="height:160px"><canvas id="spd_${id}"></canvas></div>
          <div class="chart-title" style="margin-top:18px;font-size:11px">Pos / Neg Alpha</div>
          <div class="chart-wrap" style="height:100px"><canvas id="alpha_${id}"></canvas></div>
        </div>
        <div class="chart-box">
          <div class="chart-title">Time of Sales — Recent</div>
          <div class="chart-wrap" style="height:130px"><canvas id="tsr_${id}"></canvas></div>
          <div class="chart-title" style="margin-top:18px">Time of Sales — 4Wk Running</div>
          <div class="chart-wrap" style="height:130px"><canvas id="ts4_${id}"></canvas></div>
        </div>
      </div>

      <div class="profile-section-title">Products Sold</div>
      <div class="charts-row">
        <div class="chart-box">
          <div class="chart-title">Recent (2Wk)</div>
          <div class="chart-wrap" style="height:180px"><canvas id="pr_${id}"></canvas></div>
        </div>
        <div class="chart-box">
          <div class="chart-title">4Wk Running</div>
          <div class="chart-wrap" style="height:180px"><canvas id="p4_${id}"></canvas></div>
        </div>
      </div>
      </div>

      ${canManage ? `
      <div id="team-tab-manage" style="display:none">
        <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px">
          <div class="profile-section-title" style="margin:0">Team Roster</div>
          <button onclick="App.openAddMemberModal('${safeTeamName}')"
            style="margin-left:auto;background:var(--blue-core);border:none;border-radius:6px;padding:6px 16px;color:#fff;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;cursor:pointer">+ Add Member</button>
        </div>
        <input type="text" id="manage-team-search" placeholder="Search by name..."
          oninput="Render._renderManageTeamRows('${safeTeamName}', this.value)"
          style="width:100%;box-sizing:border-box;background:rgba(255,255,255,0.5);border:1px solid rgba(26,92,229,0.2);border-radius:8px;padding:8px 14px;color:var(--white);font-family:'Barlow Condensed',sans-serif;font-size:14px;outline:none;margin-bottom:12px">
        <div class="table-scroll">
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th style="padding:10px 16px;text-align:left;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:var(--silver-dim)">Name</th>
                  <th style="padding:10px 16px;text-align:left;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:var(--silver-dim)">Role</th>
                  <th style="padding:10px 16px;text-align:center;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:var(--silver-dim)">Status</th>
                </tr>
              </thead>
              <tbody id="manage-team-body"></tbody>
            </table>
          </div>
        </div>
      </div>` : ''}
    `);

    // Store team name for manage tab refresh
    this._manageTeamName = teamName;

    // Build leaderboard table (same format as office leaderboard)
    this._activeTeamLB = { tableId, members };
    this._rebuildTeamLeaderboard();

    // Pre-populate manage team rows if canManage
    if (canManage) this._renderManageTeamRows(teamName);

    setTimeout(() => this.drawCharts(id, m), 100);
  },

  // ── Manage Team tab switching ──
  _manageTeamName: null,

  switchTeamTab(tab) {
    const overview = document.getElementById('team-tab-overview');
    const manage = document.getElementById('team-tab-manage');
    const btnOverview = document.getElementById('team-tab-btn-overview');
    const btnManage = document.getElementById('team-tab-btn-manage');
    if (!overview || !manage) return;

    if (tab === 'manage') {
      overview.style.display = 'none';
      manage.style.display = '';
      if (btnOverview) btnOverview.classList.remove('active');
      if (btnManage) btnManage.classList.add('active');
      if (this._manageTeamName) this._renderManageTeamRows(this._manageTeamName);
    } else {
      overview.style.display = '';
      manage.style.display = 'none';
      if (btnOverview) btnOverview.classList.add('active');
      if (btnManage) btnManage.classList.remove('active');
    }
  },

  _renderManageTeamRows(teamName, searchText) {
    const tbody = document.getElementById('manage-team-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    const team = App.state.teams.find(t => t.name === teamName);
    if (!team) return;

    let teamMembers = (team.members || []).slice();
    if (searchText && searchText.trim()) {
      const q = searchText.trim().toLowerCase();
      teamMembers = teamMembers.filter(p => p.name.toLowerCase().includes(q));
    }

    // Sort: active first, then alphabetical
    teamMembers.sort((a, b) => {
      const aDeact = Roster.deactivated.has(a.name) ? 1 : 0;
      const bDeact = Roster.deactivated.has(b.name) ? 1 : 0;
      if (aDeact !== bDeact) return aDeact - bDeact;
      return a.name.localeCompare(b.name);
    });

    if (teamMembers.length === 0) {
      tbody.innerHTML = '<tr><td colspan="3" style="text-align:center;color:var(--silver-dim);padding:32px;font-family:\'Barlow Condensed\',sans-serif;font-size:14px">No members found</td></tr>';
      return;
    }

    teamMembers.forEach(p => {
      const isDeactivated = Roster.deactivated.has(p.name);
      const safeName = p.name.replace(/'/g, "\\'");
      const email = p.email || Roster.getEmail(p.name) || '';
      const roleKey = p._roleKey || 'rep';
      const roleLabel = OFFICE_CONFIG.roles[roleKey]?.label || roleKey;

      const tr = document.createElement('tr');
      tr.style.cssText = `border-bottom:1px solid rgba(0,0,0,0.06);${isDeactivated ? 'opacity:0.35;' : ''}`;
      tr.innerHTML = `
        <td style="padding:12px 16px">
          <div style="font-family:'Barlow Condensed',sans-serif;font-size:15px;font-weight:700;color:${isDeactivated ? 'var(--silver-dim)' : 'var(--white)'}">
            ${p.name}
            ${isDeactivated ? '<span style="font-size:10px;letter-spacing:1px;color:#e53535;margin-left:8px;border:1px solid rgba(229,53,53,0.4);border-radius:4px;padding:1px 5px;text-transform:uppercase">Inactive</span>' : ''}
          </div>
          ${email ? `<div style="font-size:10px;color:var(--silver-dim);margin-top:2px">${email}</div>` : ''}
        </td>
        <td style="padding:12px 16px">
          <span style="font-family:'Barlow Condensed',sans-serif;font-size:13px;font-weight:600;color:var(--silver)">${roleLabel}</span>
        </td>
        <td style="padding:12px 16px;text-align:center">
          <button onclick="App.toggleDeactivate('${safeName}')"
            style="background:${isDeactivated ? 'rgba(34,197,94,0.1)' : 'rgba(229,53,53,0.1)'};border:1px solid ${isDeactivated ? 'rgba(34,197,94,0.3)' : 'rgba(229,53,53,0.3)'};border-radius:6px;color:${isDeactivated ? '#22c55e' : '#e53535'};padding:5px 14px;font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;cursor:pointer;text-transform:uppercase">
            ${isDeactivated ? 'Reactivate' : 'Deactivate'}
          </button>
        </td>`;
      tbody.appendChild(tr);
    });
  },

  // Called after deactivation to refresh manage team view if visible
  _refreshManageTeam() {
    const manage = document.getElementById('team-tab-manage');
    if (manage && manage.style.display !== 'none' && this._manageTeamName) {
      const search = document.getElementById('manage-team-search');
      this._renderManageTeamRows(this._manageTeamName, search ? search.value : '');
    }
  },

  // ── Team Leaderboard rebuild (for period toggle) ──
  _activeTeamLB: null,

  _rebuildTeamLeaderboard() {
    if (!this._activeTeamLB) return;
    const { tableId, members } = this._activeTeamLB;
    this.buildTableHeaders(tableId);
    const tbody = document.getElementById(tableId + '-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    const active = members.filter(p => !Roster.deactivated.has(p.name) && !this.NON_SALES_ROLES.has(p._roleKey));
    const sorted = [...active].sort((a, b) => this.twUnits(b) - this.twUnits(a));

    sorted.forEach((p, i) => {
      const tr = document.createElement('tr');
      tr.className = this.personRowClass(p);
      tr.innerHTML = this.personRowHTML(p, i);
      tbody.appendChild(tr);
    });

    tbody.insertAdjacentHTML('beforeend', this.groupTotalRowHTML(active, 'Team Total', 'office-total-row'));
  },

  // ── Sub-Team Breakdown (for parent teams with children) ──
  _subTeamBreakdownHTML(team) {
    if (!team.children || team.children.length === 0) return '';

    const childTeams = team.children
      .map(cId => App.state.teams.find(t => t.teamId === cId))
      .filter(Boolean)
      .sort((a, b) => b.units - a.units);

    if (childTeams.length === 0) return '';

    const rows = childTeams.map((st, i) => {
      const memberCount = st.members ? st.members.length : 0;
      return `<tr style="border-bottom:1px solid rgba(0,0,0,0.04);cursor:pointer" onclick="App.openTeamProfile('${st.name.replace(/'/g, "\\'")}')">
        <td style="padding:10px 8px 10px 16px;font-family:'Barlow Condensed',sans-serif;font-size:13px;color:var(--silver-dim)">${i + 1}</td>
        <td style="padding:10px 12px"><span class="name-text name-link">${st.emoji || '⚡'} ${st.name}</span></td>
        <td style="text-align:center;padding:10px 12px;font-family:'Barlow Condensed',sans-serif;font-size:14px;color:var(--silver)">${memberCount}</td>
        <td style="text-align:center;padding:10px 12px;font-family:'Barlow Condensed',sans-serif;font-size:14px;font-weight:700;color:var(--white)">${st.units}</td>
      </tr>`;
    }).join('');

    return `
      <div class="profile-section-title">Sub-Team Breakdown</div>
      <div style="background:rgba(255,255,255,0.5);border:1px solid rgba(26,92,229,0.2);border-radius:12px;overflow:hidden;margin-bottom:24px">
        <table style="width:100%;border-collapse:collapse">
          <thead><tr style="background:rgba(0,0,0,0.06)">
            <th style="width:36px"></th>
            <th style="text-align:left;padding:10px 12px;font-family:'Barlow Condensed',sans-serif;font-size:11px;letter-spacing:2px;color:var(--silver-dim);text-transform:uppercase">Team</th>
            <th style="text-align:center;padding:10px 12px;font-family:'Barlow Condensed',sans-serif;font-size:11px;letter-spacing:2px;color:var(--silver-dim);text-transform:uppercase">Members</th>
            <th style="text-align:center;padding:10px 12px;font-family:'Barlow Condensed',sans-serif;font-size:11px;letter-spacing:2px;color:var(--silver-dim);text-transform:uppercase">Units</th>
          </tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
      <p style="font-size:11px;color:var(--silver-dim);margin-top:-18px;margin-bottom:20px;text-align:center">Click any sub-team to view its dashboard</p>`;
  },

  // ── CHARTS ──
  drawCharts(id, m) {
    const W = '#1a3a52', GRID = 'rgba(26,92,229,0.10)', TICK = '#4a7090', BLUE = '#0099cc';
    const tip = { backgroundColor: '#0d2035', borderColor: 'rgba(0,200,255,0.35)', borderWidth: 1, titleColor: '#e8f0f7', bodyColor: '#e8f0f7', padding: 8 };
    const base = {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { labels: { color: W, font: { family: 'Barlow Condensed', size: 11 }, boxWidth: 12, padding: 6 } }, tooltip: tip },
      scales: { x: { ticks: { color: TICK, font: { size: 10 } }, grid: { color: GRID } }, y: { ticks: { color: TICK, font: { size: 10 } }, grid: { color: GRID }, beginAtZero: true } }
    };
    const DAY = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

    this.mkChart('spd_' + id, { type: 'bar', data: { labels: DAY, datasets: [
      { label: 'Recent', data: m.salesPerDay, backgroundColor: BLUE + '99', borderColor: BLUE, borderWidth: 1, borderRadius: 3 },
      { label: '4Wk Avg', data: m.fourWkDaily, type: 'line', borderColor: '#4a7090', borderWidth: 2, pointBackgroundColor: '#4a7090', pointRadius: 3, fill: false, tension: 0.3 }
    ] }, options: { ...base, plugins: { ...base.plugins, legend: { ...base.plugins.legend, position: 'top' } } } });

    this.mkChart('alpha_' + id, { type: 'bar', data: { labels: DAY, datasets: [
      { label: 'Pos', data: m.alpha.map(v => v > 0 ? v : 0), backgroundColor: BLUE + 'cc', borderRadius: 2 },
      { label: 'Neg', data: m.alpha.map(v => v < 0 ? v : 0), backgroundColor: '#334155', borderRadius: 2 }
    ] }, options: { ...base, plugins: { ...base.plugins, legend: { ...base.plugins.legend, position: 'top' } },
      scales: { x: { ...base.scales.x, stacked: true }, y: { ...base.scales.y, stacked: true } } } });

    const timeOpts = { ...base, plugins: { ...base.plugins, legend: { display: false } },
      scales: { x: { ticks: { color: TICK, font: { size: 9 } }, grid: { color: GRID } }, y: { ticks: { color: TICK, font: { size: 9 }, stepSize: 1 }, grid: { color: GRID }, beginAtZero: true } } };
    this.mkChart('tsr_' + id, { type: 'bar', data: { labels: m.timeSlots, datasets: [{ data: m.recentTime, backgroundColor: BLUE + 'bb', borderRadius: 3 }] }, options: timeOpts });
    this.mkChart('ts4_' + id, { type: 'bar', data: { labels: m.timeSlots, datasets: [{ data: m.fw4Time, backgroundColor: BLUE + 'bb', borderRadius: 3 }] }, options: timeOpts });

    const PIE = ['#2f7dff', '#1e4db7', '#22c55e', '#6366f1', '#f97316', '#e53535', '#f0b429', '#0099cc'];
    const pieOpts = { responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'right', labels: { color: W, font: { family: 'Barlow Condensed', size: 11 }, padding: 8, boxWidth: 12 } }, tooltip: tip } };
    this.mkChart('pr_' + id, { type: 'doughnut', data: { labels: m.prodLabels, datasets: [{ data: m.recentProds, backgroundColor: PIE.slice(0, m.prodLabels.length), borderColor: '#e8f0f7', borderWidth: 2 }] }, options: pieOpts });
    this.mkChart('p4_' + id, { type: 'doughnut', data: { labels: m.prodLabels, datasets: [{ data: m.fw4Prods, backgroundColor: PIE.slice(0, m.prodLabels.length), borderColor: '#e8f0f7', borderWidth: 2 }] }, options: pieOpts });
  },

  // Roles that don't make sales (excluded from leaderboards/podium/stats)
  NON_SALES_ROLES: new Set(['superadmin', 'admin', 'owner']),

  // ── FULL RENDER ──
  renderAll(people, teams) {
    const salesPeople = people.filter(p => !Roster.deactivated.has(p.name) && !this.NON_SALES_ROLES.has(p._roleKey));
    this.renderHeroStats(salesPeople);
    this.makePodium('podium', salesPeople, p => this.twUnits(p), p => this.twYeses(p), p => p.role);
    this.makePodium('team-podium', teams, t => t.units, t => t.y, t => t.emoji || '⚡');
    this.renderMainTable(people);
    this.renderTeamGrid(teams);
  }
};
