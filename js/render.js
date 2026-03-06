// ═══════════════════════════════════════════════════════
// ELEVATE — Rendering Functions
// Config-driven products, no hardcoded air/cell/fiber/voip
// ═══════════════════════════════════════════════════════

const Render = {
  activeCharts: [],
  expandedPeriods: {},
  currentView: 'days', // 'days' | 'weeks'

  // Map spreadsheet color names → CSS values for churn buckets
  _CHURN_COLOR_MAP: {
    green:  { text: '#22c55e', bg: 'rgba(34, 197, 94, 0.13)',  border: 'rgba(34, 197, 94, 0.3)' },
    red:    { text: '#e53535', bg: 'rgba(229, 53, 53, 0.13)',   border: 'rgba(229, 53, 53, 0.3)' },
    yellow: { text: '#f0b429', bg: 'rgba(240, 180, 41, 0.13)',  border: 'rgba(240, 180, 41, 0.3)' },
  },
  _getChurnColor(colorName) {
    return this._CHURN_COLOR_MAP[(colorName || '').toLowerCase()] || null;
  },

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

    const active = people.filter(p => !Roster.deactivated.has(p.name) && !this._isExcludedFromLeaderboard(p));

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

  // ── Profile header bar (mirrors main header) ──
  _profileHeaderBar(info) {
    const updated = document.getElementById('last-updated-text')?.textContent || '';
    const infoHtml = info ? `
      <div style="display:flex;flex-direction:column;align-items:flex-end;gap:1px;margin-right:4px">
        <div style="font-family:'Neue Montreal','Inter',sans-serif;font-size:14px;font-weight:700;color:var(--white);letter-spacing:0.5px">${info.name}</div>
        <div style="font-family:'Helvetica Neue','Inter',sans-serif;font-size:11px;color:var(--silver-dim);letter-spacing:0.5px">${info.sub}</div>
      </div>` : '';

    // Build nav tabs based on role — grouped: Personal | Team | Office
    const role = App.state.currentRole;
    const isSA = role === 'superadmin';
    const curNav = App.state.currentNav;
    const tabs = [];

    // ── POST SALE (pinned left) ──
    if (isSA || ['rep','l1','jd','manager','owner'].includes(role)) {
      tabs.push({ label: '+ Post Sale', action: "App.navTo('postSale')", active: curNav === 'postSale' });
      tabs.push({ separator: true });
    }

    // ── PERSONAL GROUP ──
    tabs.push({ label: 'My Profile', action: "App.navTo('profile')", active: curNav === 'profile' });
    if (isSA || ['rep','l1','jd','manager','owner'].includes(role))
      tabs.push({ label: 'My Orders', action: "App.navTo('myOrders')", active: curNav === 'myOrders' });

    // ── SEPARATOR 1 (if Team group visible) ──
    const hasTeamGroup = isSA || ['jd','manager'].includes(role);
    if (hasTeamGroup) {
      tabs.push({ separator: true });
      const myTeam = Roster.getEffectiveTeam(App.state.currentPersona, App.state.people);
      let teamLabel = 'My Team';
      if (myTeam) {
        const d = Roster.getTeamDisplay(myTeam, App.state.people, App.state.teams);
        teamLabel = (d.emoji || '') + ' ' + d.name;
      }
      tabs.push({ label: teamLabel, action: "App.navTo('team')", active: curNav === 'team' });
      tabs.push({ label: 'Roster', action: "App.navTo('teamRoster')", active: curNav === 'teamRoster' });
    }

    // ── SEPARATOR 2 ──
    tabs.push({ separator: true });

    // ── OFFICE GROUP ──
    if (isSA || ['owner','admin'].includes(role))
      tabs.push({ label: 'All Orders', action: "App.navTo('allOrders')", active: curNav === 'allOrders' });
    if (isSA || ['owner','admin'].includes(role))
      tabs.push({ label: 'People', action: "App.navTo('roster')", active: curNav === 'roster' });
    if (isSA || ['owner','admin'].includes(role))
      tabs.push({ label: 'Teams', action: "App.navTo('teams')", active: curNav === 'teams' });
    tabs.push({ label: 'Leaderboard', action: "App.navTo('leaderboard')", active: curNav === 'leaderboard' });
    const curEmail = (App.state.currentEmail || '').toLowerCase();
    const payrollMgr = (App.state.settings?.payrollManager || '').toLowerCase();
    if (isSA || role === 'owner' || (payrollMgr && curEmail === payrollMgr))
      tabs.push({ label: 'Payroll', action: "App.navTo('payroll')", active: curNav === 'payroll' });
    if (isSA || role === 'owner')
      tabs.push({ label: 'Office', action: "App.navTo('office')", active: curNav === 'office' });

    const navHtml = tabs.map(t => {
      if (t.separator) return '<span class="nav-separator"></span>';
      return `<button class="nav-tab${t.active ? ' active' : ''}" ${t.action ? `onclick="${t.action}"` : ''}>${t.label}</button>`;
    }).join('');

    return `
    <header style="position:sticky;top:0;z-index:10;background:#FFFFFF;backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);padding:12px 0 0;margin:0;border-bottom:1px solid rgba(44,110,106,0.25);box-shadow:0 2px 12px rgba(0,0,0,0.06)">
      <div style="max-width:1300px;margin:0 auto;padding:0 24px">
        <div class="header-top">
          <div class="logo-area">
            <img src="references/logos/elevate-logo-full-standard-blue.png" alt="Elevate" style="height:44px;width:auto;">
            <div style="display:flex;flex-direction:column;align-items:flex-start;gap:2px;">
              <span class="week-label">Weekly Leaderboard</span>
              <div class="live-badge"><span class="live-dot"></span>Live</div>
            </div>
          </div>
          <div class="header-right">
            <span class="last-updated">${updated}</span>
            <button class="refresh-btn" onclick="App.manualRefresh()" title="Refresh data">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
                <path d="M23 4v6h-6"/>
                <path d="M1 20v-6h6"/>
                <path d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15"/>
              </svg>
            </button>
            ${infoHtml}
            <button onclick="App.logout()" title="Sign out" style="background:none;border:1px solid rgba(229,86,74,0.3);border-radius:6px;padding:5px 10px;color:#E5564A;font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:10px;font-weight:700;letter-spacing:0.5px;text-transform:uppercase;cursor:pointer">Logout</button>
          </div>
        </div>
        <nav style="display:flex;align-items:center;gap:4px;padding:0 0 10px;overflow-x:auto">
          ${navHtml}
        </nav>
      </div>
    </header>`;
  },

  // ── PROFILE PAGE ──
  openProfilePage(html, info) {
    this.activeCharts.forEach(c => { try { c.destroy(); } catch (e) {} });
    this.activeCharts = [];
    const page = document.getElementById('profile-page');
    const inner = document.getElementById('profile-page-inner');
    if (!page || !inner) return;
    // Insert header bar before the inner container, content inside inner
    const headerEl = page.querySelector('.profile-header-bar');
    if (headerEl) headerEl.remove();
    const headerDiv = document.createElement('div');
    headerDiv.className = 'profile-header-bar';
    headerDiv.innerHTML = this._profileHeaderBar(info);
    page.insertBefore(headerDiv, inner);
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
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:80px 24px;text-align:center">
          <div style="font-size:48px;margin-bottom:16px">🔒</div>
          <div style="font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:22px;font-weight:700;color:var(--white);margin-bottom:8px">Access Restricted</div>
          <div style="color:var(--silver-dim);font-size:14px;max-width:360px;line-height:1.6">You don't have permission to view this person's full profile. Contact your Junior Director or Admin.</div>
        </div>`, { name: name, sub: 'Profile Restricted' });
      return;
    }

    const p = App.state.people.find(x => x.name === name);
    if (!p) return;

    // Owner/Admin profiles — show construction placeholder
    if (p._roleKey === 'owner' || p._roleKey === 'admin') {
      this.openProfilePage(`
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:100px 24px;text-align:center">
          <div style="font-size:64px;margin-bottom:20px;opacity:0.6;filter:sepia(0.3) hue-rotate(170deg) saturate(0.5)">🚧</div>
          <div style="font-family:'Neue Montreal','Inter',sans-serif;font-size:24px;font-weight:700;color:var(--white);margin-bottom:10px">Profile Under Construction</div>
          <div style="color:var(--silver-dim);font-size:14px;max-width:400px;line-height:1.6;font-family:'Cerebri Sans','DM Sans','Inter',sans-serif">
            This profile is being built out. Check back soon for updates.
          </div>
        </div>`, { name: p.name, sub: `${p.role} · Team: ${p.team || 'Unassigned'}` });
      return;
    }

    const m = p.metrics;
    const tw4 = this.getPeriod(p, 10);
    const twU = this.twUnits(p), twY = this.twYeses(p);
    const total = Math.max(m.totalActs, 1);
    const id = 'p' + name.replace(/[^a-zA-Z0-9]/g, '');
    const vsColor = m.vsPct >= 0 ? '#2E8B57' : '#E5564A';
    const vsArrow = m.vsPct >= 0 ? '↑' : '↓';

    const churnHTML = m.churnBuckets.map(b => {
      const hasDisco = b.pct !== 'N/A' && b.disco > 0;
      const cc = this._getChurnColor(b.color);
      const pctColor = b.pct === 'N/A' ? 'var(--silver-dim)' : (cc ? cc.text : 'var(--silver-dim)');
      const cardBg = (b.pct !== 'N/A' && cc) ? `background:${cc.bg};border-color:${cc.border}` : '';
      return `<div class="churn-bucket" style="${cardBg}">
        <div class="churn-label">${b.label}</div>
        <div class="churn-pct" style="color:${pctColor}">${b.pct === 'N/A' ? '---' : (hasDisco ? b.pct + '%' : '0%')}</div>
        <div class="churn-fraction">${b.pct === 'N/A' ? '' : (hasDisco ? `(${b.disco}/${b.activated})` : `(0/${b.activated})`)}</div>
      </div>`;
    }).join('');

    // Sales Trends + Products Sold only visible when owner/superadmin views someone else's profile
    const viewerRole = App.state.currentRole;
    const isViewingOther = name !== App.state.currentPersona;
    const showCharts = isViewingOther && (viewerRole === 'owner' || viewerRole === 'superadmin');

    this.openProfilePage(`
      <div class="profile-section-title">Sales Breakdown</div>
      <div class="breakdown-grid">
        <div class="breakdown-remark" style="background:${m.remarkColor}22;border:1px solid ${m.remarkColor}55">
          <div style="font-size:10px;letter-spacing:0.5px;color:${m.remarkColor};text-transform:uppercase;font-family:'Helvetica Neue','Inter',sans-serif;font-weight:700;margin-bottom:6px">Remarks</div>
          <div style="font-size:18px;font-weight:700;color:${m.remarkColor}">${m.remark}</div>
          <div style="font-size:11px;color:var(--silver-dim);margin-top:6px">Rec Wk Avg: <b style="color:var(--silver)">${m.recentAvg}</b> &nbsp;·&nbsp; 4Wk Avg: <b style="color:var(--silver)">${m.fourWkAvg}</b></div>
        </div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${vsColor}">${vsArrow}${Math.abs(m.vsPct)}%</div><div class="breakdown-lbl">vs 4Wk Avg</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${m.monthTotalSPEs > 0 ? (m.activePct >= 85 ? '#22c55e' : m.activePct <= 70 ? '#e53535' : '#f0b429') : 'var(--silver-dim)'}">${m.monthTotalSPEs > 0 ? m.activePct + '%' : '---'}</div><div class="breakdown-lbl">Active</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${m.monthTotalSPEs > 0 ? (m.pendingPct <= 15 ? '#22c55e' : m.pendingPct >= 30 ? '#e53535' : '#f0b429') : 'var(--silver-dim)'}">${m.monthTotalSPEs > 0 ? m.pendingPct + '%' : '---'}</div><div class="breakdown-lbl">Pending</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${m.monthTotalSPEs > 0 ? (m.cancelPct <= 5 ? '#22c55e' : m.cancelPct > 10 ? '#e53535' : '#f0b429') : 'var(--silver-dim)'}">${m.monthTotalSPEs > 0 ? m.cancelPct + '%' : '---'}</div><div class="breakdown-lbl">Cancel</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${m.monthApprovedSPEs > 0 ? (m.projDiscoPct <= 2.5 ? '#22c55e' : m.projDiscoPct >= 5 ? '#e53535' : '#f0b429') : 'var(--silver-dim)'}">${m.monthApprovedSPEs > 0 ? m.projDiscoPct + '%' : '---'}</div><div class="breakdown-lbl">Proj. Disco</div></div>
      </div>

      <div class="profile-section-title">Churn Buckets</div>
      <div class="churn-grid">${churnHTML}</div>

      ${showCharts ? `
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
      </div>` : ''}
    `, { name: p.name, sub: `${p.role} · Team: ${p.team || 'Unassigned'}` });

    if (showCharts) setTimeout(() => this.drawCharts(id, m), 100);
  },

  // ── TEAM PROFILE ──
  openTeamProfile(teamName) {
    if (!Roster.canViewTeam(teamName, App.state.currentRole, App.state.currentPersona, App.state.people, App.state.teams)) {
      this.openProfilePage(`
        <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:80px 24px;text-align:center">
          <div style="font-size:48px;margin-bottom:16px">🔒</div>
          <div style="font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:22px;font-weight:700;color:var(--white);margin-bottom:8px">Access Restricted</div>
          <div style="color:var(--silver-dim);font-size:14px;max-width:360px;line-height:1.6">You can only view your own team's dashboard.</div>
        </div>`, { name: teamName, sub: 'Team Restricted' });
      return;
    }

    const team = App.state.teams.find(t => t.name === teamName);
    if (!team) return;
    const members = team.members || [];
    const directMembers = team.directMembers || members;
    const m = team.metrics;

    if (!m) {
      this.openProfilePage(`
        <p style="color:var(--silver-dim);padding:40px;text-align:center">No members assigned to this team yet.</p>`,
        { name: team.name, sub: 'Team · 0 members' });
      return;
    }

    const id = 't' + teamName.replace(/[^a-zA-Z0-9]/g, '');
    const tableId = 'team-lb-' + id;
    const vsColor = m.vsPct >= 0 ? '#2E8B57' : '#E5564A';
    const vsArrow = m.vsPct >= 0 ? '↑' : '↓';
    const total = Math.max(m.totalActs, 1);

    const churnHTML = m.churnBuckets.map(b => {
      const hasDisco = b.pct !== 'N/A' && b.disco > 0;
      const pctColor = b.pct === 'N/A' ? 'var(--silver-dim)' : (hasDisco ? '#f97316' : 'var(--silver-dim)');
      return `<div class="churn-bucket">
        <div class="churn-label">${b.label}</div>
        <div class="churn-pct" style="color:${pctColor}">${b.pct === 'N/A' ? '---' : (hasDisco ? b.pct + '%' : '0%')}</div>
        <div class="churn-fraction">${b.pct === 'N/A' ? '' : (hasDisco ? `(${b.disco}/${b.activated})` : `(0/${b.activated})`)}</div>
      </div>`;
    }).join('');

    const headcount = (team.members || []).filter(p => !Roster.deactivated.has(p.name)).length || 1;
    const perHead = m.recentAvgPerHead != null ? m.recentAvgPerHead : parseFloat((m.recentAvg / headcount).toFixed(2));
    const fw4PerHead = m.fourWkAvgPerHead != null ? m.fourWkAvgPerHead : parseFloat((m.fourWkAvg / headcount).toFixed(2));

    this.openProfilePage(`
      <div class="profile-section-title">Team Breakdown</div>
      <div class="breakdown-grid">
        <div class="breakdown-remark" style="background:${m.remarkColor}22;border:1px solid ${m.remarkColor}55">
          <div style="font-size:10px;letter-spacing:0.5px;color:${m.remarkColor};text-transform:uppercase;font-family:'Helvetica Neue','Inter',sans-serif;font-weight:700;margin-bottom:6px">Remarks</div>
          <div style="font-size:18px;font-weight:700;color:${m.remarkColor}">${m.remark}</div>
          <div style="font-size:11px;color:var(--silver-dim);margin-top:6px">Team Avg: <b style="color:var(--silver)">${m.recentAvg}</b> &nbsp;·&nbsp; Per Rep: <b style="color:var(--silver)">${perHead}</b> &nbsp;·&nbsp; 4Wk: <b style="color:var(--silver)">${m.fourWkAvg}</b></div>
        </div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${vsColor}">${vsArrow}${Math.abs(m.vsPct)}%</div><div class="breakdown-lbl">vs 4Wk Avg</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${m.monthTotalSPEs > 0 ? (m.activePct >= 85 ? '#22c55e' : m.activePct <= 70 ? '#e53535' : '#f0b429') : 'var(--silver-dim)'}">${m.monthTotalSPEs > 0 ? m.activePct + '%' : '---'}</div><div class="breakdown-lbl">Active</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${m.monthTotalSPEs > 0 ? (m.pendingPct <= 15 ? '#22c55e' : m.pendingPct >= 30 ? '#e53535' : '#f0b429') : 'var(--silver-dim)'}">${m.monthTotalSPEs > 0 ? m.pendingPct + '%' : '---'}</div><div class="breakdown-lbl">Pending</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${m.monthTotalSPEs > 0 ? (m.cancelPct <= 5 ? '#22c55e' : m.cancelPct > 10 ? '#e53535' : '#f0b429') : 'var(--silver-dim)'}">${m.monthTotalSPEs > 0 ? m.cancelPct + '%' : '---'}</div><div class="breakdown-lbl">Cancel</div></div>
        <div class="breakdown-stat"><div class="breakdown-val" style="color:${m.monthApprovedSPEs > 0 ? (m.projDiscoPct <= 2.5 ? '#22c55e' : m.projDiscoPct >= 5 ? '#e53535' : '#f0b429') : 'var(--silver-dim)'}">${m.monthApprovedSPEs > 0 ? m.projDiscoPct + '%' : '---'}</div><div class="breakdown-lbl">Proj. Disco</div></div>
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
    `, { name: team.name, sub: `Team · ${members.length} members` });

    // Build leaderboard table (same format as office leaderboard)
    this._activeTeamLB = { tableId, members };
    this._rebuildTeamLeaderboard();

    setTimeout(() => this.drawCharts(id, m), 100);
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

    const active = members.filter(p => !Roster.deactivated.has(p.name) && !this._isExcludedFromLeaderboard(p));
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
        <td style="padding:10px 8px 10px 16px;font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:13px;color:var(--silver-dim)">${i + 1}</td>
        <td style="padding:10px 12px"><span class="name-text name-link">${st.emoji || '⚡'} ${st.name}</span></td>
        <td style="text-align:center;padding:10px 12px;font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:14px;color:var(--silver)">${memberCount}</td>
        <td style="text-align:center;padding:10px 12px;font-family:'Neue Montreal','Inter',sans-serif;font-size:14px;font-weight:700;color:var(--white)">${st.units}</td>
      </tr>`;
    }).join('');

    return `
      <div class="profile-section-title">Sub-Team Breakdown</div>
      <div style="background:rgba(255,255,255,0.5);border:1px solid rgba(0,0,0,0.2);border-radius:12px;overflow:hidden;margin-bottom:24px">
        <table style="width:100%;border-collapse:collapse">
          <thead><tr style="background:rgba(0,0,0,0.06)">
            <th style="width:36px"></th>
            <th style="text-align:left;padding:10px 12px;font-family:'Helvetica Neue','Inter',sans-serif;font-size:11px;letter-spacing:0.5px;color:var(--silver-dim);text-transform:uppercase">Team</th>
            <th style="text-align:center;padding:10px 12px;font-family:'Helvetica Neue','Inter',sans-serif;font-size:11px;letter-spacing:0.5px;color:var(--silver-dim);text-transform:uppercase">Members</th>
            <th style="text-align:center;padding:10px 12px;font-family:'Helvetica Neue','Inter',sans-serif;font-size:11px;letter-spacing:0.5px;color:var(--silver-dim);text-transform:uppercase">Units</th>
          </tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
      <p style="font-size:11px;color:var(--silver-dim);margin-top:-18px;margin-bottom:20px;text-align:center">Click any sub-team to view its dashboard</p>`;
  },

  // ── CHARTS ──
  drawCharts(id, m) {
    const W = '#4A5568', GRID = 'rgba(0,0,0,0.10)', TICK = '#708090', BLUE = '#2C6E6A';
    const tip = { backgroundColor: '#242124', borderColor: 'rgba(44,110,106,0.35)', borderWidth: 1, titleColor: '#FFFFFF', bodyColor: '#FFFFFF', padding: 8 };
    const base = {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { labels: { color: W, font: { family: "'Cerebri Sans','DM Sans','Inter',sans-serif", size: 11 }, boxWidth: 12, padding: 6 } }, tooltip: tip },
      scales: { x: { ticks: { color: TICK, font: { size: 10 } }, grid: { color: GRID } }, y: { ticks: { color: TICK, font: { size: 10 } }, grid: { color: GRID }, beginAtZero: true } }
    };
    const DAY = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

    this.mkChart('spd_' + id, { type: 'bar', data: { labels: DAY, datasets: [
      { label: 'Recent', data: m.salesPerDay, backgroundColor: BLUE + '99', borderColor: BLUE, borderWidth: 1, borderRadius: 3 },
      { label: '4Wk Avg', data: m.fourWkDaily, type: 'line', borderColor: '#708090', borderWidth: 2, pointBackgroundColor: '#708090', pointRadius: 3, fill: false, tension: 0.3 }
    ] }, options: { ...base, plugins: { ...base.plugins, legend: { ...base.plugins.legend, position: 'top' } } } });

    this.mkChart('alpha_' + id, { type: 'bar', data: { labels: DAY, datasets: [
      { label: 'Pos', data: m.alpha.map(v => v > 0 ? v : 0), backgroundColor: BLUE + 'cc', borderRadius: 2 },
      { label: 'Neg', data: m.alpha.map(v => v < 0 ? v : 0), backgroundColor: '#9AA5B1', borderRadius: 2 }
    ] }, options: { ...base, plugins: { ...base.plugins, legend: { ...base.plugins.legend, position: 'top' } },
      scales: { x: { ...base.scales.x, stacked: true }, y: { ...base.scales.y, stacked: true } } } });

    const timeOpts = { ...base, plugins: { ...base.plugins, legend: { display: false } },
      scales: { x: { ticks: { color: TICK, font: { size: 9 } }, grid: { color: GRID } }, y: { ticks: { color: TICK, font: { size: 9 }, stepSize: 1 }, grid: { color: GRID }, beginAtZero: true } } };
    this.mkChart('tsr_' + id, { type: 'bar', data: { labels: m.timeSlots, datasets: [{ data: m.recentTime, backgroundColor: BLUE + 'bb', borderRadius: 3 }] }, options: timeOpts });
    this.mkChart('ts4_' + id, { type: 'bar', data: { labels: m.timeSlots, datasets: [{ data: m.fw4Time, backgroundColor: BLUE + 'bb', borderRadius: 3 }] }, options: timeOpts });

    const PIE = ['#2C6E6A', '#1E4F46', '#2E8B57', '#6366f1', '#f97316', '#E5564A', '#f0b429', '#2C6E6A'];
    const pieOpts = { responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'right', labels: { color: W, font: { family: "'Cerebri Sans','DM Sans','Inter',sans-serif", size: 11 }, padding: 8, boxWidth: 12 } }, tooltip: tip } };
    this.mkChart('pr_' + id, { type: 'doughnut', data: { labels: m.prodLabels, datasets: [{ data: m.recentProds, backgroundColor: PIE.slice(0, m.prodLabels.length), borderColor: '#FFFFFF', borderWidth: 2 }] }, options: pieOpts });
    this.mkChart('p4_' + id, { type: 'doughnut', data: { labels: m.prodLabels, datasets: [{ data: m.fw4Prods, backgroundColor: PIE.slice(0, m.prodLabels.length), borderColor: '#FFFFFF', borderWidth: 2 }] }, options: pieOpts });
  },

  // Roles that never appear on leaderboards/podium/stats
  NON_SALES_ROLES: new Set(['superadmin', 'admin']),

  // Owners only appear on leaderboard if actively selling
  _isExcludedFromLeaderboard(p) {
    if (this.NON_SALES_ROLES.has(p._roleKey)) return true;
    if (p._roleKey === 'owner') return this.twUnits(p) === 0;
    return false;
  },

  // ── FULL RENDER ──
  renderAll(people, teams) {
    const salesPeople = people.filter(p => !Roster.deactivated.has(p.name) && !this._isExcludedFromLeaderboard(p));
    this.renderHeroStats(salesPeople);
    this.makePodium('podium', salesPeople, p => this.twUnits(p), p => this.twYeses(p), p => p.role);
    this.makePodium('team-podium', teams, t => t.units, t => t.y, t => t.emoji || '⚡');
    this.renderMainTable(people);
    this.renderTeamGrid(teams);
  }
};
