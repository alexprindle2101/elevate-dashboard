// ═══════════════════════════════════════════════════════
// CHALLENGE — Office-wide gamified sales challenges
// Owner configures point rules & penalties, reps compete
// ═══════════════════════════════════════════════════════

const Challenge = {

  // ── State ──
  _config: null,       // challenge config JSON (or null)
  _sales: null,        // { email: { dailyUnits: {}, totalUnits } }
  _blood: null,        // { "YYYY-MM-DD": { firstBlood, lastBlood } }
  _leaderboard: [],    // computed leaderboard rows
  _teamLeaderboard: [], // computed team leaderboard
  _loading: false,
  _leaderboardView: 'teams', // 'teams' or 'individual'
  _teamAssignActive: false,  // team assignment UI open
  _teamAssignData: [],       // working copy of teams during assignment

  // ── Wizard state ──
  _wizardActive: false,
  _wizardStep: 1,
  _wizardData: {},

  // ── Helpers ──
  _esc(s) { return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); },

  _isOwnerOrSuperadmin() {
    const role = App.state.currentRole;
    return role === 'owner' || role === 'superadmin';
  },

  _getChallengeStatus() {
    if (!this._config || this._config.status === 'ended') return 'ended';
    const today = new Date(); today.setHours(0,0,0,0);
    const end = new Date(this._config.endDate); end.setHours(23,59,59,999);
    if (today > end) return 'ended';
    const start = new Date(this._config.startDate); start.setHours(0,0,0,0);
    if (today < start) return 'upcoming';
    return 'active';
  },

  _daysRemaining() {
    if (!this._config) return 0;
    const today = new Date(); today.setHours(0,0,0,0);
    const end = new Date(this._config.endDate); end.setHours(0,0,0,0);
    return Math.max(0, Math.ceil((end - today) / 86400000));
  },

  // ═════════════════════════════════════════
  // ENTRY POINT
  // ═════════════════════════════════════════

  async open() {
    const page = document.getElementById('challenge-page');
    if (page) page.style.display = 'block';

    const subtitle = document.getElementById('challenge-subtitle');
    if (subtitle) subtitle.textContent = 'Loading...';

    if (this._wizardActive) {
      this._renderWizard();
      return;
    }

    if (this._teamAssignActive) {
      this._renderTeamAssignment();
      return;
    }

    this._loading = true;
    try {
      this._config = await SheetsAPI.fetchChallengeConfig(OFFICE_CONFIG);

      if (this._config && this._config.status !== 'ended') {
        const [sales, blood] = await Promise.all([
          SheetsAPI.fetchChallengeSales(OFFICE_CONFIG, this._config.startDate, this._config.endDate),
          SheetsAPI.fetchChallengeBlood(OFFICE_CONFIG)
        ]);
        this._sales = sales;
        this._blood = blood;

        // Auto-trigger blood calc for missing dates
        this._autoCalculateBlood();
      }
    } catch (err) {
      console.error('[Challenge] Failed to load:', err);
    } finally {
      this._loading = false;
    }

    // Guard: user navigated away during fetch
    if (App.state.currentNav !== 'challenge') return;

    this._render();
  },

  // ═════════════════════════════════════════
  // RENDER DISPATCH
  // ═════════════════════════════════════════

  _render() {
    const subtitle = document.getElementById('challenge-subtitle');
    const content = document.getElementById('challenge-content');
    if (!content) return;

    if (!this._config || this._config.status === 'ended') {
      if (subtitle) subtitle.textContent = '';
      this._renderNoChallenge(content);
      return;
    }

    const status = this._getChallengeStatus();
    const days = this._daysRemaining();
    if (subtitle) {
      if (status === 'upcoming') subtitle.textContent = `${this._esc(this._config.name)} — Starts ${this._config.startDate}`;
      else if (status === 'active') subtitle.textContent = `${this._esc(this._config.name)} — ${days} day${days !== 1 ? 's' : ''} remaining`;
      else subtitle.textContent = `${this._esc(this._config.name)} — Ended`;
    }

    this._computeLeaderboard();
    if (this._config.mode === 'teams') this._computeTeamLeaderboard();

    let html = '';
    if (this._isOwnerOrSuperadmin()) html += this._managementControlsHTML();

    const isTeamMode = this._config.mode === 'teams';
    const hasTeams = isTeamMode && this._config.teams && this._config.teams.length > 0;

    if (isTeamMode && !hasTeams) {
      html += '<div class="challenge-banner">Teams mode is active but teams haven\'t been assigned yet.</div>';
    }

    if (hasTeams) {
      html += this._viewToggleHTML();
      html += (this._leaderboardView === 'teams') ? this._teamLeaderboardHTML() : this._individualLeaderboardHTML(true);
    } else {
      html += this._individualLeaderboardHTML(false);
    }
    content.innerHTML = html;
  },

  // ═════════════════════════════════════════
  // NO CHALLENGE STATE
  // ═════════════════════════════════════════

  _renderNoChallenge(container) {
    const isOwner = this._isOwnerOrSuperadmin();
    container.innerHTML = `
      <div class="challenge-empty">
        <div style="font-size:48px;margin-bottom:16px">🏆</div>
        <div style="font-family:'Neue Montreal','Inter',sans-serif;font-size:22px;font-weight:700;color:var(--white);margin-bottom:8px">No Active Challenge</div>
        <div style="font-size:14px;color:var(--silver-dim);margin-bottom:24px">${isOwner ? 'Start a challenge to get your team competing.' : 'Check back soon — your owner will announce the next challenge.'}</div>
        ${isOwner ? '<button class="challenge-btn-primary" onclick="Challenge._startWizard()">Start Challenge</button>' : ''}
      </div>`;
  },

  // ═════════════════════════════════════════
  // MANAGEMENT CONTROLS (owner/superadmin)
  // ═════════════════════════════════════════

  _managementControlsHTML() {
    const c = this._config;
    const rules = c.rules || {};
    const enabledRules = Object.entries(rules).filter(([,v]) => v.enabled).map(([k]) => {
      const labels = {
        pointsPerUnit: 'Points/Unit', dailyGoals: 'Daily Goals', eventGoals: 'Event Goals',
        firstBlood: 'First Blood', lastBlood: 'Last Blood',
        activePenalty: 'Active % Penalty', churn030Penalty: '0-30 Churn Penalty'
      };
      return labels[k] || k;
    });

    return `
      <div class="challenge-mgmt">
        <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px">
          <div>
            <div style="font-size:11px;color:var(--silver-dim);font-family:'Helvetica Neue','Inter',sans-serif;letter-spacing:0.5px;text-transform:uppercase;margin-bottom:4px">Active Rules</div>
            <div style="display:flex;flex-wrap:wrap;gap:4px">${enabledRules.map(r => `<span class="challenge-rule-badge">${r}</span>`).join('')}</div>
          </div>
          <div style="display:flex;gap:8px;flex-wrap:wrap">
            ${c.mode === 'teams' ? (c.teams && c.teams.length > 0
              ? '<button class="challenge-btn-secondary" onclick="Challenge._openTeamAssignment()">Edit Teams</button>'
              : '<button class="challenge-btn-primary" onclick="Challenge._openTeamAssignment()">Set Up Teams</button>') : ''}
            <button class="challenge-btn-secondary" onclick="Challenge._recalcAllBlood()">Recalc Blood</button>
            <button class="challenge-btn-danger" onclick="Challenge._confirmEndChallenge()">End Challenge</button>
          </div>
        </div>
      </div>`;
  },

  // ═════════════════════════════════════════
  // LEADERBOARD
  // ═════════════════════════════════════════

  // ── View toggle (Teams | Individual) ──
  _viewToggleHTML() {
    const tCls = this._leaderboardView === 'teams' ? 'challenge-view-btn active' : 'challenge-view-btn';
    const iCls = this._leaderboardView === 'individual' ? 'challenge-view-btn active' : 'challenge-view-btn';
    return `<div class="challenge-view-toggle">
      <button class="${tCls}" onclick="Challenge._setView('teams')">Teams</button>
      <button class="${iCls}" onclick="Challenge._setView('individual')">Individual</button>
    </div>`;
  },

  _setView(view) {
    this._leaderboardView = view;
    this._render();
  },

  // ── Individual leaderboard ──
  _individualLeaderboardHTML(showTeamCol) {
    if (!this._leaderboard.length) {
      return '<div style="text-align:center;color:var(--silver-dim);padding:48px;font-size:14px">No sales data yet for this challenge period.</div>';
    }

    // Build email→team lookup
    const teamLookup = {};
    if (showTeamCol && this._config.teams) {
      this._config.teams.forEach(t => {
        (t.members || []).forEach(email => { teamLookup[email] = `${t.emoji || ''} ${t.name}`.trim(); });
      });
    }

    const myEmail = (App.state.currentEmail || '').toLowerCase();
    const rules = this._config.rules || {};
    let rows = '';
    this._leaderboard.forEach((r, i) => {
      const rank = i + 1;
      const medal = rank === 1 ? '🥇' : rank === 2 ? '🥈' : rank === 3 ? '🥉' : rank;
      const isMe = r.email === myEmail;

      // Build tooltip for unit pts
      const unitTip = r.rawUnitPoints !== r.unitPoints
        ? `${r.totalUnits} units × ${Number(rules.pointsPerUnit?.points) || 0} pts = ${r.rawUnitPoints} (${r.unitPoints} after penalties)`
        : `${r.totalUnits} units × ${Number(rules.pointsPerUnit?.points) || 0} pts`;

      // Build tooltip for goal pts
      const goalParts = [];
      if (r.dailyGoalPts > 0) goalParts.push(`Daily Goals: ${r.dailyGoalPts} pts`);
      if (r.eventGoalPts > 0) goalParts.push(`Event Goals: ${r.eventGoalPts} pts`);
      const goalTip = goalParts.length ? goalParts.join(' + ') : 'No goal thresholds reached';

      // Build tooltip for bonus pts
      const bonusParts = [];
      if (r.firstBloodWins > 0) bonusParts.push(`First Blood: ${r.firstBloodWins} day${r.firstBloodWins !== 1 ? 's' : ''} × ${Number(rules.firstBlood?.points) || 0} pts`);
      if (r.lastBloodWins > 0) bonusParts.push(`Last Blood: ${r.lastBloodWins} day${r.lastBloodWins !== 1 ? 's' : ''} × ${Number(rules.lastBlood?.points) || 0} pts`);
      const bonusTip = bonusParts.length ? bonusParts.join(' + ') : 'No blood wins yet';

      // Penalty display
      const penaltyDeduction = r.rawUnitPoints - r.unitPoints;
      const penaltyStr = r.penaltyPct > 0
        ? `<span class="challenge-penalty" title="−${penaltyDeduction} pts from ${r.rawUnitPoints} unit pts (${r.penaltyPct.toFixed(1)}% penalty)">-${penaltyDeduction}</span>`
        : '<span style="color:var(--silver-dim)">—</span>';

      const teamCell = showTeamCol ? `<td class="challenge-pts" style="font-size:12px">${this._esc(teamLookup[r.email] || '—')}</td>` : '';

      // Total tooltip
      const totalParts = [];
      if (r.unitPoints > 0) totalParts.push(`Unit: ${r.unitPoints}`);
      if (r.goalPoints > 0) totalParts.push(`Goals: ${r.goalPoints}`);
      if (r.bonusPoints > 0) totalParts.push(`Bonus: ${r.bonusPoints}`);
      const totalTip = totalParts.join(' + ') + ` = ${r.total}`;

      const tip = (text, content) => `<span class="challenge-tip-wrap">${content}<span class="challenge-tip">${text}</span></span>`;

      rows += `<tr class="${isMe ? 'challenge-row-me' : ''}">
        <td class="challenge-rank">${medal}</td>
        <td class="challenge-name">${this._esc(r.name)}</td>
        ${teamCell}
        <td class="challenge-pts">${tip(this._esc(unitTip), r.unitPoints)}</td>
        <td class="challenge-pts">${tip(this._esc(goalTip), r.goalPoints)}</td>
        <td class="challenge-pts">${tip(this._esc(bonusTip), r.bonusPoints)}</td>
        <td class="challenge-pts">${penaltyStr}</td>
        <td class="challenge-pts challenge-total">${tip(this._esc(totalTip), r.total)}</td>
      </tr>`;
    });

    const teamTh = showTeamCol ? '<th>Team</th>' : '';
    return `
      <div style="overflow-x:auto;border-radius:10px;border:1px solid rgba(0,0,0,0.15)">
        <table class="challenge-table">
          <thead>
            <tr>
              <th style="width:48px">#</th>
              <th style="text-align:left">Rep</th>
              ${teamTh}
              <th>Unit Pts</th>
              <th>Goal Pts</th>
              <th>Bonus Pts</th>
              <th>Penalty</th>
              <th>Score</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;
  },

  // ── Team leaderboard ──
  _teamLeaderboardHTML() {
    if (!this._teamLeaderboard.length) {
      return '<div style="text-align:center;color:var(--silver-dim);padding:48px;font-size:14px">No team data yet.</div>';
    }

    const myEmail = (App.state.currentEmail || '').toLowerCase();
    let html = '';
    this._teamLeaderboard.forEach((team, i) => {
      const rank = i + 1;
      const medal = rank === 1 ? '🥇' : rank === 2 ? '🥈' : rank === 3 ? '🥉' : rank;
      const isMyTeam = team.members.some(m => m.email === myEmail);

      // Member rows (collapsed by default, toggle with onclick)
      let memberRows = '';
      team.members.sort((a, b) => b.total - a.total).forEach(m => {
        const meClass = m.email === myEmail ? ' challenge-row-me' : '';
        memberRows += `<tr class="challenge-team-member${meClass}">
          <td></td>
          <td class="challenge-name" style="padding-left:32px;font-size:13px">${this._esc(m.name)}</td>
          <td class="challenge-pts">${m.unitPoints}</td>
          <td class="challenge-pts">${m.goalPoints}</td>
          <td class="challenge-pts">${m.bonusPoints}</td>
          <td class="challenge-pts challenge-total">${m.total}</td>
        </tr>`;
      });

      html += `
        <div class="challenge-team-card ${isMyTeam ? 'challenge-row-me' : ''}" onclick="Challenge._toggleTeamExpand(${i})" style="cursor:pointer">
          <div style="display:flex;align-items:center;justify-content:space-between">
            <div style="display:flex;align-items:center;gap:12px">
              <span class="challenge-rank" style="font-size:18px">${medal}</span>
              <span style="font-size:20px">${team.emoji || ''}</span>
              <div>
                <div style="font-family:'Neue Montreal','Inter',sans-serif;font-size:16px;font-weight:700;color:var(--white)">${this._esc(team.name)}</div>
                <div style="font-size:12px;color:var(--silver-dim)">${team.members.length} member${team.members.length !== 1 ? 's' : ''}</div>
              </div>
            </div>
            <div style="font-family:'Neue Montreal','Inter',sans-serif;font-size:24px;font-weight:800;color:var(--sc-teal)">${team.total}</div>
          </div>
          <div class="challenge-team-members" id="challenge-team-expand-${i}" style="display:none;margin-top:12px">
            <table class="challenge-table" style="border:none">
              <thead><tr>
                <th style="width:32px"></th>
                <th style="text-align:left">Rep</th>
                <th>Unit Pts</th>
                <th>Goal Pts</th>
                <th>Bonus Pts</th>
                <th>Total</th>
              </tr></thead>
              <tbody>${memberRows}</tbody>
            </table>
          </div>
        </div>`;
    });

    return html;
  },

  _toggleTeamExpand(index) {
    const el = document.getElementById(`challenge-team-expand-${index}`);
    if (el) el.style.display = el.style.display === 'none' ? 'block' : 'none';
  },

  // ═════════════════════════════════════════
  // POINT CALCULATION ENGINE
  // ═════════════════════════════════════════

  _computeLeaderboard() {
    this._leaderboard = [];
    if (!this._config || !this._sales) return;

    const rules = this._config.rules || {};
    const competingRoles = new Set(this._config.competingRoles || ['rep', 'l1', 'jd']);
    const people = App.state.people || [];

    people.forEach(p => {
      if (!competingRoles.has(p._roleKey)) return;
      const email = (Roster.getEmail(p.name) || '').toLowerCase();
      if (!email) return;

      const repSales = this._sales[email] || { dailyUnits: {}, totalUnits: 0 };

      const num = (v) => Number(v) || 0;

      // Unit points
      let rawUnitPts = 0;
      if (rules.pointsPerUnit && rules.pointsPerUnit.enabled) {
        rawUnitPts = repSales.totalUnits * num(rules.pointsPerUnit.points);
      }

      // Daily goal points (highest tier per day)
      let dailyGoalPts = 0;
      if (rules.dailyGoals && rules.dailyGoals.enabled && rules.dailyGoals.tiers) {
        const tiers = [...rules.dailyGoals.tiers].sort((a, b) => num(b.threshold) - num(a.threshold));
        Object.values(repSales.dailyUnits).forEach(dayUnits => {
          const tier = tiers.find(t => dayUnits >= num(t.threshold));
          if (tier) dailyGoalPts += num(tier.points);
        });
      }

      // Event goal points (highest tier, lifetime total)
      let eventGoalPts = 0;
      if (rules.eventGoals && rules.eventGoals.enabled && rules.eventGoals.tiers) {
        const tiers = [...rules.eventGoals.tiers].sort((a, b) => num(b.threshold) - num(a.threshold));
        const tier = tiers.find(t => repSales.totalUnits >= num(t.threshold));
        if (tier) eventGoalPts = num(tier.points);
      }

      // Blood points
      let bloodPts = 0;
      if (this._blood) {
        if (rules.firstBlood && rules.firstBlood.enabled) {
          Object.values(this._blood).forEach(day => {
            if (day.firstBlood && day.firstBlood.email === email) bloodPts += num(rules.firstBlood.points);
          });
        }
        if (rules.lastBlood && rules.lastBlood.enabled) {
          Object.values(this._blood).forEach(day => {
            if (day.lastBlood && day.lastBlood.email === email) bloodPts += num(rules.lastBlood.points);
          });
        }
      }

      // Penalties (applied to unit points only, stacked multiplicatively)
      let penaltyMultiplier = 1;
      const m = p.metrics || {};

      // Active % penalty: straight deduction — 75% active = 25% penalty
      if (rules.activePenalty && rules.activePenalty.enabled && m.monthTotalSPEs > 0) {
        const activeFrac = num(m.activePct) / 100;
        penaltyMultiplier *= activeFrac;
      }

      // 0-30 Day Churn penalty: (churnPct)^2 / 100
      if (rules.churn030Penalty && rules.churn030Penalty.enabled && m.churnBuckets && m.churnBuckets[0]) {
        const churnPct = num(m.churnBuckets[0].pct);
        const penalty = Math.min((churnPct * churnPct) / 100, 1);
        penaltyMultiplier *= (1 - penalty);
      }

      // 30 Day Churn penalty
      if (rules.churn30Penalty && rules.churn30Penalty.enabled && m.churnBuckets && m.churnBuckets[1]) {
        const churnPct = num(m.churnBuckets[1].pct);
        const penalty = Math.min((churnPct * churnPct) / 100, 1);
        penaltyMultiplier *= (1 - penalty);
      }

      const adjustedUnitPts = Math.max(0, Math.round(rawUnitPts * penaltyMultiplier));
      const penaltyPct = rawUnitPts > 0 ? ((1 - penaltyMultiplier) * 100) : 0;

      const goalPoints = dailyGoalPts + eventGoalPts;
      const total = adjustedUnitPts + goalPoints + bloodPts;

      // Count blood wins
      let firstBloodWins = 0, lastBloodWins = 0;
      if (this._blood) {
        Object.values(this._blood).forEach(day => {
          if (day.firstBlood && day.firstBlood.email === email) firstBloodWins++;
          if (day.lastBlood && day.lastBlood.email === email) lastBloodWins++;
        });
      }

      this._leaderboard.push({
        name: p.name,
        email: email,
        unitPoints: adjustedUnitPts,
        rawUnitPoints: rawUnitPts,
        totalUnits: repSales.totalUnits,
        dailyGoalPts: dailyGoalPts,
        eventGoalPts: eventGoalPts,
        goalPoints: goalPoints,
        bonusPoints: bloodPts,
        firstBloodWins: firstBloodWins,
        lastBloodWins: lastBloodWins,
        penaltyPct: penaltyPct,
        total: total
      });
    });

    // Sort by total desc, then by name
    this._leaderboard.sort((a, b) => b.total - a.total || a.name.localeCompare(b.name));
  },

  _computeTeamLeaderboard() {
    this._teamLeaderboard = [];
    if (!this._config || !this._config.teams) return;

    const emailToRow = {};
    this._leaderboard.forEach(r => { emailToRow[r.email] = r; });

    this._config.teams.forEach(team => {
      const members = (team.members || []).map(email => emailToRow[email]).filter(Boolean);
      const total = members.reduce((sum, m) => sum + m.total, 0);
      this._teamLeaderboard.push({
        name: team.name,
        emoji: team.emoji || '',
        members: members,
        total: total
      });
    });

    this._teamLeaderboard.sort((a, b) => b.total - a.total || a.name.localeCompare(b.name));
  },

  // ═════════════════════════════════════════
  // BLOOD AUTO-CALCULATION
  // ═════════════════════════════════════════

  async _autoCalculateBlood() {
    if (!this._config || !this._blood) return;
    const start = new Date(this._config.startDate); start.setHours(0,0,0,0);
    const yesterday = new Date(); yesterday.setDate(yesterday.getDate() - 1); yesterday.setHours(0,0,0,0);
    const end = new Date(this._config.endDate); end.setHours(0,0,0,0);
    const limit = yesterday < end ? yesterday : end;

    const missing = [];
    for (let d = new Date(start); d <= limit; d.setDate(d.getDate() + 1)) {
      const key = d.toISOString().split('T')[0];
      if (!this._blood[key]) missing.push(key);
    }

    if (missing.length === 0) return;
    console.log(`[Challenge] Auto-calculating blood for ${missing.length} missing date(s)`);

    // Fire and forget — don't block the UI
    for (const date of missing) {
      try {
        const result = await SheetsAPI.post(OFFICE_CONFIG, 'calculateBlood', { date });
        if (result && result.blood) this._blood[date] = result.blood;
      } catch (err) {
        console.error(`[Challenge] Blood calc failed for ${date}:`, err);
      }
    }
  },

  async _recalcAllBlood() {
    if (!this._config) return;
    const start = new Date(this._config.startDate); start.setHours(0,0,0,0);
    const yesterday = new Date(); yesterday.setDate(yesterday.getDate() - 1); yesterday.setHours(0,0,0,0);
    const end = new Date(this._config.endDate); end.setHours(0,0,0,0);
    const limit = yesterday < end ? yesterday : end;

    const subtitle = document.getElementById('challenge-subtitle');
    if (subtitle) subtitle.textContent = 'Recalculating blood...';

    for (let d = new Date(start); d <= limit; d.setDate(d.getDate() + 1)) {
      const key = d.toISOString().split('T')[0];
      try {
        const result = await SheetsAPI.post(OFFICE_CONFIG, 'calculateBlood', { date: key });
        if (result && result.blood) this._blood[key] = result.blood;
      } catch (err) {
        console.error(`[Challenge] Blood recalc failed for ${key}:`, err);
      }
    }

    if (App.state.currentNav === 'challenge') this._render();
  },

  // ═════════════════════════════════════════
  // END CHALLENGE
  // ═════════════════════════════════════════

  async _confirmEndChallenge() {
    if (!confirm('End this challenge? Final standings will be frozen.')) return;
    try {
      await SheetsAPI.post(OFFICE_CONFIG, 'endChallenge', {});
      this._config.status = 'ended';
      App.state.challengeConfig = this._config;
      App.updateNav();
      if (App.state.currentNav === 'challenge') this._render();
    } catch (err) {
      alert('Failed to end challenge: ' + err.message);
    }
  },

  // ═════════════════════════════════════════
  // SETUP WIZARD
  // ═════════════════════════════════════════

  _startWizard() {
    this._wizardActive = true;
    this._wizardStep = 1;
    this._wizardData = {
      name: '',
      startDate: new Date().toISOString().split('T')[0],
      endDate: '',
      mode: 'ffa',
      competingRoles: ['rep', 'l1', 'jd'],
      rules: {
        pointsPerUnit: { enabled: true, points: 100 },
        dailyGoals: { enabled: false, tiers: [{ threshold: 1, points: 50 }, { threshold: 2, points: 150 }, { threshold: 3, points: 300 }] },
        eventGoals: { enabled: false, tiers: [{ threshold: 10, points: 500 }, { threshold: 20, points: 1500 }] },
        firstBlood: { enabled: false, points: 200 },
        lastBlood: { enabled: false, points: 200 },
        activePenalty: { enabled: false },
        churn030Penalty: { enabled: false }
      }
    };
    this._renderWizard();
  },

  _renderWizard() {
    const subtitle = document.getElementById('challenge-subtitle');
    if (subtitle) subtitle.textContent = `Step ${this._wizardStep} of 3`;

    const content = document.getElementById('challenge-content');
    if (!content) return;

    // Step dots
    let dots = '';
    for (let i = 1; i <= 3; i++) {
      const cls = i === this._wizardStep ? 'wizard-step-dot active' : (i < this._wizardStep ? 'wizard-step-dot done' : 'wizard-step-dot');
      dots += `<div class="${cls}">${i}</div>`;
      if (i < 3) dots += '<div class="wizard-step-line"></div>';
    }

    let body = '';
    if (this._wizardStep === 1) body = this._wizardStep1HTML();
    else if (this._wizardStep === 2) body = this._wizardStep2HTML();
    else body = this._wizardStep3HTML();

    const backBtn = this._wizardStep > 1
      ? `<button class="challenge-btn-secondary" onclick="Challenge._wizardBack()">Back</button>`
      : `<button class="challenge-btn-secondary" onclick="Challenge._cancelWizard()">Cancel</button>`;
    const nextBtn = this._wizardStep < 3
      ? `<button class="challenge-btn-primary" onclick="Challenge._wizardNext()">Next</button>`
      : `<button class="challenge-btn-primary" onclick="Challenge._wizardSubmit()">Launch Challenge</button>`;

    content.innerHTML = `
      <div style="max-width:600px;margin:0 auto">
        <div class="wizard-steps" style="margin-bottom:32px">${dots}</div>
        <div class="wizard-body">${body}</div>
        <div class="wizard-nav" style="display:flex;justify-content:space-between;margin-top:24px">
          ${backBtn}${nextBtn}
        </div>
      </div>`;
  },

  // ── Step 1: Basics ──
  _wizardStep1HTML() {
    const d = this._wizardData;
    const allRoles = [
      { key: 'rep', label: 'Client Rep' },
      { key: 'l1', label: 'Team Leader' },
      { key: 'jd', label: 'Jr. Director' },
      { key: 'manager', label: 'Manager' }
    ];
    const roleChecks = allRoles.map(r => {
      const checked = d.competingRoles.includes(r.key) ? 'checked' : '';
      return `<label class="challenge-role-check"><input type="checkbox" value="${r.key}" ${checked} onchange="Challenge._collectStep1()">${r.label}</label>`;
    }).join('');

    return `
      <div class="wizard-field">
        <label class="wizard-label">Challenge Name</label>
        <input class="wizard-input" id="cw-name" value="${this._esc(d.name)}" placeholder="e.g. March Madness" oninput="Challenge._collectStep1()">
      </div>
      <div style="display:flex;gap:16px">
        <div class="wizard-field" style="flex:1">
          <label class="wizard-label">Start Date</label>
          <input class="wizard-input" type="date" id="cw-start" value="${d.startDate}" onchange="Challenge._collectStep1()">
        </div>
        <div class="wizard-field" style="flex:1">
          <label class="wizard-label">End Date</label>
          <input class="wizard-input" type="date" id="cw-end" value="${d.endDate}" onchange="Challenge._collectStep1()">
        </div>
      </div>
      <div class="wizard-field">
        <label class="wizard-label">Competing Roles</label>
        <div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:4px">${roleChecks}</div>
      </div>
      <div class="wizard-field">
        <label class="wizard-label">Mode</label>
        <div class="challenge-view-toggle" style="margin-top:4px">
          <button class="challenge-view-btn ${d.mode === 'ffa' ? 'active' : ''}" onclick="Challenge._setWizardMode('ffa')">Free-for-All</button>
          <button class="challenge-view-btn ${d.mode === 'teams' ? 'active' : ''}" onclick="Challenge._setWizardMode('teams')">Teams</button>
        </div>
        ${d.mode === 'teams' ? '<div style="font-size:12px;color:var(--silver-dim);margin-top:6px">You can assign teams after launching the challenge.</div>' : ''}
      </div>`;
  },

  _collectStep1() {
    const d = this._wizardData;
    d.name = (document.getElementById('cw-name')?.value || '').trim();
    d.startDate = document.getElementById('cw-start')?.value || '';
    d.endDate = document.getElementById('cw-end')?.value || '';
    const checks = document.querySelectorAll('.challenge-role-check input:checked');
    d.competingRoles = [...checks].map(c => c.value);
  },

  _setWizardMode(mode) {
    this._collectStep1();
    this._wizardData.mode = mode;
    this._renderWizard();
  },

  // ── Step 2: Rules ──
  _wizardStep2HTML() {
    const r = this._wizardData.rules;
    let html = '';

    // Points per Unit
    html += this._ruleCardHTML('pointsPerUnit', 'Points per Unit', 'Flat points for every unit sold', r.pointsPerUnit, 'simple');

    // Daily Goals
    html += this._ruleCardHTML('dailyGoals', 'Daily Goals', 'Hit a unit threshold each day — highest tier only', r.dailyGoals, 'tiers');

    // Event Goals
    html += this._ruleCardHTML('eventGoals', 'Event Goals', 'Hit a total unit threshold over the challenge — highest tier only', r.eventGoals, 'tiers');

    // First Blood
    html += this._ruleCardHTML('firstBlood', 'First Blood', 'First rep to log a sale each day (calculated next-day)', r.firstBlood, 'simple');

    // Last Blood
    html += this._ruleCardHTML('lastBlood', 'Last Blood', 'Last rep to log a sale each day (calculated next-day)', r.lastBlood, 'simple');

    // Penalties
    html += '<div style="font-family:\'Helvetica Neue\',\'Inter\',sans-serif;font-size:11px;letter-spacing:0.5px;text-transform:uppercase;color:var(--silver-dim);margin:24px 0 12px">Penalties (applied to unit points only)</div>';

    html += this._ruleCardHTML('activePenalty', 'Active % Penalty', 'Straight deduction — 75% active = 25% penalty on unit pts', r.activePenalty, 'toggle');
    html += this._ruleCardHTML('churn030Penalty', '0-30 Day Churn Penalty', 'Squared penalty — 5% churn = 25% penalty on unit pts', r.churn030Penalty, 'toggle');

    return html;
  },

  _ruleCardHTML(key, title, desc, rule, type) {
    const checked = rule.enabled ? 'checked' : '';
    const disabledClass = rule.enabled ? '' : ' challenge-rule-disabled';

    let body = '';
    if (type === 'simple' && rule.enabled) {
      body = `<div style="margin-top:12px">
        <label class="wizard-label">Points</label>
        <input class="wizard-input" type="number" min="0" value="${rule.points || 0}" onchange="Challenge._updateRulePoints('${key}', this.value)" style="width:120px">
      </div>`;
    } else if (type === 'tiers' && rule.enabled) {
      const tiers = rule.tiers || [];
      let tierRows = tiers.map((t, i) => `
        <div style="display:flex;gap:8px;align-items:center;margin-top:6px">
          <input class="wizard-input" type="number" min="1" value="${t.threshold}" placeholder="Units" onchange="Challenge._updateTier('${key}',${i},'threshold',this.value)" style="width:90px">
          <span style="color:var(--silver-dim);font-size:12px">units =</span>
          <input class="wizard-input" type="number" min="0" value="${t.points}" placeholder="Points" onchange="Challenge._updateTier('${key}',${i},'points',this.value)" style="width:100px">
          <span style="color:var(--silver-dim);font-size:12px">pts</span>
          <button class="challenge-btn-icon" onclick="Challenge._removeTier('${key}',${i})">✕</button>
        </div>`).join('');

      body = `<div style="margin-top:12px">
        <label class="wizard-label">Tiers (highest reached wins)</label>
        ${tierRows}
        <button class="challenge-btn-secondary" style="margin-top:8px;font-size:11px;padding:4px 12px" onclick="Challenge._addTier('${key}')">+ Add Tier</button>
      </div>`;
    }

    return `
      <div class="challenge-rule-card${disabledClass}">
        <div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px">
          <div>
            <div style="font-family:'Neue Montreal','Inter',sans-serif;font-size:15px;font-weight:700;color:var(--white)">${title}</div>
            <div style="font-size:12px;color:var(--silver-dim);margin-top:2px">${desc}</div>
          </div>
          <label class="challenge-toggle">
            <input type="checkbox" ${checked} onchange="Challenge._toggleRule('${key}', this.checked)">
            <span class="challenge-toggle-slider"></span>
          </label>
        </div>
        ${body}
      </div>`;
  },

  _toggleRule(key, enabled) {
    this._wizardData.rules[key].enabled = enabled;
    this._renderWizard();
  },

  _updateRulePoints(key, value) {
    this._wizardData.rules[key].points = parseInt(value) || 0;
  },

  _updateTier(key, index, field, value) {
    this._wizardData.rules[key].tiers[index][field] = parseInt(value) || 0;
  },

  _addTier(key) {
    this._wizardData.rules[key].tiers.push({ threshold: 0, points: 0 });
    this._renderWizard();
  },

  _removeTier(key, index) {
    this._wizardData.rules[key].tiers.splice(index, 1);
    this._renderWizard();
  },

  // ── Step 3: Review ──
  _wizardStep3HTML() {
    const d = this._wizardData;
    const r = d.rules;

    const roleLabels = { rep: 'Client Rep', l1: 'Team Leader', jd: 'Jr. Director', manager: 'Manager' };
    const roles = d.competingRoles.map(k => roleLabels[k] || k).join(', ');

    let rulesHtml = '';
    if (r.pointsPerUnit.enabled) rulesHtml += `<li><b>Points per Unit:</b> ${r.pointsPerUnit.points} pts</li>`;
    if (r.dailyGoals.enabled) {
      const tiers = r.dailyGoals.tiers.map(t => `${t.threshold}+ units = ${t.points} pts`).join(', ');
      rulesHtml += `<li><b>Daily Goals:</b> ${tiers}</li>`;
    }
    if (r.eventGoals.enabled) {
      const tiers = r.eventGoals.tiers.map(t => `${t.threshold}+ units = ${t.points} pts`).join(', ');
      rulesHtml += `<li><b>Event Goals:</b> ${tiers}</li>`;
    }
    if (r.firstBlood.enabled) rulesHtml += `<li><b>First Blood:</b> ${r.firstBlood.points} pts/day</li>`;
    if (r.lastBlood.enabled) rulesHtml += `<li><b>Last Blood:</b> ${r.lastBlood.points} pts/day</li>`;
    if (r.activePenalty.enabled) rulesHtml += `<li><b>Active % Penalty:</b> Squared penalty on unit pts</li>`;
    if (r.churn030Penalty.enabled) rulesHtml += `<li><b>0-30 Day Churn Penalty:</b> Squared penalty on unit pts</li>`;

    if (!rulesHtml) rulesHtml = '<li style="color:var(--silver-dim)">No rules enabled</li>';

    return `
      <div class="challenge-review">
        <div class="challenge-review-row"><span class="challenge-review-label">Name</span><span>${this._esc(d.name)}</span></div>
        <div class="challenge-review-row"><span class="challenge-review-label">Dates</span><span>${d.startDate} — ${d.endDate}</span></div>
        <div class="challenge-review-row"><span class="challenge-review-label">Mode</span><span>${d.mode === 'teams' ? 'Teams (assign after launch)' : 'Free-for-All'}</span></div>
        <div class="challenge-review-row"><span class="challenge-review-label">Competing</span><span>${roles}</span></div>
        <div style="margin-top:16px">
          <div class="challenge-review-label" style="margin-bottom:8px">Rules</div>
          <ul style="margin:0;padding-left:20px;font-size:14px;color:var(--white);line-height:1.8">${rulesHtml}</ul>
        </div>
      </div>`;
  },

  // ── Wizard Navigation ──
  _wizardNext() {
    if (this._wizardStep === 1) {
      this._collectStep1();
      const d = this._wizardData;
      if (!d.name) return alert('Enter a challenge name.');
      if (!d.endDate) return alert('Set an end date.');
      if (d.endDate <= d.startDate) return alert('End date must be after start date.');
      if (!d.competingRoles.length) return alert('Select at least one competing role.');
    }
    this._wizardStep++;
    this._renderWizard();
  },

  _wizardBack() {
    if (this._wizardStep === 1) return;
    if (this._wizardStep === 2) this._collectStep1(); // preserve step 1 data
    this._wizardStep--;
    this._renderWizard();
  },

  _cancelWizard() {
    this._wizardActive = false;
    this._wizardStep = 1;
    this._wizardData = {};
    this._render();
  },

  // ═════════════════════════════════════════
  // TEAM ASSIGNMENT
  // ═════════════════════════════════════════

  _openTeamAssignment() {
    this._teamAssignActive = true;
    // Deep copy existing teams or start fresh
    this._teamAssignData = (this._config.teams || []).map(t => ({
      name: t.name,
      emoji: t.emoji || '',
      members: [...(t.members || [])]
    }));
    if (this._teamAssignData.length === 0) {
      this._teamAssignData.push({ name: 'Team 1', emoji: '🔴', members: [] });
      this._teamAssignData.push({ name: 'Team 2', emoji: '🔵', members: [] });
    }
    this._renderTeamAssignment();
  },

  _renderTeamAssignment() {
    const subtitle = document.getElementById('challenge-subtitle');
    if (subtitle) subtitle.textContent = 'Assign Teams';

    const content = document.getElementById('challenge-content');
    if (!content) return;

    // Get all competing reps
    const competingRoles = new Set(this._config.competingRoles || ['rep', 'l1', 'jd']);
    const allReps = [];
    (App.state.people || []).forEach(p => {
      if (!competingRoles.has(p._roleKey)) return;
      const email = (Roster.getEmail(p.name) || '').toLowerCase();
      if (email) allReps.push({ name: p.name, email });
    });

    // Find assigned emails
    const assigned = new Set();
    this._teamAssignData.forEach(t => (t.members || []).forEach(e => assigned.add(e)));
    const unassigned = allReps.filter(r => !assigned.has(r.email));

    // Unassigned pool
    let poolHtml = unassigned.map(r =>
      `<div class="challenge-assign-chip">${this._esc(r.name)}<button class="challenge-btn-icon" onclick="Challenge._assignRep('${r.email}')">+</button></div>`
    ).join('');
    if (!poolHtml) poolHtml = '<div style="color:var(--silver-dim);font-size:13px;padding:8px">All reps assigned</div>';

    // Team columns
    const teamEmojis = ['🔴','🔵','🟢','🟡','🟣','🟠','⚫','⚪'];
    let teamsHtml = this._teamAssignData.map((team, ti) => {
      const memberChips = (team.members || []).map(email => {
        const rep = allReps.find(r => r.email === email);
        const name = rep ? rep.name : email;
        return `<div class="challenge-assign-chip">${this._esc(name)}<button class="challenge-btn-icon" onclick="Challenge._unassignRep(${ti},'${email}')">✕</button></div>`;
      }).join('');

      const emojiOptions = teamEmojis.map(e =>
        `<span class="challenge-emoji-opt ${e === team.emoji ? 'active' : ''}" onclick="Challenge._setTeamEmoji(${ti},'${e}')">${e}</span>`
      ).join('');

      return `<div class="challenge-assign-team">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
          <span style="font-size:20px">${team.emoji}</span>
          <input class="wizard-input" value="${this._esc(team.name)}" onchange="Challenge._setTeamName(${ti},this.value)" style="flex:1;font-size:14px;font-weight:700">
          <button class="challenge-btn-icon" onclick="Challenge._removeTeam(${ti})" title="Remove team">✕</button>
        </div>
        <div style="display:flex;gap:4px;margin-bottom:8px">${emojiOptions}</div>
        <div class="challenge-assign-members">${memberChips || '<div style="color:var(--silver-dim);font-size:12px;padding:4px">No members yet</div>'}</div>
      </div>`;
    }).join('');

    content.innerHTML = `
      <div class="challenge-assign-layout">
        <div class="challenge-assign-pool">
          <div class="challenge-assign-pool-header">Unassigned (${unassigned.length})</div>
          ${poolHtml}
        </div>
        <div class="challenge-assign-teams">
          ${teamsHtml}
          <button class="challenge-btn-secondary" style="width:100%;margin-top:8px" onclick="Challenge._addTeamSlot()">+ Add Team</button>
        </div>
      </div>
      <div style="display:flex;justify-content:space-between;margin-top:24px">
        <button class="challenge-btn-secondary" onclick="Challenge._cancelTeamAssignment()">Cancel</button>
        <button class="challenge-btn-primary" onclick="Challenge._saveTeams()">Save Teams</button>
      </div>`;
  },

  // Which team to assign to when clicking + on an unassigned rep
  _assignRep(email) {
    // Find the team with fewest members
    let minIdx = 0;
    let minCount = Infinity;
    this._teamAssignData.forEach((t, i) => {
      if (t.members.length < minCount) { minCount = t.members.length; minIdx = i; }
    });
    this._teamAssignData[minIdx].members.push(email);
    this._renderTeamAssignment();
  },

  _unassignRep(teamIdx, email) {
    const team = this._teamAssignData[teamIdx];
    if (team) team.members = team.members.filter(e => e !== email);
    this._renderTeamAssignment();
  },

  _setTeamName(idx, name) {
    if (this._teamAssignData[idx]) this._teamAssignData[idx].name = name;
  },

  _setTeamEmoji(idx, emoji) {
    if (this._teamAssignData[idx]) this._teamAssignData[idx].emoji = emoji;
    this._renderTeamAssignment();
  },

  _addTeamSlot() {
    const n = this._teamAssignData.length + 1;
    const emojis = ['🔴','🔵','🟢','🟡','🟣','🟠','⚫','⚪'];
    this._teamAssignData.push({ name: `Team ${n}`, emoji: emojis[(n - 1) % emojis.length], members: [] });
    this._renderTeamAssignment();
  },

  _removeTeam(idx) {
    this._teamAssignData.splice(idx, 1);
    this._renderTeamAssignment();
  },

  _cancelTeamAssignment() {
    this._teamAssignActive = false;
    this._teamAssignData = [];
    this._render();
  },

  async _saveTeams() {
    // Filter out empty teams
    const teams = this._teamAssignData.filter(t => t.members.length > 0);
    this._config.teams = teams;
    try {
      await SheetsAPI.post(OFFICE_CONFIG, 'saveChallengeConfig', { config: this._config });
      App.state.challengeConfig = this._config;
      this._teamAssignActive = false;
      this._teamAssignData = [];
      this._render();
    } catch (err) {
      alert('Failed to save teams: ' + err.message);
    }
  },

  // ═════════════════════════════════════════
  // WIZARD SUBMIT
  // ═════════════════════════════════════════

  async _wizardSubmit() {
    const d = this._wizardData;
    const config = {
      name: d.name,
      startDate: d.startDate,
      endDate: d.endDate,
      status: 'active',
      createdBy: App.state.currentEmail || '',
      mode: d.mode || 'ffa',
      competingRoles: d.competingRoles,
      teams: [],
      rules: d.rules
    };

    try {
      await SheetsAPI.post(OFFICE_CONFIG, 'saveChallengeConfig', { config });
      this._config = config;
      App.state.challengeConfig = config;
      this._sales = {};
      this._blood = {};
      this._wizardActive = false;
      this._wizardStep = 1;
      this._wizardData = {};

      // Update nav so reps can now see the tab
      App.updateNav();

      // Reload fresh data
      this.open();
    } catch (err) {
      alert('Failed to save challenge: ' + err.message);
    }
  }
};
