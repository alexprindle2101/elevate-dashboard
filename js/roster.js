// ═══════════════════════════════════════════════════════
// ELEVATE — Roster Management + Write-Back
// Email-based roster with Apps Script persistence
// ═══════════════════════════════════════════════════════

const Roster = {
  // State
  deactivated: new Set(),    // Set of person names
  emailMap: {},              // name → email (for write-back lookups)
  phoneMap: {},              // name → phone
  unlockRequests: {},        // { name: 'pending'|'approved' }
  teamCustomizations: {},    // { persona: { emoji, name } }
  _rosterStatusTab: 'active', // 'active' or 'deactivated' — sub-tab state

  // ── Initialize from roster data ──
  // rosterMap: email-keyed from API
  init(rosterMap) {
    this.emailMap = {};
    this.phoneMap = {};
    this.deactivated.clear();

    Object.entries(rosterMap).forEach(([key, r]) => {
      const name = r.name || key;
      const email = r.email || key;
      if (email.includes('@')) this.emailMap[name] = email;
      if (r.phone) this.phoneMap[name] = r.phone;
      if (r.deactivated || r.active === false) this.deactivated.add(name);
    });
  },

  // ── Load persisted state from Apps Script response ──
  initFromApi(apiData) {
    if (apiData.unlockRequests) {
      this.unlockRequests = { ...apiData.unlockRequests };
    }
    if (apiData.teamCustomizations) {
      this.teamCustomizations = { ...apiData.teamCustomizations };
    }
  },

  // ── Email lookup ──
  getEmail(name) {
    return this.emailMap[name] || '';
  },

  getPhone(name) {
    return this.phoneMap[name] || '';
  },

  // ── Get effective role ──
  getEffectiveRole(name, rosterMap) {
    // Check email-keyed roster (production)
    const email = this.emailMap[name];
    if (email && rosterMap[email]) return rosterMap[email].rank || rosterMap[email].role || 'rep';
    // Fall back to name-keyed roster
    if (rosterMap[name]) return rosterMap[name].role || 'rep';
    return 'rep';
  },

  // ── Get effective team ──
  getEffectiveTeam(name, people) {
    const p = people.find(x => x.name === name);
    return p ? p.team : null;
  },

  // ── Set role (local + write-back) ──
  async setRole(name, role, config) {
    const email = this.getEmail(name);
    if (email) {
      await SheetsAPI.post(config, 'updateRosterEntry', { email, rank: role });
    }
  },

  // ── Set team (local + write-back) ──
  async setTeam(name, team, people, config) {
    const p = people.find(x => x.name === name);
    if (p) p.team = team;

    const email = this.getEmail(name);
    if (email) {
      await SheetsAPI.post(config, 'updateRosterEntry', { email, team });
    }
  },

  // ── Set phone (local + write-back) ──
  async setPhone(name, phone, config) {
    this.phoneMap[name] = phone;
    const email = this.getEmail(name);
    if (email) {
      await SheetsAPI.post(config, 'updateRosterEntry', { email, phone });
    }
  },

  // ── Toggle deactivate (local + write-back) ──
  async toggleDeactivate(name, config) {
    const wasDeactivated = this.deactivated.has(name);
    if (wasDeactivated) {
      this.deactivated.delete(name);
    } else {
      this.deactivated.add(name);
    }

    const email = this.getEmail(name);
    if (email) {
      await SheetsAPI.post(config, 'toggleDeactivate', { email, deactivated: !wasDeactivated });
    }
  },

  // ── Add new person (JD+ creates roster entry) ──
  async addNewPerson(email, name, team, rank, phone, config) {
    await SheetsAPI.post(config, 'addRosterEntry', { email, name, team, rank, phone: phone || '' });
    this.emailMap[name] = email;
    if (phone) this.phoneMap[name] = phone;
    return { email, name, role: rank, team, phone: phone || '', active: true };
  },

  // ── Unlock request system (JD team edit) ──
  async sendUnlockRequest(personaName, config) {
    this.unlockRequests[personaName] = 'pending';
    await SheetsAPI.post(config, 'setUnlockRequest', { persona: personaName, status: 'pending' });
  },

  async approveUnlock(personaName, config) {
    this.unlockRequests[personaName] = 'approved';
    await SheetsAPI.post(config, 'setUnlockRequest', { persona: personaName, status: 'approved' });
  },

  async denyUnlock(personaName, config) {
    delete this.unlockRequests[personaName];
    await SheetsAPI.post(config, 'deleteUnlockRequest', { persona: personaName });
  },

  getUnlockStatus(personaName) {
    return this.unlockRequests[personaName] || null;
  },

  getPendingRequests() {
    return Object.entries(this.unlockRequests)
      .filter(([, status]) => status === 'pending')
      .map(([name]) => name);
  },

  // ── Team customization ──
  async setTeamCustomization(persona, emoji, name, config) {
    this.teamCustomizations[persona] = { emoji, name };
    delete this.unlockRequests[persona];
    await SheetsAPI.post(config, 'setTeamCustomization', { persona, emoji, displayName: name });
  },

  getTeamDisplay(teamName, people, teams) {
    for (const [persona, custom] of Object.entries(this.teamCustomizations)) {
      const p = people.find(x => x.name === persona);
      if (p && p.team === teamName) return { emoji: custom.emoji, name: custom.name };
    }
    // Fall back to actual team emoji from _Teams hierarchy
    if (teams) {
      const team = teams.find(t => t.name === teamName);
      if (team && team.emoji) return { emoji: team.emoji, name: teamName };
    }
    return { emoji: '⚡', name: teamName };
  },

  // ── Visibility (RBAC) — hierarchy-aware ──
  canViewMetrics(targetName, currentRole, currentPersona, people, teams) {
    if (currentRole === 'superadmin' || currentRole === 'owner' || currentRole === 'admin' || currentRole === 'qc') return true;
    if (targetName === currentPersona) return true;
    // Managers and JDs: scoped to their own team (+ descendant sub-teams)
    if (currentRole === 'manager' || currentRole === 'jd') {
      const myTeam = this.getEffectiveTeam(currentPersona, people);
      const theirTeam = this.getEffectiveTeam(targetName, people);
      if (myTeam !== null && myTeam === theirTeam) return true;
      // Check if their team is a descendant of my team
      if (myTeam && theirTeam && teams) {
        const myTeamObj = teams.find(t => t.name === myTeam);
        if (myTeamObj && myTeamObj.allDescendantIds) {
          const theirTeamObj = teams.find(t => t.name === theirTeam);
          if (theirTeamObj && myTeamObj.allDescendantIds.includes(theirTeamObj.teamId)) return true;
        }
      }
    }
    return false;
  },

  canViewTeam(teamName, currentRole, currentPersona, people, teams) {
    if (currentRole === 'superadmin' || currentRole === 'owner' || currentRole === 'admin' || currentRole === 'qc') return true;
    // Managers and JDs: scoped to their own team (+ descendant sub-teams)
    if (currentRole === 'manager' || currentRole === 'jd') {
      const myTeam = this.getEffectiveTeam(currentPersona, people);
      if (myTeam === teamName) return true;
      // Check if teamName is a descendant of my team
      if (myTeam && teams) {
        const myTeamObj = teams.find(t => t.name === myTeam);
        if (myTeamObj && myTeamObj.allDescendantIds) {
          const targetTeamObj = teams.find(t => t.name === teamName);
          if (targetTeamObj && myTeamObj.allDescendantIds.includes(targetTeamObj.teamId)) return true;
        }
      }
    }
    return false;
  },

  // ── Cached state for filtering ──
  _rosterPeople: [],
  _rosterRole: 'rep',
  _rosterConfig: null,

  // ── Render the roster page ──
  renderRoster(people, currentRole, config) {
    const page = document.getElementById('roster-page');
    if (!page) return;
    page.style.display = 'block';

    // Cache for re-filtering
    this._rosterPeople = people;
    this._rosterRole = currentRole;
    this._rosterConfig = config;

    // Reset sub-tab to Active on page open
    this._rosterStatusTab = 'active';
    const tabBtnActive = document.getElementById('roster-tab-active');
    const tabBtnDeactivated = document.getElementById('roster-tab-deactivated');
    if (tabBtnActive) tabBtnActive.classList.add('active');
    if (tabBtnDeactivated) tabBtnDeactivated.classList.remove('active');

    // Show/hide Add Member button based on role
    const addBtn = document.getElementById('add-member-btn');
    if (addBtn) {
      const canAdd = (currentRole === 'superadmin' || currentRole === 'owner' || currentRole === 'manager' || currentRole === 'admin' || currentRole === 'jd');
      addBtn.style.display = canAdd ? '' : 'none';
    }

    // Pending unlock requests
    const pending = this.getPendingRequests();
    const pendingSection = document.getElementById('pending-requests-section');
    const pendingList = document.getElementById('pending-requests-list');
    if (pendingSection && pendingList) {
      if (pending.length > 0) {
        pendingSection.style.display = 'block';
        pendingList.innerHTML = pending.map(name => `
          <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px">
            <span style="color:var(--white);font-size:13px;font-family:'Neue Montreal','Inter',sans-serif;font-weight:700">${name}</span>
            <span style="color:var(--silver-dim);font-size:11px">requested team edit access</span>
            <button onclick="App.approveUnlock('${name}')" style="background:rgba(46,139,87,0.13);border:1px solid rgba(46,139,87,0.33);border-radius:6px;color:#2E8B57;padding:3px 10px;font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;cursor:pointer;text-transform:uppercase">Approve</button>
            <button onclick="App.denyUnlock('${name}')" style="background:rgba(229,86,74,0.1);border:1px solid rgba(229,86,74,0.3);border-radius:6px;color:#E5564A;padding:3px 10px;font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;cursor:pointer;text-transform:uppercase">Deny</button>
          </div>`).join('');
      } else {
        pendingSection.style.display = 'none';
      }
    }

    // Populate filter dropdowns
    this._populateFilterDropdowns(people, config);

    // Render with current filters
    this.applyFilters();
  },

  // ── Populate team & role filter dropdowns ──
  _populateFilterDropdowns(people, config) {
    const teamSel = document.getElementById('roster-filter-team');
    if (teamSel) {
      const prev = teamSel.value;
      const teamNames = (App.state.teams && App.state.teams.length > 0)
        ? App.state.teams.filter(t => t.teamId !== '_unassigned').map(t => t.name)
        : config.teams;
      teamSel.innerHTML = '<option value="">All Teams</option>'
        + teamNames.map(t => `<option value="${t}">${t}</option>`).join('')
        + '<option value="Unassigned">Unassigned</option>';
      teamSel.value = prev || '';
    }

    const roleSel = document.getElementById('roster-filter-role');
    if (roleSel) {
      const prev = roleSel.value;
      roleSel.innerHTML = '<option value="">All Roles</option>'
        + Object.entries(OFFICE_CONFIG.roles)
          .filter(([key]) => key !== 'superadmin')
          .map(([key, val]) => `<option value="${key}">${val.label}</option>`)
          .join('');
      roleSel.value = prev || '';
    }
  },

  // ── Sub-tab toggle (Active / Deactivated) ──
  setStatusTab(tab) {
    this._rosterStatusTab = tab;
    const btnActive = document.getElementById('roster-tab-active');
    const btnDeactivated = document.getElementById('roster-tab-deactivated');
    if (btnActive) btnActive.classList.toggle('active', tab === 'active');
    if (btnDeactivated) btnDeactivated.classList.toggle('active', tab === 'deactivated');
    this.applyFilters();
  },

  // ── Apply filters & sort, then render rows ──
  applyFilters() {
    const people = this._rosterPeople;
    const config = this._rosterConfig;
    if (!people || !config) return;

    const search = (document.getElementById('roster-search')?.value || '').toLowerCase().trim();
    const teamFilter = document.getElementById('roster-filter-team')?.value || '';
    const roleFilter = document.getElementById('roster-filter-role')?.value || '';
    const statusTab = this._rosterStatusTab || 'active';
    const sortVal = document.getElementById('roster-sort')?.value || 'name-asc';

    // Filter by sub-tab + search/team/role
    let filtered = people.filter(p => {
      if (statusTab === 'active' && this.deactivated.has(p.name)) return false;
      if (statusTab === 'deactivated' && !this.deactivated.has(p.name)) return false;
      if (search && !p.name.toLowerCase().includes(search)) return false;
      if (teamFilter && (p.team || 'Unassigned') !== teamFilter) return false;
      if (roleFilter && (p._roleKey || 'rep') !== roleFilter) return false;
      return true;
    });

    // Sort
    const roleRank = (key) => OFFICE_CONFIG.roles[key]?.rank || 0;
    if (sortVal === 'name-asc') filtered.sort((a, b) => a.name.localeCompare(b.name));
    else if (sortVal === 'name-desc') filtered.sort((a, b) => b.name.localeCompare(a.name));
    else if (sortVal === 'role-desc') filtered.sort((a, b) => roleRank(b._roleKey || 'rep') - roleRank(a._roleKey || 'rep'));
    else if (sortVal === 'role-asc') filtered.sort((a, b) => roleRank(a._roleKey || 'rep') - roleRank(b._roleKey || 'rep'));
    else if (sortVal === 'team-asc') filtered.sort((a, b) => (a.team || 'Unassigned').localeCompare(b.team || 'Unassigned'));

    // Total in current tab (before search/team/role filters)
    const totalInTab = people.filter(p =>
      statusTab === 'active' ? !this.deactivated.has(p.name) : this.deactivated.has(p.name)
    ).length;

    // Update subtitle
    const subtitle = document.getElementById('roster-subtitle');
    if (subtitle) {
      subtitle.textContent = statusTab === 'active'
        ? `${totalInTab} active people`
        : `${totalInTab} deactivated people`;
    }

    // Update count
    const countEl = document.getElementById('roster-count');
    if (countEl) {
      countEl.textContent = filtered.length === totalInTab
        ? `Showing all ${totalInTab}`
        : `Showing ${filtered.length} of ${totalInTab}`;
    }

    // Render rows
    this._renderRosterRows(filtered, config);
  },

  // ── Render filtered roster rows ──
  _renderRosterRows(filtered, config) {
    const tbody = document.getElementById('roster-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (filtered.length === 0) {
      tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:var(--silver-dim);padding:32px;font-family:\'Cerebri Sans\',\'DM Sans\',\'Inter\',sans-serif;font-size:14px">No people match the current filters</td></tr>';
      return;
    }

    // Role permission: current user's rank determines what they can assign
    const myRole = this._rosterRole || App.state.currentRole || 'rep';
    const myRank = OFFICE_CONFIG.roles[myRole]?.rank || 0;

    const teamNames = (App.state.teams && App.state.teams.length > 0)
      ? App.state.teams.filter(t => t.teamId !== '_unassigned').map(t => t.name)
      : config.teams;

    filtered.forEach(p => {
      const isDeactivated = this.deactivated.has(p.name);
      const effRole = p._roleKey || 'rep';
      const effTeam = p.team || 'Unassigned';
      const email = p.email || this.getEmail(p.name) || '';
      const safeEmail = email.replace(/'/g, "\\'");
      const safeName = p.name.replace(/'/g, "\\'");
      const phone = Roster.getPhone(p.name);

      // Build per-row role options based on permission hierarchy
      const targetRank = OFFICE_CONFIG.roles[effRole]?.rank || 0;
      const canChangeRole = myRank > targetRank || myRole === 'superadmin';

      let roleCell;
      if (!canChangeRole) {
        // Current user can't change this person's role — show read-only label
        const roleLabel = OFFICE_CONFIG.roles[effRole]?.label || effRole;
        roleCell = `<span style="font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:12px;color:var(--silver-dim);padding:5px 0">${roleLabel}</span>`;
      } else {
        // Filter roles: only those with rank < current user's rank
        // Special rules: owner never assignable, admin only by owner/admin/superadmin
        const roleOptions = Object.entries(OFFICE_CONFIG.roles)
          .filter(([key, val]) => {
            if (key === 'superadmin') return false;  // never assignable
            if (key === 'owner') return false;        // sourced from admin portal
            if (key === 'admin') return myRole === 'owner' || myRole === 'admin' || myRole === 'superadmin';
            return val.rank < myRank || myRole === 'superadmin';
          })
          .map(([key, val]) => ({ key, label: val.label }));
        const roleSelect = roleOptions.map(r =>
          `<option value="${r.key}"${effRole === r.key ? ' selected' : ''}>${r.label}</option>`
        ).join('');
        roleCell = `<select onchange="App.setPersonRole('${safeName}',this.value)"
            style="background:rgba(0,0,0,0.06);border:1px solid rgba(0,0,0,0.25);border-radius:6px;color:var(--white);padding:5px 8px;font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:12px;cursor:pointer;outline:none">
            ${roleSelect}
          </select>`;
      }

      const teamSelect = teamNames.map(t =>
        `<option value="${t}"${effTeam === t ? ' selected' : ''}>${t}</option>`
      ).join('') + `<option value="Unassigned"${effTeam === 'Unassigned' ? ' selected' : ''}>Unassigned</option>`;

      const tr = document.createElement('tr');
      tr.style.cssText = `border-bottom:1px solid rgba(0,0,0,0.06);`;
      tr.innerHTML = `
        <td style="padding:12px 16px">
          <div id="roster-display-${safeEmail}" style="display:flex;align-items:center;gap:8px">
            <div style="flex:1">
              <div style="font-family:'Neue Montreal','Inter',sans-serif;font-size:15px;font-weight:700;color:var(--white)">
                ${p.name}
              </div>
              ${email ? `<div style="font-size:10px;color:var(--silver-dim);margin-top:2px">${email}</div>` : ''}
            </div>
            <button onclick="Roster.startEditPerson('${safeEmail}')" title="Edit name & email"
              style="background:none;border:1px solid rgba(0,0,0,0.2);border-radius:6px;padding:3px 7px;cursor:pointer;font-size:12px;color:var(--silver-dim);line-height:1">✏️</button>
          </div>
          <div id="roster-edit-${safeEmail}" style="display:none">
            <input id="roster-edit-name-${safeEmail}" type="text" value="${p.name.replace(/"/g, '&quot;')}"
              style="width:100%;box-sizing:border-box;background:rgba(255,255,255,0.6);border:1px solid rgba(0,0,0,0.3);border-radius:6px;padding:5px 8px;color:var(--white);font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:14px;font-weight:700;outline:none;margin-bottom:4px">
            <input id="roster-edit-email-${safeEmail}" type="email" value="${email}"
              style="width:100%;box-sizing:border-box;background:rgba(255,255,255,0.6);border:1px solid rgba(0,0,0,0.3);border-radius:6px;padding:4px 8px;color:var(--white);font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:11px;outline:none;margin-bottom:6px">
            <div style="display:flex;gap:6px">
              <button onclick="App.savePersonInfo('${safeEmail}')"
                style="background:var(--blue-core);border:none;border-radius:5px;padding:3px 10px;color:#fff;font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;cursor:pointer">Save</button>
              <button onclick="Roster.cancelEditPerson('${safeEmail}')"
                style="background:rgba(0,0,0,0.05);border:1px solid rgba(0,0,0,0.2);border-radius:5px;padding:3px 10px;color:var(--silver);font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;cursor:pointer">Cancel</button>
            </div>
          </div>
        </td>
        <td style="padding:12px 16px">
          ${roleCell}
        </td>
        <td style="padding:12px 16px">
          <select onchange="App.setPersonTeam('${safeName}',this.value)"
            style="background:rgba(0,0,0,0.06);border:1px solid rgba(0,0,0,0.25);border-radius:6px;color:var(--white);padding:5px 8px;font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:12px;cursor:pointer;outline:none">
            ${teamSelect}
          </select>
        </td>
        <td style="padding:12px 16px">
          <div id="roster-phone-display-${safeEmail}" style="display:flex;align-items:center;gap:6px">
            ${phone ? `<span style="position:relative;display:inline-block">
              <span onclick="Roster.showPhoneMenu(this,'${phone.replace(/'/g, "\\'")}','${phone.replace(/[^+\\d]/g, '')}')" style="font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:13px;color:var(--sc-cyan);cursor:pointer">${phone}</span>
            </span>` : `<span style="font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:13px;color:var(--silver-dim)">—</span>`}
            <button onclick="Roster.startEditPhone('${safeEmail}')" title="Edit phone"
              style="background:none;border:none;padding:2px;cursor:pointer;font-size:11px;color:var(--silver-dim);line-height:1;opacity:0.5">✏️</button>
          </div>
          <div id="roster-phone-edit-${safeEmail}" style="display:none">
            <div style="display:flex;gap:4px;align-items:center">
              <input id="roster-edit-phone-${safeEmail}" type="tel" value="${(phone || '').replace(/"/g, '&quot;')}"
                placeholder="(555) 123-4567"
                style="width:130px;box-sizing:border-box;background:rgba(255,255,255,0.6);border:1px solid rgba(0,0,0,0.3);border-radius:6px;padding:4px 8px;color:var(--white);font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:12px;outline:none">
              <button onclick="App.savePersonPhone('${safeEmail}')"
                style="background:var(--blue-core);border:none;border-radius:5px;padding:3px 8px;color:#fff;font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:10px;font-weight:700;cursor:pointer">✓</button>
              <button onclick="Roster.cancelEditPhone('${safeEmail}')"
                style="background:rgba(0,0,0,0.05);border:1px solid rgba(0,0,0,0.2);border-radius:5px;padding:3px 8px;color:var(--silver);font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:10px;font-weight:700;cursor:pointer">✕</button>
            </div>
          </div>
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

  // ── Inline edit toggle ──
  startEditPerson(email) {
    const display = document.getElementById('roster-display-' + email);
    const edit = document.getElementById('roster-edit-' + email);
    if (display) display.style.display = 'none';
    if (edit) edit.style.display = 'block';
  },

  cancelEditPerson(email) {
    const display = document.getElementById('roster-display-' + email);
    const edit = document.getElementById('roster-edit-' + email);
    if (display) display.style.display = 'flex';
    if (edit) edit.style.display = 'none';
  },

  startEditPhone(email) {
    const display = document.getElementById('roster-phone-display-' + email);
    const edit = document.getElementById('roster-phone-edit-' + email);
    if (display) display.style.display = 'none';
    if (edit) edit.style.display = 'block';
  },

  cancelEditPhone(email) {
    const display = document.getElementById('roster-phone-display-' + email);
    const edit = document.getElementById('roster-phone-edit-' + email);
    if (display) display.style.display = 'flex';
    if (edit) edit.style.display = 'none';
  },

  // ── Phone click menu (Call / Copy) ──
  showPhoneMenu(el, display, digits) {
    // Remove any existing menu
    const old = document.getElementById('phone-action-menu');
    if (old) old.remove();

    const menu = document.createElement('div');
    menu.id = 'phone-action-menu';
    menu.style.cssText = 'position:absolute;top:100%;left:0;z-index:999;background:var(--card-bg);border:1px solid rgba(0,200,255,0.25);border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,0.25);padding:4px 0;min-width:130px;margin-top:4px';

    const btnStyle = 'display:flex;align-items:center;gap:8px;width:100%;padding:8px 14px;border:none;background:none;cursor:pointer;font-family:"Cerebri Sans","DM Sans","Inter",sans-serif;font-size:13px;color:var(--white);text-align:left';
    const hoverIn = "this.style.background='rgba(0,200,255,0.1)'";
    const hoverOut = "this.style.background='none'";

    menu.innerHTML = `
      <button onmouseover="${hoverIn}" onmouseout="${hoverOut}" style="${btnStyle}" onclick="window.location.href='tel:${digits}';document.getElementById('phone-action-menu')?.remove()">
        📞 <span>Call</span>
      </button>
      <button onmouseover="${hoverIn}" onmouseout="${hoverOut}" style="${btnStyle}" onclick="Roster._copyPhone('${display.replace(/'/g, "\\'")}')">
        📋 <span>Copy</span>
      </button>`;

    el.parentElement.appendChild(menu);

    // Close on outside click
    setTimeout(() => {
      const close = (e) => { if (!menu.contains(e.target)) { menu.remove(); document.removeEventListener('click', close); } };
      document.addEventListener('click', close);
    }, 10);
  },

  _copyPhone(phone) {
    navigator.clipboard.writeText(phone).then(() => {
      const menu = document.getElementById('phone-action-menu');
      if (menu) {
        menu.innerHTML = '<div style="padding:8px 14px;font-family:\'Cerebri Sans\',\'DM Sans\',\'Inter\',sans-serif;font-size:13px;color:var(--sc-green)">✓ Copied!</div>';
        setTimeout(() => menu.remove(), 800);
      }
    });
  }
};
