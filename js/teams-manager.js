// ═══════════════════════════════════════════════════════
// ELEVATE — Teams Manager
// Manages the _Teams sheet hierarchy, renders
// the Teams management page, and handles CRUD.
// ═══════════════════════════════════════════════════════

const TeamsManager = {

  // Raw team definitions from API (teamId → { teamId, name, parentId, leaderId, emoji, createdDate })
  teamsData: {},

  init(teamsData) {
    this.teamsData = teamsData || {};
  },

  // ── Build a tree of teams for rendering ──
  buildTree() {
    const byId = {};
    Object.values(this.teamsData).forEach(t => {
      byId[t.teamId] = { ...t, children: [] };
    });

    const roots = [];
    Object.values(byId).forEach(t => {
      if (t.parentId && byId[t.parentId]) {
        byId[t.parentId].children.push(t);
      } else {
        roots.push(t);
      }
    });

    // Sort roots and children alphabetically
    const sortByName = (a, b) => a.name.localeCompare(b.name);
    roots.sort(sortByName);
    Object.values(byId).forEach(t => t.children.sort(sortByName));

    return roots;
  },

  // ── Generate a teamId slug from a name ──
  generateTeamId(name) {
    return name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
  },

  // ── Get all team names (for dropdowns) ──
  getAllTeamNames() {
    return Object.values(this.teamsData).map(t => t.name).sort();
  },

  // ── Get available parents (excludes a team and its descendants to prevent cycles) ──
  getAvailableParents(excludeTeamId) {
    if (!excludeTeamId) return Object.values(this.teamsData);

    // Build descendant set
    const descendants = new Set();
    const walk = (id) => {
      Object.values(this.teamsData).forEach(t => {
        if (t.parentId === id && !descendants.has(t.teamId)) {
          descendants.add(t.teamId);
          walk(t.teamId);
        }
      });
    };
    descendants.add(excludeTeamId);
    walk(excludeTeamId);

    return Object.values(this.teamsData).filter(t => !descendants.has(t.teamId));
  },

  // ── Render the Teams management page ──
  renderTeamsPage(people, currentRole, config) {
    const page = document.getElementById('teams-page');
    if (!page) return;
    page.style.display = 'block';

    const teamCount = Object.keys(this.teamsData).length;
    const canEdit = ['superadmin', 'owner', 'manager', 'admin'].includes(currentRole);

    const subtitle = document.getElementById('teams-subtitle');
    if (subtitle) subtitle.textContent = `${teamCount} teams`;

    // Show/hide create button based on role
    const createBtn = document.getElementById('create-team-btn');
    if (createBtn) createBtn.style.display = canEdit ? '' : 'none';

    // Build tree and render rows
    const tree = this.buildTree();
    const tbody = document.getElementById('teams-tree-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    const renderNode = (node, level) => {
      const teamObj = App.state.teams.find(t => t.teamId === node.teamId);
      const directCount = teamObj ? teamObj.directMembers.length : 0;
      const totalCount = teamObj ? teamObj.members.length : 0;
      const units = teamObj ? teamObj.units : 0;

      // Find leader name from roster
      let leaderName = '';
      let leaderRole = '';
      if (node.leaderId) {
        const rosterEntry = App.state.roster[node.leaderId];
        if (rosterEntry) {
          leaderName = rosterEntry.name || node.leaderId;
          leaderRole = OFFICE_CONFIG.roles[rosterEntry.rank]?.label || rosterEntry.rank || '';
        } else {
          leaderName = node.leaderId;
        }
      }

      const indent = level * 24;
      const connector = level > 0 ? '<span style="color:var(--silver-dim);margin-right:6px">└</span>' : '';
      const memberText = node.children.length > 0 && totalCount !== directCount
        ? `${directCount} direct / ${totalCount} total`
        : `${totalCount}`;

      const row = document.createElement('tr');
      row.className = 'teams-tree-row';
      row.innerHTML = `
        <td style="padding-left:${indent + 12}px">
          ${connector}<span style="margin-right:4px">${node.emoji || '⚡'}</span>
          <span class="name-link" onclick="App.openTeamProfile('${node.name.replace(/'/g, "\\'")}')">${node.name}</span>
        </td>
        <td>${leaderName ? `${leaderName}${leaderRole ? ` <span style="font-size:10px;color:var(--silver-dim)">(${leaderRole})</span>` : ''}` : '<span style="color:var(--silver-dim)">—</span>'}</td>
        <td style="text-align:center">${memberText}</td>
        <td style="text-align:center">${units}</td>
        <td style="text-align:right">${canEdit ? `
          <button onclick="App.openEditTeamModal('${node.teamId}')" style="background:none;border:1px solid rgba(0,0,0,0.3);border-radius:6px;padding:4px 10px;color:var(--blue-core);font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;cursor:pointer;margin-right:4px">Edit</button>
          <button onclick="App.confirmDeleteTeam('${node.teamId}','${node.name.replace(/'/g, "\\'")}')" style="background:none;border:1px solid rgba(229,53,53,0.3);border-radius:6px;padding:4px 10px;color:var(--red);font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;cursor:pointer">Del</button>
        ` : ''}</td>
      `;
      tbody.appendChild(row);

      // Render children
      node.children.forEach(child => renderNode(child, level + 1));
    };

    tree.forEach(root => renderNode(root, 0));

    if (teamCount === 0) {
      tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:var(--silver-dim);padding:40px">No teams configured yet. Create your first team to get started.</td></tr>';
    }
  },

  // ── Populate create/edit modal fields ──
  populateModal(teamId) {
    const isEdit = !!teamId;
    const team = isEdit ? this.teamsData[teamId] : null;

    const title = document.getElementById('team-modal-title');
    if (title) title.textContent = isEdit ? 'Edit Team' : 'Create Team';

    const submitBtn = document.getElementById('team-modal-submit');
    if (submitBtn) submitBtn.textContent = isEdit ? 'SAVE CHANGES' : 'CREATE TEAM';

    // Store editing state
    const idField = document.getElementById('team-modal-id');
    if (idField) idField.value = teamId || '';

    const nameField = document.getElementById('team-modal-name');
    if (nameField) nameField.value = team ? team.name : '';

    // Parent dropdown
    const parentSel = document.getElementById('team-modal-parent');
    if (parentSel) {
      const available = this.getAvailableParents(teamId);
      parentSel.innerHTML = '<option value="">None (Top Level)</option>' +
        available.map(t => `<option value="${t.teamId}"${team && team.parentId === t.teamId ? ' selected' : ''}>${t.emoji || '⚡'} ${t.name}</option>`).join('');
    }

    // Leader dropdown — JD/L1/Manager roles from roster
    const leaderSel = document.getElementById('team-modal-leader');
    if (leaderSel) {
      const leaderRoles = new Set(['jd', 'l1', 'manager', 'owner']);
      const leaders = Object.entries(App.state.roster)
        .filter(([, info]) => leaderRoles.has(info.rank || info.role) && !info.deactivated)
        .sort(([, a], [, b]) => (a.name || '').localeCompare(b.name || ''));

      leaderSel.innerHTML = '<option value="">Unassigned</option>' +
        leaders.map(([email, info]) => {
          const roleLabel = OFFICE_CONFIG.roles[info.rank || info.role]?.label || info.rank || '';
          return `<option value="${email}"${team && team.leaderId === email ? ' selected' : ''}>${info.name} (${roleLabel})</option>`;
        }).join('');
    }

    // Emoji picker
    const emojiDisplay = document.getElementById('team-modal-emoji-display');
    if (emojiDisplay) emojiDisplay.textContent = team?.emoji || '⚡';
    this._selectedTeamEmoji = team?.emoji || '⚡';

    App._renderEmojiPicker('team-modal-emoji-picker', 'team-modal-emoji-display', 'TeamsManager.selectTeamEmoji', () => this.highlightTeamEmoji());

    // Error
    const errorEl = document.getElementById('team-modal-error');
    if (errorEl) errorEl.textContent = '';
  },

  _selectedTeamEmoji: '⚡',

  selectTeamEmoji(e) {
    this._selectedTeamEmoji = e;
    const display = document.getElementById('team-modal-emoji-display');
    if (display) display.textContent = e;
    this.highlightTeamEmoji();
  },

  highlightTeamEmoji() {
    document.querySelectorAll('#team-modal-emoji-picker span[id^="team-emoji-opt-"]').forEach(s => {
      s.style.borderColor = 'transparent'; s.style.background = 'transparent';
    });
    const el = document.getElementById('team-emoji-opt-' + this._selectedTeamEmoji.codePointAt(0));
    if (el) { el.style.borderColor = 'var(--sc-cyan)'; el.style.background = 'rgba(0,0,0,0.2)'; }
  },

  // ── Get form values from modal ──
  getModalValues() {
    return {
      teamId: document.getElementById('team-modal-id')?.value?.trim() || '',
      name: document.getElementById('team-modal-name')?.value?.trim() || '',
      parentId: document.getElementById('team-modal-parent')?.value || '',
      leaderId: document.getElementById('team-modal-leader')?.value || '',
      emoji: this._selectedTeamEmoji || '⚡'
    };
  }
};
