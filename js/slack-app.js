// ═══════════════════════════════════════════════════════
// Aptel Slack Channel Auditor — App Controller
// State management, Excel parsing, Slack fetching,
// comparison logic
//
// Excel schema:
//   People sheet:      Name | Email | SlackEmail | Department | Level
//   Departments sheet: Department | Channel  (base channels for everyone in dept)
//   Roles sheet:       Department | Level | Channel  (dept-specific level channels)
//
// Expected channels = dept base channels + dept-level specific channels
// ═══════════════════════════════════════════════════════

const SlackApp = {

  // ── State ──
  state: {
    currentPage: 'audit',
    excelData: null,
    slackChannels: [],
    slackUsers: [],
    slackUserGroups: [],
    slackUserMap: {},
    slackChannelMemberMap: {},
    comparisonResults: [],
    isLoading: false,
    lastRefresh: null,
    searchQuery: '',
    filterMode: 'all',
    peopleSearchQuery: '',
    pendingPeopleUpdates: new Map(), // email → { level }
  },


  // ═══════════════════════════════════════════
  // Initialization
  // ═══════════════════════════════════════════

  init() {
    console.log('[SlackApp] Initializing...');
    this.loadAll();
  },

  // Fetch sheet + Slack data, auto-populate people from Slack, then compare
  async loadAll() {
    SlackRender.hideError();
    SlackRender.showLoading('Loading configuration...');

    try {
      await this.loadSheetData();
      await this.loadSlackData();

      // Auto-populate: add any Slack users missing from the People sheet
      const synced = await this._syncSlackUsersToSheet();
      if (synced) {
        // Re-render info bar with updated counts
        SlackRender.renderExcelInfo(this.state.excelData);
      } else {
        SlackRender.renderExcelInfo(this.state.excelData);
      }
    } catch (err) {
      console.error('[SlackApp] Init error:', err);
      SlackRender.showError(`Failed to load: ${err.message}`);
    } finally {
      SlackRender.hideLoading();
    }
  },

  // Fetch sheet data from Google Sheets via Worker proxy
  async loadSheetData() {
    const url = `${SLACK_CONFIG.workerUrl}/sheet`;
    console.log(`[SlackApp] Fetching sheet data from: ${url}`);

    const res = await fetch(url);
    if (!res.ok) throw new Error(`Sheet fetch failed (${res.status})`);

    const raw = await res.json();
    const sheetData = this.parseSheetData(raw);
    this.state.excelData = sheetData;
    const depts = Object.keys(sheetData.deptMappings).length;
    const roles = Object.keys(sheetData.roleMappings).length;
    console.log(`[SlackApp] Parsed: ${sheetData.people.length} people, ${depts} departments, ${roles} role combos`);
  },


  // ═══════════════════════════════════════════
  // Sheet Data Parsing
  // ═══════════════════════════════════════════

  parseSheetData(raw) {
    const col = SLACK_CONFIG.columns;
    const result = { people: [], deptMappings: {}, roleMappings: {} };

    // ── People ──
    if (raw.people) {
      result.people = raw.people
        .map(row => {
          const deptRaw = String(row[col.department] || '').trim();
          const levelRaw = String(row[col.level] || '').trim();
          return {
            name: String(row[col.name] || '').trim(),
            email: String(row[col.email] || '').trim().toLowerCase(),
            slackEmail: String(row[col.slackEmail] || '').trim().toLowerCase(),
            departments: deptRaw.split(',').map(d => d.trim()).filter(Boolean),
            level: levelRaw,
            displayDept: deptRaw,
            displayLevel: levelRaw,
          };
        })
        .filter(p => p.name && p.email);
    }

    // ── Departments (Department → [channels]) ──
    if (raw.departments) {
      for (const row of raw.departments) {
        const dept = String(row[col.department] || '').trim();
        const ch = this._normalizeChannel(String(row[col.channel] || ''));
        if (!dept || !ch) continue;
        if (!result.deptMappings[dept]) result.deptMappings[dept] = [];
        result.deptMappings[dept].push(ch);
      }
      for (const k of Object.keys(result.deptMappings)) {
        result.deptMappings[k] = [...new Set(result.deptMappings[k])];
      }
    }

    // ── Roles (Department|Level → [channels]) ──
    if (raw.roles) {
      for (const row of raw.roles) {
        const dept = String(row[col.department] || '').trim();
        const level = String(row[col.level] || '').trim();
        const ch = this._normalizeChannel(String(row[col.channel] || ''));
        if (!dept || !level || !ch) continue;
        const key = `${dept}|${level}`;
        if (!result.roleMappings[key]) result.roleMappings[key] = [];
        result.roleMappings[key].push(ch);
      }
      for (const k of Object.keys(result.roleMappings)) {
        result.roleMappings[k] = [...new Set(result.roleMappings[k])];
      }
    }

    return result;
  },

  _normalizeChannel(ch) {
    return ch.replace(/^#/, '').toLowerCase().trim();
  },


  // ═══════════════════════════════════════════
  // Slack Data Loading
  // ═══════════════════════════════════════════

  async loadSlackData() {
    if (this.state.isLoading) return;
    this.state.isLoading = true;

    const url = SLACK_CONFIG.workerUrl;
    SlackRender.setStatus('Fetching Slack data...', false);
    SlackRender.renderSkeletonTable();
    const refreshBtn = document.getElementById('refresh-btn');
    if (refreshBtn) refreshBtn.disabled = true;

    try {
      const [channelsRes, usersRes, ugRes] = await Promise.all([
        fetch(`${url}/channels`).then(r => {
          if (!r.ok) throw new Error(`Channels: ${r.status} ${r.statusText}`);
          return r.json();
        }),
        fetch(`${url}/users`).then(r => {
          if (!r.ok) throw new Error(`Users: ${r.status} ${r.statusText}`);
          return r.json();
        }),
        fetch(`${url}/usergroups`).then(r => {
          if (!r.ok) throw new Error(`User Groups: ${r.status} ${r.statusText}`);
          return r.json();
        }).catch(() => ({ usergroups: [] })), // graceful fallback if scope missing
      ]);

      if (channelsRes.error) throw new Error(channelsRes.error);
      if (usersRes.error) throw new Error(usersRes.error);

      this.state.slackChannels = channelsRes.channels || [];
      this.state.slackUsers = usersRes.users || [];
      this.state.slackUserGroups = ugRes.usergroups || [];
      this.state.lastRefresh = new Date();

      this._buildLookups();
      this._autoDetectDepartments();
      console.log(`[SlackApp] Loaded: ${this.state.slackChannels.length} channels, ${this.state.slackUsers.length} users, ${this.state.slackUserGroups.length} user groups`);

      this.computeComparison();
      SlackRender.hideError();
      const time = this.state.lastRefresh.toLocaleTimeString();
      SlackRender.setStatus(`Updated ${time}`, true);

      // Auto-sync: add people to any missing expected channels
      this._syncMissingChannels();

    } catch (err) {
      console.error('[SlackApp] Slack fetch error:', err);
      SlackRender.showError(`Failed to fetch Slack data: ${err.message}`);
      SlackRender.setStatus('Error', false);
      document.getElementById('table-container').innerHTML = '';
    } finally {
      this.state.isLoading = false;
      if (refreshBtn) refreshBtn.disabled = false;
    }
  },

  _buildLookups() {
    this.state.slackUserMap = {};
    for (const u of this.state.slackUsers) {
      if (u.email) this.state.slackUserMap[u.email] = u;
    }

    this.state.slackChannelMemberMap = {};
    for (const ch of this.state.slackChannels) {
      this.state.slackChannelMemberMap[ch.name.toLowerCase()] = new Set(ch.members || []);
    }
  },

  // Auto-detect departments from Slack User Group membership
  // ALWAYS overrides with user group data — user groups are the source of truth for departments
  _autoDetectDepartments() {
    const { excelData, slackUserGroups, slackUserMap } = this.state;
    if (!excelData || !slackUserGroups.length) return;

    // Build lookup: lowercase dept name -> actual dept name
    const deptNames = Object.keys(excelData.deptMappings);
    const deptLookup = {};
    for (const d of deptNames) {
      deptLookup[d.toLowerCase()] = d;
    }
    for (const key of Object.keys(excelData.roleMappings)) {
      const dept = key.split('|')[0];
      if (!deptLookup[dept.toLowerCase()]) deptLookup[dept.toLowerCase()] = dept;
    }

    // Map user group -> department (match by name or handle)
    const groupToDept = {};
    for (const ug of slackUserGroups) {
      const nameMatch = deptLookup[ug.name.toLowerCase()];
      const handleMatch = deptLookup[ug.handle.toLowerCase()];
      if (nameMatch) groupToDept[ug.id] = nameMatch;
      else if (handleMatch) groupToDept[ug.id] = handleMatch;
    }

    if (!Object.keys(groupToDept).length) {
      console.log('[SlackApp] No user group → department matches found');
      return;
    }

    console.log('[SlackApp] User group → department mappings:', groupToDept);

    // Build userId -> [departments] from user group membership
    const userDepts = {};
    for (const ug of slackUserGroups) {
      const dept = groupToDept[ug.id];
      if (!dept) continue;
      for (const userId of ug.users) {
        if (!userDepts[userId]) userDepts[userId] = [];
        if (!userDepts[userId].includes(dept)) userDepts[userId].push(dept);
      }
    }

    // Always set departments from user groups — user groups are the source of truth
    const updates = [];
    let changed = 0;
    for (const person of excelData.people) {
      const lookupEmail = person.slackEmail || person.email;
      const slackUser = slackUserMap[lookupEmail];
      if (!slackUser) continue;

      const detected = userDepts[slackUser.id];
      if (detected && detected.length) {
        const newDept = detected.sort().join(', ');
        const oldDept = person.departments.sort().join(', ');
        if (newDept !== oldDept) {
          person.departments = detected;
          person.displayDept = detected.join(', ');
          updates.push({ email: person.email, department: person.displayDept });
          changed++;
        }
      }
    }

    if (changed) {
      console.log(`[SlackApp] Updated departments for ${changed} people from user groups`);
      // Write changes back to Google Sheet
      this._writeDepartmentUpdates(updates);
    }
  },

  // Write department updates back to Google Sheet via Worker
  async _writeDepartmentUpdates(updates) {
    if (!updates.length) return;
    try {
      const res = await fetch(`${SLACK_CONFIG.workerUrl}/sheet/write`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'writePeople', updates }),
      });
      const data = await res.json();
      console.log(`[SlackApp] Sheet write result:`, data);
    } catch (err) {
      console.warn('[SlackApp] Failed to write departments to sheet:', err.message);
    }
  },

  // Auto-populate People sheet from Slack users
  // Adds any Slack user not already in the sheet, detects departments from user groups
  async _syncSlackUsersToSheet() {
    const { excelData, slackUsers, slackUserGroups } = this.state;
    if (!slackUsers.length) return false;

    // Build set of emails already in the sheet
    const existingEmails = new Set(
      (excelData?.people || []).map(p => (p.slackEmail || p.email).toLowerCase())
    );

    // Find Slack users not in the sheet
    const newUsers = slackUsers.filter(u => u.email && !existingEmails.has(u.email.toLowerCase()));

    if (!newUsers.length && excelData?.people?.length) {
      console.log('[SlackApp] All Slack users already in sheet');
      return false;
    }

    // Build userId → departments from user groups
    const deptLookup = this._buildDeptLookup();
    const userDepts = {};
    for (const ug of slackUserGroups) {
      const dept = deptLookup[ug.id];
      if (!dept) continue;
      for (const userId of ug.users) {
        if (!userDepts[userId]) userDepts[userId] = [];
        if (!userDepts[userId].includes(dept)) userDepts[userId].push(dept);
      }
    }

    // Build people entries for new users
    const newPeople = newUsers.map(u => ({
      name: u.realName || u.name,
      email: u.email,
      slackEmail: u.email,
      department: (userDepts[u.id] || []).join(', '),
      level: '',
    }));

    if (newPeople.length) {
      console.log(`[SlackApp] Adding ${newPeople.length} new Slack users to sheet`);

      // Write to Google Sheet — full sync with existing + new
      const allPeople = [
        ...(excelData?.people || []).map(p => ({
          name: p.name,
          email: p.email,
          slackEmail: p.slackEmail,
          department: p.displayDept,
          level: p.displayLevel,
        })),
        ...newPeople,
      ];

      try {
        const res = await fetch(`${SLACK_CONFIG.workerUrl}/sheet/write`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: 'syncPeople', people: allPeople }),
        });
        const data = await res.json();
        console.log(`[SlackApp] Sheet sync result:`, data);
      } catch (err) {
        console.warn('[SlackApp] Failed to sync people to sheet:', err.message);
      }

      // Update local state with the new people
      for (const p of newPeople) {
        const depts = p.department.split(',').map(d => d.trim()).filter(Boolean);
        excelData.people.push({
          name: p.name,
          email: p.email.toLowerCase(),
          slackEmail: p.slackEmail.toLowerCase(),
          departments: depts,
          level: p.level,
          displayDept: p.department,
          displayLevel: p.level,
        });
      }

      return true;
    }

    return false;
  },

  // Helper: build user group ID → department name lookup
  _buildDeptLookup() {
    const { excelData, slackUserGroups } = this.state;
    const deptLookup = {};

    // Collect all known department names
    const allDepts = {};
    if (excelData) {
      for (const d of Object.keys(excelData.deptMappings || {})) {
        allDepts[d.toLowerCase()] = d;
      }
      for (const key of Object.keys(excelData.roleMappings || {})) {
        const dept = key.split('|')[0];
        if (!allDepts[dept.toLowerCase()]) allDepts[dept.toLowerCase()] = dept;
      }
    }

    // Match user group name/handle to department names
    const groupToDept = {};
    for (const ug of slackUserGroups) {
      const nameMatch = allDepts[ug.name.toLowerCase()];
      const handleMatch = allDepts[ug.handle.toLowerCase()];
      if (nameMatch) groupToDept[ug.id] = nameMatch;
      else if (handleMatch) groupToDept[ug.id] = handleMatch;
    }

    return groupToDept;
  },


  // ═══════════════════════════════════════════
  // Comparison Engine
  // ═══════════════════════════════════════════

  computeComparison() {
    const { excelData, slackUserMap } = this.state;
    if (!excelData) return;

    const results = [];

    for (const person of excelData.people) {
      const lookupEmail = person.slackEmail || person.email;
      const slackUser = slackUserMap[lookupEmail];

      // Expected = dept base channels + dept-specific level channels for each department
      const deptChannels = person.departments.flatMap(d => excelData.deptMappings[d] || []);
      const roleChannels = person.departments.flatMap(d => {
        const key = `${d}|${person.level}`;
        return excelData.roleMappings[key] || [];
      });
      const expectedChannels = [...new Set([...deptChannels, ...roleChannels])].sort();

      const roleDisplay = [person.displayDept, person.displayLevel].filter(Boolean).join(' | ');

      if (!slackUser) {
        results.push({
          name: person.name,
          email: person.email,
          department: person.displayDept,
          level: person.displayLevel,
          role: roleDisplay,
          slackUser: null,
          expectedChannels,
          actualChannels: [],
          matched: [],
          missing: expectedChannels.slice(),
          extra: [],
          status: 'notFound',
        });
        continue;
      }

      if (!expectedChannels.length) {
        const actual = this._getUserChannels(slackUser.id);
        results.push({
          name: person.name,
          email: person.email,
          department: person.displayDept || '(none)',
          level: person.displayLevel || '(none)',
          role: roleDisplay || '(none)',
          slackUser,
          expectedChannels: [],
          actualChannels: actual,
          matched: [],
          missing: [],
          extra: actual.slice(),
          status: 'noMapping',
        });
        continue;
      }

      const actualChannels = this._getUserChannels(slackUser.id);
      const expectedSet = new Set(expectedChannels);
      const actualSet = new Set(actualChannels);

      const matched = expectedChannels.filter(ch => actualSet.has(ch));
      const missing = expectedChannels.filter(ch => !actualSet.has(ch));
      const extra = actualChannels.filter(ch => !expectedSet.has(ch));

      let status = 'match';
      if (missing.length && extra.length) status = 'extra';
      else if (missing.length) status = 'missing';
      else if (extra.length) status = 'extra';

      results.push({
        name: person.name,
        email: person.email,
        department: person.displayDept,
        level: person.displayLevel,
        role: roleDisplay,
        slackUser,
        expectedChannels,
        actualChannels,
        matched,
        missing,
        extra,
        status,
      });
    }

    const statusOrder = { missing: 0, extra: 1, noMapping: 2, notFound: 3, match: 4 };
    results.sort((a, b) => (statusOrder[a.status] ?? 5) - (statusOrder[b.status] ?? 5));

    this.state.comparisonResults = results;

    SlackRender.renderSummary(results);
    SlackRender.updateFilterCounts(results);
    SlackRender.renderTable(results, this.state.filterMode, this.state.searchQuery);
  },

  _getUserChannels(userId) {
    const channels = [];
    for (const ch of this.state.slackChannels) {
      const members = this.state.slackChannelMemberMap[ch.name.toLowerCase()];
      if (members && members.has(userId)) {
        channels.push(ch.name.toLowerCase());
      }
    }
    return channels.sort();
  },

  // Auto-sync: add people to missing expected channels
  async _syncMissingChannels() {
    const results = this.state.comparisonResults;
    if (!results.length) return;

    // Collect all add actions for people with missing channels
    const actions = [];
    for (const r of results) {
      if (!r.slackUser || !r.missing.length) continue;
      for (const ch of r.missing) {
        actions.push({ userId: r.slackUser.id, channel: ch, action: 'add' });
      }
    }

    if (!actions.length) {
      console.log('[SlackApp] No missing channels to sync');
      return;
    }

    console.log(`[SlackApp] Syncing ${actions.length} missing channel assignments...`);

    try {
      const res = await fetch(`${SLACK_CONFIG.workerUrl}/slack/sync`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ actions }),
      });
      const data = await res.json();
      console.log(`[SlackApp] Sync result: ${data.succeeded}/${data.total} succeeded`);

      if (data.succeeded > 0) {
        SlackRender.showToast(`Added ${data.succeeded} missing channel assignment${data.succeeded !== 1 ? 's' : ''}`);

        // Update local state so the table reflects changes without a full refresh
        for (const result of data.results || []) {
          if (!result.ok) continue;
          const chName = result.channel?.toLowerCase();
          const memberSet = this.state.slackChannelMemberMap[chName];
          if (memberSet) memberSet.add(result.userId);
        }

        // Re-run comparison to update the table
        this.computeComparison();
      }
    } catch (err) {
      console.warn('[SlackApp] Channel sync failed:', err.message);
    }
  },


  // ═══════════════════════════════════════════
  // Navigation
  // ═══════════════════════════════════════════

  navTo(page) {
    const pages = ['audit', 'people', 'channels'];
    for (const p of pages) {
      const el = document.getElementById(`page-${p}`);
      if (el) el.style.display = p === page ? '' : 'none';
    }
    document.querySelectorAll('.sidebar-link').forEach(link => {
      link.classList.toggle('active', link.dataset.page === page);
    });
    this.state.currentPage = page;

    // Render the page content
    if (page === 'people') {
      this._renderPeoplePage();
    } else if (page === 'channels') {
      this._renderChannelsPage();
    }
  },


  // ═══════════════════════════════════════════
  // People Management
  // ═══════════════════════════════════════════

  _renderPeoplePage() {
    const people = this.state.excelData?.people || [];
    SlackRender.renderPeopleTable(people, this.state.peopleSearchQuery, this.state.pendingPeopleUpdates);
    SlackRender.renderPeopleSaveBar(this.state.pendingPeopleUpdates.size);
  },

  setPeopleSearch(query) {
    this.state.peopleSearchQuery = query;
    this._renderPeoplePage();
  },

  setPeopleLevel(email, newLevel) {
    const person = (this.state.excelData?.people || []).find(p => p.email === email);
    if (!person) return;

    person.level = newLevel;
    person.displayLevel = newLevel;

    this.state.pendingPeopleUpdates.set(email, { level: newLevel });
    SlackRender.renderPeopleSaveBar(this.state.pendingPeopleUpdates.size);

    // Highlight the changed select
    const selects = document.querySelectorAll('.level-select');
    selects.forEach(sel => {
      const row = sel.closest('tr');
      const emailCell = row?.querySelector('.email-cell');
      if (emailCell && emailCell.textContent.trim() === email) {
        sel.classList.add('changed');
        row.classList.add('row-dirty');
      }
    });
  },

  async savePeopleChanges() {
    const updates = [];
    for (const [email, data] of this.state.pendingPeopleUpdates) {
      updates.push({ email, ...data });
    }

    if (!updates.length) return;

    try {
      const res = await fetch(`${SLACK_CONFIG.workerUrl}/sheet/write`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'writePeople', updates }),
      });
      const data = await res.json();
      console.log('[SlackApp] People save result:', data);

      this.state.pendingPeopleUpdates.clear();
      this._renderPeoplePage();
      SlackRender.showToast(`Saved ${updates.length} change${updates.length !== 1 ? 's' : ''}`);

      // Re-run comparison and sync channels with new levels
      this.computeComparison();
      this._syncMissingChannels();
    } catch (err) {
      console.error('[SlackApp] People save error:', err);
      SlackRender.showToast('Failed to save: ' + err.message, true);
    }
  },


  // ═══════════════════════════════════════════
  // Channel Mapping Management
  // ═══════════════════════════════════════════

  _renderChannelsPage() {
    const data = this.state.excelData;
    if (!data) return;
    SlackRender.renderDeptMappings(data.deptMappings);
    SlackRender.renderRoleMappings(data.roleMappings);
    SlackRender.populateChannelsFormData(data.deptMappings, data.roleMappings, this.state.slackChannels);
  },

  async addDeptMapping() {
    const deptInput = document.getElementById('add-dept-name');
    const chInput = document.getElementById('add-dept-channel');
    const dept = (deptInput.value || '').trim();
    const ch = this._normalizeChannel(chInput.value || '');

    if (!dept || !ch) {
      SlackRender.showToast('Enter both a department and channel name', true);
      return;
    }

    if (!this.state.excelData.deptMappings[dept]) {
      this.state.excelData.deptMappings[dept] = [];
    }
    if (this.state.excelData.deptMappings[dept].includes(ch)) {
      SlackRender.showToast(`#${ch} is already mapped to ${dept}`, true);
      return;
    }

    this.state.excelData.deptMappings[dept].push(ch);
    deptInput.value = '';
    chInput.value = '';
    this._renderChannelsPage();

    await this._writeDeptMappingsToSheet();
    SlackRender.showToast(`Added #${ch} to ${dept}`);
  },

  async removeDeptMapping(dept, channel) {
    const arr = this.state.excelData.deptMappings[dept];
    if (!arr) return;
    const idx = arr.indexOf(channel);
    if (idx === -1) return;

    arr.splice(idx, 1);
    if (!arr.length) delete this.state.excelData.deptMappings[dept];
    this._renderChannelsPage();

    await this._writeDeptMappingsToSheet();
    SlackRender.showToast(`Removed #${channel} from ${dept}`);
  },

  async addRoleMapping() {
    const deptInput = document.getElementById('add-role-dept');
    const levelSelect = document.getElementById('add-role-level');
    const chInput = document.getElementById('add-role-channel');
    const dept = (deptInput.value || '').trim();
    const level = levelSelect.value;
    const ch = this._normalizeChannel(chInput.value || '');

    if (!dept || !level || !ch) {
      SlackRender.showToast('Enter department, level, and channel', true);
      return;
    }

    const key = `${dept}|${level}`;
    if (!this.state.excelData.roleMappings[key]) {
      this.state.excelData.roleMappings[key] = [];
    }
    if (this.state.excelData.roleMappings[key].includes(ch)) {
      SlackRender.showToast(`#${ch} is already mapped to ${key}`, true);
      return;
    }

    this.state.excelData.roleMappings[key].push(ch);
    deptInput.value = '';
    levelSelect.value = '';
    chInput.value = '';
    this._renderChannelsPage();

    await this._writeRoleMappingsToSheet();
    SlackRender.showToast(`Added #${ch} to ${dept} | ${level}`);
  },

  async removeRoleMapping(dept, level, channel) {
    const key = `${dept}|${level}`;
    const arr = this.state.excelData.roleMappings[key];
    if (!arr) return;
    const idx = arr.indexOf(channel);
    if (idx === -1) return;

    arr.splice(idx, 1);
    if (!arr.length) delete this.state.excelData.roleMappings[key];
    this._renderChannelsPage();

    await this._writeRoleMappingsToSheet();
    SlackRender.showToast(`Removed #${channel} from ${dept} | ${level}`);
  },

  async _writeDeptMappingsToSheet() {
    try {
      await fetch(`${SLACK_CONFIG.workerUrl}/sheet/write`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'writeDepartments',
          mappings: this.state.excelData.deptMappings,
        }),
      });
    } catch (err) {
      console.error('[SlackApp] Failed to write dept mappings:', err);
      SlackRender.showToast('Failed to save to sheet', true);
    }
  },

  async _writeRoleMappingsToSheet() {
    try {
      await fetch(`${SLACK_CONFIG.workerUrl}/sheet/write`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'writeRoles',
          mappings: this.state.excelData.roleMappings,
        }),
      });
    } catch (err) {
      console.error('[SlackApp] Failed to write role mappings:', err);
      SlackRender.showToast('Failed to save to sheet', true);
    }
  },


  // ═══════════════════════════════════════════
  // User Actions
  // ═══════════════════════════════════════════

  refresh() {
    this.loadAll();
  },

  setFilter(mode) {
    this.state.filterMode = mode;
    SlackRender.setActiveFilter(mode);
    SlackRender.renderTable(this.state.comparisonResults, mode, this.state.searchQuery);
  },

  setSearch(query) {
    this.state.searchQuery = query;
    SlackRender.renderTable(this.state.comparisonResults, this.state.filterMode, query);
  },

  dismissError() {
    SlackRender.hideError();
  },
};

// ── Boot ──
document.addEventListener('DOMContentLoaded', () => SlackApp.init());
