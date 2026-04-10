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
    excelData: null,
    // excelData shape: {
    //   people: [{ name, email, slackEmail, departments: [], level, displayDept, displayLevel }],
    //   deptMappings: { 'Sales': ['ch1','ch2'], 'QC': ['ch3'] },
    //   roleMappings: { 'QC|Lead': ['qc-managers'], 'Sales|SWAT': ['all-swat'] },
    // }
    slackChannels: [],
    slackUsers: [],
    slackUserMap: {},
    slackChannelMemberMap: {},
    comparisonResults: [],
    isLoading: false,
    lastRefresh: null,
    searchQuery: '',
    filterMode: 'all',
  },


  // ═══════════════════════════════════════════
  // Initialization
  // ═══════════════════════════════════════════

  init() {
    console.log('[SlackApp] Initializing...');
    const saved = localStorage.getItem(SLACK_CONFIG.excelStorageKey);
    if (saved) {
      try {
        this.state.excelData = JSON.parse(saved);
        console.log('[SlackApp] Restored Excel data from localStorage');
        SlackRender.renderExcelInfo(this.state.excelData);
        this.loadSlackData();
      } catch (e) {
        console.warn('[SlackApp] Corrupt localStorage data, clearing');
        localStorage.removeItem(SLACK_CONFIG.excelStorageKey);
      }
    }
  },


  // ═══════════════════════════════════════════
  // File Upload
  // ═══════════════════════════════════════════

  handleFileSelect(event) {
    const file = event.target.files?.[0];
    if (file) this.handleFileUpload(file);
    event.target.value = '';
  },

  handleFileDrop(event) {
    const file = event.dataTransfer?.files?.[0];
    if (file) this.handleFileUpload(file);
  },

  async handleFileUpload(file) {
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls'].includes(ext)) {
      SlackRender.showError('Please upload an .xlsx or .xls file');
      return;
    }

    console.log(`[SlackApp] Parsing file: ${file.name}`);
    SlackRender.hideError();

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const excelData = this.parseExcel(workbook);

      if (!excelData.people.length) {
        SlackRender.showError('No people found in the People sheet. Check column headers: Name, Email, Department, Level');
        return;
      }

      const hasMappings = Object.keys(excelData.deptMappings).length || Object.keys(excelData.roleMappings).length;
      if (!hasMappings) {
        SlackRender.showError('No channel mappings found. Check Departments sheet (Department, Channel) and/or Roles sheet (Department, Level, Channel)');
        return;
      }

      this.state.excelData = excelData;
      localStorage.setItem(SLACK_CONFIG.excelStorageKey, JSON.stringify(excelData));

      const depts = Object.keys(excelData.deptMappings).length;
      const roles = Object.keys(excelData.roleMappings).length;
      console.log(`[SlackApp] Parsed: ${excelData.people.length} people, ${depts} departments, ${roles} role combos`);
      SlackRender.renderExcelInfo(excelData);
      this.loadSlackData();

    } catch (err) {
      console.error('[SlackApp] Excel parse error:', err);
      SlackRender.showError(`Failed to parse Excel file: ${err.message}`);
    }
  },


  // ═══════════════════════════════════════════
  // Excel Parsing
  // ═══════════════════════════════════════════

  parseExcel(workbook) {
    const cfg = SLACK_CONFIG;
    const result = { people: [], deptMappings: {}, roleMappings: {} };

    // ── People sheet ──
    const peopleSheet = workbook.Sheets[cfg.expectedSheets.people];
    if (peopleSheet) {
      const rows = XLSX.utils.sheet_to_json(peopleSheet, { defval: '' });
      result.people = rows
        .map(row => {
          const deptRaw = String(row[cfg.peopleColumns.department] || '').trim();
          const levelRaw = String(row[cfg.peopleColumns.level] || '').trim();
          return {
            name: String(row[cfg.peopleColumns.name] || '').trim(),
            email: String(row[cfg.peopleColumns.email] || '').trim().toLowerCase(),
            slackEmail: String(row[cfg.peopleColumns.slackEmail] || '').trim().toLowerCase(),
            departments: deptRaw.split(',').map(d => d.trim()).filter(Boolean),
            level: levelRaw,
            displayDept: deptRaw,
            displayLevel: levelRaw,
          };
        })
        .filter(p => p.name && p.email);
    } else {
      console.warn(`[SlackApp] Sheet "${cfg.expectedSheets.people}" not found. Available: ${workbook.SheetNames.join(', ')}`);
    }

    // ── Departments sheet (2-col: Department | Channel) ──
    const deptSheet = workbook.Sheets[cfg.expectedSheets.departments];
    if (deptSheet) {
      const rows = XLSX.utils.sheet_to_json(deptSheet, { defval: '' });
      for (const row of rows) {
        const dept = String(row[cfg.deptColumns.department] || '').trim();
        const ch = this._normalizeChannel(String(row[cfg.deptColumns.channel] || ''));
        if (!dept || !ch) continue;
        if (!result.deptMappings[dept]) result.deptMappings[dept] = [];
        result.deptMappings[dept].push(ch);
      }
      for (const k of Object.keys(result.deptMappings)) {
        result.deptMappings[k] = [...new Set(result.deptMappings[k])];
      }
    } else {
      console.warn(`[SlackApp] Sheet "${cfg.expectedSheets.departments}" not found`);
    }

    // ── Roles sheet (3-col: Department | Level | Channel) ──
    // Keyed as "Department|Level" for lookup
    const roleSheet = workbook.Sheets[cfg.expectedSheets.roles];
    if (roleSheet) {
      const rows = XLSX.utils.sheet_to_json(roleSheet, { defval: '' });
      for (const row of rows) {
        const dept = String(row[cfg.roleColumns.department] || '').trim();
        const level = String(row[cfg.roleColumns.level] || '').trim();
        const ch = this._normalizeChannel(String(row[cfg.roleColumns.channel] || ''));
        if (!dept || !level || !ch) continue;
        const key = `${dept}|${level}`;
        if (!result.roleMappings[key]) result.roleMappings[key] = [];
        result.roleMappings[key].push(ch);
      }
      for (const k of Object.keys(result.roleMappings)) {
        result.roleMappings[k] = [...new Set(result.roleMappings[k])];
      }
    } else {
      console.warn(`[SlackApp] Sheet "${cfg.expectedSheets.roles}" not found`);
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
      const [channelsRes, usersRes] = await Promise.all([
        fetch(`${url}/channels`).then(r => {
          if (!r.ok) throw new Error(`Channels: ${r.status} ${r.statusText}`);
          return r.json();
        }),
        fetch(`${url}/users`).then(r => {
          if (!r.ok) throw new Error(`Users: ${r.status} ${r.statusText}`);
          return r.json();
        }),
      ]);

      if (channelsRes.error) throw new Error(channelsRes.error);
      if (usersRes.error) throw new Error(usersRes.error);

      this.state.slackChannels = channelsRes.channels || [];
      this.state.slackUsers = usersRes.users || [];
      this.state.lastRefresh = new Date();

      this._buildLookups();
      console.log(`[SlackApp] Loaded: ${this.state.slackChannels.length} channels, ${this.state.slackUsers.length} users`);

      this.computeComparison();
      SlackRender.hideError();
      const time = this.state.lastRefresh.toLocaleTimeString();
      SlackRender.setStatus(`Updated ${time}`, true);

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


  // ═══════════════════════════════════════════
  // User Actions
  // ═══════════════════════════════════════════

  refresh() {
    if (!this.state.excelData) return;
    this.loadSlackData();
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

  clearExcelData() {
    localStorage.removeItem(SLACK_CONFIG.excelStorageKey);
    this.state.excelData = null;
    this.state.comparisonResults = [];
    this.state.slackChannels = [];
    this.state.slackUsers = [];
    this.state.slackUserMap = {};
    this.state.slackChannelMemberMap = {};
    this.state.filterMode = 'all';
    this.state.searchQuery = '';
    SlackRender.resetToEmpty();
  },

  dismissError() {
    SlackRender.hideError();
  },
};

// ── Boot ──
document.addEventListener('DOMContentLoaded', () => SlackApp.init());
