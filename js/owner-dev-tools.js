// ═══════════════════════════════════════════════════════
// Owner Development Dashboard — Tools Module
// ═══════════════════════════════════════════════════════
// Self-contained module for the Tools tab.
// First tool: Applicant Stream Reformatter.

const OwnerDevTools = {

  // ── State ──
  _initialized: false,
  _mode: 'week',        // 'week' or 'day'
  _selectedDay: 3,      // 0=Sun … 6=Sat (default Wed)
  _officeNumber: '',
  _officeMappings: {},   // { officeNumber: ownerName } — from server
  _outputRows: [],       // accumulated rows for current session
  _loading: false,
  _editingMappings: false,

  // ── localStorage keys (session rows + mode only) ──
  _LS_ROWS_KEY: 'asrf_session_rows',
  _LS_MODE_KEY: 'asrf_mode',

  // ── Output column definitions ──
  OUTPUT_COLS: [
    'Manager', 'Recruiter', 'Opens', 'AI Booked', 'Recruiter Booked',
    '% Of Call List', '1st Rounds Calendar', '1st Rounds Showed', 'Turned to 2nd',
    'Retention', 'Conversion', '2nd Rounds Booked', '2nd Rounds Showed',
    'Retention (2nd)', 'New Starts Scheduled', 'New Starts Showed', 'New Starts Showed (Retention)'
  ],

  DAY_LABELS: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],


  // ══════════════════════════════════════════════════════
  // RENDER
  // ══════════════════════════════════════════════════════

  render() {
    const container = document.getElementById('tools-content');
    if (!container) return;

    // Restore saved state from localStorage
    this._restoreState();

    container.innerHTML = this._buildHTML();
    this._bindEvents();
    this._updateModeUI();
    this._renderOutputTable();
    this._renderOfficeTags();

    // Fetch office mappings from server on first render
    if (!this._initialized) {
      this._initialized = true;
      this._fetchOfficeIds();
    }
  },


  // ══════════════════════════════════════════════════════
  // HTML BUILDER
  // ══════════════════════════════════════════════════════

  _buildHTML() {
    return `
      <div class="tools-header">
        <h2>Applicant Stream Reformatter</h2>
        <p class="tools-subtitle">Paste Applicant Stream data one office at a time to build formatted output rows.</p>
      </div>

      <!-- INPUT CARD -->
      <div class="tools-card">
        <div class="tools-form-row">
          <div class="tools-field">
            <label class="tools-label">Office #</label>
            <input type="text" id="asrf-office-num" class="tools-input" placeholder="e.g. 101" value="${this._escHtml(this._officeNumber)}">
          </div>
          <div class="tools-field">
            <label class="tools-label">Mode</label>
            <div class="tools-radio-group" id="asrf-mode-group">
              <button class="tools-radio-btn ${this._mode === 'week' ? 'active' : ''}" data-mode="week">Full Week</button>
              <button class="tools-radio-btn ${this._mode === 'day' ? 'active' : ''}" data-mode="day">Single Day</button>
            </div>
          </div>
          <div class="tools-field" id="asrf-day-picker-wrap" style="display:${this._mode === 'day' ? '' : 'none'}">
            <label class="tools-label">Day</label>
            <div class="tools-radio-group" id="asrf-day-group">
              ${this.DAY_LABELS.map((d, i) => `<button class="tools-radio-btn tools-day-btn ${this._selectedDay === i ? 'active' : ''}" data-day="${i}">${d}</button>`).join('')}
            </div>
          </div>
        </div>
        <textarea id="asrf-paste-area" class="tools-textarea" placeholder="Paste Applicant Stream data here...\n\nCopy the full table including category rows (Starting Open Applicants, Emails Received, Interviews Booked, etc.) and paste here." rows="10"></textarea>
        <div class="tools-form-row tools-actions">
          <button class="tools-btn tools-btn-primary" id="asrf-process-btn">Process</button>
          <button class="tools-btn tools-btn-secondary" id="asrf-clear-paste-btn">Clear Paste</button>
        </div>
      </div>

      <!-- OUTPUT TABLE CARD -->
      <div class="tools-card">
        <div class="tools-card-header">
          <h3>Output</h3>
          <span id="asrf-row-count" class="tools-muted">0 rows</span>
        </div>
        <div class="tools-table-wrap">
          <table class="tools-output-table" id="asrf-output-table">
            <thead>
              <tr>${this.OUTPUT_COLS.map(c => `<th>${c}</th>`).join('')}</tr>
            </thead>
            <tbody id="asrf-output-tbody"></tbody>
          </table>
        </div>
        <div class="tools-form-row tools-actions" style="margin-top:12px">
          <button class="tools-btn tools-btn-primary" id="asrf-copy-btn">Copy All to Clipboard</button>
          <button class="tools-btn tools-btn-secondary" id="asrf-screenshot-btn">📸 Screenshot View</button>
          <button class="tools-btn tools-btn-danger" id="asrf-clear-rows-btn">Clear All Rows</button>
        </div>
      </div>

      <!-- SAVED OFFICE MAPPINGS CARD -->
      <div class="tools-card">
        <div class="tools-card-header">
          <h3>Saved Office Mappings</h3>
          <div style="display:flex;gap:6px">
            <button class="tools-btn tools-btn-icon" id="asrf-edit-mappings-btn" title="Edit mappings">&#x270E;</button>
            <button class="tools-btn tools-btn-icon" id="asrf-refresh-btn" title="Refresh from server">&#x21bb;</button>
          </div>
        </div>
        <div id="asrf-office-tags" class="tools-office-tags">
          <span class="tools-muted">Loading...</span>
        </div>
      </div>

      <!-- OWNER NAME PROMPT OVERLAY -->
      <div id="asrf-prompt-overlay" class="tools-prompt-overlay" style="display:none">
        <div class="tools-prompt-card">
          <h3>New Office Number</h3>
          <p id="asrf-prompt-msg">Office #<strong id="asrf-prompt-num"></strong> is not recognized. What is the owner's name?</p>
          <input type="text" id="asrf-prompt-name" class="tools-input" placeholder="Owner name">
          <div class="tools-form-row tools-actions" style="margin-top:12px">
            <button class="tools-btn tools-btn-primary" id="asrf-prompt-save">Save & Continue</button>
            <button class="tools-btn tools-btn-secondary" id="asrf-prompt-cancel">Cancel</button>
          </div>
        </div>
      </div>
    `;
  },


  // ══════════════════════════════════════════════════════
  // EVENT BINDING
  // ══════════════════════════════════════════════════════

  _bindEvents() {
    // Mode toggle
    document.getElementById('asrf-mode-group').addEventListener('click', e => {
      const btn = e.target.closest('[data-mode]');
      if (!btn) return;
      this._mode = btn.dataset.mode;
      this._updateModeUI();
      this._saveMode();
    });

    // Day picker
    document.getElementById('asrf-day-group').addEventListener('click', e => {
      const btn = e.target.closest('[data-day]');
      if (!btn) return;
      this._selectedDay = parseInt(btn.dataset.day);
      this._updateModeUI();
      this._saveMode();
    });

    // Process
    document.getElementById('asrf-process-btn').addEventListener('click', () => this._onProcess());

    // Clear paste
    document.getElementById('asrf-clear-paste-btn').addEventListener('click', () => {
      document.getElementById('asrf-paste-area').value = '';
    });

    // Copy all
    document.getElementById('asrf-copy-btn').addEventListener('click', () => this._onCopyAll());

    // Screenshot view
    document.getElementById('asrf-screenshot-btn').addEventListener('click', () => this._onScreenshotView());

    // Clear rows
    document.getElementById('asrf-clear-rows-btn').addEventListener('click', () => {
      if (!this._outputRows.length) return;
      this._outputRows = [];
      this._saveRows();
      this._renderOutputTable();
      this._toast('Output cleared', 'info');
    });

    // Refresh office mappings
    document.getElementById('asrf-refresh-btn').addEventListener('click', () => this._fetchOfficeIds());

    // Edit office mappings
    document.getElementById('asrf-edit-mappings-btn').addEventListener('click', () => this._toggleEditMappings());

    // Prompt overlay
    document.getElementById('asrf-prompt-cancel').addEventListener('click', () => {
      document.getElementById('asrf-prompt-overlay').style.display = 'none';
      this._promptCallback = null;
    });
    document.getElementById('asrf-prompt-save').addEventListener('click', () => this._onPromptSave());
    document.getElementById('asrf-prompt-name').addEventListener('keydown', e => {
      if (e.key === 'Enter') this._onPromptSave();
    });
  },


  // ══════════════════════════════════════════════════════
  // MODE UI
  // ══════════════════════════════════════════════════════

  _updateModeUI() {
    // Mode buttons
    document.querySelectorAll('#asrf-mode-group .tools-radio-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.mode === this._mode);
    });
    // Day picker visibility
    const dayWrap = document.getElementById('asrf-day-picker-wrap');
    if (dayWrap) dayWrap.style.display = this._mode === 'day' ? '' : 'none';
    // Day buttons
    document.querySelectorAll('#asrf-day-group .tools-day-btn').forEach(btn => {
      btn.classList.toggle('active', parseInt(btn.dataset.day) === this._selectedDay);
    });
  },


  // ══════════════════════════════════════════════════════
  // PROCESS FLOW
  // ══════════════════════════════════════════════════════

  _onProcess() {
    const officeNum = document.getElementById('asrf-office-num').value.trim();
    if (!officeNum) return this._toast('Enter an office number', 'error');

    const raw = document.getElementById('asrf-paste-area').value.trim();
    if (!raw) return this._toast('Paste Applicant Stream data first', 'error');

    this._officeNumber = officeNum;

    // Check if we have a mapping for this office
    if (!this._officeMappings[officeNum]) {
      this._showOwnerPrompt(officeNum, () => this._processData(officeNum, raw));
      return;
    }

    this._processData(officeNum, raw);
  },

  _processData(officeNum, raw) {
    try {
      const parsed = this._parseApplicantStream(raw);
      const colIndex = this._mode === 'week' ? 7 : this._selectedDay;
      const row = this._extractRow(parsed, colIndex, officeNum);

      this._outputRows.push(row);
      this._saveRows();
      this._renderOutputTable();

      // Clear textarea for next paste
      document.getElementById('asrf-paste-area').value = '';

      this._toast(`Row added for Office #${officeNum}`, 'success');
    } catch (err) {
      console.error('[ASRF] Parse error:', err);
      this._toast('Error parsing data: ' + err.message, 'error');
    }
  },


  // ══════════════════════════════════════════════════════
  // PARSER
  // ══════════════════════════════════════════════════════

  _parseApplicantStream(raw) {
    const lines = raw.split('\n').filter(l => l.trim() !== '');
    const result = { categories: {}, categoryOrder: [] };
    let currentCategory = null;

    // Known category-level labels (used as fallback detection)
    const KNOWN_CATEGORIES = [
      'starting open applicants', 'emails received', 'removed from process',
      'sent to call list', 'manual apps entry', 'no answers',
      'left message one', 'left message two', 'left message three',
      'removed from call list', 'administrative booking data',
      'interviews booked', 'retention call list',
      'total daily bob', 'second interviews booked',
      'total first interviews', 'first interviews showed up',
      'cancel first interview', 'retention first interviews',
      'total second interviews', 'second interviews showed up',
      'retention second interviews', 'first showed up booked second',
      'retention first showed up booked second',
      'removed from calendar', 'removed from second calendar', 'removed from third calendar',
      'total training', 'training showed up', 'retention training',
      'total new starts scheduled', 'new starts showed up',
      'retention new starts scheduled',
      'offered job from second round', 'offered job from third round',
      'disqualifed from first calendar', 'disqualifed from second calendar',
      'disqualifed from third calendar',
      'declined from first calendar', 'declined from second calendar',
      'declined from third calendar'
    ];

    for (let i = 0; i < lines.length; i++) {
      const rawLine = lines[i];
      const cells = rawLine.split('\t');

      // Find label: first non-empty cell (sub-rows may have leading empty tabs)
      let label = '';
      let labelCol = 0;
      for (let c = 0; c < cells.length; c++) {
        const trimmed = (cells[c] || '').trim();
        if (trimmed) { label = trimmed; labelCol = c; break; }
      }
      if (!label) continue;

      // Skip section headers like "Administrative Booking Data:"
      if (label.endsWith(':') && cells.length <= 2) continue;

      // Parse value cells — they start after the label column
      // For category rows (labelCol=0): values are cells[1..8]
      // For sub-rows with offset labels: values still map to day columns
      // The data always has 7 day columns + 1 total at fixed positions from the right
      const values = [];
      const totalCells = cells.length;
      // Take the last 8 numeric cells (7 days + total) regardless of label position
      const valueStart = Math.max(labelCol + 1, totalCells - 8);
      for (let c = valueStart; c < totalCells; c++) {
        values.push(this._parseValue(cells[c]));
      }
      // Pad to 8 if shorter
      while (values.length < 8) values.push(0);

      // Determine if this is a category or sub-row
      const labelLower = label.toLowerCase();
      const isKnownCategory = KNOWN_CATEGORIES.some(k => labelLower.startsWith(k));

      if (isKnownCategory) {
        // Definitely a category row
        currentCategory = label;
        result.categories[label] = { values, subs: {} };
        result.categoryOrder.push(label);
      } else if (currentCategory) {
        // Sub-row under current category (person names, AI Messaging, etc.)
        result.categories[currentCategory].subs[label] = { values };
      } else {
        // No current category yet — treat as category
        currentCategory = label;
        result.categories[label] = { values, subs: {} };
        result.categoryOrder.push(label);
      }
    }

    return result;
  },

  _parseValue(cell) {
    if (!cell || typeof cell !== 'string') return 0;
    const s = cell.trim();
    if (!s || s === '-' || s === '#DIV/0!' || s === '#DIV/01' || s.startsWith('#')) return 0;
    // Handle percentages: "55 %" or "55%" → 55
    const cleaned = s.replace(/[%\s,]/g, '');
    const num = parseFloat(cleaned);
    return isNaN(num) ? 0 : num;
  },

  _findCategory(parsed, label) {
    // Exact match
    if (parsed.categories[label]) return parsed.categories[label];
    // Case-insensitive match
    const lower = label.toLowerCase();
    for (const key of Object.keys(parsed.categories)) {
      if (key.toLowerCase() === lower) return parsed.categories[key];
    }
    // Substring match (category starts with label)
    for (const key of Object.keys(parsed.categories)) {
      if (key.toLowerCase().startsWith(lower)) return parsed.categories[key];
    }
    // Label starts with category
    for (const key of Object.keys(parsed.categories)) {
      if (lower.startsWith(key.toLowerCase())) return parsed.categories[key];
    }
    return null;
  },

  _getCatValue(parsed, label, colIndex) {
    const cat = this._findCategory(parsed, label);
    if (!cat) return 0;
    return cat.values[colIndex] || 0;
  },

  _getSubValue(parsed, categoryLabel, subLabel, colIndex) {
    const cat = this._findCategory(parsed, categoryLabel);
    if (!cat || !cat.subs) return 0;
    // Exact sub match
    if (cat.subs[subLabel]) return cat.subs[subLabel].values[colIndex] || 0;
    // Case-insensitive sub match
    const lower = subLabel.toLowerCase();
    for (const key of Object.keys(cat.subs)) {
      if (key.toLowerCase() === lower) return cat.subs[key].values[colIndex] || 0;
    }
    return 0;
  },


  // ══════════════════════════════════════════════════════
  // ROW EXTRACTION
  // ══════════════════════════════════════════════════════

  _extractRow(parsed, colIndex, officeNum) {
    const manager = this._officeMappings[officeNum] || 'Unknown';

    // Recruiter names: sub-rows under "Interviews Booked" excluding "AI Messaging"
    const ibCat = this._findCategory(parsed, 'Interviews Booked');
    const recruiterNames = [];
    let aiBooked = 0;
    let recruiterBooked = 0;

    if (ibCat && ibCat.subs) {
      for (const [name, data] of Object.entries(ibCat.subs)) {
        if (name.toLowerCase() === 'ai messaging') {
          aiBooked = data.values[colIndex] || 0;
        } else {
          recruiterNames.push(name);
          recruiterBooked += (data.values[colIndex] || 0);
        }
      }
    }

    // Opens = Sent to Call List + Manual Apps Entry
    const opens = this._getCatValue(parsed, 'Sent to Call List', colIndex)
                + this._getCatValue(parsed, 'Manual Apps Entry', colIndex);

    // Training fallback logic
    const trainingCat = this._findCategory(parsed, 'Total Training');
    const trainingAllZero = trainingCat
      ? trainingCat.values.every(v => v === 0)
      : true;

    let newStartsScheduled, newStartsShowed, newStartsRetention;
    if (trainingAllZero) {
      newStartsScheduled = this._getCatValue(parsed, 'Total New Starts Scheduled', colIndex);
      newStartsShowed    = this._getCatValue(parsed, 'New Starts Showed Up', colIndex);
      newStartsRetention = this._getCatValue(parsed, 'Retention New Starts Scheduled', colIndex);
    } else {
      newStartsScheduled = this._getCatValue(parsed, 'Total Training', colIndex);
      newStartsShowed    = this._getCatValue(parsed, 'Training Showed Up', colIndex);
      newStartsRetention = this._getCatValue(parsed, 'Retention Training', colIndex);
    }

    return {
      manager,
      recruiter: recruiterNames.join(' & '),
      opens,
      aiBooked,
      recruiterBooked,
      pctCallList:      this._getCatValue(parsed, 'Retention Call List', colIndex),
      firstCalendar:    this._getCatValue(parsed, 'Total First Interviews', colIndex),
      firstShowed:      this._getCatValue(parsed, 'First Interviews Showed Up', colIndex),
      turnedTo2nd:      this._getCatValue(parsed, 'First Showed Up Booked Second', colIndex),
      retention1st:     this._getCatValue(parsed, 'Retention First Interviews', colIndex),
      conversion:       this._getCatValue(parsed, 'Retention First Showed Up Booked Second', colIndex),
      secondBooked:     this._getCatValue(parsed, 'Total Second Interviews', colIndex),
      secondShowed:     this._getCatValue(parsed, 'Second Interviews Showed Up', colIndex),
      retention2nd:     this._getCatValue(parsed, 'Retention Second Interviews', colIndex),
      newStartsScheduled,
      newStartsShowed,
      newStartsRetention
    };
  },


  // ══════════════════════════════════════════════════════
  // OUTPUT TABLE
  // ══════════════════════════════════════════════════════

  _renderOutputTable() {
    const tbody = document.getElementById('asrf-output-tbody');
    if (!tbody) return;

    if (!this._outputRows.length) {
      tbody.innerHTML = '<tr><td colspan="17" class="tools-muted" style="text-align:center;padding:20px">No rows yet. Process an office to add rows.</td></tr>';
    } else {
      tbody.innerHTML = this._outputRows.map((row, idx) => `
        <tr>
          <td>${this._escHtml(row.manager)}</td>
          <td>${this._escHtml(row.recruiter)}</td>
          <td>${row.opens}</td>
          <td>${row.aiBooked}</td>
          <td>${row.recruiterBooked}</td>
          <td>${this._fmtPct(row.pctCallList)}</td>
          <td>${row.firstCalendar}</td>
          <td>${row.firstShowed}</td>
          <td>${row.turnedTo2nd}</td>
          <td>${this._fmtPct(row.retention1st)}</td>
          <td>${this._fmtPct(row.conversion)}</td>
          <td>${row.secondBooked}</td>
          <td>${row.secondShowed}</td>
          <td>${this._fmtPct(row.retention2nd)}</td>
          <td>${row.newStartsScheduled}</td>
          <td>${row.newStartsShowed}</td>
          <td>${this._fmtPct(row.newStartsRetention)}</td>
        </tr>
      `).join('');
    }

    const countEl = document.getElementById('asrf-row-count');
    if (countEl) countEl.textContent = `${this._outputRows.length} row${this._outputRows.length !== 1 ? 's' : ''}`;
  },

  _fmtPct(val) {
    if (val === 0 || val === '0' || !val) return '0%';
    return Math.round(val) + '%';
  },


  // ══════════════════════════════════════════════════════
  // OFFICE TAGS
  // ══════════════════════════════════════════════════════

  _renderOfficeTags() {
    const container = document.getElementById('asrf-office-tags');
    if (!container) return;

    const entries = Object.entries(this._officeMappings);
    if (!entries.length) {
      container.innerHTML = '<span class="tools-muted">No saved office mappings yet.</span>';
      return;
    }

    const sorted = entries.sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true }));

    if (this._editingMappings) {
      container.innerHTML = sorted.map(([num, name]) =>
        `<div class="tools-office-edit-row" data-office="${this._escHtml(num)}">
          <span class="tools-office-tag-num">#${this._escHtml(num)}</span>
          <input type="text" class="tools-input tools-edit-name-input" value="${this._escHtml(name)}" data-office="${this._escHtml(num)}">
          <button class="tools-btn-inline tools-btn-save-mapping" data-office="${this._escHtml(num)}" title="Save">&#x2713;</button>
          <button class="tools-btn-inline tools-btn-delete-mapping" data-office="${this._escHtml(num)}" title="Delete">&times;</button>
        </div>`
      ).join('');

      // Bind edit row events
      container.querySelectorAll('.tools-btn-save-mapping').forEach(btn => {
        btn.addEventListener('click', () => this._onSaveMapping(btn.dataset.office));
      });
      container.querySelectorAll('.tools-btn-delete-mapping').forEach(btn => {
        btn.addEventListener('click', () => this._onDeleteMapping(btn.dataset.office));
      });
      container.querySelectorAll('.tools-edit-name-input').forEach(input => {
        input.addEventListener('keydown', e => {
          if (e.key === 'Enter') this._onSaveMapping(input.dataset.office);
        });
      });
    } else {
      container.innerHTML = sorted.map(([num, name]) =>
        `<span class="tools-office-tag">#${this._escHtml(num)} &rarr; ${this._escHtml(name)}</span>`
      ).join('');
    }
  },

  _toggleEditMappings() {
    this._editingMappings = !this._editingMappings;
    const btn = document.getElementById('asrf-edit-mappings-btn');
    if (btn) btn.classList.toggle('active', this._editingMappings);
    this._renderOfficeTags();
  },

  async _onSaveMapping(officeNum) {
    const input = document.querySelector(`.tools-edit-name-input[data-office="${officeNum}"]`);
    if (!input) return;
    const newName = input.value.trim();
    if (!newName) return this._toast('Name cannot be empty', 'error');

    this._officeMappings[officeNum] = newName;
    try {
      const email = (typeof OwnerDev !== 'undefined' && OwnerDev.state.session)
        ? OwnerDev.state.session.email : '';
      await OwnerDev._post('odSaveOfficeId', { officeNumber: officeNum, ownerName: newName, email });
      this._toast(`Updated Office #${officeNum}`, 'success');
    } catch (err) {
      console.warn('[ASRF] Failed to save:', err);
    }
    this._renderOfficeTags();
  },

  async _onDeleteMapping(officeNum) {
    delete this._officeMappings[officeNum];
    try {
      const email = (typeof OwnerDev !== 'undefined' && OwnerDev.state.session)
        ? OwnerDev.state.session.email : '';
      await OwnerDev._post('odDeleteOfficeId', { officeNumber: officeNum, email });
    } catch (err) {
      console.warn('[ASRF] Failed to delete:', err);
    }
    this._toast(`Removed Office #${officeNum}`, 'info');
    this._renderOfficeTags();
  },


  // ══════════════════════════════════════════════════════
  // OWNER NAME PROMPT
  // ══════════════════════════════════════════════════════

  _promptCallback: null,

  _showOwnerPrompt(officeNum, callback) {
    this._promptCallback = callback;
    document.getElementById('asrf-prompt-num').textContent = officeNum;
    document.getElementById('asrf-prompt-name').value = '';
    document.getElementById('asrf-prompt-overlay').style.display = '';
    setTimeout(() => document.getElementById('asrf-prompt-name').focus(), 50);
  },

  async _onPromptSave() {
    const name = document.getElementById('asrf-prompt-name').value.trim();
    if (!name) return this._toast('Enter an owner name', 'error');

    const officeNum = document.getElementById('asrf-office-num').value.trim();

    // Save to server
    try {
      const email = (typeof OwnerDev !== 'undefined' && OwnerDev.state.session)
        ? OwnerDev.state.session.email : '';
      await OwnerDev._post('odSaveOfficeId', { officeNumber: officeNum, ownerName: name, email });
    } catch (err) {
      console.warn('[ASRF] Failed to save office ID to server:', err);
    }

    // Update local state
    this._officeMappings[officeNum] = name;
    this._renderOfficeTags();

    // Close prompt and continue
    document.getElementById('asrf-prompt-overlay').style.display = 'none';
    if (this._promptCallback) {
      const cb = this._promptCallback;
      this._promptCallback = null;
      cb();
    }
  },


  // ══════════════════════════════════════════════════════
  // CLIPBOARD
  // ══════════════════════════════════════════════════════

  _onCopyAll() {
    if (!this._outputRows.length) return this._toast('No rows to copy', 'error');

    const rows = this._outputRows.map(r => [
      r.manager, r.recruiter, r.opens, r.aiBooked, r.recruiterBooked,
      this._fmtPct(r.pctCallList), r.firstCalendar, r.firstShowed, r.turnedTo2nd,
      this._fmtPct(r.retention1st), this._fmtPct(r.conversion),
      r.secondBooked, r.secondShowed, this._fmtPct(r.retention2nd),
      r.newStartsScheduled, r.newStartsShowed, this._fmtPct(r.newStartsRetention)
    ].join('\t'));

    const text = rows.join('\n');

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => {
        this._toast(`Copied ${this._outputRows.length} rows to clipboard`, 'success');
      }).catch(() => this._copyFallback(text));
    } else {
      this._copyFallback(text);
    }
  },

  _copyFallback(text) {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try {
      document.execCommand('copy');
      this._toast(`Copied ${this._outputRows.length} rows to clipboard`, 'success');
    } catch (e) {
      this._toast('Copy failed — select and copy manually', 'error');
    }
    document.body.removeChild(ta);
  },


  // ══════════════════════════════════════════════════════
  // SCREENSHOT VIEW
  // ══════════════════════════════════════════════════════

  _onScreenshotView() {
    if (!this._outputRows.length) return this._toast('No rows to display', 'error');

    // Remove existing overlay if any
    const existing = document.getElementById('asrf-screenshot-overlay');
    if (existing) existing.remove();

    const today = new Date();
    const dateStr = (today.getMonth() + 1) + '/' + today.getDate() + '/' + today.getFullYear();

    const headerCols = [
      { label: 'Manager', align: 'left' },
      { label: 'Recruiter', align: 'left' },
      { label: 'Opens', align: 'right' },
      { label: 'AI Bkd', align: 'right' },
      { label: 'Rec Bkd', align: 'right' },
      { label: '% Call List', align: 'right', pct: true },
      { label: '1st Cal', align: 'right' },
      { label: '1st Show', align: 'right' },
      { label: '→ 2nd', align: 'right' },
      { label: 'Ret 1st', align: 'right', pct: true },
      { label: 'Conv', align: 'right', pct: true },
      { label: '2nd Bkd', align: 'right' },
      { label: '2nd Show', align: 'right' },
      { label: 'Ret 2nd', align: 'right', pct: true },
      { label: 'NS Sched', align: 'right' },
      { label: 'NS Show', align: 'right' },
      { label: 'NS Ret', align: 'right', pct: true }
    ];

    const rowValues = (r) => [
      r.manager, r.recruiter, r.opens, r.aiBooked, r.recruiterBooked,
      r.pctCallList, r.firstCalendar, r.firstShowed, r.turnedTo2nd,
      r.retention1st, r.conversion, r.secondBooked, r.secondShowed,
      r.retention2nd, r.newStartsScheduled, r.newStartsShowed, r.newStartsRetention
    ];

    const pctColor = (val) => {
      const n = parseFloat(val) || 0;
      if (n >= 70) return '#1e6f34';
      if (n >= 50) return '#2ea043';
      if (n >= 30) return '#9a6500';
      return '#8b1a1a';
    };

    const fmtCell = (val, col) => {
      if (col.pct) {
        const n = parseFloat(val) || 0;
        return `<span style="color:${pctColor(n)};font-weight:600;">${Math.round(n)}%</span>`;
      }
      return val;
    };

    const theadHtml = headerCols.map(c =>
      `<th style="text-align:${c.align};padding:10px 12px;font-size:11px;font-weight:700;color:#fff;white-space:nowrap;letter-spacing:0.3px;">${c.label}</th>`
    ).join('');

    const tbodyHtml = this._outputRows.map((r, idx) => {
      const vals = rowValues(r);
      const bg = idx % 2 === 0 ? '#fff' : 'var(--gray-50, #f8f9fb)';
      const cells = vals.map((v, ci) => {
        const col = headerCols[ci];
        const weight = ci <= 1 ? 'font-weight:600;' : '';
        const align = col.align;
        const font = ci > 1 ? 'font-variant-numeric:tabular-nums;' : '';
        return `<td style="text-align:${align};padding:8px 12px;${weight}${font}white-space:nowrap;">${fmtCell(v, col)}</td>`;
      }).join('');
      return `<tr style="background:${bg};border-bottom:1px solid rgba(0,0,0,0.05);">${cells}</tr>`;
    }).join('');

    const overlay = document.createElement('div');
    overlay.id = 'asrf-screenshot-overlay';
    overlay.className = 'tools-screenshot-overlay';
    overlay.innerHTML = `
      <div class="tools-screenshot-modal">
        <button class="tools-screenshot-close" onclick="document.getElementById('asrf-screenshot-overlay').remove()">&times;</button>
        <div class="tools-screenshot-content">
          <div class="tools-screenshot-header">
            <div style="display:flex;align-items:center;gap:12px;">
              <span style="font-size:18px;font-weight:800;color:var(--gray-800,#1a202c);">Applicant Stream Report</span>
            </div>
            <span style="font-size:12px;color:var(--gray-400,#98a3b3);font-weight:500;">${dateStr}</span>
          </div>
          <div style="overflow-x:auto;">
            <table class="tools-screenshot-table">
              <thead><tr>${theadHtml}</tr></thead>
              <tbody>${tbodyHtml}</tbody>
            </table>
          </div>
        </div>
      </div>`;

    // Close on backdrop click
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) overlay.remove();
    });

    document.body.appendChild(overlay);
  },


  // ══════════════════════════════════════════════════════
  // SERVER: FETCH / SAVE OFFICE IDS
  // ══════════════════════════════════════════════════════

  async _fetchOfficeIds() {
    const refreshBtn = document.getElementById('asrf-refresh-btn');
    if (refreshBtn) refreshBtn.classList.add('spinning');

    try {
      const resp = await OwnerDev._api('odGetOfficeIds');
      if (resp && resp.success && Array.isArray(resp.officeIds)) {
        this._officeMappings = {};
        for (const row of resp.officeIds) {
          const num = String(row.officeNumber || '').trim();
          const name = String(row.ownerName || '').trim();
          if (num && name) this._officeMappings[num] = name;
        }
        this._renderOfficeTags();
      }
    } catch (err) {
      console.warn('[ASRF] Failed to fetch office IDs:', err);
      this._toast('Could not load office mappings', 'error');
    } finally {
      if (refreshBtn) refreshBtn.classList.remove('spinning');
    }
  },


  // ══════════════════════════════════════════════════════
  // LOCALSTORAGE (session rows + mode only)
  // ══════════════════════════════════════════════════════

  _restoreState() {
    try {
      const rows = localStorage.getItem(this._LS_ROWS_KEY);
      if (rows) this._outputRows = JSON.parse(rows);
    } catch (e) { /* ignore */ }

    try {
      const mode = localStorage.getItem(this._LS_MODE_KEY);
      if (mode) {
        const m = JSON.parse(mode);
        if (m.mode) this._mode = m.mode;
        if (typeof m.day === 'number') this._selectedDay = m.day;
      }
    } catch (e) { /* ignore */ }
  },

  _saveRows() {
    try {
      localStorage.setItem(this._LS_ROWS_KEY, JSON.stringify(this._outputRows));
    } catch (e) { /* ignore */ }
  },

  _saveMode() {
    try {
      localStorage.setItem(this._LS_MODE_KEY, JSON.stringify({ mode: this._mode, day: this._selectedDay }));
    } catch (e) { /* ignore */ }
  },


  // ══════════════════════════════════════════════════════
  // UTILITIES
  // ══════════════════════════════════════════════════════

  _escHtml(str) {
    if (!str) return '';
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  },

  _toast(msg, type) {
    if (typeof OwnerDev !== 'undefined' && OwnerDev._toast) {
      OwnerDev._toast(msg, type);
    } else {
      console.log('[ASRF Toast]', type, msg);
    }
  }
};
