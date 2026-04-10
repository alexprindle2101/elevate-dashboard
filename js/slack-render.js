// ═══════════════════════════════════════════════════════
// Aptel Slack Channel Auditor — Render Module
// Pure DOM rendering — no state mutation, no data fetching
// ═══════════════════════════════════════════════════════

const SlackRender = {

  // ── HTML escape ──
  _esc(str) {
    const d = document.createElement('div');
    d.textContent = str ?? '';
    return d.innerHTML;
  },

  // ── Show / hide helpers ──
  show(id) { document.getElementById(id).style.display = ''; },
  hide(id) { document.getElementById(id).style.display = 'none'; },

  // ── Loading screen ──
  showLoading(msg) {
    document.getElementById('loading-text').textContent = msg || 'Loading...';
    this.show('loading-screen');
  },
  hideLoading() { this.hide('loading-screen'); },

  // ── Error banner ──
  showError(msg) {
    document.getElementById('error-msg').textContent = msg;
    this.show('error-banner');
  },
  hideError() { this.hide('error-banner'); },

  // ── Status pill ──
  setStatus(text, connected) {
    const pill = document.getElementById('status-pill');
    const span = document.getElementById('status-text');
    span.textContent = text;
    pill.className = connected ? 'status-pill connected' : 'status-pill';
  },

  // ── Excel info bar ──
  renderExcelInfo(excelData) {
    if (!excelData) { this.hide('excel-info'); return; }
    const p = excelData.people.length;
    const d = Object.keys(excelData.deptMappings || {}).length;
    const r = Object.keys(excelData.roleMappings || {}).length;
    const allChannels = new Set([
      ...Object.values(excelData.deptMappings || {}).flat(),
      ...Object.values(excelData.roleMappings || {}).flat(),
    ]);
    document.getElementById('excel-info-text').textContent =
      `Loaded: ${p} people, ${d} departments, ${r} role combos, ${allChannels.size} unique channels`;
    this.show('excel-info');
    this.hide('empty-state');
  },

  // ── Summary cards ──
  renderSummary(results) {
    if (!results || !results.length) { this.hide('summary-section'); return; }

    const total = results.length;
    const ok = results.filter(r => r.status === 'match').length;
    const mismatches = results.filter(r => r.status === 'missing' || r.status === 'extra').length;
    const notFound = results.filter(r => r.status === 'notFound').length;
    const noRole = results.filter(r => r.status === 'noMapping').length;

    const el = document.getElementById('summary-section');
    el.innerHTML = `
      <div class="stat-card">
        <div class="stat-card-value">${total}</div>
        <div class="stat-card-label">Total People</div>
      </div>
      <div class="stat-card ok">
        <div class="stat-card-value">${ok}</div>
        <div class="stat-card-label">Matched</div>
      </div>
      <div class="stat-card danger">
        <div class="stat-card-value">${mismatches}</div>
        <div class="stat-card-label">Mismatches</div>
      </div>
      <div class="stat-card">
        <div class="stat-card-value">${notFound}</div>
        <div class="stat-card-label">Not in Slack</div>
      </div>
      ${noRole ? `
      <div class="stat-card warning">
        <div class="stat-card-value">${noRole}</div>
        <div class="stat-card-label">No Role Mapping</div>
      </div>` : ''}
    `;
    el.style.display = '';
  },

  // ── Filter tab counts ──
  updateFilterCounts(results) {
    if (!results) return;
    const all = results.length;
    const ok = results.filter(r => r.status === 'match').length;
    const mismatches = results.filter(r => r.status === 'missing' || r.status === 'extra').length;
    const notFound = results.filter(r => r.status === 'notFound' || r.status === 'noMapping').length;

    document.getElementById('count-all').textContent = all;
    document.getElementById('count-mismatches').textContent = mismatches;
    document.getElementById('count-ok').textContent = ok;
    document.getElementById('count-not-found').textContent = notFound;
  },

  // ── Active filter highlight ──
  setActiveFilter(mode) {
    document.querySelectorAll('.filter-tab').forEach(tab => {
      tab.classList.toggle('active', tab.dataset.filter === mode);
    });
  },

  // ── Comparison table ──
  renderTable(results, filterMode, searchQuery) {
    const container = document.getElementById('table-container');
    if (!results || !results.length) {
      container.innerHTML = '';
      return;
    }

    // Filter
    let rows = results;
    if (filterMode === 'mismatches') {
      rows = rows.filter(r => r.status === 'missing' || r.status === 'extra');
    } else if (filterMode === 'ok') {
      rows = rows.filter(r => r.status === 'match');
    } else if (filterMode === 'not-found') {
      rows = rows.filter(r => r.status === 'notFound' || r.status === 'noMapping');
    }

    // Search
    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      rows = rows.filter(r =>
        r.name.toLowerCase().includes(q) ||
        r.email.toLowerCase().includes(q) ||
        (r.department || '').toLowerCase().includes(q) ||
        (r.level || '').toLowerCase().includes(q)
      );
    }

    if (!rows.length) {
      container.innerHTML = `
        <div style="text-align:center;padding:60px 20px;color:var(--gray-500);">
          <div style="font-size:24px;margin-bottom:8px;">🔍</div>
          <div style="font-weight:600;">No results match your filters</div>
        </div>`;
      return;
    }

    const html = `
      <table class="slack-table">
        <thead>
          <tr>
            <th>Name</th>
            <th>Email</th>
            <th>Department</th>
            <th>Level</th>
            <th>Expected Channels</th>
            <th>Actual Channels</th>
            <th>Status</th>
          </tr>
        </thead>
        <tbody>
          ${rows.map(r => this._renderRow(r)).join('')}
        </tbody>
      </table>
    `;
    container.innerHTML = html;
    this.show('toolbar');
  },

  _renderRow(r) {
    const statusClass = SLACK_CONFIG.statusColors[r.status] || '';
    const statusLabel = SLACK_CONFIG.statusLabels[r.status] || r.status;

    // Build channel pills
    const expectedPills = this._channelPills(r.expectedChannels, r.matched, r.missing, 'expected');
    const actualPills = this._channelPills(r.actualChannels, r.matched, r.extra, 'actual');

    // Detail summary
    let detail = '';
    if (r.missing.length) detail += `<span style="color:var(--tomato);font-size:12px;font-weight:600;">+${r.missing.length} missing</span> `;
    if (r.extra.length) detail += `<span style="color:var(--amber);font-size:12px;font-weight:600;">-${r.extra.length} extra</span>`;

    return `
      <tr>
        <td class="name-cell">${this._esc(r.name)}</td>
        <td class="email-cell">${this._esc(r.email)}</td>
        <td class="role-cell">${this._esc(r.department || '')}</td>
        <td class="role-cell">${this._esc(r.level || '')}</td>
        <td><div class="channel-pills">${expectedPills}</div></td>
        <td><div class="channel-pills">${actualPills}</div></td>
        <td>
          <span class="status-badge ${statusClass}">
            <span class="dot"></span>
            ${this._esc(statusLabel)}
          </span>
          ${detail ? `<div style="margin-top:4px">${detail}</div>` : ''}
        </td>
      </tr>
    `;
  },

  _channelPills(channels, matched, highlighted, mode) {
    if (!channels || !channels.length) {
      return '<span style="color:var(--gray-400);font-size:12px;">—</span>';
    }

    return channels.map(ch => {
      let cls = 'pill-ok';
      let icon = '✓';

      if (mode === 'expected' && highlighted.includes(ch)) {
        // Channel is in expected but missing from actual
        cls = 'pill-missing';
        icon = '✗';
      } else if (mode === 'actual' && highlighted.includes(ch)) {
        // Channel is in actual but not in expected
        cls = 'pill-extra';
        icon = '?';
      } else if (matched.includes(ch)) {
        cls = 'pill-ok';
        icon = '✓';
      }

      return `<span class="channel-pill ${cls}"><span class="pill-icon">${icon}</span>#${this._esc(ch)}</span>`;
    }).join('');
  },

  // ── Skeleton loading rows ──
  renderSkeletonTable() {
    const container = document.getElementById('table-container');
    const rows = Array.from({ length: 8 }, () => `
      <div class="skeleton-row">
        <div class="skeleton skeleton-cell" style="width:120px"></div>
        <div class="skeleton skeleton-cell" style="width:180px"></div>
        <div class="skeleton skeleton-cell" style="width:80px"></div>
        <div class="skeleton skeleton-cell" style="width:200px"></div>
        <div class="skeleton skeleton-cell" style="width:200px"></div>
        <div class="skeleton skeleton-cell" style="width:80px"></div>
      </div>
    `).join('');

    container.innerHTML = `
      <div style="background:var(--white);border:1px solid var(--gray-100);border-radius:var(--radius);overflow:hidden;">
        <div style="padding:12px 16px;background:var(--gray-50);border-bottom:1px solid var(--gray-100);display:flex;gap:16px;">
          <div class="skeleton" style="width:60px;height:12px"></div>
          <div class="skeleton" style="width:60px;height:12px"></div>
          <div class="skeleton" style="width:40px;height:12px"></div>
          <div class="skeleton" style="width:120px;height:12px"></div>
          <div class="skeleton" style="width:120px;height:12px"></div>
          <div class="skeleton" style="width:60px;height:12px"></div>
        </div>
        ${rows}
      </div>
    `;
  },

  // ── Reset to empty state ──
  resetToEmpty() {
    this.hide('excel-info');
    this.hide('summary-section');
    this.hide('toolbar');
    document.getElementById('table-container').innerHTML = '';
    document.getElementById('search-input').value = '';
    this.show('empty-state');
    this.setStatus('No data loaded', false);
  },
};
