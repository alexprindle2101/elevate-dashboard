// ═══════════════════════════════════════════════════════
// ELEVATE — Orders Management
// Displays individual order rows with filters, notes
// ═══════════════════════════════════════════════════════

const Orders = {

  // ── State ──
  _orders: [],
  _mode: 'all',
  _loading: false,
  _expandedDsi: null,       // currently expanded DSI for drill-down
  _detailCache: {},          // dsi → device detail array (client-side cache)
  _loadingDetail: false,

  // ── Fetch orders from server ──
  async fetchOrders(config, mode) {
    this._mode = mode;
    this._loading = true;
    this._expandedDsi = null;
    try {
      const email = (mode === 'my')
        ? (Roster.getEmail(App.state.currentPersona) || App.state.currentEmail)
        : null;
      this._orders = await SheetsAPI.fetchOrders(config, email);
      // Attach Tableau data to each order
      const tableauKeys = Object.keys(App.state.tableauDsi);
      let matched = 0;
      this._orders.forEach(o => {
        o.tableau = App.state.tableauDsi[o.dsi] || null;
        if (o.tableau) matched++;
      });
      console.log(`[Orders] ${matched}/${this._orders.length} orders matched Tableau DSIs`);
      if (this._orders.length > 0 && matched === 0 && tableauKeys.length > 0) {
        console.log('[Orders] Sample order DSI:', JSON.stringify(this._orders[0].dsi));
        console.log('[Orders] Sample Tableau DSI:', JSON.stringify(tableauKeys[0]));
      }
    } catch (err) {
      console.error('Failed to fetch orders:', err);
      this._orders = [];
    } finally {
      this._loading = false;
    }
  },

  // ── Status override logic ──
  // Returns { label, color, source, badges }
  _getEffectiveStatus(order) {
    // If Order Log col AM has a non-Pending value → manual override
    if (order.status && order.status !== 'Pending') {
      const color = order.status === 'Active' ? 'var(--green)'
        : order.status === 'Cancelled' ? 'var(--red)'
        : order.status === 'Complete' ? 'var(--sc-cyan)'
        : 'var(--yellow)';
      return { label: order.status, color, source: 'override', badges: null };
    }

    // If Tableau data exists → derive from device statuses
    if (order.tableau && order.tableau.statusCounts) {
      // Remap device statuses (fiber→Pending Install, air→Active, posted+approved→Active)
      const counts = this._remapDeviceStatuses(order.tableau);
      const total = order.tableau.totalDevices || 0;
      return { label: null, color: null, source: 'tableau', badges: counts, total };
    }

    // Default fallback
    return { label: 'N/A', color: 'var(--silver-dim)', source: 'default', badges: null };
  },

  // ── Remap device statuses based on product type ──
  _remapDeviceStatuses(tableau) {
    if (!tableau.devices || tableau.devices.length === 0) {
      return { ...tableau.statusCounts };
    }
    const ACTIVE_STATUSES = ['Posted', 'Delivered', 'Confirmed'];
    const KEEP_AS_IS = ['Canceled', 'Disconnected'];
    const remapped = {};
    tableau.devices.forEach(d => {
      const pt = (d.productType || '').toUpperCase();
      const isFiber = pt.includes('INTERNET');
      const isAir = pt.includes('AIR') || pt.includes('AWB');
      let status = d.dtrStatus || '';
      const orderStatus = (d.orderStatus || '').toLowerCase();

      // AIR/AWB: if firstStreaming exists → Active
      if (isAir && d.firstStreaming) {
        status = 'Active';
      }
      // "Posted" + "Approved" orderStatus → Active (no "Posted" label)
      else if (status === 'Posted' && orderStatus === 'approved') {
        status = 'Active';
      }
      // Fiber: pending-type → Pending Install
      else if (isFiber && status && !ACTIVE_STATUSES.includes(status) && !KEEP_AS_IS.includes(status)) {
        status = 'Pending Install';
      }

      remapped[status] = (remapped[status] || 0) + 1;
    });
    return remapped;
  },

  // ── Format date string to MM/DD/YYYY ──
  _formatDate(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return `${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}/${d.getFullYear()}`;
  },

  // ── Status color map for Tableau DTR statuses ──
  _dtrStatusColor(status) {
    const map = {
      'Active': 'var(--green)', 'Posted': 'var(--green)', 'Delivered': 'var(--yellow)', 'Confirmed': 'var(--green)',
      'Shipped': 'var(--sc-cyan)', 'Scheduled': 'var(--sc-cyan)',
      'Open': 'var(--yellow)', 'Pending': 'var(--yellow)',
      'Port Approved': 'var(--blue-core)', 'Porting Issue': '#cc6600', 'Pending Install': 'var(--sc-teal)', 'BYOD': 'var(--blue-core)', 'Backordered': 'var(--orange)',
      'Canceled': '#E5564A', 'Cancelled': '#E5564A', 'Disconnected': '#9B3030'
    };
    return map[status] || 'var(--silver-dim)';
  },

  // ── Render mini status badges for Tableau source ──
  _renderStatusBadges(badges, total) {
    const parts = [];
    // Sort: active first, then pending, then bad
    const order = ['Active', 'Posted', 'Delivered', 'Confirmed', 'Shipped', 'Scheduled', 'Open', 'Pending', 'Port Approved', 'Porting Issue', 'Pending Install', 'BYOD', 'Backordered', 'Canceled', 'Disconnected'];
    const sorted = Object.entries(badges).sort((a, b) => {
      const ai = order.indexOf(a[0]), bi = order.indexOf(b[0]);
      return (ai === -1 ? 99 : ai) - (bi === -1 ? 99 : bi);
    });

    sorted.forEach(([status, count]) => {
      const color = this._dtrStatusColor(status);
      // Abbreviate long status names
      const short = status === 'Disconnected' ? 'Disco'
        : status === 'Port Approved' ? 'Port'
        : status === 'Porting Issue' ? 'Port Issue'
        : status === 'Pending Install' ? 'P. Install'
        : status === 'Backordered' ? 'B/O'
        : status;
      parts.push(`<span style="font-size:10px;font-weight:700;letter-spacing:0.5px;color:${color};background:${color}18;border:1px solid ${color}44;border-radius:4px;padding:2px 6px;white-space:nowrap">${count} ${short}</span>`);
    });

    return `<div style="display:flex;flex-wrap:wrap;gap:3px;justify-content:center">${parts.join('')}</div>`;
  },

  // ── Render the orders page ──
  renderOrdersPage(mode, config) {
    const pageId = (mode === 'all') ? 'all-orders-page' : 'my-orders-page';
    const page = document.getElementById(pageId);
    if (!page) return;
    page.style.display = 'block';

    const subtitle = document.getElementById(pageId + '-subtitle');
    if (subtitle) {
      subtitle.textContent = this._loading
        ? 'Loading orders...'
        : `${this._orders.length} orders \u00b7 Past 30 days`;
    }

    // Populate rep filter dropdown (All Orders only)
    if (mode === 'all') {
      const repSel = document.getElementById('all-orders-filter-rep');
      if (repSel) {
        const reps = [...new Set((App.state.people || []).map(p => p.name).filter(Boolean))].sort();
        const current = repSel.value;
        repSel.innerHTML = '<option value="">All Reps</option>'
          + reps.map(r => `<option value="${r}">${r}</option>`).join('');
        if (current) repSel.value = current;
      }
    }

    this.applyFilters(mode);
  },

  // ── Apply filters and re-render rows ──
  applyFilters(mode) {
    const prefix = (mode === 'all') ? 'all-orders' : 'my-orders';
    const search = (document.getElementById(prefix + '-search')?.value || '').toLowerCase().trim();
    const statusFilter = document.getElementById(prefix + '-filter-status')?.value || '';
    const repFilter = document.getElementById(prefix + '-filter-rep')?.value || '';
    const dateFilter = document.getElementById(prefix + '-filter-date')?.value || '';
    const productFilter = document.getElementById(prefix + '-filter-product')?.value || '';
    const hideCompleted = document.getElementById(prefix + '-hide-completed')?.checked || false;
    const hideNoted = document.getElementById(prefix + '-hide-noted')?.checked || false;

    const COMPLETED_STATUSES = ['active', 'canceled', 'disconnected'];
    const MONTHS = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};

    let filtered = this._orders.filter(o => {
      if (search) {
        const haystack = (o.repName + ' ' + o.dsi).toLowerCase();
        if (!haystack.includes(search)) return false;
      }
      // Status filter — matches Tableau DTR statuses or col AM override
      if (statusFilter) {
        if (o.tableau && o.tableau.statusCounts && o.tableau.statusCounts[statusFilter]) {
          // Has this DTR status in Tableau data — passes
        } else if (o.status === statusFilter) {
          // Matches col AM override — passes
        } else {
          return false;
        }
      }
      if (repFilter && o.repName !== repFilter) return false;
      if (productFilter && !(o[productFilter] > 0)) return false;
      if (dateFilter) {
        const orderDate = new Date(o.dateOfSale);
        const now = new Date();
        now.setHours(0, 0, 0, 0);
        if (dateFilter === 'today') {
          const todayStr = now.toISOString().split('T')[0];
          if (o.dateOfSale !== todayStr) return false;
        } else if (dateFilter === 'yesterday') {
          const yesterday = new Date(now.getTime() - 86400000);
          const yStr = yesterday.toISOString().split('T')[0];
          if (o.dateOfSale !== yStr) return false;
        } else {
          const days = parseInt(dateFilter);
          const cutoff = new Date(now.getTime() - days * 86400000);
          if (orderDate < cutoff) return false;
        }
      }
      // Hide completed — orders where ALL remapped device statuses are Active, Canceled, or Disconnected
      if (hideCompleted) {
        if (o.tableau && o.tableau.devices && o.tableau.devices.length > 0) {
          const remapped = Orders._remapDeviceStatuses(o.tableau);
          const allCompleted = Object.keys(remapped).every(s =>
            COMPLETED_STATUSES.includes(s.toLowerCase())
          );
          if (allCompleted) return false;
        } else if (o.status) {
          if (COMPLETED_STATUSES.includes((o.status || '').toLowerCase())) return false;
        }
      }
      // Hide recently noted — orders with an admin/owner/manager note within the past 3 days
      if (hideNoted && o.notes) {
        const lines = o.notes.split('\n');
        const now = new Date();
        now.setHours(0, 0, 0, 0);
        const cutoff = new Date(now.getTime() - 3 * 86400000);
        const hasRecentAdminNote = lines.some(line => {
          const m = line.match(/^\[(\w{3})\s+(\d{1,2})\s*[—–\-]\s*(.+?)\]/);
          if (!m) return false;
          const mi = MONTHS[m[1]];
          if (mi === undefined) return false;
          // Check if author is an admin-level role (use _roleKey, not .role which is the display label)
          const author = m[3].trim();
          const adminRoles = ['admin', 'owner', 'superadmin', 'manager'];
          const person = (App.state.people || []).find(p => p.name === author);
          // Also check roster in case admin has no sales data and isn't in people array
          const rosterEntry = !person ? Object.values(App.state.roster || {}).find(r => r.name === author) : null;
          const roleKey = person ? person._roleKey : (rosterEntry ? rosterEntry.rank : null);
          if (!roleKey || !adminRoles.includes(roleKey)) return false;
          const noteDate = new Date(now.getFullYear(), mi, parseInt(m[2]));
          if (noteDate > now) noteDate.setFullYear(noteDate.getFullYear() - 1);
          return noteDate >= cutoff;
        });
        if (hasRecentAdminNote) return false;
      }
      return true;
    });

    const countEl = document.getElementById(prefix + '-count');
    if (countEl) {
      countEl.textContent = filtered.length === this._orders.length
        ? `Showing all ${this._orders.length}`
        : `Showing ${filtered.length} of ${this._orders.length}`;
    }

    this._renderOrderRows(filtered, mode);
  },

  // ── Check if current role can edit orders ──
  _canEdit() {
    const role = App.state.currentRole;
    return ['superadmin', 'owner', 'manager', 'admin'].includes(role);
  },

  // ── Render filtered order rows ──
  _renderOrderRows(orders, mode) {
    const prefix = (mode === 'all') ? 'all-orders' : 'my-orders';
    const tbody = document.getElementById(prefix + '-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    const canEdit = mode === 'all' && this._canEdit();
    const baseCols = (mode === 'all') ? 7 : 6;
    const colSpan = canEdit ? baseCols + 1 : baseCols;

    if (orders.length === 0) {
      tbody.innerHTML = `<tr><td colspan="${colSpan}" style="text-align:center;color:var(--silver-dim);padding:32px;font-family:'Cerebri Sans','DM Sans','Inter',sans-serif;font-size:14px">No orders match the current filters</td></tr>`;
      return;
    }

    orders.forEach(o => {
      const soldParts = [];
      OFFICE_CONFIG.columns.products.forEach(prod => {
        const val = o[prod.key] || 0;
        if (val > 0) {
          let label = prod.type === 'boolean' ? prod.label : `${prod.label} x${val}`;
          // Append package detail from new schema fields
          if (prod.key === 'fiber' && o.fiberPackage) label += ` (${o.fiberPackage})`;
          if (prod.key === 'voip' && o.oomaPackage) label += ` (${o.oomaPackage})`;
          // Cell new/BYOD breakdown removed — just show "Cell xN"
          soldParts.push(label);
        }
      });
      // DTV (excluded from main product list but show if present)
      if (o.dtv > 0) {
        let dtvLabel = 'DTV';
        if (o.dtvPackage) dtvLabel += ` (${o.dtvPackage})`;
        soldParts.push(dtvLabel);
      }
      const soldStr = soldParts.length > 0 ? soldParts.join(', ') : '\u2014';
      // Tower badge (tracked but excluded from leaderboard)
      const towerBadge = (o.orderChannel === 'Tower')
        ? '<span style="font-size:9px;font-weight:700;letter-spacing:0.5px;color:var(--orange);background:rgba(249,115,22,0.12);border:1px solid rgba(249,115,22,0.3);border-radius:4px;padding:1px 5px;margin-right:4px;text-transform:uppercase">TOWER</span>'
        : '';
      // Codes-used indicator
      const codesIndicator = o.codesUsedBy
        ? `<div style="font-size:10px;color:var(--orange);margin-top:2px">Codes: ${this._escapeHtml((App.state?.roster?.[o.codesUsedBy]?.name) || o.codesUsedBy)}</div>`
        : '';
      // Campaign badge (e.g., Ooma orders get a distinct badge)
      const campaignBadge = (o.campaign && o.campaign !== 'attb2b')
        ? `<span style="font-size:9px;font-weight:700;letter-spacing:0.5px;color:var(--sc-teal);background:rgba(0,229,204,0.12);border:1px solid rgba(0,229,204,0.3);border-radius:4px;padding:1px 5px;margin-right:4px;text-transform:uppercase">${this._escapeHtml(o.campaign)}</span>`
        : '';

      // Use effective status (with Tableau override logic)
      const effective = this._getEffectiveStatus(o);
      let statusHtml;
      if (effective.source === 'tableau') {
        statusHtml = this._renderStatusBadges(effective.badges, effective.total);
      } else {
        const overrideTag = effective.source === 'override'
          ? ' <span style="font-size:9px;color:var(--silver-dim);font-weight:400">(Override)</span>' : '';
        statusHtml = `<span style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:${effective.color};background:${effective.color}18;border:1px solid ${effective.color}44;border-radius:6px;padding:3px 10px">${this._escapeHtml(effective.label)}</span>${overrideTag}`;
      }

      const noteLines = o.notes ? o.notes.split('\n') : [];
      const notePreview = noteLines.length > 0
        ? noteLines[noteLines.length - 1].substring(0, 60) + (noteLines[noteLines.length - 1].length > 60 ? '...' : '')
        : '';
      const noteCount = noteLines.length;

      const tickets = o.tickets || [];
      const openTickets = tickets.filter(t => !t.resolved).length;

      const escapedDsi = (o.dsi || '').replace(/'/g, "\\'");
      const hasDrillDown = !!o.tableau;
      const isExpanded = this._expandedDsi === o.dsi;
      const dsiClickable = hasDrillDown
        ? `<span class="name-link" style="cursor:pointer;color:var(--blue-core)" onclick="Orders.toggleDrillDown('${escapedDsi}','${prefix}')">${this._escapeHtml(o.dsi) || '\u2014'} <span style="font-size:9px">${isExpanded ? '\u25B2' : '\u25BC'}</span></span>`
        : `<span style="color:var(--silver)">${this._escapeHtml(o.dsi) || '\u2014'}</span>`;

      const tr = document.createElement('tr');
      tr.className = 'orders-row';
      tr.setAttribute('data-dsi', o.dsi || '');
      tr.innerHTML = `
        ${mode === 'all' ? `<td style="padding:10px 12px;font-weight:700;color:var(--white)">${this._escapeHtml(o.repName)}</td>` : ''}
        <td style="padding:10px 12px">${dsiClickable}</td>
        <td style="padding:10px 12px;color:var(--silver)">${o.dateOfSale}</td>
        <td style="padding:10px 12px;color:var(--white)">${towerBadge}${campaignBadge}${soldStr}${codesIndicator}</td>
        <td style="padding:10px 12px;text-align:center">${statusHtml}</td>
        <td style="padding:10px 12px;font-size:11px;color:var(--silver-dim);max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${this._escapeHtml(notePreview)}${noteCount > 1 ? ` <span style="color:var(--blue-core)">(${noteCount})</span>` : ''}</td>
        <td style="padding:10px 8px;text-align:right;white-space:nowrap">
          ${canEdit ? `<button onclick="Orders.openEditModal(${o.rowIndex})" style="background:none;border:1px solid rgba(0,0,0,0.3);border-radius:6px;padding:4px 10px;color:var(--blue-core);font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;cursor:pointer;margin-right:4px">Edit</button>` : ''}
          <button onclick="Orders.openNoteModal(${o.rowIndex},'${escapedDsi}')"
            style="background:none;border:1px solid rgba(0,0,0,0.3);border-radius:6px;padding:4px 10px;color:var(--blue-core);font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;cursor:pointer">Notes${noteCount > 0 ? ' (' + noteCount + ')' : ''}</button>${openTickets > 0 ? `<span style="display:inline-flex;align-items:center;gap:2px;margin-left:6px;font-size:11px;font-weight:700;color:var(--orange);background:rgba(249,115,22,0.12);border:1px solid rgba(249,115,22,0.3);border-radius:6px;padding:3px 8px;font-family:'Helvetica Neue','Inter',sans-serif">\uD83C\uDFAB ${openTickets}</span>` : ''}
        </td>`;
      tbody.appendChild(tr);

      // If this DSI is expanded, render drill-down row
      if (isExpanded) {
        this._renderDrillDownRow(tbody, o.dsi, colSpan);
      }
    });
  },

  // ── Toggle drill-down for a DSI ──
  toggleDrillDown(dsi, prefix) {
    if (this._expandedDsi === dsi) {
      this._expandedDsi = null;
    } else {
      this._expandedDsi = dsi;
      // Use embedded devices from summary (already loaded with dashboard)
      if (!this._detailCache[dsi]) {
        const tableau = App.state.tableauDsi[dsi];
        this._detailCache[dsi] = (tableau && tableau.devices) ? tableau.devices : [];
      }
    }
    this.applyFilters(this._mode);
  },

  // ── Render the inline drill-down detail row ──
  _renderDrillDownRow(tbody, dsi, colSpan) {
    const tr = document.createElement('tr');
    tr.className = 'drill-down-row';
    tr.style.background = 'rgba(44,110,106,0.04)';

    const devices = this._detailCache[dsi];
    if (!devices || devices.length === 0) {
      tr.innerHTML = `<td colspan="${colSpan}" style="padding:12px 24px;font-size:12px;color:var(--silver-dim)">No device detail available for this DSI</td>`;
      tbody.appendChild(tr);
      return;
    }

    // Check if any device has an install date (for fiber column)
    const hasFiber = devices.some(d => (d.productType || '').toUpperCase().includes('INTERNET'));

    const escapedDsi = dsi.replace(/'/g, "\\'");
    let html = `<td colspan="${colSpan}" style="padding:8px 16px">
      <div style="display:flex;justify-content:flex-end;margin-bottom:6px">
        <button onclick="Orders._openSaraPlus('${escapedDsi}')" style="background:rgba(44,110,106,0.1);border:1px solid rgba(44,110,106,0.3);border-radius:6px;padding:4px 12px;color:var(--sc-cyan);font-family:'Neue Haas Grotesk','Helvetica Neue','Inter',sans-serif;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;cursor:pointer;display:flex;align-items:center;gap:4px">
          <span style="font-size:12px">&#x1F50D;</span> SARA+
        </button>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:11px;font-family:'Cerebri Sans','DM Sans','Inter',sans-serif">
        <thead>
          <tr style="color:var(--silver-dim);text-transform:uppercase;letter-spacing:1px;font-size:10px;font-weight:700">
            <th style="padding:4px 8px;text-align:left">SPE</th>
            <th style="padding:4px 8px;text-align:left">Product</th>
            <th style="padding:4px 8px;text-align:center">CRU/IRU</th>
            <th style="padding:4px 8px;text-align:center">Status</th>
            <th style="padding:4px 8px;text-align:left">Device</th>
            ${hasFiber ? '<th style="padding:4px 8px;text-align:left">Install Date</th>' : ''}
            <th style="padding:4px 8px;text-align:left">Disco Reason</th>
          </tr>
        </thead><tbody>`;

    const ACTIVE_STATUSES = ['Posted', 'Delivered', 'Confirmed'];
    const KEEP_AS_IS = ['Canceled', 'Disconnected'];

    devices.forEach(d => {
      const pt = (d.productType || '').toUpperCase();
      const isFiber = pt.includes('INTERNET');
      const isAir = pt.includes('AIR') || pt.includes('AWB');
      const orderStatus = (d.orderStatus || '').toLowerCase();
      let displayStatus = d.dtrStatus;

      // AIR/AWB: if firstStreaming exists → Active
      if (isAir && d.firstStreaming) {
        displayStatus = 'Active';
      }
      // Posted + Approved → Active
      else if (displayStatus === 'Posted' && orderStatus === 'approved') {
        displayStatus = 'Active';
      }
      // Fiber: pending-type → Pending Install
      else if (isFiber && displayStatus && !ACTIVE_STATUSES.includes(displayStatus) && !KEEP_AS_IS.includes(displayStatus)) {
        displayStatus = 'Pending Install';
      }
      const statusColor = this._dtrStatusColor(displayStatus);
      html += `<tr style="border-top:1px solid rgba(0,0,0,0.1)">
        <td style="padding:4px 8px;color:var(--silver)">${this._escapeHtml(d.spe)}</td>
        <td style="padding:4px 8px;color:var(--white)">${this._escapeHtml(d.productType)}</td>
        <td style="padding:4px 8px;text-align:center;color:var(--silver)">${this._escapeHtml(d.cruIru)}</td>
        <td style="padding:4px 8px;text-align:center">
          <span style="font-size:10px;font-weight:700;color:${statusColor};background:${statusColor}18;border:1px solid ${statusColor}44;border-radius:4px;padding:2px 6px">${this._escapeHtml(displayStatus)}</span>
        </td>
        <td style="padding:4px 8px;color:var(--silver);max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${this._escapeHtml(d.phone) || '\u2014'}</td>
        ${hasFiber ? `<td style="padding:4px 8px;color:var(--silver)">${isFiber && d.installDate ? this._formatDate(d.installDate) : '\u2014'}</td>` : ''}
        <td style="padding:4px 8px;color:${d.discoReason ? 'var(--red)' : 'var(--silver-dim)'}">${this._escapeHtml(d.discoReason) || '\u2014'}</td>
      </tr>`;
    });

    html += '</tbody></table></td>';
    tr.innerHTML = html;
    tbody.appendChild(tr);
  },

  // ── Open SARA+ with DSI copied to clipboard (reuse same tab) ──
  _saraWindow: null,
  _openSaraPlus(dsi) {
    navigator.clipboard.writeText(dsi).then(() => {
      this._showToast(`${dsi} copied — paste into SARA+ search`);
      this._openSaraWindow();
    }).catch(() => {
      this._openSaraWindow();
    });
  },
  _openSaraWindow() {
    // Reuse the same named window so session persists
    if (this._saraWindow && !this._saraWindow.closed) {
      this._saraWindow.focus();
    } else {
      this._saraWindow = window.open('https://www.saraplus.com', 'saraplus');
    }
  },

  _showToast(msg) {
    let toast = document.getElementById('sc-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'sc-toast';
      toast.style.cssText = 'position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#242124;color:#FEFAF3;padding:10px 20px;border-radius:8px;font-family:"Cerebri Sans","DM Sans","Inter",sans-serif;font-size:13px;font-weight:600;letter-spacing:0.5px;z-index:99999;opacity:0;transition:opacity 0.3s;pointer-events:none;box-shadow:0 4px 12px rgba(0,0,0,0.3)';
      document.body.appendChild(toast);
    }
    toast.textContent = msg;
    toast.style.opacity = '1';
    clearTimeout(this._toastTimer);
    this._toastTimer = setTimeout(() => { toast.style.opacity = '0'; }, 3000);
  },

  // ── Note Modal ──
  _activeRowIndex: null,
  _activeDsi: '',
  _savingNote: false,

  openNoteModal(rowIndex, dsi) {
    this._activeRowIndex = rowIndex;
    this._activeDsi = dsi;
    const modal = document.getElementById('order-note-modal');
    if (!modal) return;
    modal.style.display = 'flex';

    const order = this._orders.find(o => o.rowIndex === rowIndex);
    const notesDisplay = document.getElementById('order-note-history');
    if (notesDisplay) {
      if (order && order.notes) {
        notesDisplay.innerHTML = order.notes.split('\n').map(line =>
          `<div style="padding:6px 0;border-bottom:1px solid rgba(0,0,0,0.1);font-size:12px;color:var(--silver);line-height:1.5">${this._escapeHtml(line)}</div>`
        ).join('');
      } else {
        notesDisplay.innerHTML = '<div style="color:var(--silver-dim);font-size:12px;padding:12px 0">No notes yet</div>';
      }
    }

    // Render tickets
    this._renderTicketsList(order);

    const dsiLabel = document.getElementById('order-note-dsi');
    if (dsiLabel) dsiLabel.textContent = dsi ? `DSI: ${dsi}` : 'Order Note';

    const input = document.getElementById('order-note-input');
    if (input) input.value = '';

    const errorEl = document.getElementById('order-note-error');
    if (errorEl) errorEl.textContent = '';

    // Clear ticket inputs
    const ticketIdInput = document.getElementById('ticket-id-input');
    const ticketTextInput = document.getElementById('ticket-text-input');
    const ticketError = document.getElementById('ticket-error');
    if (ticketIdInput) ticketIdInput.value = '';
    if (ticketTextInput) ticketTextInput.value = '';
    if (ticketError) ticketError.textContent = '';
  },

  closeNoteModal() {
    const modal = document.getElementById('order-note-modal');
    if (modal) modal.style.display = 'none';
    this._activeRowIndex = null;
  },

  // ── Ticket rendering ──
  _renderTicketsList(order) {
    const container = document.getElementById('order-note-tickets');
    if (!container) return;

    const tickets = (order && order.tickets) || [];
    if (tickets.length === 0) {
      container.innerHTML = '<div style="color:var(--silver-dim);font-size:12px;padding:12px 0">No tickets</div>';
      return;
    }

    container.innerHTML = tickets.map(t => {
      const escapedId = this._escapeHtml(t.id).replace(/'/g, "\\'");
      const resolvedStyle = t.resolved ? 'text-decoration:line-through;opacity:0.6' : '';
      return `<div style="padding:8px 0;border-bottom:1px solid rgba(0,0,0,0.1);display:flex;align-items:flex-start;gap:8px">
        <input type="checkbox" ${t.resolved ? 'checked' : ''} onchange="Orders.toggleTicket('${escapedId}')" style="cursor:pointer;margin-top:2px;flex-shrink:0">
        <div style="flex:1;min-width:0;${resolvedStyle}">
          <div style="font-size:13px;color:var(--white);font-weight:700;font-family:'Neue Montreal','Inter',sans-serif">${this._escapeHtml(t.id)} <span style="font-weight:400;color:var(--silver)">\u2014 ${this._escapeHtml(t.text || '')}</span></div>
          <div style="font-size:11px;color:var(--silver-dim);margin-top:2px">${this._escapeHtml(t.date || '')} \u2014 ${this._escapeHtml(t.author || '')}</div>
        </div>
      </div>`;
    }).join('');
  },

  // ── Add ticket ──
  _savingTicket: false,

  async addTicket() {
    if (this._savingTicket || !this._activeRowIndex) return;

    const idInput = document.getElementById('ticket-id-input');
    const textInput = document.getElementById('ticket-text-input');
    const errorEl = document.getElementById('ticket-error');
    const ticketId = (idInput?.value || '').trim();
    const ticketText = (textInput?.value || '').trim();

    if (!ticketId) {
      if (errorEl) errorEl.textContent = 'Ticket ID is required';
      return;
    }

    this._savingTicket = true;
    const btn = document.getElementById('ticket-submit-btn');
    if (btn) { btn.textContent = 'Saving...'; btn.style.opacity = '0.5'; }

    try {
      const result = await SheetsAPI.post(OFFICE_CONFIG, 'addTicket', {
        rowIndex: this._activeRowIndex,
        ticketId: ticketId,
        ticketText: ticketText,
        authorName: App.state.currentPersona
      });

      if (result.data?.error) {
        if (errorEl) errorEl.textContent = result.data.error;
        return;
      }

      // Update local cache
      const order = this._orders.find(o => o.rowIndex === this._activeRowIndex);
      if (order && result.data?.tickets) {
        order.tickets = result.data.tickets;
      }

      // Re-render modal and table
      this.openNoteModal(this._activeRowIndex, this._activeDsi);
      this.applyFilters(this._mode);
      App.showToast('Ticket added');
    } catch (err) {
      if (errorEl) errorEl.textContent = 'Failed: ' + err.message;
    } finally {
      this._savingTicket = false;
      if (btn) { btn.textContent = 'ADD TICKET'; btn.style.opacity = ''; }
    }
  },

  // ── Toggle ticket resolved ──
  async toggleTicket(ticketId) {
    if (!this._activeRowIndex) return;

    try {
      const result = await SheetsAPI.post(OFFICE_CONFIG, 'toggleTicket', {
        rowIndex: this._activeRowIndex,
        ticketId: ticketId
      });

      if (result.data?.error) {
        App.showToast(result.data.error);
        return;
      }

      const order = this._orders.find(o => o.rowIndex === this._activeRowIndex);
      if (order && result.data?.tickets) {
        order.tickets = result.data.tickets;
      }

      this._renderTicketsList(order);
      this.applyFilters(this._mode);
    } catch (err) {
      App.showToast('Failed: ' + err.message);
    }
  },

  async submitNote() {
    if (this._savingNote || !this._activeRowIndex) return;

    const input = document.getElementById('order-note-input');
    const noteText = (input?.value || '').trim();
    const errorEl = document.getElementById('order-note-error');

    if (!noteText) {
      if (errorEl) errorEl.textContent = 'Please enter a note';
      return;
    }

    this._savingNote = true;
    const btn = document.getElementById('order-note-submit-btn');
    if (btn) { btn.textContent = 'Saving...'; btn.style.opacity = '0.5'; }

    try {
      const result = await SheetsAPI.post(OFFICE_CONFIG, 'writeOrderNote', {
        rowIndex: this._activeRowIndex,
        authorName: App.state.currentPersona,
        noteText: noteText
      });

      if (result.data?.error) {
        if (errorEl) errorEl.textContent = result.data.error;
        return;
      }

      // Update local cache
      const order = this._orders.find(o => o.rowIndex === this._activeRowIndex);
      if (order && result.data?.notes) {
        order.notes = result.data.notes;
      }

      // Re-render modal and table
      this.openNoteModal(this._activeRowIndex, this._activeDsi);
      if (this._mode === 'payroll') {
        App._filterPayrollOrders();
      } else {
        this.applyFilters(this._mode);
      }
      App.showToast('Note added');
    } catch (err) {
      if (errorEl) errorEl.textContent = 'Failed: ' + err.message;
    } finally {
      this._savingNote = false;
      if (btn) { btn.textContent = 'ADD NOTE'; btn.style.opacity = ''; }
    }
  },

  // ── Edit Modal ──
  _editRowIndex: null,
  _savingEdit: false,

  openEditModal(rowIndex) {
    this._editRowIndex = rowIndex;
    const order = this._orders.find(o => o.rowIndex === rowIndex);
    if (!order) return;

    const modal = document.getElementById('order-edit-modal');
    if (!modal) return;
    modal.style.display = 'flex';

    const setVal = (id, val) => { const el = document.getElementById(id); if (el) el.value = val; };
    setVal('order-edit-rep', order.repName);
    setVal('order-edit-dsi', order.dsi);
    setVal('order-edit-date', order.dateOfSale);
    setVal('order-edit-air', order.air);
    setVal('order-edit-cell', order.cell);
    setVal('order-edit-fiber', order.fiber);
    setVal('order-edit-voip', order.voip);
    setVal('order-edit-status', order.status);

    const errorEl = document.getElementById('order-edit-error');
    if (errorEl) errorEl.textContent = '';
  },

  closeEditModal() {
    const modal = document.getElementById('order-edit-modal');
    if (modal) modal.style.display = 'none';
    this._editRowIndex = null;
  },

  async submitEdit() {
    if (this._savingEdit || !this._editRowIndex) return;

    const getVal = (id) => document.getElementById(id)?.value ?? '';
    const errorEl = document.getElementById('order-edit-error');

    const payload = {
      rowIndex: this._editRowIndex,
      repName: getVal('order-edit-rep').trim(),
      dsi: getVal('order-edit-dsi').trim(),
      dateOfSale: getVal('order-edit-date'),
      air: Number(getVal('order-edit-air')) || 0,
      cell: Number(getVal('order-edit-cell')) || 0,
      fiber: Number(getVal('order-edit-fiber')) || 0,
      voip: Number(getVal('order-edit-voip')) || 0,
      status: getVal('order-edit-status').trim()
    };

    if (!payload.repName) {
      if (errorEl) errorEl.textContent = 'Rep name is required';
      return;
    }

    this._savingEdit = true;
    const btn = document.getElementById('order-edit-submit-btn');
    if (btn) { btn.textContent = 'Saving...'; btn.style.opacity = '0.5'; }

    try {
      const result = await SheetsAPI.post(OFFICE_CONFIG, 'updateOrder', payload);

      if (result.data?.error) {
        if (errorEl) errorEl.textContent = result.data.error;
        return;
      }

      // Update local cache
      const order = this._orders.find(o => o.rowIndex === this._editRowIndex);
      if (order) {
        order.repName = payload.repName;
        order.dsi = payload.dsi;
        order.dateOfSale = payload.dateOfSale;
        order.air = payload.air;
        order.cell = payload.cell;
        order.fiber = payload.fiber;
        order.voip = payload.voip;
        order.status = payload.status;
      }

      this.closeEditModal();
      this.applyFilters(this._mode);
      App.showToast('Order updated');
    } catch (err) {
      if (errorEl) errorEl.textContent = 'Failed: ' + err.message;
    } finally {
      this._savingEdit = false;
      if (btn) { btn.textContent = 'SAVE CHANGES'; btn.style.opacity = ''; }
    }
  },

  _escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }
};
