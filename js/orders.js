// ═══════════════════════════════════════════════════════
// ELEVATE — Orders Management
// Displays individual order rows with filters, notes
// ═══════════════════════════════════════════════════════

const Orders = {

  // ── State ──
  _orders: [],
  _mode: 'all',
  _loading: false,

  // ── Fetch orders from server ──
  async fetchOrders(config, mode) {
    this._mode = mode;
    this._loading = true;
    try {
      const email = (mode === 'my')
        ? (Roster.getEmail(App.state.currentPersona) || App.state.currentEmail)
        : null;
      this._orders = await SheetsAPI.fetchOrders(config, email);
    } catch (err) {
      console.error('Failed to fetch orders:', err);
      this._orders = [];
    } finally {
      this._loading = false;
    }
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

    let filtered = this._orders.filter(o => {
      if (search) {
        const haystack = (o.repName + ' ' + o.dsi).toLowerCase();
        if (!haystack.includes(search)) return false;
      }
      if (statusFilter && o.status !== statusFilter) return false;
      if (repFilter && o.repName !== repFilter) return false;
      if (productFilter && !(o[productFilter] > 0)) return false;
      if (dateFilter) {
        const orderDate = new Date(o.dateOfSale);
        const now = new Date();
        now.setHours(0, 0, 0, 0);
        if (dateFilter === 'today') {
          const todayStr = now.toISOString().split('T')[0];
          if (o.dateOfSale !== todayStr) return false;
        } else {
          const days = parseInt(dateFilter);
          const cutoff = new Date(now.getTime() - days * 86400000);
          if (orderDate < cutoff) return false;
        }
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
      tbody.innerHTML = `<tr><td colspan="${colSpan}" style="text-align:center;color:var(--silver-dim);padding:32px;font-family:'Barlow Condensed',sans-serif;font-size:14px">No orders match the current filters</td></tr>`;
      return;
    }

    orders.forEach(o => {
      const soldParts = [];
      OFFICE_CONFIG.columns.products.forEach(prod => {
        const val = o[prod.key] || 0;
        if (val > 0) {
          soldParts.push(prod.type === 'boolean' ? prod.label : `${prod.label} x${val}`);
        }
      });
      const soldStr = soldParts.length > 0 ? soldParts.join(', ') : '\u2014';

      const statusColor = o.status === 'Active' ? 'var(--green)'
        : o.status === 'Cancelled' ? 'var(--red)'
        : o.status === 'Complete' ? 'var(--sc-cyan)'
        : 'var(--yellow)';

      const noteLines = o.notes ? o.notes.split('\n') : [];
      const notePreview = noteLines.length > 0
        ? noteLines[noteLines.length - 1].substring(0, 60) + (noteLines[noteLines.length - 1].length > 60 ? '...' : '')
        : '';
      const noteCount = noteLines.length;

      const escapedDsi = (o.dsi || '').replace(/'/g, "\\'");

      const tr = document.createElement('tr');
      tr.innerHTML = `
        ${mode === 'all' ? `<td style="padding:10px 12px;font-weight:700;color:var(--white)">${this._escapeHtml(o.repName)}</td>` : ''}
        <td style="padding:10px 12px;color:var(--silver)">${this._escapeHtml(o.dsi) || '\u2014'}</td>
        <td style="padding:10px 12px;color:var(--silver)">${o.dateOfSale}</td>
        <td style="padding:10px 12px;color:var(--white)">${soldStr}</td>
        <td style="padding:10px 12px;text-align:center">
          <span style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:${statusColor};background:${statusColor}18;border:1px solid ${statusColor}44;border-radius:6px;padding:3px 10px">${this._escapeHtml(o.status)}</span>
        </td>
        <td style="padding:10px 12px;font-size:11px;color:var(--silver-dim);max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${this._escapeHtml(notePreview)}${noteCount > 1 ? ` <span style="color:var(--blue-core)">(${noteCount})</span>` : ''}</td>
        <td style="padding:10px 8px;text-align:right;white-space:nowrap">
          ${canEdit ? `<button onclick="Orders.openEditModal(${o.rowIndex})" style="background:none;border:1px solid rgba(26,92,229,0.3);border-radius:6px;padding:4px 10px;color:var(--blue-core);font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;cursor:pointer;margin-right:4px">Edit</button>` : ''}
          <button onclick="Orders.openNoteModal(${o.rowIndex},'${escapedDsi}')"
            style="background:none;border:1px solid rgba(26,92,229,0.3);border-radius:6px;padding:4px 10px;color:var(--blue-core);font-family:'Barlow Condensed',sans-serif;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;cursor:pointer">Notes${noteCount > 0 ? ' (' + noteCount + ')' : ''}</button>
        </td>`;
      tbody.appendChild(tr);
    });
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
          `<div style="padding:6px 0;border-bottom:1px solid rgba(26,92,229,0.1);font-size:12px;color:var(--silver);line-height:1.5">${this._escapeHtml(line)}</div>`
        ).join('');
      } else {
        notesDisplay.innerHTML = '<div style="color:var(--silver-dim);font-size:12px;padding:12px 0">No notes yet</div>';
      }
    }

    const dsiLabel = document.getElementById('order-note-dsi');
    if (dsiLabel) dsiLabel.textContent = dsi ? `DSI: ${dsi}` : 'Order Note';

    const input = document.getElementById('order-note-input');
    if (input) input.value = '';

    const errorEl = document.getElementById('order-note-error');
    if (errorEl) errorEl.textContent = '';
  },

  closeNoteModal() {
    const modal = document.getElementById('order-note-modal');
    if (modal) modal.style.display = 'none';
    this._activeRowIndex = null;
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
