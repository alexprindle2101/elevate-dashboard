// ═══════════════════════════════════════════════════════
// Aptel Admin Dashboard — Render Module
// ═══════════════════════════════════════════════════════

const AdminRender = {

  // ═══════════════════════════════════════════════════════
  // OFFICES PAGE
  // ═══════════════════════════════════════════════════════

  renderOffices(offices) {
    const grid = document.getElementById('offices-grid');
    if (!grid) return;

    let html = '';

    // Office cards
    offices.forEach(office => {
      const templateCfg = ADMIN_CONFIG.templates[office.templateType];
      const templateLabel = templateCfg ? templateCfg.label : office.templateType;
      const statusClass = office.status || 'setup';
      const statusLabel = statusClass.charAt(0).toUpperCase() + statusClass.slice(1);

      const logoSrc = office.logoIconUrl || office.logoUrl || 'references/logos/aptel-symbol-black.png';

      html += `
        <div class="office-card" onclick="AdminApp.openOffice('${office.officeId}')">
          <div class="office-card-header">
            <div class="office-card-identity">
              <img class="office-card-logo" src="${this._esc(logoSrc)}" alt="">
              <div class="office-card-name">${this._esc(office.name || 'Unnamed Office')}</div>
            </div>
            <span class="office-card-template">${this._esc(templateLabel)}</span>
          </div>
          <div class="office-card-details">
            <div class="office-card-detail">
              <span class="label">Status</span>
              <span class="status-badge ${statusClass}">
                <span class="dot"></span> ${statusLabel}
              </span>
            </div>
            <div class="office-card-detail">
              <span class="label">Owner</span>
              <span>${this._esc(office.ownerName || office.ownerEmail || '—')}</span>
            </div>
            <div class="office-card-detail">
              <span class="label">Sheet</span>
              <span>${office.sheetId ? this._truncate(office.sheetId, 20) : '—'}</span>
            </div>
          </div>
          <div class="office-card-actions" onclick="event.stopPropagation()">
            <button class="btn btn-secondary btn-sm" onclick="AdminApp.showEditOfficeModal('${office.officeId}')">Edit</button>
            <button class="btn btn-secondary btn-sm" onclick="AdminApp.openOffice('${office.officeId}')">Open Dashboard</button>
            <button class="btn btn-danger btn-sm" onclick="AdminApp.deleteOffice('${office.officeId}')">Delete</button>
          </div>
        </div>
      `;
    });

    // Add Office card (dashed)
    html += `
      <div class="add-card" onclick="AdminApp.showAddOfficeModal()">
        <span class="icon">+</span>
        <span class="label">Add Office</span>
      </div>
    `;

    grid.innerHTML = html;
  },


  // ═══════════════════════════════════════════════════════
  // PEOPLE PAGE
  // ═══════════════════════════════════════════════════════

  renderPeople(adminRoster) {
    const wrap = document.getElementById('people-table-wrap');
    if (!wrap) return;

    const admins = Object.values(adminRoster);

    if (admins.length === 0) {
      wrap.innerHTML = `
        <div class="empty-state">
          <div class="icon">👥</div>
          <h3>No Admins Yet</h3>
          <p>Add your first admin user to get started.</p>
        </div>
      `;
      return;
    }

    let rows = '';
    admins.forEach(admin => {
      const statusClass = admin.deactivated ? 'inactive' : 'active';
      const statusLabel = admin.deactivated ? 'Deactivated' : 'Active';
      const roleLabel = admin.role === 'superadmin' ? 'Super Admin' : admin.role;

      rows += `
        <tr>
          <td><strong>${this._esc(admin.name || '—')}</strong></td>
          <td>${this._esc(admin.email)}</td>
          <td>${this._esc(roleLabel)}</td>
          <td>
            <span class="status-badge ${statusClass}">
              <span class="dot"></span> ${statusLabel}
            </span>
          </td>
          <td>${admin.hasPinSet ? 'Yes' : 'No'}</td>
          <td>${this._formatDate(admin.dateAdded)}</td>
          <td>
            <button class="btn btn-secondary btn-sm" onclick="AdminApp.showEditAdminModal('${this._esc(admin.email)}')">Edit</button>
            <button class="btn btn-sm ${admin.deactivated ? 'btn-primary' : 'btn-danger'}"
                    onclick="AdminApp.toggleAdminDeactivated('${this._esc(admin.email)}')">
              ${admin.deactivated ? 'Reactivate' : 'Deactivate'}
            </button>
          </td>
        </tr>
      `;
    });

    wrap.innerHTML = `
      <table class="admin-table">
        <thead>
          <tr>
            <th>Name</th>
            <th>Email</th>
            <th>Role</th>
            <th>Status</th>
            <th>PIN Set</th>
            <th>Added</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    `;
  },


  // ═══════════════════════════════════════════════════════
  // SETTINGS PAGE
  // ═══════════════════════════════════════════════════════

  renderSettings(offices, adminRoster) {
    const officeCount = document.getElementById('settings-office-count');
    const adminCount = document.getElementById('settings-admin-count');

    const activeOffices = offices.filter(o => o.status === 'active').length;
    const totalOffices = offices.length;
    const totalAdmins = Object.keys(adminRoster).length;
    const activeAdmins = Object.values(adminRoster).filter(a => !a.deactivated).length;

    if (officeCount) {
      officeCount.innerHTML = `<strong>Active Offices:</strong> ${activeOffices} of ${totalOffices}`;
    }
    if (adminCount) {
      adminCount.innerHTML = `<strong>Admin Users:</strong> ${activeAdmins} active (${totalAdmins} total)`;
    }
  },


  // ═══════════════════════════════════════════════════════
  // OFFICE MODAL
  // ═══════════════════════════════════════════════════════

  populateOfficeModal(office) {
    const title = document.getElementById('office-modal-title');
    const nameInput = document.getElementById('office-name');
    const templateSelect = document.getElementById('office-template');
    const sheetIdInput = document.getElementById('office-sheet-id');
    const scriptUrlInput = document.getElementById('office-script-url');
    const apiKeyInput = document.getElementById('office-api-key');
    const ownerEmailInput = document.getElementById('office-owner-email');
    const ownerNameInput = document.getElementById('office-owner-name');
    const logoUrlInput = document.getElementById('office-logo-url');
    const logoIconUrlInput = document.getElementById('office-logo-icon-url');
    const statusSelect = document.getElementById('office-status');
    const error = document.getElementById('office-modal-error');

    // Set title
    if (title) title.textContent = office ? 'Edit Office' : 'Add Office';
    if (error) error.textContent = '';

    // Populate template dropdown
    if (templateSelect) {
      templateSelect.innerHTML = '';
      Object.entries(ADMIN_CONFIG.templates).forEach(([key, tmpl]) => {
        const opt = document.createElement('option');
        opt.value = key;
        opt.textContent = tmpl.label;
        if (office && office.templateType === key) opt.selected = true;
        templateSelect.appendChild(opt);
      });
    }

    // Fill fields
    if (nameInput) nameInput.value = office ? office.name : '';
    if (sheetIdInput) sheetIdInput.value = office ? office.sheetId : '';
    if (scriptUrlInput) scriptUrlInput.value = office ? office.appsScriptUrl : '';
    if (apiKeyInput) apiKeyInput.value = office ? office.apiKey : '';
    if (ownerEmailInput) ownerEmailInput.value = office ? office.ownerEmail : '';
    if (ownerNameInput) ownerNameInput.value = office ? office.ownerName : '';
    if (logoUrlInput) logoUrlInput.value = office ? office.logoUrl : '';
    if (logoIconUrlInput) logoIconUrlInput.value = office ? office.logoIconUrl : '';
    if (statusSelect) statusSelect.value = office ? office.status : 'setup';
  },


  // ═══════════════════════════════════════════════════════
  // ADMIN MODAL
  // ═══════════════════════════════════════════════════════

  populateAdminModal(admin) {
    const title = document.getElementById('admin-modal-title');
    const emailInput = document.getElementById('admin-email');
    const nameInput = document.getElementById('admin-name');
    const roleSelect = document.getElementById('admin-role');
    const error = document.getElementById('admin-modal-error');

    if (title) title.textContent = admin ? 'Edit Admin' : 'Add Admin';
    if (error) error.textContent = '';

    if (emailInput) {
      emailInput.value = admin ? admin.email : '';
      // Disable email editing on existing admins (email is the primary key)
      emailInput.disabled = !!admin;
    }
    if (nameInput) nameInput.value = admin ? admin.name : '';
    if (roleSelect) roleSelect.value = admin ? admin.role : 'superadmin';
  },


  // ═══════════════════════════════════════════════════════
  // HELPERS
  // ═══════════════════════════════════════════════════════

  _esc(str) {
    const el = document.createElement('span');
    el.textContent = str || '';
    return el.innerHTML;
  },

  _truncate(str, len) {
    if (!str) return '';
    const safe = this._esc(str);
    if (str.length <= len) return safe;
    return this._esc(str.slice(0, len)) + '&hellip;';
  },

  _formatDate(dateStr) {
    if (!dateStr) return '—';
    try {
      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return '—';
      return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    } catch (_) {
      return '—';
    }
  }
};
