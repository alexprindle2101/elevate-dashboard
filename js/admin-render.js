// ═══════════════════════════════════════════════════════
// Aptel Admin Dashboard — Render Module
// ═══════════════════════════════════════════════════════

const AdminRender = {

  // ═══════════════════════════════════════════════════════
  // OFFICES PAGE
  // ═══════════════════════════════════════════════════════

  renderOffices(offices, role, userType) {
    const grid = document.getElementById('offices-grid');
    if (!grid) return;

    const isA3 = role === 'a3' && userType !== 'owner';
    const isOwner = userType === 'owner';
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
            ${isOwner ? '' : `
            <div class="office-card-detail">
              <span class="label">Owner</span>
              <span>${this._esc(office.ownerName || office.ownerEmail || '—')}</span>
            </div>
            <div class="office-card-detail">
              <span class="label">Role</span>
              <span>${this._ownerLevelLabel(this._resolveOwnerLevel(office))}</span>
            </div>
            `}
          </div>
          <div class="office-card-actions" onclick="event.stopPropagation()">
            ${isA3 ? `<button class="btn btn-secondary btn-sm" onclick="AdminApp.showEditOfficeModal('${office.officeId}')">Edit</button>` : ''}
            <button class="btn btn-${isOwner ? 'primary' : 'secondary'} btn-sm" onclick="AdminApp.openOffice('${office.officeId}')">Open Dashboard</button>
            ${isA3 ? `<button class="btn btn-danger btn-sm" onclick="AdminApp.deleteOffice('${office.officeId}')">Delete</button>` : ''}
          </div>
        </div>
      `;
    });

    // Add Office card — a3 only (not owners)
    if (isA3) {
      html += `
        <div class="add-card" onclick="AdminApp.showAddOfficeModal()">
          <span class="icon">+</span>
          <span class="label">Add Office</span>
        </div>
      `;
    }

    grid.innerHTML = html;
  },


  // ═══════════════════════════════════════════════════════
  // PEOPLE PAGE
  // ═══════════════════════════════════════════════════════

  renderPeople(adminRoster, role, currentEmail) {
    const wrap = document.getElementById('people-table-wrap');
    if (!wrap) return;

    // Update page header — hide Add Admin button for a1
    const addBtn = document.querySelector('#page-people .page-header .btn-primary');
    if (addBtn) addBtn.style.display = (role === 'a1') ? 'none' : '';

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

    const isA1 = role === 'a1';
    const isA2 = role === 'a2';
    const isA3 = role === 'a3';

    let rows = '';
    admins.forEach(admin => {
      const statusClass = admin.deactivated ? 'inactive' : 'active';
      const statusLabel = admin.deactivated ? 'Deactivated' : 'Active';

      // Map role code to label
      const roleCfg = ADMIN_CONFIG.adminRoles[admin.role];
      const roleLabel = roleCfg ? roleCfg.label : (admin.role === 'superadmin' ? 'Super Admin' : admin.role);

      // Determine if current user can manage this admin
      const canManage = isA3 || (isA2 && admin.managedBy === currentEmail);

      // Scope info columns for a2/a3
      let scopeInfo = '—';
      if (admin.role === 'a1' && admin.assignedOffices) {
        scopeInfo = admin.assignedOffices.split(',').length + ' office(s)';
      } else if (admin.role === 'a2' && admin.assignedOwner) {
        const ownerData = AdminApp.state.owners[admin.assignedOwner];
        scopeInfo = ownerData ? ownerData.name : admin.assignedOwner;
      } else if (admin.role === 'a3') {
        scopeInfo = 'Full access';
      }

      rows += `
        <tr>
          <td><strong>${this._esc(admin.name || '—')}</strong></td>
          <td>${this._esc(admin.email)}</td>
          <td>${this._esc(roleLabel)}</td>
          <td>${this._esc(scopeInfo)}</td>
          <td>
            <span class="status-badge ${statusClass}">
              <span class="dot"></span> ${statusLabel}
            </span>
          </td>
          <td>${admin.hasPinSet ? 'Yes' : 'No'}</td>
          <td>${this._formatDate(admin.dateAdded)}</td>
          ${isA1 ? '' : `
          <td>
            ${canManage ? `
              <button class="btn btn-secondary btn-sm" onclick="AdminApp.showEditAdminModal('${this._esc(admin.email)}')">Edit</button>
              <button class="btn btn-sm ${admin.deactivated ? 'btn-primary' : 'btn-danger'}"
                      onclick="AdminApp.toggleAdminDeactivated('${this._esc(admin.email)}')">
                ${admin.deactivated ? 'Reactivate' : 'Deactivate'}
              </button>
            ` : ''}
          </td>
          `}
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
            <th>Scope</th>
            <th>Status</th>
            <th>PIN Set</th>
            <th>Added</th>
            ${isA1 ? '' : '<th>Actions</th>'}
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
    const ownerSelect = document.getElementById('office-owner-select');
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

      // Auto-fill campaign fields when template changes (new offices only)
      if (!office) {
        const newSelect = templateSelect.cloneNode(true);
        templateSelect.parentNode.replaceChild(newSelect, templateSelect);
        newSelect.addEventListener('change', () => {
          this._autoFillCampaignFields(newSelect.value);
        });
        // Trigger initial auto-fill for default template
        setTimeout(() => this._autoFillCampaignFields(newSelect.value), 0);
      }
    }

    // Fill fields
    if (nameInput) nameInput.value = office ? office.name : '';
    if (sheetIdInput) sheetIdInput.value = office ? office.sheetId : '';
    if (scriptUrlInput) scriptUrlInput.value = office ? office.appsScriptUrl : '';
    if (apiKeyInput) apiKeyInput.value = office ? office.apiKey : '';

    // Populate owner dropdown from _Owners tab data
    if (ownerSelect) {
      ownerSelect.innerHTML = '<option value="">(No owner assigned)</option>';
      const activeOwners = Object.values(AdminApp.state.owners).filter(o => !o.deactivated);
      activeOwners.sort((a, b) => a.name.localeCompare(b.name));
      activeOwners.forEach(owner => {
        const opt = document.createElement('option');
        opt.value = owner.email;
        opt.textContent = `${owner.name || owner.email} (${this._ownerLevelLabel(owner.level)})`;
        if (office && (office.ownerEmail || '').toLowerCase() === owner.email) opt.selected = true;
        ownerSelect.appendChild(opt);
      });
    }

    // Populate payroll manager dropdown with eligible admins
    const payrollSelect = document.getElementById('office-payroll-manager');
    if (payrollSelect) {
      payrollSelect.innerHTML = '<option value="">(None)</option>';
      const officeOwnerEmail = (office ? (office.ownerEmail || '') : (ownerSelect?.value || '')).toLowerCase();
      const officeId = office ? office.officeId : '';

      // Collect eligible admins: a3 (all offices), a2 (assigned to this owner), a1 (assigned to this office)
      const eligible = [];
      Object.values(AdminApp.state.adminRoster).forEach(admin => {
        if (admin.deactivated) return;
        if (admin.role === 'a3') { eligible.push(admin); return; }
        if (admin.role === 'a2' && officeOwnerEmail && admin.assignedOwner === officeOwnerEmail) { eligible.push(admin); return; }
        if (admin.role === 'a1' && officeId) {
          const ids = (admin.assignedOffices || '').split(',').map(s => s.trim().toLowerCase());
          if (ids.includes(officeId.toLowerCase())) eligible.push(admin);
        }
      });
      eligible.sort((a, b) => (a.name || a.email).localeCompare(b.name || b.email));
      eligible.forEach(admin => {
        const opt = document.createElement('option');
        opt.value = admin.email;
        opt.textContent = `${admin.name || admin.email} (${ADMIN_CONFIG.adminRoles[admin.role]?.label || admin.role})`;
        if (office && (office.payrollManagerEmail || '').toLowerCase() === admin.email) opt.selected = true;
        payrollSelect.appendChild(opt);
      });
    }

    if (logoUrlInput) logoUrlInput.value = office ? office.logoUrl : '';
    if (logoIconUrlInput) logoIconUrlInput.value = office ? office.logoIconUrl : '';
    const headerLogoStyleSelect = document.getElementById('office-header-logo-style');
    if (headerLogoStyleSelect) headerLogoStyleSelect.value = office ? (office.headerLogoStyle || 'icon') : 'icon';
    if (statusSelect) statusSelect.value = office ? office.status : 'setup';

    // Advanced Settings: auto-expand when editing (fields have values), collapse for new
    const advancedToggle = document.getElementById('advanced-toggle-btn');
    const advancedFields = document.getElementById('advanced-fields');
    if (advancedToggle && advancedFields) {
      const hasValues = office && (office.sheetId || office.appsScriptUrl || office.apiKey);
      if (hasValues) {
        advancedToggle.classList.add('open');
        advancedFields.classList.remove('collapsed');
      } else {
        advancedToggle.classList.remove('open');
        advancedFields.classList.add('collapsed');
      }
    }
  },


  // ═══════════════════════════════════════════════════════
  // ADMIN MODAL — expanded with role-based fields
  // ═══════════════════════════════════════════════════════

  populateAdminModal(admin, options) {
    options = options || { availableRoles: ['a3'], availableOwners: [], availableOffices: [] };

    const title = document.getElementById('admin-modal-title');
    const emailInput = document.getElementById('admin-email');
    const nameInput = document.getElementById('admin-name');
    const roleSelect = document.getElementById('admin-role');
    const ownerGroup = document.getElementById('admin-assigned-owner-group');
    const ownerSelect = document.getElementById('admin-assigned-owner');
    const officesGroup = document.getElementById('admin-assigned-offices-group');
    const officesList = document.getElementById('admin-assigned-offices-list');
    const error = document.getElementById('admin-modal-error');

    if (title) title.textContent = admin ? 'Edit Admin' : 'Add Admin';
    if (error) error.textContent = '';

    if (emailInput) {
      emailInput.value = admin ? admin.email : '';
      // Disable email editing on existing admins (email is the primary key)
      emailInput.disabled = !!admin;
    }
    if (nameInput) nameInput.value = admin ? admin.name : '';

    // Populate role dropdown with available roles
    if (roleSelect) {
      roleSelect.innerHTML = '';
      (options.availableRoles || []).forEach(roleKey => {
        const roleCfg = ADMIN_CONFIG.adminRoles[roleKey];
        const opt = document.createElement('option');
        opt.value = roleKey;
        opt.textContent = roleCfg ? roleCfg.label : roleKey;
        if (admin && admin.role === roleKey) opt.selected = true;
        roleSelect.appendChild(opt);
      });

      // Default to a1 for new admins
      if (!admin && roleSelect.options.length > 0) {
        roleSelect.value = 'a1';
      }

      // Wire up role change to toggle conditional fields
      const newSelect = roleSelect.cloneNode(true);
      roleSelect.parentNode.replaceChild(newSelect, roleSelect);
      newSelect.addEventListener('change', () => this._toggleAdminRoleFields(newSelect.value));

      // Set initial value for existing admin
      if (admin) newSelect.value = admin.role || 'a1';
    }

    // Populate assigned owner dropdown
    if (ownerSelect) {
      ownerSelect.innerHTML = '<option value="">(Select owner organization)</option>';
      (options.availableOwners || []).forEach(owner => {
        const opt = document.createElement('option');
        opt.value = owner.email;
        opt.textContent = `${owner.name || owner.email} (${this._ownerLevelLabel(owner.level)})`;
        if (admin && admin.assignedOwner === owner.email) opt.selected = true;
        ownerSelect.appendChild(opt);
      });
    }

    // Populate assigned offices checkbox list
    if (officesList) {
      officesList.innerHTML = '';
      const assignedSet = new Set((admin && admin.assignedOffices) ? admin.assignedOffices.split(',') : []);

      (options.availableOffices || []).forEach(office => {
        const checked = assignedSet.has(office.officeId) ? 'checked' : '';
        officesList.innerHTML += `
          <label style="display:flex;align-items:center;gap:8px;padding:6px 0;font-size:13px;cursor:pointer">
            <input type="checkbox" value="${this._esc(office.officeId)}" ${checked}
                   style="width:16px;height:16px;accent-color:var(--teal)">
            <span>${this._esc(office.name)}</span>
          </label>
        `;
      });

      if ((options.availableOffices || []).length === 0) {
        officesList.innerHTML = '<div style="color:var(--gray-400);font-size:13px;padding:8px 0">No active offices available</div>';
      }
    }

    // Toggle conditional fields based on current role
    const currentRole = admin ? admin.role : 'a1';
    this._toggleAdminRoleFields(currentRole);
  },

  _toggleAdminRoleFields(role) {
    const ownerGroup = document.getElementById('admin-assigned-owner-group');
    const officesGroup = document.getElementById('admin-assigned-offices-group');

    // Show assignedOwner for a2, assignedOffices for a1, hide both for a3
    if (ownerGroup) ownerGroup.style.display = (role === 'a2') ? 'block' : 'none';
    if (officesGroup) officesGroup.style.display = (role === 'a1') ? 'block' : 'none';
  },


  // ═══════════════════════════════════════════════════════
  // OWNERS PAGE
  // ═══════════════════════════════════════════════════════

  renderOwners(rootNodes, totalCount, role) {
    const wrap = document.getElementById('owners-tree-wrap');
    const countEl = document.getElementById('owners-count');
    if (!wrap) return;

    const isA3 = role === 'a3';

    // Update page header — hide Add Owner button for non-a3
    const addBtn = document.querySelector('#page-owners .page-header .btn-primary');
    if (addBtn) addBtn.style.display = isA3 ? '' : 'none';

    if (countEl) countEl.textContent = totalCount + ' owner' + (totalCount !== 1 ? 's' : '');

    if (totalCount === 0) {
      wrap.innerHTML = `
        <div class="empty-state">
          <div class="icon">👑</div>
          <h3>No Owners Yet</h3>
          <p>Add your first owner to start building the promotion hierarchy.</p>
        </div>
      `;
      return;
    }

    let rows = '';
    const renderNode = (node, depth) => {
      const indent = depth * 28;
      const connector = depth > 0 ? '<span style="color:var(--gray-400);margin-right:6px">└</span>' : '';
      const levelLabel = this._ownerLevelLabel(node.level);
      const statusClass = node.deactivated ? 'inactive' : 'active';
      const statusLabel = node.deactivated ? 'Deactivated' : 'Active';
      const officesBadge = node.officeCount > 0
        ? `<span style="background:var(--teal);color:#fff;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600">${node.officeCount}</span>`
        : '<span style="color:var(--gray-400)">0</span>';
      const downlineCount = this._countDescendants(node);
      const downlineBadge = downlineCount > 0
        ? `<span style="background:var(--blue-core);color:#fff;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600">${downlineCount}</span>`
        : '<span style="color:var(--gray-400)">0</span>';

      rows += `
        <tr>
          <td style="padding-left:${indent + 12}px">
            ${connector}<strong>${this._esc(node.name || '—')}</strong>
          </td>
          <td>${this._esc(node.email)}</td>
          <td>${this._esc(levelLabel)}</td>
          <td style="text-align:center">${officesBadge}</td>
          <td style="text-align:center">${downlineBadge}</td>
          <td>
            <span class="status-badge ${statusClass}">
              <span class="dot"></span> ${statusLabel}
            </span>
          </td>
          ${isA3 ? `
          <td>
            <button class="btn btn-secondary btn-sm" onclick="AdminApp.showEditOwnerModal('${this._esc(node.email)}')">Edit</button>
            <button class="btn btn-sm ${node.deactivated ? 'btn-primary' : 'btn-danger'}"
                    onclick="AdminApp.toggleOwnerDeactivated('${this._esc(node.email)}')">
              ${node.deactivated ? 'Reactivate' : 'Deactivate'}
            </button>
            <button class="btn btn-danger btn-sm" onclick="AdminApp.deleteOwner('${this._esc(node.email)}')">Delete</button>
          </td>
          ` : '<td></td>'}
        </tr>
      `;

      node.children.forEach(child => renderNode(child, depth + 1));
    };

    rootNodes.forEach(root => renderNode(root, 0));

    wrap.innerHTML = `
      <table class="admin-table">
        <thead>
          <tr>
            <th>Name</th>
            <th>Email</th>
            <th>Level</th>
            <th style="text-align:center">Offices</th>
            <th style="text-align:center">Downline</th>
            <th>Status</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    `;
  },


  // ═══════════════════════════════════════════════════════
  // OWNER MODAL
  // ═══════════════════════════════════════════════════════

  populateOwnerModal(owner, availableUplines) {
    const title = document.getElementById('owner-modal-title');
    const emailInput = document.getElementById('owner-email');
    const nameInput = document.getElementById('owner-name');
    const levelSelect = document.getElementById('owner-level');
    const uplineSelect = document.getElementById('owner-upline');
    const phoneInput = document.getElementById('owner-phone');
    const notesInput = document.getElementById('owner-notes');
    const error = document.getElementById('owner-modal-error');

    if (title) title.textContent = owner ? 'Edit Owner' : 'Add Owner';
    if (error) error.textContent = '';

    if (emailInput) {
      emailInput.value = owner ? owner.email : '';
      emailInput.disabled = !!owner; // email is primary key
    }
    if (nameInput) nameInput.value = owner ? owner.name : '';

    // Populate level dropdown
    if (levelSelect) {
      levelSelect.innerHTML = '';
      Object.entries(ADMIN_CONFIG.ownerLevels).forEach(([key, lvl]) => {
        const opt = document.createElement('option');
        opt.value = key;
        opt.textContent = lvl.label;
        if (owner && owner.level === key) opt.selected = true;
        levelSelect.appendChild(opt);
      });
    }

    // Populate upline dropdown (cycle-safe)
    if (uplineSelect) {
      uplineSelect.innerHTML = '<option value="">(None — top of hierarchy)</option>';
      (availableUplines || []).forEach(up => {
        const opt = document.createElement('option');
        opt.value = up.email;
        opt.textContent = `${up.name || up.email} (${this._ownerLevelLabel(up.level)})`;
        if (owner && owner.uplineEmail === up.email) opt.selected = true;
        uplineSelect.appendChild(opt);
      });
    }

    if (phoneInput) phoneInput.value = owner ? owner.phone : '';
    if (notesInput) notesInput.value = owner ? owner.notes : '';

    // Show PIN status for existing owners
    const pinStatusEl = document.getElementById('owner-pin-status');
    if (pinStatusEl) {
      if (owner) {
        const pinSet = owner.hasPinSet;
        pinStatusEl.innerHTML = `<span style="font-size:13px;color:${pinSet ? 'var(--green)' : 'var(--gray-400)'}">
          ${pinSet ? '&#10003; PIN set — can log in' : '&#9679; No PIN — hasn\'t logged in yet'}
        </span>`;
        pinStatusEl.style.display = 'block';
      } else {
        pinStatusEl.style.display = 'none';
      }
    }
  },


  // ═══════════════════════════════════════════════════════
  // HELPERS
  // ═══════════════════════════════════════════════════════

  // Auto-fill Sheet ID, Script URL, and API Key from campaign config
  _autoFillCampaignFields(templateType) {
    const campaignCfg = (ADMIN_CONFIG.campaign || {})[templateType];
    if (!campaignCfg) return;

    const sheetIdInput = document.getElementById('office-sheet-id');
    const scriptUrlInput = document.getElementById('office-script-url');
    const apiKeyInput = document.getElementById('office-api-key');

    // Only auto-fill if fields are empty (don't overwrite manual edits)
    if (sheetIdInput && !sheetIdInput.value) sheetIdInput.value = campaignCfg.sheetId || '';
    if (scriptUrlInput && !scriptUrlInput.value) scriptUrlInput.value = campaignCfg.appsScriptUrl || '';
    if (apiKeyInput && !apiKeyInput.value) apiKeyInput.value = campaignCfg.apiKey || '';
  },

  _countDescendants(node) {
    let count = node.children.length;
    node.children.forEach(child => { count += this._countDescendants(child); });
    return count;
  },

  // Look up the owner's real level from _Owners data (source of truth)
  // Falls back to the office row's ownerLevel if owner not found
  _resolveOwnerLevel(office) {
    const email = (office.ownerEmail || '').toLowerCase();
    if (email && AdminApp.state.owners[email]) {
      return AdminApp.state.owners[email].level || office.ownerLevel || 'o1';
    }
    return office.ownerLevel || 'o1';
  },

  _ownerLevelLabel(level) {
    const cfg = ADMIN_CONFIG.ownerLevels[level];
    return cfg ? cfg.label : 'Owner';
  },

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
