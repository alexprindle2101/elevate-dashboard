// ═══════════════════════════════════════════════════════
// Aptel Admin Dashboard — App Controller
// ═══════════════════════════════════════════════════════

const AdminApp = {
  state: {
    adminRoster: {},
    offices: [],
    currentPage: 'offices',
    currentEmail: '',
    currentName: '',
    currentRole: 'a3',
    userType: 'admin',  // 'admin' or 'owner'
    assignedOwner: '',
    assignedOffices: '',
    owners: {},
    editingOfficeId: null,  // null = add mode, string = edit mode
    editingAdminEmail: null,
    editingOwnerEmail: null
  },

  // ═══════════════════════════════════════════════════════
  // INIT
  // ═══════════════════════════════════════════════════════

  async init() {
    // Check existing session
    const session = this.getSession();
    if (session) {
      // Invalidate stale sessions that are missing the user's name
      // (caused by a previous login bug — force re-login to capture name)
      if (!session.name) {
        console.warn('[Admin] Session missing name — forcing re-login');
        localStorage.removeItem(ADMIN_CONFIG.sessionKey);
        this.showLoginScreen();
        return;
      }
      this.state.currentEmail = session.email;
      this.state.currentName = session.name;
      this.state.currentRole = session.role || 'a3';
      this.state.userType = session.userType || 'admin';
      await this.loadData();
      return;
    }

    // No session — show login
    this.showLoginScreen();
  },

  // ═══════════════════════════════════════════════════════
  // AUTH (replicates auth.js pattern for admin context)
  // ═══════════════════════════════════════════════════════

  getSession() {
    try {
      const raw = localStorage.getItem(ADMIN_CONFIG.sessionKey);
      if (!raw) return null;
      const session = JSON.parse(raw);
      if (Date.now() - session.loginTime > ADMIN_CONFIG.sessionDuration) {
        localStorage.removeItem(ADMIN_CONFIG.sessionKey);
        return null;
      }
      return session;
    } catch (_) { return null; }
  },

  saveSession(email, name, role, userType) {
    const session = { email, name, role, userType: userType || 'admin', loginTime: Date.now() };
    localStorage.setItem(ADMIN_CONFIG.sessionKey, JSON.stringify(session));
    return session;
  },

  showLoginScreen() {
    const screen = document.getElementById('login-screen');
    const dashboard = document.getElementById('admin-dashboard');
    if (screen) screen.style.display = 'flex';
    if (dashboard) dashboard.style.display = 'none';
    this._showEmailStep();
  },

  _showEmailStep() {
    const title = document.querySelector('.login-step-title');
    const subtitle = document.querySelector('.login-step-subtitle');
    if (title) title.textContent = 'Admin Sign In';
    if (subtitle) subtitle.textContent = 'Enter your admin email to continue';

    const emailWrap = document.getElementById('login-email-wrap');
    const pinWrap = document.getElementById('login-pin-wrap');
    const pinCreateWrap = document.getElementById('login-pin-create-wrap');
    const backLink = document.getElementById('login-back-link');
    if (emailWrap) emailWrap.style.display = 'block';
    if (pinWrap) pinWrap.style.display = 'none';
    if (pinCreateWrap) pinCreateWrap.style.display = 'none';
    if (backLink) backLink.style.display = 'none';

    const input = document.getElementById('login-email');
    const btn = document.getElementById('login-btn');
    const error = document.getElementById('login-error');
    if (input) input.value = '';
    if (error) error.textContent = '';
    if (btn) { btn.textContent = 'Continue'; btn.disabled = false; }

    const doSubmit = () => {
      const val = document.getElementById('login-email')?.value?.trim() || '';
      this.handleEmailStep(val);
    };

    if (btn) {
      const newBtn = btn.cloneNode(true);
      btn.parentNode.replaceChild(newBtn, btn);
      newBtn.addEventListener('click', doSubmit);
    }
    if (input) {
      const newInput = input.cloneNode(true);
      input.parentNode.replaceChild(newInput, input);
      newInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') doSubmit(); });
      setTimeout(() => newInput.focus(), 100);
    }
  },

  async handleEmailStep(email) {
    const error = document.getElementById('login-error');
    if (!email) {
      if (error) error.textContent = 'Please enter your email or alias';
      return;
    }

    // Resolve alias (e.g. "alex" → full email)
    const aliases = ADMIN_CONFIG.loginAliases || {};
    const resolved = aliases[email.toLowerCase()] || email;
    email = resolved;

    if (!email.includes('@')) {
      if (error) error.textContent = 'Unknown alias. Enter a full email or a configured alias.';
      return;
    }

    const btn = document.getElementById('login-btn');
    if (btn) { btn.textContent = 'Checking...'; btn.disabled = true; }
    if (error) error.textContent = '';

    try {
      // Validate against admin roster via Apps Script
      const resp = await this._post('validatePin', { email, pin: '' });

      if (!resp.success && resp.error === 'Email not found') {
        if (error) error.textContent = 'Email not found in admin roster';
        if (btn) { btn.textContent = 'Continue'; btn.disabled = false; }
        return;
      }

      this._loginEmail = email;
      this._loginUserType = resp.userType || 'admin';

      // Only set name/role if the response included them
      // (PIN error responses from validatePin don't include name)
      if (resp.name) this._loginName = resp.name;
      if (resp.role) this._loginRole = resp.role;

      if (resp.firstLogin) {
        this._showPinCreateStep();
      } else {
        this._showPinStep();
      }
    } catch (err) {
      if (error) error.textContent = 'Connection error. Check admin config.';
      if (btn) { btn.textContent = 'Continue'; btn.disabled = false; }
    }
  },

  _showPinStep() {
    const title = document.querySelector('.login-step-title');
    const subtitle = document.querySelector('.login-step-subtitle');
    if (title) title.textContent = 'Enter Your PIN';
    if (subtitle) subtitle.textContent = this._loginEmail;

    const emailWrap = document.getElementById('login-email-wrap');
    const pinWrap = document.getElementById('login-pin-wrap');
    const pinCreateWrap = document.getElementById('login-pin-create-wrap');
    const backLink = document.getElementById('login-back-link');
    if (emailWrap) emailWrap.style.display = 'none';
    if (pinWrap) pinWrap.style.display = 'block';
    if (pinCreateWrap) pinCreateWrap.style.display = 'none';
    if (backLink) backLink.style.display = 'inline';

    const pinInput = document.getElementById('login-pin');
    const btn = document.getElementById('login-btn');
    const error = document.getElementById('login-error');
    if (pinInput) pinInput.value = '';
    if (error) error.textContent = '';
    if (btn) { btn.textContent = 'Sign In'; btn.disabled = false; }

    const doSubmit = () => {
      const val = document.getElementById('login-pin')?.value || '';
      this.handlePinValidation(val);
    };

    if (btn) {
      const newBtn = btn.cloneNode(true);
      btn.parentNode.replaceChild(newBtn, btn);
      newBtn.addEventListener('click', doSubmit);
    }
    if (pinInput) {
      const newInput = pinInput.cloneNode(true);
      pinInput.parentNode.replaceChild(newInput, pinInput);
      newInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') doSubmit(); });
      setTimeout(() => newInput.focus(), 100);
    }
    if (backLink) {
      const newBack = backLink.cloneNode(true);
      backLink.parentNode.replaceChild(newBack, backLink);
      newBack.style.display = 'inline';
      newBack.addEventListener('click', () => this._showEmailStep());
    }
  },

  async handlePinValidation(pin) {
    const error = document.getElementById('login-error');
    const btn = document.getElementById('login-btn');
    if (!pin || pin.length < 4) {
      if (error) error.textContent = 'PIN must be 4-6 digits';
      return;
    }
    if (btn) { btn.textContent = 'Verifying...'; btn.disabled = true; }

    try {
      const resp = await this._post('validatePin', { email: this._loginEmail, pin });
      if (resp.success) {
        // Update name/role from successful PIN response (email step may not have included them)
        if (resp.name) this._loginName = resp.name;
        if (resp.role) this._loginRole = resp.role;
        if (resp.userType) this._loginUserType = resp.userType;
        this._completeLogin();
      } else {
        if (error) error.textContent = resp.error || 'Incorrect PIN';
        if (btn) { btn.textContent = 'Sign In'; btn.disabled = false; }
      }
    } catch (err) {
      if (error) error.textContent = 'Connection error';
      if (btn) { btn.textContent = 'Sign In'; btn.disabled = false; }
    }
  },

  _showPinCreateStep() {
    const title = document.querySelector('.login-step-title');
    const subtitle = document.querySelector('.login-step-subtitle');
    if (title) title.textContent = 'Create Your PIN';
    if (subtitle) subtitle.textContent = 'Set a 4-6 digit PIN for ' + this._loginEmail;

    const emailWrap = document.getElementById('login-email-wrap');
    const pinWrap = document.getElementById('login-pin-wrap');
    const pinCreateWrap = document.getElementById('login-pin-create-wrap');
    const backLink = document.getElementById('login-back-link');
    if (emailWrap) emailWrap.style.display = 'none';
    if (pinWrap) pinWrap.style.display = 'none';
    if (pinCreateWrap) pinCreateWrap.style.display = 'block';
    if (backLink) backLink.style.display = 'inline';

    const btn = document.getElementById('login-btn');
    const error = document.getElementById('login-error');
    if (error) error.textContent = '';
    if (btn) { btn.textContent = 'Set PIN & Sign In'; btn.disabled = false; }

    const doSubmit = () => {
      const pin1 = document.getElementById('login-pin-new')?.value || '';
      const pin2 = document.getElementById('login-pin-confirm')?.value || '';
      this.handlePinCreation(pin1, pin2);
    };

    if (btn) {
      const newBtn = btn.cloneNode(true);
      btn.parentNode.replaceChild(newBtn, btn);
      newBtn.addEventListener('click', doSubmit);
    }

    const pinNew = document.getElementById('login-pin-new');
    if (pinNew) {
      const newInput = pinNew.cloneNode(true);
      pinNew.parentNode.replaceChild(newInput, pinNew);
      setTimeout(() => newInput.focus(), 100);
    }

    if (backLink) {
      const newBack = backLink.cloneNode(true);
      backLink.parentNode.replaceChild(newBack, backLink);
      newBack.style.display = 'inline';
      newBack.addEventListener('click', () => this._showEmailStep());
    }
  },

  async handlePinCreation(pin, confirmPin) {
    const error = document.getElementById('login-error');
    const btn = document.getElementById('login-btn');

    if (!pin || pin.length < 4 || pin.length > 6) {
      if (error) error.textContent = 'PIN must be 4-6 digits';
      return;
    }
    if (pin !== confirmPin) {
      if (error) error.textContent = 'PINs do not match';
      return;
    }
    if (btn) { btn.textContent = 'Setting up...'; btn.disabled = true; }

    try {
      const resp = await this._post('createPin', { email: this._loginEmail, pin });
      if (resp.success) {
        this._completeLogin();
      } else {
        if (error) error.textContent = resp.error || 'Failed to set PIN';
        if (btn) { btn.textContent = 'Set PIN & Sign In'; btn.disabled = false; }
      }
    } catch (err) {
      if (error) error.textContent = 'Connection error';
      if (btn) { btn.textContent = 'Set PIN & Sign In'; btn.disabled = false; }
    }
  },

  async _completeLogin() {
    const userType = this._loginUserType || 'admin';
    this.saveSession(this._loginEmail, this._loginName, this._loginRole, userType);
    this.state.currentEmail = this._loginEmail;
    this.state.currentName = this._loginName;
    this.state.currentRole = this._loginRole || 'a3';
    this.state.userType = userType;

    const loginScreen = document.getElementById('login-screen');
    if (loginScreen) loginScreen.style.display = 'none';

    await this.loadData();
  },

  logout() {
    localStorage.removeItem(ADMIN_CONFIG.sessionKey);
    this.state.currentEmail = '';
    this.state.currentName = '';
    this.state.currentRole = 'a3';
    this.state.userType = 'admin';
    this.state.assignedOwner = '';
    this.state.assignedOffices = '';
    this.state.adminRoster = {};
    this.state.offices = [];
    this.state.owners = {};
    const dashboard = document.getElementById('admin-dashboard');
    if (dashboard) dashboard.style.display = 'none';
    this.showLoginScreen();
  },

  // ═══════════════════════════════════════════════════════
  // DATA
  // ═══════════════════════════════════════════════════════

  async loadData() {
    this.showLoading('Loading admin data...');

    try {
      // Use readScoped endpoint — returns role-filtered data based on email
      const url = `${ADMIN_CONFIG.appsScriptUrl}?key=${encodeURIComponent(ADMIN_CONFIG.apiKey)}&action=readScoped&email=${encodeURIComponent(this.state.currentEmail)}`;
      const resp = await fetch(url);
      const data = await resp.json();

      if (data.error) {
        console.error('Admin API error:', data.error);
        this.hideLoading();
        return;
      }

      this.state.adminRoster = data.adminRoster || {};
      this.state.offices = data.offices || [];
      this.state.owners = data.owners || {};

      // Update role + scope from server response (authoritative)
      if (data.role) {
        this.state.currentRole = data.role;
      }
      if (data.userType) {
        this.state.userType = data.userType;
      }
      // Update session with server-confirmed role + userType
      this.saveSession(this.state.currentEmail, this.state.currentName, this.state.currentRole, this.state.userType);
      if (data.assignedOwner !== undefined) this.state.assignedOwner = data.assignedOwner;
      if (data.assignedOffices !== undefined) this.state.assignedOffices = data.assignedOffices;

      this.hideLoading();
      this.showDashboard();
    } catch (err) {
      console.error('Failed to load admin data:', err);
      this.hideLoading();
      // Show dashboard anyway with empty state
      this.showDashboard();
    }
  },

  showLoading(msg) {
    const screen = document.getElementById('loading-screen');
    const text = document.getElementById('loading-text');
    if (screen) screen.style.display = 'flex';
    if (text) text.textContent = msg || 'Loading...';
  },

  hideLoading() {
    const screen = document.getElementById('loading-screen');
    if (screen) screen.style.display = 'none';
  },

  showDashboard() {
    const dashboard = document.getElementById('admin-dashboard');
    if (dashboard) dashboard.style.display = 'flex';

    // Update sidebar user info with role label
    const nameEl = document.getElementById('sidebar-user-name');
    const roleEl = document.getElementById('sidebar-user-role');
    if (nameEl) nameEl.textContent = this.state.currentName || this.state.currentEmail;

    if (this.state.userType === 'owner') {
      const ownerCfg = ADMIN_CONFIG.ownerLevels[this.state.currentRole];
      if (roleEl) roleEl.textContent = ownerCfg ? ownerCfg.label : 'Owner';
    } else {
      const roleCfg = ADMIN_CONFIG.adminRoles[this.state.currentRole];
      if (roleEl) roleEl.textContent = roleCfg ? roleCfg.label : 'Admin';
    }

    // Apply RBAC sidebar visibility
    this._applyRBACSidebar();

    // Render current page
    this.navTo(this.state.currentPage);
  },

  // ═══════════════════════════════════════════════════════
  // RBAC — Sidebar Visibility & Helpers
  // ═══════════════════════════════════════════════════════

  _applyRBACSidebar() {
    const role = this.state.currentRole;
    const userType = this.state.userType;
    const links = {
      offices: document.querySelector('.sidebar-link[data-page="offices"]'),
      owners: document.querySelector('.sidebar-link[data-page="owners"]'),
      people: document.querySelector('.sidebar-link[data-page="people"]'),
      settings: document.querySelector('.sidebar-link[data-page="settings"]')
    };

    if (userType === 'owner') {
      // Owners: Offices only
      if (links.owners) links.owners.style.display = 'none';
      if (links.people) links.people.style.display = 'none';
      if (links.settings) links.settings.style.display = 'none';
    } else {
      // a1: Offices + People only
      // a2: Offices + Owners + People
      // a3: All 4 tabs
      if (links.owners) links.owners.style.display = (role === 'a1') ? 'none' : '';
      if (links.people) links.people.style.display = '';
      if (links.settings) links.settings.style.display = (role === 'a3') ? '' : 'none';
    }
  },

  _roleRank() {
    const cfg = ADMIN_CONFIG.adminRoles[this.state.currentRole];
    return cfg ? cfg.rank : 0;
  },

  _canManageAdmin(email) {
    const role = this.state.currentRole;
    if (role === 'a3') return true;
    if (role === 'a2') {
      const admin = this.state.adminRoster[email];
      return admin && admin.managedBy === this.state.currentEmail;
    }
    return false;
  },

  _getAdminModalOptions() {
    const role = this.state.currentRole;

    // Available roles the current user can assign
    let availableRoles = [];
    if (role === 'a3') {
      availableRoles = ['a1', 'a2', 'a3'];
    } else if (role === 'a2') {
      availableRoles = ['a1', 'a2'];
    }

    // Available owners for assignedOwner dropdown (active only)
    const availableOwners = Object.values(this.state.owners).filter(o => !o.deactivated);

    // Available offices for assignedOffices checkboxes (active only)
    const availableOffices = this.state.offices.filter(o => o.status === 'active');

    return { availableRoles, availableOwners, availableOffices };
  },

  // ═══════════════════════════════════════════════════════
  // NAVIGATION
  // ═══════════════════════════════════════════════════════

  navTo(page) {
    const role = this.state.currentRole;
    const userType = this.state.userType;

    // RBAC guards — owners can only see offices
    if (userType === 'owner' && page !== 'offices') page = 'offices';

    // Admin RBAC guards
    if (userType === 'admin') {
      if (page === 'owners' && role === 'a1') page = 'offices';
      if (page === 'settings' && role !== 'a3') page = 'offices';
    }

    this.state.currentPage = page;

    // Update sidebar active state
    document.querySelectorAll('.sidebar-link').forEach(link => {
      link.classList.toggle('active', link.dataset.page === page);
    });

    // Hide all pages, show target
    ['offices', 'owners', 'people', 'settings'].forEach(p => {
      const el = document.getElementById('page-' + p);
      if (el) el.style.display = (p === page) ? 'block' : 'none';
    });

    // Render page content — pass role for conditional rendering
    switch (page) {
      case 'offices': {
        let visibleOffices = this.state.offices;
        if (userType === 'owner') {
          // Server already scopes offices for owners, but double-check client-side
          // (owner sees their own offices + downline offices if o2+)
        } else if (role === 'a1' && this.state.assignedOffices) {
          const allowed = new Set(this.state.assignedOffices.split(','));
          visibleOffices = this.state.offices.filter(o => allowed.has(o.officeId));
        }
        AdminRender.renderOffices(visibleOffices, role, userType);
        break;
      }
      case 'owners':
        AdminRender.renderOwners(this.buildOwnerTree(), Object.keys(this.state.owners).length, role);
        break;
      case 'people':
        AdminRender.renderPeople(this.state.adminRoster, role, this.state.currentEmail);
        break;
      case 'settings':
        AdminRender.renderSettings(this.state.offices, this.state.adminRoster);
        break;
    }
  },

  // ═══════════════════════════════════════════════════════
  // OFFICE CRUD — a3 only for add/edit/delete
  // ═══════════════════════════════════════════════════════

  showAddOfficeModal() {
    if (this.state.currentRole !== 'a3') return;
    this.state.editingOfficeId = null;
    AdminRender.populateOfficeModal(null);
    document.getElementById('office-modal')?.classList.add('open');
  },

  showEditOfficeModal(officeId) {
    if (this.state.currentRole !== 'a3') return;
    this.state.editingOfficeId = officeId;
    const office = this.state.offices.find(o => o.officeId === officeId);
    AdminRender.populateOfficeModal(office);
    document.getElementById('office-modal')?.classList.add('open');
  },

  closeOfficeModal() {
    document.getElementById('office-modal')?.classList.remove('open');
    const error = document.getElementById('office-modal-error');
    if (error) error.textContent = '';
  },

  async saveOffice() {
    if (this.state.currentRole !== 'a3') return;

    const name = document.getElementById('office-name')?.value?.trim();
    const templateType = document.getElementById('office-template')?.value;
    const sheetId = document.getElementById('office-sheet-id')?.value?.trim();
    const appsScriptUrl = document.getElementById('office-script-url')?.value?.trim();
    const apiKey = document.getElementById('office-api-key')?.value?.trim();

    // Owner comes from dropdown linked to _Owners tab
    const ownerSelect = document.getElementById('office-owner-select');
    const selectedOwnerEmail = ownerSelect?.value || '';
    let ownerEmail = selectedOwnerEmail;
    let ownerName = '';
    let ownerLevel = 'o1';
    if (selectedOwnerEmail && this.state.owners[selectedOwnerEmail]) {
      const ownerData = this.state.owners[selectedOwnerEmail];
      ownerName = ownerData.name;
      ownerLevel = ownerData.level;
    }

    const payrollManagerEmail = document.getElementById('office-payroll-manager')?.value || '';
    const payrollMode = document.getElementById('office-payroll-mode')?.value || 'commission-split';
    const logoUrl = document.getElementById('office-logo-url')?.value?.trim();
    const logoIconUrl = document.getElementById('office-logo-icon-url')?.value?.trim();
    const headerLogoStyle = document.getElementById('office-header-logo-style')?.value || 'icon';
    const status = document.getElementById('office-status')?.value;
    const error = document.getElementById('office-modal-error');

    if (!name) { if (error) error.textContent = 'Office name is required'; return; }

    const saveBtn = document.getElementById('office-modal-save');
    if (saveBtn) { saveBtn.textContent = 'Saving...'; saveBtn.disabled = true; }

    try {
      const payload = { name, templateType, sheetId, appsScriptUrl, apiKey, ownerEmail, ownerName, ownerLevel, payrollManagerEmail, payrollMode, logoUrl, logoIconUrl, headerLogoStyle, status };

      if (this.state.editingOfficeId) {
        payload.officeId = this.state.editingOfficeId;
        await this._post('updateOffice', payload);
      } else {
        // Create new office — AdminCode.gs returns the generated officeId
        const resp = await this._post('addOffice', payload);

        // Auto-create per-office tabs in the campaign sheet
        if (resp.officeId && templateType) {
          const campaignCfg = (ADMIN_CONFIG.campaign || {})[templateType];
          if (campaignCfg) {
            try {
              console.log('[Admin] Creating office tabs for', resp.officeId, 'in campaign sheet');
              await fetch(campaignCfg.appsScriptUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({
                  key: campaignCfg.apiKey,
                  action: 'createOfficeTabs',
                  officeId: resp.officeId,
                  sheetId: campaignCfg.sheetId
                })
              });
              console.log('[Admin] Office tabs created successfully');
            } catch (tabErr) {
              console.warn('[Admin] Tab creation failed (non-blocking):', tabErr.message);
            }
          }
        }
      }

      this.closeOfficeModal();
      await this.loadData();
    } catch (err) {
      if (error) error.textContent = 'Failed to save: ' + err.message;
    } finally {
      if (saveBtn) { saveBtn.textContent = 'Save Office'; saveBtn.disabled = false; }
    }
  },

  async deleteOffice(officeId) {
    if (this.state.currentRole !== 'a3') return;

    const office = this.state.offices.find(o => o.officeId === officeId);
    if (!office) return;
    if (!confirm(`Delete office "${office.name}"? This cannot be undone.`)) return;

    try {
      await this._post('deleteOffice', { officeId });
      await this.loadData();
    } catch (err) {
      alert('Failed to delete: ' + err.message);
    }
  },

  openOffice(officeId) {
    const office = this.state.offices.find(o => o.officeId === officeId);
    if (!office) return;

    const template = ADMIN_CONFIG.templates[office.templateType] || ADMIN_CONFIG.templates['att-b2b'];
    const config = {
      officeId: office.officeId,  // Per-office tab routing
      sheetId: office.sheetId,
      appsScriptUrl: office.appsScriptUrl,
      apiKey: office.apiKey,
      officeName: office.name,
      logoUrl: office.logoUrl || '',
      logoIconUrl: office.logoIconUrl || '',
      headerLogoStyle: office.headerLogoStyle || 'icon',
      payrollManagerEmail: office.payrollManagerEmail || '',
      payrollMode: office.payrollMode || 'commission-split'
    };

    const encoded = btoa(JSON.stringify(config));

    // Build SSO adminAuth token
    const adminAuth = {
      email: this.state.currentEmail,
      name: this.state.currentName,
      role: this.state.currentRole,
      userType: this.state.userType,
      source: 'admin-portal',
      timestamp: Date.now(),
      assignedOffices: this.state.assignedOffices,
      assignedOwner: this.state.assignedOwner
    };
    const authEncoded = btoa(JSON.stringify(adminAuth));

    const url = template.file + '?office=' + encoded + '&adminAuth=' + authEncoded;
    window.open(url, '_blank');
  },

  // ═══════════════════════════════════════════════════════
  // ADMIN USER CRUD — a2 (managed) and a3 (all)
  // ═══════════════════════════════════════════════════════

  showAddAdminModal() {
    if (this._roleRank() < 2) return; // a2 and a3 only
    this.state.editingAdminEmail = null;
    AdminRender.populateAdminModal(null, this._getAdminModalOptions());
    document.getElementById('admin-modal')?.classList.add('open');
  },

  showEditAdminModal(email) {
    if (!this._canManageAdmin(email)) return;
    this.state.editingAdminEmail = email;
    const admin = this.state.adminRoster[email];
    AdminRender.populateAdminModal(admin, this._getAdminModalOptions());
    document.getElementById('admin-modal')?.classList.add('open');
  },

  closeAdminModal() {
    document.getElementById('admin-modal')?.classList.remove('open');
    const error = document.getElementById('admin-modal-error');
    if (error) error.textContent = '';
  },

  async saveAdmin() {
    if (this._roleRank() < 2) return;

    const email = document.getElementById('admin-email')?.value?.trim()?.toLowerCase();
    const name = document.getElementById('admin-name')?.value?.trim();
    const role = document.getElementById('admin-role')?.value;
    const error = document.getElementById('admin-modal-error');

    if (!email || !email.includes('@')) { if (error) error.textContent = 'Valid email required'; return; }
    if (!name) { if (error) error.textContent = 'Name is required'; return; }

    // Collect new fields
    const assignedOwner = document.getElementById('admin-assigned-owner')?.value || '';

    // Collect assigned offices from checkboxes
    const officeCheckboxes = document.querySelectorAll('#admin-assigned-offices-list input[type="checkbox"]:checked');
    const assignedOffices = Array.from(officeCheckboxes).map(cb => cb.value).join(',');

    const saveBtn = document.getElementById('admin-modal-save');
    if (saveBtn) { saveBtn.textContent = 'Saving...'; saveBtn.disabled = true; }

    try {
      const payload = { email, name, role, assignedOwner, assignedOffices };

      if (this.state.editingAdminEmail) {
        await this._post('updateAdmin', payload);
      } else {
        // Auto-set managedBy to current user on create
        payload.managedBy = this.state.currentEmail;
        await this._post('addAdmin', payload);
      }

      this.closeAdminModal();
      await this.loadData();
    } catch (err) {
      if (error) error.textContent = 'Failed to save: ' + err.message;
    } finally {
      if (saveBtn) { saveBtn.textContent = 'Save'; saveBtn.disabled = false; }
    }
  },

  async toggleAdminDeactivated(email) {
    if (!this._canManageAdmin(email)) return;

    const admin = this.state.adminRoster[email];
    if (!admin) return;
    const newState = !admin.deactivated;
    if (newState && !confirm(`Deactivate ${admin.name}?`)) return;

    try {
      await this._post('updateAdmin', { email, deactivated: newState ? 'TRUE' : 'FALSE' });
      await this.loadData();
    } catch (err) {
      alert('Failed to update: ' + err.message);
    }
  },

  // ═══════════════════════════════════════════════════════
  // OWNER HIERARCHY
  // ═══════════════════════════════════════════════════════

  buildOwnerTree() {
    const owners = this.state.owners;
    const byEmail = {};

    // Clone each owner into a tree node with children array
    Object.values(owners).forEach(o => {
      byEmail[o.email] = { ...o, children: [], officeCount: 0 };
    });

    // Count offices per owner
    this.state.offices.forEach(office => {
      const email = (office.ownerEmail || '').toLowerCase();
      if (byEmail[email]) byEmail[email].officeCount++;
    });

    // Link parent-child via uplineEmail
    const roots = [];
    Object.values(byEmail).forEach(node => {
      if (node.uplineEmail && byEmail[node.uplineEmail]) {
        byEmail[node.uplineEmail].children.push(node);
      } else {
        roots.push(node);
      }
    });

    // Sort alphabetically at each level
    const sortByName = (a, b) => a.name.localeCompare(b.name);
    roots.sort(sortByName);
    Object.values(byEmail).forEach(node => node.children.sort(sortByName));

    return roots;
  },

  getAvailableUplines(excludeEmail) {
    const owners = this.state.owners;
    if (!excludeEmail) return Object.values(owners);

    // Walk downline recursively to find all descendants
    const descendants = new Set();
    const walk = (email) => {
      Object.values(owners).forEach(o => {
        if (o.uplineEmail === email && !descendants.has(o.email)) {
          descendants.add(o.email);
          walk(o.email);
        }
      });
    };
    descendants.add(excludeEmail);
    walk(excludeEmail);

    return Object.values(owners).filter(o => !descendants.has(o.email));
  },

  // ═══════════════════════════════════════════════════════
  // OWNER CRUD — a3 only
  // ═══════════════════════════════════════════════════════

  showAddOwnerModal() {
    if (this.state.currentRole !== 'a3') return;
    this.state.editingOwnerEmail = null;
    AdminRender.populateOwnerModal(null, this.getAvailableUplines(null));
    document.getElementById('owner-modal')?.classList.add('open');
  },

  showEditOwnerModal(email) {
    if (this.state.currentRole !== 'a3') return;
    this.state.editingOwnerEmail = email;
    const owner = this.state.owners[email];
    AdminRender.populateOwnerModal(owner, this.getAvailableUplines(email));
    document.getElementById('owner-modal')?.classList.add('open');
  },

  closeOwnerModal() {
    document.getElementById('owner-modal')?.classList.remove('open');
    const error = document.getElementById('owner-modal-error');
    if (error) error.textContent = '';
  },

  async saveOwner() {
    if (this.state.currentRole !== 'a3') return;

    const email = document.getElementById('owner-email')?.value?.trim()?.toLowerCase();
    const name = document.getElementById('owner-name')?.value?.trim();
    const level = document.getElementById('owner-level')?.value;
    const uplineEmail = document.getElementById('owner-upline')?.value || '';
    const phone = document.getElementById('owner-phone')?.value?.trim() || '';
    const notes = document.getElementById('owner-notes')?.value?.trim() || '';
    const error = document.getElementById('owner-modal-error');

    if (!email || !email.includes('@')) { if (error) error.textContent = 'Valid email required'; return; }
    if (!name) { if (error) error.textContent = 'Name is required'; return; }

    const saveBtn = document.getElementById('owner-modal-save');
    if (saveBtn) { saveBtn.textContent = 'Saving...'; saveBtn.disabled = true; }

    try {
      if (this.state.editingOwnerEmail) {
        await this._post('updateOwner', { email, name, level, uplineEmail, phone, notes });
      } else {
        await this._post('addOwner', { email, name, level, uplineEmail, phone, notes });
      }
      this.closeOwnerModal();
      await this.loadData();
    } catch (err) {
      if (error) error.textContent = 'Failed to save: ' + err.message;
    } finally {
      if (saveBtn) { saveBtn.textContent = 'Save'; saveBtn.disabled = false; }
    }
  },

  async deleteOwner(email) {
    if (this.state.currentRole !== 'a3') return;

    const owner = this.state.owners[email];
    if (!owner) return;

    const linkedOffices = this.state.offices.filter(o =>
      (o.ownerEmail || '').toLowerCase() === email.toLowerCase()
    );
    const downline = Object.values(this.state.owners).filter(o => o.uplineEmail === email);

    let msg = `Delete owner "${owner.name}"?`;
    if (linkedOffices.length > 0) msg += ` Warning: ${linkedOffices.length} office(s) reference this owner.`;
    if (downline.length > 0) msg += ` ${downline.length} downline owner(s) will become top-level.`;
    if (!confirm(msg)) return;

    try {
      await this._post('deleteOwner', { email });
      await this.loadData();
    } catch (err) {
      alert('Failed to delete: ' + err.message);
    }
  },

  async toggleOwnerDeactivated(email) {
    if (this.state.currentRole !== 'a3') return;

    const owner = this.state.owners[email];
    if (!owner) return;
    const newState = !owner.deactivated;
    if (newState && !confirm(`Deactivate ${owner.name}?`)) return;

    try {
      await this._post('updateOwner', { email, deactivated: newState ? 'TRUE' : 'FALSE' });
      await this.loadData();
    } catch (err) {
      alert('Failed to update: ' + err.message);
    }
  },

  // ═══════════════════════════════════════════════════════
  // API HELPER
  // ═══════════════════════════════════════════════════════

  async _post(action, payload) {
    const body = JSON.stringify({
      key: ADMIN_CONFIG.apiKey,
      action,
      ...payload
    });
    console.log('[Admin API] POST', action, '→', ADMIN_CONFIG.appsScriptUrl);

    const resp = await fetch(ADMIN_CONFIG.appsScriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      redirect: 'follow',
      body
    });

    const text = await resp.text();
    console.log('[Admin API] Response status:', resp.status, 'body:', text.substring(0, 200));

    try {
      return JSON.parse(text);
    } catch (e) {
      console.error('[Admin API] Failed to parse JSON:', text.substring(0, 500));
      throw new Error('Invalid response from server');
    }
  }
};

// ═══════════════════════════════════════════════════════
// BOOT
// ═══════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => AdminApp.init());
