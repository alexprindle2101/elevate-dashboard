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
      this.state.currentEmail = session.email;
      this.state.currentName = session.name;
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

  saveSession(email, name, role) {
    const session = { email, name, role, loginTime: Date.now() };
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

      if (resp.firstLogin) {
        this._loginName = resp.name;
        this._loginRole = resp.role;
        this._showPinCreateStep();
      } else {
        this._loginName = resp.name;
        this._loginRole = resp.role;
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
    this.saveSession(this._loginEmail, this._loginName, this._loginRole);
    this.state.currentEmail = this._loginEmail;
    this.state.currentName = this._loginName;

    const loginScreen = document.getElementById('login-screen');
    if (loginScreen) loginScreen.style.display = 'none';

    await this.loadData();
  },

  logout() {
    localStorage.removeItem(ADMIN_CONFIG.sessionKey);
    this.state.currentEmail = '';
    this.state.currentName = '';
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
      const url = `${ADMIN_CONFIG.appsScriptUrl}?key=${encodeURIComponent(ADMIN_CONFIG.apiKey)}`;
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

    // Update sidebar user info
    const nameEl = document.getElementById('sidebar-user-name');
    const roleEl = document.getElementById('sidebar-user-role');
    if (nameEl) nameEl.textContent = this.state.currentName || this.state.currentEmail;
    if (roleEl) roleEl.textContent = 'Super Admin';

    // Render current page
    this.navTo(this.state.currentPage);
  },

  // ═══════════════════════════════════════════════════════
  // NAVIGATION
  // ═══════════════════════════════════════════════════════

  navTo(page) {
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

    // Render page content
    switch (page) {
      case 'offices':
        AdminRender.renderOffices(this.state.offices);
        break;
      case 'owners':
        AdminRender.renderOwners(this.buildOwnerTree(), Object.keys(this.state.owners).length);
        break;
      case 'people':
        AdminRender.renderPeople(this.state.adminRoster);
        break;
      case 'settings':
        AdminRender.renderSettings(this.state.offices, this.state.adminRoster);
        break;
    }
  },

  // ═══════════════════════════════════════════════════════
  // OFFICE CRUD
  // ═══════════════════════════════════════════════════════

  showAddOfficeModal() {
    this.state.editingOfficeId = null;
    AdminRender.populateOfficeModal(null);
    document.getElementById('office-modal')?.classList.add('open');
  },

  showEditOfficeModal(officeId) {
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
    const name = document.getElementById('office-name')?.value?.trim();
    const templateType = document.getElementById('office-template')?.value;
    const sheetId = document.getElementById('office-sheet-id')?.value?.trim();
    const appsScriptUrl = document.getElementById('office-script-url')?.value?.trim();
    const apiKey = document.getElementById('office-api-key')?.value?.trim();
    const ownerEmail = document.getElementById('office-owner-email')?.value?.trim();
    const ownerName = document.getElementById('office-owner-name')?.value?.trim();
    const ownerLevel = document.getElementById('office-owner-level')?.value;
    const logoUrl = document.getElementById('office-logo-url')?.value?.trim();
    const logoIconUrl = document.getElementById('office-logo-icon-url')?.value?.trim();
    const status = document.getElementById('office-status')?.value;
    const error = document.getElementById('office-modal-error');

    if (!name) { if (error) error.textContent = 'Office name is required'; return; }

    const saveBtn = document.getElementById('office-modal-save');
    if (saveBtn) { saveBtn.textContent = 'Saving...'; saveBtn.disabled = true; }

    try {
      const payload = { name, templateType, sheetId, appsScriptUrl, apiKey, ownerEmail, ownerName, ownerLevel, logoUrl, logoIconUrl, status };

      if (this.state.editingOfficeId) {
        payload.officeId = this.state.editingOfficeId;
        await this._post('updateOffice', payload);
      } else {
        await this._post('addOffice', payload);
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
      sheetId: office.sheetId,
      appsScriptUrl: office.appsScriptUrl,
      apiKey: office.apiKey,
      officeName: office.name,
      logoUrl: office.logoUrl || '',
      logoIconUrl: office.logoIconUrl || ''
    };

    const encoded = btoa(JSON.stringify(config));
    const url = template.file + '?office=' + encoded;
    window.open(url, '_blank');
  },

  // ═══════════════════════════════════════════════════════
  // ADMIN USER CRUD
  // ═══════════════════════════════════════════════════════

  showAddAdminModal() {
    this.state.editingAdminEmail = null;
    AdminRender.populateAdminModal(null);
    document.getElementById('admin-modal')?.classList.add('open');
  },

  showEditAdminModal(email) {
    this.state.editingAdminEmail = email;
    const admin = this.state.adminRoster[email];
    AdminRender.populateAdminModal(admin);
    document.getElementById('admin-modal')?.classList.add('open');
  },

  closeAdminModal() {
    document.getElementById('admin-modal')?.classList.remove('open');
    const error = document.getElementById('admin-modal-error');
    if (error) error.textContent = '';
  },

  async saveAdmin() {
    const email = document.getElementById('admin-email')?.value?.trim()?.toLowerCase();
    const name = document.getElementById('admin-name')?.value?.trim();
    const role = document.getElementById('admin-role')?.value;
    const error = document.getElementById('admin-modal-error');

    if (!email || !email.includes('@')) { if (error) error.textContent = 'Valid email required'; return; }
    if (!name) { if (error) error.textContent = 'Name is required'; return; }

    const saveBtn = document.getElementById('admin-modal-save');
    if (saveBtn) { saveBtn.textContent = 'Saving...'; saveBtn.disabled = true; }

    try {
      if (this.state.editingAdminEmail) {
        await this._post('updateAdmin', { email, name, role });
      } else {
        await this._post('addAdmin', { email, name, role });
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
  // OWNER CRUD
  // ═══════════════════════════════════════════════════════

  showAddOwnerModal() {
    this.state.editingOwnerEmail = null;
    AdminRender.populateOwnerModal(null, this.getAvailableUplines(null));
    document.getElementById('owner-modal')?.classList.add('open');
  },

  showEditOwnerModal(email) {
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
