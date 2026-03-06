// ═══════════════════════════════════════════════════════
// ELEVATE — Email + PIN Authentication System
// Two-step login: email → PIN (server-side validated)
// ═══════════════════════════════════════════════════════

const Auth = {
  SESSION_KEY: 'elevate_session',

  // ══════════════════════════════════════════════
  // SESSION MANAGEMENT
  // ══════════════════════════════════════════════

  getSession() {
    try {
      const raw = localStorage.getItem(this.SESSION_KEY);
      if (!raw) return null;
      const session = JSON.parse(raw);
      if (Date.now() - session.loginTime > OFFICE_CONFIG.sessionDuration) {
        this.clearSession();
        return null;
      }
      return session;
    } catch {
      this.clearSession();
      return null;
    }
  },

  saveSession(session) {
    localStorage.setItem(this.SESSION_KEY, JSON.stringify(session));
  },

  clearSession() {
    localStorage.removeItem(this.SESSION_KEY);
  },

  // ══════════════════════════════════════════════
  // LOGIN LOGIC
  // ══════════════════════════════════════════════

  // Step 1: Check email against roster (client-side only)
  checkEmail(email, rosterMap) {
    const cleanEmail = String(email).trim().toLowerCase();
    if (!cleanEmail) return { ok: false, error: 'Please enter your email' };

    const match = rosterMap[cleanEmail];
    if (!match) {
      console.warn('Login failed — email not found in roster:', cleanEmail);
      console.warn('Roster contains', Object.keys(rosterMap).length, 'entries:', Object.keys(rosterMap));
      return { ok: false, error: 'Email not found. Contact your JD or Admin to be added.' };
    }
    if (match.deactivated) {
      return { ok: false, error: 'Account deactivated. Contact your Admin.' };
    }

    return { ok: true, email: cleanEmail, hasPin: match.hasPin || false, rosterEntry: match };
  },

  // Step 2a: Validate existing PIN (server-side)
  async validatePin(email, pin, config) {
    try {
      const resp = await SheetsAPI.post(config, 'validatePin', { email, pin });
      if (!resp.ok) return { ok: false, error: 'Connection error' };
      const data = resp.data;
      if (data.valid === true) return { ok: true };
      return { ok: false, error: data.error || 'Incorrect PIN' };
    } catch (err) {
      return { ok: false, error: 'Connection error: ' + err.message };
    }
  },

  // Step 2b: Create PIN on first login (server-side)
  async createPin(email, pin, config) {
    try {
      const resp = await SheetsAPI.post(config, 'setPin', { email, pin });
      if (!resp.ok) return { ok: false, error: 'Connection error' };
      const data = resp.data;
      if (data.error) return { ok: false, error: data.error };
      return { ok: true };
    } catch (err) {
      return { ok: false, error: 'Connection error: ' + err.message };
    }
  },

  // Create session after PIN validation succeeds
  createSession(email, rosterEntry) {
    const session = {
      email: email,
      name: rosterEntry.name,
      role: rosterEntry.rank || 'rep',
      team: rosterEntry.team,
      office: OFFICE_CONFIG.officeName,
      loginTime: Date.now()
    };
    this.saveSession(session);
    return session;
  },

  // Create session from admin portal SSO token (bypasses PIN login)
  createAdminSSOSession(adminAuth) {
    // Map admin portal roles to office dashboard roles
    const roleMap = { a1: 'admin', a2: 'admin', a3: 'superadmin' };
    const officeRole = roleMap[adminAuth.role] || 'admin';

    const session = {
      email: adminAuth.email,
      name: adminAuth.name,
      role: officeRole,
      team: '',
      office: OFFICE_CONFIG.officeName,
      loginTime: Date.now(),
      source: 'admin-portal',
      adminRole: adminAuth.role,
      assignedOffices: adminAuth.assignedOffices || '',
      assignedOwner: adminAuth.assignedOwner || ''
    };
    this.saveSession(session);
    return session;
  },

  logout() {
    this.clearSession();
    window.location.reload();
  },

  // ══════════════════════════════════════════════
  // LOGIN UI — Step Management
  // ══════════════════════════════════════════════

  // Show login screen starting at email step
  showLoginScreen(onEmailSubmit) {
    const screen = document.getElementById('login-screen');
    if (!screen) return;
    screen.style.display = 'flex';
    this._showEmailStep(onEmailSubmit);
  },

  // Step 1: Email entry
  _showEmailStep(onEmailSubmit) {
    const title = document.querySelector('.login-step-title');
    const subtitle = document.querySelector('.login-step-subtitle');
    if (title) title.textContent = 'Sign In';
    if (subtitle) subtitle.textContent = 'Enter your email to access the dashboard';

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
      const currentInput = document.getElementById('login-email');
      if (onEmailSubmit) onEmailSubmit(currentInput?.value || '');
    };

    if (btn) {
      const newBtn = btn.cloneNode(true);
      btn.parentNode.replaceChild(newBtn, btn);
      newBtn.addEventListener('click', doSubmit);
    }

    if (input) {
      const newInput = input.cloneNode(true);
      input.parentNode.replaceChild(newInput, input);
      newInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') doSubmit();
      });
      setTimeout(() => newInput.focus(), 100);
    }

  },

  // Step 2a: Enter existing PIN
  showPinStep(email, onPinSubmit, onBack) {
    const title = document.querySelector('.login-step-title');
    const subtitle = document.querySelector('.login-step-subtitle');
    if (title) title.textContent = 'Enter Your PIN';
    if (subtitle) subtitle.textContent = email;

    const emailWrap = document.getElementById('login-email-wrap');
    const pinWrap = document.getElementById('login-pin-wrap');
    const pinCreateWrap = document.getElementById('login-pin-create-wrap');
    if (emailWrap) emailWrap.style.display = 'none';
    if (pinWrap) pinWrap.style.display = 'block';
    if (pinCreateWrap) pinCreateWrap.style.display = 'none';

    const pinInput = document.getElementById('login-pin');
    const btn = document.getElementById('login-btn');
    const error = document.getElementById('login-error');
    const backLink = document.getElementById('login-back-link');

    if (pinInput) pinInput.value = '';
    if (error) error.textContent = '';
    if (btn) { btn.textContent = 'Sign In'; btn.disabled = false; }
    if (backLink) backLink.style.display = 'inline';

    const doSubmit = () => {
      const val = document.getElementById('login-pin')?.value || '';
      if (onPinSubmit) onPinSubmit(val);
    };

    if (btn) {
      const newBtn = btn.cloneNode(true);
      btn.parentNode.replaceChild(newBtn, btn);
      newBtn.addEventListener('click', doSubmit);
    }

    if (pinInput) {
      const newInput = pinInput.cloneNode(true);
      pinInput.parentNode.replaceChild(newInput, pinInput);
      newInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') doSubmit();
      });
      setTimeout(() => newInput.focus(), 100);
    }

    if (backLink) {
      const newBack = backLink.cloneNode(true);
      backLink.parentNode.replaceChild(newBack, backLink);
      newBack.style.display = 'inline';
      newBack.addEventListener('click', () => { if (onBack) onBack(); });
    }
  },

  // Step 2b: Create PIN (first login)
  showPinCreateStep(email, onPinCreate, onBack) {
    const title = document.querySelector('.login-step-title');
    const subtitle = document.querySelector('.login-step-subtitle');
    if (title) title.textContent = 'Create Your PIN';
    if (subtitle) subtitle.textContent = 'Set a 4-6 digit PIN for ' + email;

    const emailWrap = document.getElementById('login-email-wrap');
    const pinWrap = document.getElementById('login-pin-wrap');
    const pinCreateWrap = document.getElementById('login-pin-create-wrap');
    if (emailWrap) emailWrap.style.display = 'none';
    if (pinWrap) pinWrap.style.display = 'none';
    if (pinCreateWrap) pinCreateWrap.style.display = 'block';

    const pinNew = document.getElementById('login-pin-new');
    const pinConfirm = document.getElementById('login-pin-confirm');
    const btn = document.getElementById('login-btn');
    const error = document.getElementById('login-error');
    const backLink = document.getElementById('login-back-link');

    if (pinNew) pinNew.value = '';
    if (pinConfirm) pinConfirm.value = '';
    if (error) error.textContent = '';
    if (btn) { btn.textContent = 'Set PIN & Sign In'; btn.disabled = false; }
    if (backLink) backLink.style.display = 'inline';

    const doSubmit = () => {
      const newVal = document.getElementById('login-pin-new')?.value || '';
      const confirmVal = document.getElementById('login-pin-confirm')?.value || '';
      if (onPinCreate) onPinCreate(newVal, confirmVal);
    };

    if (btn) {
      const newBtn = btn.cloneNode(true);
      btn.parentNode.replaceChild(newBtn, btn);
      newBtn.addEventListener('click', doSubmit);
    }

    if (pinConfirm) {
      const newInput = pinConfirm.cloneNode(true);
      pinConfirm.parentNode.replaceChild(newInput, pinConfirm);
      newInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') doSubmit();
      });
    }

    if (pinNew) {
      const newInput = pinNew.cloneNode(true);
      pinNew.parentNode.replaceChild(newInput, pinNew);
      setTimeout(() => newInput.focus(), 100);
    }

    if (backLink) {
      const newBack = backLink.cloneNode(true);
      backLink.parentNode.replaceChild(newBack, backLink);
      newBack.style.display = 'inline';
      newBack.addEventListener('click', () => { if (onBack) onBack(); });
    }
  },

  hideLoginScreen() {
    const screen = document.getElementById('login-screen');
    if (screen) screen.style.display = 'none';
  },

  showLoginError(msg) {
    const error = document.getElementById('login-error');
    if (error) error.textContent = msg;
  },

  // ══════════════════════════════════════════════
  // LOADING SCREEN
  // ══════════════════════════════════════════════

  showLoading(msg) {
    const screen = document.getElementById('loading-screen');
    const text = document.getElementById('loading-text');
    if (screen) screen.style.display = 'flex';
    if (text && msg) text.textContent = msg;
  },

  hideLoading() {
    const screen = document.getElementById('loading-screen');
    if (screen) screen.style.display = 'none';
  }
};
