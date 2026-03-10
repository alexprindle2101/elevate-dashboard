// ═══════════════════════════════════════════════════════
// NDS POST SALE — Multi-step wizard form
// ═══════════════════════════════════════════════════════
// Simplified AT&T NDS version: Air + Wireless only.
// No Fiber, VoIP, DTV, or Ooma campaign support.

const PostSale = {
  // ── State ──
  _step: 1,
  _campaign: 'attb2b',   // always attb2b for NDS
  _submitting: false,
  _formData: {},
  _products: {
    air:      { on: false, icon: '📡', label: 'Internet Air' },
    wireless: { on: false, icon: '📱', label: 'Wireless' }
  },

  // ── Open / Reset ──
  open() {
    const page = document.getElementById('post-sale-page');
    if (page) page.style.display = 'block';
    this._resetForm();
    this._render();
  },

  _resetForm() {
    this._step = 1;
    this._campaign = 'attb2b';
    this._submitting = false;
    this._formData = {
      dateOfSale: (() => { const n = new Date(); return n.getFullYear() + '-' + String(n.getMonth()+1).padStart(2,'0') + '-' + String(n.getDate()).padStart(2,'0'); })(),
      dsi: '',
      accountType: 'Business',
      accountNotes: '',
      trainee: false,
      traineeName: '',
      orderChannel: 'Sara',
      codesUsed: false,
      codesUsedBy: '',
      codesUsedByName: '',
      activationSupport: false,
      newPhones: 0,
      byods: 0,
      hashtags: ''
    };
    Object.keys(this._products).forEach(k => this._products[k].on = false);
    this._clearDraft();
  },

  _resetForAnother() {
    // Keep date, clear products and sale-specific fields
    const keepDate = this._formData.dateOfSale;
    this._resetForm();
    this._formData.dateOfSale = keepDate;
    this._step = 1;
    this._render();
  },

  // ── Draft Persistence ──
  _saveDraft() {
    try {
      const draft = { campaign: this._campaign, formData: this._formData, products: {} };
      Object.keys(this._products).forEach(k => draft.products[k] = this._products[k].on);
      sessionStorage.setItem('postSaleDraft', JSON.stringify(draft));
    } catch (e) { /* ignore */ }
  },

  _restoreDraft() {
    try {
      const raw = sessionStorage.getItem('postSaleDraft');
      if (!raw) { this._resetForm(); return; }
      const draft = JSON.parse(raw);
      this._campaign = 'attb2b';
      this._formData = { ...this._formData, ...draft.formData };
      if (draft.products) {
        Object.keys(draft.products).forEach(k => {
          if (this._products[k]) this._products[k].on = draft.products[k];
        });
      }
      this._step = 1;
    } catch (e) { this._resetForm(); }
  },

  _clearDraft() {
    try { sessionStorage.removeItem('postSaleDraft'); } catch (e) { /* ignore */ }
  },

  // ── Collect values from live DOM ──
  _collectStep1() {
    const v = (id) => { const el = document.getElementById(id); return el ? el.value : ''; };
    this._formData.dateOfSale = v('ps-date') || this._formData.dateOfSale;
    this._formData.dsi = v('ps-dsi');
    this._formData.accountNotes = v('ps-notes');
    this._formData.traineeName = v('ps-trainee-name');
    // Codes-used-by dropdown
    const codesSelect = document.getElementById('ps-codes-used-by');
    if (codesSelect && this._formData.codesUsed) {
      this._formData.codesUsedBy = codesSelect.value;
      const opt = codesSelect.options[codesSelect.selectedIndex];
      this._formData.codesUsedByName = opt && opt.value ? opt.textContent : '';
    }
    this._formData.hashtags = v('ps-hashtags');
    this._saveDraft();
  },

  _collectStep2() {
    const n = (id) => { const el = document.getElementById(id); return el ? (parseInt(el.value) || 0) : 0; };
    this._formData.newPhones = n('ps-new-phones');
    this._formData.byods = n('ps-byods');
    this._saveDraft();
  },

  // ── Master Render ──
  _render() {
    const session = Auth.getSession();
    const sub = document.getElementById('post-sale-subtitle');
    if (sub) sub.textContent = 'Logging sale for ' + (session?.name || 'you');
    this._renderSteps();
    this._renderBody();
    this._renderNav();
  },

  // ── Step Indicator ──
  _renderSteps() {
    const total = 4;
    const el = document.getElementById('post-sale-steps');
    if (!el) return;
    let html = '';
    for (let i = 1; i <= total; i++) {
      const cls = i < this._step ? 'completed' : i === this._step ? 'active' : '';
      const icon = i < this._step ? '✓' : i;
      html += `<div class="wizard-step-dot ${cls}">${icon}</div>`;
      if (i < total) html += `<div class="wizard-step-line ${i < this._step ? 'completed' : ''}"></div>`;
    }
    el.innerHTML = html;
  },

  // ── Body ──
  _renderBody() {
    const el = document.getElementById('post-sale-body');
    if (!el) return;
    if (this._step === 1) el.innerHTML = this._step1HTML();
    else if (this._step === 2) el.innerHTML = this._step2HTML();
    else if (this._step === 3) el.innerHTML = this._reviewHTML();
    else if (this._step === 4) el.innerHTML = this._successHTML();
  },

  // ── Step 1: Sale Info ──
  _step1HTML() {
    const d = this._formData;
    return `<div class="wizard-step-content">
      <div class="wizard-field">
        <label class="wizard-label">Date of Sale</label>
        <input type="date" class="wizard-input" id="ps-date" value="${d.dateOfSale}" max="${new Date().toISOString().split('T')[0]}">
      </div>

      <div class="wizard-field" id="ps-dsi-field">
        <label class="wizard-label">SPM Number</label>
        <input type="text" class="wizard-input" id="ps-dsi" value="${this._esc(d.dsi)}" placeholder="Enter 12+ character SPM" oninput="PostSale._updateDSIHint()">
        <div class="wizard-hint" id="ps-dsi-hint">${d.dsi.length}/12 min characters</div>
        <div class="wizard-error">SPM must be at least 12 characters</div>
      </div>

      <div class="wizard-field">
        <label class="wizard-label">Type of Account</label>
        <div class="toggle-group" id="ps-account-type-toggle">
          <button class="toggle-btn ${d.accountType === 'Consumer' ? 'active' : ''}" onclick="PostSale.setAccountType('Consumer')">Consumer</button>
          <button class="toggle-btn ${d.accountType === 'Business' ? 'active' : ''}" onclick="PostSale.setAccountType('Business')">Business</button>
        </div>
      </div>

      <div class="wizard-field">
        <label class="wizard-label">Did you have a trainee?</label>
        <div class="toggle-group">
          <button class="toggle-btn ${d.trainee ? 'active' : ''}" onclick="PostSale.setTrainee(true)">Yes</button>
          <button class="toggle-btn ${!d.trainee ? 'active' : ''}" onclick="PostSale.setTrainee(false)">No</button>
        </div>
      </div>

      <div id="ps-trainee-wrap" style="display:${d.trainee ? 'block' : 'none'}">
        <div class="wizard-field">
          <label class="wizard-label">Trainee's Name</label>
          <input type="text" class="wizard-input" id="ps-trainee-name" value="${this._esc(d.traineeName)}" placeholder="Trainee's full name">
        </div>
      </div>

      <div class="wizard-field">
        <label class="wizard-label">How was this order processed?</label>
        <div class="toggle-group" id="ps-order-channel-toggle">
          <button class="toggle-btn ${d.orderChannel === 'Sara' ? 'active' : ''}" onclick="PostSale.setOrderChannel('Sara')">Sara</button>
          <button class="toggle-btn ${d.orderChannel === 'Tower' ? 'active' : ''}" onclick="PostSale.setOrderChannel('Tower')">Tower</button>
        </div>
        ${d.orderChannel === 'Tower' ? '<div class="wizard-hint" style="color:var(--orange);margin-top:4px">Tower orders are tracked but excluded from the leaderboard.</div>' : ''}
      </div>

      <div class="wizard-field">
        <label class="wizard-label">Was this sale made under someone else's codes?</label>
        <div class="toggle-group" id="ps-codes-used-toggle">
          <button class="toggle-btn ${d.codesUsed ? 'active' : ''}" onclick="PostSale.setCodesUsed(true)">Yes</button>
          <button class="toggle-btn ${!d.codesUsed ? 'active' : ''}" onclick="PostSale.setCodesUsed(false)">No</button>
        </div>
      </div>

      <div id="ps-codes-used-wrap" style="display:${d.codesUsed ? 'block' : 'none'}">
        <div class="wizard-field">
          <label class="wizard-label">Whose codes were used?</label>
          <select class="wizard-input" id="ps-codes-used-by">
            <option value="">Select rep...</option>
            ${this._buildRepOptions(d.codesUsedBy)}
          </select>
        </div>
      </div>

      <div class="wizard-field">
        <label class="wizard-label">Additional Notes <span style="font-weight:400;text-transform:none">(optional)</span></label>
        <textarea class="wizard-input" id="ps-notes" rows="2" placeholder="Any extra details about the account...">${this._esc(d.accountNotes)}</textarea>
      </div>

      <div class="wizard-field">
        <label class="wizard-label">Hashtags <span style="font-weight:400;text-transform:none">(optional — posted to Discord)</span></label>
        <input type="text" class="wizard-input" id="ps-hashtags" value="${this._esc(d.hashtags)}" placeholder="#grindteam #letsgoo">
      </div>
    </div>`;
  },

  // ── Step 2: Products (Air + Wireless only) ──
  _step2HTML() {
    const prods = this._products;
    let cardsHTML = '';
    const keys = ['air', 'wireless'];
    keys.forEach(k => {
      const p = prods[k];
      cardsHTML += `
        <div class="product-card ${p.on ? 'selected' : ''}" onclick="PostSale.toggleProduct('${k}')">
          <div class="product-check">✓</div>
          <div class="product-icon">${p.icon}</div>
          <div class="product-label">${p.label}</div>
        </div>`;
      // Sub-fields for this product
      cardsHTML += this._productSubHTML(k);
    });

    return `<div class="wizard-step-content">
      <div style="font-family:var(--font-subheading);font-size:11px;font-weight:700;letter-spacing:0.5px;text-transform:uppercase;color:var(--silver-dim);margin-bottom:12px">Select Products Sold</div>
      <div class="product-cards" id="ps-product-grid">${cardsHTML}</div>
      <div class="wizard-error" id="ps-products-error" style="text-align:center;margin-top:8px">Select at least one product</div>
    </div>`;
  },

  _productSubHTML(key) {
    const d = this._formData;
    const open = this._products[key].on;
    let inner = '';
    switch (key) {
      case 'air':
        // Activation Support is Elevate-only (off_001)
        if (OFFICE_CONFIG.officeId === 'off_001') {
          inner = `
            <div class="wizard-field">
              <label class="wizard-label">Booked for Activation Support?</label>
              <div style="margin-bottom:8px"><a href="https://aspireteam.zohobookings.com/#/activation-support" target="_blank" rel="noopener" style="color:var(--sc-cyan);font-size:13px;text-decoration:underline">Book Activation Support Here</a></div>
              <div class="toggle-group" id="ps-activation-toggle">
                <button class="toggle-btn ${d.activationSupport ? 'active' : ''}" onclick="PostSale.setField('activationSupport',true)">Yes</button>
                <button class="toggle-btn ${!d.activationSupport ? 'active' : ''}" onclick="PostSale.setField('activationSupport',false)">No</button>
              </div>
              <div class="wizard-error" id="ps-activation-error">You must book the activation support appointment to proceed</div>
            </div>`;
        }
        break;
      case 'wireless':
        inner = `
          <div style="display:flex;gap:12px">
            <div class="wizard-field" style="flex:1">
              <label class="wizard-label">New Phones</label>
              <input type="number" class="wizard-input" id="ps-new-phones" value="${d.newPhones}" min="0" oninput="PostSale._updateCellTotal()">
            </div>
            <div class="wizard-field" style="flex:1">
              <label class="wizard-label">BYODs</label>
              <input type="number" class="wizard-input" id="ps-byods" value="${d.byods}" min="0" oninput="PostSale._updateCellTotal()">
            </div>
          </div>
          <div class="wizard-hint">Total lines: <b id="ps-cell-total">${d.newPhones + d.byods}</b></div>
          <div class="wizard-error" id="ps-wireless-error">Enter at least 1 phone or BYOD</div>`;
        break;
    }
    return `<div class="product-subfields ${open ? 'open' : ''}" id="ps-sub-${key}">${inner}</div>`;
  },

  // ── Step 3: Review ──
  _reviewHTML() {
    const d = this._formData;
    const session = Auth.getSession();
    const prods = this._products;

    let productsHTML = '';
    const sold = [];
    if (prods.air.on) sold.push('Internet Air' + (d.activationSupport ? ' (Activation Booked)' : ''));
    if (prods.wireless.on) {
      const total = (d.newPhones || 0) + (d.byods || 0);
      const parts = [];
      if (d.newPhones > 0) parts.push(d.newPhones + ' New');
      if (d.byods > 0) parts.push(d.byods + ' BYOD');
      sold.push('Wireless x' + total + ' (' + parts.join(', ') + ')');
    }
    productsHTML = sold.map(s => `<div class="review-row"><span class="review-row-value">✓ ${s}</span></div>`).join('');

    const editStep1 = `onclick="PostSale.goToStep(1)"`;
    const editStep2 = `onclick="PostSale.goToStep(2)"`;

    return `<div class="wizard-step-content">
      <div class="review-card">
        <div class="review-card-title">Sale Info <span class="review-edit" ${editStep1}>Edit</span></div>
        <div class="review-row"><span class="review-row-label">Rep</span><span class="review-row-value">${this._esc(session?.name || '')}</span></div>
        <div class="review-row"><span class="review-row-label">Date</span><span class="review-row-value">${this._formatDate(d.dateOfSale)}</span></div>
        <div class="review-row"><span class="review-row-label">Campaign</span><span class="review-row-value">AT&T B2B</span></div>
        <div class="review-row"><span class="review-row-label">SPM</span><span class="review-row-value">${this._esc(d.dsi)}</span></div>
        <div class="review-row"><span class="review-row-label">Account Type</span><span class="review-row-value">${d.accountType}</span></div>
        ${d.trainee ? `<div class="review-row"><span class="review-row-label">Trainee</span><span class="review-row-value">${this._esc(d.traineeName)}</span></div>` : ''}
        <div class="review-row"><span class="review-row-label">Processed Via</span><span class="review-row-value${d.orderChannel === 'Tower' ? '" style="color:var(--orange);font-weight:700' : ''}">${d.orderChannel}${d.orderChannel === 'Tower' ? ' (no leaderboard)' : ''}</span></div>
        ${d.codesUsed && d.codesUsedBy ? `<div class="review-row"><span class="review-row-label">Codes Used</span><span class="review-row-value" style="color:var(--orange)">${this._esc(d.codesUsedByName || d.codesUsedBy)}</span></div>` : ''}
        ${d.accountNotes ? `<div class="review-row"><span class="review-row-label">Notes</span><span class="review-row-value" style="max-width:60%;text-align:right">${this._esc(d.accountNotes)}</span></div>` : ''}
        ${d.hashtags ? `<div class="review-row"><span class="review-row-label">Hashtags</span><span class="review-row-value" style="color:var(--sc-cyan)">${this._esc(d.hashtags)}</span></div>` : ''}
      </div>

      <div class="review-card">
        <div class="review-card-title">Products <span class="review-edit" ${editStep2}>Edit</span></div>
        ${productsHTML}
      </div>
    </div>`;
  },

  // ── Success ──
  _successHTML() {
    const d = this._formData;
    const soldItems = [];
    if (this._products.air.on) soldItems.push('Air');
    if (this._products.wireless.on) soldItems.push('Cell x' + ((d.newPhones || 0) + (d.byods || 0)));
    const summary = d.dsi + ' — ' + soldItems.join(', ');

    return `<div class="wizard-step-content" style="text-align:center;padding-top:40px">
      <div class="success-check">✓</div>
      <div class="success-summary">
        <div class="success-title">Sale Logged!</div>
        <div class="success-subtitle">${this._formatDate(d.dateOfSale)}</div>
        <div style="font-size:14px;color:var(--white);font-weight:600;margin-top:12px">${this._esc(summary)}</div>
      </div>
      <div class="success-actions">
        <button class="wizard-btn wizard-btn-primary" onclick="PostSale._resetForAnother()">Log Another Sale</button>
        <button class="wizard-btn wizard-btn-back" onclick="App.navTo('leaderboard')">Back to Dashboard</button>
      </div>
    </div>`;
  },

  // ── Nav Buttons ──
  _renderNav() {
    const el = document.getElementById('post-sale-nav');
    if (!el) return;
    const total = 4;
    const isReview = this._step === 3;
    const isSuccess = this._step === total;

    if (isSuccess) { el.innerHTML = ''; return; }

    const backBtn = this._step > 1
      ? `<button class="wizard-btn wizard-btn-back" onclick="PostSale.prevStep()">Back</button>`
      : '<div></div>';

    const nextLabel = isReview ? 'Submit Sale' : 'Next';
    const nextAction = isReview ? 'PostSale.submit()' : 'PostSale.nextStep()';
    const nextBtn = `<button class="wizard-btn wizard-btn-primary" id="ps-next-btn" onclick="${nextAction}" ${this._submitting ? 'disabled' : ''}>${this._submitting ? 'Submitting...' : nextLabel}</button>`;

    el.innerHTML = backBtn + nextBtn;
  },

  // ── Navigation ──
  goToStep(n) {
    this._collectCurrentStep();
    this._step = n;
    this._render();
    // Scroll to top of form
    const page = document.getElementById('post-sale-page');
    if (page) page.scrollTop = 0;
  },

  nextStep() {
    this._collectCurrentStep();
    if (!this._validateCurrentStep()) return;
    this._step++;
    this._render();
    const page = document.getElementById('post-sale-page');
    if (page) page.scrollTop = 0;
  },

  prevStep() {
    this._collectCurrentStep();
    if (this._step > 1) this._step--;
    this._render();
  },

  _collectCurrentStep() {
    if (this._step === 1) this._collectStep1();
    else if (this._step === 2) this._collectStep2();
  },

  // ── Validation ──
  _validateCurrentStep() {
    if (this._step === 1) return this._validateStep1();
    if (this._step === 2) return this._validateStep2();
    return true;
  },

  _validateStep1() {
    let valid = true;
    const d = this._formData;

    if (!d.dateOfSale) {
      valid = false;
    } else if (d.dateOfSale > new Date().toISOString().split('T')[0]) {
      valid = false;
      alert('Sale date cannot be in the future.');
    }

    const dsiField = document.getElementById('ps-dsi-field');
    if ((d.dsi || '').length < 12) {
      valid = false;
      if (dsiField) dsiField.classList.add('has-error');
    } else {
      if (dsiField) dsiField.classList.remove('has-error');
    }

    if (d.trainee && !d.traineeName.trim()) {
      // Trainee name is nice-to-have, don't block
    }

    // Codes-used-by: if "Yes" was selected, a rep must be chosen
    if (d.codesUsed && !d.codesUsedBy) {
      valid = false;
      const codesWrap = document.getElementById('ps-codes-used-wrap');
      if (codesWrap) {
        const selectField = codesWrap.querySelector('.wizard-field');
        if (selectField) selectField.classList.add('has-error');
      }
    }

    return valid;
  },

  _validateStep2() {
    let valid = true;
    const prods = this._products;
    const d = this._formData;

    // At least one product selected
    const anyOn = Object.values(prods).some(p => p.on);
    const errEl = document.getElementById('ps-products-error');
    if (!anyOn) {
      if (errEl) errEl.style.display = 'block';
      return false;
    }
    if (errEl) errEl.style.display = 'none';

    // Air: activation support must be Yes (Elevate only)
    if (OFFICE_CONFIG.officeId === 'off_001' && prods.air.on && !d.activationSupport) {
      valid = false;
      const actErr = document.getElementById('ps-activation-error');
      if (actErr) actErr.style.display = 'block';
      alert('You must book the activation support appointment before submitting.');
      return false;
    }

    // Wireless: at least 1 phone or byod
    if (prods.wireless.on) {
      const total = (d.newPhones || 0) + (d.byods || 0);
      if (total < 1) {
        valid = false;
        const wErr = document.getElementById('ps-wireless-error');
        if (wErr) wErr.style.display = 'block';
      }
    }

    return valid;
  },

  // ── Submission ──

  async submit() {
    this._collectCurrentStep();
    if (this._submitting) return;
    this._submitting = true;
    this._renderNav();

    const payload = this._buildPayload();
    try {
      const result = await SheetsAPI.post(OFFICE_CONFIG, 'addSale', payload);
      this._submitting = false;
      if (result.ok || (result.data && result.data.ok)) {
        this._clearDraft();
        // Fire Discord webhook (fire-and-forget — don't block success)
        this._fireDiscordWebhook(payload);
        // Advance to success step
        this._step = 4;
        this._render();
      } else {
        const errMsg = (result.data && result.data.error) || 'Something went wrong. Please try again.';
        alert(errMsg);
        this._renderNav();
      }
    } catch (err) {
      this._submitting = false;
      alert('Network error — please check your connection and try again.');
      this._renderNav();
    }
  },

  _fireDiscordWebhook(payload) {
    try {
      // Build message matching #production format
      let msg = '';
      let units = 0;
      // Line 1: **Rep** made a sale with AT&T: B2B!
      msg += '**' + payload.repName + '** made a sale with AT&T: B2B!\n';
      // Line 2: Account type
      msg += (payload.accountType || 'Business') + ' Account\n';
      // Line 3: SPM
      msg += payload.dsi + '\n';
      // Product bullet lines
      if (payload.air) { msg += '• Internet Air\n'; units++; }
      if (payload.newPhones || payload.byods) {
        msg += '• ' + (payload.newPhones || 0) + ' New Phone(s)|' + (payload.byods || 0) + ' BYOD(s)\n';
        units += (payload.newPhones || 0) + (payload.byods || 0);
      }

      // Hashtags (optional, from form)
      const tags = this._formData.hashtags?.trim();
      if (tags) msg += tags + '\n';

      // Team emoji x units
      const emoji = this._getTeamEmoji();
      if (emoji && units > 0) msg += emoji.repeat(Math.min(units, 20));

      const finalMsg = msg.trim();

      // Per-office Discord webhook URL (from _Offices tab) — no fallback
      const webhookUrl = OFFICE_CONFIG.discordWebhookUrl;
      if (!webhookUrl) {
        console.warn('[PostSale] No Discord webhook configured for this office');
        return;
      }

      const webhookBody = JSON.stringify({ content: finalMsg });

      const opts = { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: webhookBody };
      fetch(webhookUrl, opts)
        .then(r => { if (!r.ok) console.warn('[PostSale] Discord webhook HTTP', r.status); })
        .catch(err => {
          console.warn('[PostSale] Discord webhook failed, retrying…', err.message);
          // Retry once after 2s
          setTimeout(() => {
            fetch(webhookUrl, opts)
              .then(r => { if (!r.ok) console.warn('[PostSale] Retry HTTP', r.status); })
              .catch(e2 => console.warn('[PostSale] Retry failed:', e2.message));
          }, 2000);
        });
    } catch (e) { console.warn('[PostSale] Webhook build error:', e.message); }
  },

  _buildPayload() {
    const session = Auth.getSession();
    const d = this._formData;
    const prods = this._products;

    const newPhones = prods.wireless.on ? (d.newPhones || 0) : 0;
    const byods = prods.wireless.on ? (d.byods || 0) : 0;

    return {
      email: session?.email || '',
      repName: session?.name || '',
      campaign: 'attb2b',
      dateOfSale: d.dateOfSale,
      dsi: d.dsi,
      accountType: d.accountType,
      accountNotes: d.accountNotes || '',
      trainee: d.trainee,
      traineeName: d.trainee ? d.traineeName : '',
      // Products (Air + Wireless only)
      air: prods.air.on ? 1 : 0,
      activationSupport: prods.air.on && d.activationSupport,
      newPhones: newPhones,
      byods: byods,
      // Order channel & codes
      orderChannel: d.orderChannel || 'Sara',
      codesUsedBy: d.codesUsed ? (d.codesUsedBy || '') : ''
    };
  },

  // ── UI Actions ──
  setAccountType(type) {
    this._formData.accountType = type;
    this._saveDraft();
    const group = document.getElementById('ps-account-type-toggle');
    if (group) group.querySelectorAll('.toggle-btn').forEach(btn => {
      btn.classList.toggle('active', btn.textContent.trim() === type);
    });
  },

  setTrainee(val) {
    this._formData.trainee = val;
    this._saveDraft();
    const wrap = document.getElementById('ps-trainee-wrap');
    if (wrap) wrap.style.display = val ? 'block' : 'none';
    // Update toggle buttons
    const btns = document.querySelectorAll('#post-sale-body .wizard-field');
    btns.forEach(field => {
      const label = field.querySelector('.wizard-label');
      if (label && label.textContent.includes('trainee')) {
        field.querySelectorAll('.toggle-btn').forEach(btn => {
          btn.classList.toggle('active', (btn.textContent.trim() === 'Yes') === val);
        });
      }
    });
  },

  setOrderChannel(val) {
    this._formData.orderChannel = val;
    this._saveDraft();
    // Update toggle buttons
    const toggle = document.getElementById('ps-order-channel-toggle');
    if (toggle) toggle.querySelectorAll('.toggle-btn').forEach(btn => {
      btn.classList.toggle('active', btn.textContent.trim() === val);
    });
    // Re-render to show/hide Tower hint
    this._renderBody();
  },

  setCodesUsed(val) {
    this._formData.codesUsed = val;
    if (!val) {
      this._formData.codesUsedBy = '';
      this._formData.codesUsedByName = '';
    }
    this._saveDraft();
    const wrap = document.getElementById('ps-codes-used-wrap');
    if (wrap) wrap.style.display = val ? 'block' : 'none';
    // Update toggle buttons
    const toggle = document.getElementById('ps-codes-used-toggle');
    if (toggle) toggle.querySelectorAll('.toggle-btn').forEach(btn => {
      btn.classList.toggle('active', (btn.textContent.trim() === 'Yes') === val);
    });
  },

  _buildRepOptions(selectedEmail) {
    const roster = App.state?.roster || {};
    const session = Auth.getSession();
    const myEmail = (session?.email || '').toLowerCase();
    const LEADER_RANKS = new Set(['l1', 'jd', 'manager', 'owner']);
    return Object.entries(roster)
      .filter(([email, r]) => !r.deactivated && email !== myEmail && LEADER_RANKS.has(r.rank))
      .sort((a, b) => (a[1].name || a[0]).localeCompare(b[1].name || b[0]))
      .map(([email, r]) =>
        `<option value="${email}" ${email === selectedEmail ? 'selected' : ''}>${this._esc(r.name || email)}</option>`)
      .join('');
  },

  setField(key, val) {
    this._formData[key] = val;
    this._saveDraft();
    // Update toggle buttons in current context
    const subEl = document.getElementById('ps-sub-' + (key === 'activationSupport' ? 'air' : key));
    if (subEl) {
      subEl.querySelectorAll('.toggle-btn').forEach(btn => {
        btn.classList.toggle('active', (btn.textContent.trim() === 'Yes') === val);
      });
      // Clear activation error when Yes is selected
      if (key === 'activationSupport' && val) {
        const actErr = document.getElementById('ps-activation-error');
        if (actErr) actErr.style.display = 'none';
      }
    }
  },

  toggleProduct(key) {
    this._collectStep2();
    this._products[key].on = !this._products[key].on;
    this._saveDraft();
    // Re-render step 2
    const body = document.getElementById('post-sale-body');
    if (body) body.innerHTML = this._step2HTML();
  },

  // ── Live update helpers ──
  _updateCellTotal() {
    const phones = parseInt(document.getElementById('ps-new-phones')?.value) || 0;
    const byods = parseInt(document.getElementById('ps-byods')?.value) || 0;
    const totalEl = document.getElementById('ps-cell-total');
    if (totalEl) totalEl.textContent = phones + byods;
    // Also update formData live
    this._formData.newPhones = phones;
    this._formData.byods = byods;
    this._saveDraft();
  },

  _updateDSIHint() {
    const input = document.getElementById('ps-dsi');
    const hint = document.getElementById('ps-dsi-hint');
    if (input && hint) {
      const len = input.value.length;
      hint.textContent = len + '/12 min characters';
      hint.style.color = len >= 12 ? 'var(--green)' : 'var(--silver-dim)';
    }
    // Clear error if valid
    const field = document.getElementById('ps-dsi-field');
    if (field && input && input.value.length >= 12) field.classList.remove('has-error');
  },

  // ── Helpers ──
  _getTeamEmoji() {
    try {
      const session = Auth.getSession();
      if (!session?.email || !App.state?.roster) return '';
      const email = session.email.toLowerCase();
      const person = App.state.roster.find(r => r.email?.toLowerCase() === email);
      const teamName = person?.team;
      if (!teamName || !App.state.teamsData) return '';
      const team = App.state.teamsData.find(t => t.name === teamName);
      return team?.emoji || '';
    } catch (e) { return ''; }
  },

  _esc(s) { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); },

  _formatDate(iso) {
    if (!iso) return '—';
    const d = new Date(iso + 'T12:00:00');
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    return months[d.getMonth()] + ' ' + d.getDate() + ', ' + d.getFullYear();
  }
};
