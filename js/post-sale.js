// ═══════════════════════════════════════════════════════
// POST SALE — Multi-step wizard form
// ═══════════════════════════════════════════════════════
// Replaces external Google Form with in-dashboard sale logging.
// Conditional branching: only shows relevant follow-up fields.

const PostSale = {
  // ── State ──
  _step: 1,
  _campaign: 'attb2b',   // 'attb2b' | 'ooma'
  _submitting: false,
  _formData: {},
  _products: {
    air:      { on: false, icon: '📡', label: 'Internet Air' },
    wireless: { on: false, icon: '📱', label: 'Wireless' },
    fiber:    { on: false, icon: '🌐', label: 'Fiber' },
    voip:     { on: false, icon: '🎧', label: 'VoIP' },
    dtv:      { on: false, icon: '📺', label: 'DIRECTV' }
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
      fiberPackage: '',
      installDate: '',
      voipQty: 0,
      dtvPackage: '',
      clientName: '',
      oomaPackage: 'Ooma Pro',
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
      this._campaign = draft.campaign || 'attb2b';
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
    if (this._campaign === 'attb2b') {
      this._formData.dsi = v('ps-dsi');
      this._formData.accountNotes = v('ps-notes');
    } else {
      this._formData.clientName = v('ps-client-name');
    }
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
    const v = (id) => { const el = document.getElementById(id); return el ? el.value : ''; };
    const n = (id) => { const el = document.getElementById(id); return el ? (parseInt(el.value) || 0) : 0; };
    this._formData.newPhones = n('ps-new-phones');
    this._formData.byods = n('ps-byods');
    this._formData.fiberPackage = v('ps-fiber-pkg');
    this._formData.installDate = v('ps-install-date');
    this._formData.voipQty = n('ps-voip-qty');
    this._formData.dtvPackage = v('ps-dtv-pkg');
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
    const total = this._campaign === 'ooma' ? 3 : 4;
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
    else if (this._step === 2) el.innerHTML = this._campaign === 'ooma' ? this._reviewHTML() : this._step2HTML();
    else if (this._step === 3) el.innerHTML = this._campaign === 'ooma' ? this._successHTML() : this._reviewHTML();
    else if (this._step === 4) el.innerHTML = this._successHTML();
  },

  // ── Step 1: Sale Info ──
  _step1HTML() {
    const d = this._formData;
    return `<div class="wizard-step-content">
      <div class="wizard-label" style="margin-bottom:8px">Campaign</div>
      <div class="campaign-cards">
        <div class="campaign-card ${this._campaign === 'attb2b' ? 'selected' : ''}" onclick="PostSale.setCampaign('attb2b')">
          <img src="references/logos/logo-att.png" alt="AT&T B2B" class="campaign-logo">
        </div>
        <div class="campaign-card ${this._campaign === 'ooma' ? 'selected' : ''}" onclick="PostSale.setCampaign('ooma')">
          <img src="references/logos/logo-ooma.png" alt="Ooma" class="campaign-logo">
        </div>
      </div>

      <div class="wizard-field">
        <label class="wizard-label">Date of Sale</label>
        <input type="date" class="wizard-input" id="ps-date" value="${d.dateOfSale}">
      </div>

      ${this._campaign === 'attb2b' ? `
        <div class="wizard-field" id="ps-dsi-field">
          <label class="wizard-label">DSI Number</label>
          <input type="text" class="wizard-input" id="ps-dsi" value="${this._esc(d.dsi)}" placeholder="Enter 12+ character DSI" oninput="PostSale._updateDSIHint()">
          <div class="wizard-hint" id="ps-dsi-hint">${d.dsi.length}/12 min characters</div>
          <div class="wizard-error">DSI must be at least 12 characters</div>
        </div>

        <div class="wizard-field">
          <label class="wizard-label">Type of Account</label>
          <div class="toggle-group" id="ps-account-type-toggle">
            <button class="toggle-btn ${d.accountType === 'Consumer' ? 'active' : ''}" onclick="PostSale.setAccountType('Consumer')">Consumer</button>
            <button class="toggle-btn ${d.accountType === 'Business' ? 'active' : ''}" onclick="PostSale.setAccountType('Business')">Business</button>
          </div>
        </div>
      ` : `
        <div class="wizard-field">
          <label class="wizard-label">Client Name</label>
          <input type="text" class="wizard-input" id="ps-client-name" value="${this._esc(d.clientName)}" placeholder="Customer's name">
          <div class="wizard-error">Client name is required</div>
        </div>

        <div class="wizard-field">
          <label class="wizard-label">Package</label>
          <div class="toggle-group">
            <button class="toggle-btn ${d.oomaPackage === 'Ooma Essentials' ? 'active' : ''}" onclick="PostSale.setOomaPackage('Ooma Essentials')">Essentials</button>
            <button class="toggle-btn ${d.oomaPackage === 'Ooma Pro' ? 'active' : ''}" onclick="PostSale.setOomaPackage('Ooma Pro')">Pro</button>
            <button class="toggle-btn ${d.oomaPackage === 'Ooma Pro Plus' ? 'active' : ''}" onclick="PostSale.setOomaPackage('Ooma Pro Plus')">Pro Plus</button>
          </div>
        </div>
      `}

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

      ${this._campaign === 'attb2b' ? `
        <div class="wizard-field">
          <label class="wizard-label">Additional Notes <span style="font-weight:400;text-transform:none">(optional)</span></label>
          <textarea class="wizard-input" id="ps-notes" rows="2" placeholder="Any extra details about the account...">${this._esc(d.accountNotes)}</textarea>
        </div>
      ` : ''}

      <div class="wizard-field">
        <label class="wizard-label">Hashtags <span style="font-weight:400;text-transform:none">(optional — posted to Discord)</span></label>
        <input type="text" class="wizard-input" id="ps-hashtags" value="${this._esc(d.hashtags)}" placeholder="#grindteam #letsgoo">
      </div>
    </div>`;
  },

  // ── Step 2: Products (AT&T B2B only) ──
  _step2HTML() {
    const prods = this._products;
    let cardsHTML = '';
    const keys = ['air', 'wireless', 'fiber', 'voip', 'dtv'];
    keys.forEach(k => {
      const p = prods[k];
      // VoIP only available when Fiber is selected
      if (k === 'voip' && !prods.fiber.on) return;
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
        inner = `
          <div class="wizard-field">
            <label class="wizard-label">Booked for Activation Support?</label>
            <div class="toggle-group" id="ps-activation-toggle">
              <button class="toggle-btn ${d.activationSupport ? 'active' : ''}" onclick="PostSale.setField('activationSupport',true)">Yes</button>
              <button class="toggle-btn ${!d.activationSupport ? 'active' : ''}" onclick="PostSale.setField('activationSupport',false)">No</button>
            </div>
            <div class="wizard-error" id="ps-activation-error">You must book the activation support appointment to proceed</div>
          </div>`;
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
      case 'fiber':
        inner = `
          <div class="wizard-field" id="ps-fiber-pkg-field">
            <label class="wizard-label">Package</label>
            <select class="wizard-input" id="ps-fiber-pkg" onchange="PostSale._saveDraft()">
              <option value="" ${!d.fiberPackage ? 'selected' : ''}>Select package...</option>
              <option value="Fiber 300" ${d.fiberPackage === 'Fiber 300' ? 'selected' : ''}>Fiber 300</option>
              <option value="Fiber 500" ${d.fiberPackage === 'Fiber 500' ? 'selected' : ''}>Fiber 500</option>
              <option value="Fiber 1 GIG" ${d.fiberPackage === 'Fiber 1 GIG' ? 'selected' : ''}>Fiber 1 GIG</option>
              <option value="Fiber 2 GIG" ${d.fiberPackage === 'Fiber 2 GIG' ? 'selected' : ''}>Fiber 2 GIG</option>
              <option value="Fiber 5 GIG" ${d.fiberPackage === 'Fiber 5 GIG' ? 'selected' : ''}>Fiber 5 GIG</option>
            </select>
            <div class="wizard-error">Select a Fiber package</div>
          </div>
          <div class="wizard-field" id="ps-install-date-field">
            <label class="wizard-label">Install Date</label>
            <input type="date" class="wizard-input" id="ps-install-date" value="${d.installDate}" onchange="PostSale._saveDraft()">
            <div class="wizard-error">Install date is required</div>
          </div>`;
        break;
      case 'voip':
        inner = `
          <div class="wizard-field" id="ps-voip-qty-field">
            <label class="wizard-label">Quantity of Lines</label>
            <input type="number" class="wizard-input" id="ps-voip-qty" value="${d.voipQty}" min="1" oninput="PostSale._saveDraft()">
            <div class="wizard-error">Enter at least 1 VoIP line</div>
          </div>`;
        break;
      case 'dtv':
        inner = `
          <div class="wizard-field" id="ps-dtv-pkg-field">
            <label class="wizard-label">DIRECTV Package</label>
            <select class="wizard-input" id="ps-dtv-pkg" onchange="PostSale._saveDraft()">
              <option value="" ${!d.dtvPackage ? 'selected' : ''}>Select package...</option>
              <option value="Entertainment" ${d.dtvPackage === 'Entertainment' ? 'selected' : ''}>Entertainment</option>
              <option value="Choice" ${d.dtvPackage === 'Choice' ? 'selected' : ''}>Choice</option>
              <option value="Ultimate" ${d.dtvPackage === 'Ultimate' ? 'selected' : ''}>Ultimate</option>
              <option value="Premier" ${d.dtvPackage === 'Premier' ? 'selected' : ''}>Premier</option>
            </select>
            <div class="wizard-error">Select a DIRECTV package</div>
          </div>`;
        break;
    }
    return `<div class="product-subfields ${open ? 'open' : ''}" id="ps-sub-${key}">${inner}</div>`;
  },

  // ── Step 3 (AT&T) / Step 2 (Ooma): Review ──
  _reviewHTML() {
    const d = this._formData;
    const session = Auth.getSession();
    const prods = this._products;

    let productsHTML = '';
    if (this._campaign === 'attb2b') {
      const sold = [];
      if (prods.air.on) sold.push('Internet Air' + (d.activationSupport ? ' (Activation Booked)' : ''));
      if (prods.wireless.on) {
        const total = (d.newPhones || 0) + (d.byods || 0);
        const parts = [];
        if (d.newPhones > 0) parts.push(d.newPhones + ' New');
        if (d.byods > 0) parts.push(d.byods + ' BYOD');
        sold.push('Wireless x' + total + ' (' + parts.join(', ') + ')');
      }
      if (prods.fiber.on) sold.push('Fiber' + (d.fiberPackage ? ' — ' + d.fiberPackage : ''));
      if (prods.voip.on) sold.push('VoIP x' + (d.voipQty || 1));
      if (prods.dtv.on) sold.push('DIRECTV' + (d.dtvPackage ? ' — ' + d.dtvPackage : ''));
      productsHTML = sold.map(s => `<div class="review-row"><span class="review-row-value">✓ ${s}</span></div>`).join('');
    }

    const editStep1 = `onclick="PostSale.goToStep(1)"`;
    const editStep2 = `onclick="PostSale.goToStep(2)"`;

    return `<div class="wizard-step-content">
      <div class="review-card">
        <div class="review-card-title">Sale Info <span class="review-edit" ${editStep1}>Edit</span></div>
        <div class="review-row"><span class="review-row-label">Rep</span><span class="review-row-value">${this._esc(session?.name || '')}</span></div>
        <div class="review-row"><span class="review-row-label">Date</span><span class="review-row-value">${this._formatDate(d.dateOfSale)}</span></div>
        <div class="review-row"><span class="review-row-label">Campaign</span><span class="review-row-value">${this._campaign === 'attb2b' ? 'AT&T B2B' : 'Ooma'}</span></div>
        ${this._campaign === 'attb2b' ? `
          <div class="review-row"><span class="review-row-label">DSI</span><span class="review-row-value">${this._esc(d.dsi)}</span></div>
          <div class="review-row"><span class="review-row-label">Account Type</span><span class="review-row-value">${d.accountType}</span></div>
        ` : `
          <div class="review-row"><span class="review-row-label">Client</span><span class="review-row-value">${this._esc(d.clientName)}</span></div>
          <div class="review-row"><span class="review-row-label">Package</span><span class="review-row-value">${d.oomaPackage}</span></div>
        `}
        ${d.trainee ? `<div class="review-row"><span class="review-row-label">Trainee</span><span class="review-row-value">${this._esc(d.traineeName)}</span></div>` : ''}
        <div class="review-row"><span class="review-row-label">Processed Via</span><span class="review-row-value${d.orderChannel === 'Tower' ? '" style="color:var(--orange);font-weight:700' : ''}">${d.orderChannel}${d.orderChannel === 'Tower' ? ' (no leaderboard)' : ''}</span></div>
        ${d.codesUsed && d.codesUsedBy ? `<div class="review-row"><span class="review-row-label">Codes Used</span><span class="review-row-value" style="color:var(--orange)">${this._esc(d.codesUsedByName || d.codesUsedBy)}</span></div>` : ''}
        ${d.accountNotes ? `<div class="review-row"><span class="review-row-label">Notes</span><span class="review-row-value" style="max-width:60%;text-align:right">${this._esc(d.accountNotes)}</span></div>` : ''}
        ${d.hashtags ? `<div class="review-row"><span class="review-row-label">Hashtags</span><span class="review-row-value" style="color:var(--sc-cyan)">${this._esc(d.hashtags)}</span></div>` : ''}
      </div>

      ${this._campaign === 'attb2b' ? `
        <div class="review-card">
          <div class="review-card-title">Products <span class="review-edit" ${editStep2}>Edit</span></div>
          ${productsHTML}
        </div>
      ` : ''}
    </div>`;
  },

  // ── Success ──
  _successHTML() {
    const d = this._formData;
    const session = Auth.getSession();
    let summary = '';
    if (this._campaign === 'attb2b') {
      const soldItems = [];
      if (this._products.air.on) soldItems.push('Air');
      if (this._products.wireless.on) soldItems.push('Cell x' + ((d.newPhones || 0) + (d.byods || 0)));
      if (this._products.fiber.on) soldItems.push('Fiber');
      if (this._products.voip.on) soldItems.push('VoIP x' + (d.voipQty || 1));
      if (this._products.dtv.on) soldItems.push('DTV');
      summary = d.dsi + ' — ' + soldItems.join(', ');
    } else {
      summary = d.clientName + ' — ' + d.oomaPackage;
    }

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
    const total = this._campaign === 'ooma' ? 3 : 4;
    const isReview = (this._campaign === 'ooma' && this._step === 2) || (this._campaign === 'attb2b' && this._step === 3);
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
    else if (this._step === 2 && this._campaign === 'attb2b') this._collectStep2();
  },

  // ── Validation ──
  _validateCurrentStep() {
    if (this._step === 1) return this._validateStep1();
    if (this._step === 2 && this._campaign === 'attb2b') return this._validateStep2();
    return true;
  },

  _validateStep1() {
    let valid = true;
    const d = this._formData;

    if (!d.dateOfSale) {
      valid = false;
    }

    if (this._campaign === 'attb2b') {
      const dsiField = document.getElementById('ps-dsi-field');
      if ((d.dsi || '').length < 12) {
        valid = false;
        if (dsiField) dsiField.classList.add('has-error');
      } else {
        if (dsiField) dsiField.classList.remove('has-error');
      }
    } else {
      // Ooma: require client name
      const nameField = document.getElementById('ps-client-name');
      if (!d.clientName.trim()) {
        valid = false;
        if (nameField) nameField.parentElement.classList.add('has-error');
      }
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

    // Air: activation support must be Yes
    if (prods.air.on && !d.activationSupport) {
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

    // Fiber: package and install date required
    if (prods.fiber.on) {
      if (!d.fiberPackage) {
        valid = false;
        const f = document.getElementById('ps-fiber-pkg-field');
        if (f) f.classList.add('has-error');
      }
      if (!d.installDate) {
        valid = false;
        const f = document.getElementById('ps-install-date-field');
        if (f) f.classList.add('has-error');
      }
    }

    // VoIP: at least 1 line
    if (prods.voip.on && (d.voipQty || 0) < 1) {
      valid = false;
      const f = document.getElementById('ps-voip-qty-field');
      if (f) f.classList.add('has-error');
    }

    // DTV: package required
    if (prods.dtv.on && !d.dtvPackage) {
      valid = false;
      const f = document.getElementById('ps-dtv-pkg-field');
      if (f) f.classList.add('has-error');
    }

    return valid;
  },

  // ── Submission ──
  _DEFAULT_WEBHOOK_URL: 'https://hook.us2.make.com/rqxy9beu6ybplh8axdq4p6euuv4mc8jj',

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
        const total = this._campaign === 'ooma' ? 3 : 4;
        this._step = total;
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
      if (this._campaign === 'attb2b') {
        // Line 1: **Rep** made a sale with AT&T: B2B!
        msg += '**' + payload.repName + '** made a sale with AT&T: B2B!\n';
        // Line 2: Account type
        msg += (payload.accountType || 'Business') + ' Account\n';
        // Line 3: DSI
        msg += payload.dsi + '\n';
        // Product bullet lines
        if (payload.air) { msg += '• Internet Air\n'; units++; }
        if (payload.newPhones || payload.byods) {
          msg += '• ' + (payload.newPhones || 0) + ' New Phone(s)|' + (payload.byods || 0) + ' BYOD(s)\n';
          units += (payload.newPhones || 0) + (payload.byods || 0);
        }
        if (payload.fiber) { msg += '• ' + (payload.fiberPackage || 'Fiber') + '\n'; units++; }
        if (payload.voipQty) { msg += '• ' + payload.voipQty + ' VoIP(s)\n'; units += payload.voipQty; }
        if (payload.dtv) msg += '• DIRECTV ' + (payload.dtvPackage || '') + '\n';
      } else {
        // Ooma format
        msg += '**' + payload.repName + '** made a sale with Ooma!\n';
        msg += payload.clientName + '\n';
        msg += '• ' + (payload.oomaPackage || 'Ooma Pro') + '\n';
        units = 1;
      }

      // Hashtags (optional, from form)
      const tags = this._formData.hashtags?.trim();
      if (tags) msg += tags + '\n';

      // Team emoji × units
      const emoji = this._getTeamEmoji();
      if (emoji && units > 0) msg += emoji.repeat(Math.min(units, 20));

      const finalMsg = msg.trim();

      // Per-office webhook URL (from _Offices tab), fallback to default Make.com webhook
      const webhookUrl = OFFICE_CONFIG.discordWebhookUrl || this._DEFAULT_WEBHOOK_URL;

      // Direct Discord webhooks use { content }, Make.com uses { message }
      const isDiscordDirect = webhookUrl.includes('discord.com/api/webhooks');
      const webhookBody = isDiscordDirect
        ? JSON.stringify({ content: finalMsg })
        : JSON.stringify({ message: finalMsg });

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
      campaign: this._campaign,
      dateOfSale: d.dateOfSale,
      dsi: this._campaign === 'attb2b' ? d.dsi : '',
      accountType: this._campaign === 'attb2b' ? d.accountType : '',
      accountNotes: d.accountNotes || '',
      trainee: d.trainee,
      traineeName: d.trainee ? d.traineeName : '',
      // Products (AT&T B2B)
      air: prods.air.on ? 1 : 0,
      activationSupport: prods.air.on && d.activationSupport,
      newPhones: newPhones,
      byods: byods,
      fiber: prods.fiber.on ? 1 : 0,
      fiberPackage: prods.fiber.on ? d.fiberPackage : '',
      installDate: prods.fiber.on ? d.installDate : '',
      voipQty: prods.voip.on ? (d.voipQty || 1) : 0,
      dtv: prods.dtv.on ? 1 : 0,
      dtvPackage: prods.dtv.on ? d.dtvPackage : '',
      // Ooma
      clientName: this._campaign === 'ooma' ? d.clientName : '',
      oomaPackage: this._campaign === 'ooma' ? d.oomaPackage : '',
      // Order channel & codes
      orderChannel: d.orderChannel || 'Sara',
      codesUsedBy: d.codesUsed ? (d.codesUsedBy || '') : ''
    };
  },

  // ── UI Actions ──
  setCampaign(c) {
    this._collectStep1();
    this._campaign = c;
    this._saveDraft();
    this._render();
  },

  setAccountType(type) {
    this._formData.accountType = type;
    this._saveDraft();
    const group = document.getElementById('ps-account-type-toggle');
    if (group) group.querySelectorAll('.toggle-btn').forEach(btn => {
      btn.classList.toggle('active', btn.textContent.trim() === type);
    });
  },

  setOomaPackage(pkg) {
    this._formData.oomaPackage = pkg;
    this._saveDraft();
    this._render(); // re-render to update toggles
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
    return Object.entries(roster)
      .filter(([email, r]) => !r.deactivated && email !== myEmail)
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
    // Turning off Fiber also turns off VoIP (can't have VoIP without Fiber)
    if (key === 'fiber' && !this._products.fiber.on) {
      this._products.voip.on = false;
      this._formData.voipQty = 0;
    }
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
