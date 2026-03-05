// ═══════════════════════════════════════════════════════
// ELEVATE — Data Pipeline
// Transforms raw Google Sheet rows into the data
// structures that rendering functions consume.
// ═══════════════════════════════════════════════════════

const DataPipeline = {

  // ──────────────────────────────────────────────
  // CORE: Build people + teams from sheet data
  // ──────────────────────────────────────────────
  buildFromSheets(salesData, rosterData, config) {
    const roster = this.parseRoster(rosterData, config);
    const people = this.aggregateSales(salesData, roster, config);
    const teams  = this.buildTeams(people, config);
    return { people, teams, roster };
  },

  // ──────────────────────────────────────────────
  // BUILD FROM APPS SCRIPT (pre-aggregated data)
  // Transforms Code.gs doGet response → modular format
  // ──────────────────────────────────────────────
  buildFromAppsScript(apiData, config) {
    // Build roster map (email-keyed)
    const roster = {};
    if (apiData.roster) {
      Object.entries(apiData.roster).forEach(([email, info]) => {
        roster[email] = {
          name: info.name,
          email: email,
          role: info.rank || 'rep',
          rank: info.rank || 'rep',
          team: info.team || 'Unassigned',
          active: !info.deactivated,
          deactivated: info.deactivated || false,
          dateAdded: info.dateAdded || '',
          phone: info.phone || ''
        };
      });
    }

    // Build people from pre-aggregated data (skip deactivated)
    const people = (apiData.people || [])
      .filter(p => !p.deactivated)
      .map(p => this._buildPersonFromApi(p, config));

    // Use hierarchy builder if _Teams data exists, else flat fallback
    const teamsData = apiData.teams || {};
    const hasTeamsData = Object.keys(teamsData).length > 0;
    const teams = hasTeamsData
      ? this.buildTeamHierarchy(people, teamsData, config)
      : this.buildTeams(people, config);
    return { people, teams, roster };
  },

  // Convert flat period { y, air, cell, fiber, voip, dtv, units }
  // → modular format { y, units, products: { air, cell, fiber, voip } }
  _convertApiPeriod(flat, config) {
    const period = { y: flat?.y || 0, units: flat?.units || 0, products: {} };
    config.columns.products.forEach(prod => {
      period.products[prod.key] = flat?.[prod.key] || 0;
    });
    return period;
  },

  _buildPersonFromApi(p, config) {
    const days = (p.days || []).map(d => this._convertApiPeriod(d, config));
    while (days.length < 7) days.push(this.emptyPeriod(config));

    const lwDays = (p.lwDays || []).map(d => this._convertApiPeriod(d, config));
    while (lwDays.length < 7) lwDays.push(this.emptyPeriod(config));
    const w2Days = (p.w2Days || []).map(d => this._convertApiPeriod(d, config));
    while (w2Days.length < 7) w2Days.push(this.emptyPeriod(config));
    const w3Days = (p.w3Days || []).map(d => this._convertApiPeriod(d, config));
    while (w3Days.length < 7) w3Days.push(this.emptyPeriod(config));
    const w4Days = (p.w4Days || []).map(d => this._convertApiPeriod(d, config));
    while (w4Days.length < 7) w4Days.push(this.emptyPeriod(config));
    const w5Days = (p.w5Days || []).map(d => this._convertApiPeriod(d, config));
    while (w5Days.length < 7) w5Days.push(this.emptyPeriod(config));

    const lw = this._convertApiPeriod(p.priorWeek, config);
    const w2 = this._convertApiPeriod(p.twoWkPrior, config);
    const w3 = this._convertApiPeriod(p.threeWkPrior, config);
    const w4 = this._convertApiPeriod(p.fourWkPrior, config);
    const w5 = this._convertApiPeriod(p.fiveWkPrior, config);

    const accum = {
      name: p.name,
      role: OFFICE_CONFIG.roles[p.rank]?.label || p.rank,
      _roleKey: p.rank || 'rep',
      team: p.team || 'Unassigned',
      days,
      lwDays,
      w2Days,
      w3Days,
      w4Days,
      w5Days,
      lw,
      w2,
      w3,
      w4,
      w5,
      _recentTime: p.recentTime || new Array(config.timeSlots.length).fill(0),
      _fw4Time: p.fw4Time || new Array(config.timeSlots.length).fill(0),
      _autoAdded: false
    };

    const person = this._finalizePerson(accum, config);
    person.email = p.email || '';
    return person;
  },

  // ──────────────────────────────────────────────
  // ROSTER PARSING (for direct Sheets API)
  // ──────────────────────────────────────────────
  parseRoster(rosterData, config) {
    const rosterMap = {};  // name → { role, team, active }

    rosterData.data.forEach(row => {
      const name = String(row[ROSTER_COLUMNS.name] || '').trim();
      if (!name) return;

      rosterMap[name] = {
        name,
        role:   String(row[ROSTER_COLUMNS.role] || 'rep').toLowerCase(),
        team:   String(row[ROSTER_COLUMNS.team] || 'Unassigned').trim(),
        active: String(row[ROSTER_COLUMNS.active] || 'Yes').toLowerCase() !== 'no'
      };
    });

    return rosterMap;
  },

  // ──────────────────────────────────────────────
  // DATE HELPERS
  // ──────────────────────────────────────────────
  getWeekStart(date) {
    const d = new Date(date);
    const day = d.getDay();
    const diff = d.getDate() - day + (day === 0 ? -6 : 1); // Monday
    const monday = new Date(d);
    monday.setDate(diff);
    monday.setHours(0, 0, 0, 0);
    return monday;
  },

  getCurrentWeekStart() {
    return this.getWeekStart(new Date());
  },

  getDayIndex(date, weekStart) {
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    const diff = Math.floor((d - weekStart) / (24 * 60 * 60 * 1000));
    return diff >= 0 && diff < 7 ? diff : -1;
  },

  getWeekOffset(date, currentWeekStart) {
    const saleWeekStart = this.getWeekStart(date);
    const diff = Math.floor((currentWeekStart - saleWeekStart) / (7 * 24 * 60 * 60 * 1000));
    return diff; // 0 = this week, 1 = last week, 2 = 2 weeks ago, 3 = 3 weeks ago
  },

  parseDate(dateStr) {
    if (!dateStr) return null;
    // Handle various date formats
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? null : d;
  },

  parseTime(timestampStr) {
    if (!timestampStr) return null;
    const d = new Date(timestampStr);
    if (isNaN(d.getTime())) return null;
    return d.getHours() + d.getMinutes() / 60; // e.g., 14.5 = 2:30 PM
  },

  // ──────────────────────────────────────────────
  // PRODUCT EXTRACTION from a single sheet row
  // ──────────────────────────────────────────────
  extractProducts(row, config) {
    const products = {};
    let yeses = 0;
    let units = 0;

    config.columns.products.forEach(prod => {
      let val = 0;

      if (prod.type === 'boolean') {
        const raw = String(row[prod.column] || '').trim().toLowerCase();
        val = (raw === 'yes' || raw === 'y' || raw === 'true' || raw === '1') ? 1 : 0;
      } else if (prod.type === 'quantity') {
        val = parseInt(row[prod.column], 10) || 0;
      } else if (prod.type === 'sum') {
        val = (prod.columns || []).reduce((sum, col) => {
          return sum + (parseInt(row[col], 10) || 0);
        }, 0);
      }

      products[prod.key] = val;
      if (val > 0) yeses++;
      units += val;
    });

    return { products, yeses, units };
  },

  // ──────────────────────────────────────────────
  // CREATE EMPTY PERIOD
  // ──────────────────────────────────────────────
  emptyPeriod(config) {
    const p = { y: 0, units: 0, products: {} };
    config.columns.products.forEach(prod => { p.products[prod.key] = 0; });
    return p;
  },

  // ──────────────────────────────────────────────
  // ADD TO PERIOD
  // ──────────────────────────────────────────────
  addToPeriod(period, extracted) {
    period.y += extracted.yeses;
    period.units += extracted.units;
    Object.keys(extracted.products).forEach(key => {
      period.products[key] = (period.products[key] || 0) + extracted.products[key];
    });
  },

  // ──────────────────────────────────────────────
  // SUM PERIODS
  // ──────────────────────────────────────────────
  sumPeriods(periods, config) {
    const result = this.emptyPeriod(config);
    periods.forEach(p => {
      result.y += p.y;
      result.units += p.units;
      config.columns.products.forEach(prod => {
        result.products[prod.key] += p.products[prod.key] || 0;
      });
    });
    return result;
  },

  // ──────────────────────────────────────────────
  // TIME SLOT BUCKETING
  // ──────────────────────────────────────────────
  getTimeSlotIndex(decimalTime, config) {
    if (decimalTime === null) return -1;
    for (let i = 0; i < config.timeRanges.length; i++) {
      if (decimalTime >= config.timeRanges[i].start && decimalTime < config.timeRanges[i].end) {
        return i;
      }
    }
    return -1;
  },

  // ──────────────────────────────────────────────
  // AGGREGATE SALES → PEOPLE
  // ──────────────────────────────────────────────
  aggregateSales(salesData, rosterMap, config) {
    const currentWeekStart = this.getCurrentWeekStart();
    const peopleMap = {}; // name → accumulation structure

    // Initialize all roster members
    Object.values(rosterMap).forEach(r => {
      if (!r.active) return;
      peopleMap[r.name] = this._initPersonAccum(r, config);
    });

    // Process each sale row
    salesData.data.forEach(row => {
      const repName = String(row[config.columns.repName] || '').trim();
      if (!repName) return;

      // Auto-add: if rep is in sales but not in roster
      if (!peopleMap[repName]) {
        const rosterEntry = rosterMap[repName] || {
          name: repName, role: 'rep', team: 'Unassigned', pin: '', active: true
        };
        if (!rosterMap[repName]) {
          rosterMap[repName] = rosterEntry; // Track new additions
          rosterEntry._autoAdded = true;
        }
        if (!rosterEntry.active) return;
        peopleMap[repName] = this._initPersonAccum(rosterEntry, config);
      }

      const accum = peopleMap[repName];
      const saleDate = this.parseDate(row[config.columns.dateOfSale]);
      if (!saleDate) return;

      const extracted = this.extractProducts(row, config);
      const weekOffset = this.getWeekOffset(saleDate, currentWeekStart);

      // Bucket into the right period
      if (weekOffset === 0) {
        const dayIdx = this.getDayIndex(saleDate, currentWeekStart);
        if (dayIdx >= 0 && dayIdx < 7) {
          this.addToPeriod(accum.days[dayIdx], extracted);
        }
      } else if (weekOffset === 1) {
        this.addToPeriod(accum.lw, extracted);
        const lwIdx = this.getDayIndex(saleDate, new Date(currentWeekStart.getTime() - 7 * 86400000));
        if (lwIdx >= 0 && lwIdx < 7) this.addToPeriod(accum.lwDays[lwIdx], extracted);
      } else if (weekOffset === 2) {
        this.addToPeriod(accum.w2, extracted);
        const w2Idx = this.getDayIndex(saleDate, new Date(currentWeekStart.getTime() - 14 * 86400000));
        if (w2Idx >= 0 && w2Idx < 7) this.addToPeriod(accum.w2Days[w2Idx], extracted);
      } else if (weekOffset === 3) {
        this.addToPeriod(accum.w3, extracted);
        const w3Idx = this.getDayIndex(saleDate, new Date(currentWeekStart.getTime() - 21 * 86400000));
        if (w3Idx >= 0 && w3Idx < 7) this.addToPeriod(accum.w3Days[w3Idx], extracted);
      } else if (weekOffset === 4) {
        this.addToPeriod(accum.w4, extracted);
        const w4Idx = this.getDayIndex(saleDate, new Date(currentWeekStart.getTime() - 28 * 86400000));
        if (w4Idx >= 0 && w4Idx < 7) this.addToPeriod(accum.w4Days[w4Idx], extracted);
      } else if (weekOffset === 5) {
        this.addToPeriod(accum.w5, extracted);
        const w5Idx = this.getDayIndex(saleDate, new Date(currentWeekStart.getTime() - 35 * 86400000));
        if (w5Idx >= 0 && w5Idx < 7) this.addToPeriod(accum.w5Days[w5Idx], extracted);
      }

      // Time-of-sale bucketing: recent = tw+lw, fw4 = w2-w5
      const decimalTime = this.parseTime(row[config.columns.timestamp]);
      const slotIdx = this.getTimeSlotIndex(decimalTime, config);
      if (slotIdx >= 0) {
        if (weekOffset <= 1) {
          accum._recentTime[slotIdx]++;
        }
        if (weekOffset >= 2 && weekOffset <= 5) {
          accum._fw4Time[slotIdx]++;
        }
      }
    });

    // Convert accumulators to final person objects
    return Object.values(peopleMap).map(accum => this._finalizePerson(accum, config));
  },

  _initPersonAccum(rosterEntry, config) {
    return {
      name: rosterEntry.name,
      role: OFFICE_CONFIG.roles[rosterEntry.role]?.label || rosterEntry.role,
      _roleKey: rosterEntry.role,
      team: rosterEntry.team,
      days: Array.from({ length: 7 }, () => this.emptyPeriod(config)),
      lwDays: Array.from({ length: 7 }, () => this.emptyPeriod(config)),
      w2Days: Array.from({ length: 7 }, () => this.emptyPeriod(config)),
      w3Days: Array.from({ length: 7 }, () => this.emptyPeriod(config)),
      w4Days: Array.from({ length: 7 }, () => this.emptyPeriod(config)),
      w5Days: Array.from({ length: 7 }, () => this.emptyPeriod(config)),
      lw:  this.emptyPeriod(config),
      w2:  this.emptyPeriod(config),
      w3:  this.emptyPeriod(config),
      w4:  this.emptyPeriod(config),
      w5:  this.emptyPeriod(config),
      _recentTime: new Array(config.timeSlots.length).fill(0),
      _fw4Time:    new Array(config.timeSlots.length).fill(0),
      _autoAdded:  rosterEntry._autoAdded || false
    };
  },

  _finalizePerson(accum, config) {
    const days = accum.days;
    const lw = accum.lw;
    const w2 = accum.w2;
    const w3 = accum.w3;
    const w4 = accum.w4 || this.emptyPeriod(config);
    const w5 = accum.w5 || this.emptyPeriod(config);

    // This week totals
    const twPeriod = this.sumPeriods(days, config);
    const twUnits = twPeriod.units;
    const twYeses = twPeriod.y;

    // Recent = tw + lw (past ~14 days)
    const recentPeriod = this.sumPeriods([...days, lw], config);
    // 4Wk running = w2 + w3 + w4 + w5 (days 15–42)
    const fw4Period = this.sumPeriods([w2, w3, w4, w5], config);
    // All-time total for this window
    const totalPeriod = this.sumPeriods([...days, lw, w2, w3, w4, w5], config);
    const totalActs = Math.max(totalPeriod.units, 1);

    // Averages and trends (recent = 2wk avg, fourWk = 4wk avg from w2-w5)
    const recentAvg = parseFloat((recentPeriod.units / 2).toFixed(2));
    const fourWkAvg = parseFloat((fw4Period.units / 4).toFixed(2));
    const vsPct = fourWkAvg > 0
      ? parseFloat(((recentAvg - fourWkAvg) / fourWkAvg * 100).toFixed(1))
      : 0;

    // Remarks
    let remark, remarkColor;
    if (vsPct >= 20)      { remark = 'Strong upward trend'; remarkColor = '#22c55e'; }
    else if (vsPct >= 5)  { remark = 'Slight improvement';  remarkColor = '#86efac'; }
    else if (vsPct >= -5) { remark = 'Holding steady';      remarkColor = '#f0b429'; }
    else if (vsPct >= -20){ remark = 'Meaningful decline';   remarkColor = '#f97316'; }
    else                  { remark = 'Significant drop';     remarkColor = '#e53535'; }

    // Churn buckets — N/A placeholders (requires order status data)
    const churnBuckets = [
      { label: '0–30 Day',   activated: 0, disco: 0, pct: 'N/A' },
      { label: '30–60 Day',  activated: 0, disco: 0, pct: 'N/A' },
      { label: '60–90 Day',  activated: 0, disco: 0, pct: 'N/A' },
      { label: '90–120 Day', activated: 0, disco: 0, pct: 'N/A' },
      { label: '120–150 Day',activated: 0, disco: 0, pct: 'N/A' }
    ];

    // Sales per day: 2-week avg per day-of-week
    // Days before today: avg(thisWeek[day], lastWeek[day])
    // Today and future days: avg(lastWeek[day], w2[day])
    const lwDays = accum.lwDays || Array.from({ length: 7 }, () => this.emptyPeriod(config));
    const w2Days = accum.w2Days || Array.from({ length: 7 }, () => this.emptyPeriod(config));
    const w3Days = accum.w3Days || Array.from({ length: 7 }, () => this.emptyPeriod(config));
    const w4Days = accum.w4Days || Array.from({ length: 7 }, () => this.emptyPeriod(config));
    const w5Days = accum.w5Days || Array.from({ length: 7 }, () => this.emptyPeriod(config));
    const todayDow = (new Date().getDay() + 6) % 7; // 0=Mon, 6=Sun
    const salesPerDay = days.map((d, i) => {
      if (i < todayDow) {
        // Day already happened this week: avg of this week + last week
        return parseFloat(((d.units + lwDays[i].units) / 2).toFixed(1));
      } else {
        // Today or hasn't happened yet: avg of last week + 2 weeks ago
        return parseFloat(((lwDays[i].units + w2Days[i].units) / 2).toFixed(1));
      }
    });
    // 4Wk Avg line: per-day-of-week avg from w2+w3+w4+w5
    const fourWkDaily = days.map((d, i) => {
      return parseFloat(((w2Days[i].units + w3Days[i].units + w4Days[i].units + w5Days[i].units) / 4).toFixed(1));
    });
    const alpha = salesPerDay.map((v, i) => parseFloat((v - fourWkDaily[i]).toFixed(1)));

    // Product labels and data
    const prodLabels = config.columns.products.map(p => p.label);
    const recentProds = config.columns.products.map(p => recentPeriod.products[p.key] || 0);
    const fw4Prods = config.columns.products.map(p => fw4Period.products[p.key] || 0);

    // Active/pending/cancel — N/A without order data, use estimates
    const active = totalActs;
    const pending = 0;
    const cancel = 0;
    const projDisco = 0;

    return {
      name: accum.name,
      role: accum.role,
      _roleKey: accum._roleKey,
      team: accum.team,
      days,
      lwDays,
      w2Days,
      w3Days,
      w4Days,
      w5Days,
      lw,
      w2,
      w3,
      w4,
      w5,
      _autoAdded: accum._autoAdded,
      metrics: {
        totalActs, active, pending, cancel, projDisco,
        recentAvg, fourWkAvg, vsPct, remark, remarkColor,
        churnBuckets, salesPerDay, fourWkDaily, alpha,
        timeSlots: config.timeSlots,
        recentTime: accum._recentTime,
        fw4Time: accum._fw4Time,
        prodLabels, recentProds, fw4Prods
      }
    };
  },

  // ──────────────────────────────────────────────
  // BUILD TEAMS from people
  // ──────────────────────────────────────────────
  buildTeams(people, config) {
    const teamMap = {};

    // Initialize all configured teams
    config.teams.forEach(name => {
      teamMap[name] = { name, members: [], units: 0, y: 0 };
    });

    // Assign people to teams
    people.forEach(p => {
      const teamName = p.team || 'Unassigned';
      if (!teamMap[teamName]) {
        teamMap[teamName] = { name: teamName, members: [], units: 0, y: 0 };
      }
      teamMap[teamName].members.push(p);
    });

    // Build team metrics
    return Object.values(teamMap).map(team => {
      if (team.members.length === 0) {
        return { ...team, metrics: null };
      }

      // Aggregate member data
      const twUnits = team.members.reduce((s, p) => s + p.days.reduce((a, d) => a + d.units, 0), 0);
      const twYeses = team.members.reduce((s, p) => s + p.days.reduce((a, d) => a + d.y, 0), 0);
      team.units = twUnits;
      team.y = twYeses;

      // Sum member periods for team-level metrics
      const sumMemberPeriod = (getPeriod) => {
        const result = this.emptyPeriod(config);
        team.members.forEach(p => {
          const period = getPeriod(p);
          result.y += period.y;
          result.units += period.units;
          config.columns.products.forEach(prod => {
            result.products[prod.key] += period.products[prod.key] || 0;
          });
        });
        return result;
      };

      const tw = sumMemberPeriod(p => this.sumPeriods(p.days, config));
      const lw = sumMemberPeriod(p => p.lw);
      const w2 = sumMemberPeriod(p => p.w2);
      const w3 = sumMemberPeriod(p => p.w3);
      const w4 = sumMemberPeriod(p => p.w4 || this.emptyPeriod(config));
      const w5 = sumMemberPeriod(p => p.w5 || this.emptyPeriod(config));

      // Recent = tw + lw, 4Wk running = w2+w3+w4+w5
      const recentPeriod = this.sumPeriods([tw, lw], config);
      const fw4Period = this.sumPeriods([w2, w3, w4, w5], config);
      const total = this.sumPeriods([tw, lw, w2, w3, w4, w5], config);
      const totalActs = Math.max(total.units, 1);

      const recentAvg = parseFloat((recentPeriod.units / 2).toFixed(2));
      const fourWkAvg = parseFloat((fw4Period.units / 4).toFixed(2));
      const vsPct = fourWkAvg > 0
        ? parseFloat(((recentAvg - fourWkAvg) / fourWkAvg * 100).toFixed(1))
        : 0;

      let remark, remarkColor;
      if (vsPct >= 20)      { remark = 'Strong upward trend'; remarkColor = '#22c55e'; }
      else if (vsPct >= 5)  { remark = 'Slight improvement';  remarkColor = '#86efac'; }
      else if (vsPct >= -5) { remark = 'Holding steady';      remarkColor = '#f0b429'; }
      else if (vsPct >= -20){ remark = 'Meaningful decline';   remarkColor = '#f97316'; }
      else                  { remark = 'Significant drop';     remarkColor = '#e53535'; }

      const churnBuckets = [
        { label: '0–30 Day',   activated: 0, disco: 0, pct: 'N/A' },
        { label: '30–60 Day',  activated: 0, disco: 0, pct: 'N/A' },
        { label: '60–90 Day',  activated: 0, disco: 0, pct: 'N/A' },
        { label: '90–120 Day', activated: 0, disco: 0, pct: 'N/A' },
        { label: '120–150 Day',activated: 0, disco: 0, pct: 'N/A' }
      ];

      // Sales per day: 2-week avg per day-of-week for team
      const todayDow = (new Date().getDay() + 6) % 7;
      const salesPerDay = Array.from({ length: 7 }, (_, i) => {
        const twDay = team.members.reduce((s, p) => s + p.days[i].units, 0);
        const lwDay = team.members.reduce((s, p) => s + (p.lwDays ? p.lwDays[i].units : 0), 0);
        const w2Day = team.members.reduce((s, p) => s + (p.w2Days ? p.w2Days[i].units : 0), 0);
        if (i < todayDow) return parseFloat(((twDay + lwDay) / 2).toFixed(1));
        return parseFloat(((lwDay + w2Day) / 2).toFixed(1));
      });
      // 4Wk Avg line: per-day-of-week avg from w2+w3+w4+w5
      const fourWkDaily = Array.from({ length: 7 }, (_, i) => {
        const w2Day = team.members.reduce((s, p) => s + (p.w2Days ? p.w2Days[i].units : 0), 0);
        const w3Day = team.members.reduce((s, p) => s + (p.w3Days ? p.w3Days[i].units : 0), 0);
        const w4Day = team.members.reduce((s, p) => s + (p.w4Days ? p.w4Days[i].units : 0), 0);
        const w5Day = team.members.reduce((s, p) => s + (p.w5Days ? p.w5Days[i].units : 0), 0);
        return parseFloat(((w2Day + w3Day + w4Day + w5Day) / 4).toFixed(1));
      });
      const alpha = salesPerDay.map((v, i) => parseFloat((v - fourWkDaily[i]).toFixed(1)));

      const recentTime = team.members.reduce((acc, p) =>
        acc.map((v, i) => v + (p.metrics.recentTime[i] || 0)),
        new Array(config.timeSlots.length).fill(0)
      );
      const fw4Time = team.members.reduce((acc, p) =>
        acc.map((v, i) => v + (p.metrics.fw4Time[i] || 0)),
        new Array(config.timeSlots.length).fill(0)
      );

      const prodLabels = config.columns.products.map(p => p.label);
      const recentProds = config.columns.products.map(p => recentPeriod.products[p.key] || 0);
      const fw4Prods = config.columns.products.map(p => fw4Period.products[p.key] || 0);

      team.metrics = {
        totalActs,
        active: totalActs,
        pending: 0,
        cancel: 0,
        projDisco: 0,
        recentAvg, fourWkAvg, vsPct, remark, remarkColor,
        churnBuckets, salesPerDay, fourWkDaily, alpha,
        timeSlots: config.timeSlots,
        recentTime, fw4Time, prodLabels, recentProds, fw4Prods
      };

      return team;
    });
  },

  // ──────────────────────────────────────────────
  // BUILD TEAM HIERARCHY from _Teams sheet data
  // ──────────────────────────────────────────────
  buildTeamHierarchy(people, teamsData, config) {
    // Step 1: Build lookup maps
    const defs = {};       // teamId → definition
    const nameToId = {};   // team name → teamId

    Object.values(teamsData).forEach(td => {
      defs[td.teamId] = { ...td, children: [] };
      nameToId[td.name] = td.teamId;
    });

    // Step 2: Link parent-child
    Object.values(defs).forEach(d => {
      if (d.parentId && defs[d.parentId]) {
        defs[d.parentId].children.push(d.teamId);
      }
    });

    // Step 3: Compute all descendant IDs (recursive)
    const getDescendants = (teamId) => {
      const desc = [];
      const d = defs[teamId];
      if (!d) return desc;
      d.children.forEach(cId => {
        desc.push(cId);
        desc.push(...getDescendants(cId));
      });
      return desc;
    };
    Object.keys(defs).forEach(id => {
      defs[id].allDescendantIds = getDescendants(id);
    });

    // Step 4: Assign people to their DIRECT team
    const directMembers = {};  // teamId → [person, ...]
    Object.keys(defs).forEach(id => { directMembers[id] = []; });

    const unassigned = [];
    people.forEach(p => {
      const teamName = p.team || 'Unassigned';
      const teamId = nameToId[teamName];
      if (teamId && directMembers[teamId]) {
        directMembers[teamId].push(p);
      } else {
        unassigned.push(p);
      }
    });

    // Step 5: Build team objects with combined members and metrics
    const teams = Object.values(defs).map(d => {
      const direct = directMembers[d.teamId] || [];
      const descendant = d.allDescendantIds.flatMap(dId => directMembers[dId] || []);
      const combined = [...direct, ...descendant];

      const team = {
        name: d.name,
        teamId: d.teamId,
        parentId: d.parentId || '',
        leaderId: d.leaderId || '',
        emoji: d.emoji || '',
        children: d.children,
        allDescendantIds: d.allDescendantIds,
        directMembers: direct,
        members: combined,
        isSubTeam: !!d.parentId,
        units: 0,
        y: 0
      };

      // Use existing metrics builder logic
      this._buildTeamMetrics(team, config);
      return team;
    });

    // Add Unassigned if needed
    if (unassigned.length > 0) {
      const uTeam = {
        name: 'Unassigned',
        teamId: '_unassigned',
        parentId: '',
        leaderId: '',
        emoji: '',
        children: [],
        allDescendantIds: [],
        directMembers: unassigned,
        members: unassigned,
        isSubTeam: false,
        units: 0,
        y: 0
      };
      this._buildTeamMetrics(uTeam, config);
      teams.push(uTeam);
    }

    return teams;
  },

  // Shared team metrics builder (used by both buildTeams and buildTeamHierarchy)
  _buildTeamMetrics(team, config) {
    if (team.members.length === 0) {
      team.metrics = null;
      return;
    }

    const twUnits = team.members.reduce((s, p) => s + p.days.reduce((a, d) => a + d.units, 0), 0);
    const twYeses = team.members.reduce((s, p) => s + p.days.reduce((a, d) => a + d.y, 0), 0);
    team.units = twUnits;
    team.y = twYeses;

    const sumMemberPeriod = (getPeriod) => {
      const result = this.emptyPeriod(config);
      team.members.forEach(p => {
        const period = getPeriod(p);
        result.y += period.y;
        result.units += period.units;
        config.columns.products.forEach(prod => {
          result.products[prod.key] += period.products[prod.key] || 0;
        });
      });
      return result;
    };

    const tw = sumMemberPeriod(p => this.sumPeriods(p.days, config));
    const lw = sumMemberPeriod(p => p.lw);
    const w2 = sumMemberPeriod(p => p.w2);
    const w3 = sumMemberPeriod(p => p.w3);
    const w4 = sumMemberPeriod(p => p.w4 || this.emptyPeriod(config));
    const w5 = sumMemberPeriod(p => p.w5 || this.emptyPeriod(config));

    // Recent = tw + lw, 4Wk running = w2+w3+w4+w5
    const recentPeriod = this.sumPeriods([tw, lw], config);
    const fw4Period = this.sumPeriods([w2, w3, w4, w5], config);
    const total = this.sumPeriods([tw, lw, w2, w3, w4, w5], config);
    const totalActs = Math.max(total.units, 1);

    const recentAvg = parseFloat((recentPeriod.units / 2).toFixed(2));
    const fourWkAvg = parseFloat((fw4Period.units / 4).toFixed(2));
    const vsPct = fourWkAvg > 0
      ? parseFloat(((recentAvg - fourWkAvg) / fourWkAvg * 100).toFixed(1))
      : 0;

    let remark, remarkColor;
    if (vsPct >= 20)      { remark = 'Strong upward trend'; remarkColor = '#22c55e'; }
    else if (vsPct >= 5)  { remark = 'Slight improvement';  remarkColor = '#86efac'; }
    else if (vsPct >= -5) { remark = 'Holding steady';      remarkColor = '#f0b429'; }
    else if (vsPct >= -20){ remark = 'Meaningful decline';   remarkColor = '#f97316'; }
    else                  { remark = 'Significant drop';     remarkColor = '#e53535'; }

    const churnBuckets = [
      { label: '0–30 Day',   activated: 0, disco: 0, pct: 'N/A' },
      { label: '30–60 Day',  activated: 0, disco: 0, pct: 'N/A' },
      { label: '60–90 Day',  activated: 0, disco: 0, pct: 'N/A' },
      { label: '90–120 Day', activated: 0, disco: 0, pct: 'N/A' },
      { label: '120–150 Day',activated: 0, disco: 0, pct: 'N/A' }
    ];

    // Sales per day: 2-week avg per day-of-week for team
    const todayDow = (new Date().getDay() + 6) % 7;
    const salesPerDay = Array.from({ length: 7 }, (_, i) => {
      const twDay = team.members.reduce((s, p) => s + p.days[i].units, 0);
      const lwDay = team.members.reduce((s, p) => s + (p.lwDays ? p.lwDays[i].units : 0), 0);
      const w2Day = team.members.reduce((s, p) => s + (p.w2Days ? p.w2Days[i].units : 0), 0);
      if (i < todayDow) return parseFloat(((twDay + lwDay) / 2).toFixed(1));
      return parseFloat(((lwDay + w2Day) / 2).toFixed(1));
    });
    // 4Wk Avg line: per-day-of-week avg from w2+w3+w4+w5
    const fourWkDaily = Array.from({ length: 7 }, (_, i) => {
      const w2Day = team.members.reduce((s, p) => s + (p.w2Days ? p.w2Days[i].units : 0), 0);
      const w3Day = team.members.reduce((s, p) => s + (p.w3Days ? p.w3Days[i].units : 0), 0);
      const w4Day = team.members.reduce((s, p) => s + (p.w4Days ? p.w4Days[i].units : 0), 0);
      const w5Day = team.members.reduce((s, p) => s + (p.w5Days ? p.w5Days[i].units : 0), 0);
      return parseFloat(((w2Day + w3Day + w4Day + w5Day) / 4).toFixed(1));
    });
    const alpha = salesPerDay.map((v, i) => parseFloat((v - fourWkDaily[i]).toFixed(1)));

    const recentTime = team.members.reduce((acc, p) =>
      acc.map((v, i) => v + (p.metrics.recentTime[i] || 0)),
      new Array(config.timeSlots.length).fill(0)
    );
    const fw4Time = team.members.reduce((acc, p) =>
      acc.map((v, i) => v + (p.metrics.fw4Time[i] || 0)),
      new Array(config.timeSlots.length).fill(0)
    );

    const prodLabels = config.columns.products.map(p => p.label);
    const recentProds = config.columns.products.map(p => recentPeriod.products[p.key] || 0);
    const fw4Prods = config.columns.products.map(p => fw4Period.products[p.key] || 0);

    team.metrics = {
      totalActs,
      active: totalActs,
      pending: 0,
      cancel: 0,
      projDisco: 0,
      recentAvg, fourWkAvg, vsPct, remark, remarkColor,
      churnBuckets, salesPerDay, fourWkDaily, alpha,
      timeSlots: config.timeSlots,
      recentTime, fw4Time, prodLabels, recentProds, fw4Prods
    };
  },

  // ──────────────────────────────────────────────
  // TABLEAU ENRICHMENT — DTR status classification
  // ──────────────────────────────────────────────
  _ACTIVE_STATUSES: ['Posted', 'Delivered', 'Confirmed'],
  _PENDING_STATUSES: ['Open', 'Pending', 'Scheduled', 'Shipped', 'Port Approved', 'BYOD', 'Backordered'],
  _CANCEL_STATUSES: ['Canceled'],
  _DISCO_STATUSES: ['Disconnected'],

  _classifyStatusCounts(statusCounts) {
    let active = 0, pending = 0, cancel = 0, disco = 0, total = 0;
    Object.entries(statusCounts || {}).forEach(([status, count]) => {
      total += count;
      if (this._ACTIVE_STATUSES.includes(status)) active += count;
      else if (this._PENDING_STATUSES.includes(status)) pending += count;
      else if (this._CANCEL_STATUSES.includes(status)) cancel += count;
      else if (this._DISCO_STATUSES.includes(status)) disco += count;
      else pending += count; // unknown statuses default to pending
    });
    return { active, pending, cancel, disco, total };
  },

  // Enrich person metrics with Tableau rep-level data
  enrichWithTableau(people, tableauByRep) {
    if (!tableauByRep || Object.keys(tableauByRep).length === 0) return;

    people.forEach(p => {
      const email = p.email;
      if (!email) return;
      const rep = tableauByRep[email];
      if (!rep) return;

      const classified = this._classifyStatusCounts(rep.statusCounts);
      const m = p.metrics;
      m.active = classified.active;
      m.pending = classified.pending;
      m.cancel = classified.cancel;
      m.projDisco = classified.disco;
      m.totalDevices = classified.total;
      m.activationRate = classified.total > 0
        ? parseFloat((classified.active / classified.total * 100).toFixed(1))
        : 0;
      m.productBreakdown = rep.productCounts || {};
      m.tableauName = rep.tableauName || '';
    });
  },

  // Enrich team metrics by aggregating member Tableau data
  enrichTeamsWithTableau(teams) {
    if (!teams) return;

    teams.forEach(team => {
      if (!team.metrics || !team.members || team.members.length === 0) return;

      let active = 0, pending = 0, cancel = 0, disco = 0, totalDevices = 0;
      const productBreakdown = {};

      team.members.forEach(p => {
        const m = p.metrics;
        if (m.totalDevices > 0) {
          active += m.active;
          pending += m.pending;
          cancel += m.cancel;
          disco += m.projDisco;
          totalDevices += m.totalDevices;
          Object.entries(m.productBreakdown || {}).forEach(([pt, count]) => {
            productBreakdown[pt] = (productBreakdown[pt] || 0) + count;
          });
        }
      });

      if (totalDevices > 0) {
        team.metrics.active = active;
        team.metrics.pending = pending;
        team.metrics.cancel = cancel;
        team.metrics.projDisco = disco;
        team.metrics.totalDevices = totalDevices;
        team.metrics.activationRate = parseFloat((active / totalDevices * 100).toFixed(1));
        team.metrics.productBreakdown = productBreakdown;
      }
    });
  }
};
