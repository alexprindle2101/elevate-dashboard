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
          phone: info.phone || '',
          tableauName: info.tableauName || ''
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

  // Rolling averages from per-day arrays.
  // Rec Wk Avg = units from past 14 days (excluding today) ÷ 2
  // 4Wk Avg    = units from days 15–42 before today ÷ 4
  // Accepts flat array of day objects (oldest first, 42 entries = 6 weeks)
  _rollingAvgs(allDays) {
    const todayDow = (new Date().getDay() + 6) % 7; // Mon=0 .. Sun=6
    const todayIdx = 35 + todayDow; // position of today in the 42-entry array

    // Rec Wk Avg: days [todayIdx-14 .. todayIdx-1]
    let recentUnits = 0;
    for (let i = todayIdx - 14; i <= todayIdx - 1; i++) {
      if (i >= 0 && i < allDays.length) recentUnits += (allDays[i].units || 0);
    }
    const recentAvg = parseFloat((recentUnits / 2).toFixed(2));

    // 4Wk Avg: days [todayIdx-42 .. todayIdx-15]
    let fw4Units = 0;
    for (let i = todayIdx - 42; i <= todayIdx - 15; i++) {
      if (i >= 0 && i < allDays.length) fw4Units += (allDays[i].units || 0);
    }
    const fourWkAvg = parseFloat((fw4Units / 4).toFixed(2));

    const vsPct = fourWkAvg > 0
      ? parseFloat(((recentAvg - fourWkAvg) / fourWkAvg * 100).toFixed(1))
      : 0;

    return { recentAvg, fourWkAvg, vsPct };
  },

  // Build flat chronological day array from 6-week per-day arrays (oldest first)
  _flattenDays(dayArrays) {
    // dayArrays = { w5Days, w4Days, w3Days, w2Days, lwDays, days }
    return [
      ...(dayArrays.w5Days || []),
      ...(dayArrays.w4Days || []),
      ...(dayArrays.w3Days || []),
      ...(dayArrays.w2Days || []),
      ...(dayArrays.lwDays || []),
      ...(dayArrays.days || [])
    ];
  },

  // Aggregate team members' per-day arrays into a single flat 42-entry array
  _buildTeamDayArray(members, config) {
    const result = Array.from({ length: 42 }, () => ({ units: 0 }));
    members.forEach(p => {
      const flat = this._flattenDays(p);
      for (let i = 0; i < 42 && i < flat.length; i++) {
        result[i].units += (flat[i].units || 0);
      }
    });
    return result;
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

    // All-time total for this window
    const totalPeriod = this.sumPeriods([...days, lw, w2, w3, w4, w5], config);
    const totalActs = Math.max(totalPeriod.units, 1);

    // Rolling averages from per-day data (14-day recent, 28-day prior)
    const allDays = this._flattenDays(accum);
    const { recentAvg, fourWkAvg, vsPct } = this._rollingAvgs(allDays);

    // Remarks
    let remark, remarkColor;
    if (vsPct >= 20)      { remark = 'Strong upward trend'; remarkColor = '#22c55e'; }
    else if (vsPct >= 5)  { remark = 'Slight improvement';  remarkColor = '#86efac'; }
    else if (vsPct >= -5) { remark = 'Holding steady';      remarkColor = '#f0b429'; }
    else if (vsPct >= -20){ remark = 'Meaningful decline';   remarkColor = '#f97316'; }
    else                  { remark = 'Significant drop';     remarkColor = '#e53535'; }

    // Churn buckets — defaults (enriched from _TableauChurnReport when available)
    const churnBuckets = [
      { label: '0-30 Day', activated: 0, disco: 0, pct: 'N/A' },
      { label: '30 Day',   activated: 0, disco: 0, pct: 'N/A' },
      { label: '60 Day',   activated: 0, disco: 0, pct: 'N/A' },
      { label: '90 Day',   activated: 0, disco: 0, pct: 'N/A' },
      { label: '120 Day',  activated: 0, disco: 0, pct: 'N/A' }
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

    // Product labels and data (recent = tw+lw, 4wk = w2+w3+w4+w5)
    const recentPeriod = this.sumPeriods([twPeriod, lw], config);
    const fw4Period = this.sumPeriods([w2, w3, w4, w5], config);
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

      const total = this.sumPeriods([tw, lw, w2, w3, w4, w5], config);
      const totalActs = Math.max(total.units, 1);

      // Build team-level per-day arrays for rolling avg calc
      const teamAllDays = this._buildTeamDayArray(team.members, config);
      const { recentAvg, fourWkAvg, vsPct } = this._rollingAvgs(teamAllDays);

      let remark, remarkColor;
      if (vsPct >= 20)      { remark = 'Strong upward trend'; remarkColor = '#22c55e'; }
      else if (vsPct >= 5)  { remark = 'Slight improvement';  remarkColor = '#86efac'; }
      else if (vsPct >= -5) { remark = 'Holding steady';      remarkColor = '#f0b429'; }
      else if (vsPct >= -20){ remark = 'Meaningful decline';   remarkColor = '#f97316'; }
      else                  { remark = 'Significant drop';     remarkColor = '#e53535'; }

      const churnBuckets = [
        { label: '0-30 Day', activated: 0, disco: 0, pct: 'N/A' },
        { label: '30 Day',   activated: 0, disco: 0, pct: 'N/A' },
        { label: '60 Day',   activated: 0, disco: 0, pct: 'N/A' },
        { label: '90 Day',   activated: 0, disco: 0, pct: 'N/A' },
        { label: '120 Day',  activated: 0, disco: 0, pct: 'N/A' }
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

      const recentPeriod = this.sumPeriods([tw, lw], config);
      const fw4Period = this.sumPeriods([w2, w3, w4, w5], config);
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

    const total = this.sumPeriods([tw, lw, w2, w3, w4, w5], config);
    const totalActs = Math.max(total.units, 1);

    // Build team-level per-day arrays for rolling avg calc
    const teamAllDays = this._buildTeamDayArray(team.members, config);
    const { recentAvg, fourWkAvg, vsPct } = this._rollingAvgs(teamAllDays);

    let remark, remarkColor;
    if (vsPct >= 20)      { remark = 'Strong upward trend'; remarkColor = '#22c55e'; }
    else if (vsPct >= 5)  { remark = 'Slight improvement';  remarkColor = '#86efac'; }
    else if (vsPct >= -5) { remark = 'Holding steady';      remarkColor = '#f0b429'; }
    else if (vsPct >= -20){ remark = 'Meaningful decline';   remarkColor = '#f97316'; }
    else                  { remark = 'Significant drop';     remarkColor = '#e53535'; }

    const churnBuckets = [
      { label: '0-30 Day', activated: 0, disco: 0, pct: 'N/A' },
      { label: '30 Day',   activated: 0, disco: 0, pct: 'N/A' },
      { label: '60 Day',   activated: 0, disco: 0, pct: 'N/A' },
      { label: '90 Day',   activated: 0, disco: 0, pct: 'N/A' },
      { label: '120 Day',  activated: 0, disco: 0, pct: 'N/A' }
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

    const recentPeriod = this.sumPeriods([tw, lw], config);
    const fw4Period = this.sumPeriods([w2, w3, w4, w5], config);
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
  enrichWithTableau(people, tableauByRep, tableauByName, roster) {
    if (!tableauByRep || Object.keys(tableauByRep).length === 0) return;

    people.forEach(p => {
      const email = p.email;
      if (!email) return;

      // If the rep has a stored tableauName (from roster picker), use name-keyed lookup
      const storedName = (roster && roster[email] && roster[email].tableauName) || '';
      let rep;
      if (storedName && tableauByName && tableauByName[storedName]) {
        rep = tableauByName[storedName];
      } else {
        rep = tableauByRep[email];
      }
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
      // Bonus tier data from Tableau
      p.bonusTier = rep.bonusTier || '';
      p.payoutReason = rep.payoutReason || '';
      // Wireless SPE counts in 30-day window
      m.monthTotalSPEs = rep.monthTotalSPEs || 0;
      m.monthApprovedSPEs = rep.monthApprovedSPEs || 0;
      m.monthPendingSPEs = rep.monthPendingSPEs || 0;
      m.monthCanceledSPEs = rep.monthCanceledSPEs || 0;
      m.monthDiscoSPEs = rep.monthDiscoSPEs || 0;
      // Percentages: Active/Pending/Cancel vs Total, ProjDisco vs Approved
      const tot = m.monthTotalSPEs;
      const pct = (n, d) => d > 0 ? parseFloat((n / d * 100).toFixed(1)) : 0;
      m.activePct = pct(m.monthApprovedSPEs, tot);
      m.pendingPct = pct(m.monthPendingSPEs, tot);
      m.cancelPct = pct(m.monthCanceledSPEs, tot);
      m.projDiscoPct = pct(m.monthDiscoSPEs, m.monthApprovedSPEs);
    });
  },

  // Enrich team metrics by aggregating member Tableau data
  enrichTeamsWithTableau(teams) {
    if (!teams) return;

    teams.forEach(team => {
      if (!team.metrics || !team.members || team.members.length === 0) return;

      let active = 0, pending = 0, cancel = 0, disco = 0, totalDevices = 0;
      let monthTotal = 0, monthApproved = 0, monthPending = 0, monthCanceled = 0, monthDisco = 0;
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
        // Aggregate monthly wireless SPE counts
        monthTotal += m.monthTotalSPEs || 0;
        monthApproved += m.monthApprovedSPEs || 0;
        monthPending += m.monthPendingSPEs || 0;
        monthCanceled += m.monthCanceledSPEs || 0;
        monthDisco += m.monthDiscoSPEs || 0;
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

      // Team-level monthly wireless SPE percentages
      team.metrics.monthTotalSPEs = monthTotal;
      team.metrics.monthApprovedSPEs = monthApproved;
      team.metrics.monthPendingSPEs = monthPending;
      team.metrics.monthCanceledSPEs = monthCanceled;
      team.metrics.monthDiscoSPEs = monthDisco;
      const pct = (n, d) => d > 0 ? parseFloat((n / d * 100).toFixed(1)) : 0;
      team.metrics.activePct = pct(monthApproved, monthTotal);
      team.metrics.pendingPct = pct(monthPending, monthTotal);
      team.metrics.cancelPct = pct(monthCanceled, monthTotal);
      team.metrics.projDiscoPct = pct(monthDisco, monthApproved);

      // Per-headcount averages for team context
      const activeMembers = team.members.filter(p => !p._deactivated);
      const headcount = activeMembers.length || 1;
      team.metrics.recentAvgPerHead = parseFloat((team.metrics.recentAvg / headcount).toFixed(2));
      team.metrics.fourWkAvgPerHead = parseFloat((team.metrics.fourWkAvg / headcount).toFixed(2));
    });
  },

  // Column keys from _TableauChurnReport matching bucket labels
  _CHURN_BUCKET_COLS: ['0-30 Day', '30 Day', '60 Day', '90 Day', '120 Day'],

  // Per-bucket churn color thresholds: [yellowMin, redMin]
  // < yellowMin = green, yellowMin–redMin = yellow, > redMin = red
  _CHURN_COLOR_THRESHOLDS: [
    { yellowMin: 2.5, redMin: 3.0 },   // 0-30 Day
    { yellowMin: 5.0, redMin: 7.0 },   // 30 Day
    { yellowMin: 9.0, redMin: 10.0 },  // 60 Day
    { yellowMin: 11.0, redMin: 14.0 }, // 90 Day
    { yellowMin: 14.0, redMin: 17.0 }, // 120 Day
  ],

  _getChurnBucketColor(pct, bucketIdx) {
    const t = this._CHURN_COLOR_THRESHOLDS[bucketIdx];
    if (!t) return '';
    if (pct < t.yellowMin) return 'Green';
    if (pct > t.redMin) return 'Red';
    return 'Yellow';
  },

  // Enrich person metrics with _TableauChurnReport data
  enrichWithChurnReport(people, churnReport) {
    if (!churnReport || churnReport.length === 0) {
      console.warn('[Churn] No churnReport data received (empty or missing)');
      return;
    }
    const sampleKeys = Object.keys(churnReport[0] || {});
    console.log('[Churn] Received', churnReport.length, 'rows. Keys:', sampleKeys.join(', '));

    // Dynamically find the actual header keys by partial/case-insensitive match
    const _findKey = (keys, patterns) => {
      for (const p of patterns) {
        const lp = p.toLowerCase();
        const found = keys.find(k => k.toLowerCase() === lp);
        if (found) return found;
      }
      // Partial match fallback
      for (const p of patterns) {
        const lp = p.toLowerCase();
        const found = keys.find(k => k.toLowerCase().includes(lp));
        if (found) return found;
      }
      return null;
    };

    const repNameKey = _findKey(sampleKeys, ['rep.Full Name', 'Rep', 'Full Name', 'Rep Name', 'Name']);
    const metricTypeKey = _findKey(sampleKeys, ['metricType', 'Metric Type', 'metric_type', 'Measure Names']);
    const colorKey = _findKey(sampleKeys, ['30-60 Color Churn (copy)', 'Color Churn', 'Churn Color']);

    console.log('[Churn] Resolved keys → repName:', repNameKey, '| metricType:', metricTypeKey, '| color:', colorKey);

    if (!repNameKey || !metricTypeKey) {
      console.error('[Churn] Cannot find required columns. Available keys:', sampleKeys);
      console.error('[Churn] First row sample:', JSON.stringify(churnReport[0]));
      return;
    }

    // Match bucket columns by exact match first, then substring.
    // Order matters: '0-30 Day' must not steal '30 Day', so use exact-first strategy.
    const BUCKET_NAMES = ['0-30 Day', '30 Day', '60 Day', '90 Day', '120 Day'];
    const usedKeys = new Set();
    const cols = BUCKET_NAMES.map(name => {
      // Exact match (case-insensitive)
      let found = sampleKeys.find(k => k.toLowerCase() === name.toLowerCase() && !usedKeys.has(k));
      if (!found) {
        // Substring fallback — but skip already-used keys
        found = sampleKeys.find(k => k.toLowerCase().includes(name.toLowerCase()) && !usedKeys.has(k));
      }
      if (found) usedKeys.add(found);
      return found || name;
    });
    console.log('[Churn] Bucket columns:', cols);

    // Group rows by rep name, summing activated + disco counts
    // Result: byRep[name] = { activated:[5], disco:[5], hasData:[5] }
    const byRep = {};
    churnReport.forEach(row => {
      const repName = String(row[repNameKey] || '').trim();
      const metricType = String(row[metricTypeKey] || '').trim();
      if (!repName || !metricType || repName === 'Grand Total') return;

      if (!byRep[repName]) {
        byRep[repName] = {
          activated: new Array(5).fill(0),
          disco: new Array(5).fill(0),
          hasData: new Array(5).fill(false)
        };
      }
      const rd = byRep[repName];

      cols.forEach((col, i) => {
        const val = row[col];
        if (val === undefined || val === '') return;

        if (metricType === 'Activated SPE/SP') {
          rd.activated[i] += parseInt(val) || 0;
          rd.hasData[i] = true;
        } else if (metricType === 'Disconnect count (SPE/SP)') {
          rd.disco[i] += parseInt(val) || 0;
        }
      });
    });

    const repNames = Object.keys(byRep);
    console.log('[Churn] Aggregated', repNames.length, 'reps from report:', repNames.slice(0, 5));

    let matched = 0;
    people.forEach(p => {
      // Match by tableauName (set during Tableau enrichment) since churn report
      // uses full Tableau names which differ from roster display names
      const repData = byRep[p.metrics.tableauName] || byRep[p.name];
      if (!repData) return;
      matched++;

      cols.forEach((col, i) => {
        if (!repData.hasData[i]) return;
        const bucket = p.metrics.churnBuckets[i];
        bucket.activated = repData.activated[i];
        bucket.disco = repData.disco[i];
        bucket.pct = bucket.activated > 0
          ? parseFloat((bucket.disco / bucket.activated * 100).toFixed(1))
          : 0;
        bucket.color = this._getChurnBucketColor(bucket.pct, i);
      });
    });
    console.log('[Churn] Matched', matched, '/', people.length, 'people');
  },

  // Enrich team metrics by aggregating member churn data
  enrichTeamsWithChurn(teams) {
    if (!teams) return;

    teams.forEach(team => {
      if (!team.metrics || !team.members || team.members.length === 0) return;

      team.members.forEach(p => {
        p.metrics.churnBuckets.forEach((bucket, i) => {
          if (bucket.pct === 'N/A') return;
          const tb = team.metrics.churnBuckets[i];
          tb.activated += bucket.activated;
          tb.disco += bucket.disco;
        });
      });

      // Recalculate team-level percentages and assign color from hardcoded thresholds
      team.metrics.churnBuckets.forEach((tb, i) => {
        if (tb.activated > 0) {
          tb.pct = parseFloat((tb.disco / tb.activated * 100).toFixed(1));
          tb.color = this._getChurnBucketColor(tb.pct, i);
        }
      });
    });
  }
};
