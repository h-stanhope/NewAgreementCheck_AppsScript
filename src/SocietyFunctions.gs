/** ================================
 * Society Functions
 * - Build "Society Issues" overview (sales currency view)
 * - Reuses QUAL, readers, and writer utilities
 * ================================ */

/** Entry point: builds a sheet "Society Issues" with only Investigate? = Yes rows.
 *  Values (Expected/Income/Variance) are in the SOCIETY'S SALES CURRENCY.
 */
function SOC_buildSocietyIssuesOverview() {
  const ctx = CORE_buildContext_();

  const statusesMap   = READ_readSalesforceStatuses_(ctx);
  const societyHome   = READ_readSocietyHome_(ctx);
  const chartPlays    = READ_readChartMetric_(ctx);
  const benchmarks    = READ_readBenchmarks_(ctx);
  const incomes       = READ_readIncome_(ctx);
  const fx            = SOC_FX_load_(ctx);

  const salesCcyBySociety = SOC_buildSalesCcyBySociety_(incomes);

  const outRows = [];

  statusesMap.forEach((socArr, uuid) => {
    if (!Array.isArray(socArr) || !socArr.length) return;
    const contribName = CORE_properCase_(socArr[0]?.contributorName || '');

    socArr.forEach(st => {
      const society   = CORE_canonSociety_(st.societyName || '');
      if (!society) return;

      const territory = CORE_canonCountry_(societyHome.get(society) || '');
      const regStatus = (st.status || '').toString().toLowerCase();

      // 1) Plays via the SAME engine used by the audit
      const { totalPlays, qualifiedPlays, chartSubset } = EXP_computePlaysForTerritory_({
        ctx, uuid, society, territory, chartPlays, contentByKey: null
      });

      // 2) Expected in BENCHMARK currency (sum only qualifying spins)
      const { expectedBench, benchCcy, hasBenchmark } =
        SOC_computeExpectedInSalesCcy_({ ctx, chartSubset, society, territory, benchmarks, fx });

      // 3) Income in SALES currency (+ determine target sales CCY)
      const { incomeSales, salesCcy } =
        SOC_sumIncomeInSalesCcy_({ uuid, society, territory, incomes, fx, salesCcyBySociety });

      // Decide target sales currency for this row
      const targetSalesCcy = salesCcy || salesCcyBySociety.get(society) || (ctx.config.DEFAULT_CURRENCY || 'USD');

      // 4) Convert Expected → SALES currency (use most recent FX; no period id for expected)
      let expectedSales = 0;
      if (hasBenchmark && benchCcy) {
        expectedSales = fx.convert(expectedBench, benchCcy, targetSalesCcy, null);
      }

      // 5) Variance in SALES currency
      const varianceAbs = (incomeSales || 0) - (expectedSales || 0);
      const variancePct = expectedSales > 0 ? (incomeSales / expectedSales) - 1 : '';

      // 6) Status / Investigate via GLOBAL rules
      const classified = LOGIC_classifyRow_(ctx, {
        registrationStatus: regStatus,
        hasBenchmark,
        qualifiedPlays,
        income: incomeSales,
        expected: expectedSales
      });

      if (!classified.investigate) return; // Only show issues

      outRows.push({
        contributor: contribName,
        uuid,
        society,
        territory,
        registrationStatus: regStatus,
        playsTotal: totalPlays,
        playsQualified: qualifiedPlays,
        expected: expectedSales,
        income: incomeSales,
        varianceAbs,
        variancePct,
        status: classified.status,
        investigate: classified.investigate,
        notes: classified.notes,
        currency: targetSalesCcy
      });
    });
  });

  SOC_writeSocietyIssues_(ctx, outRows);
}

/* ===============================
   Currency & FX helpers (local)
   =============================== */

/** Read FX rates sheet and build converters.
 *  Sheet: FX Rates with columns: FROM_CURRENCY_CODE | TO_CURRENCY_CODE | Concat ccy | STATEMENT_PERIOD_ID | RATE
 */
function SOC_FX_load_(ctx) {
  const sh = ctx.ss.getSheetByName(ctx.config.SHEETS.FX);
  const data = sh ? sh.getDataRange().getValues() : [];
  if (!data || data.length < 2) {
    return {
      convert: (amt, from, to, periodId) => (from && to && from.toUpperCase() === to.toUpperCase()) ? amt : amt
    };
  }

  const head = data[0].map(x => String(x || '').trim().toLowerCase());
  const rows = data.slice(1);

  const idx = {
    from: head.indexOf('from_currency_code'),
    to: head.indexOf('to_currency_code'),
    pid: head.indexOf('statement_period_id'),
    rate: head.indexOf('rate')
  };

  // Map: "USD||GBP" -> [{pid, rate}], sorted by pid asc
  const table = new Map();
  rows.forEach(r => {
    const from = CORE_str_(r[idx.from]).toUpperCase();
    const to   = CORE_str_(r[idx.to]).toUpperCase();
    const pid  = CORE_str_(r[idx.pid]);
    const rate = Number(r[idx.rate]) || 0;
    if (!from || !to || !rate) return;
    const key = `${from}||${to}`;
    if (!table.has(key)) table.set(key, []);
    table.get(key).push({ pid, rate });
  });
  table.forEach(arr => arr.sort((a,b) => (a.pid > b.pid ? 1 : a.pid < b.pid ? -1 : 0)));

  function findRate(from, to, periodId) {
    if (!from || !to) return null;
    const key = `${from.toUpperCase()}||${to.toUpperCase()}`;
    const arr = table.get(key) || [];
    if (!arr.length) return null;
    if (periodId) {
      const hit = arr.find(x => x.pid === String(periodId));
      if (hit) return hit.rate;
    }
    // most-recent fallback
    return arr[arr.length - 1].rate;
  }

  function convert(amount, from, to, periodId) {
    amount = Number(amount) || 0;
    if (!amount) return 0;
    if (!from || !to) return amount;
    const f = from.toUpperCase(), t = to.toUpperCase();
    if (f === t) return amount;

    // Try direct
    const r1 = findRate(f, t, periodId);
    if (r1) return amount * r1;

    // Try inverse
    const r2 = findRate(t, f, periodId);
    if (r2) return amount / r2;

    // No rate – return amount unchanged
    return amount;
  }

  return { convert };
}

/** Build a map of society -> most common sales currency found in income rows */
function SOC_buildSalesCcyBySociety_(incomes) {
  const map = new Map();
  const tally = new Map(); // soc -> Map<ccy, count>
  incomes.forEach(x => {
    const soc = CORE_canonSociety_(x.societyName || '');
    const ccy = (x.salesCcy || '').toUpperCase();
    if (!soc || !ccy) return;
    if (!tally.has(soc)) tally.set(soc, new Map());
    const m = tally.get(soc);
    m.set(ccy, (m.get(ccy) || 0) + 1);
  });
  tally.forEach((m, soc) => {
    let best = '', bestN = -1;
    m.forEach((n, ccy) => { if (n > bestN) { best = ccy; bestN = n; } });
    if (best) map.set(soc, best);
  });
  return map;
}

/* ==========================================
   Expected & Income (sales currency) compute
   ========================================== */

/** Expected in benchmark currency for a contributor+society+territory (from chartSubset + benchmarks). */
function SOC_computeExpectedInSalesCcy_({ ctx, chartSubset, society, territory, benchmarks, fx }) {
  const key = `${territory}||${society}`;
  const rec = benchmarks.get(key);
  if (!rec || !rec.rate || !rec.currency) {
    return { expectedBench: 0, benchCcy: '', hasBenchmark: false };
  }
  const benchRate = Number(rec.rate) || 0;
  const benchCcy  = (rec.currency || '').toUpperCase();

  let expectedBench = 0;
  for (const p of chartSubset) {
    if (!p.qualifies) continue;             // only qualifying spins
    const plays = Number(p.playCount) || 0;
    if (!plays) continue;

    let perSpin = benchRate;

    // SESSION_MUSICIAN (Non-Featured) -> divide by 4
    const type = (p.contributorType || '').toUpperCase();
    if (type === 'SESSION_MUSICIAN') {
      perSpin = perSpin / 4;
    } else {
      // Featured; divide by number of primary performers if > 1
      const nPrim = Number(p.numPrimary) || 1;
      if (nPrim > 1) perSpin = perSpin / nPrim;
    }

    expectedBench += perSpin * plays;
  }

  return { expectedBench, benchCcy, hasBenchmark: true };
}

/** Sum income in sales currency for uuid+society, limited to society home territory (if provided). */
function SOC_sumIncomeInSalesCcy_({ uuid, society, territory, incomes, fx, salesCcyBySociety }) {
  // Determine a target sales currency:
  // 1) Most common sales ccy for this society across all income lines (precomputed)
  // 2) If this contributor has income lines with a sales ccy, prefer that
  let target = salesCcyBySociety.get(society) || '';

  // Contributor-specific currency preference
  const byThis = incomes.filter(x =>
    x.uuid === uuid && CORE_canonSociety_(x.societyName || '') === society &&
    (!territory || CORE_canonCountry_(x.countryName || '') === territory)
  );

  const counts = new Map();
  byThis.forEach(x => {
    const c = (x.salesCcy || '').toUpperCase();
    if (!c) return;
    counts.set(c, (counts.get(c) || 0) + 1);
  });
  let best = '', nBest = -1;
  counts.forEach((n, c) => { if (n > nBest) { best = c; nBest = n; } });
  if (best) target = best;

  // Fallback default if still unknown
  if (!target) target = 'USD';

  // Sum converting any differing sales currency using periodId (line-by-line)
  let sum = 0;
  for (const x of byThis) {
    const amt = Number(x.amountSales) || 0;               // Column T
    if (!amt) continue;
    const from = (x.salesCcy || '').toUpperCase();        // Column Q
    const pid  = x.periodId || '';
    sum += fx.convert(amt, from, target, pid);
  }
  return { incomeSales: sum, salesCcy: target };
}

/* ====================================
   Status / Investigate classification
   ==================================== */

// Back-compat alias: delegate to the global classifier to keep one source of truth
function SOC_classifyRow_(args) {
  const ctx = CORE_buildContext_();
  return LOGIC_classifyRow_(ctx, args || {});
}

/* =============================
   Writer for Society Issues
   ============================= */
/* =============================
   Writer for Society Issues  (sorted)
   ============================= */
function SOC_writeSocietyIssues_(ctx, rows) {
  const ss = ctx.ss;
  const name = 'Society Issues';
  const sh = SHEETWR_getOrCreate_(ss, name);

  // Clear old content, keep formatting rules/widths
  sh.clear({ contentsOnly: true });

  // Headers: drop (CCY) since mixed currencies per row
  const headers = ctx.config.OUTPUT_HEADERS_BASE.map(h => h.replace(/\s*\(CCY\)/g, ''));
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Only Investigate rows are passed already, but guard anyway
  let data = (rows || []).filter(r => r && r.investigate);

  // ---- NEW: sort (Society A→Z, then Status priority, then Contributor A→Z) ----
  const STATUS = ctx.config.STATUS;
  const rank = {};
  rank[STATUS.MISSING]       = 0;
  rank[STATUS.UNDER]         = 1;
  rank[STATUS.OVER]          = 2;
  rank[STATUS.ONTRACK]       = 3;
  rank[STATUS.NOT_EXPECTED]  = 4;
  data.sort((a, b) => {
    const sa = String(a.society || '').toLowerCase();
    const sb = String(b.society || '').toLowerCase();
    if (sa < sb) return -1;
    if (sa > sb) return 1;

    const ra = (rank[a.status] ?? 99);
    const rb = (rank[b.status] ?? 99);
    if (ra !== rb) return ra - rb;

    const ca = String(a.contributor || '').toLowerCase();
    const cb = String(b.contributor || '').toLowerCase();
    if (ca < cb) return -1;
    if (ca > cb) return 1;

    return 0;
  });
  // ---------------------------------------------------------------------------

  if (data.length) {
    const values = data.map(r => ([
      r.contributor,
      r.uuid,
      r.society,
      r.territory,
      CORE_properCase_(r.registrationStatus),
      r.playsTotal,
      r.playsQualified,
      r.expected,     // sales ccy
      r.income,       // sales ccy
      r.varianceAbs,  // sales ccy
      r.variancePct,
      r.status,
      r.investigate ? 'Yes' : 'No',
      r.notes
    ]));
    sh.getRange(2, 1, values.length, headers.length).setValues(values);

    // Per-row money formats for Expected/Income/Variance Abs (cols H,I,J)
    const n = values.length;
    const moneyFormats = [];
    for (let i = 0; i < n; i++) {
      const ccy =
        (data[i].currency || data[i].currencyCode || '').toString().toUpperCase() ||
        (ctx.config.DEFAULT_CURRENCY || 'USD');
      const symbol = (typeof SHEETWR_currencySymbol_ === 'function') ? SHEETWR_currencySymbol_(ccy) : '';
      const fmt = symbol ? `${symbol}#,##0.00` : (ccy ? `[$${ccy}] #,##0.00` : '#,##0.00');
      moneyFormats.push([fmt, fmt, fmt]);
    }
    sh.getRange(2, 8, n, 3).setNumberFormats(moneyFormats);

    // Plays & %
    sh.getRange(2, 6, n, 2).setNumberFormat('#,##0');
    sh.getRange(2, 11, n, 1).setNumberFormat('0.0%');

    // Conditional coloring by Status (12th column)
    if (typeof SHEETWR_applyStatusConditionalFormats_ === 'function') {
      SHEETWR_applyStatusConditionalFormats_(sh, 12, n, headers.length, /*startRow=*/2);
    }
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}

/** ============================================
 * Society Functions — Non-Home Territory Income
 * Builds a sheet listing income a performer received
 * from a society OUTSIDE that society's home territory.
 *
 * Columns: Contributor | UUID | Society | Territory | Income  (in Payee CCY)
 * ============================================ */

/** Entry point */
function SOC_buildNonHomeTerritoryIncome() {
  const ctx           = CORE_buildContext_();
  const contributors  = READ_readContributors_(ctx);        // Map<uuid,{id,name}>
  const statuses      = READ_readSalesforceStatuses_(ctx);  // Map<uuid, SocietyStatus[]>
  const societyHome   = READ_readSocietyHome_(ctx);         // Map<Society -> Home Territory>
  const incomes       = READ_readIncome_(ctx);              // GIR lines with payee/sales ccy, periodId, countryName
  const fx            = SOC_FX_load_(ctx);                  // { convert(amount, from, to, periodId) }

  // Determine dominant Payee CCY per contributor (agreement currency)
  const contributorCurrency = FX_determineContributorCurrencies_(ctx, incomes);

  // Build quick lookup: uuid -> Set(society) they have in Salesforce Statuses
  const socByUUID = new Map();
  statuses.forEach((arr, uuid) => {
    const set = new Set();
    (arr || []).forEach(s => { const sc = CORE_canonSociety_(s.societyName || ''); if (sc) set.add(sc); });
    socByUUID.set(uuid, set);
  });

  // Aggregate: key = uuid||society||territoryNonHome
  const agg = new Map();
  function keyOf(uuid, soc, terr) { return `${uuid}||${soc}||${terr}`; }

  incomes.forEach(row => {
    const uuid = row.uuid;
    if (!uuid) return;

    const soc = CORE_canonSociety_(row.societyName || '');
    if (!soc) return;

    // Only consider societies that appear on this performer's Salesforce statuses
    const allowedSoc = socByUUID.get(uuid);
    if (!allowedSoc || !allowedSoc.has(soc)) return;

    const lineTerr = CORE_canonCountry_(row.countryName || '');
    if (!lineTerr) return; // need a territory to compare

    const homeTerr = societyHome.get(soc) || '';
    if (!homeTerr) return; // if we don't know the home territory, skip to avoid false positives

    if (lineTerr === homeTerr) return; // only NON-home territory lines

    // Convert this line's payee amount into the contributor's payee currency (agreement ccy)
    const targetCCY = (contributorCurrency.get(uuid) || ctx.config.DEFAULT_CURRENCY || 'USD').toUpperCase();
    const fromCCY   = (row.payeeCcy || '').toUpperCase();
    const amtPayee  = Number(row.amountPayee) || 0;
    if (!amtPayee) return;

    const converted = fx.convert(amtPayee, fromCCY, targetCCY, row.periodId || '');

    const k = keyOf(uuid, soc, lineTerr);
    if (!agg.has(k)) {
      // resolve a nice contributor name
      const displayName = (function() {
        // prefer name from contributors tab; fall back to first status entry’s name if needed
        const c = contributors.get(uuid);
        if (c && c.name) return CORE_properCase_(c.name);
        const arr = statuses.get(uuid) || [];
        return CORE_properCase_(arr[0]?.contributorName || uuid);
      })();

      agg.set(k, {
        contributor: displayName,
        uuid,
        society: soc,
        territory: lineTerr,
        income: 0,
        currency: targetCCY
      });
    }
    const bucket = agg.get(k);
    bucket.income += converted;
  });

  // To array + sort: Contributor (A→Z), Society (A→Z), Territory (A→Z)
  const rows = Array.from(agg.values()).sort((a, b) => {
    const cA = (a.contributor || '').toLowerCase(), cB = (b.contributor || '').toLowerCase();
    if (cA !== cB) return cA < cB ? -1 : 1;
    const sA = (a.society || '').toLowerCase(), sB = (b.society || '').toLowerCase();
    if (sA !== sB) return sA < sB ? -1 : 1;
    const tA = (a.territory || '').toLowerCase(), tB = (b.territory || '').toLowerCase();
    return tA < tB ? -1 : tA > tB ? 1 : 0;
  });

  SOC_writeNonHomeTerritoryIncome_(ctx, rows);
}

/** Writer */
function SOC_writeNonHomeTerritoryIncome_(ctx, rows) {
  const ss = ctx.ss;
  const name = 'Non-Home Territory Income';
  const sh = SHEETWR_getOrCreate_(ss, name);
  sh.clear({ contentsOnly: true });

  const headers = ['Contributor', 'UUID', 'Society', 'Territory', 'Income'];
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);

  if (rows.length) {
    const values = rows.map(r => ([
      r.contributor,
      r.uuid,
      r.society,
      r.territory,
      r.income
    ]));
    sh.getRange(2,1,values.length,headers.length).setValues(values);

    // Per-row currency format for Income (col 5)
    const n = values.length;
    const moneyFormats = [];
    for (let i = 0; i < n; i++) {
      const ccy = (rows[i].currency || ctx.config.DEFAULT_CURRENCY || 'USD').toUpperCase();
      const symbol = (typeof SHEETWR_currencySymbol_ === 'function') ? SHEETWR_currencySymbol_(ccy) : '';
      moneyFormats.push([ symbol ? `${symbol}#,##0.00` : `[$${ccy}] #,##0.00` ]);
    }
    sh.getRange(2,5,n,1).setNumberFormats(moneyFormats);
  }

  sh.autoResizeColumns(1, headers.length);
}

