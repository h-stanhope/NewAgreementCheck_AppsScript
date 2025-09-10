/** ============================================
 * Recording Features
 * - Build a user-facing Recording Overview
 * - Per-recording qualification + expected/income/variance in payee ccy
 * - Adds a new status: "Non-Qualifying" for tracks that do not qualify
 * - Skips non-track-detail societies for income-based recording views
 * - Adds "Most Played Recordings" for non-track-detail societies
 * ============================================ */

/** Entry point: Recording Overview (PAYEE CCY), only for societies with track-level income */
function REC_buildRecordingOverview() {
  const ctx = CORE_buildContext_();

  const contributors = READ_readContributors_(ctx);
  const statuses     = READ_readSalesforceStatuses_(ctx);
  const societyHome  = READ_readSocietyHome_(ctx);
  const chartPlays   = READ_readChartMetric_(ctx);
  const benchmarks   = READ_readBenchmarks_(ctx);
  const fx           = SOC_FX_load_(ctx);
  const incomes      = READ_readIncome_(ctx);

  const contributorCurrency = FX_determineContributorCurrencies_(ctx, incomes);

  // NEW: figure out which societies have track-level income
  const { trackDetail } = REC_trackDetailSocieties_(incomes);

  const rows = REC_buildRecordingRows_(ctx, {
    contributors, statuses, societyHome, chartPlays, benchmarks, fx, incomes, contributorCurrency,
    // extra filter: include only societies with track detail
    includeSociety: (soc) => trackDetail.has(soc)
  });

  REC_writeRecordingOverview_(ctx, rows);
}

/** Determine which societies provide track-level income detail. */
function REC_trackDetailSocieties_(incomes) {
  const all = new Set();
  const track = new Set();
  for (const x of incomes || []) {
    const soc = CORE_canonSociety_(x.societyName || '');
    if (!soc) continue;
    all.add(soc);
    if (CORE_str_(x.contributionId) || CORE_str_(x.isrc) || CORE_str_(x.soundRecordingId)) {
      track.add(soc);
    }
  }
  const nonTrack = new Set(Array.from(all).filter(s => !track.has(s)));
  return { trackDetail: track, nonTrack, allSocieties: all };
}

/** Build all rows for Recording Overview (one row per [uuid,society,track]) */
function REC_buildRecordingRows_(ctx, deps) {
  const {
    contributors, statuses, societyHome, chartPlays, benchmarks, fx, incomes, contributorCurrency
  } = deps;

  const contentIdx = READ_readContent_(ctx); // enrich identity + qual fields

  // Helper: payee ccy per contributor
  function payeeCcyFor(uuid) {
    const c = contributorCurrency.get(uuid);
    if (c) return String(c).toUpperCase();
    const meta = contributors.get(uuid);
    if (meta && meta.currency) return String(meta.currency).toUpperCase();
    return (ctx.config.DEFAULT_CURRENCY || 'USD');
  }

  // Title/Artist normalizer for clustering
  function normKey(s) {
    return String(s || '')
      .toLowerCase()
      .replace(/&/g, 'and')
      .replace(/[^a-z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // -------- Precompute cluster income in PAYEE CCY (per contributor) --------
  // key: uuid||society||territory||titleNorm||artistNorm||targetPayeeCcy
  const clusterPayee = new Map();
  incomes.forEach(x => {
    const uuid = x.uuid; if (!uuid) return;
    const soc  = CORE_canonSociety_(x.societyName || ''); if (!soc) return;
    const terr = CORE_canonCountry_(x.countryName || '');
    const titleNorm  = normKey(x.recordingTitle);
    const artistNorm = normKey(x.mainArtist);
    if (!titleNorm || !artistNorm) return;

    const target = payeeCcyFor(uuid);
    const amt = REC_safeConvert_(fx, Number(x.amountPayee) || 0, (x.payeeCcy || '').toUpperCase(), target, x.periodId);
    if (!amt) return;

    const key = `${uuid}||${soc}||${terr}||${titleNorm}||${artistNorm}||${target}`;
    if (!clusterPayee.has(key)) clusterPayee.set(key, { total: 0, byIsrc: new Map() });
    const bucket = clusterPayee.get(key);
    bucket.total += amt;
    const isrc = CORE_str_(x.isrc);
    if (isrc) bucket.byIsrc.set(isrc, (bucket.byIsrc.get(isrc) || 0) + amt);
  });

  function payeeClusterFor(uuid, society, territory, title, artist) {
    const key = `${uuid}||${society}||${territory}||${normKey(title)}||${normKey(artist)}||${payeeCcyFor(uuid)}`;
    return clusterPayee.get(key) || { total: 0, byIsrc: new Map() };
  }

  // Money formatter for notes
  function fmtMoney(ccy, amt) {
    const a = Number(amt) || 0;
    const sym = (typeof SHEETWR_currencySymbol_ === 'function') ? SHEETWR_currencySymbol_(ccy) : '';
    return sym ? `${sym}${a.toFixed(2)}` : `${ccy} ${a.toFixed(2)}`;
  }

  const out = [];

  statuses.forEach((socArr, uuid) => {
    if (!Array.isArray(socArr) || !socArr.length) return;

    const contribMeta = contributors.get(uuid) || { id: uuid, name: socArr[0]?.contributorName || '' };
    const payeeCcy    = payeeCcyFor(uuid);

    socArr.forEach(st => {
      const society = CORE_canonSociety_(st.societyName || '');
      if (!society) return;

      // Optional society filter (e.g., only track-detail societies)
      if (deps.includeSociety && !deps.includeSociety(society)) return;

      const territory = CORE_canonCountry_(societyHome.get(society) || '');
      const regStatus = (st.status || '').toString().toLowerCase();

      // Plays for this contributor in this society's home territory
      const playsSubset = chartPlays.filter(p => (p.uuid === uuid) && (!territory || CORE_canonCountry_(p.countryName || '') === territory));
      if (!playsSubset.length) return;

      // Benchmark details (if any)
      const bmKey  = `${territory}||${society}`;
      const bmRec  = benchmarks.get(bmKey);
      const hasBM  = !!(bmRec && bmRec.rate && bmRec.currency);
      const bmRate = hasBM ? Number(bmRec.rate) : 0;
      const bmCcy  = hasBM ? String(bmRec.currency || '').toUpperCase() : '';

      // Group by track (ISRC / Contribution ID / NR ID / meta)
      const grp = new Map();

      for (const p of playsSubset) {
        const px = (contentIdx && typeof contentIdx.enrichPlay === 'function') ? contentIdx.enrichPlay(p, uuid) : p;

        const key = (function trackKey(px) {
          const cid = CORE_str_(px.contributionId);
          const isrc = CORE_str_(px.isrc);
          const nrid = CORE_str_(px.nrId || px.nrID || px.nr_id);
          if (cid) return `CID:${cid}`;
          if (isrc) return `ISRC:${isrc}`;
          if (nrid) return `NR:${nrid}`;
          return `META:${CORE_str_(px.recordingTitle)}|${CORE_str_(px.version || px.recordingVersion)}|${CORE_str_(px.primaryArtist)}`;
        })(px);

        if (!grp.has(key)) {
          grp.set(key, {
            key,
            // Identity
            recordingTitle: CORE_str_(px.recordingTitle) || CORE_str_(px.songName) || '',
            versionTitle:   CORE_str_(px.version || px.recordingVersion) || '',
            mainArtist:     CORE_str_(px.primaryArtist || px.primaryArtistDisplayedAs) || '',
            isrc:           CORE_str_(px.isrc),
            contributionId: CORE_str_(px.contributionId),
            nrId:           CORE_str_(px.nrId || px.nrID || px.nr_id),
            contributorType: (px.contributorType || '').toString().toUpperCase(),
            roles:           CORE_str_(px.roles),
            numPrimary:      Number(px.numPrimary || px.numberOfPrimaryPerformers || 1) || 1,

            totalPlays: 0,
            qualifiedPlays: 0,
            expectedBench: 0,
            qualifiesAny: false,
            // keep last blockers for context (optional)
            lastFailRule: '',
            featuredOnly: null,
            featuredPassed: null,
            rolesExcluded: null
          });
        }

        const g = grp.get(key);
        const plays = Number(px.playCount) || 0;
        g.totalPlays += plays;

        // Qualification (engine)
        const crit = {
          society,
          territory,
          contributorType: (px.contributorType || '').toString().toUpperCase(),
          roles: px.roles || '',
          cor1: px.cor1 || '',
          cor2: px.cor2 || '',
          coc:  px.coc  || '',
          com:  px.com  || '',
          col:  px.col  || '',
          isrc: px.isrc || '',
          contributionId: px.contributionId || ''
        };
        const expl = QUAL.explain(crit);

        if (expl.qualified) {
          g.qualifiesAny = true;
          g.qualifiedPlays += plays;

          if (hasBM && bmRate) {
            let perSpin = bmRate;
            const type = (px.contributorType || '').toString().toUpperCase();
            if (type === 'SESSION_MUSICIAN') perSpin = perSpin / 4;
            else {
              const np = Number(px.numPrimary || 1) || 1;
              if (np > 1) perSpin = perSpin / np;
            }
            g.expectedBench += perSpin * plays;
          }
        } else {
          g.lastFailRule   = expl.rule || g.lastFailRule || 'No gate matched';
          g.featuredOnly   = (expl.featuredOnly != null ? expl.featuredOnly : g.featuredOnly);
          g.featuredPassed = (expl.featuredPassed != null ? expl.featuredPassed : g.featuredPassed);
          g.rolesExcluded  = (expl.rolesExcluded != null ? expl.rolesExcluded : g.rolesExcluded);
        }

        // Backfill identity
        if (!g.recordingTitle) g.recordingTitle = CORE_str_(px.recordingTitle) || CORE_str_(px.songName) || g.recordingTitle;
        if (!g.versionTitle)   g.versionTitle   = CORE_str_(px.version || px.recordingVersion) || g.versionTitle;
        if (!g.mainArtist)     g.mainArtist     = CORE_str_(px.primaryArtist || px.primaryArtistDisplayedAs) || g.mainArtist;
        if (!g.isrc)           g.isrc           = CORE_str_(px.isrc) || g.isrc;
        if (!g.nrId)           g.nrId           = CORE_str_(px.nrId || px.nrID || px.nr_id) || g.nrId;
        if (!g.contributorType) g.contributorType = (px.contributorType || '').toString().toUpperCase() || g.contributorType;
        if (!g.roles)           g.roles           = CORE_str_(px.roles) || g.roles;
      }

      // Emit rows per track group
      grp.forEach(g => {
        const expected = (hasBM && bmCcy)
          ? REC_safeConvert_(fx, Number(g.expectedBench) || 0, bmCcy, payeeCcy, null)
          : 0;

        const income = REC_sumIncomeForTrackInPayeeCcy_({
          uuid, society, territory, track: g, incomes, fx, targetPayeeCcy: payeeCcy
        });

        const varianceAbs = (income || 0) - (expected || 0);
        const variancePct = expected > 0 ? (income / expected) - 1 : null;

        // Base classification (per-track income only)
        const base = LOGIC_classifyRow_(ctx, {
          registrationStatus: regStatus,
          hasBenchmark: hasBM,
          qualifiedPlays: g.qualifiedPlays,
          income,
          expected
        });

        let status = base.status;
        let investigate = base.investigate;
        let notes = base.notes || '';

        // ---- Data Mismatch logic: combine cluster income (same Title+Artist in payee ccy) ----
        const bucket = payeeClusterFor(uuid, society, territory, g.recordingTitle, g.mainArtist);
        const thisIsrcIncome = (g.isrc && bucket.byIsrc.size) ? (bucket.byIsrc.get(g.isrc) || 0) : income;
        const otherIncome = Math.max(0, (bucket.total || 0) - (thisIsrcIncome || 0));
        const combinedIncome = (income || 0) + (otherIncome || 0);

        if (otherIncome > 0) {
          const combined = LOGIC_classifyRow_(ctx, {
            registrationStatus: regStatus,
            hasBenchmark: hasBM,
            qualifiedPlays: g.qualifiedPlays,
            income: combinedIncome,
            expected
          });

          const otherStr = fmtMoney(payeeCcy, otherIncome);
          const combStr  = fmtMoney(payeeCcy, combinedIncome);

          if (status === ctx.config.STATUS.MISSING) {
            if (combined.status === ctx.config.STATUS.UNDER) {
              status = ctx.config.STATUS.UNDER;
              investigate = true;
              notes = `Data Mismatch: other ISRC income ${otherStr} (combined ${combStr}). ${notes || ''}`.trim();
            } else {
              status = combined.status;
              investigate = false;
              notes = `Data Mismatch: other ISRC income ${otherStr} (combined ${combStr}).`.trim();
            }
          } else if (status === ctx.config.STATUS.UNDER && (combined.status === ctx.config.STATUS.ONTRACK || combined.status === ctx.config.STATUS.OVER)) {
            status = combined.status;
            investigate = false;
            notes = `Underperforming upgraded via Data Mismatch: other ISRC income ${otherStr} (combined ${combStr}).`;
          } else if (combined.status === status && (status === ctx.config.STATUS.ONTRACK || status === ctx.config.STATUS.OVER)) {
            notes = notes ? `${notes} | Also detected Data Mismatch: other ISRC income ${otherStr} (combined ${combStr}).` :
                            `Data Mismatch detected: other ISRC income ${otherStr} (combined ${combStr}).`;
          }
        }

        out.push({
          contributor: contribMeta.name,
          uuid,
          society,
          territory,
          registrationStatus: regStatus,

          recordingTitle: g.recordingTitle,
          versionTitle: g.versionTitle,
          mainArtist: g.mainArtist,
          isrc: g.isrc,
          contributionId: g.contributionId,
          nrId: g.nrId,
          contributorType: g.contributorType,
          roles: g.roles,

          playsTotal: g.totalPlays,
          playsQualified: g.qualifiedPlays,

          expected,
          income,
          varianceAbs,
          variancePct,
          status,
          investigate,
          notes,
          currency: payeeCcy
        });
      });
    });
  });

  // Order: Contributor → Society → Status priority → Recording Title
  const STATUS = ctx.config.STATUS;
  const rank = {};
  rank[STATUS.MISSING]       = 0;
  rank[STATUS.UNDER]         = 1;
  rank['Non-Qualifying']     = 2;
  rank[STATUS.OVER]          = 3;
  rank[STATUS.ONTRACK]       = 4;
  rank[STATUS.NOT_EXPECTED]  = 5;

  out.sort((a, b) => {
    const ca = (a.contributor || '').toLowerCase();
    const cb = (b.contributor || '').toLowerCase();
    if (ca < cb) return -1; if (ca > cb) return 1;

    const sa = (a.society || '').toLowerCase();
    const sb = (b.society || '').toLowerCase();
    if (sa < sb) return -1; if (sa > sb) return 1;

    const ra = (rank[a.status] ?? 99);
    const rb = (rank[b.status] ?? 99);
    if (ra !== rb) return ra - rb;

    const ta = (a.recordingTitle || '').toLowerCase();
    const tb = (b.recordingTitle || '').toLowerCase();
    if (ta < tb) return -1; if (ta > tb) return 1;

    return 0;
  });

  return out;
}

/** Safe FX conversion helper used by Recording features */
function REC_safeConvert_(fx, amount, from, to, periodId) {
  amount = Number(amount) || 0;
  if (!amount) return 0;

  try {
    if (fx && typeof fx.convert === 'function') {
      return fx.convert(amount, from, to, periodId);
    }
  } catch (e) {
    // ignore
  }

  const f = (from || '').toUpperCase();
  const t = (to || '').toUpperCase();
  return (!f || !t || f === t) ? amount : amount;
}

/** Sum income in PAYEE currency for a specific track (match by Contribution ID, else ISRC, else Sound Recording ID) */
function REC_sumIncomeForTrackInPayeeCcy_({ uuid, society, territory, track, incomes, fx, targetPayeeCcy }) {
  let sum = 0;
  const cid = track.contributionId;
  const isrc = track.isrc;
  const nrId = track.nrId;

  for (const x of incomes) {
    if (x.uuid !== uuid) continue;
    if (CORE_canonSociety_(x.societyName || '') !== society) continue;
    if (territory && CORE_canonCountry_(x.countryName || '') !== territory) continue;

    // Track-level matching
    const hit =
      (cid && CORE_str_(x.contributionId) === cid) ||
      (isrc && CORE_str_(x.isrc) === isrc) ||
      (nrId && CORE_str_(x.soundRecordingId) === nrId);

    if (!hit) continue;

    const amt  = Number(x.amountPayee) || 0;
    if (!amt) continue;
    const from = (x.payeeCcy || '').toUpperCase();
    const pid  = x.periodId || '';
    sum += REC_safeConvert_(fx, amt, from, targetPayeeCcy, pid);
  }
  return sum;
}

/** Writer: "Recording Overview" sheet (per-row currency symbols, status coloring) */
function REC_writeRecordingOverview_(ctx, rows) {
  const ss = ctx.ss;
  const name = 'Recording Overview';
  const sh = SHEETWR_getOrCreate_(ss, name);
  sh.clear({ contentsOnly: true });

  const headers = [
    'Contributor','UUID','Society','Territory','Registration Status',
    'Recording Title','Recording Version','Main Artist','ISRC','Contribution ID',
    'Contributor Type','Roles',
    'Plays (Total)','Plays (Qualified)',
    'Expected','Income','Variance Abs','Variance %',
    'Status','Investigate?','Notes'
  ];
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);

  if (rows.length) {
    const values = rows.map(r => ([
      r.contributor,
      r.uuid,
      r.society,
      r.territory,
      CORE_properCase_(r.registrationStatus),
      r.recordingTitle,
      r.versionTitle,
      r.mainArtist,
      r.isrc,
      r.contributionId,
      r.contributorType,
      r.roles,
      r.playsTotal,
      r.playsQualified,
      r.expected,
      r.income,
      r.varianceAbs,
      r.variancePct,
      r.status,
      r.investigate ? 'Yes' : 'No',
      r.notes
    ]));
    sh.getRange(2,1,values.length,headers.length).setValues(values);

    const n = values.length;

    // Number formats
    sh.getRange(2, 13, n, 2).setNumberFormat('#,##0');  // Plays total/qualified
    // Per-row money formats for Expected/Income/Variance Abs (cols 15..17)
    const moneyFormats = [];
    for (let i = 0; i < n; i++) {
      const ccy =
        (rows[i].currency || rows[i].currencyCode || '').toString().toUpperCase() ||
        (ctx.config.DEFAULT_CURRENCY || 'USD');
      const symbol = (typeof SHEETWR_currencySymbol_ === 'function') ? SHEETWR_currencySymbol_(ccy) : '';
      const fmt = symbol ? `${symbol}#,##0.00` : (ccy ? `[$${ccy}] #,##0.00` : '#,##0.00');
      moneyFormats.push([fmt, fmt, fmt]);
    }
    sh.getRange(2, 15, n, 3).setNumberFormats(moneyFormats);

    // Variance % (col 18)
    sh.getRange(2, 18, n, 1).setNumberFormat('0.0%');

    // Status banding (col 19) using shared function
    if (typeof SHEETWR_applyStatusConditionalFormats_ === 'function') {
      SHEETWR_applyStatusConditionalFormats_(sh, 19, n, headers.length, /*startRow=*/2);
      const rules = sh.getConditionalFormatRules() || [];
      sh.setConditionalFormatRules(rules);
    }
  }

  sh.autoResizeColumns(1, headers.length);
}

/** Build "Recording Issues" — only Missing/Underperforming tracks, in SALES currency.
 * Excludes "Missing" rows that are actually Data Mismatch (money on sibling ISRCs).
 * Skips non-track-detail societies entirely.
 */
function REC_buildRecordingIssues() {
  const ctx          = CORE_buildContext_();
  const contributors = READ_readContributors_(ctx);
  const statuses     = READ_readSalesforceStatuses_(ctx);
  const societyHome  = READ_readSocietyHome_(ctx);
  const chartPlays   = READ_readChartMetric_(ctx);
  const benchmarks   = READ_readBenchmarks_(ctx);
  const incomes      = READ_readIncome_(ctx);
  const fx           = SOC_FX_load_(ctx);
  const contentIdx   = READ_readContent_(ctx);

  // Only include societies with track detail
  const { trackDetail } = REC_trackDetailSocieties_(incomes);

  // Normalizer
  function normKey(s) {
    return String(s || '').toLowerCase().replace(/&/g,'and').replace(/[^a-z0-9]+/g,' ').replace(/\s+/g,' ').trim();
  }

  // Choose a target Sales CCY for a (uuid,society,territory)
  function targetSalesCcyFor(uuid, society, territory) {
    // Prefer this contributor’s most common sales ccy for this (soc,terr)
    const counts = new Map();
    incomes.forEach(x => {
      if (x.uuid !== uuid) return;
      if (CORE_canonSociety_(x.societyName || '') !== society) return;
      if (territory && CORE_canonCountry_(x.countryName || '') !== territory) return;
      const c = (x.salesCcy || '').toUpperCase();
      if (!c) return;
      counts.set(c, (counts.get(c) || 0) + 1);
    });
    if (counts.size) {
      let best = '', n = -1;
      counts.forEach((v,k) => { if (v > n) { n = v; best = k; } });
      if (best) return best;
    }
    // fallback to society's general sales ccy
    const socCcy = SOC_buildSalesCcyBySociety_(incomes).get(society);
    return socCcy || (ctx.config.DEFAULT_CURRENCY || 'USD');
  }

  // -------- Precompute cluster income in SALES CCY --------
  // key: uuid||society||territory||titleNorm||artistNorm||targetSalesCcy
  const clusterSales = new Map();
  // cache target ccy per triple
  const tgtCache = new Map(); // key: uuid||soc||terr -> ccy
  function tgtKey(uuid, soc, terr){ return `${uuid}||${soc}||${terr}`; }

  incomes.forEach(x => {
    const uuid = x.uuid; if (!uuid) return;
    const soc  = CORE_canonSociety_(x.societyName || ''); if (!soc) return;
    const terr = CORE_canonCountry_(x.countryName || '');
    const titleNorm  = normKey(x.recordingTitle);
    const artistNorm = normKey(x.mainArtist);
    if (!titleNorm || !artistNorm) return;

    const tk = tgtKey(uuid, soc, terr);
    let target = tgtCache.get(tk);
    if (!target) { target = targetSalesCcyFor(uuid, soc, terr); tgtCache.set(tk, target); }

    const amt = REC_safeConvert_(fx, Number(x.amountSales) || 0, (x.salesCcy || '').toUpperCase(), target, x.periodId);
    if (!amt) return;

    const key = `${uuid}||${soc}||${terr}||${titleNorm}||${artistNorm}||${target}`;
    if (!clusterSales.has(key)) clusterSales.set(key, { total: 0, byIsrc: new Map() });
    const bucket = clusterSales.get(key);
    bucket.total += amt;
    const isrc = CORE_str_(x.isrc);
    if (isrc) bucket.byIsrc.set(isrc, (bucket.byIsrc.get(isrc) || 0) + amt);
  });

  function salesClusterFor(uuid, society, territory, title, artist, targetSalesCcy) {
    const key = `${uuid}||${society}||${territory}||${normKey(title)}||${normKey(artist)}||${targetSalesCcy}`;
    return clusterSales.get(key) || { total: 0, byIsrc: new Map() };
  }

  const out = [];

  statuses.forEach((socArr, uuid) => {
    if (!Array.isArray(socArr) || !socArr.length) return;
    const contrib = contributors.get(uuid) || { id: uuid, name: socArr[0]?.contributorName || '' };

    const myPlays = chartPlays.filter(p => p.uuid === uuid);

    socArr.forEach(st => {
      const society   = CORE_canonSociety_(st.societyName || '');
      if (!society) return;
      if (!trackDetail.has(society)) return; // skip non-track-detail societies

      const territory = CORE_canonCountry_(societyHome.get(society) || '');
      const regStatus = (st.status || '').toString().toLowerCase();

      const tgtSalesCcy = targetSalesCcyFor(uuid, society, territory);

      const playsSubset = myPlays.filter(p => !territory || CORE_canonCountry_(p.countryName || '') === territory);
      if (!playsSubset.length) return;

      const bmKey = `${territory}||${society}`;
      const bmRec = benchmarks.get(bmKey);
      const hasBM  = !!(bmRec && bmRec.rate && bmRec.currency);
      const bmRate = hasBM ? Number(bmRec.rate) : 0;
      const bmCcy  = hasBM ? String(bmRec.currency || '').toUpperCase() : '';

      const grp = new Map();

      for (const p of playsSubset) {
        const px = (contentIdx && typeof contentIdx.enrichPlay === 'function') ? contentIdx.enrichPlay(p, uuid) : p;
        const key = (function trackKey(px) {
          const cid = CORE_str_(px.contributionId);
          const isrc = CORE_str_(px.isrc);
          const nrid = CORE_str_(px.nrId || px.nrID || px.nr_id);
          if (cid) return `CID:${cid}`;
          if (isrc) return `ISRC:${isrc}`;
          if (nrid) return `NR:${nrid}`;
          return `META:${CORE_str_(px.recordingTitle)}|${CORE_str_(px.version || px.recordingVersion)}|${CORE_str_(px.primaryArtist)}`;
        })(px);

        if (!grp.has(key)) {
          grp.set(key, {
            key,
            recordingTitle: CORE_str_(px.recordingTitle) || CORE_str_(px.songName) || '',
            versionTitle:   CORE_str_(px.version || px.recordingVersion) || '',
            mainArtist:     CORE_str_(px.primaryArtist || px.primaryArtistDisplayedAs) || '',
            isrc:           CORE_str_(px.isrc),
            contributionId: CORE_str_(px.contributionId),
            nrId:           CORE_str_(px.nrId || px.nrID || px.nr_id),
            contributorType: (px.contributorType || '').toString().toUpperCase(),
            roles: CORE_str_(px.roles),
            numPrimary: Number(px.numPrimary || px.numberOfPrimaryPerformers || 1) || 1,
            totalPlays: 0,
            qualifiedPlays: 0,
            expectedBench: 0,
            qualifiesAny: false
          });
        }
        const g = grp.get(key);
        const plays = Number(px.playCount) || 0;
        g.totalPlays += plays;

        const crit = {
          society, territory,
          contributorType: (px.contributorType || '').toString().toUpperCase(),
          roles: px.roles || '',
          cor1: px.cor1 || '', cor2: px.cor2 || '',
          coc: px.coc || '', com: px.com || '', col: px.col || '',
          isrc: px.isrc || '', contributionId: px.contributionId || ''
        };
        const expl = QUAL.explain(crit);
        if (expl.qualified) {
          g.qualifiesAny = true;
          g.qualifiedPlays += plays;
          if (hasBM && bmRate) {
            let perSpin = bmRate;
            const type = (px.contributorType || '').toString().toUpperCase();
            if (type === 'SESSION_MUSICIAN') perSpin = perSpin / 4;
            else {
              const np = Number(px.numPrimary || 1) || 1;
              if (np > 1) perSpin = perSpin / np;
            }
            g.expectedBench += perSpin * plays;
          }
        }
      }

      grp.forEach(g => {
        const expectedSales = (hasBM && bmCcy) ? REC_safeConvert_(fx, Number(g.expectedBench) || 0, bmCcy, tgtSalesCcy, null) : 0;

        const incomeSales = REC_sumIncomeForTrackInSalesCcy_({
          uuid, society, territory, track: g, incomes, fx, targetSalesCcy: tgtSalesCcy
        });

        const cls = LOGIC_classifyRow_(ctx, {
          registrationStatus: regStatus,
          hasBenchmark: hasBM,
          qualifiedPlays: g.qualifiedPlays,
          income: incomeSales,
          expected: expectedSales
        });

        // If it's Missing but cluster shows income on other ISRCs in SALES CCY → skip (data mismatch)
        if (cls.status === ctx.config.STATUS.MISSING) {
          const bucket = salesClusterFor(uuid, society, territory, g.recordingTitle, g.mainArtist, tgtSalesCcy);
          const thisIsrcIncome = (g.isrc && bucket.byIsrc.size) ? (bucket.byIsrc.get(g.isrc) || 0) : incomeSales;
          const otherIncome = Math.max(0, (bucket.total || 0) - (thisIsrcIncome || 0));
          if (otherIncome > 0) {
            return; // not a pay problem; omit from issues
          }
        }

        if (cls.status !== ctx.config.STATUS.MISSING && cls.status !== ctx.config.STATUS.UNDER) return;

        out.push({
          contributor: contrib.name,
          uuid,
          society,
          territory,
          registrationStatus: regStatus,
          recordingTitle: g.recordingTitle,
          versionTitle: g.versionTitle,
          mainArtist: g.mainArtist,
          isrc: g.isrc,
          contributionId: g.contributionId,
          nrId: g.nrId,
          contributorType: g.contributorType,
          roles: g.roles,
          playsTotal: g.totalPlays,
          playsQualified: g.qualifiedPlays,
          expected: expectedSales,
          income: incomeSales,
          varianceAbs: (incomeSales || 0) - (expectedSales || 0),
          variancePct: expectedSales > 0 ? (incomeSales / expectedSales) - 1 : null,
          status: cls.status,
          investigate: true,
          notes: cls.notes,
          currency: tgtSalesCcy
        });
      });
    });
  });

  // Sort: Society (A→Z) → Performer (A→Z) → Status (Missing, Underperforming)
  const rank = { 'Missing': 0, 'Underperforming': 1 };
  out.sort((a, b) => {
    const sA = (a.society || '').toLowerCase(), sB = (b.society || '').toLowerCase();
    if (sA !== sB) return sA < sB ? -1 : 1;
    const cA = (a.contributor || '').toLowerCase(), cB = (b.contributor || '').toLowerCase();
    if (cA !== cB) return cA < cB ? -1 : 1;
    const rA = rank[a.status] ?? 99, rB = rank[b.status] ?? 99;
    if (rA !== rB) return rA - rB;
    const tA = (a.recordingTitle || '').toLowerCase(), tB = (b.recordingTitle || '').toLowerCase();
    return tA < tB ? -1 : tA > tB ? 1 : 0;
  });

  REC_writeRecordingIssues_(ctx, out);
}

/** Sum income in SALES currency for a specific track */
function REC_sumIncomeForTrackInSalesCcy_({ uuid, society, territory, track, incomes, fx, targetSalesCcy }) {
  let sum = 0;
  const cid  = track.contributionId;
  const isrc = track.isrc;
  const nrId = track.nrId;

  for (const x of incomes) {
    if (x.uuid !== uuid) continue;
    if (CORE_canonSociety_(x.societyName || '') !== society) continue;
    if (territory && CORE_canonCountry_(x.countryName || '') !== territory) continue;

    // Track-level match
    const hit =
      (cid && CORE_str_(x.contributionId) === cid) ||
      (isrc && CORE_str_(x.isrc) === isrc) ||
      (nrId && CORE_str_(x.soundRecordingId) === nrId);
    if (!hit) continue;

    const amt  = Number(x.amountSales) || 0;
    if (!amt) continue;
    const from = (x.salesCcy || '').toUpperCase();
    const pid  = x.periodId || '';
    sum += REC_safeConvert_(fx, amt, from, targetSalesCcy, pid);
  }
  return sum;
}

function REC_writeRecordingIssues_(ctx, rows) {
  const ss = ctx.ss;
  const name = 'Recording Issues';
  const sh = SHEETWR_getOrCreate_(ss, name);
  sh.clear({ contentsOnly: true });

  // Reuse Recording Overview headers (no CCY suffix; per-row symbols instead)
  const headers = [
    'Contributor','UUID','Society','Territory','Registration Status',
    'Recording Title','Recording Version','Main Artist','ISRC','Contribution ID',
    'Contributor Type','Roles',
    'Plays (Total)','Plays (Qualified)',
    'Expected','Income','Variance Abs','Variance %',
    'Status','Investigate?','Notes'
  ];
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);

  if (rows.length) {
    const values = rows.map(r => ([
      r.contributor,
      r.uuid,
      r.society,
      r.territory,
      CORE_properCase_(r.registrationStatus),
      r.recordingTitle,
      r.versionTitle,
      r.mainArtist,
      r.isrc,
      r.contributionId,
      r.contributorType,
      r.roles,
      r.playsTotal,
      r.playsQualified,
      r.expected,     // SALES ccy
      r.income,       // SALES ccy
      r.varianceAbs,  // SALES ccy
      r.variancePct,
      r.status,
      'Yes',          // This sheet only lists issues
      r.notes
    ]));
    sh.getRange(2,1,values.length,headers.length).setValues(values);

    const n = values.length;

    // Number formats
    sh.getRange(2, 13, n, 2).setNumberFormat('#,##0');
    sh.getRange(2, 18, n, 1).setNumberFormat('0.0%');

    // Per-row currency formats for Expected/Income/Variance Abs (cols 15..17)
    const moneyFormats = [];
    for (let i = 0; i < n; i++) {
      const ccy =
        (rows[i].currency || rows[i].currencyCode || '').toString().toUpperCase() ||
        (ctx.config.DEFAULT_CURRENCY || 'USD');
      const symbol = (typeof SHEETWR_currencySymbol_ === 'function') ? SHEETWR_currencySymbol_(ccy) : '';
      const fmt = symbol ? `${symbol}#,##0.00` : (ccy ? `[$${ccy}] #,##0.00` : '#,##0.00');
      moneyFormats.push([fmt, fmt, fmt]);
    }
    sh.getRange(2, 15, n, 3).setNumberFormats(moneyFormats);

    // Conditional coloring by Status (col 19)
    if (typeof SHEETWR_applyStatusConditionalFormats_ === 'function') {
      SHEETWR_applyStatusConditionalFormats_(sh, 19, n, headers.length, /*startRow=*/2);
    }
  }

  sh.autoResizeColumns(1, headers.length);
}

/** Build a sheet listing the most played recordings for societies with NO track detail. */
function REC_buildMostPlayedRecordings() {
  const ctx = CORE_buildContext_();

  const contributors = READ_readContributors_(ctx);
  const statuses     = READ_readSalesforceStatuses_(ctx);
  const societyHome  = READ_readSocietyHome_(ctx);
  const chartPlays   = READ_readChartMetric_(ctx);
  const incomes      = READ_readIncome_(ctx);

  const { nonTrack } = REC_trackDetailSocieties_(incomes);
  const contentIdx   = READ_readContent_(ctx);

  const out = [];

  statuses.forEach((socArr, uuid) => {
    if (!Array.isArray(socArr) || !socArr.length) return;
    const contribName = CORE_properCase_(socArr[0]?.contributorName || contributors.get(uuid)?.name || uuid);

    socArr.forEach(st => {
      const society   = CORE_canonSociety_(st.societyName || '');
      if (!society || !nonTrack.has(society)) return;

      const territory = CORE_canonCountry_(societyHome.get(society) || '');
      const regStatus = (st.status || '').toString().toLowerCase();

      const plays = chartPlays.filter(p => p.uuid === uuid && (!territory || p.countryName === territory));
      if (!plays.length) return;

      // group by track
      const grp = new Map();
      for (const p of plays) {
        const px = (contentIdx && typeof contentIdx.enrichPlay === 'function') ? contentIdx.enrichPlay(p, uuid) : p;
        const key = CORE_str_(px.contributionId) || CORE_str_(px.isrc) ||
                    `${CORE_str_(px.recordingTitle)}|${CORE_str_(px.recordingVersion)}|${CORE_str_(px.primaryArtist)}`;

        if (!grp.has(key)) {
          grp.set(key, {
            recordingTitle: CORE_str_(px.recordingTitle) || CORE_str_(px.songName) || '',
            versionTitle: CORE_str_(px.version || px.recordingVersion) || '',
            mainArtist: CORE_str_(px.primaryArtist || px.primaryArtistDisplayedAs) || '',
            isrc: CORE_str_(px.isrc),
            contributionId: CORE_str_(px.contributionId),
            contributorType: (px.contributorType || '').toString().toUpperCase(),
            roles: CORE_str_(px.roles),
            totalPlays: 0, qualifiedPlays: 0,
            lastFailRule: '', qualifiesAny: false
          });
        }
        const g = grp.get(key);
        const n = Number(px.playCount) || 0;
        g.totalPlays += n;

        // qualify
        const crit = {
          society, territory,
          contributorType: (px.contributorType || '').toString().toUpperCase(),
          roles: px.roles || '',
          cor1: px.cor1 || '', cor2: px.cor2 || '',
          coc: px.coc || '', com: px.com || '', col: px.col || '',
          isrc: px.isrc || '', contributionId: px.contributionId || ''
        };
        const expl = QUAL.explain(crit);
        if (expl.qualified) {
          g.qualifiesAny = true;
          g.qualifiedPlays += n;
        } else {
          g.lastFailRule = expl.rule || g.lastFailRule || 'No gate matched';
        }
      }

      grp.forEach(g => {
        let status, notes;
        if (!g.qualifiesAny) {
          status = 'Non-Qualifying';
          notes = g.lastFailRule || 'Does not qualify.';
        } else {
          status = 'To Register'; // green like On Track
          notes = 'Qualifies in home territory; ensure track is registered with this society.';
        }
        out.push({
          contributor: contribName, uuid, society, territory, registrationStatus: regStatus,
          recordingTitle: g.recordingTitle, versionTitle: g.versionTitle, mainArtist: g.mainArtist,
          isrc: g.isrc, contributionId: g.contributionId, contributorType: g.contributorType, roles: g.roles,
          playsTotal: g.totalPlays, playsQualified: g.qualifiedPlays,
          status, notes
        });
      });
    });
  });

  // Sort: Society → Contributor → Status (To Register first) → Plays Qualified desc
  out.sort((a,b)=>{
    const s = a.society.localeCompare(b.society); if (s) return s;
    const c = a.contributor.localeCompare(b.contributor); if (c) return c;
    const rank = (x)=> x==='To Register'?0:1;
    const r = rank(a.status)-rank(b.status); if (r) return r;
    return (b.playsQualified||0)-(a.playsQualified||0);
  });

  REC_writeMostPlayedRecordings_(ctx, out);
}

/** Write the "Most Played Recordings" sheet with simple formats + status colors. */
function REC_writeMostPlayedRecordings_(ctx, rows) {
  const sh = SHEETWR_getOrCreate_(ctx.ss, 'Most Played Recordings');
  sh.clear({ contentsOnly: true });

  const headers = [
    'Contributor','UUID','Society','Territory','Registration Status',
    'Recording Title','Recording Version','Main Artist','ISRC','Contribution ID',
    'Contributor Type','Roles','Plays (Total)','Plays (Qualified)','Status','Notes'
  ];
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);

  if (rows && rows.length) {
    const vals = rows.map(r => [
      r.contributor, r.uuid, r.society, r.territory, CORE_properCase_(r.registrationStatus),
      r.recordingTitle, r.versionTitle, r.mainArtist, r.isrc, r.contributionId,
      r.contributorType, r.roles, r.playsTotal, r.playsQualified, r.status, r.notes
    ]);
    sh.getRange(2,1,vals.length,headers.length).setValues(vals);
    const n = vals.length;
    sh.getRange(2,13,n,2).setNumberFormat('#,##0'); // plays
    // Status coloring (green for To Register, red for Non-Qualifying)
    const allData = sh.getRange(2,1,n,headers.length);
    const rules = [];
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('To Register').setBackground('#b6d7a8').setRanges([allData]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Non-Qualifying').setBackground('#f4cccc').setRanges([allData]).build());
    sh.setConditionalFormatRules(rules);
  }

  sh.autoResizeColumns(1, headers.length);
}
