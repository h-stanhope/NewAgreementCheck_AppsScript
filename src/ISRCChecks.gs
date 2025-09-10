/** ============================================
 * ISRCChecks — Audit plays vs income alignment by Title + Main Artist
 * - Clusters within (UUID, Society, Home Territory, TitleNorm, ArtistNorm)
 * - Compares ISRCs getting plays vs ISRCs receiving income
 * - Emits only issues: Mismatch (top ISRC differs or no overlap) or Ambiguous (income w/o ISRC detail)
 * ============================================ */

function ISRC_buildMatchAudit() {
  const ctx          = CORE_buildContext_();
  const contributors = READ_readContributors_(ctx);
  const statuses     = READ_readSalesforceStatuses_(ctx);
  const societyHome  = READ_readSocietyHome_(ctx);
  const chartPlays   = READ_readChartMetric_(ctx);   // has title/artist columns (and we enrich further)
  const incomes      = READ_readIncome_(ctx);        // now includes recordingTitle + mainArtist
  const contentIdx   = READ_readContent_(ctx);       // for enrichment (title/artist if Chart lacks)

  // Normalizer for clustering by Title + Artist
  function normKey(s) {
    return String(s || '')
      .toLowerCase()
      .replace(/&/g, 'and')
      .replace(/[^a-z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // Build income clusters keyed by (uuid||society||territory||titleNorm||artistNorm)
  const incByCluster = new Map();
  incomes.forEach(x => {
    const uuid = x.uuid;
    if (!uuid) return;
    const soc  = CORE_canonSociety_(x.societyName || '');
    const terr = CORE_canonCountry_(x.countryName || '');
    if (!soc || !terr) return;

    const titleNorm  = normKey(x.recordingTitle);
    const artistNorm = normKey(x.mainArtist);
    if (!titleNorm || !artistNorm) return;

    const key = `${uuid}||${soc}||${terr}||${titleNorm}||${artistNorm}`;
    if (!incByCluster.has(key)) {
      incByCluster.set(key, {
        byIsrc: new Map(),       // ISRC -> sum (payee)
        total: 0,
        payeeCcy: (x.payeeCcy || '').toUpperCase(),
        displayTitle: x.recordingTitle,
        displayArtist: x.mainArtist
      });
    }
    const b = incByCluster.get(key);
    const amt = Number(x.amountPayee) || 0;
    b.total += amt;

    const isrc = CORE_str_(x.isrc);
    if (isrc) b.byIsrc.set(isrc, (b.byIsrc.get(isrc) || 0) + amt);

    // Keep first seen identity for display
    if (!b.displayTitle && x.recordingTitle) b.displayTitle = x.recordingTitle;
    if (!b.displayArtist && x.mainArtist)    b.displayArtist = x.mainArtist;
    if (!b.payeeCcy && x.payeeCcy)          b.payeeCcy = String(x.payeeCcy).toUpperCase();
  });

  const out = [];

  // Drive by statuses to respect society home territory
  statuses.forEach((socArr, uuid) => {
    if (!Array.isArray(socArr) || !socArr.length) return;
    const contrib = contributors.get(uuid) || { id: uuid, name: socArr[0]?.contributorName || '' };

    // contributor's plays once
    const myPlays = chartPlays.filter(p => p.uuid === uuid);

    socArr.forEach(st => {
      const society   = CORE_canonSociety_(st.societyName || '');
      if (!society) return;
      const territory = CORE_canonCountry_(societyHome.get(society) || '');

      // Only plays in the home territory
      const playsSub = myPlays.filter(p => !territory || CORE_canonCountry_(p.countryName || '') === territory);
      if (!playsSub.length) return;

      // Build play clusters by Title+Artist
      const playsByCluster = new Map(); // key -> {byIsrc Map, total, displayTitle, displayArtist}
      for (const p of playsSub) {
        const px = (contentIdx && typeof contentIdx.enrichPlay === 'function') ? contentIdx.enrichPlay(p, uuid) : p;

        const title = px.recordingTitle || px.songName;
        const artist = px.primaryArtist || px.primaryArtistDisplayedAs;
        const titleNorm  = normKey(title);
        const artistNorm = normKey(artist);
        if (!titleNorm || !artistNorm) continue;

        const key = `${uuid}||${society}||${territory}||${titleNorm}||${artistNorm}`;
        if (!playsByCluster.has(key)) {
          playsByCluster.set(key, {
            byIsrc: new Map(), // ISRC -> plays
            total: 0,
            displayTitle: title || '',
            displayArtist: artist || ''
          });
        }
        const g = playsByCluster.get(key);
        const isrc = CORE_str_(px.isrc);
        const plays = Number(px.playCount) || 0;

        if (isrc) g.byIsrc.set(isrc, (g.byIsrc.get(isrc) || 0) + plays);
        g.total += plays;

        // backfill identity if missing
        if (!g.displayTitle && title)  g.displayTitle = title;
        if (!g.displayArtist && artist) g.displayArtist = artist;
      }

      // Compare against income clusters with the same key
      playsByCluster.forEach((gp, key) => {
        const inc = incByCluster.get(key);
        if (!inc) return; // no income for this title/artist cluster in this society/territory

        if (!(gp.total > 0 && inc.total > 0)) return;

        // Determine top-play ISRC
        let topPlayIsrc = '', topPlayCount = 0;
        gp.byIsrc.forEach((cnt, isrc) => { if (cnt > topPlayCount) { topPlayIsrc = isrc; topPlayCount = cnt; } });

        // Determine top-income ISRC
        let topIncIsrc = '', topIncAmt = 0;
        inc.byIsrc.forEach((amt, isrc) => { if (amt > topIncAmt) { topIncIsrc = isrc; topIncAmt = amt; } });

        // Build readable lists
        function listStr(map, valueFmt) {
          const arr = [];
          map.forEach((v, k) => arr.push({ k, v }));
          arr.sort((a, b) => b.v - a.v);
          return arr.map(x => `${x.k}:${valueFmt(x.v)}`).join(', ');
        }
        const playsList  = listStr(gp.byIsrc, v => `${v}`);
        const incomeList = listStr(inc.byIsrc, v => `${v.toFixed(2)}`);

        // Decide outcome
        const playSet = new Set(Array.from(gp.byIsrc.keys()));
        const incSet  = new Set(Array.from(inc.byIsrc.keys()));
        const hasIncomeIsrcs = incSet.size > 0;

        let matchStatus = 'OK';
        let reason = 'Aligned on ISRC';
        if (!hasIncomeIsrcs) {
          matchStatus = 'Ambiguous';
          reason = 'Income lines lack ISRC detail';
        } else {
          let overlap = false;
          playSet.forEach(k => { if (incSet.has(k)) overlap = true; });
          if (!overlap) {
            matchStatus = 'Mismatch';
            reason = 'No overlap between play ISRCs and income ISRCs';
          } else if (topPlayIsrc && topIncIsrc && topPlayIsrc !== topIncIsrc) {
            matchStatus = 'Mismatch';
            reason = 'Top-play ISRC ≠ top-income ISRC';
          }
        }

        if (matchStatus === 'OK') return; // only issues

        out.push({
          contributor: contrib.name,
          uuid,
          society,
          territory,

          title: gp.displayTitle || inc.displayTitle || '',
          artist: gp.displayArtist || inc.displayArtist || '',

          topPlayIsrc,
          topPlayCount,
          playsTotal: gp.total,
          playsList,

          topIncIsrc,
          topIncAmt,
          incomeTotal: inc.total,
          payeeCcy: inc.payeeCcy || (ctx.config.DEFAULT_CURRENCY || 'USD'),
          incomeList,

          matchStatus, // Mismatch | Ambiguous
          reason
        });
      });
    });
  });

  ISRC_writeAudit_(ctx, out);
}

/** Writer */
function ISRC_writeAudit_(ctx, rows) {
  const ss = ctx.ss;
  const name = 'ISRC Match Audit';
  const sh = SHEETWR_getOrCreate_(ss, name);
  sh.clear({ contentsOnly: true });

  const headers = [
    'Contributor','UUID','Society','Territory',
    'Recording Title','Main Artist',
    'Top Play ISRC','Top Play Count','Play ISRCs','Total Plays',
    'Top Income ISRC','Top Income (Payee)','Total Income (Payee)','Payee CCY','Income ISRCs',
    'Match?','Reason'
  ];
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);

  // Sort: Society → Contributor → Match? (Mismatch first) → Title
  const rank = { 'Mismatch': 0, 'Ambiguous': 1, 'OK': 2 };
  rows.sort((a,b) => {
    const sA=(a.society||'').toLowerCase(), sB=(b.society||'').toLowerCase();
    if (sA!==sB) return sA<sB?-1:1;
    const cA=(a.contributor||'').toLowerCase(), cB=(b.contributor||'').toLowerCase();
    if (cA!==cB) return cA<cB?-1:1;
    const rA=rank[a.matchStatus]??99, rB=rank[b.matchStatus]??99;
    if (rA!==rB) return rA-rB;
    const tA=(a.title||'').toLowerCase(), tB=(b.title||'').toLowerCase();
    return tA<tB?-1:tA>tB?1:0;
  });

  if (rows.length) {
    const values = rows.map(r => ([
      r.contributor, r.uuid, r.society, r.territory,
      r.title, r.artist,
      r.topPlayIsrc, r.topPlayCount, r.playsList, r.playsTotal,
      r.topIncIsrc, r.topIncAmt, r.incomeTotal, r.payeeCcy, r.incomeList,
      r.matchStatus, r.reason
    ]));
    sh.getRange(2,1,values.length,headers.length).setValues(values);

    const n = values.length;
    sh.getRange(2,8,n,1).setNumberFormat('#,##0');     // Top Play Count
    sh.getRange(2,10,n,1).setNumberFormat('#,##0');    // Total Plays

    // Money columns K,L,M
    const moneyFormats = [];
    for (let i = 0; i < n; i++) {
      const symbol = (typeof SHEETWR_currencySymbol_==='function') ? SHEETWR_currencySymbol_(rows[i].payeeCcy || 'USD') : '';
      const fmt = symbol ? `${symbol}#,##0.00` : '#,##0.00';
      moneyFormats.push([fmt, fmt]); // Top Income, Total Income
    }
    sh.getRange(2,12,n,2).setNumberFormats(moneyFormats);

    // Conditional coloring on Match?
    const mRange = sh.getRange(2,16,n,1);
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Mismatch').setBackground('#f4cccc').setRanges([mRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Ambiguous').setBackground('#fff2cc').setRanges([mRange]).build()
    ];
    sh.setConditionalFormatRules(rules);
  }

  sh.autoResizeColumns(1, headers.length);
}
