/** 
 * NR Analysis — Entry points & menu
 * Kollective Neighbouring Rights
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('NR Analysis')
    .addSubMenu(ui.createMenu('Performer')
      .addItem('Performer Tabs', 'USECASE_buildPerformerTabs')
      .addItem('Performer Overview', 'USECASE_buildOverviewOnly')
    )
    .addSubMenu(ui.createMenu('Society')
      .addItem('Society Issues', 'SOC_buildSocietyIssuesOverview')
      .addItem('Non-Home Territory Income', 'SOC_buildNonHomeTerritoryIncome')
    )
    .addSubMenu(ui.createMenu('Recording')
      .addItem('Recording Overview', 'REC_buildRecordingOverview')
      .addItem('Recording Issues', 'REC_buildRecordingIssues')
      .addItem('Most Played Recordings', 'REC_buildMostPlayedRecordings')
    )
    .addSeparator()
    .addSubMenu(ui.createMenu('Diagnostics')
      .addItem('Build Qualification Audit', 'DIAG_buildQualificationAudit')
      .addItem('Log Random Sample (200)', 'DIAG_logRandomSample')
      .addItem('Reconcile Overview vs Engine', 'DIAG_reconcileOverviewVsEngine')
    )
    .addSeparator()
    .addItem('Export Full Analysis (Start/Resume)', 'EXPORT_startFullAnalysis')
    .addItem('Abort Full Export', 'EXPORT_abortFullAnalysis')
    .addItem('Rebuild All', 'USECASE_rebuildAll')
    .addItem('Remove All', 'USECASE_removeAll')
    .addToUi();
}

/** Build (or refresh) all performer tabs and the overview. */
function USECASE_buildPerformerTabs() {
  const ctx = CORE_buildContext_();
  const rowsByContributor = USECASE_buildPerformerRows_(ctx);
  SHEETWR_writeAllPerformerTabs_(ctx, rowsByContributor);
  SHEETWR_writeOverview_(ctx, rowsByContributor);
}

/** Only rebuild the consolidated overview (from fresh engine rows). */
function USECASE_buildOverviewOnly() {
  const ctx = CORE_buildContext_();
  const rowsByContributor = USECASE_buildPerformerRows_(ctx);
  SHEETWR_writeOverview_(ctx, rowsByContributor);
}

/** ---------------------------
 * Rebuild All
 * - Clears ALL feature outputs
 * - Rebuilds Performer tabs+overview, Society Issues, Non-Home Income, Recording Overview, Recording Issues
 * (Diagnostics are not rebuilt automatically.)
 * --------------------------- */
function USECASE_rebuildAll() {
  const ctx = CORE_buildContext_();

  // Delete all outputs first
  SHEETWR_deleteExistingPerformerTabs_(ctx);
  SHEETWR_deleteOverview_(ctx);
  SHEETWR_deleteSocietyIssues_(ctx);
  SHEETWR_deleteNonHomeIncome_(ctx);
  SHEETWR_deleteRecordingOverview_(ctx);
  SHEETWR_deleteRecordingIssues_(ctx);
  // (Optional) also clear diagnostics to avoid stale views
  SHEETWR_deleteQualificationAudit_(ctx);
  SHEETWR_deleteReconciliationSheet_(ctx);
  // (Optional) if you use the ISRC audit and want it cleared too:
  SHEETWR_deleteIsrcAudit_(ctx);

  // Rebuild feature outputs
  USECASE_buildPerformerTabs();            // Performer tabs + Performer Overview
  SOC_buildSocietyIssuesOverview();        // Society Issues
  SOC_buildNonHomeTerritoryIncome();       // Non-Home Territory Income
  REC_buildRecordingOverview();            // Recording Overview
  REC_buildRecordingIssues();              // Recording Issues
}

/** ---------------------------
 * Remove All
 * - Clears ALL feature outputs (and diagnostics) but does not rebuild
 * --------------------------- */
function USECASE_removeAll() {
  const ctx = CORE_buildContext_();
  SHEETWR_deleteExistingPerformerTabs_(ctx);
  SHEETWR_deleteOverview_(ctx);
  SHEETWR_deleteSocietyIssues_(ctx);
  SHEETWR_deleteNonHomeIncome_(ctx);
  SHEETWR_deleteRecordingOverview_(ctx);
  SHEETWR_deleteRecordingIssues_(ctx);
  // (Optional) diagnostics
  SHEETWR_deleteQualificationAudit_(ctx);
  SHEETWR_deleteReconciliationSheet_(ctx);
  SHEETWR_deleteIsrcAudit_(ctx);
}

/**
 * Orchestrates row building for all contributors.
 * Returns: Map<contributorUUID, { contributor, currency, rows } >
 */
function USECASE_buildPerformerRows_(ctx) {
  return PERF_buildPerformerRows_(ctx);
}

/* =======================
   Small delete helpers
   ======================= */

function SHEETWR_deleteIfExists_(ctx, name) {
  if (!name) return;
  const sh = ctx.ss.getSheetByName(name);
  if (sh) ctx.ss.deleteSheet(sh);
}

function SHEETWR_deleteRecordingOverview_(ctx) {
  const name = (ctx.config.SHEETS && ctx.config.SHEETS.REC_OVERVIEW) || 'Recording Overview';
  SHEETWR_deleteIfExists_(ctx, name);
}

function SHEETWR_deleteRecordingIssues_(ctx) {
  const name = (ctx.config.SHEETS && ctx.config.SHEETS.REC_ISSUES) || 'Recording Issues';
  SHEETWR_deleteIfExists_(ctx, name);
}

function SHEETWR_deleteNonHomeIncome_(ctx) {
  const name = (ctx.config.SHEETS && ctx.config.SHEETS.NON_HOME_INCOME) || 'Non-Home Territory Income';
  SHEETWR_deleteIfExists_(ctx, name);
}

function SHEETWR_deleteIsrcAudit_(ctx) {
  const name = (ctx.config.SHEETS && ctx.config.SHEETS.ISRC_AUDIT) || 'ISRC Match Audit';
  SHEETWR_deleteIfExists_(ctx, name);
}

function SHEETWR_deleteQualificationAudit_(ctx) {
  const name = (ctx.config.SHEETS && ctx.config.SHEETS.QUAL_AUDIT) || 'Qualification Audit';
  SHEETWR_deleteIfExists_(ctx, name);
}

function SHEETWR_deleteReconciliationSheet_(ctx) {
  const name = 'Reconciliation - Plays';
  SHEETWR_deleteIfExists_(ctx, name);
}

/**
 * Build a per-track, per-society diagnostic sheet showing why a play qualifies or not.
 * Scope: all contributors in "Contributor List"; only plays in each society's home territory.
 */
function DIAG_buildQualificationAudit() {
  const ctx = CORE_buildContext_();
  const contributors = READ_readContributors_(ctx);
  const statuses = READ_readSalesforceStatuses_(ctx);
  const societyHome = READ_readSocietyHome_(ctx);
  const chartPlays = READ_readChartMetric_(ctx);

  // Create/clear the audit sheet
  const sh = (function() {
    let s = ctx.ss.getSheetByName(ctx.config.SHEETS.QUAL_AUDIT);
    if (!s) s = ctx.ss.insertSheet(ctx.config.SHEETS.QUAL_AUDIT);
    s.clear({ contentsOnly: true });
    return s;
  })();

  const headers = [
    'Contributor', 'UUID', 'Society', 'Territory (Home)', 'Airplay Country',
    'ISRC', 'Contribution ID', 'Contributor Type', 'Roles', 'Play Count',
    'Qualifies?', 'Rule / Gate', 'Future Year', 'Featured Requirement', 'Featured Passed', 'Roles Excluded',
    // optional detail flags to verify all gates were evaluated
    'Hit: Citizenship','Hit: Residency','Hit: CoL','Hit: CoR','Hit: CoP','Hit: CoM','Hit: ADAMI EEA','Hit: ADAMI RC+CoL'
  ];
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sh.setFrozenRows(1);

  const out = [];

  for (const [uuid, contrib] of contributors.entries()) {
    const myStatuses = statuses.get(uuid) || [];
    if (!myStatuses.length) continue;

    const myPlays = chartPlays.filter(p => p.uuid === uuid);

    for (const st of myStatuses) {
      const society = CORE_canonSociety_(st.societyName || '');
      if (!society) continue;
      const territory = CORE_canonCountry_(societyHome.get(society) || '');

      const subset = myPlays.filter(p => !territory || p.countryName === territory);
      for (const p of subset) {
        const crit = {
          society,
          territory,
          contributorType: (p.contributorType || '').toString().toUpperCase(),
          roles: p.roles || '',
          cor1: p.cor1 || '',
          cor2: p.cor2 || '',
          coc: p.coc || '',
          com: p.com || '',
          col: p.col || '',
          isrc: p.isrc || '',
          contributionId: p.contributionId || ''
        };
        const expl = QUAL.explain(crit);

        const ruleText = expl.rule || (expl.qualified ? 'Qualified (no gate logged)' : (expl.societyRowFound ? '' : 'No society row'));

        out.push([
          contrib.name, uuid, society, territory, p.countryName,
          p.isrc, p.contributionId, p.contributorType, p.roles, p.playCount,
          expl.qualified ? 'Yes' : 'No',
          ruleText,
          expl.futureYear || '',
          expl.featuredOnly || '',
          expl.featuredPassed === null ? '' : (expl.featuredPassed ? 'Yes' : 'No'),
          expl.rolesExcluded ? 'Yes' : 'No',
          // flags
          expl.checks?.Citizenship ? 'Yes' : 'No',
          expl.checks?.Residency   ? 'Yes' : 'No',
          expl.checks?.CoL        ? 'Yes' : 'No',
          expl.checks?.CoR        ? 'Yes' : 'No',
          expl.checks?.CoP        ? 'Yes' : 'No',
          expl.checks?.CoM        ? 'Yes' : 'No',
          expl.checks?.ADAMI_EEA  ? 'Yes' : 'No',
          expl.checks?.ADAMI_RC_CoL ? 'Yes' : 'No'
        ]);
      }
    }
  }

  if (out.length) {
    sh.getRange(2,1,out.length,headers.length).setValues(out);
  }

  // Light formatting
  const lastCol = headers.length;
  const dataRows = Math.max(0, sh.getLastRow() - 1);
  if (dataRows > 0) {
    sh.getRange(2, 10, dataRows, 1).setNumberFormat('#,##0'); // Play Count
    const qRange = sh.getRange(2, 11, dataRows, 1);
    const yesRule = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Yes').setBackground('#b6d7a8').setRanges([qRange]).build();
    const noRule  = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('No').setBackground('#f4cccc').setRanges([qRange]).build();
    sh.setConditionalFormatRules([yesRule, noRule]);
    sh.autoResizeColumns(1, lastCol);
  }
}

/**
 * Log a random sample of ~200 contributor-society-track evaluations.
 */
function DIAG_logRandomSample() {
  const ctx = CORE_buildContext_();
  const contributors = READ_readContributors_(ctx);
  const statuses = READ_readSalesforceStatuses_(ctx);
  const societyHome = READ_readSocietyHome_(ctx);
  const chartPlays = READ_readChartMetric_(ctx);

  const tuples = [];
  for (const [uuid, c] of contributors.entries()) {
    const myStatuses = statuses.get(uuid) || [];
    if (!myStatuses.length) continue;
    const myPlays = chartPlays.filter(p => p.uuid === uuid);
    for (const st of myStatuses) {
      const society = CORE_canonSociety_(st.societyName || '');
      const terr = CORE_canonCountry_(societyHome.get(society) || '');
      myPlays.forEach(p => {
        if (!terr || p.countryName === terr) {
          tuples.push({ uuid, name: c.name, society, terr, p });
        }
      });
    }
  }

  tuples.sort(() => Math.random() - 0.5);
  const sample = tuples.slice(0, Math.min(200, tuples.length));

  sample.forEach(t => {
    const crit = {
      society: t.society, territory: t.terr,
      contributorType: (t.p.contributorType || '').toString().toUpperCase(),
      roles: t.p.roles || '',
      cor1: t.p.cor1 || '', cor2: t.p.cor2 || '',
      coc: t.p.coc || '', com: t.p.com || '', col: t.p.col || '',
      isrc: t.p.isrc || '', contributionId: t.p.contributionId || ''
    };
    const expl = QUAL.explain(crit);
    console.log(
      `[${t.name} | ${t.uuid}] ${t.society} @ ${t.terr} :: ISRC=${t.p.isrc} plays=${t.p.playCount} → ${expl.qualified ? 'QUAL' : 'NO'} | rule="${expl.rule}" ` +
      `featuredOnly=${expl.featuredOnly} passed=${expl.featuredPassed} rolesExcluded=${expl.rolesExcluded}`
    );
  });

  console.log(`Diagnostics sample logged: ${sample.length} rows.`);
}

/**
 * Reconcile the "Performer Overview" against a fresh engine recomputation.
 * Always writes ALL rows (not just diffs) so you can filter/sort in-sheet.
 */
function DIAG_reconcileOverviewVsEngine() {
  const ctx = CORE_buildContext_();
  const overviewName = ctx.config.SHEETS.OVERVIEW_OUT || 'Performer Overview';
  const sh = ctx.ss.getSheetByName(overviewName);
  if (!sh) throw new Error(`Missing sheet: ${overviewName}`);

  const data = sh.getDataRange().getValues();
  if (!data.length) throw new Error('Overview is empty.');
  const headers = data[0].map(v => String(v || ''));
  const rows = data.slice(1);
  const h = CORE_headerIndexMap_(headers);

  function col(...cands) {
    for (const c of cands) {
      const k = String(c).toLowerCase().trim();
      if (h.has(k)) return h.get(k);
    }
    for (const c of cands) {
      const q = String(c).toLowerCase().trim();
      const idx = headers.findIndex(x => String(x).toLowerCase().includes(q));
      if (idx >= 0) return idx;
    }
    return -1;
  }

  const iContributor = col('contributor');
  const iUUID        = col('uuid','contributor id');
  const iSoc         = col('society');
  const iTerr        = col('territory (home)','territory','home territory');
  const iPTotal      = col('plays (total)','plays total');
  const iPQ          = col('plays (qualified)','plays qualified');

  if (iUUID === -1 || iSoc === -1) throw new Error('Could not find UUID and/or Society columns in Overview.');
  if (iPTotal === -1 || iPQ === -1) throw new Error('Could not find Plays (Total)/(Qualified) columns in Overview.');

  // Read chart plays once
  const chartPlays = READ_readChartMetric_(ctx);

  // Cache engine results per key
  const cache = new Map();
  function engineResult(uuid, society, territory) {
    const key = `${uuid}||${society}||${territory}`;
    if (cache.has(key)) return cache.get(key);
    const res = EXP_computePlaysForTerritory_({
      ctx,
      uuid,
      society: CORE_canonSociety_(society),
      territory: territory ? CORE_canonCountry_(territory) : '',
      chartPlays,
      contentByKey: null
    });
    cache.set(key, res);
    return res;
  }

  const outName = 'Reconciliation - Plays';
  let rec = ctx.ss.getSheetByName(outName);
  if (!rec) rec = ctx.ss.insertSheet(outName); else rec.clear({ contentsOnly: true });

  const outHead = [
    'Contributor','UUID','Society','Territory',
    'Overview Total','Engine Total','Δ Total',
    'Overview Qualified','Engine Qualified','Δ Qualified'
  ];
  rec.getRange(1,1,1,outHead.length).setValues([outHead]).setFontWeight('bold');
  rec.setFrozenRows(1);

  const out = [];
  for (const r of rows) {
    const uuid = CORE_str_(r[iUUID]); if (!uuid) continue;
    const contributor = iContributor >= 0 ? CORE_str_(r[iContributor]) : '';
    const society = CORE_canonSociety_(CORE_str_(r[iSoc]));
    const territory = iTerr >= 0 ? CORE_canonCountry_(CORE_str_(r[iTerr])) : '';

    const ovPT = Number(r[iPTotal]) || 0;
    const ovPQ = Number(r[iPQ]) || 0;

    const eng = engineResult(uuid, society, territory);
    const et = eng.totalPlays;
    const eq = eng.qualifiedPlays;

    out.push([contributor, uuid, society, territory, ovPT, et, et - ovPT, ovPQ, eq, eq - ovPQ]);
  }

  if (out.length) {
    rec.getRange(2,1,out.length,outHead.length).setValues(out);
  }

  const rowsCount = Math.max(0, rec.getLastRow()-1);
  if (rowsCount > 0) {
    rec.getRange(2,5,rowsCount,6).setNumberFormat('#,##0');
    const deltaCols = [7,10];
    const rules = [];
    deltaCols.forEach(c => {
      const rng = rec.getRange(2,c,rowsCount,1);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberNotEqualTo(0).setBackground('#fff2cc').setRanges([rng]).build());
    });
    rec.setConditionalFormatRules(rules);
    rec.autoResizeColumns(1, outHead.length);
  }
}
