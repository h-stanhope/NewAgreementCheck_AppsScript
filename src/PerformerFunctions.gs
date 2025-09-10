/** ============================================
 * Performer Functions
 * - Build performer rows (per contributor UUID)
 * - Reuses shared readers, QUAL engine, FX & eval logic
 * - Returns Map<uuid, { contributor, currency, rows }>
 * ============================================ */

/**
 * Build all performer rows (grouped by contributor UUID).
 * Each row matches the schema expected by SheetWriters.
 *
 * Uses:
 *  - EXP_computePlaysForTerritory_: same qualification engine as the audit
 *  - EXP_computeExpectedForTerritory_: sums qualifying spins → expected in contributor (payee) currency
 *  - FX_sumIncomeForTerritoryInContributorCcy_: sums income → contributor (payee) currency
 *  - LOGIC_classifyRow_: global status logic (Missing/Under/Over/On Track/Not Expected)
 *  - EVAL_compareRowsByStatusThenVariance_: sorting (Missing → Under → Over → On Track → Not Expected, then variance asc)
 */
function PERF_buildPerformerRows_(ctx) {
  const contributors = READ_readContributors_(ctx);       // Map<uuid,{id,name,(optional)currency}>
  const statuses     = READ_readSalesforceStatuses_(ctx); // Map<uuid, SocietyStatus[]>
  const societyHome  = READ_readSocietyHome_(ctx);        // Map<Society, Territory>
  const chartPlays   = READ_readChartMetric_(ctx);        // Array of ChartPlay rows (parsed)
  const contentByKey = READ_readContent_(ctx);            // Map<(ISRC|ContributionId), content info> (fallback enrichment)
  const benchmarks   = READ_readBenchmarks_(ctx);         // Map<`${territory}||${society}`, {rate,currency,...}>
  const fx           = READ_readFxRates_(ctx);            // FX lookups (+ latest fallbacks for Expected)
  const incomeRows   = READ_readIncome_(ctx);             // Income lines (with Payee/Sales & Statement Period ID)

  // Determine a per-contributor agreement currency (dominant Payee Currency across all income)
  const contributorCurrency = FX_determineContributorCurrencies_(ctx, incomeRows);

  const outByContributor = new Map();

  for (const [uuid, contributor] of contributors.entries()) {
    const myStatuses = statuses.get(uuid) || [];
    const myCurrency = (contributorCurrency.get(uuid) || ctx.config.DEFAULT_CURRENCY || 'USD')
      .toString().toUpperCase();

    const rows = [];
    for (const st of myStatuses) {
      const society = CORE_canonSociety_(st.societyName || '');
      if (!society) continue;

      const territory = CORE_canonCountry_(societyHome.get(society) || '');
      const registrationStatus = (st.status || '').toString().toLowerCase();

      // 1) Plays (same engine as Qualification Audit)
      const playsAgg = EXP_computePlaysForTerritory_({
        ctx, uuid, society, territory, chartPlays, contentByKey
      });
      const qualifiedPlays = Number(playsAgg.qualifiedPlays) || 0;

      // 2) Benchmark presence for (territory, society)
      const benchKey = `${territory}||${society}`;
      const benchRec = benchmarks.get(benchKey);
      const hasBenchmark = !!(benchRec && benchRec.rate && benchRec.currency);

      // 3) Expected in contributor currency (payee ccy), from qualifying spins in chartSubset
      const expectedInfo = EXP_computeExpectedForTerritory_({
        ctx,
        uuid,
        society,
        territory,
        benchmarks,
        fx,
        chartPlaysSubset: playsAgg.chartSubset,
        contentByKey,
        contributorCurrency: myCurrency
      });
      const expectedAmount = Number(expectedInfo.expectedAmount) || 0;

      // 4) Income in contributor currency (period-based FX per income line)
      const incomeInfo = FX_sumIncomeForTerritoryInContributorCcy_({
        ctx,
        uuid,
        society,
        territory,
        incomeRows,
        fx,
        contributorCurrency: myCurrency
      });
      const incomeAmount = Number(incomeInfo.incomeAmount) || 0;

      // 5) Variance (payee currency)
      const varianceAbs = incomeAmount - expectedAmount;
      const variancePct = expectedAmount > 0 ? (incomeAmount / expectedAmount) - 1 : null;

      // 6) Global Status / Investigate / Notes
      const cls = LOGIC_classifyRow_(ctx, {
        registrationStatus,
        hasBenchmark,
        qualifiedPlays,
        income: incomeAmount,
        expected: expectedAmount
      });

      rows.push({
        contributor: contributor.name,
        uuid,
        society,
        territory,
        registrationStatus,
        playsTotal: Number(playsAgg.totalPlays) || 0,
        playsQualified: qualifiedPlays,
        expected: expectedAmount,
        income: incomeAmount,
        varianceAbs,
        variancePct,
        status: cls.status,
        investigate: cls.investigate,
        notes: cls.notes
      });
    }

    // Ordering: Missing → Underperforming → Overperforming → On Track → Not Expected; then by Variance (asc)
    rows.sort((a, b) => EVAL_compareRowsByStatusThenVariance_(a, b));

    outByContributor.set(uuid, {
      contributor,
      currency: myCurrency,   // used for the column headings per performer tab
      rows
    });
  }

  return outByContributor;
}

/** Convenience: build & write tabs + overview in one go (optional). */
function PERF_buildAndWritePerformerOutputs_() {
  const ctx = CORE_buildContext_();
  const rowsByContributor = PERF_buildPerformerRows_(ctx);
  SHEETWR_writeAllPerformerTabs_(ctx, rowsByContributor);
  SHEETWR_writeOverview_(ctx, rowsByContributor);
}
