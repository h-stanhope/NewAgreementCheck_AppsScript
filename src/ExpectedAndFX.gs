/** FX helpers & Expected calculations */

/**
 * Pick a dominant Payee Currency per contributor from their income lines.
 * Returns Map<uuid, currency>
 */
function FX_determineContributorCurrencies_(ctx, incomeRows) {
  const map = new Map();
  const sums = new Map(); // key `${uuid}||${ccy}` => amount

  for (const r of incomeRows) {
    const key = `${r.uuid}||${r.payeeCcy}`;
    sums.set(key, (sums.get(key) || 0) + (r.amountPayee || 0));
  }

  // choose max
  for (const key of sums.keys()) {
    const [uuid, ccy] = key.split('||');
    const cur = map.get(uuid);
    const curAmt = cur ? (sums.get(`${uuid}||${cur}`) || 0) : -1;
    const thisAmt = sums.get(key) || 0;
    if (!cur || thisAmt > curAmt) map.set(uuid, ccy);
  }

  return map;
}

/** Get FX rate from -> to by period (or latest if period not provided) */
function FX_getRate_(fx, from, to, periodId) {
  if (!from || !to || from === to) return 1;
  if (periodId) {
    const key = `${from}||${to}||${periodId}`;
    const r = fx.exact.get(key);
    if (r) return r;
  }
  const last = fx.latest.get(`${from}||${to}`);
  if (last) return last;
  // If missing direct, try via USD as simple bridge
  const mid1 = fx.latest.get(`${from}||USD`);
  const mid2 = fx.latest.get(`USD||${to}`);
  if (mid1 && mid2) return mid1 * mid2;
  return null; // indicate missing
}

/** Sum income for (uuid, society, territory) in contributor currency, period-aware FX */
function FX_sumIncomeForTerritoryInContributorCcy_({ ctx, uuid, society, territory, incomeRows, fx, contributorCurrency }) {
  let sum = 0;
  for (const r of incomeRows) {
    if (r.uuid !== uuid) continue;
    if (r.societyName !== society) continue;
    if (territory && r.countryName !== territory) continue;

    const from = r.payeeCcy || contributorCurrency;
    const rate = FX_getRate_(fx, from, contributorCurrency, r.periodId);
    const amt = (r.amountPayee || 0) * (rate || 1);
    sum += amt;
  }
  return { incomeAmount: sum };
}

/**
 * Expected value for qualified plays in a territory.
 * - Uses Benchmarks (rate per spin in benchmark currency)
 * - Adjusts for contributor share:
 *    * SESSION_MUSICIAN => /4
 *    * MAIN/FEATURING => / numberOfPrimaryPerformers
 * - Converts to contributor currency using LATEST FX (no statement period for expected)
 */
function EXP_computeExpectedForTerritory_({ ctx, uuid, society, territory, benchmarks, fx, chartPlaysSubset, contentByKey, contributorCurrency }) {
  const bm = benchmarks.get(`${territory}||${society}`);
  if (!bm || !bm.rate || !bm.currency) {
    return { expectedAmount: null }; // per your rule: no expected if no benchmark
  }

  let total = 0; // in benchmark currency
  for (const p of chartPlaysSubset) {
    if (!p.qualifies) continue;

    const contributorType = (p.contributorType || '').toUpperCase();
    const isNonFeatured = contributorType === 'SESSION_MUSICIAN';
    const isFeatured = contributorType === 'MAIN_PERFORMER' || contributorType === 'FEATURING_PERFORMER';

    // share adjustment
    let perSpin = bm.rate;
    if (isNonFeatured) {
      perSpin = perSpin / 4;
    } else if (isFeatured) {
      const n = Math.max(1, Number(p.numPrimary) || 1);
      perSpin = perSpin / n;
    }

    total += perSpin * (p.qualifiedCount || 0);
  }

  const rate = FX_getRate_(fx, bm.currency, contributorCurrency, /*period*/ null);
  const expectedInContributorCcy = (rate == null) ? null : total * rate;

  return { expectedAmount: expectedInContributorCcy };
}

/** Variance & Status evaluation */
function EVAL_buildStatusAndVariance_({ ctx, registrationStatus, expectedAmount, qualifiedPlays, incomeAmount }) {
  const expecting = ctx.config.EXPECTING_STATUS_SET.has(String(registrationStatus || '').toLowerCase());
  const S = ctx.config.STATUS;
  const T = ctx.config.THRESHOLDS;

  // No benchmark expected
  if (expectedAmount == null) {
    if (expecting && qualifiedPlays > ctx.config.MIN_QUAL_PLAYS_FOR_EXPECTATION) {
      // We expect SOME income. Mark Missing if none received; else On Track
      if ((incomeAmount || 0) <= 0) {
        return { status: S.MISSING, investigate: true, notes: 'No benchmark; >20 qualified plays; income missing.' };
      }
      return { status: S.ONTRACK, investigate: false, notes: 'No benchmark; >20 qualified plays; income received.' };
    }
    // Not expecting by status OR too few qualified plays – treat as On Track (no investigation)
    if (!expecting) {
      return { status: S.NOT_EXPECTED, investigate: false, notes: 'Registration not expecting income.' };
    }
    return { status: S.ONTRACK, investigate: false, notes: 'No benchmark; ≤20 qualified plays.' };
  }

  // We have expected — compute variance
  const exp = Number(expectedAmount) || 0;
  const inc = Number(incomeAmount) || 0;
  const varPct = exp ? (inc / exp - 1) : null;

  // If registration says not expecting, override status
  if (!expecting) {
    const base = (varPct == null) ? S.NOT_EXPECTED : S.NOT_EXPECTED;
    const notes = inc > 0 ? 'Registration not expecting, but income received.' : 'Registration not expecting.';
    return { status: base, investigate: false, notes };
  }

  if (varPct == null) {
    // Shouldn't happen when expected exists, but guard anyway
    return { status: S.ONTRACK, investigate: false, notes: 'Expected undefined.' };
  }

  if (varPct <= T.UNDER_MAX) {
    return { status: S.UNDER, investigate: true, notes: 'Underperforming vs expected.' };
  }
  if (varPct >= T.OVER_MIN) {
    return { status: S.OVER, investigate: false, notes: 'Overperforming vs expected.' };
  }
  return { status: S.ONTRACK, investigate: false, notes: 'Within acceptable variance.' };
}

/** Sorting: Missing → Underperforming → Overperforming → On Track → Not Expected; then by Variance% asc when applicable */
function EVAL_compareRowsByStatusThenVariance_(a, b) {
  const order = ['Missing', 'Underperforming', 'Overperforming', 'On Track', 'Not Expected'];
  const ai = order.indexOf(a.status); const bi = order.indexOf(b.status);
  if (ai !== bi) return ai - bi;

  const aHasVar = (a.variancePct != null);
  const bHasVar = (b.variancePct != null);
  if (aHasVar && bHasVar) return a.variancePct - b.variancePct;

  // When variance not applicable, keep as-is (stable) or sort by qualified plays desc as a mild heuristic
  if (aHasVar !== bHasVar) return aHasVar ? -1 : 1;
  return (b.playsQualified || 0) - (a.playsQualified || 0);
}
