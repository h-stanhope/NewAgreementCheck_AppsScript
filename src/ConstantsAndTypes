/** Constants & lightweight types */

const CONFIG = {
  SHEETS: {
    CONTRIBUTORS: 'Contributor List',
    INCOME: 'Generated Income Report',
    STATUSES: 'Salesforce Statuses',
    CONTENT: 'Content',
    CHART: 'Chart Metric - Radio Play',
    SOC_HOME: 'Society Home',
    FX: 'FX Rates',
    BENCHMARKS: 'Benchmarks',
    OVERVIEW_OUT: 'Performer Overview',
    QUAL_AUDIT: 'Qualification Audit'
  },

  // Expected output columns (currency code placeholders are replaced at write time)
  OUTPUT_HEADERS_BASE: [
    'Contributor', 'UUID', 'Society', 'Territory', 'Registration Status',
    'Plays (Total)', 'Plays (Qualified)',
    'Expected (CCY)', 'Income (CCY)', 'Variance Abs (CCY)', 'Variance %',
    'Status', 'Investigate?', 'Notes'
  ],

  // Which statuses mean "we expect income"
  EXPECTING_STATUS_SET: new Set([
    'no agreement registration required',
    'requires registration',
    'registration sent',
    'registration confirmed'
  ]),

  // Status labels (single source of truth)
  STATUS: {
    MISSING: 'Missing',
    UNDER: 'Underperforming',
    OVER: 'Overperforming',
    ONTRACK: 'On Track',
    NOT_EXPECTED: 'Not Expecting',
    DATA_MISMATCH: 'Data Mismatch'
  },

  // Thresholds (variance % = income/expected - 1)
  THRESHOLDS: {
    UNDER_MAX: -0.30, // ≤ -30%
    OVER_MIN: 0.50    // ≥ +50%
  },

  // Qualifying plays threshold for "we should expect *some* income" when no benchmarks
  MIN_QUAL_PLAYS_FOR_EXPECTATION: 20,

  // Fallback default currency (shouldn't be needed given your data)
  DEFAULT_CURRENCY: 'USD',

  // --- Qualification reference (central workbook) ---
  QUAL: {
    // Update if your central reference workbook URL changes:
    CENTRAL_URL: 'https://docs.google.com/spreadsheets/d/1ffYCppr7lQ4PFw8SIJPFSvnfK7941RCP5apc0kFPkRA/edit',
    SHEETS: {
      QUALIFICATION: 'Qualification',
      COUNTRY_LIST: 'Country List',
      EXCLUDED_ROLES: 'Excluded Roles',
      QUAL_YEAR: 'Qualification Year'
    },
    DEBUG: false
  },

  // ---- Canonicalisation (aliases → one canonical label) ----
  SOCIETY_ALIASES: [
    { canonical: 'SoundExchange', aliases: ['Sound Exchange', 'SOUNDEXCHANGE'] },
    { canonical: 'AFM & SAG-AFTRA ("The Fund")', aliases: ['AFM & SAG-AFTRA Fund', 'The Fund', 'AFM', 'AFM & SAG-AFTRA'] },
    { canonical: 'GRAMEX DK', aliases: ['GRAMEX DENMARK'] },
    { canonical: 'GRAMEX FI', aliases: ['GRAMEX FINLAND'] },
    { canonical: 'GRAMO', aliases: ['GRAMO NORWAY'] },
    { canonical: 'ISAMRA', aliases: ['ISRA'] },
    { canonical: 'Nuovo IMAIE', aliases: ['NUOVO IMAIE'] },
    { canonical: 'PI (Serbia)', aliases: ['PI SERBIA', 'PI Serbia'] }

  ],

  COUNTRY_ALIASES: [
    { canonical: 'United States', aliases: ['US', 'U.S.', 'USA', 'U.S.A.', 'United States of America', 'America', 'U S A', 'U S'] },
    { canonical: 'United Kingdom', aliases: ['UK', 'U.K.', 'Great Britain', 'GB', 'GBR', 'Britain', 'England', 'Scotland', 'Wales', 'Northern Ireland'] },
    { canonical: 'Czech Republic', aliases: ['Czechia', 'Czech Rep', 'CZ', 'Czech Republic (Czechia)'] }
  ]
};

/** Build a context object once */
function CORE_buildContext_() {
  return {
    ss: SpreadsheetApp.getActive(),
    config: CONFIG,
    now: new Date()
  };
}

/** Utility: case-insensitive header map */
function CORE_headerIndexMap_(headerRow) {
  const map = new Map();
  headerRow.forEach((h, i) => map.set(String(h || '').trim().toLowerCase(), i));
  return map;
}

/** Utility: safe get by header */
function CORE_idx_(hmap, name) {
  const key = String(name || '').trim().toLowerCase();
  if (!hmap.has(key)) return -1;
  return hmap.get(key);
}

/** Utility: coerce number */
function CORE_num_(v) {
  if (v == null || v === '' || isNaN(v)) return 0;
  return Number(v);
}

/** Utility: coerce string */
function CORE_str_(v) { return (v == null) ? '' : String(v).trim(); }

/** Utility: sanitize sheet/tab name */
function CORE_safeSheetName_(name) {
  const bad = /[\\/?*[\]]/g;
  let n = String(name || '').replace(bad, ' ');
  if (n.length > 99) n = n.substring(0, 99);
  return n || 'Sheet';
}

/** Utility: convert to Proper Case (handles hyphens) */
function CORE_properCase_(s) {
  const str = (s == null) ? '' : String(s);
  return str.split(' ').map(word =>
    word.split('-').map(part =>
      part ? part.charAt(0).toUpperCase() + part.slice(1).toLowerCase() : part
    ).join('-')
  ).join(' ');
}

/** Normalization key: lowercase, drop "the", remove non-alphanum */
function CORE_normKey_(s) {
  const x = (s == null ? '' : String(s)).toLowerCase();
  return x.replace(/\bthe\b/g, '').replace(/[^a-z0-9]+/g, '');
}

let __ALIAS_SOC_REV = null;
let __ALIAS_COUNTRY_REV = null;

function CORE_buildReverseAliasMap_(groups) {
  const map = new Map();
  (groups || []).forEach(g => {
    const can = g.canonical;
    const canKey = CORE_normKey_(can);
    map.set(canKey, can);
    (g.aliases || []).forEach(a => map.set(CORE_normKey_(a), can));
  });
  return map;
}

/** Canonical society label (for reading/outputs & lookups) */
function CORE_canonSociety_(s) {
  if (!__ALIAS_SOC_REV) __ALIAS_SOC_REV = CORE_buildReverseAliasMap_(CONFIG.SOCIETY_ALIASES);
  const key = CORE_normKey_((s || '').trim());
  return __ALIAS_SOC_REV.get(key) || (s || '');
}

/** Canonical country/territory label */
function CORE_canonCountry_(s) {
  if (!__ALIAS_COUNTRY_REV) __ALIAS_COUNTRY_REV = CORE_buildReverseAliasMap_(CONFIG.COUNTRY_ALIASES);
  const key = CORE_normKey_(s || '');
  return __ALIAS_COUNTRY_REV.get(key) || (s || '');
}

/** For the qualification engine: return ALL normalized keys that should map to the same society row */
function CORE_getSocietyAliasKeys_(label) {
  const target = CORE_normKey_(label);
  const out = new Set([target]);
  (CONFIG.SOCIETY_ALIASES || []).forEach(g => {
    const all = [g.canonical].concat(g.aliases || []);
    const normed = all.map(CORE_normKey_);
    if (normed.includes(target)) normed.forEach(k => out.add(k));
  });
  return Array.from(out);
}

/**
 * Global status classifier used by Performer tabs, Overview, and Society Issues.
 * Inputs are already in the comparison currency (i.e., the same unit for income vs expected).
 */
function LOGIC_classifyRow_(ctx, { registrationStatus, hasBenchmark, qualifiedPlays, income, expected }) {
  const expectSet   = ctx.config.EXPECTING_STATUS_SET; // lower-cased strings
  const spinThresh  = ctx.config.MIN_QUAL_PLAYS_FOR_EXPECTATION || 20;
  const underMax    = ctx.config.THRESHOLDS.UNDER_MAX; // e.g. -0.30
  const overMin     = ctx.config.THRESHOLDS.OVER_MIN;  // e.g. +0.50

  const reg = String(registrationStatus || '').toLowerCase();
  const expecting = expectSet.has(reg);

  // Default result
  let status = ctx.config.STATUS.NOT_EXPECTED;
  let investigate = false;
  let notes = '';

  if (!expecting) {
    // Registration says we don't expect income
    return { status, investigate, notes: 'Registration does not expect income.' };
  }

  const inc = Number(income)   || 0;
  const exp = Number(expected) || 0;

  // If we have a benchmark but spins are below threshold => Not Expected (new rule)
  if (hasBenchmark && (Number(qualifiedPlays) || 0) < spinThresh) {
    return { status: ctx.config.STATUS.ONTRACK, investigate: false, notes: `Benchmark present; below ${spinThresh} qualifying spins.` };
  }

  // No benchmark path
  if (!hasBenchmark) {
    if ((Number(qualifiedPlays) || 0) >= spinThresh) {
      if (inc <= 0) {
        return { status: ctx.config.STATUS.MISSING, investigate: true, notes: 'No benchmark; ≥ threshold spins; income missing.' };
      }
      return { status: ctx.config.STATUS.ONTRACK, investigate: false, notes: 'No benchmark; ≥ threshold spins; income present.' };
    }
    return { status: ctx.config.STATUS.ONTRACK, investigate: false, notes: 'No benchmark; below spin threshold.' };
  }

  // With benchmark (qualified spins >= threshold at this point)
  if (exp <= 0) {
    // Edge case: zero expected even with benchmark (e.g., no qualifying spins after gates)
    if (inc > 0) return { status: ctx.config.STATUS.ONTRACK, investigate: false, notes: 'No qualifying spins; income present.' };
    return { status: ctx.config.STATUS.NOT_EXPECTED, investigate: false, notes: 'No qualifying spins; no income expected.' };
  }

  // *** New rule: when there IS a benchmark and NO income, classify as Missing (not Underperforming) ***
  if (inc <= 0) {
    return { status: ctx.config.STATUS.MISSING, investigate: true, notes: 'Benchmark present; income missing.' };
  }

  const variancePct = (inc / exp) - 1;
  if (variancePct <= underMax) {
    return { status: ctx.config.STATUS.UNDER, investigate: true, notes: 'Underperforming vs expected.' };
  }
  if (variancePct >= overMin) {
    return { status: ctx.config.STATUS.OVER, investigate: false, notes: 'Overperforming.' };
  }
  return { status: ctx.config.STATUS.ONTRACK, investigate: false, notes: 'On track.' };
}

