/**
 * Qualification Engine (baked-in)
 * --------------------------------
 * Implements NRQualification.qualifies(criteria) using the central reference workbook,
 * with society/country alias tolerance and an explain() API for diagnostics.
 *
 * criteria: {
 *   society, territory,
 *   contributorType,            // MAIN_PERFORMER | FEATURING_PERFORMER | SESSION_MUSICIAN
 *   roles,                      // comma-separated roles string
 *   cor1, cor2,                 // Country of Recording 1/2
 *   coc,                        // Country of Contribution (producer/label concept)
 *   com,                        // Country of Mastering
 *   col,                        // Country of Label Which Funded Recording
 *   isrc, contributionId        // (not used in rules; future-proof)
 * }
 */

var NRQualification = (function () {
  // ====== Cache (per execution) ======
  var _loaded = false;
  var _qualRows = [];                    // qualification sheet rows
  var _socToRow = new Map();             // society -> row index (exact label)
  var _socNormKeyToRow = new Map();      // normalized keys (including aliases) -> row index
  var _countrySets = new Map();          // header -> Set(countryLower)
  var _excludedRoles = new Map();        // society -> Set(role)
  var _qYearMap = new Map();             // society -> Map(countryLower -> year)
  var _currentYear = new Date().getFullYear();

  // ====== Debug toggle ======
  function DEBUG() { return !!(CONFIG && CONFIG.QUAL && CONFIG.QUAL.DEBUG); }
  function log(){ if (DEBUG()) Logger.log.apply(Logger, arguments); }

  // ====== Load central reference once ======
  function ensureLoaded_() {
    if (_loaded) return;

    var url = CONFIG.QUAL.CENTRAL_URL;
    if (!url) throw new Error('CONFIG.QUAL.CENTRAL_URL is not set.');

    var book = SpreadsheetApp.openByUrl(url);

    var shQual     = book.getSheetByName(CONFIG.QUAL.SHEETS.QUALIFICATION);
    var shCountries= book.getSheetByName(CONFIG.QUAL.SHEETS.COUNTRY_LIST);
    var shExcluded = book.getSheetByName(CONFIG.QUAL.SHEETS.EXCLUDED_ROLES);
    var shYear     = book.getSheetByName(CONFIG.QUAL.SHEETS.QUAL_YEAR);

    if (!shQual || !shCountries || !shExcluded || !shYear) {
      throw new Error('One or more Qualification reference sheets are missing in the central workbook.');
    }

    // Load Qualification rows
    _qualRows = shQual.getDataRange().getValues(); // row[0]=headers
    _socToRow.clear();
    _socNormKeyToRow.clear();

    for (var r = 1; r < _qualRows.length; r++) {
      var soc = (_qualRows[r][0] || '').toString().trim();
      if (!soc) continue;
      _socToRow.set(soc, r);

      // Exact normalized label
      var baseKey = CORE_normKey_(soc);
      _socNormKeyToRow.set(baseKey, r);

      // Map known aliases to the same row
      var aliasKeys = CORE_getSocietyAliasKeys_(soc); // includes baseKey
      aliasKeys.forEach(function(k){ _socNormKeyToRow.set(k, r); });
    }

    // Preload supporting maps/sets
    _countrySets   = buildCountrySets_(shCountries);      // header -> Set(countryLower)
    _excludedRoles = buildExcludedRolesMap_(shExcluded);  // society -> Set(role)
    _qYearMap      = buildQualificationYearMap_(shYear);  // society -> Map(countryLower, year)

    _loaded = true;
  }

  // ====== Fast helpers (adapted from provided code) ======
  function buildCountrySets_(countryListSheet) {
    var data = countryListSheet.getDataRange().getValues();
    if (!data || !data.length) return new Map();
    var headers = data[0].map(function (h) { return (h || '').toString().trim(); });
    var sets = new Map();
    for (var c = 0; c < headers.length; c++) {
      var name = headers[c]; if (!name) continue;
      var set = new Set();
      for (var r = 1; r < data.length; r++) {
        var val = (data[r][c] || '').toString().trim().toLowerCase();
        if (val) set.add(val);
      }
      sets.set(name, set);
    }
    return sets;
  }

  function countryInCategoryFast_(country, category) {
    var set = _countrySets.get(category);
    var key = (country || '').toString().trim().toLowerCase();
    return !!(set && set.has(key));
  }

  // + / - modifiers with preloaded sets
  function checkSingleCountryWithModifiersFast_(country, listCategory, society) {
    if (!country) return false;
    var parts = (listCategory || '').toString().split(' ');
    var baseList = parts[0] || '';
    var modifier = parts.slice(1).join(' ') || '';

    var c = (country || '').toString().trim();

    var inSocMinus = countryInCategoryFast_(c, society + ' -');
    var inSocPlus  = countryInCategoryFast_(c, society + ' +');
    var inBase     = countryInCategoryFast_(c, baseList);

    if (modifier.indexOf('+') !== -1 && modifier.indexOf('-') !== -1) {
      if (inSocMinus) return false;
      if (inSocPlus)  return true;
      return inBase;
    }
    if (modifier.indexOf('-') !== -1) {
      if (inSocMinus) return false;
      return inBase;
    }
    if (modifier.indexOf('+') !== -1) {
      if (inSocPlus)  return true;
      return inBase;
    }
    return inBase;
  }

  function checkCountryListWithModifiersFast_(countries, listCategory, society) {
    if (!countries || !countries.length || !listCategory) return false;
    for (var i = 0; i < countries.length; i++) {
      if (checkSingleCountryWithModifiersFast_(countries[i], listCategory, society)) {
        return true;
      }
    }
    return false;
  }

  function buildExcludedRolesMap_(excludedRolesSheet) {
    var data = excludedRolesSheet.getDataRange().getValues();
    var map = new Map();
    if (!data || !data.length) return map;
    var headers = data[0].map(function (h) { return (h || '').toString().trim(); });
    for (var c = 0; c < headers.length; c++) {
      var society = headers[c];
      if (!society) continue;
      var set = new Set();
      for (var r = 1; r < data.length; r++) {
        var role = (data[r][c] || '').toString().trim();
        if (role) set.add(role);
      }
      map.set(society, set);
    }
    return map;
  }

  function buildQualificationYearMap_(qualificationYearSheet) {
    var data = qualificationYearSheet.getDataRange().getValues();
    var map = new Map();
    if (!data || data.length < 2) return map;
    for (var i = 1; i < data.length; i++) {
      var society = (data[i][0] || '').toString().trim();
      var country = (data[i][1] || '').toString().trim().toLowerCase();
      var year = Number(data[i][2]) || 0;
      if (!society || !country || !year) continue;
      if (!map.has(society)) map.set(society, new Map());
      map.get(society).set(country, year);
    }
    return map;
  }

  // Returns earliest qualifying year (<= current year) or null
  function checkQualificationYearFast_(society, arr1, arr2, arr3) {
    var socMap = _qYearMap.get(society);
    if (!socMap) return null;
    var all = new Set(
      [].concat(arr1 || [], arr2 || [], arr3 || [])
        .map(function (x) { return (x || '').toString().trim().toLowerCase(); })
        .filter(Boolean)
    );
    var minYear = null;
    all.forEach(function (c) {
      var yr = socMap.get(c);
      if (yr && _currentYear >= yr) {
        minYear = (minYear === null ? yr : Math.min(minYear, yr));
      }
    });
    return minYear;
  }

  function parseList_(val) {
    if (!val) return [];
    return val.toString().split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  }

  // ====== Core evaluate logic (shared) ======
// ====== Core evaluate logic (shared) ======
function _evaluate(criteria, wantExplain) {
  ensureLoaded_();

  var societyIn = (criteria.society || '').toString().trim();
  if (!societyIn) {
    return wantExplain ? { qualified: false, rule: 'No society provided', societyRowFound: false, notes: ['No society provided'] } : { qualified: false };
  }

  // Alias-aware row lookup
  var rowIdx = _socNormKeyToRow.get(CORE_normKey_(societyIn));
  if (rowIdx == null) {
    var canon = CORE_canonSociety_(societyIn);
    rowIdx = _socNormKeyToRow.get(CORE_normKey_(canon));
  }
  if (rowIdx == null) {
    var objNF = { qualified: false, rule: 'Society not in matrix', societyRowFound: false, notes: ['Society not found in qualification matrix: ' + societyIn] };
    return wantExplain ? objNF : { qualified: false };
  }
  var row = _qualRows[rowIdx];

  // Column semantics
  var allQualify                  = row[2];   // "Yes" means everything qualifies (but still must pass Featured/Roles)
  var societyCitizenship          = row[3];
  var societyResidency            = row[4];
  var countryOfProducer           = row[5];   // label/producer list category
  var countryOfRecordingList      = row[6];
  var societyCountryOfPerformance = row[8];
  var countryOfMasteringList      = row[9];
  var featuredOnly                = row[10];  // "Featured" | "Non-Featured" | ""

  // Normalize input fields
  var contributorTypeRaw = (criteria.contributorType || '').toString().trim().toUpperCase();
  var rolesStr = (criteria.roles || '').toString();

  var countryOfRecording1 = parseList_(criteria.cor1);
  var countryOfRecording2 = parseList_(criteria.cor2);
  var countryOfRecording3 = []; // not present in our dataset
  var countryOfLabel      = parseList_(criteria.col);
  var masteringTerritory  = parseList_(criteria.com);
  var countryOfPerformance= []; // not present in our dataset
  var citizenship         = []; // not present in our dataset
  var residency           = []; // not present in our dataset

  var checks = {
    ADAMI_EEA:false, ADAMI_RC_CoL:false,
    Citizenship:false, Residency:false, CoL:false, CoR:false, CoP:false, CoM:false
  };

  // We will set qualifiesFlag/ruleHit and STILL run the Featured/Roles gates afterwards.
  var qualifiesFlag = false;
  var ruleHit = '';

  // 1) All Qualify -> mark as qualified by gate, but DO NOT return yet.
  if (allQualify === 'Yes') {
    qualifiesFlag = true;
    ruleHit = 'All Qualify';
  }

  // 2) ADAMI special case (if no gate hit yet)
  if (!qualifiesFlag && societyIn === 'ADAMI') {
    var inEEA =
      checkCountryListWithModifiersFast_(countryOfRecording1, 'EEA', 'ADAMI') ||
      checkCountryListWithModifiersFast_(countryOfRecording2, 'EEA', 'ADAMI') ||
      checkCountryListWithModifiersFast_(countryOfRecording3, 'EEA', 'ADAMI');

    var inRConCoR =
      checkCountryListWithModifiersFast_(countryOfRecording1, 'RC', 'ADAMI') ||
      checkCountryListWithModifiersFast_(countryOfRecording2, 'RC', 'ADAMI') ||
      checkCountryListWithModifiersFast_(countryOfRecording3, 'RC', 'ADAMI');

    var inRConCoL = checkCountryListWithModifiersFast_(countryOfLabel, 'RC', 'ADAMI');

    if (inEEA) { qualifiesFlag = true; ruleHit = 'ADAMI: CoR (EEA)'; checks.ADAMI_EEA = true; }
    else if (inRConCoR && inRConCoL) { qualifiesFlag = true; ruleHit = 'ADAMI: CoR+CoL (RC)'; checks.ADAMI_RC_CoL = true; }
  }

  // 3) Standard progressive checks (stop at first success)
  if (!qualifiesFlag && societyCitizenship && citizenship.length &&
      checkCountryListWithModifiersFast_(citizenship, societyCitizenship, societyIn)) {
    qualifiesFlag = true; ruleHit = 'Citizenship'; checks.Citizenship = true;
  }
  if (!qualifiesFlag && societyResidency && residency.length &&
      checkCountryListWithModifiersFast_(residency, societyResidency, societyIn)) {
    qualifiesFlag = true; ruleHit = 'Residency'; checks.Residency = true;
  }
  if (!qualifiesFlag && countryOfProducer && countryOfLabel.length &&
      checkCountryListWithModifiersFast_(countryOfLabel, countryOfProducer, societyIn)) {
    qualifiesFlag = true; ruleHit = 'CoL'; checks.CoL = true;
  }
  var anyRecordingCountries = (countryOfRecording1.length + countryOfRecording2.length + countryOfRecording3.length) > 0;
  if (!qualifiesFlag && anyRecordingCountries && countryOfRecordingList &&
      ( checkCountryListWithModifiersFast_(countryOfRecording1, countryOfRecordingList, societyIn) ||
        checkCountryListWithModifiersFast_(countryOfRecording2, countryOfRecordingList, societyIn) ||
        checkCountryListWithModifiersFast_(countryOfRecording3, countryOfRecordingList, societyIn) )) {
    qualifiesFlag = true; ruleHit = 'CoR'; checks.CoR = true;
  }
  if (!qualifiesFlag && societyCountryOfPerformance && countryOfPerformance.length &&
      checkCountryListWithModifiersFast_(countryOfPerformance, societyCountryOfPerformance, societyIn)) {
    qualifiesFlag = true; ruleHit = 'CoP'; checks.CoP = true;
  }
  if (!qualifiesFlag && countryOfMasteringList && masteringTerritory.length &&
      checkCountryListWithModifiersFast_(masteringTerritory, countryOfMasteringList, societyIn)) {
    qualifiesFlag = true; ruleHit = 'CoM'; checks.CoM = true;
  }

  // 4) Future-qualification via year (CoR-based) — only if still not qualified
  var futureYear = null;
  if (!qualifiesFlag) {
    futureYear = checkQualificationYearFast_(societyIn, countryOfRecording1, countryOfRecording2, countryOfRecording3);
    if (futureYear) { qualifiesFlag = true; ruleHit = 'Future Year (CoR)'; }
  }

  // 5) Featured/Non-Featured constraint AFTER qualification
  var featuredPassed = null;
  if (qualifiesFlag && featuredOnly) {
    if (featuredOnly === 'Featured') {
      featuredPassed = (contributorTypeRaw === 'MAIN_PERFORMER' || contributorTypeRaw === 'FEATURING_PERFORMER');
      if (!featuredPassed) { qualifiesFlag = false; ruleHit += (ruleHit ? ' → ' : '') + 'blocked by Featured requirement'; }
    } else if (featuredOnly === 'Non-Featured') {
      featuredPassed = (contributorTypeRaw === 'SESSION_MUSICIAN');
      if (!featuredPassed) { qualifiesFlag = false; ruleHit += (ruleHit ? ' → ' : '') + 'blocked by Non-Featured requirement'; }
    } else {
      featuredPassed = true;
    }
  }

  // 6) Roles exclusion AFTER featured check
  var rolesExcluded = false;
  if (qualifiesFlag && rolesStr) {
    var roles = rolesStr.split(',').map(function (r) { return r.trim(); }).filter(Boolean);
    var excludedSet = _excludedRoles.get(societyIn);
    if (excludedSet && roles.length) {
      var allExcluded = roles.every(function (role) { return excludedSet.has(role); });
      if (allExcluded) {
        rolesExcluded = true;
        qualifiesFlag = false;
        ruleHit += (ruleHit ? ' → ' : '') + 'blocked by Roles exclusion';
      }
    }
  }

  if (wantExplain) {
    return {
      qualified: !!qualifiesFlag,
      rule: ruleHit || 'No gate matched',
      societyRowFound: true,
      featuredOnly,
      featuredPassed,
      rolesExcluded,
      futureYear,
      checks,
      notes: []
    };
  }
  return { qualified: !!qualifiesFlag };
}

  // ====== Public API ======
  function qualifies(criteria) {
    return _evaluate(criteria, /*wantExplain*/ false).qualified;
  }

  function explain(criteria) {
    return _evaluate(criteria, /*wantExplain*/ true);
  }

  // Public API
  return { qualifies: qualifies, explain: explain };
})();

/**
 * Thin compatibility wrapper used by the rest of our pipeline.
 */
const QUAL = (function () {
  function qualifies(criteria) {
    return (typeof NRQualification !== 'undefined' && NRQualification && typeof NRQualification.qualifies === 'function')
      ? !!NRQualification.qualifies(criteria) : false;
  }
  function explain(criteria) {
    return (typeof NRQualification !== 'undefined' && NRQualification && typeof NRQualification.explain === 'function')
      ? NRQualification.explain(criteria) : { qualified:false, rule:'', societyRowFound:false, notes:['Explain unavailable'] };
  }
  return { qualifies, explain };
})();

/**
 * Compute total and qualified plays for a contributor in a society's home territory.
 * Returns: { totalPlays, qualifiedPlays, chartSubset }
 */
/**
 * Compute total and qualified plays for a contributor in a society's home territory,
 * using the SAME engine as the Qualification Audit (QUAL.explain).
 * Returns: { totalPlays, qualifiedPlays, chartSubset }
 */
function EXP_computePlaysForTerritory_({ ctx, uuid, society, territory, chartPlays, contentByKey }) {
  // Canonicalize inputs so lookups match the engine’s matrix
  society   = CORE_canonSociety_(society);
  territory = territory ? CORE_canonCountry_(territory) : '';

  let total = 0;
  let qualified = 0;
  const subset = [];

  for (const p of chartPlays) {
    if (p.uuid !== uuid) continue;
    if (territory && p.countryName !== territory) continue;

    const playCount = Number(p.playCount) || 0;
    total += playCount;

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

    // *** Single source of truth ***
    const expl = QUAL.explain(crit);

    const qCount = expl.qualified ? playCount : 0;
    qualified += qCount;

    // Keep rich details for any downstream need (debug, drilldowns, etc.)
    subset.push({
      ...p,
      qualifies: expl.qualified,
      qualifiedCount: qCount,
      rule: expl.rule,
      featuredOnly: expl.featuredOnly,
      featuredPassed: expl.featuredPassed,
      rolesExcluded: expl.rolesExcluded,
      futureYear: expl.futureYear
    });
  }

  return { totalPlays: total, qualifiedPlays: qualified, chartSubset: subset };
}
