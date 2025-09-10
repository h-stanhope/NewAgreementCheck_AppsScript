/** Reads & indices */

function READ_sheet_(ctx, name) {
  const sh = ctx.ss.getSheetByName(name);
  if (!sh) throw new Error(`Missing sheet: ${name}`);
  const range = sh.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return { headers: [], rows: [] };
  const headers = values[0];
  const rows = values.slice(1);
  return { headers, rows };
}

/** Contributor List: A=Contributor ID | B=Contributor Name */
function READ_readContributors_(ctx) {
  const { headers, rows } = READ_sheet_(ctx, ctx.config.SHEETS.CONTRIBUTORS);
  const map = new Map();
  // Be tolerant to header names
  const hmap = CORE_headerIndexMap_(headers);
  const iId = CORE_idx_(hmap, 'contributor id') !== -1 ? CORE_idx_(hmap, 'contributor id') : 0;
  const iName = CORE_idx_(hmap, 'contributor name') !== -1 ? CORE_idx_(hmap, 'contributor name') : 1;

  for (const r of rows) {
    const id = CORE_str_(r[iId]);
    if (!id) continue;
    map.set(id, { id, name: CORE_str_(r[iName]) || id });
  }
  return map;
}

/** Salesforce Statuses */
function READ_readSalesforceStatuses_(ctx) {
  const { headers, rows } = READ_sheet_(ctx, ctx.config.SHEETS.STATUSES);
  const h = CORE_headerIndexMap_(headers);
  const idx = {
    contributorName: CORE_idx_(h, 'contributor: contibutor name') !== -1 ? CORE_idx_(h, 'contributor: contibutor name') : CORE_idx_(h, 'contributor: contributor name'),
    uuid: CORE_idx_(h, 'contributor uuid'),
    society: CORE_idx_(h, 'society: society name'),
    status: CORE_idx_(h, 'status')
  };
  const map = new Map(); // Map<uuid, SocietyStatus[]>
  for (const r of rows) {
    const uuid = CORE_str_(r[idx.uuid]);
    const societyRaw = CORE_str_(r[idx.society]);
    if (!uuid || !societyRaw) continue;
    const entry = {
      contributorName: CORE_str_(r[idx.contributorName]),
      uuid,
      societyName: CORE_canonSociety_(societyRaw),
      status: CORE_str_(r[idx.status]).toLowerCase()
    };
    if (!map.has(uuid)) map.set(uuid, []);
    map.get(uuid).push(entry);
  }
  return map;
}

/** Society Home: Society | Home Territory */
function READ_readSocietyHome_(ctx) {
  const { headers, rows } = READ_sheet_(ctx, ctx.config.SHEETS.SOC_HOME);
  const h = CORE_headerIndexMap_(headers);
  const iSoc = CORE_idx_(h, 'society') !== -1 ? CORE_idx_(h, 'society') : 0;
  const iTerr = CORE_idx_(h, 'home territory') !== -1 ? CORE_idx_(h, 'home territory') : 1;
  const map = new Map(); // Map<societyCanonical, territoryCanonical>
  for (const r of rows) {
    const s = CORE_canonSociety_(CORE_str_(r[iSoc]));
    if (!s) continue;
    const terr = CORE_canonCountry_(CORE_str_(r[iTerr]));
    map.set(s, terr);
  }
  return map;
}

/**
 * Content (NEW SCHEMA)
 * Headers:
 *   Contributor ID | Contributor Name | NR ID | Recording Title | Recording Version
 *   Primary Artist Displayed As | ISRC | Number of Primary Performers
 *   Contributor Type | Roles
 *   Country of Recording 1 | Country of Recording 2 | Country of Contribution
 *   Country of Label Which Funded Recording | Country of Mastering
 *
 * Returns an index object with:
 *   - byIsrc, byNrId, byUuidIsrc (Maps)
 *   - findForPlay(uuid, isrc, nrId)
 *   - enrichPlay(play, uuid)  // fills missing qual/identity fields from Content
 *
 * Back-compat: exposes a .get(key) that mimics the old Map lookup (tries ISRC, then NR ID),
 * returning an object with { numPrimary, roles, contributorType } like before.
 */
function READ_readContent_(ctx) {
  const { headers, rows } = READ_sheet_(ctx, ctx.config.SHEETS.CONTENT);
  if (!headers.length || !rows.length) return CONTENT_emptyIndex_();

  const h = CORE_headerIndexMap_(headers);

  function pick(...names) {
    for (const n of names) {
      const i = CORE_idx_(h, n);
      if (i !== -1) return i;
    }
    return -1;
  }

  const idx = {
    uuid: pick('contributor id'),
    contributorName: pick('contributor name'),
    nrId: pick('nr id'),
    recordingTitle: pick('recording title'),
    recordingVersion: pick('recording version'),
    primaryArtist: pick('primary artist displayed as'),
    isrc: pick('isrc'),
    numPrimary: pick('number of primary performers'),
    contributorType: pick('contributor type'),
    roles: pick('roles'),
    cor1: pick('country of recording 1'),
    cor2: pick('country of recording 2'),
    coc: pick('country of contribution'),
    col: pick('country of label which funded recording'),
    com: pick('country of mastering')
  };

  const byIsrc = new Map();
  const byNrId = new Map();
  const byUuidIsrc = new Map();

  // Back-compat shim (old callers used a Map keyed by ISRC or Contribution ID)
  const legacyMap = new Map();

  function pushMap(map, key, obj) {
    if (!key) return;
    const k = String(key).trim();
    if (!k) return;
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(obj);
  }

  for (const r of rows) {
    const rec = {
      uuid: CORE_str_(r[idx.uuid]),
      contributorName: CORE_str_(r[idx.contributorName]),
      nrId: CORE_str_(r[idx.nrId]),
      recordingTitle: CORE_str_(r[idx.recordingTitle]),
      recordingVersion: CORE_str_(r[idx.recordingVersion]),
      primaryArtist: CORE_str_(r[idx.primaryArtist]),
      isrc: CORE_str_(r[idx.isrc]),
      numPrimary: CORE_num_(r[idx.numPrimary]) || null,
      contributorType: CORE_str_(r[idx.contributorType]).toUpperCase(),
      roles: CORE_str_(r[idx.roles]),
      cor1: CORE_canonCountry_(CORE_str_(r[idx.cor1])),
      cor2: CORE_canonCountry_(CORE_str_(r[idx.cor2])),
      coc: CORE_canonCountry_(CORE_str_(r[idx.coc])),
      col: CORE_canonCountry_(CORE_str_(r[idx.col])),
      com: CORE_canonCountry_(CORE_str_(r[idx.com]))
    };

    // Indices
    pushMap(byIsrc, rec.isrc, rec);
    pushMap(byNrId, rec.nrId, rec);
    pushMap(byUuidIsrc, `${rec.uuid}||${rec.isrc}`, rec);

    // Legacy compatibility map (old code expects numPrimary/roles/contributorType by ISRC; Contribution ID no longer exists in Content)
    if (rec.isrc) {
      legacyMap.set(rec.isrc, {
        numPrimary: rec.numPrimary,
        roles: rec.roles,
        contributorType: rec.contributorType
      });
    }
    if (rec.nrId) {
      legacyMap.set(rec.nrId, {
        numPrimary: rec.numPrimary,
        roles: rec.roles,
        contributorType: rec.contributorType
      });
    }
  }

  function findForPlay(uuid, isrc, nrId) {
    // strongest: uuid+isrc
    const c1 = byUuidIsrc.get(`${uuid}||${isrc}`); if (c1 && c1.length) return c1[0];
    // then isrc
    const c2 = byIsrc.get(isrc); if (c2 && c2.length) return c2[0];
    // then nr id
    const c3 = byNrId.get(nrId); if (c3 && c3.length) return c3[0];
    return null;
  }

  function enrichPlay(play, uuid) {
    const p = Object.assign({}, play);
    const isrc = CORE_str_(p.isrc);
    const nrId = CORE_str_(p.nrId || p.nrID || p.nr_id);
    const hit = findForPlay(uuid, isrc, nrId);

    if (hit) {
      // identity
      p.recordingTitle = p.recordingTitle || hit.recordingTitle;
      p.version        = p.version || p.recordingVersion || hit.recordingVersion;
      p.primaryArtist  = p.primaryArtist || p.primaryArtistDisplayedAs || hit.primaryArtist;

      // splits & roles
      p.contributorType = (p.contributorType || hit.contributorType || '').toString().toUpperCase();
      p.roles           = p.roles || hit.roles;
      p.numPrimary      = Number(p.numPrimary || p.numberOfPrimaryPerformers || hit.numPrimary || 1) || 1;

      // qualification countries
      p.cor1 = p.cor1 || hit.cor1;
      p.cor2 = p.cor2 || hit.cor2;
      p.coc  = p.coc  || hit.coc;
      p.col  = p.col  || hit.col;
      p.com  = p.com  || hit.com;
    }
    return p;
  }

  // Return a content index object (plus back-compat .get)
  return {
    byIsrc,
    byNrId,
    byUuidIsrc,
    findForPlay,
    enrichPlay,
    // Back-compat: allow contentByKey.get(key) like old Map
    get: function(key) { return legacyMap.get(String(key || '')); }
  };
}

function CONTENT_emptyIndex_() {
  return {
    byIsrc: new Map(),
    byNrId: new Map(),
    byUuidIsrc: new Map(),
    findForPlay: function(){ return null; },
    enrichPlay: function(p){ return p; },
    get: function(){ return undefined; }
  };
}

/** Chart Metric - Radio Play */
function READ_readChartMetric_(ctx) {
  const { headers, rows } = READ_sheet_(ctx, ctx.config.SHEETS.CHART);
  const h = CORE_headerIndexMap_(headers);
  const idx = {
    isrc: CORE_idx_(h, 'isrc'),
    nrId: CORE_idx_(h, 'nr id'),                               // NEW: NR ID
    contributorId: CORE_idx_(h, 'contributor id'),
    contributionId: CORE_idx_(h, 'contribution id'),
    contributorType: CORE_idx_(h, 'contributor type'),
    numPrimary: CORE_idx_(h, 'number of primary performers'),
    countryName: CORE_idx_(h, 'country name'),
    playYear: CORE_idx_(h, 'play year'),
    cor1: CORE_idx_(h, 'country of recording 1'),
    cor2: CORE_idx_(h, 'country of recording 2'),
    coc: CORE_idx_(h, 'country of contribution'),
    com: CORE_idx_(h, 'country of mastering'),
    col: CORE_idx_(h, 'country of label which funded recording'),
    roles: CORE_idx_(h, 'roles'),
    playCount: CORE_idx_(h, 'play count'),
    songName: CORE_idx_(h, 'song name'),                       // optional niceties
    recordingVersion: CORE_idx_(h, 'recording version'),
    primaryArtistDisplayedAs: CORE_idx_(h, 'primary artist displayed as')
  };
  const out = [];
  for (const r of rows) {
    const uuid = CORE_str_(r[idx.contributorId]);
    if (!uuid) continue;
    out.push({
      uuid,
      isrc: CORE_str_(r[idx.isrc]),
      nrId: CORE_str_(r[idx.nrId]),                            // NEW
      contributionId: CORE_str_(r[idx.contributionId]),
      countryName: CORE_canonCountry_(CORE_str_(r[idx.countryName])),
      playYear: CORE_str_(r[idx.playYear]),
      contributorType: CORE_str_(r[idx.contributorType]),
      numPrimary: CORE_num_(r[idx.numPrimary]) || 1,
      roles: CORE_str_(r[idx.roles]),
      cor1: CORE_canonCountry_(CORE_str_(r[idx.cor1])),
      cor2: CORE_canonCountry_(CORE_str_(r[idx.cor2])),
      coc: CORE_canonCountry_(CORE_str_(r[idx.coc])),
      com: CORE_canonCountry_(CORE_str_(r[idx.com])),
      col: CORE_canonCountry_(CORE_str_(r[idx.col])),
      playCount: CORE_num_(r[idx.playCount]),
      // identity helpers if present
      recordingTitle: CORE_str_(r[idx.songName]),
      recordingVersion: CORE_str_(r[idx.recordingVersion]),
      primaryArtistDisplayedAs: CORE_str_(r[idx.primaryArtistDisplayedAs])
    });
  }
  return out;
}

/** Benchmarks: Territory | Society | RatePerSpin | Currency | TrackLevelAvailable | KeyTerritory */
function READ_readBenchmarks_(ctx) {
  const { headers, rows } = READ_sheet_(ctx, ctx.config.SHEETS.BENCHMARKS);
  const h = CORE_headerIndexMap_(headers);
  const idx = {
    territory: CORE_idx_(h, 'territory'),
    society: CORE_idx_(h, 'society'),
    rate: CORE_idx_(h, 'rateperspin'),
    currency: CORE_idx_(h, 'currency'),
    trackLevel: CORE_idx_(h, 'tracklevelavailable'),
    keyTerritory: CORE_idx_(h, 'keyterritory')
  };
  const map = new Map(); // key `${territory}||${society}`
  for (const r of rows) {
    const terr = CORE_canonCountry_(CORE_str_(r[idx.territory]));
    const soc = CORE_canonSociety_(CORE_str_(r[idx.society]));
    if (!terr || !soc) continue;
    map.set(`${terr}||${soc}`, {
      territory: terr,
      society: soc,
      rate: Number(r[idx.rate]) || 0,
      currency: CORE_str_(r[idx.currency]),
      trackLevel: CORE_str_(r[idx.trackLevel]),
      keyTerritory: CORE_str_(r[idx.keyTerritory])
    });
  }
  return map;
}

/** FX Rates: FROM | TO | Concat | STATEMENT_PERIOD_ID | RATE */
function READ_readFxRates_(ctx) {
  const { headers, rows } = READ_sheet_(ctx, ctx.config.SHEETS.FX);
  const h = CORE_headerIndexMap_(headers);
  const idx = {
    from: CORE_idx_(h, 'from_currency_code'),
    to: CORE_idx_(h, 'to_currency_code'),
    period: CORE_idx_(h, 'statement_period_id'),
    rate: CORE_idx_(h, 'rate')
  };
  const exact = new Map();  // key `${from}||${to}||${period}` => rate
  const latest = new Map(); // key `${from}||${to}` => rate (last seen considered latest)
  for (const r of rows) {
    const from = CORE_str_(r[idx.from]);
    const to = CORE_str_(r[idx.to]);
    const period = CORE_str_(r[idx.period]);
    const rate = Number(r[idx.rate]) || 0;
    if (!from || !to || !rate) continue;
    exact.set(`${from}||${to}||${period}`, rate);
    latest.set(`${from}||${to}`, rate); // updated each row; last seen = latest
  }
  return { exact, latest };
}

/** Generated Income Report (with leading Statement Period ID)
 * Reads track identifiers + identity so we can cluster by Title + Main Artist.
 */
function READ_readIncome_(ctx) {
  const { headers, rows } = READ_sheet_(ctx, ctx.config.SHEETS.INCOME);
  const h = CORE_headerIndexMap_(headers);

  function col(...names) {
    for (const n of names) {
      const i = CORE_idx_(h, n);
      if (i !== -1) return i;
    }
    return -1;
  }

  const idx = {
    periodId:        col('statement period id'),
    contributorId:   col('contributor id'),
    contributorName: col('contributor name'),
    societyName:     col('society name'),
    countryName:     col('country name'),

    // identity for title/artist clustering
    recordingTitle:  col('recording title','song name'),
    recordingVersion:col('recording version'),
    mainArtist:      col('main artist','primary artist displayed as'),

    // currencies & amounts
    salesCcy:        col('sales currency'),
    payeeCcy:        col('payee currency'),
    totalUSD:        col('total usd'),
    amtSales:        col('gross revenue after withholding tax sales currency'),
    amtPayee:        col('gross revenue after withholding tax payee currency'),

    // track identifiers
    isrc:            col('isrc'),
    contributionId:  col('contribution id'),
    soundRecordingId:col('sound recording id')
  };

  const out = [];
  for (const r of rows) {
    const uuid = CORE_str_(r[idx.contributorId]);
    if (!uuid) continue;

    out.push({
      uuid,
      societyName: CORE_canonSociety_(CORE_str_(r[idx.societyName])),
      countryName: CORE_canonCountry_(CORE_str_(r[idx.countryName])),

      recordingTitle: CORE_str_(r[idx.recordingTitle]),
      recordingVersion: CORE_str_(r[idx.recordingVersion]),
      mainArtist: CORE_str_(r[idx.mainArtist]),

      payeeCcy: CORE_str_(r[idx.payeeCcy]),
      salesCcy: CORE_str_(r[idx.salesCcy]),
      periodId: CORE_str_(r[idx.periodId]),

      amountPayee: Number(r[idx.amtPayee]) || 0,
      amountSales: Number(r[idx.amtSales]) || 0,

      isrc: CORE_str_(r[idx.isrc]),
      contributionId: CORE_str_(r[idx.contributionId]),
      soundRecordingId: CORE_str_(r[idx.soundRecordingId])
    });
  }
  return out;
}


