/*********************************************************
 * monday.com — Push “Society Issues” rows as sub-items
 * + Diagnostics (rules & lookups)
 **********************************************************/

const MON_CFG = {
  ENDPOINT: 'https://api.monday.com/v2',

  // Main tracking board + target group (Test)
  BOARD_ID: '4762290141',          // Perf - Royalty Tracking Team
  GROUP_ID: 'group_mkvf1ts9',      // "Test"

  // PARENT (main query) columns — specify by Title (ids resolved dynamically)
  MAIN_COLS: {
    owner:      { id: null, title: 'Query Owner' },          // People
    status:     { id: null, title: 'Query Status' },         // Status (color) or Dropdown
    type:       { id: null, title: 'Query Type' },           // Status or Dropdown
    reviewDate: { id: null, title: 'Review Date' }           // Date
  },

  // Direct link columns (from your column dump)
  MAIN_LINK_COL_ID:    'link_mkvmm6z4', // Analysis Query Link (parent)
  SUBITEM_LINK_COL_ID: 'link_mkvm6jq5', // Analysis Sub Query Link (sub-item)

  // Who to set as Query Owner
  DEFAULT_QUERY_OWNER_ID: null,
  DEFAULT_QUERY_OWNER_NAME: 'Harry Stanhope',

  // Sub-item board column IDs (from your subitem columns list)
  SUBITEM_COLS: {
    contributorRelation: 'contributor_list',   // connect_boards → Contributor List
    societyRelation:     'connect_boards',     // connect_boards → Society List
    subitemOwner:        'person',             // people
    subitemStatus:       'status',             // status
    trackingDropdown:    'tracking_sub_query', // dropdown
    reviewDate:          'resolution_due_date' // date
  },

  // Lookup boards
  CONTRIBUTOR_LIST_BOARD_ID: '4740231620',   // Perf - Contributor List
  SOCIETY_LIST_BOARD_ID:     '4739922971',   // Perf - Society List

  // People fallback title on Society List board
  SOCIETY_OWNER_COL_TITLE: 'Society Owner',

  // Relationship gating
  ALLOWED_RELATIONSHIP_STATUSES: ['Good', 'Escalation In Progress'],

  // Source sheets
  SOURCE_SHEET:            'Society Issues',
  RECORDING_ISSUES_SHEET:  'Recording Issues',
  MOST_PLAYED_SHEET:       'Most Played Recordings',
  RULES_SHEET:             'Monday Issues Rules',
  SALESFORCE_SHEET:        'Salesforce Statuses', // should have UUID + (Account/Client Priority)

  // Diagnostics
  DIAG_RULES_SHEET:  'Diag – Rules Map',
  DIAG_PREVIEW_SHEET:'Diag – Preview',
  DIAG_PREVIEW_LIMIT: 200, // max Society Issues rows to preview

  // Batch size for the test runner
  TEST_BATCH_ROWS: 5,

  // Comment generation tuning
  COMMENT: {
    MAX_RECORDING_BULLETS: 5,
    DEFAULT_PERIOD:        'this review period'
  }
};

/* ==============================
   Token + GraphQL (prefixed MI_)
   ============================== */

function MI_token_() {
  const t = PropertiesService.getScriptProperties().getProperty('MONDAY_TOKEN');
  if (!t) throw new Error('Missing Script Property MONDAY_TOKEN (set it in File → Project properties → Script properties).');
  return t.trim();
}

function MI_gql_(query, variables) {
  const res = UrlFetchApp.fetch(MON_CFG.ENDPOINT, {
    method: 'post',
    headers: { Authorization: MI_token_(), 'Content-Type': 'application/json' },
    payload: JSON.stringify({ query, variables: variables || {} }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code >= 400) throw new Error(`monday API HTTP ${code}: ${body}`);
  let json;
  try { json = JSON.parse(body); } catch (e) { throw new Error('monday API: invalid JSON: ' + body); }
  if (json.errors) throw new Error('monday API error: ' + JSON.stringify(json.errors));
  return json.data;
}

/* =============
   Normalisers
   ============= */

function MI_norm_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/\(.*?\)/g, '')
    .replace(/&/g, 'and')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
function MI_stripParen_(s) {
  return String(s || '').replace(/\(.*?\)/g, '').trim();
}
function MI_normLabel_(s) {
  const t = String(s || '').trim();
  if (!t) return '';
  return t.replace(/\s+/g, ' ').trim();
}

/* =========================================================
   Board columns + item column_values helpers (with settings)
   ========================================================= */

function MI_getBoardColumns_(boardId) {
  const q = `
    query ($ids:[ID!]!) {
      boards(ids:$ids) {
        id
        columns { id title type settings_str }
      }
    }`;
  const data = MI_gql_(q, { ids: [String(boardId)] });
  const cols = data?.boards?.[0]?.columns || [];
  return cols.map(c => {
    let settings = {};
    try { settings = c.settings_str ? JSON.parse(c.settings_str) : {}; } catch (e) {}
    return {
      id: String(c.id),
      title: String(c.title || ''),
      type: String(c.type || ''),
      settings
    };
  });
}

function MI_getItemColumnValues_(itemId) {
  const q = `
    query ($ids:[ID!]!) {
      items(ids:$ids) {
        id
        board { id }
        column_values { id type value text }
      }
    }`;
  const data = MI_gql_(q, { ids: [String(itemId)] });
  const it = data?.items?.[0] || null;
  return it ? (it.column_values || []) : [];
}

/* =========================================================
   Society Owner (People) — fetch user id from Society List
   ========================================================= */

function MI_getSocietyOwnerUserId_(societyItemId) {
  const colsMeta = MI_getBoardColumns_(MON_CFG.SOCIETY_LIST_BOARD_ID);
  const metaById = new Map(colsMeta.map(m => [m.id, m]));
  const cvals = MI_getItemColumnValues_(societyItemId);
  const targetTitle = (MON_CFG.SOCIETY_OWNER_COL_TITLE || '').toLowerCase().trim();

  function extractPersonId(val) {
    if (!val || val === 'null') return null;
    try {
      const obj = JSON.parse(val);
      const arr = obj.personsAndTeams || [];
      const hit = arr.find(x => x && x.kind === 'person' && x.id != null);
      return hit ? Number(hit.id) : null;
    } catch (_) { return null; }
  }

  if (targetTitle) {
    const ownerMeta = colsMeta.find(m => m.type === 'people' && m.title.toLowerCase().trim() === targetTitle);
    if (ownerMeta) {
      const cv = cvals.find(v => String(v.id) === ownerMeta.id);
      const pid = cv ? extractPersonId(cv.value) : null;
      if (pid) return pid;
    }
  }
  for (const cv of cvals) {
    const meta = metaById.get(String(cv.id));
    if (!meta || meta.type !== 'people') continue;
    const pid = extractPersonId(cv.value);
    if (pid) return pid;
  }
  for (const cv of cvals) {
    const meta = metaById.get(String(cv.id));
    if (meta && meta.type === 'people') return extractPersonId(cv.value);
  }
  return null;
}

/* ===========================================
   Read “Society Issues” rows (returns objects)
   =========================================== */

function MI_readSocietyIssuesRows_(maxRows) {
  const sh = SpreadsheetApp.getActive().getSheetByName(MON_CFG.SOURCE_SHEET);
  if (!sh) throw new Error(`Missing sheet: ${MON_CFG.SOURCE_SHEET}`);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];

  const headers = vals[0].map(h => String(h || '').toLowerCase().trim());
  const col = (label) => headers.indexOf(String(label).toLowerCase());

  const want = Number(maxRows || MON_CFG.TEST_BATCH_ROWS) || 5;
  const out = [];
  for (let r = 1; r < vals.length && out.length < want; r++) {
    const row = vals[r];
    const contributor = row[col('contributor')];
    const society     = row[col('society')];
    if (!String(contributor || society || '').trim()) continue;

    out.push({
      rowIndex:  r + 1,
      contributor,
      uuid:        row[col('uuid')],
      society,
      territory:   row[col('territory')],
      status:      row[col('status')],
      notes:       row[col('notes')]
    });
  }
  return out;
}

/* ======================================================
   Parent item → find or create + ENSURE main columns set
   ====================================================== */

const __PARENT_FIELDS_SET__ = new Set(); // cache per parent item id
let __MAIN_BOARD_COL_META__ = null;      // cache main board columns

function MI_resolveColumnIdByTitle_(boardCols, title) {
  if (!title) return null;
  const t = String(title).trim().toLowerCase();
  const hit = boardCols.find(c => String(c.title).trim().toLowerCase() === t);
  return hit ? String(hit.id) : null;
}

function MI_findUserIdByName_(name) {
  if (!name) return null;
  const q = `query { users(kind: non_guests, limit: 500) { id name email } }`;
  try {
    const data = MI_gql_(q, {});
    const users = data?.users || [];
    const nm = String(name).trim().toLowerCase();
    const hit = users.find(u => String(u.name || '').trim().toLowerCase() === nm);
    return hit ? Number(hit.id) : null;
  } catch (e) {
    console.log('WARN: users() lookup failed: ' + e);
    return null;
  }
}

function MI_changeColumnValue_(itemId, columnId, valueObj) {
  if (!itemId || !columnId) return;
  const m = `
    mutation ($item: ID!, $board: ID!, $column: String!, $value: JSON!) {
      change_column_value(item_id: $item, board_id: $board, column_id: $column, value: $value) { id }
    }`;
  MI_gql_(m, {
    item:   String(itemId),
    board:  String(MON_CFG.BOARD_ID),
    column: String(columnId),
    value:  JSON.stringify(valueObj)
  });
}

/* --- resolve status index / dropdown id from settings_str (case-insensitive) --- */
function MI_statusIndexFromLabel_(colMeta, wantedLabel) {
  if (!colMeta || !['color','status'].includes(colMeta.type)) return null;
  const settings = colMeta.settings || {};

  const labelsObj = settings.labels || settings.labels_ids || null;
  if (labelsObj && typeof labelsObj === 'object' && !Array.isArray(labelsObj)) {
    const entries = Object.entries(labelsObj);
    const w = String(wantedLabel || '').trim().toLowerCase();
    for (const [idx, label] of entries) {
      if (String(label || '').trim().toLowerCase() === w) return Number(idx);
    }
  }

  const labelsColors = settings.labels_colors;
  if (Array.isArray(labelsColors)) {
    const w = String(wantedLabel || '').trim().toLowerCase();
    const hit = labelsColors.find(x => String(x?.name || '').trim().toLowerCase() === w);
    if (hit && hit.index != null) return Number(hit.index);
  }
  return null;
}
function MI_dropdownIdFromLabel_(colMeta, wantedLabel) {
  if (!colMeta || colMeta.type !== 'dropdown') return null;
  const settings = colMeta.settings || {};
  const arr = settings.labels || [];
  const w = String(wantedLabel || '').trim().toLowerCase();
  for (let i = 0; i < arr.length; i++) {
    if (String(arr[i] || '').trim().toLowerCase() === w) return i; // ids are indexes
  }
  return null;
}

function MI_ensureParentMainQueryFields_(parentItemId) {
  const pid = String(parentItemId);
  if (__PARENT_FIELDS_SET__.has(pid)) return;

  if (!__MAIN_BOARD_COL_META__) __MAIN_BOARD_COL_META__ = MI_getBoardColumns_(MON_CFG.BOARD_ID);
  const cols = __MAIN_BOARD_COL_META__;

  const ownerColId  = MON_CFG.MAIN_COLS.owner.id      || MI_resolveColumnIdByTitle_(cols, MON_CFG.MAIN_COLS.owner.title);
  const statusColId = MON_CFG.MAIN_COLS.status.id     || MI_resolveColumnIdByTitle_(cols, MON_CFG.MAIN_COLS.status.title);
  const typeColId   = MON_CFG.MAIN_COLS.type.id       || MI_resolveColumnIdByTitle_(cols, MON_CFG.MAIN_COLS.type.title);
  const dateColId   = MON_CFG.MAIN_COLS.reviewDate.id || MI_resolveColumnIdByTitle_(cols, MON_CFG.MAIN_COLS.reviewDate.title);

  const statusMeta = cols.find(c => c.id === statusColId) || null;
  const typeMeta   = cols.find(c => c.id === typeColId)   || null;

  // Owner (People)
  let ownerUserId = MON_CFG.DEFAULT_QUERY_OWNER_ID ? Number(MON_CFG.DEFAULT_QUERY_OWNER_ID) : null;
  if (!ownerUserId && MON_CFG.DEFAULT_QUERY_OWNER_NAME) {
    ownerUserId = MI_findUserIdByName_(MON_CFG.DEFAULT_QUERY_OWNER_NAME) || null;
  }
  if (ownerColId && ownerUserId) {
    try {
      MI_changeColumnValue_(pid, ownerColId, { personsAndTeams: [{ id: ownerUserId, kind: 'person' }] });
    } catch (e) {
      console.log(`WARN: could not set Query Owner on ${pid}: ${e}`);
    }
  }

  // Query Status → "Working on it"
  if (statusColId && statusMeta) {
    try {
      if (['status','color'].includes(statusMeta.type)) {
        let idx = MI_statusIndexFromLabel_(statusMeta, 'Working On It');
        if (idx == null) idx = MI_statusIndexFromLabel_(statusMeta, 'Working on it');
        if (idx != null) MI_changeColumnValue_(pid, statusColId, { index: idx });
        else             MI_changeColumnValue_(pid, statusColId, { label: 'Working on it' });
      } else if (statusMeta.type === 'dropdown') {
        const id = MI_dropdownIdFromLabel_(statusMeta, 'Working On It') ?? MI_dropdownIdFromLabel_(statusMeta, 'Working on it');
        if (id != null) MI_changeColumnValue_(pid, statusColId, { ids: [id] });
        else            MI_changeColumnValue_(pid, statusColId, { labels: ['Working on it'] });
      }
    } catch (e) {
      console.log(`WARN: could not set Query Status on ${pid}: ${e}`);
    }
  }

  // Query Type → "Tracking - Contributor"
  if (typeColId && typeMeta) {
    try {
      if (typeMeta.type === 'dropdown') {
        const id = MI_dropdownIdFromLabel_(typeMeta, 'Tracking - Contributor');
        if (id != null) MI_changeColumnValue_(pid, typeColId, { ids: [id] });
        else            MI_changeColumnValue_(pid, typeColId, { labels: ['Tracking - Contributor'] });
      } else if (['status','color'].includes(typeMeta.type)) {
        const idx = MI_statusIndexFromLabel_(typeMeta, 'Tracking - Contributor');
        if (idx != null) MI_changeColumnValue_(pid, typeColId, { index: idx });
        else             MI_changeColumnValue_(pid, typeColId, { label: 'Tracking - Contributor' });
      }
    } catch (e) {
      console.log(`WARN: could not set Query Type on ${pid}: ${e}`);
    }
  }

  // Review Date = +30 days
  if (dateColId) {
    try {
      MI_changeColumnValue_(pid, dateColId, { date: MI_isoDatePlusDays_(30) });
    } catch (e) {
      console.log(`WARN: could not set Review Date on ${pid}: ${e}`);
    }
  }

  __PARENT_FIELDS_SET__.add(pid);
}

/** Optionally set the parent link (Analysis Query Link). */
function MI_setParentLinkIfAny_(parentItemId, url, text) {
  if (!MON_CFG.MAIN_LINK_COL_ID || !url) return;
  try {
    MI_changeColumnValue_(String(parentItemId), MON_CFG.MAIN_LINK_COL_ID, { url: String(url), text: String(text || url) });
  } catch (e) {
    console.log(`WARN: could not set parent link on ${parentItemId}: ${e}`);
  }
}

function MI_findOrCreateParentItem_(societyName, linkUrl, linkText) {
  const wanted = `Bulk Performer Checks – ${String(societyName || '').trim()}`;
  const listQ = `
    query ($ids:[ID!]!, $cursor:String) {
      boards(ids:$ids) {
        id
        items_page(limit:200, cursor:$cursor) {
          cursor
          items { id name group { id } }
        }
      }
    }`;

  let cursor = null;
  let foundId = null;
  while (true) {
    const data = MI_gql_(listQ, { ids: [String(MON_CFG.BOARD_ID)], cursor });
    const b = data?.boards?.[0];
    if (!b) break;
    const page = b.items_page || {};
    const hit = (page.items || []).find(it => it?.group?.id === MON_CFG.GROUP_ID && String(it.name || '') === wanted);
    if (hit) { foundId = String(hit.id); break; }
    if (!page.cursor) break;
    cursor = page.cursor;
  }

  if (!foundId) {
    const createM = `
      mutation ($boardId: ID!, $groupId: String!, $name: String!) {
        create_item(board_id: $boardId, group_id: $groupId, item_name: $name) { id }
      }`;
    const created = MI_gql_(createM, {
      boardId: String(MON_CFG.BOARD_ID),
      groupId: String(MON_CFG.GROUP_ID),
      name:    String(wanted)
    });
    foundId = created?.create_item?.id;
    if (!foundId) throw new Error('Could not create parent item.');
    console.log(`Created parent: "${wanted}" (${foundId})`);
  }

  // Ensure main query fields are set (once per parent)
  MI_ensureParentMainQueryFields_(foundId);

  // Set parent link if provided
  if (linkUrl) MI_setParentLinkIfAny_(foundId, linkUrl, linkText);

  return String(foundId);
}

/* ====================================================
   Create sub-item & set all relations/columns at once
   ==================================================== */

function MI_createSubitemWithAllValues_(
  parentItemId,
  subitemName,
  contributorItemId,
  societyItemId,
  societyOwnerUserId,
  statusLabel,
  trackingLabels,
  reviewIsoDate /* 'YYYY-MM-DD' */,
  linkUrl /* optional */,
  linkText /* optional */
) {
  const colVals = {};

  if (contributorItemId) {
    colVals[MON_CFG.SUBITEM_COLS.contributorRelation] = {
      board_id: Number(MON_CFG.CONTRIBUTOR_LIST_BOARD_ID),
      item_ids: [Number(contributorItemId)]
    };
  }
  if (societyItemId) {
    colVals[MON_CFG.SUBITEM_COLS.societyRelation] = {
      board_id: Number(MON_CFG.SOCIETY_LIST_BOARD_ID),
      item_ids: [Number(societyItemId)]
    };
  }
  if (societyOwnerUserId) {
    colVals[MON_CFG.SUBITEM_COLS.subitemOwner] = {
      personsAndTeams: [{ id: Number(societyOwnerUserId), kind: 'person' }]
    };
  }
  if (statusLabel) {
    colVals[MON_CFG.SUBITEM_COLS.subitemStatus] = { label: String(statusLabel) };
  }
  if (trackingLabels && trackingLabels.length) {
    colVals[MON_CFG.SUBITEM_COLS.trackingDropdown] = { labels: trackingLabels.map(String) };
  }
  if (MON_CFG.SUBITEM_COLS.reviewDate && reviewIsoDate) {
    colVals[MON_CFG.SUBITEM_COLS.reviewDate] = { date: String(reviewIsoDate) };
  }
  if (MON_CFG.SUBITEM_LINK_COL_ID && linkUrl) {
    colVals[MON_CFG.SUBITEM_LINK_COL_ID] = { url: String(linkUrl), text: String(linkText || linkUrl) };
  }

  const m = `
    mutation ($parent: ID!, $name: String!, $vals: JSON!) {
      create_subitem (parent_item_id: $parent, item_name: $name, column_values: $vals) { id }
    }`;

  const res = MI_gql_(m, {
    parent: String(parentItemId),
    name:   String(subitemName),
    vals:   JSON.stringify(colVals)
  });
  const sid = res?.create_subitem?.id;
  if (!sid) throw new Error('Sub-item not created.');
  return String(sid);
}

/* ============== small date helper ============== */
function MI_isoDatePlusDays_(days) {
  const tz = Session.getScriptTimeZone() || 'UTC';
  const d = new Date();
  d.setDate(d.getDate() + (Number(days) || 0));
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

/* ====================================================
   Updates (comments) with retry
   ==================================================== */

function MI_postUpdate_(itemId, body) {
  if (!itemId || !body) return;
  const m = `
    mutation ($id: ID!, $body: String!) {
      create_update (item_id: $id, body: $body) { id }
    }`;
  MI_gql_(m, { id: String(itemId), body: String(body) });
}

function MI_postUpdateWithRetry_(itemId, body) {
  const tries = [0, 600, 1200];
  for (let i = 0; i < tries.length; i++) {
    try {
      if (tries[i] > 0) Utilities.sleep(tries[i]);
      MI_postUpdate_(itemId, body);
      return true;
    } catch (e) {
      if (i === tries.length - 1) {
        console.log('WARN: failed to post update on item ' + itemId + ': ' + e);
        return false;
      }
    }
  }
  return false;
}

/* ====================================================
   RULES & GATING (with DIAGNOSTICS)
   ==================================================== */

let __RULES_CACHE__ = null;

function MI_loadIssueRules_() {
  if (__RULES_CACHE__) return __RULES_CACHE__;
  const sh = SpreadsheetApp.getActive().getSheetByName(MON_CFG.RULES_SHEET);
  const map = new Map();
  if (!sh) { __RULES_CACHE__ = map; return map; }

  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) { __RULES_CACHE__ = map; return map; }

  const headers = vals[0].map(h => String(h || '').toLowerCase().trim());
  const c = (labelArr) => {
    for (const l of labelArr) {
      const idx = headers.indexOf(String(l).toLowerCase());
      if (idx !== -1) return idx;
    }
    return -1;
  };
  const iPriority = c(['priority (from society list in monday)', 'priority']);
  const iStatus   = c(['status (from society issues)', 'status']);
  const iAccount  = c(['account priority (from salesforce status)', 'account priority', 'client priority']);
  const iRaise    = c(['raise?', 'raise']);

  for (let r = 1; r < vals.length; r++) {
    const pr = MI_normLabel_(vals[r][iPriority] || 'Not Allocated');
    const st = MI_normLabel_(vals[r][iStatus] || '');
    const ac = MI_normLabel_(vals[r][iAccount] || '');
    const ra = String(vals[r][iRaise] || '').toLowerCase().trim();
    const key = `${pr}|${st}|${ac}`;
    const raise = (ra === 'yes' || ra === 'y' || ra === 'true' || ra === '1');
    map.set(key, raise);
  }
  __RULES_CACHE__ = map;
  return map;
}

function MI_shouldRaiseByRules_(socPriority, issueStatus, acctPriority) {
  const rules = MI_loadIssueRules_();
  const pr = MI_normLabel_(socPriority || 'Not Allocated');
  const st = MI_normLabel_(issueStatus || '');
  const ac = MI_normLabel_(acctPriority || '');
  const key = `${pr}|${st}|${ac}`;
  return { key, raise: !!rules.get(key), exists: rules.has(key) };
}

/** Prefer Contributor BOARD by contributorId; fallback to Salesforce sheet by UUID. */
function MI_getContributorAccountPriority_(contribItemId, uuid) {
  const fromBoard = MI_getContributorAccountPriorityFromBoard_(contribItemId);
  if (fromBoard) return { value: fromBoard, source: 'Contributor Board' };
  const fromSheet = MI_getAccountPriorityFromSheet_(uuid);
  if (fromSheet) return { value: fromSheet, source: 'Salesforce Sheet' };
  return { value: '', source: 'None' };
}

function MI_getContributorAccountPriorityFromBoard_(contribItemId) {
  if (!contribItemId) return '';
  const colsMeta = MI_getBoardColumns_(MON_CFG.CONTRIBUTOR_LIST_BOARD_ID);
  const metaByTitle = new Map(colsMeta.map(m => [m.title.toLowerCase().trim(), m]));
  const cvals = MI_getItemColumnValues_(contribItemId);
  const byId = new Map(cvals.map(cv => [String(cv.id), cv]));

  const titles = ['client priority', 'account priority', 'client tier'];
  for (const t of titles) {
    const meta = metaByTitle.get(t);
    if (!meta) continue;
    const cv = byId.get(meta.id);
    const txt = String(cv?.text || '').trim();
    if (txt) return MI_normLabel_(txt);
    try {
      const parsed = cv?.value ? JSON.parse(cv.value) : null;
      if (parsed) {
        if (meta.type === 'dropdown') {
          if (Array.isArray(parsed.labels) && parsed.labels[0]) return MI_normLabel_(parsed.labels[0]);
        } else if (meta.type === 'status' || meta.type === 'color') {
          if (parsed.label) return MI_normLabel_(parsed.label);
        } else if (meta.type === 'mirror') {
          if (parsed.display_value) return MI_normLabel_(parsed.display_value);
          if (cv?.text) return MI_normLabel_(cv.text);
        }
      }
    } catch(_) {}
  }
  return '';
}

// Normalise the four tiers to canonical labels or '' if unknown
function MI_cleanPriorityLabel_(s) {
  const t = String(s || '').trim().toLowerCase();
  if (!t) return '';
  if (t.startsWith('bronze'))    return 'Bronze';
  if (t.startsWith('silver'))    return 'Silver';
  if (t.startsWith('gold'))      return 'Gold';
  if (t.startsWith('platinum'))  return 'Platinum';
  return '';
}

/**
 * Read Account Priority for a performer UUID from the Salesforce sheet.
 * Expects headers: "UUID" and "Account Priority" (case-insensitive).
 * Falls back to common synonyms if needed.
 */
function MI_getAccountPriorityFromSheet_(uuid) {
  if (!uuid) return '';
  const sh = SpreadsheetApp.getActive().getSheetByName(MON_CFG.SALESFORCE_SHEET);
  if (!sh) return '';

  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return '';

  const headers = vals[0].map(h => String(h || '').trim());
  const lower   = headers.map(h => h.toLowerCase());

  // UUID column — exact or contains
  let iUUID = lower.indexOf('uuid');
  if (iUUID === -1) iUUID = lower.findIndex(h => h.includes('uuid'));
  if (iUUID === -1) return '';

  // Account Priority column — prefer exact "Account Priority"
  let iPri = lower.indexOf('account priority');
  // Soft fallbacks if someone renamed it in the past
  if (iPri === -1) iPri = lower.indexOf('acct priority');
  if (iPri === -1) iPri = lower.indexOf('account priority (from salesforce status)');
  if (iPri === -1) return '';

  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][iUUID] || '').trim() === String(uuid).trim()) {
      return MI_cleanPriorityLabel_(vals[r][iPri]);
    }
  }
  return '';
}


/** Read Priority + Relationship Status from the Society List board item. */
function MI_getSocietyMeta_(societyItemId) {
  const colsMeta = MI_getBoardColumns_(MON_CFG.SOCIETY_LIST_BOARD_ID);
  const metaByTitle = new Map(colsMeta.map(m => [m.title.toLowerCase().trim(), m]));
  const cvals = MI_getItemColumnValues_(societyItemId);
  const byId = new Map(cvals.map(cv => [String(cv.id), cv]));

  function readTextByTitles(titles) {
    for (const t of titles) {
      const meta = metaByTitle.get(t.toLowerCase());
      if (!meta) continue;
      const cv = byId.get(meta.id);
      const txt = String(cv?.text || '').trim();
      if (txt) return txt;
      let parsed;
      try { parsed = cv?.value ? JSON.parse(cv.value) : null; } catch(_) {}
      if (parsed) {
        if (meta.type === 'dropdown') {
          if (Array.isArray(parsed.labels) && parsed.labels.length) return String(parsed.labels[0]);
        } else if (meta.type === 'status' || meta.type === 'color') {
          if (parsed.label) return String(parsed.label);
        } else if (meta.type === 'mirror') {
          if (parsed.display_value) return String(parsed.display_value);
        }
      }
    }
    return '';
  }

  const priority = readTextByTitles(['priority']);
  const relationship = readTextByTitles(['relationship status', 'relationship']);
  return { priority: priority || 'Not Allocated', relationship: relationship || '' };
}

/* ====================================================
   Comment building helpers
   ==================================================== */

function MI_pickIdx_(headers, candidates) {
  const lower = headers.map(h => String(h || '').trim().toLowerCase());
  for (const c of candidates) {
    const idx = lower.indexOf(c.toLowerCase());
    if (idx !== -1) return idx;
  }
  return -1;
}

function MI_fmtRecordingLabel_(title, version, artist) {
  const t = String(title || '').trim();
  const v = String(version || '').trim();
  const a = String(artist || '').trim();
  const tv = v ? `${t} (${v})` : t;
  return a ? `${tv} - ${a}` : tv;
}

function MI_getSocietyRecordingRows_(sheetName, uuid, society) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0];

  const iUUID    = MI_pickIdx_(headers, ['uuid']);
  const iSoc     = MI_pickIdx_(headers, ['society']);
  const iTitle   = MI_pickIdx_(headers, ['title','recording title','track title']);
  const iMainArt = MI_pickIdx_(headers, ['main artist','main_artist','primary artist']);
  const iArtist  = MI_pickIdx_(headers, ['artist','artist name']);
  const iVersion = MI_pickIdx_(headers, ['version','track version','mix']);
  const iPlays   = MI_pickIdx_(headers, ['plays (qualified)','plays','streams','qualified plays']);
  const iStatus  = MI_pickIdx_(headers, ['status','recording status']);

  const out = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    if (iUUID !== -1 && String(row[iUUID] || '').trim() !== String(uuid).trim()) continue;
    if (iSoc  !== -1 && String(row[iSoc]  || '').trim().toLowerCase() !== String(society || '').trim().toLowerCase()) continue;

    const title   = iTitle   !== -1 ? row[iTitle]   : '';
    const version = iVersion !== -1 ? row[iVersion] : '';
    const artist  = iMainArt !== -1 ? row[iMainArt] : (iArtist !== -1 ? row[iArtist] : '');

    out.push({
      title:  String(title || '').trim(),
      version:String(version || '').trim(),
      artist: String(artist || '').trim(),
      plays:  iPlays !== -1 ? Number(row[iPlays] || 0) : null,
      status: iStatus !== -1 ? String(row[iStatus] || '').trim() : ''
    });
  }
  return out;
}

function MI_getMostPlayedForUUID_(uuid) {
  const sh = SpreadsheetApp.getActive().getSheetByName(MON_CFG.MOST_PLAYED_SHEET);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0];

  const iUUID    = MI_pickIdx_(headers, ['uuid']);
  const iTitle   = MI_pickIdx_(headers, ['title','recording title','track title']);
  const iMainArt = MI_pickIdx_(headers, ['main artist','main_artist','primary artist']);
  const iArtist  = MI_pickIdx_(headers, ['artist','artist name']);
  const iVersion = MI_pickIdx_(headers, ['version','track version','mix']);
  const iPlays   = MI_pickIdx_(headers, ['plays (qualified)','plays','streams','qualified plays']);

  const out = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    if (iUUID !== -1 && String(row[iUUID] || '').trim() !== String(uuid).trim()) continue;

    const title   = iTitle   !== -1 ? row[iTitle]   : '';
    const version = iVersion !== -1 ? row[iVersion] : '';
    const artist  = iMainArt !== -1 ? row[iMainArt] : (iArtist !== -1 ? row[iArtist] : '');

    out.push({
      title:  String(title || '').trim(),
      version:String(version || '').trim(),
      artist: String(artist || '').trim(),
      plays:  iPlays !== -1 ? Number(row[iPlays] || 0) : null
    });
  }
  out.sort((a,b) => (Number(b.plays||0) - Number(a.plays||0)));
  return out.slice(0, MON_CFG.COMMENT.MAX_RECORDING_BULLETS);
}

function MI_fmtPlays_(n) {
  if (n == null || isNaN(n)) return '';
  try { return Utilities.formatString('%s', Math.round(Number(n))); } catch (e) { return String(n); }
}

function MI_buildMissingComment_(p) {
  const { contributorName, societyName, period, recordings, fallbackMostPlayed, performerFileUrl, reviewDate } = p;
  const lines = [];
  lines.push(`Missing Contributor — ${contributorName}`);
  lines.push('');
  lines.push('Issues Identified');
  lines.push(`• ${contributorName} appears to be missing from ${societyName}’s payee/contributor list for ${period || MON_CFG.COMMENT.DEFAULT_PERIOD}.`);
  lines.push('• No payments or matches were found for this contributor at this society.');
  lines.push('');
  lines.push('Recordings To Check');
  if (recordings && recordings.length) {
    recordings.forEach(r => {
      const label = MI_fmtRecordingLabel_(r.title, r.version, r.artist);
      lines.push(`• ${label} : ${MI_fmtPlays_(r.plays)} plays`);
    });
  } else {
    lines.push('No track-level rows at this society. Using Most Played Recordings for the performer:');
    (fallbackMostPlayed || []).forEach(r => {
      const label = MI_fmtRecordingLabel_(r.title, r.version, r.artist);
      lines.push(`• ${label} : ${MI_fmtPlays_(r.plays)} plays`);
    });
  }
  lines.push('');
  lines.push('Actions Required');
  lines.push(`• Confirm if ${contributorName} is registered/linked correctly.`);
  lines.push('');
  lines.push('References');
  if (performerFileUrl) lines.push(`• Performer Overview (GSheet): ${performerFileUrl}`);
  if (reviewDate)       lines.push(`• Review date: ${reviewDate}`);
  lines.push('');
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd HH:mm');
  lines.push(`Note: This comment was auto-generated by the Income Tracking automation on ${ts}.`);
  return lines.join('\n');
}

function MI_buildLowIncomeComment_(p) {
  const { contributorName, societyName, period, recordings, fallbackMostPlayed, performerFileUrl, reviewDate, yoyDelta, QoQDelta, trendFlags } = p;
  const lines = [];
  lines.push(`Low Income — ${contributorName}`);
  lines.push('');
  lines.push('Issues Identified');
  const yoyStr = yoyDelta ? `${yoyDelta} YoY` : '';
  const qoqStr = QoQDelta ? `${QoQDelta} QoQ` : '';
  const joinDeltas = [yoyStr, qoqStr].filter(Boolean).join(', ');
  lines.push(`• Income from ${societyName} for ${contributorName} is below expected in ${period || MON_CFG.COMMENT.DEFAULT_PERIOD}${joinDeltas ? ` (${joinDeltas})` : ''}.`);
  if (trendFlags) lines.push(`• Trend flags: ${trendFlags}.`);
  lines.push('');
  lines.push('Recordings To Review');
  if (recordings && recordings.length) {
    recordings.forEach(r => {
      const label  = MI_fmtRecordingLabel_(r.title, r.version, r.artist);
      const status = r.status ? ` — Status: ${r.status}` : '';
      lines.push(`• ${label} : ${MI_fmtPlays_(r.plays)} plays${status}`);
    });
  } else {
    lines.push('No track-level rows at this society. Using Most Played Recordings for the performer:');
    (fallbackMostPlayed || []).forEach(r => {
      const label = MI_fmtRecordingLabel_(r.title, r.version, r.artist);
      lines.push(`• ${label} : ${MI_fmtPlays_(r.plays)} plays`);
    });
  }
  lines.push('');
  lines.push('Actions Required');
  lines.push(`• Check if ${contributorName} is registered correctly.`);
  lines.push('• Check the above recordings have been linked to their account.');
  lines.push('');
  lines.push('References');
  if (performerFileUrl) lines.push(`• Performer Overview (GSheet): ${performerFileUrl}`);
  if (reviewDate)       lines.push(`• Review date: ${reviewDate}`);
  lines.push('');
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'UTC', 'yyyy-MM-dd HH:mm');
  lines.push(`Note: This comment was auto-generated by the Income Tracking automation on ${ts}.`);
  return lines.join('\n');
}

function MI_buildCommentForRow_(rowObj, type, options) {
  const { uuid, contributor, society, territory } = rowObj;
  const maxN = MON_CFG.COMMENT.MAX_RECORDING_BULLETS;

  let recs = MI_getSocietyRecordingRows_(MON_CFG.RECORDING_ISSUES_SHEET, uuid, society);
  if (recs.length) {
    recs.sort((a,b)=> (Number(b.plays||0) - Number(a.plays||0)));
    recs = recs.slice(0, maxN);
  }
  const fallback = recs.length ? [] : MI_getMostPlayedForUUID_(uuid);

  const base = {
    contributorName: String(contributor || '').trim(),
    societyName:     String(society || '').trim(),
    period:          options?.period,
    recordings:      recs,
    fallbackMostPlayed: fallback,
    performerFileUrl: options?.performerFileUrl || '',
    reviewDate:      options?.reviewDate || '',
    yoyDelta:        options?.yoyDelta || '',
    QoQDelta:        options?.QoQDelta || '',
    trendFlags:      options?.trendFlags || ''
  };

  return (type === 'missing')
    ? MI_buildMissingComment_(base)
    : MI_buildLowIncomeComment_(base);
}

/* ===========================================
   PUBLIC ENTRIES (simple “first” + “batch”)
   =========================================== */

function MONDAY_pushFirstSocietyIssueAsSubitem() {
  const rows = MI_readSocietyIssuesRows_(1);
  if (!rows.length) throw new Error('No data rows found in Society Issues.');
  const r = rows[0];

  const contribMap = MI_loadNameIdMapFromBoard_(MON_CFG.CONTRIBUTOR_LIST_BOARD_ID);
  const societyMap = MI_loadNameIdMapFromBoard_(MON_CFG.SOCIETY_LIST_BOARD_ID);

  const contributorName = String(r.contributor || '').trim();
  const societyName     = String(r.society || '').trim();
  const territory       = String(r.territory || '').trim();

  const contribId =
    contribMap.get(MI_norm_(contributorName)) ||
    contribMap.get(MI_norm_(MI_stripParen_(contributorName))) || null;

  const societyId =
    societyMap.get(MI_norm_(societyName)) ||
    societyMap.get(MI_norm_(MI_stripParen_(societyName))) || null;

  if (!contribId) { console.log(`WARNING: Could not resolve contributor "${contributorName}"`); return; }
  if (!societyId) { console.log(`WARNING: Could not resolve society "${societyName}"`); return; }

  // === RULES & RELATIONSHIP GATING ===
  const acctObj = MI_getContributorAccountPriority_(String(contribId), r.uuid);
  const acctPriority = acctObj.value;
  const { priority: societyPriority, relationship } = MI_getSocietyMeta_(String(societyId));
  const relOk = MON_CFG.ALLOWED_RELATIONSHIP_STATUSES
    .map(x => x.toLowerCase()).includes(String(relationship || '').toLowerCase());
  if (!relOk) { console.log(`Monday: SKIP relationship "${relationship || 'Unknown'}" for ${societyName}.`); return; }

  const issueStatus = String(r.status || '').trim();
  const rule = MI_shouldRaiseByRules_(societyPriority || 'Not Allocated', issueStatus, acctPriority);
  if (!rule.raise) {
    console.log(`Monday: SKIP by rules: ${societyPriority} | ${MI_normLabel_(issueStatus)} | ${acctPriority}`);
    return;
  }

  const statusLabel = 'For Society';
  const trackingLabels =
    (String(r.status || '').toLowerCase().startsWith('missing')) ? ['Missing Contributor']
    : (String(r.status || '').toLowerCase().startsWith('under') ? ['Low Income'] : []);

  const reviewISO = MI_isoDatePlusDays_(14);
  const subTitle = `${societyName} – ${contributorName}${territory ? ` (${territory})` : ''}`;

  const societyFileUrl   = ''; // optional: set if you have it
  const performerFileUrl = ''; // optional: set if you have it

  const parentId = MI_findOrCreateParentItem_(societyName, societyFileUrl, societyFileUrl ? 'Open Society Analysis' : undefined);
  const ownerUserId = MI_getSocietyOwnerUserId_(String(societyId)) || null;

  const subId = MI_createSubitemWithAllValues_(
    parentId, subTitle, contribId, societyId, ownerUserId, statusLabel, trackingLabels, reviewISO,
    performerFileUrl, performerFileUrl ? 'Open Performer Analysis' : undefined
  );

  Utilities.sleep(600);
  const type = (trackingLabels[0] === 'Missing Contributor') ? 'missing' : 'low';
  const comment = MI_buildCommentForRow_(r, type, { reviewDate: reviewISO, performerFileUrl });
  MI_postUpdateWithRetry_(subId, comment);
}

function MONDAY_pushSocietyIssuesBatch() {
  const rows = MI_readSocietyIssuesRows_(MON_CFG.TEST_BATCH_ROWS);
  if (!rows.length) throw new Error('No data rows found in Society Issues.');

  const contribMap = MI_loadNameIdMapFromBoard_(MON_CFG.CONTRIBUTOR_LIST_BOARD_ID);
  const societyMap = MI_loadNameIdMapFromBoard_(MON_CFG.SOCIETY_LIST_BOARD_ID);

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    try {
      const contributorName = String(r.contributor || '').trim();
      const societyName     = String(r.society || '').trim();
      const territory       = String(r.territory || '').trim();

      const contribId =
        contribMap.get(MI_norm_(contributorName)) ||
        contribMap.get(MI_norm_(MI_stripParen_(contributorName))) || null;

      const societyId =
        societyMap.get(MI_norm_(societyName)) ||
        societyMap.get(MI_norm_(MI_stripParen_(societyName))) || null;

      if (!contribId) { console.log(`Row ${r.rowIndex}: WARN no contributor match for "${contributorName}"`); continue; }
      if (!societyId) { console.log(`Row ${r.rowIndex}: WARN no society match for "${societyName}"`); continue; }

      // === RULES & RELATIONSHIP GATING ===
      const acctObj = MI_getContributorAccountPriority_(String(contribId), r.uuid);
      const acctPriority = acctObj.value;
      const { priority: societyPriority, relationship } = MI_getSocietyMeta_(String(societyId));
      const relOk = MON_CFG.ALLOWED_RELATIONSHIP_STATUSES
        .map(x => x.toLowerCase()).includes(String(relationship || '').toLowerCase());
      if (!relOk) { console.log(`Monday: SKIP relationship "${relationship || 'Unknown'}" for ${societyName}.`); continue; }

      const issueStatus = String(r.status || '').trim();
      const rule = MI_shouldRaiseByRules_(societyPriority || 'Not Allocated', issueStatus, acctPriority);
      if (!rule.raise) {
        console.log(`Monday: SKIP by rules: ${societyPriority} | ${MI_normLabel_(issueStatus)} | ${acctPriority}`);
        continue;
      }

      const statusLabel = 'For Society';
      const trackingLabels =
        (String(r.status || '').toLowerCase().startsWith('missing')) ? ['Missing Contributor']
          : (String(r.status || '').toLowerCase().startsWith('under') ? ['Low Income'] : []);

      const reviewISO = MI_isoDatePlusDays_(14);
      const subTitle = `${societyName} – ${contributorName}${territory ? ` (${territory})` : ''}`;

      const societyFileUrl   = ''; // optional
      const performerFileUrl = ''; // optional

      const parentId = MI_findOrCreateParentItem_(societyName, societyFileUrl, societyFileUrl ? 'Open Society Analysis' : undefined);
      const ownerUserId = MI_getSocietyOwnerUserId_(String(societyId)) || null;

      const subId = MI_createSubitemWithAllValues_(
        parentId, subTitle, contribId, societyId, ownerUserId, statusLabel, trackingLabels, reviewISO,
        performerFileUrl, performerFileUrl ? 'Open Performer Analysis' : undefined
      );

      Utilities.sleep(600);
      const type = (trackingLabels[0] === 'Missing Contributor') ? 'missing' : 'low';
      const comment = MI_buildCommentForRow_(r, type, { reviewDate: reviewISO, performerFileUrl });
      MI_postUpdateWithRetry_(subId, comment);

      console.log(`Row ${r.rowIndex}: created subitem ${subId} → "${subTitle}"`);
      Utilities.sleep(200);

    } catch (e) {
      console.log(`Row ${r.rowIndex}: ERROR ${e}`);
    }
  }
}

/* ============== helpers ============== */

function MI_loadNameIdMapFromBoard_(boardId) {
  const id = String(boardId).trim();
  const query = `
    query ($ids: [ID!]!, $cursor: String) {
      boards(ids: $ids) {
        id
        items_page (limit: 200, cursor: $cursor) {
          cursor
          items { id name group { id } }
        }
      }
    }`;

  const map = new Map();
  let cursor = null;
  let total = 0;

  while (true) {
    const data = MI_gql_(query, { ids: [id], cursor });
    const b = data?.boards?.[0];
    if (!b) break;
    const page = b.items_page || {};
    const items = page.items || [];
    items.forEach(it => {
      const nm = String(it.name || '').trim();
      const itemId = String(it.id || '').trim();
      if (!nm || !itemId) return;
      map.set(MI_norm_(nm), itemId);
      const alt = MI_norm_(MI_stripParen_(nm));
      if (alt && !map.has(alt)) map.set(alt, itemId);
    });
    total += items.length;
    if (!page.cursor) break;
    cursor = page.cursor;
  }

  console.log(`Loaded ${total} items from board ${boardId}.`);
  return map;
}

function socityIdSafe_(id) { return String(id).trim(); }

/* =========================================================
   DIAGNOSTICS
   ========================================================= */

/** Dump important lookup columns for quick sanity-check (to logs). */
function MONDAY_dumpLookupColumns() {
  const dump = (boardId, title) => {
    const cols = MI_getBoardColumns_(boardId);
    console.log(`\n${title} — ${cols.length} columns (id | type | title)`);
    cols.forEach(c => console.log(`${c.id} | ${c.type} | ${c.title}`));
  };
  dump(MON_CFG.CONTRIBUTOR_LIST_BOARD_ID, `Contributor List board ${MON_CFG.CONTRIBUTOR_LIST_BOARD_ID}`);
  dump(MON_CFG.SOCIETY_LIST_BOARD_ID,     `Society List board ${MON_CFG.SOCIETY_LIST_BOARD_ID}`);
}

/** Write the normalized rules map into a sheet (“Diag – Rules Map”). */
function MONDAY_writeRulesMapSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MON_CFG.DIAG_RULES_SHEET);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(MON_CFG.DIAG_RULES_SHEET);

  const rules = MI_loadIssueRules_();
  const rows = [['Society Priority','Status','Account Priority','Raise?','Composite Key']];
  for (const [key, raise] of rules.entries()) {
    const [pr, st, ac] = key.split('|');
    rows.push([pr, st, ac, raise ? 'Yes' : 'No', key]);
  }
  sh.getRange(1,1,rows.length, rows[0].length).setValues(rows);
  try { sh.setFrozenRows(1); sh.autoResizeColumns(1, rows[0].length); } catch(e){}
}

/** Build a detailed preview (“Diag – Preview”) showing exactly what’s compared and the final decision. */
function MONDAY_runDiagnostics() {
  // Load a larger set of Society Issues rows
  const shSI = SpreadsheetApp.getActive().getSheetByName(MON_CFG.SOURCE_SHEET);
  if (!shSI) throw new Error(`Missing sheet: ${MON_CFG.SOURCE_SHEET}`);
  const vals = shSI.getDataRange().getValues();
  if (vals.length < 2) throw new Error('No data rows in Society Issues');

  // Build index helpers
  const headers = vals[0].map(h => String(h || '').toLowerCase().trim());
  const col = (label) => headers.indexOf(String(label).toLowerCase());

  // Build lookup maps for board ids
  const contribMap = MI_loadNameIdMapFromBoard_(MON_CFG.CONTRIBUTOR_LIST_BOARD_ID);
  const societyMap = MI_loadNameIdMapFromBoard_(MON_CFG.SOCIETY_LIST_BOARD_ID);

  // Prepare diag sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(MON_CFG.DIAG_PREVIEW_SHEET);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(MON_CFG.DIAG_PREVIEW_SHEET);

  const out = [[
    'Row','Contributor','UUID','Contributor Item ID',
    'Acct Priority (Board)','Acct Priority (SF Sheet)','Acct Priority (Chosen, Source)',
    'Society','Society Item ID','Society Priority','Relationship',
    'Issue Status (Raw)','Issue Status (Norm)',
    'Relationship Allowed','Rule Key','Rule Exists','Rule Says Raise','FINAL: Raise?','Skip Reason'
  ]];

  const limit = Math.min(MON_CFG.DIAG_PREVIEW_LIMIT, vals.length - 1);
  for (let r = 1; r <= limit; r++) {
    const row = vals[r];
    const contributor = row[col('contributor')];
    const uuid        = row[col('uuid')];
    const society     = row[col('society')];
    const statusRaw   = row[col('status')];

    if (!String(contributor || society || '').trim()) continue;

    // Resolve board item ids by NAME (as production does)
    const contribId = contribMap.get(MI_norm_(String(contributor || ''))) ||
                      contribMap.get(MI_norm_(MI_stripParen_(String(contributor || '')))) || '';
    const societyId = societyMap.get(MI_norm_(String(society || ''))) ||
                      societyMap.get(MI_norm_(MI_stripParen_(String(society || '')))) || '';

    // Account priority (Contributor ID first, then SF sheet by UUID)
    const apBoard = MI_getContributorAccountPriorityFromBoard_(String(contribId || '')) || '';
    const apSheet = MI_getAccountPriorityFromSheet_(String(uuid || '')) || '';
    const apChosenObj = MI_getContributorAccountPriority_(String(contribId || ''), String(uuid || ''));
    const apChosen = apChosenObj.value;
    const apSource = apChosenObj.source;

    // Society meta by society item id
    let socPriority = '';
    let relationship = '';
    if (societyId) {
      const meta = MI_getSocietyMeta_(String(societyId));
      socPriority = MI_normLabel_(meta.priority || 'Not Allocated');
      relationship = MI_normLabel_(meta.relationship || '');
    }

    // Relationship gating
    const relAllowed = MON_CFG.ALLOWED_RELATIONSHIP_STATUSES
      .map(x => x.toLowerCase()).includes(String(relationship || '').toLowerCase());

    // Status normalization + rule
    const issueNorm = MI_normLabel_(String(statusRaw || ''));
    const rule = MI_shouldRaiseByRules_(socPriority || 'Not Allocated', issueNorm, apChosen);
    const finalRaise = !!(relAllowed && rule.raise);

    // Skip reason (if not raised)
    let skip = '';
    if (!relAllowed) skip = `Relationship "${relationship}" not allowed`;
    else if (!rule.exists) skip = `No rule for key`;
    else if (!rule.raise)  skip = `Rule says No`;

    out.push([
      r,
      String(contributor || ''), String(uuid || ''), String(contribId || ''),
      apBoard, apSheet, apChosen ? `${apChosen} (${apSource})` : '',
      String(society || ''), String(societyId || ''), socPriority, relationship,
      String(statusRaw || ''), issueNorm,
      relAllowed ? 'Yes' : 'No', rule.key, rule.exists ? 'Yes' : 'No', rule.raise ? 'Yes' : 'No',
      finalRaise ? 'Yes' : 'No', skip
    ]);
  }

  sh.getRange(1,1,out.length,out[0].length).setValues(out);
  try { sh.setFrozenRows(1); sh.autoResizeColumns(1, out[0].length); } catch(e){}

  // Also write the normalized rules for side-by-side review
  MONDAY_writeRulesMapSheet_();
}
