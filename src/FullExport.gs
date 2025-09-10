/** =========================================================
 * Full Export (Chunked, Resumable)
 * - Uses ONLY the overview sheets as sources (no per-performer tabs built)
 * - Copies formatting by duplicating overview sheets, then writes filtered rows
 * - Processes performers/societies in batches with time-based triggers
 * - Optional Monday phase AFTER society exports (chunked)
 * - Cleans up overview tabs and the Export Log at the very end
 * ========================================================= */

const FULL_EXPORT = {
  BASE_FOLDER_ID: '', // e.g. '0AOxxx...'; optional if you use folder-name search
  BASE_FOLDER_NAME: 'Income Tracking',

  // Folder nodes
  YEAR_NODE_NAME: (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy'),
  REVIEW_NODE: 'Performer Reviews',
  DATE_NODE_NAME: (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
  TIME_NODE_NAME: (d) => Utilities.formatDate(d, Session.getScriptTimeZone(), 'HHmm'),

  // Subfolders
  PERFORMER_BREAKDOWN: 'Performer Breakdown',
  SOCIETY_BREAKDOWN: 'Society Breakdown',

  // Overview workbook name
  OVERVIEW_FILE_NAME: (d) => `Overview - ` +
    Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH.mm'),

  // Source sheet names
  SHEET_NAMES: {
    CONTRIBUTORS: 'Contributor List',
    PERF_OVERVIEW: 'Performer Overview',
    SOC_ISSUES: 'Society Issues',
    NON_HOME: 'Non-Home Territory Income',
    REC_OVERVIEW: 'Recording Overview',
    REC_ISSUES: 'Recording Issues',
    REC_MOST: 'Most Played Recordings'
  },

  // Batch sizes
  BATCH: {
    PERFORMERS_PER_RUN: 20,
    SOCIETIES_PER_RUN: 15,
    MONDAY_PERFORMERS_PER_RUN: 10 // chunk Monday push to avoid timeouts
  },

  // Trigger spacing (ms)
  TRIGGER_DELAY_MS: 45 * 1000
};

/* =========================================================
   PUBLIC ENTRYPOINTS
   ========================================================= */

/** Start (or resume) a full export. Use this in your menu. */
function EXPORT_startFullAnalysis() {
  const ctx = CORE_buildContext_();

  // confirmation
  const nContribs = READ_readContributors_(ctx).size;
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    `You are about to run a full export for ${nContribs} contributors. Do you wish to proceed?`,
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  // Ask for Monday.com raising
  let mondayEnabled = false;
  try {
    const mresp = ui.alert(
      'Would you like to raise any resulting issues in Monday.com?',
      ui.ButtonSet.YES_NO
    );
    mondayEnabled = (mresp === ui.Button.YES);
  } catch (e) { /* headless */ }

  // Prepare/refresh overview sheets (fast)
  try {
    USECASE_buildOverviewOnly();
    SOC_buildSocietyIssuesOverview();
    SOC_buildNonHomeTerritoryIncome();
    REC_buildRecordingOverview();
    REC_buildRecordingIssues();
    if (typeof REC_buildMostPlayedRecordings === 'function') {
      REC_buildMostPlayedRecordings();
    }
  } catch (e) {
    console.warn('Pre-export refresh (overviews) failed:', e);
  }

  // Initialize state
  const startedAt = new Date();
  const exportFolder = FULL_buildFolderTree_(startedAt);
  const exportFolderId = exportFolder.getId();

  // Build lists
  const contributors = READ_readContributors_(ctx);
  const performerIds = Array.from(contributors.keys());

  const statuses = READ_readSalesforceStatuses_(ctx);
  const societySet = new Set();
  statuses.forEach(arr => (arr || []).forEach(s => {
    const soc = CORE_canonSociety_(s.societyName || '');
    if (soc) societySet.add(soc);
  }));
  const societies = Array.from(societySet).sort((a,b)=>a.localeCompare(b));

  // Persist state
  EXPORT_stateSave_({
    startedAt: startedAt.toISOString(),
    folderId: exportFolderId,
    performerFolderId: FULL_getOrCreateChild_(exportFolder, FULL_EXPORT.PERFORMER_BREAKDOWN).getId(),
    societyFolderId: FULL_getOrCreateChild_(exportFolder, FULL_EXPORT.SOCIETY_BREAKDOWN).getId(),
    overviewDone: false,

    // performer stage
    performerIds,
    performerIdx: 0,

    // society stage
    societies,
    societyIdx: 0,

    // file URL indices we’ll build during export
    perfFileUrlByUUID: {},        // uuid -> performer GSheet URL
    societyFileUrlByName: {},     // society -> society GSheet URL

    // monday stage (AFTER societies)
    mondayEnabled: !!mondayEnabled,
    mondayIdx: 0,
    mondayDone: !mondayEnabled // if disabled, mark done up-front
  });

  // Write overview workbook now (once)
  FULL_writeOverviewWorkbookByState_();

  // Kick first batch
  EXPORT_continueFullAnalysis_();
}

/** Abort and clean triggers/state (safe to run anytime). */
function EXPORT_abortFullAnalysis() {
  EXPORT_clearTriggers_();
  EXPORT_stateClear_();
  SpreadsheetApp.getActive().toast('Full export aborted and state cleared.');
}

/* =========================================================
   SCHEDULER / STATE
   ========================================================= */

const EXPORT_STATE_KEY = 'FULL_EXPORT_STATE';

function EXPORT_stateSave_(obj) {
  const sp = PropertiesService.getScriptProperties();
  sp.setProperty(EXPORT_STATE_KEY, JSON.stringify(obj || {}));
}
function EXPORT_stateLoad_() {
  const sp = PropertiesService.getScriptProperties();
  const raw = sp.getProperty(EXPORT_STATE_KEY);
  return raw ? JSON.parse(raw) : null;
}
function EXPORT_stateClear_() {
  PropertiesService.getScriptProperties().deleteProperty(EXPORT_STATE_KEY);
}

function EXPORT_scheduleNext_() {
  ScriptApp.newTrigger('EXPORT_continueFullAnalysis_')
    .timeBased()
    .after(FULL_EXPORT.TRIGGER_DELAY_MS)
    .create();
}
function EXPORT_clearTriggers_() {
  const all = ScriptApp.getProjectTriggers() || [];
  all.forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'EXPORT_continueFullAnalysis_') {
      try { ScriptApp.deleteTrigger(t); } catch(e) {}
    }
  });
}

/** Continue processing: performers → societies → monday → cleanup → done. */
function EXPORT_continueFullAnalysis_() {
  // Avoid multiple concurrent triggers
  EXPORT_clearTriggers_();

  // Cache reset per run chunk (for fast filtered copy)
  if (__FULL_MATRIX_CACHE__ && __FULL_MATRIX_CACHE__.clear) __FULL_MATRIX_CACHE__.clear();

  let st = EXPORT_stateLoad_();
  if (!st) {
    SpreadsheetApp.getActive().toast('No export in progress.');
    return;
  }

  // 1) Ensure overview workbook exists
  if (!st.overviewDone) {
    FULL_writeOverviewWorkbookByState_(); // idempotent
    st = EXPORT_stateLoad_();
    if (!st.overviewDone) { EXPORT_scheduleNext_(); return; }
  }

  // 2) Performer batch
  const pNext = processPerformerBatch_(st);
  if (pNext) { EXPORT_scheduleNext_(); return; }

  // 3) Society batch
  const sNext = processSocietyBatch_(st);
  if (sNext) { EXPORT_scheduleNext_(); return; }

  // 4) Monday batch (AFTER societies)
  if (!st.mondayDone) {
    const mNext = processMondayBatch_(st);
    if (mNext) { EXPORT_scheduleNext_(); return; }
    // mark done in state
    st = EXPORT_stateLoad_();
  }

  // 5) Done
  const url = DriveApp.getFolderById(st.folderId).getUrl();
  EXPORT_logProgress_(`Full export completed: ${url}`, /*final=*/true);
  EXPORT_stateClear_();

  try { SpreadsheetApp.getUi().alert(`Full export completed.\n\nFolder:\n${url}`); } catch(e) {}

  // 6) Cleanup temporary sheets (overviews + Export Log in main file)
  EXPORT_cleanupTemporarySheets_();
}

/** Logs progress to (or creates) an "Export Log" sheet. */
function EXPORT_logProgress_(msg, final) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName('Export Log');
  if (!log) log = ss.insertSheet('Export Log');
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  log.appendRow([ts, msg]);
  if (final) {
    try { log.getRange(log.getLastRow(), 1, 1, 2).setFontWeight('bold'); } catch(e) {}
  }
}

/* =========================================================
   FOLDER / SPREADSHEET HELPERS
   ========================================================= */

const __FULL_MATRIX_CACHE__ = new Map();

function FULL_findBaseFolder_() {
  if (FULL_EXPORT.BASE_FOLDER_ID) {
    try { return DriveApp.getFolderById(FULL_EXPORT.BASE_FOLDER_ID); } catch (e) {}
  }
  const it = DriveApp.searchFolders(
    `title = '${FULL_EXPORT.BASE_FOLDER_NAME.replace(/'/g, "\\'")}' and trashed = false`
  );
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(FULL_EXPORT.BASE_FOLDER_NAME);
}
function FULL_getOrCreateChild_(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}
function FULL_buildFolderTree_(ts) {
  const base = FULL_findBaseFolder_();
  const yearFolder   = FULL_getOrCreateChild_(base, FULL_EXPORT.YEAR_NODE_NAME(ts));
  const reviewFolder = FULL_getOrCreateChild_(yearFolder, FULL_EXPORT.REVIEW_NODE);
  const dateFolder   = FULL_getOrCreateChild_(reviewFolder, FULL_EXPORT.DATE_NODE_NAME(ts));
  const timeFolder   = FULL_getOrCreateChild_(dateFolder, FULL_EXPORT.TIME_NODE_NAME(ts));
  FULL_getOrCreateChild_(timeFolder, FULL_EXPORT.PERFORMER_BREAKDOWN);
  FULL_getOrCreateChild_(timeFolder, FULL_EXPORT.SOCIETY_BREAKDOWN);
  return timeFolder;
}

function FULL_newSpreadsheetInFolder_(name, folderId) {
  const ss = SpreadsheetApp.create(name);
  try {
    const file = DriveApp.getFileById(ss.getId());
    const folder = DriveApp.getFolderById(folderId);
    folder.addFile(file);
  } catch(e) {}
  return ss; // keep default sheet; we delete once real sheet exists
}
function FULL_removeDefaultIfSafe_(ss) {
  const sheets = ss.getSheets();
  if (sheets.length <= 1) return;
  const def = sheets.find(sh => sh.getName() === 'Sheet1');
  if (def) ss.deleteSheet(def);
}
function FULL_copyWholeSheetIfNonEmpty_(srcSheet, dstSs, dstName) {
  if (!srcSheet) return null;
  if (srcSheet.getLastRow() < 2) return null;
  const copy = srcSheet.copyTo(dstSs);
  copy.setName(dstName);
  FULL_removeDefaultIfSafe_(dstSs);
  return copy;
}

/* =======================
   Fast matrices + filtering
   ======================= */

function FULL_getSheetMatrices_(srcSheet) {
  const key = srcSheet.getSheetId();
  if (__FULL_MATRIX_CACHE__.has(key)) return __FULL_MATRIX_CACHE__.get(key);

  const lastRow = srcSheet.getLastRow();
  const lastCol = srcSheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    const empty = { headers: [], values: [], formats: [], lastRow: 0, lastCol: 0, hmap: new Map() };
    __FULL_MATRIX_CACHE__.set(key, empty);
    return empty;
  }
  const range = srcSheet.getRange(1, 1, lastRow, lastCol);
  const values  = range.getValues();
  const formats = range.getNumberFormats();
  const headers = (values[0] || []).map(v => String(v || ''));
  const hmap    = CORE_headerIndexMap_(headers);
  const out = { headers, values, formats, lastRow, lastCol, hmap };
  __FULL_MATRIX_CACHE__.set(key, out);
  return out;
}

function FULL_copyFilteredWithFormatting_(srcSheet, dstSs, dstName, predicate) {
  if (!srcSheet) return null;
  const { headers, values, formats, lastRow, lastCol, hmap } = FULL_getSheetMatrices_(srcSheet);
  if (!headers.length || lastRow < 2) return null;

  const filteredVals = [];
  const filteredFmts = [];
  for (let r = 1; r < lastRow; r++) {
    const rowVals = values[r];
    const keep = !!predicate({
      headers,
      get: (name) => {
        const i = CORE_idx_(hmap, name);
        return i >= 0 ? rowVals[i] : '';
      },
      row: rowVals
    });
    if (keep) {
      filteredVals.push(rowVals.slice(0, lastCol));
      filteredFmts.push(formats[r].slice(0, lastCol));
    }
  }
  if (!filteredVals.length) return null;

  const copy = srcSheet.copyTo(dstSs);
  copy.setName(dstName);
  FULL_removeDefaultIfSafe_(dstSs);

  const maxRows = copy.getMaxRows(), maxCols = copy.getMaxColumns();
  if (maxRows > 1) copy.getRange(2, 1, maxRows - 1, maxCols).clearContent();

  copy.getRange(2, 1, filteredVals.length, lastCol).setValues(filteredVals);
  copy.getRange(2, 1, filteredVals.length, lastCol).setNumberFormats(filteredFmts);

  try { copy.setFrozenRows(1); } catch(e) {}
  try { copy.autoResizeColumns(1, lastCol); } catch(e) {}

  return copy;
}

function nameFromUUID_(uuid) {
  try {
    const ctx = CORE_buildContext_();
    const contributors = READ_readContributors_(ctx); // Map<uuid,{name:...}>
    const c = contributors.get(uuid);
    return c ? CORE_properCase_(c.name || uuid) : uuid;
  } catch (e) {
    return uuid; // safe fallback if data layer isn't available
  }
}

/* =========================================================
   OVERVIEW (single workbook) — executed once per export
   ========================================================= */

function FULL_writeOverviewWorkbookByState_() {
  let st = EXPORT_stateLoad_();
  if (!st) return;

  const src = SpreadsheetApp.getActiveSpreadsheet();
  const exportFolder = DriveApp.getFolderById(st.folderId);

  const ss = FULL_newSpreadsheetInFolder_(FULL_EXPORT.OVERVIEW_FILE_NAME(new Date()), st.folderId);

  const N = FULL_EXPORT.SHEET_NAMES;

  FULL_copyWholeSheetIfNonEmpty_(src.getSheetByName(N.CONTRIBUTORS),  ss, N.CONTRIBUTORS);
  FULL_copyWholeSheetIfNonEmpty_(src.getSheetByName(N.PERF_OVERVIEW), ss, N.PERF_OVERVIEW);
  FULL_copyWholeSheetIfNonEmpty_(src.getSheetByName(N.SOC_ISSUES),    ss, N.SOC_ISSUES);
  FULL_copyWholeSheetIfNonEmpty_(src.getSheetByName(N.NON_HOME),      ss, N.NON_HOME);
  FULL_copyWholeSheetIfNonEmpty_(src.getSheetByName(N.REC_OVERVIEW),  ss, N.REC_OVERVIEW);
  FULL_copyWholeSheetIfNonEmpty_(src.getSheetByName(N.REC_ISSUES),    ss, N.REC_ISSUES);
  FULL_copyWholeSheetIfNonEmpty_(src.getSheetByName(N.REC_MOST),      ss, N.REC_MOST);

  FULL_removeDefaultIfSafe_(ss);

  st.overviewDone = true;
  EXPORT_stateSave_(st);
  EXPORT_logProgress_(`Overview workbook created: ${exportFolder.getUrl()}`);
}

/* =========================================================
   BATCH PROCESSORS
   ========================================================= */

/** Returns true if more performers remain (schedules next). */
function processPerformerBatch_(st) {
  const src = SpreadsheetApp.getActiveSpreadsheet();
  const N = FULL_EXPORT.SHEET_NAMES;
  const perfSheet = src.getSheetByName(N.PERF_OVERVIEW);
  const socIssues = src.getSheetByName(N.SOC_ISSUES);
  const nonHome   = src.getSheetByName(N.NON_HOME);
  const recOver   = src.getSheetByName(N.REC_OVERVIEW);
  const recIssues = src.getSheetByName(N.REC_ISSUES);
  const recMost   = src.getSheetByName(N.REC_MOST);

  const start = st.performerIdx || 0;
  const end = Math.min(start + FULL_EXPORT.BATCH.PERFORMERS_PER_RUN, st.performerIds.length);

  for (let i = start; i < end; i++) {
    const uuid = st.performerIds[i];
    const fileName = `${nameFromUUID_(uuid)} - ${uuid}`;
    const ss = FULL_newSpreadsheetInFolder_(fileName, st.performerFolderId);

    FULL_copyFilteredWithFormatting_(perfSheet, ss, 'Performer Overview',
      o => String(o.get('UUID')).trim() === uuid
    );
    FULL_copyFilteredWithFormatting_(socIssues, ss, N.SOC_ISSUES,
      o => String(o.get('UUID')).trim() === uuid
    );
    FULL_copyFilteredWithFormatting_(nonHome, ss, N.NON_HOME,
      o => String(o.get('UUID')).trim() === uuid
    );
    FULL_copyFilteredWithFormatting_(recOver, ss, N.REC_OVERVIEW,
      o => String(o.get('UUID')).trim() === uuid
    );
    FULL_copyFilteredWithFormatting_(recIssues, ss, N.REC_ISSUES,
      o => String(o.get('UUID')).trim() === uuid
    );
    FULL_copyFilteredWithFormatting_(recMost, ss, N.REC_MOST,
      o => String(o.get('UUID')).trim() === uuid
    );

    FULL_removeDefaultIfSafe_(ss);
    EXPORT_logProgress_(`Performer exported: ${fileName}`);

    // NEW: store performer file URL in state for Monday comments
    try {
      const url = ss.getUrl();
      const st2 = EXPORT_stateLoad_() || st;
      st2.perfFileUrlByUUID = st2.perfFileUrlByUUID || {};
      st2.perfFileUrlByUUID[uuid] = url;
      EXPORT_stateSave_(st2);
      st = st2;
    } catch (e) {
      // non-fatal
    }
  }

  st.performerIdx = end;
  EXPORT_stateSave_(st);

  const more = end < st.performerIds.length;
  if (more) EXPORT_logProgress_(`Performer batch done: ${end}/${st.performerIds.length}`);
  return more;
}

/** Returns true if more societies remain (schedules next). */
function processSocietyBatch_(st) {
  const src = SpreadsheetApp.getActiveSpreadsheet();
  const N = FULL_EXPORT.SHEET_NAMES;
  const perfSheet = src.getSheetByName(N.PERF_OVERVIEW);
  const socIssues = src.getSheetByName(N.SOC_ISSUES);
  const nonHome   = src.getSheetByName(N.NON_HOME);
  const recOver   = src.getSheetByName(N.REC_OVERVIEW);
  const recIssues = src.getSheetByName(N.REC_ISSUES);
  const recMost   = src.getSheetByName(N.REC_MOST);

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const start = st.societyIdx || 0;
  const end = Math.min(start + FULL_EXPORT.BATCH.SOCIETIES_PER_RUN, st.societies.length);

  for (let i = start; i < end; i++) {
    const society = st.societies[i];
    const fileName = `${society} - Bulk Performer Analysis - ${today}`;
    const ss = FULL_newSpreadsheetInFolder_(fileName, st.societyFolderId);

    FULL_copyFilteredWithFormatting_(perfSheet, ss, 'Society Overview',
      o => CORE_canonSociety_(String(o.get('Society') || '')) === society
    );
    FULL_copyFilteredWithFormatting_(socIssues, ss, N.SOC_ISSUES,
      o => CORE_canonSociety_(String(o.get('Society') || '')) === society
    );
    FULL_copyFilteredWithFormatting_(nonHome, ss, N.NON_HOME,
      o => CORE_canonSociety_(String(o.get('Society') || '')) === society
    );
    FULL_copyFilteredWithFormatting_(recOver, ss, N.REC_OVERVIEW,
      o => CORE_canonSociety_(String(o.get('Society') || '')) === society
    );
    FULL_copyFilteredWithFormatting_(recIssues, ss, N.REC_ISSUES,
      o => CORE_canonSociety_(String(o.get('Society') || '')) === society
    );
    FULL_copyFilteredWithFormatting_(recMost, ss, N.REC_MOST,
      o => CORE_canonSociety_(String(o.get('Society') || '')) === society
    );

    FULL_removeDefaultIfSafe_(ss);
    EXPORT_logProgress_(`Society exported: ${fileName}`);

    // NEW: store society file URL in state (kept for future use)
    try {
      const url = ss.getUrl();
      const st2 = EXPORT_stateLoad_() || st;
      st2.societyFileUrlByName = st2.societyFileUrlByName || {};
      st2.societyFileUrlByName[society] = url;
      EXPORT_stateSave_(st2);
      st = st2;
    } catch (e) {
      // non-fatal
    }
  }

  st.societyIdx = end;
  EXPORT_stateSave_(st);

  const more = end < st.societies.length;
  if (more) EXPORT_logProgress_(`Society batch done: ${end}/${st.societies.length}`);
  return more;
}

/* =========================================================
   MONDAY PHASE (AFTER societies)
   ========================================================= */

// Avoid repeating a warning
let __MONDAY_HELPERS_MISSING_WARNED__ = false;

/* =========================
   Monday raising — rule gate
   ========================= */

const MON_RULES = {
  TAB: 'Monday Issues Rules',
  REL_ALLOWED: new Set(['good', 'escalation in progress']) // corrected per your note
};

function RULES_norm_(s) { return String(s||'').trim().toLowerCase(); }
function RULES_normStatus_(s) {
  const v = RULES_norm_(s);
  if (v.startsWith('miss')) return 'missing';
  if (v.startsWith('under')) return 'underperforming';
  return v;
}

/** Load matrix key -> boolean from the "Monday Issues Rules" tab. */
function RULES_loadMatrix_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(MON_RULES.TAB);
  if (!sh) throw new Error(`Missing rules sheet: ${MON_RULES.TAB}`);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return new Map();

  const hdr = vals[0].map(x => String(x||'').trim().toLowerCase());
  const cPri = hdr.indexOf('priority (from society list in monday)');
  const cSta = hdr.indexOf('status (from society issues)');
  const cAcc = hdr.indexOf('account priority (from salesforce status)');
  const cYes = hdr.indexOf('raise?');
  if (cPri===-1 || cSta===-1 || cAcc===-1 || cYes===-1) {
    throw new Error('Rules sheet headers not found (check exact header text).');
  }

  const m = new Map();
  for (let r=1;r<vals.length;r++) {
    const pri = RULES_norm_(vals[r][cPri]);
    const sta = RULES_normStatus_(vals[r][cSta]);
    const acc = RULES_norm_(vals[r][cAcc]);
    const yn  = RULES_norm_(vals[r][cYes]);
    if (!pri || !sta || !acc) continue;
    const key = `${pri}|${sta}|${acc}`;
    m.set(key, yn === 'yes');
  }
  return m;
}

/** Read Account Priority for a performer UUID from Performer Overview. */
function EXPORT_accountPriorityForUUID_(uuid) {
  const sh = SpreadsheetApp.getActive().getSheetByName(FULL_EXPORT.SHEET_NAMES.PERF_OVERVIEW);
  if (!sh) return '';
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return '';
  const hdr = vals[0].map(x => String(x||'').trim().toLowerCase());
  const cUUID = hdr.indexOf('uuid');
  let cAcc = hdr.indexOf('account priority (from salesforce status)');
  if (cAcc === -1) cAcc = hdr.indexOf('account priority'); // fallback
  if (cUUID===-1 || cAcc===-1) return '';

  for (let r=1;r<vals.length;r++) {
    if (String(vals[r][cUUID]).trim() === String(uuid).trim()) {
      return String(vals[r][cAcc] || '').trim();
    }
  }
  return '';
}

/** Fetch Society Priority + Relationship (text) from Society List board item. */
function MI_getSocietyPriorityAndRelationship_(societyItemId) {
  const boardId = MON_CFG.SOCIETY_LIST_BOARD_ID;
  const colsMeta = MI_getBoardColumns_(boardId);
  const cvs = MI_getItemColumnValues_(String(societyItemId));

  function byTitle(part) {
    const p = RULES_norm_(part);
    // priority: prefer a status/color/dropdown whose title contains 'priority'
    // relationship: title contains 'relationship'
    return colsMeta.find(c => RULES_norm_(c.title).includes(p));
  }
  const priMeta = byTitle('priority');
  const relMeta = byTitle('relationship');

  function readText(meta) {
    if (!meta) return '';
    const cv = cvs.find(v => String(v.id) === String(meta.id));
    return (cv && cv.text) ? String(cv.text).trim() : '';
  }
  return {
    priority: readText(priMeta),
    relationship: readText(relMeta)
  };
}

/** Decide if a row should raise. Mirrors the diagnostic logic. */
function RULES_shouldRaise_(row, maps) {
  const { uuid, society, status } = row;
  const rules = maps.rules;
  const societyId = maps.societyMap.get(MI_norm_(society)) ||
                    maps.societyMap.get(MI_norm_(MI_stripParen_(society)));

  if (!societyId) {
    return { raise:false, why:`No Society ID for "${society}"` };
  }
  const socMeta = MI_getSocietyPriorityAndRelationship_(String(societyId));
  const rel = RULES_norm_(socMeta.relationship);
  if (!MON_RULES.REL_ALLOWED.has(rel)) {
    return { raise:false, why:`Relationship "${socMeta.relationship}" not allowed` };
  }

  const pri = RULES_norm_(socMeta.priority);
  const accRaw = EXPORT_accountPriorityForUUID_(uuid);
  const acc = RULES_norm_(accRaw);
  const sta = RULES_normStatus_(status);

  const key = `${pri}|${sta}|${acc}`;
  const hasRule = rules.has(key);
  const allow = !!rules.get(key);

  if (!hasRule) return { raise:false, why:`No rule for key ${pri}|${sta}|${acc}` };
  if (!allow)   return { raise:false, why:`Rule says No for ${pri}|${sta}|${acc}` };

  return { raise:true, why:`Rule says Yes for ${pri}|${sta}|${acc}` };
}


/** Process a batch of performers and push *all* their Society Issues to Monday. */
function processMondayBatch_(st) {
  // If disabled or already done, nothing to do
  if (!st.mondayEnabled || st.mondayDone) return false;

  // Check monday helpers are available; if not, skip this phase entirely
  if (typeof MON_CFG === 'undefined' ||
      typeof MI_gql_ !== 'function' ||
      typeof MI_loadNameIdMapFromBoard_ !== 'function' ||
      typeof MI_getSocietyOwnerUserId_ !== 'function' ||
      typeof MI_findOrCreateParentItem_ !== 'function' ||
      typeof MI_createSubitemWithAllValues_ !== 'function') {
    if (!__MONDAY_HELPERS_MISSING_WARNED__) {
      EXPORT_logProgress_('Monday integration not loaded (MON_CFG/MI_* missing). Skipping Monday.', false);
      __MONDAY_HELPERS_MISSING_WARNED__ = true;
    }
    st.mondayDone = true;
    EXPORT_stateSave_(st);
    return false;
  }

  const start = st.mondayIdx || 0;
  const end = Math.min(start + FULL_EXPORT.BATCH.MONDAY_PERFORMERS_PER_RUN, st.performerIds.length);

  for (let i = start; i < end; i++) {
    const uuid = st.performerIds[i];
    try {
      EXPORT_pushMondayForUUID_(uuid); // de-dupes subitems by title under each parent
    } catch (e) {
      EXPORT_logProgress_(`Monday push failed for ${nameFromUUID_(uuid)} (${uuid}): ${e}`);
    }
  }

  st.mondayIdx = end;
  st.mondayDone = (end >= st.performerIds.length);
  EXPORT_stateSave_(st);

  const more = !st.mondayDone;
  if (more) EXPORT_logProgress_(`Monday batch done: ${end}/${st.performerIds.length}`);
  else EXPORT_logProgress_('Monday phase completed.');
  return more;
}

/** Push ALL Society Issues rows for a performer UUID, with de-dup + comments + links. */
function EXPORT_pushMondayForUUID_(uuid) {
  // 1) Read all Society Issues rows for this UUID
  const rows = EXPORT_societyIssuesRowsForUUID_(uuid);
  if (!rows.length) return;

  // 2) Resolve lookup maps once (Contributor & Society boards)
  const contribMap = MI_loadNameIdMapFromBoard_(MON_CFG.CONTRIBUTOR_LIST_BOARD_ID);
  const societyMap = MI_loadNameIdMapFromBoard_(MON_CFG.SOCIETY_LIST_BOARD_ID);

  // 3) Load rules matrix once
  const rules = RULES_loadMatrix_();

  // 4) Performer file URL (for comments)
  const st = EXPORT_stateLoad_() || {};
  const perfUrlMap = st.perfFileUrlByUUID || {};
  const performerFileUrl = perfUrlMap[uuid] || '';

  // 5) Caches to limit API calls per society
  const parentIdBySoc = new Map();    // societyName → parent item id
  const ownerIdBySocItem = new Map(); // societyItemId → people userId
  const subitemNameSets = new Map();  // parentId → Set(existing subitem names)

  rows.forEach(r => {
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

      if (!contribId) console.log(`Monday: UUID ${uuid}: no contributor match for "${contributorName}"`);
      if (!societyId) console.log(`Monday: UUID ${uuid}: no society match for "${societyName}"`);

      // -------- RULE GATE (mirrors diagnostics) --------
      const gate = RULES_shouldRaise_(
        { uuid, society: societyName, status: r.status },
        { rules, societyMap }
      );
      if (!gate.raise) {
        console.log(`Monday: SKIP — ${societyName} | ${contributorName} — ${gate.why}`);
        return;
      }

      // -------- Find/create parent (ensures main fields) --------
      let parentId = parentIdBySoc.get(societyName);
      if (!parentId) {
        parentId = MI_findOrCreateParentItem_(societyName);
        parentIdBySoc.set(societyName, parentId);
      }

      // De-dup by subitem title under the parent
      let existing = subitemNameSets.get(parentId);
      if (!existing) {
        existing = EXPORT_fetchSubitemNameSet_(parentId);
        subitemNameSets.set(parentId, existing);
      }
      const subTitle = `${societyName} – ${contributorName}${territory ? ` (${territory})` : ''}`;
      if (existing.has(subTitle)) {
        console.log(`Monday: SKIP duplicate subitem title under ${societyName}: "${subTitle}"`);
        return;
      }

      // Society Owner (People)
      let ownerUserId = null;
      if (societyId) {
        ownerUserId = ownerIdBySocItem.get(societyId);
        if (typeof ownerUserId === 'undefined') {
          ownerUserId = MI_getSocietyOwnerUserId_(String(societyId)) || null;
          ownerIdBySocItem.set(societyId, ownerUserId);
        }
      }

      // Status/Dropdown mapping + Review Date
      const statusLabel = 'For Society';
      const trackingLabels =
        (String(r.status || '').toLowerCase().startsWith('missing')) ? ['Missing Contributor']
          : (String(r.status || '').toLowerCase().startsWith('under') ? ['Low Income'] : []);
      const reviewISO = (typeof MI_isoDatePlusDays_ === 'function') ? MI_isoDatePlusDays_(14)
                                                                     : EXPORT_isoDatePlusDays_(14);

      // Create subitem
      const subId = MI_createSubitemWithAllValues_(
        parentId, subTitle, contribId, societyId, ownerUserId, statusLabel, trackingLabels, reviewISO
      );
      existing.add(subTitle);
      console.log(`Monday: RAISE — ${societyName} | ${contributorName} — ${gate.why} → subitem ${subId}`);

      // Set links + post comment
      if (typeof MI_buildCommentForRow_ === 'function' && typeof MI_postUpdate_ === 'function') {
        const type = (trackingLabels[0] === 'Missing Contributor') ? 'missing' : 'low';
        const comment = MI_buildCommentForRow_(r, type, {
          performerFileUrl: performerFileUrl,
          reviewDate: reviewISO,
          ownerName: 'Harry Stanhope'
        });
        MI_postUpdate_(subId, comment);
      }

      // Parent/Sub links (already added in a previous patch if present)
      try {
        if (typeof MI_setAnalysisLinksForItems_ === 'function') {
          const parentFileUrl = (st.societyFileUrlByName || {})[societyName] || '';
          MI_setAnalysisLinksForItems_({ parentId, parentFileUrl, subId, performerFileUrl });
        }
      } catch (e) {
        console.log('Monday: WARN setting links: ' + e);
      }

      Utilities.sleep(120);
    } catch (e) {
      console.log(`Monday: error for UUID ${uuid} row ${r.rowIndex}: ${e}`);
    }
  });

  EXPORT_logProgress_(`Monday: processed ${rows.length} issues for ${nameFromUUID_(uuid)} (${uuid})`);
}



/** Read all Society Issues rows for a given UUID. */
function EXPORT_societyIssuesRowsForUUID_(uuid) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(FULL_EXPORT.SHEET_NAMES.SOC_ISSUES);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];

  const headers = vals[0].map(h => String(h || '').toLowerCase().trim());
  const col = (label) => headers.indexOf(String(label).toLowerCase());
  const iUUID = col('uuid');
  if (iUUID === -1) return [];

  const out = [];
  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][iUUID] || '').trim() !== uuid) continue;
    out.push({
      rowIndex: r + 1,
      contributor: vals[r][col('contributor')],
      uuid:        vals[r][iUUID],
      society:     vals[r][col('society')],
      territory:   vals[r][col('territory')],
      status:      vals[r][col('status')],
      notes:       vals[r][col('notes')]
    });
  }
  return out;
}

/** Fetch existing subitem names Set for a parent item (to avoid duplicates). */
function EXPORT_fetchSubitemNameSet_(parentItemId) {
  try {
    const q = `
      query ($ids:[ID!]!) {
        items(ids:$ids) {
          id
          subitems { id name }
        }
      }`;
    const data = MI_gql_(q, { ids: [String(parentItemId)] });
    const arr = data?.items?.[0]?.subitems || [];
    return new Set(arr.map(s => String(s.name || '')));
  } catch (e) {
    console.log('Monday: could not fetch existing subitems for ' + parentItemId + ': ' + e);
    return new Set();
  }
}

/** Tiny ISO date helper if MI_isoDatePlusDays_ isn’t available. */
function EXPORT_isoDatePlusDays_(days) {
  const tz = Session.getScriptTimeZone() || 'UTC';
  const d = new Date();
  d.setDate(d.getDate() + (Number(days) || 0));
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

/* =========================================================
   CLEANUP: remove overview tabs & Export Log from main file
   ========================================================= */
function EXPORT_cleanupTemporarySheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = [
    FULL_EXPORT.SHEET_NAMES.PERF_OVERVIEW,
    FULL_EXPORT.SHEET_NAMES.SOC_ISSUES,
    FULL_EXPORT.SHEET_NAMES.NON_HOME,
    FULL_EXPORT.SHEET_NAMES.REC_OVERVIEW,
    FULL_EXPORT.SHEET_NAMES.REC_ISSUES,
    FULL_EXPORT.SHEET_NAMES.REC_MOST,
    'Export Log'
  ];
  const minSheetsToKeep = 1; // safety

  names.forEach(n => {
    try {
      const sh = ss.getSheetByName(n);
      if (!sh) return;
      if (ss.getSheets().length <= minSheetsToKeep) return; // don't delete last sheet
      ss.deleteSheet(sh);
    } catch (e) {
      // ignore, keep going
    }
  });
}
