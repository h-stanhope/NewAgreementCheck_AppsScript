/*********************************************************
 * FullExport Test Harness (non-invasive)
 * - Limits data volume for quick end-to-end testing
 * - Drives the real FullExport pipeline with sliced lists
 * - Optional tiny Monday push limited per-contributor
 * - Does not modify existing files or functions
 **********************************************************/

/**
 * Start a mini FullExport using the same pipeline/state,
 * but only for the first N contributors and first M societies.
 *
 * @param {number} maxPerformers - default 1
 * @param {number} maxSocieties  - default 5
 * @param {boolean} mondayInsidePipeline - default true. If false, pipeline skips Monday
 *        and you can run TEST_runMiniMondayPush() afterwards for a tiny, capped push.
 */
function TEST_startMiniFullExport(maxPerformers, maxSocieties, mondayInsidePipeline) {
  maxPerformers = Number(maxPerformers || 1);
  maxSocieties  = Number(maxSocieties  || 5);
  const runMonday = (mondayInsidePipeline === undefined) ? true : !!mondayInsidePipeline;

  // Refresh overviews (same as FullExport)
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
    console.warn('TestHarness: overview refresh warning:', e);
  }

  // Build the *same* folder tree & init state
  const startedAt   = new Date();
  const exportFolder = FULL_buildFolderTree_(startedAt);
  const exportFolderId = exportFolder.getId();

  // First N contributors (sorted by name for determinism)
  const { uuids: performerIds } = __TEST_pickFirstContributors_(maxPerformers);
  if (!performerIds.length) throw new Error('TestHarness: no contributors found.');

  // First M societies *seen for those contributors* on Society Issues
  const societies = __TEST_firstSocietiesForContributors_(performerIds, maxSocieties);

  // Persist state (exact shape the pipeline expects)
  EXPORT_stateSave_({
    startedAt: startedAt.toISOString(),
    folderId: exportFolderId,
    performerFolderId: FULL_getOrCreateChild_(exportFolder, FULL_EXPORT.PERFORMER_BREAKDOWN).getId(),
    societyFolderId:   FULL_getOrCreateChild_(exportFolder, FULL_EXPORT.SOCIETY_BREAKDOWN).getId(),
    overviewDone: false,

    // Performer stage (LIMITED)
    performerIds,
    performerIdx: 0,

    // Society stage (LIMITED)
    societies,
    societyIdx: 0,

    // File URL indexes (populated by FullExport during export)
    perfFileUrlByUUID: {},
    societyFileUrlByName: {},

    // Monday stage
    mondayEnabled: !!runMonday,
    mondayIdx: 0,
    mondayDone: !runMonday
  });

  // Create overview workbook once (same writer)
  FULL_writeOverviewWorkbookByState_();

  // Kick the normal pipeline (will handle performers → societies → monday → cleanup)
  EXPORT_continueFullAnalysis_();
  SpreadsheetApp.getActive().toast(`Mini FullExport started: ${performerIds.length} contribs, ${societies.length} societies. Monday=${runMonday}`);
}

/**
 * After a mini export (often with mondayInsidePipeline = false), run a *tiny*
 * Monday push for the first N contributors, limited to K rows each.
 *
 * This reuses the MondayIntegration helpers and posts the same subitems, owner/status/type,
 * review date, and the rich comment (including performer file URL from state).
 *
 * @param {number} maxPerformers          - default 1
 * @param {number} maxRowsPerContributor  - default 5
 */
function TEST_runMiniMondayPush(maxPerformers, maxRowsPerContributor) {
  maxPerformers = Number(maxPerformers || 1);
  maxRowsPerContributor = Number(maxRowsPerContributor || 5);

  // Ensure Monday helpers exist
  if (typeof MON_CFG === 'undefined' ||
      typeof MI_gql_ !== 'function' ||
      typeof MI_loadNameIdMapFromBoard_ !== 'function' ||
      typeof MI_getSocietyOwnerUserId_ !== 'function' ||
      typeof MI_findOrCreateParentItem_ !== 'function' ||
      typeof MI_createSubitemWithAllValues_ !== 'function') {
    throw new Error('Monday integration helpers (MON_CFG / MI_*) are not loaded.');
  }

  // Get first N contributors (same as TEST_startMiniFullExport selection)
  const { uuids: performerIds } = __TEST_pickFirstContributors_(maxPerformers);
  if (!performerIds.length) throw new Error('TestHarness: no contributors found for Monday push.');

  // Resolve lookup maps once
  const contribMap = MI_loadNameIdMapFromBoard_(MON_CFG.CONTRIBUTOR_LIST_BOARD_ID);
  const societyMap = MI_loadNameIdMapFromBoard_(MON_CFG.SOCIETY_LIST_BOARD_ID);

  // Try to read performer file URL map saved by FullExport
  const st = EXPORT_stateLoad_() || {};
  const perfUrlByUUID = st.perfFileUrlByUUID || {};

  performerIds.forEach(uuid => {
    try {
      // First K Society Issues rows for this contributor
      const rows = __TEST_readSocietyIssuesRowsForUUID_Limited_(uuid, maxRowsPerContributor);
      if (!rows.length) {
        console.log(`TestHarness/Monday: no Society Issues rows for ${uuid}`);
        return;
      }

      // A small cache for this contributor’s run
      const parentIdBySoc  = new Map();
      const ownerIdBySocItem = new Map();
      const subitemNameSets = new Map();

      const performerFileUrl = perfUrlByUUID[uuid] || '';

      rows.forEach(r => {
        const contributorName = String(r.contributor || '').trim();
        const societyName     = String(r.society || '').trim();
        const territory       = String(r.territory || '').trim();

        const contribId =
          contribMap.get(__TEST_norm_(contributorName)) ||
          contribMap.get(__TEST_norm_(__TEST_stripParen_(contributorName))) || null;

        const societyId =
          societyMap.get(__TEST_norm_(societyName)) ||
          societyMap.get(__TEST_norm_(__TEST_stripParen_(societyName))) || null;

        if (!contribId) console.log(`TestHarness/Monday: no contributor match for "${contributorName}"`);
        if (!societyId) console.log(`TestHarness/Monday: no society match for "${societyName}"`);

        // Parent per society
        let parentId = parentIdBySoc.get(societyName);
        if (!parentId) {
          parentId = MI_findOrCreateParentItem_(societyName);
          parentIdBySoc.set(societyName, parentId);
        }

        // De-dup existing subitems under this parent
        let existing = subitemNameSets.get(parentId);
        if (!existing) {
          existing = __TEST_fetchSubitemNameSet_(parentId);
          subitemNameSets.set(parentId, existing);
        }

        const subTitle = `${societyName} – ${contributorName}${territory ? ` (${territory})` : ''}`;
        if (existing.has(subTitle)) return;

        // Owner from society board
        let ownerUserId = null;
        if (societyId) {
          ownerUserId = ownerIdBySocItem.get(societyId);
          if (typeof ownerUserId === 'undefined') {
            ownerUserId = MI_getSocietyOwnerUserId_(String(societyId)) || null;
            ownerIdBySocItem.set(societyId, ownerUserId);
          }
        }

        // Map status → dropdown label + 2-week review date (same as FullExport)
        const statusLabel = 'For Society';
        const trackingLabels =
          (String(r.status || '').toLowerCase().startsWith('missing')) ? ['Missing Contributor']
            : (String(r.status || '').toLowerCase().startsWith('under') ? ['Low Income'] : []);
        const reviewISO = (typeof MI_isoDatePlusDays_ === 'function')
          ? MI_isoDatePlusDays_(14)
          : __TEST_isoDatePlusDays_(14);

        // Create the subitem (sets owner, status, dropdown, review date)
        const subId = MI_createSubitemWithAllValues_(
          parentId, subTitle, contribId, societyId, ownerUserId, statusLabel, trackingLabels, reviewISO
        );
        existing.add(subTitle);
        console.log(`TestHarness/Monday: created subitem ${subId} for ${subTitle}`);

        // Post the rich comment with performer file link, if helpers exist
        if (typeof MI_buildCommentForRow_ === 'function' && typeof MI_postUpdate_ === 'function') {
          const type = (trackingLabels[0] === 'Missing Contributor') ? 'missing' : 'low';
          const comment = MI_buildCommentForRow_(r, type, {
            performerFileUrl: performerFileUrl,
            reviewDate: reviewISO,
            ownerName: 'Harry Stanhope'
          });
          MI_postUpdate_(subId, comment);
        }

        Utilities.sleep(150);
      });

      SpreadsheetApp.getActive().toast(`TestHarness/Monday: pushed ${rows.length} for ${nameFromUUID_(uuid)} (${uuid})`);
    } catch (e) {
      console.log(`TestHarness/Monday: error for ${uuid}: ${e}`);
    }
  });
}

/** Optional: remove the overview tabs & Export Log after a test run */
function TEST_cleanupMini() {
  EXPORT_cleanupTemporarySheets_();
  SpreadsheetApp.getActive().toast('Mini cleanup done.');
}

/* ===========================
   Small local helpers
   =========================== */

function __TEST_pickFirstContributors_(maxPerformers) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(FULL_EXPORT.SHEET_NAMES.CONTRIBUTORS) || ss.getSheetByName('Contributor List');
  if (!sh) throw new Error('TestHarness: missing "Contributor List" sheet.');

  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) throw new Error('TestHarness: "Contributor List" has no data rows.');

  const headers = vals[0].map(h => String(h||'').trim().toLowerCase());
  let iUuid = headers.findIndex(h => /\buuid\b/.test(h));
  if (iUuid === -1) iUuid = headers.findIndex(h => /(performer|contributor).*\buuid\b/.test(h));
  if (iUuid === -1) iUuid = headers.findIndex(h => h === 'id' || h.endsWith(' id'));
  if (iUuid === -1) throw new Error('TestHarness: could not find UUID column.');

  let iName = headers.findIndex(h => /(contributor|performer).*\bname\b/.test(h));
  if (iName === -1) iName = headers.findIndex(h => h === 'name' || h.endsWith(' name'));

  const rows = [];
  for (let r = 1; r < vals.length; r++) {
    const uuid = String(vals[r][iUuid] || '').trim();
    if (!uuid) continue;
    const name = iName >= 0 ? String(vals[r][iName] || '').trim() : uuid;
    rows.push({ uuid, name });
  }
  rows.sort((a,b) => String(a.name||a.uuid).localeCompare(String(b.name||b.uuid)));

  const limited = rows.slice(0, Math.max(1, maxPerformers));
  return { uuids: limited.map(x => x.uuid), names: limited.map(x => x.name) };
}

function __TEST_firstSocietiesForContributors_(uuids, maxSocieties) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(FULL_EXPORT.SHEET_NAMES.SOC_ISSUES) || ss.getSheetByName('Society Issues');
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];

  const headers = vals[0].map(h => String(h||'').trim().toLowerCase());
  const iUUID = headers.indexOf('uuid');
  const iSoc  = headers.indexOf('society');
  if (iUUID === -1 || iSoc === -1) return [];

  const want = Math.max(1, Number(maxSocieties||5));
  const set = new Set();
  for (let r = 1; r < vals.length && set.size < want; r++) {
    const rowUUID = String(vals[r][iUUID] || '').trim();
    if (!uuids.includes(rowUUID)) continue;
    const soc = CORE_canonSociety_(String(vals[r][iSoc] || ''));
    if (!soc) continue;
    if (!set.has(soc)) set.add(soc);
  }
  return Array.from(set);
}

function __TEST_readSocietyIssuesRowsForUUID_Limited_(uuid, maxRows) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(FULL_EXPORT.SHEET_NAMES.SOC_ISSUES) || ss.getSheetByName('Society Issues');
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];

  const headers = vals[0].map(h => String(h || '').toLowerCase().trim());
  const col = (label) => headers.indexOf(String(label).toLowerCase());
  const iUUID = col('uuid');
  if (iUUID === -1) return [];

  const out = [];
  const want = Math.max(1, Number(maxRows || 5));
  for (let r = 1; r < vals.length && out.length < want; r++) {
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

function __TEST_fetchSubitemNameSet_(parentItemId) {
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
    console.log('TestHarness: could not fetch existing subitems for ' + parentItemId + ': ' + e);
    return new Set();
  }
}

function __TEST_norm_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/\(.*?\)/g, '')
    .replace(/&/g, 'and')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}
function __TEST_stripParen_(s) {
  return String(s || '').replace(/\(.*?\)/g, '').trim();
}
function __TEST_isoDatePlusDays_(days) {
  const tz = Session.getScriptTimeZone() || 'UTC';
  const d = new Date();
  d.setDate(d.getDate() + (Number(days) || 0));
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

function MONDAY__smokeTestQueries() {
  try {
    // 1) columns of Society List board
    const cols = MI_getBoardColumns_(MON_CFG.SOCIETY_LIST_BOARD_ID);
    console.log('Columns loaded:', cols.length);

    // 2) first society item id from the map + its column_values
    const socMap = MI_loadNameIdMapFromBoard_(MON_CFG.SOCIETY_LIST_BOARD_ID);
    const firstSocId = socMap.values().next().value;
    const cvals = MI_getItemColumnValues_(firstSocId);
    console.log('First society item column_values:', cvals.length);

    // 3) ensure we can list/create a parent
    const pid = MI_findOrCreateParentItem_('TEST SOCIETY');
    console.log('Parent item id:', pid);
  } catch (e) {
    console.error('Smoke test failed:', e);
  }
}
