/*********************************************************
 * Pilot: run FullExport-style flow for ONE contributor,
 * limited to their first N societies, then push N Monday
 * subitems from Society Issues for that contributor.
 *
 * FILE LINKS:
 *  - Parent item ← Society workbook (Google Sheet URL)
 *  - Subitem     ← Performer workbook (Google Sheet URL)
 * If Link columns are provided, we set them; else we post
 * an update with the URL.
 **********************************************************/

const PILOT = {
  SOCIETY_LIMIT: 5,     // how many societies for this contributor
  MONDAY_ROWS:  5,      // how many Society Issues rows to send to Monday

  // Optional Link columns (leave blank to post updates instead)
  // Parent board (main item)
  PARENT_LINK_COL_ID: '',   // e.g. 'link' on parent board
  // Subitem board (sub-items)
  SUBITEM_LINK_COL_ID: ''   // e.g. 'link' on subitem board
};

/** MAIN entrypoint */
function PILOT_runOneContributorTest() {
  // 0) Build/refresh overviews exactly like FullExport
  try {
    USECASE_buildOverviewOnly();
    SOC_buildSocietyIssuesOverview();
    SOC_buildNonHomeTerritoryIncome();
    REC_buildRecordingOverview();
    REC_buildRecordingIssues();
    if (typeof REC_buildMostPlayedRecordings === 'function') {
      REC_buildMostPlayedRecordings(); // if present
    }
  } catch (e) {
    console.log('Pilot: overview refresh warning: ' + e);
  }

  // 1) Pick ONE contributor (robust)
  const picked = PILOT_pickFirstContributor_();
  if (!picked || !picked.uuid) {
    throw new Error('Pilot: could not determine a contributor UUID.');
  }
  const { uuid, name } = picked;
  console.log(`Pilot: running for ${name} (${uuid})`);

  // 2) Build the same folder tree and overview workbook (same location/names)
  const startedAt = new Date();
  const exportFolder = FULL_buildFolderTree_(startedAt); // same tree
  const folderId = exportFolder.getId();
  const performerFolderId = FULL_getOrCreateChild_(exportFolder, FULL_EXPORT.PERFORMER_BREAKDOWN).getId();
  const societyFolderId   = FULL_getOrCreateChild_(exportFolder, FULL_EXPORT.SOCIETY_BREAKDOWN).getId();

  // Reuse the "overview workbook" writer by seeding minimal state
  EXPORT_stateSave_({
    startedAt: startedAt.toISOString(),
    folderId,
    performerFolderId,
    societyFolderId,
    overviewDone: false,
    performerIds: [uuid],
    performerIdx: 0,
    societies: [],
    societyIdx: 0
  });
  FULL_writeOverviewWorkbookByState_(); // creates Overview workbook once

  // 3) Create the ONE performer workbook (we'll need its URL for subitems)
  const perfSS = P_makePerformerWorkbook_(uuid, performerFolderId);
  const perfUrl = perfSS.getUrl();

  // 4) Determine this contributor’s first N societies,
  //    then export society files and collect their URLs for parent attachments
  const societies = PILOT_firstSocietiesForContributor_(uuid, PILOT.SOCIETY_LIMIT);
  let societyUrlMap = new Map(); // societyName -> URL
  if (societies.length) {
    societyUrlMap = PILOT_exportSocietyFiles_AndReturnUrls_(societies, societyFolderId);
  } else {
    console.log('Pilot: contributor had no societies on Society Issues (after refresh).');
  }

  // 5) Push Monday subitems for this contributor’s first N Society Issues rows
  try {
    PILOT_pushMondayForContributor_(uuid, PILOT.MONDAY_ROWS, perfUrl, societyUrlMap);
  } catch (e) {
    console.log('Pilot: Monday push failed (continuing): ' + e);
  }

  // Done log
  const url = DriveApp.getFolderById(folderId).getUrl();
  EXPORT_logProgress_(`Pilot export (1 contributor) completed: ${url}`, /*final=*/true);
}

/* -------------------------------------------------------
 * Contributor picker (robust)
 * -----------------------------------------------------*/
function PILOT_pickFirstContributor_() {
  // 1) Prefer data-layer if available
  try {
    const ctx = CORE_buildContext_();
    const contribs = READ_readContributors_(ctx); // Map<uuid,{name:...}>
    if (contribs && contribs.size) {
      const first = Array.from(contribs.entries())
        .sort((a,b)=> String(a[1]?.name||a[0]).localeCompare(String(b[1]?.name||b[0])))[0];
      if (first) {
        const [uuid, c] = first;
        return { uuid, name: CORE_properCase_(c?.name || uuid) };
      }
    }
  } catch (e) {
    console.log('Pilot: data-layer pick failed, falling back to sheet. ' + e);
  }

  // 2) Fallback: first non-blank UUID in Contributor List sheet
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(FULL_EXPORT.SHEET_NAMES.CONTRIBUTORS) || ss.getSheetByName('Contributor List');
  if (!sh) throw new Error('Pilot: missing "Contributor List" sheet.');
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) throw new Error('Pilot: "Contributor List" has no data rows.');

  const headers = vals[0].map(h => String(h||'').trim().toLowerCase());
  let iUuid = headers.findIndex(h => /\buuid\b/.test(h));
  if (iUuid === -1) iUuid = headers.findIndex(h => /(performer|contributor).*\buuid\b/.test(h));
  if (iUuid === -1) iUuid = headers.findIndex(h => h === 'id' || h.endsWith(' id'));
  if (iUuid === -1) throw new Error('Pilot: could not find a UUID column on "Contributor List".');

  let iName = headers.findIndex(h => /(contributor|performer).*\bname\b/.test(h));
  if (iName === -1) iName = headers.findIndex(h => h === 'name' || h.endsWith(' name'));

  for (let r = 1; r < vals.length; r++) {
    const uuid = String(vals[r][iUuid] || '').trim();
    if (!uuid) continue;
    const nm = iName >= 0 ? String(vals[r][iName] || '').trim() : uuid;
    return { uuid, name: CORE_properCase_(nm || uuid) };
  }
  throw new Error('Pilot: no UUID rows found in "Contributor List".');
}

/* -------------------------------------------------------
 * Find this contributor’s first N societies (from Society Issues)
 * -----------------------------------------------------*/
function PILOT_firstSocietiesForContributor_(uuid, maxN) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(FULL_EXPORT.SHEET_NAMES.SOC_ISSUES);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];

  const headers = vals[0].map(h => String(h||'').trim().toLowerCase());
  const iUUID = headers.indexOf('uuid');
  const iSoc  = headers.indexOf('society');
  if (iUUID === -1 || iSoc === -1) return [];

  const seen = new Set();
  const list = [];
  for (let r = 1; r < vals.length && list.length < (maxN||5); r++) {
    if (String(vals[r][iUUID] || '').trim() !== uuid) continue;
    const soc = CORE_canonSociety_(String(vals[r][iSoc] || ''));
    if (!soc || seen.has(soc)) continue;
    seen.add(soc);
    list.push(soc);
  }
  console.log(`Pilot: first ${list.length} societies for contributor ${uuid}: ${list.join(', ')}`);
  return list;
}

/* -------------------------------------------------------
 * Make one performer workbook (returns Spreadsheet)
 * -----------------------------------------------------*/
function P_makePerformerWorkbook_(uuid, performerFolderId) {
  const src = SpreadsheetApp.getActiveSpreadsheet();
  const N = FULL_EXPORT.SHEET_NAMES;
  const perfSheet = src.getSheetByName(N.PERF_OVERVIEW);
  const socIssues = src.getSheetByName(N.SOC_ISSUES);
  const nonHome   = src.getSheetByName(N.NON_HOME);
  const recOver   = src.getSheetByName(N.REC_OVERVIEW);
  const recIssues = src.getSheetByName(N.REC_ISSUES);
  const recMost   = src.getSheetByName(N.REC_MOST);

  const fileName = `${nameFromUUID_(uuid)} - ${uuid}`;
  const ss = FULL_newSpreadsheetInFolder_(fileName, performerFolderId);

  FULL_copyFilteredWithFormatting_(perfSheet, ss, 'Performer Overview', o => String(o.get('UUID')).trim() === uuid);
  FULL_copyFilteredWithFormatting_(socIssues, ss, N.SOC_ISSUES,        o => String(o.get('UUID')).trim() === uuid);
  FULL_copyFilteredWithFormatting_(nonHome,   ss, N.NON_HOME,          o => String(o.get('UUID')).trim() === uuid);
  FULL_copyFilteredWithFormatting_(recOver,   ss, N.REC_OVERVIEW,      o => String(o.get('UUID')).trim() === uuid);
  FULL_copyFilteredWithFormatting_(recIssues, ss, N.REC_ISSUES,        o => String(o.get('UUID')).trim() === uuid);
  if (recMost) {
    FULL_copyFilteredWithFormatting_(recMost, ss, N.REC_MOST,          o => String(o.get('UUID')).trim() === uuid);
  }

  FULL_removeDefaultIfSafe_(ss);
  EXPORT_logProgress_(`Pilot: Performer exported: ${fileName}`);
  return ss;
}

/* -------------------------------------------------------
 * Export society files and return a Map(societyName -> URL)
 * -----------------------------------------------------*/
function PILOT_exportSocietyFiles_AndReturnUrls_(societies, societyFolderId) {
  const src = SpreadsheetApp.getActiveSpreadsheet();
  const N = FULL_EXPORT.SHEET_NAMES;
  const perfSheet = src.getSheetByName(N.PERF_OVERVIEW);
  const socIssues = src.getSheetByName(N.SOC_ISSUES);
  const nonHome   = src.getSheetByName(N.NON_HOME);
  const recOver   = src.getSheetByName(N.REC_OVERVIEW);
  const recIssues = src.getSheetByName(N.REC_ISSUES);
  const recMost   = src.getSheetByName(N.REC_MOST);

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const urlMap = new Map();

  societies.forEach(society => {
    const fileName = `${society} - Bulk Performer Analysis - ${today}`;
    const ss = FULL_newSpreadsheetInFolder_(fileName, societyFolderId);

    const pred = o => CORE_canonSociety_(String(o.get('Society') || '')) === society;

    FULL_copyFilteredWithFormatting_(perfSheet, ss, 'Society Overview', pred);
    FULL_copyFilteredWithFormatting_(socIssues, ss, N.SOC_ISSUES,       pred);
    FULL_copyFilteredWithFormatting_(nonHome,   ss, N.NON_HOME,         pred);
    FULL_copyFilteredWithFormatting_(recOver,   ss, N.REC_OVERVIEW,     pred);
    FULL_copyFilteredWithFormatting_(recIssues, ss, N.REC_ISSUES,       pred);
    if (recMost) {
      FULL_copyFilteredWithFormatting_(recMost, ss, N.REC_MOST,         pred);
    }

    FULL_removeDefaultIfSafe_(ss);
    EXPORT_logProgress_(`Pilot: Society exported: ${fileName}`);

    try {
      urlMap.set(society, ss.getUrl());
    } catch (e) {
      console.log(`WARN: could not get URL for ${society}: ${e}`);
    }
  });

  return urlMap;
}

/* -------------------------------------------------------
 * Monday push: first N rows for THIS contributor only
 * + attach Google Sheet URLs (parent/subitem)
 * -----------------------------------------------------*/
function PILOT_pushMondayForContributor_(uuid, maxRows, perfUrl, societyUrlMap /* Map */) {
  if (typeof MON_CFG === 'undefined' || typeof MI_gql_ !== 'function') {
    throw new Error('Monday Integration is not loaded (MON_CFG/MI_* missing).');
  }

  // 1) Read first N rows for this UUID from Society Issues
  const rows = PILOT_readSocietyIssuesRowsForUUID_(uuid, maxRows || 5);
  if (!rows.length) {
    console.log('Pilot: no Society Issues rows for UUID ' + uuid);
    return;
  }

  // 2) Resolve lookup maps once
  const contribMap = MI_loadNameIdMapFromBoard_(MON_CFG.CONTRIBUTOR_LIST_BOARD_ID);
  const societyMap = MI_loadNameIdMapFromBoard_(MON_CFG.SOCIETY_LIST_BOARD_ID);

  const reviewISO = (typeof MI_isoDatePlusDays_ === 'function') ? MI_isoDatePlusDays_(14) : PILOT_isoDatePlusDays_(14);

  // 3) Process rows
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

      if (!contribId) console.log(`Pilot/Monday row ${r.rowIndex}: WARN no contributor match for "${contributorName}"`);
      if (!societyId) console.log(`Pilot/Monday row ${r.rowIndex}: WARN no society match for "${societyName}"`);

      const ownerUserId = societyId ? MI_getSocietyOwnerUserId_(String(societyId)) : null;

      // Parent in Test group (created once per society if not present)
      const parentId = MI_findOrCreateParentItem_(societyName);

      // Attach the society Google Sheet URL to parent:
      const socUrl = societyUrlMap && societyUrlMap.get(societyName);
      if (socUrl) {
        try {
          if (PILOT.PARENT_LINK_COL_ID) {
            P_setLinkColumn_(MON_CFG.BOARD_ID, parentId, PILOT.PARENT_LINK_COL_ID, socUrl, `Society Workbook (${societyName})`);
          } else {
            P_addUpdateWithLink_(parentId, `Society workbook: ${socUrl}`);
          }
        } catch (e) {
          console.log(`WARN: could not attach link to parent (${parentId}): ${e}`);
        }
      }

      // Status/Dropdown mapping
      const statusLabel = 'For Society';
      const trackingLabels =
        (String(r.status || '').toLowerCase().indexOf('missing') === 0)
          ? ['Missing Contributor']
          : (String(r.status || '').toLowerCase().indexOf('under') === 0 ? ['Low Income'] : []);

      const subTitle = `${societyName} – ${contributorName}${territory ? ` (${territory})` : ''}`;

      // Create subitem w/ relations + fields
      const subId = MI_createSubitemWithAllValues_(
        parentId, subTitle, contribId, societyId, ownerUserId, statusLabel, trackingLabels, reviewISO
      );
      console.log(`Pilot/Monday row ${r.rowIndex}: created subitem ${subId} for ${subTitle}`);

      // Attach the performer Google Sheet URL to the subitem
      if (perfUrl) {
        try {
          if (PILOT.SUBITEM_LINK_COL_ID) {
            const subBoardId = P_getBoardIdForItem_(subId);
            P_setLinkColumn_(subBoardId, subId, PILOT.SUBITEM_LINK_COL_ID, perfUrl, `Performer Workbook`);
          } else {
            P_addUpdateWithLink_(subId, `Performer workbook: ${perfUrl}`);
          }
        } catch (e) {
          console.log(`WARN: could not attach link to subitem (${subId}): ${e}`);
        }
      }

      Utilities.sleep(200);
    } catch (e) {
      console.log(`Pilot/Monday row ${r.rowIndex}: ERROR ${e}`);
    }
  });
}

/* Read first N Society Issues rows for a given UUID */
function PILOT_readSocietyIssuesRowsForUUID_(uuid, maxRows) {
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

/* ======== Monday helpers for links/updates ======== */

/** Set a Link column on an item/subitem. */
function P_setLinkColumn_(boardId, itemId, columnId, url, text) {
  const q = `
    mutation ($bid: ID!, $iid: ID!, $cid: String!, $val: JSON!) {
      change_column_value(board_id: $bid, item_id: $iid, column_id: $cid, value: $val) { id }
    }`;
  const val = { url: String(url), text: String(text || url) };
  MI_gql_(q, { bid: String(boardId), iid: String(itemId), cid: String(columnId), val });
}

/** Post a simple update with the link. */
function P_addUpdateWithLink_(itemId, body) {
  const q = `mutation ($iid: ID!, $body: String!) {
    create_update (item_id: $iid, body: $body) { id }
  }`;
  MI_gql_(q, { iid: String(itemId), body: String(body) });
}

/** Fetch the board id that owns a given (sub)item id. */
function P_getBoardIdForItem_(itemId) {
  const q = `query ($ids: [ID!]!) {
    items(ids: $ids) { id board { id } }
  }`;
  const data = MI_gql_(q, { ids: [String(itemId)] });
  const it = data?.items?.[0];
  const bid = it?.board?.id;
  if (!bid) throw new Error('Could not determine board_id for item ' + itemId);
  return String(bid);
}

/* local tiny helper if MI_isoDatePlusDays_ isn’t in scope */
function PILOT_isoDatePlusDays_(days) {
  const tz = Session.getScriptTimeZone() || 'UTC';
  const d = new Date();
  d.setDate(d.getDate() + (Number(days) || 0));
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}
