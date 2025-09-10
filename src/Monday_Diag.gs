/* =======================
   monday column diagnostics
   ======================= */

/** Find the subitems board id by inspecting any existing subitem. */
function MONDAY_diagGetSubitemsBoardId() {
  const q = `
    query($ids:[ID!]!,$cursor:String){
      boards(ids:$ids){
        id
        items_page(limit:200,cursor:$cursor){
          cursor
          items{
            id
            subitems { id board { id } }
          }
        }
      }
    }`;
  let cursor = null;
  while (true) {
    const data = MI_gql_(q, { ids: [String(MON_CFG.BOARD_ID)], cursor });
    const page = data?.boards?.[0]?.items_page || {};
    const items = page.items || [];
    for (const it of items) {
      const subs = it.subitems || [];
      if (subs.length && subs[0].board && subs[0].board.id) {
        const subBoardId = String(subs[0].board.id);
        Logger.log(`Subitems board id: ${subBoardId}`);
        return subBoardId;
      }
    }
    if (!page.cursor) break;
    cursor = page.cursor;
  }
  Logger.log('Subitems board id not found. (Create at least one subitem, then re-run.)');
  return null;
}

/** Dump id/type/title for ALL columns on a given board id. */
function MONDAY_diagDumpBoardColumns(boardId) {
  const cols = MI_getBoardColumns_(String(boardId));
  Logger.log(`\nBoard ${boardId} — ${cols.length} columns  (id | type | title)`);
  cols.forEach(c => Logger.log(`${c.id} | ${c.type} | ${c.title}`));
  return cols;
}

/** Convenience: dump columns for the main board and its subitems board. */
function MONDAY_diagDumpMainAndSubitemColumns() {
  MONDAY_diagDumpBoardColumns(MON_CFG.BOARD_ID);
  const subBoardId = MONDAY_diagGetSubitemsBoardId();
  if (subBoardId) MONDAY_diagDumpBoardColumns(subBoardId);
}

/** Lookup a column id by its title (case-insensitive). */
function MONDAY_diagFindColumnIdByTitle(boardId, title) {
  const cols = MI_getBoardColumns_(String(boardId));
  const t = String(title || '').trim().toLowerCase();
  const hit = cols.find(c => String(c.title || '').trim().toLowerCase() === t);
  if (hit) {
    Logger.log(`Board ${boardId}: "${title}" → id=${hit.id} (type=${hit.type})`);
    return hit.id;
  }
  Logger.log(`Board ${boardId}: column titled "${title}" not found.`);
  return null;
}

/* =======================
   FYI: how to set a Link column
   ======================= */
// Parent (main board) after you know its Link column id:
// MI_changeColumnValue_(parentItemId, PARENT_LINK_COL_ID, { url: "https://...", text: "Open file" });

// For subitems, include it during create_subitem (preferred):
// In MI_createSubitemWithAllValues_ add:
//   if (MON_CFG.SUBITEM_COLS.link && linkUrl) {
//     colVals[MON_CFG.SUBITEM_COLS.link] = { url: String(linkUrl), text: String(linkText || linkUrl) };
//   }
// (You’ll need to know SUBITEM_COLS.link = "<subitems link column id>")

