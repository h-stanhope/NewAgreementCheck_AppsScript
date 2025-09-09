# NewAgreementCheck – Google Apps Script

An automation that analyses performer income data, generates per-performer and per-society workbooks, and raises structured **Queries** and **Sub-Queries** in **monday.com** for society owners to review.

---

## TL;DR (What it does)

1. **Reads data** from a Google Sheets workbook:
   - *Performer Overview*, *Society Issues*, *Non-Home Territory Income*, *Recording Overview*, *Recording Issues*, *Most Played Recordings*.
2. **Exports** filtered workbooks:
   - **Performer Breakdown** (one file per performer UUID)
   - **Society Breakdown** (one file per society)
   - Plus a single **Overview** workbook snapshot.
3. **Raises monday.com items** (optional):
   - Creates/Finds a parent *“Bulk Performer Checks – {Society}”* item on the Royalty Tracking board (group: **Test**).
   - Creates sub-items per Society Issue row, sets relations/owners/status/due dates, and posts a contextual comment including the key recordings to check.
   - Obeys a **Rules table** and **Relationship status** gating logic so only the right cases are raised.

---

## Architecture (High-level)

- **FullExport.gs**  
  Orchestrates the end-to-end “Full Export”: builds overview tabs, writes output files, then runs the Monday phase in batches with triggers.

- **MondayIntergration.gs**  
  Encapsulates Monday API calls (GraphQL), column/id resolution, parent creation, sub-item creation, comment building, Rules evaluation, Relationship gating, and link-back columns:
  - Parent link column: **Analysis Query Link** (`link_mkvmm6z4`)
  - Sub-item link column: **Analysis Sub Query Link** (`link_mkvm6jq5`)

- **SheetReaders/SheetWriters/SocietyFunctions/PerformerFunctions/RecordingFunctions/ExpectedAndFX/QualificationEngineAdapter/ISRCCheks**  
  Domain utilities that prepare the overview tabs and business logic used by the export.

- **FullExport_TestHarness.gs / PilotOneContributor.gs**  
  Helpers to run smaller test slices.

- **Monday_Diag.gs**  
  Diagnostic helpers:
  - Dump board columns/ids
  - Trace which **Priority / Issue Status / Account Priority** rule would apply
  - Resolve contributor priorities and society relationship/priority directly from Monday boards
  - Export a row-by-row audit sheet (*“Monday Push – Diag”*) showing why an issue was or wasn’t raised.

---

## monday.com configuration

**Royalty Tracking board (parents):**
- Board ID: `4762290141`
- Group ID: `group_mkvf1ts9` (Test)
- Relevant columns:
  - People: **Query Owner** (`task_owner`)
  - Status/Dropdown: **Query Status** (`query_status`)
  - Dropdown/Status: **Query Type** (`dropdown`)
  - Date: **Review Date** (`date8`)
  - Link: **Analysis Query Link** (`link_mkvmm6z4`)

**Subitems board:**
- Board ID: `4762290182` (discovered via API)
- Relevant columns:
  - People: **Sub Item Owner** (`person`)
  - Status: **Sub Item Status** (`status`)
  - Connect Boards: **Contributor List** (`contributor_list`)
  - Connect Boards: **Society** (`connect_boards`)
  - Dropdown: **Tracking Sub Query** (`tracking_sub_query`)
  - Date: **Review Date** (`resolution_due_date`)
  - Link: **Analysis Sub Query Link** (`link_mkvm6jq5`)

**Lookup boards:**
- **Contributor List** board: `4740231620`
- **Society List** board: `4739922971`
  - *Society Owner* (people) column title used to assign sub-item owner
  - Society “**Priority**” and “**Relationship**” also read from here

---

## Rules & Gating

1) **Relationship gating**  
Only raise if the society’s **Relationship** ∈ { `Good`, `Escalation In Progress` } for that society item. Any other relationship → **skip**.

2) **Rules table** (Google Sheet tab: **Monday Issues Rules**)  

Columns:
- `Priority (From Society List in Monday)` – e.g. Priority 1..4 or Not Allocated  
- `Status (From Society Issues)` – `Missing` or `Underperforming`  
- `Account Priority (From Salesforce Status)` – Bronze/Silver/Gold/Platinum  
- `Raise?` – `Yes` or `No`

The exporter computes:
- **Society Priority** from the Society List board  
- **Issue Status** from *Society Issues* (“Missing”/“Underperforming”)  
- **Account Priority** from the Salesforce data (sheet header **Account Priority**, column **F**)

A row is raised only if **(Relationship allowed) AND (Rules table says Raise = Yes)**.

All intermediate decisions (priority/status matches and rule key) can be inspected on the *“Monday Push – Diag”* sheet produced by diagnostics.

---

## Data source (current vs. Looker – future)

Today the tool reads from named tabs in the active Google Sheet:

- `Contributor List`
- `Performer Overview`
- `Society Issues`
- `Non-Home Territory Income`
- `Recording Overview`
- `Recording Issues`
- `Most Played Recordings`
- `Monday Issues Rules`

**Planned:** swap these readers to a **Looker** API adapter while keeping a **fallback to Sheets** so the tool still runs without Looker credentials.

Suggested adapter shape:

// src/DataSourceAdapter.gs (future)
const DATA_SOURCE = {
  // "gsheet" | "looker"
  MODE: "gsheet",

  getContributorPriorityMap() { /* if MODE=gsheet read from Salesforce tab; if MODE=looker call endpoints */ },
  getSocietyRelationshipAndPriorityMap() { /* resolve from Monday + Looker metadata if applicable */ },
  getSocietyIssuesRows() { /* return same objects the exporter expects */ },
  getRecordingIssuesFor(uuid, society) { /* same shape used by comment builder */ },
  // ...
};
You’ll pass DATA_SOURCE into the existing readers (or proxy calls inside the current reader functions). When Looker creds are present, set MODE="looker"; otherwise default to "gsheet".

Setup
Script properties / secrets
In Apps Script: Project Settings → Script properties
Set:

MONDAY_TOKEN – monday.com API token (Workspace scope)

(Optional) DEFAULT_QUERY_OWNER_ID if you prefer ID over resolving “Harry Stanhope” by name

Scopes (manifest)
Ensure appsscript.json includes the Sheets/Drive/UrlFetch scopes. Example:

{
  "timeZone": "Europe/London",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "oauthScopes": [
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
  ]
}
How to run
Full export (end-to-end + Monday phase)
Open the main Google Sheet.

Run EXPORT_startFullAnalysis (menu or Apps Script).

Confirms Monday raising (Yes/No).

Creates folder tree: Income Tracking/<year>/Performer Reviews/<date>/<time>/...

Writes Performer Breakdown, Society Breakdown, Overview.

Runs Monday phase in chunks (processMondayBatch_) after society files are created, so link columns can point at the right files.

To stop at any time: EXPORT_abortFullAnalysis.

Test small batches
MONDAY_pushSocietyIssuesBatch() – push first N rows (per TEST_BATCH_ROWS) without running the whole export.

MONDAY_pushFirstSocietyIssueAsSubitem() – single row smoke test.

MONDAY_dumpBoardColumns() – prints board/column IDs for parents and subitems.

MONDAY_buildDiagSheetForRange() (from Monday_Diag.gs) – builds Monday Push – Diag with every key the rules engine evaluates and the reason for raise/skip.

Link-back columns
Parent item gets: Analysis Query Link → the generated Society workbook URL.

Each sub-item gets: Analysis Sub Query Link → the generated Performer workbook URL.

Both links are set at creation time so society owners can jump straight to the evidence.

Key entrypoints (by file)
FullExport.gs

EXPORT_startFullAnalysis

EXPORT_continueFullAnalysis_

EXPORT_abortFullAnalysis

EXPORT_pushMondayForUUID_ (per-performer push with de-dup)

MondayIntergration.gs

MI_findOrCreateParentItem_

MI_createSubitemWithAllValues_

MI_buildCommentForRow_ (formats “Title (Version) – Artist : N plays”)

Rules: reads Monday Issues Rules tab and Society Relationship gate

Monday_Diag.gs

MONDAY_dumpBoardColumns

MONDAY_buildDiagSheetForRange

Handoff prompt (for a new collaborator)
Project summary:
This Apps Script exports performer/society analysis from Google Sheets into Drive folders, then raises Monday.com Queries/Sub-Queries with rich context (owners set, statuses, due dates, and a comment listing the top recordings to check). A Rules table and a Relationship gate control which rows become Monday tasks.

Where to start:
Read README.md, then open src/FullExport.gs and src/MondayIntergration.gs. Run MONDAY_dumpBoardColumns to confirm column IDs match the workspace. Use MONDAY_buildDiagSheetForRange to validate rule gating on sample rows.

Next milestone:
Add a Looker data adapter that returns the same row shapes used today (see the “Data source” section). Keep Sheets as fallback behind a simple DATA_SOURCE.MODE flag.

Contributing
Use feature branches.

Keep Monday board IDs, group IDs, and column IDs in config/README.md for fast updates.

If you change a sheet/tab name, update the constants in MondayIntergration.gs and FullExport.gs.

Troubleshooting
Everything is “SKIP by rules”
Run MONDAY_buildDiagSheetForRange and inspect:

Society Relationship (must be Good or Escalation In Progress)

Society Priority (from Monday) + Issue Status (from Society Issues) + Account Priority (from Salesforce Account Priority column F)

Rule Exists / Rule Says Raise

Parent fields not setting
Verify Query Owner, Query Status, Query Type, Review Date column IDs with MONDAY_dumpBoardColumns.

Links not set
Ensure the export stage ran before Monday stage so file URLs exist in state.

---

## Final tips

- Commit **exactly** the `.gs` files you showed into `src/`. Keep filenames as-is so anyone can find the entrypoints quickly.
- Add your latest board/column diagnostics as `docs/monday_columns.md`. It’s gold for future debugging.
- When you wire Looker, create `src/DataSourceAdapter.gs` and keep the public signatures identical to what the readers currently return—this makes the swap safe and keeps Monday logic untouched.

If you want, I can also generate stub files (`DataSourceAdapter.gs`, `docs/monday_columns.md`, and `config/README.md`) with ready-to-fill sections—just say the word.
::contentReference[oaicite:0]{index=0}
