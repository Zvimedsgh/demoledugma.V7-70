# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repository is

This is a **VBA project hosted inside Excel macro-enabled workbooks (.xlsm)**. There is no build system, no package manager, no automated tests. Source of truth for code is the standalone `.bas` files at the repo root and under `DEMO/`; those are imported into the workbook via the VBA IDE.

All UI text, sheet names, error messages, and dialog content are in **Hebrew**, embedded as `ChrW(...)` sequences so the `.bas` file stays ASCII-safe. Treat `ChrW` chains as untouchable string literals unless you understand the exact Hebrew word being assembled.

## Files & artifacts

- `DEMO/modDemoReports_V7.70.bas` — **the entire system** (5440 lines, single VBA `Module1`). This is what you almost always edit.
- `modDemoBuilder.bas` — one-time anonymizer that turns real `SOURCE/YYYY.xlsx` into demo files. Imported, run once, then removed. Not part of the runtime system.
- `DEMO/Demo_Reports_Syatem_V7.70.xlsm` — the live workbook that hosts the module. (Note the misspelling "Syatem" in the filename — match it exactly.)
- `DEMO/2019.xlsx`, `DEMO/2020.xlsx` — anonymized yearly source data the demo workbook reads from.
- `SOURCE/2024.xlsx`, `SOURCE/2025.xlsx` — real yearly data; not used by the demo workbook unless `FILES_FOLDER` is repointed.
- `REPORTS/` — output folder for generated `.pptx` / `.pdf` reports.
- `DEMO/חריגים_לטיפול_<year>.xlsx` — workbooks emitted by `SendForReview` for emailing to underwriters.
- `2019_files/`, `2020_files/` — unrelated web-page assets (saved HTML scrape). Ignore.

Version numbers like `V7.70` are baked into the filename, the `Attribute VB_Name` line, and the header changelog. When iterating, ask the user whether to bump the version or stay on the current one.

## Development workflow

There is no CLI flow. Iteration goes through the VBA IDE inside Excel:

1. Open `DEMO/Demo_Reports_Syatem_V7.70.xlsm` in Excel.
2. `Alt+F11` opens the VBA editor.
3. To replace the module: right-click `Module1` in the project tree → Remove (don't export). Then `File → Import File…` and pick the updated `.bas`. The module re-imports under the name in its `Attribute VB_Name` header (`Module1`).
4. Save the workbook (`Ctrl+S`). Excel will keep the `.xlsm` extension.
5. Run macros via `Alt+F8` or via the buttons on the home sheet (`דף הבית`).

Editing the `.bas` outside Excel is fine — it's a plain UTF-8/ASCII text file. The VBA IDE will reformat indentation on import, so don't rely on whitespace being preserved verbatim.

### Entry-point macros (run from Alt+F8 or home-page buttons)

| Macro | Purpose |
|---|---|
| `BuildReview` | Scan the current-year source, write the `לטיפול_<year>` (review) sheet with rows that have missing fields, premium over threshold, or unmapped branch. User picks an action per row (`תקן` / `התעלם` / `העבר לבדיקה`). |
| `ApplyCorrectionsAndBuildReports` | Consume the user's choices, build per-year `בסיס_<year>` sheets, then produce comparison sheets per category (`חברות`, `ענפים`, `ענף מרכז`, `טלרים`, `סוכנים`, `חודשים`) and a `סיכום` sheet. |
| `BuildPresentation` | Export charts from comparison sheets as temp `.gif`s, build a PowerPoint deck via late-binding. |
| `SendForReview` | Bound to the "סיימתי לעדכן" button on each review sheet. Copies the "transfer for review" rows into a temp `.xlsx`, opens an Outlook compose window with the file attached. |
| `SetupMainSheet`, `SetupSettingsSheet`, `RebuildHomeButtons` | One-time / repair utilities for the workbook UI. Run from the Macros dialog if buttons go missing. |
| `EmbedSourceSheets` | Copy external `YYYY.xlsx` files into hidden `__src_YYYY` sheets so the workbook is self-contained. |
| `HideWorkSheets` / `ShowHiddenSheets` / `ToggleHiddenSheets` | Whole-workbook visibility toggles (also hides the ribbon). |

There is no test framework. "Testing" means running the macros against the demo data and checking the output sheets and the generated `.pptx`.

## Architecture

### The pipeline

```
SOURCE (YYYY.xlsx or __src_YYYY sheet)
        │
        ▼  BuildReview            ─► לטיפול_<year>   (issues with actions: תקן/התעלם/העבר לבדיקה)
        │
        ▼  ApplyCorrectionsAndBuildReports
        │      │
        │      ├─► בסיס_<year>     (22-col normalized base, currency-converted, threshold-filtered, fix-applied)
        │      ├─► חברות / ענפים / ענף מרכז / טלרים / סוכנים / חודשים   (comparison vs. base-year)
        │      └─► סיכום           (counts + totals)
        │
        ▼  BuildPresentation       ─► REPORTS/מצגת הנהלה <year>.pptx
        │
        ▼  SendForReview           ─► Outlook compose w/ חריגים_<sheet>.xlsx attached
```

### Sheet name resolution (critical)

User-visible sheet names are Hebrew and would corrupt if hardcoded as literals in an ASCII `.bas`. They are built via private helper functions near the top of the module:

- `CONTROL_SHEET_NAME()` → `"דף הבית"` (home/control sheet)
- `MANAGEMENT_SHEET_NAME()` → `"הגדרות"` (settings/parameters)
- `REVIEW_SHEET_NAME()` → `"לטיפול"` (per-year review prefix; actual sheets are `לטיפול_<year>`)
- `SHEET_COMPANIES/BRANCH/MAINBRANCH/TELLERS/AGENTS/MONTHS/SUMMARY()`
- Base sheets: `"בסיס_" & yearVal` (built inline; legacy English `"base_<year>"` is auto-renamed on encounter)
- Internal embedded source: `"__src_" & yearVal` (very-hidden)
- Helper lists sheet: `"רשימות"` (very-hidden; populated by `BuildReview`)

**Always reach for these helpers, never inline a Hebrew literal.** If you need a new Hebrew label, add a new helper function.

### Configuration model

There are **two** parallel configuration layers:

1. **Named ranges on the control sheet** (`rngCurrentYear`, `rngBaseYear`, `rngPeriodType`, `rngPeriodValue`, `rngFilterType`, `rngFilterValue`, `rngClientName`, `rngDateType`, `rngDOLAR`, `rngFILES_FOLDER`). Read directly via `wsMain.Range("rngXxx").Value2`. Reads are usually wrapped in `On Error Resume Next` and fall back to `""`, so a missing named range silently disables the corresponding feature.

2. **Tables on the `הגדרות` (management) sheet**:
   - Columns E–H (5–8): checked-field definitions — Hebrew name, Excel-letter column on source, "CHECK" flag, machine key. Loaded by `LoadCheckedFields`.
   - Columns J–K (10–11): string/numeric parameters (`PREMIUM_THRESHOLD`, `ERROR_EMAIL`, `FILES_FOLDER`, `REPORTS_FOLDER`, `Agency_Name`, `English_Name`). Looked up via `GetMgmtParam` / `GetStringParameter` / `GetNumericParameter`.
   - Columns N–O (14–15): "helper" dictionary — Hebrew translations for codes like `MISSING_<FIELD>`, `PREMIUM_OVER_THRESHOLD`, plus `REVIEW_*_HEADER` keys used to find columns by header text. Loaded by `LoadHelperDictionary`; defaults are inlined as fallbacks.
   - Column A–B: branch → main-branch mapping. Loaded by `LoadBranchMapping`.
   - Column S (19): UI message strings (confirmation, errors, "done" messages). Indexed by `MSG_ROW_*` constants. Row S20 stores the "don't show again" flag for the pre-run dialog.

### Source data resolution

`OpenSourceFor(yearVal, ByRef wbExternal)` tries:
1. Internal hidden sheet `__src_<year>` in the current workbook.
2. External `<FILES_FOLDER>/<year>.xlsx`, then `.xls`.

The caller must close `wbExternal` only if it's non-Nothing (internal-sheet path leaves it Nothing). Source data sheet inside an external workbook is `TmpClientPolicyListEx` if present, else `Worksheets(1)`.

### Base sheet schema

The `בסיס_<year>` sheet has 22 fixed columns defined by `BASE_COL_*` constants (ID, Year, Month, Identity, Customer, CustName, Policy, Addendum, Company, CompNum, BranchName, BranchNum, MainBranch, AgentName, AgentNum, Teller, TellerNum, Action, Premium, Commission, Issue, ToFix). Source columns are defined by `RAW_*` constants. Any new field must be added in both places.

### Bulk I/O pattern

`BuildBaseSheet` (line ~1446) is the reference implementation: read whole source via `Range.Value2` into a `Variant` 2D array, write whole output in one assignment. Other hot loops (`BuildReview`, `BuildComparisonSheet`, `SendForReview`) still do cell-by-cell I/O and are correspondingly slow — match the bulk pattern when adding new work in those subs.

### Reason-based row fixes

User fixes in the review sheet are stored as Hebrew "reason → fix value" pairs keyed by source row. `ApplyRowFixes` (line ~1588) routes each fix to a base column by Hebrew substring match (`סוכן` → AgentName, `טלר` → Teller, `חברה` → Company, `ענף` → BranchName, `פרמיה` → Premium, `עמלת` → Commission). If you add a new fix-able field, both the review-sheet column and the substring match must be added.

### Error handling pattern

Each public sub follows: save `Application.ScreenUpdating / DisplayAlerts / EnableEvents / Calculation` to `prev*`, set them off, run, restore on `CLEANUP`, restore-or-force-default on `ERR_HANDLER`. Inside, line numbers (`10`, `20`, …) are sprinkled so `Erl` in the error handler reports useful line numbers. Coverage is partial — only "significant" lines are numbered. `debugStep` is a string variable updated at each phase boundary in `ApplyCorrectionsAndBuildReports` for higher-level diagnostics.

### Hebrew Unicode message boxes

`MsgBoxU` wraps `MessageBoxW` (Win32) instead of VBA's `MsgBox`, because the built-in dialog mangles Hebrew. It applies `MB_RTLREADING | MB_RIGHT | MB_SYSTEMMODAL` automatically. Always use `MsgBoxU` for any user-facing message, never `MsgBox`. `PositionNextMsgBox(x, y)` installs a one-shot CBT hook to move the next `MsgBoxU` window — used to keep dialogs from covering Outlook/PowerPoint.

### Sheet protection password

`SHEET_PROTECT_PWD = "961814"` is declared at line ~155. Every entry-point macro unprotects all sheets at the start. This password is a deterrent only (documented inline) — don't treat it as security.

## Conventions when editing

- **Hebrew text goes through helpers or `ChrW` chains** — never paste literal Hebrew into the `.bas`. If you do, it will be lost when the file is re-saved as ASCII.
- **New sheet names** get a helper function next to the existing `SHEET_*` helpers (lines ~256–289).
- **New configuration values** go on the `הגדרות` sheet (column J–K for params, N–O for translations) — don't hardcode.
- **Don't add a new module.** The single-module layout is intentional for ease of distribution. New code goes into `modDemoReports_V7.70.bas`.
- **When changing version**, update the filename of the `.bas`, the `Attribute VB_Name`/header comment block, and (if applicable) the workbook filename. Add a `CHANGES IN <ver>` block near the top of the file.
- **Don't introduce literal English in user-facing strings** — column S messages, dialog text, button captions are all Hebrew via `ChrW`.
