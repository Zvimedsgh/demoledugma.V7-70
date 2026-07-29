# modReports — Insurance Reporting System Documentation

**Version:** 7.70
**Date:** 2026-05-10
**Module:** `modReports` (Module1 in VBE)

---

## Overview

`modReports` is an Excel/VBA module that processes insurance policy data exported from an agency-management system, detects data-quality exceptions, applies corrections, and generates comparative reports plus a PowerPoint presentation.

The system runs entirely inside one macro-enabled Excel workbook (`.xlsm`). Source data can live in external `YYYY.xlsx` files in a configured folder, or be embedded inside the workbook as hidden `__src_<year>` sheets.

---

## Architecture

### Three public macros (the buttons on the home sheet)

| Button | Sub | Purpose |
|--------|-----|---------|
| **1** | `BuildReview` | Scans the source year, produces a "letipul" review sheet listing all exceptions found |
| **2** | `ApplyCorrectionsAndBuildReports` | Reads decisions from the review sheet, builds `base_<year>` sheets, generates 6 comparison sheets + summary |
| **3** | `BuildPresentation` | Generates a PowerPoint deck with title slide, total summary, 5 chart types × 5 categories, and data tables |

### Supporting public subs

- `SetupMainSheet` — one-time setup of the home sheet (buttons, dropdowns, named ranges, conditional formatting)
- `SetupSettingsSheet` — adds navigation buttons to the settings sheet
- `EmbedSourceSheets` — copies external `YYYY.xlsx` files into the workbook as hidden `__src_<year>` sheets (for self-contained portability)
- `SendForReview` — fired by the "Done updating" button on the review sheet; creates an Outlook email with rows marked "transfer for review"
- `ShowHiddenSheets` / `HideWorkSheets` / `ToggleHiddenSheets` — visibility toggles, also toggle the Ribbon
- `HideRibbon` / `ShowRibbon` — Excel Ribbon visibility
- `RebuildHomeButtons` — recreates the 3 client-row buttons on the home sheet
- `SearchClientName` / `DoClientSearch` / `ConfirmClientSelection` / `CancelClientSearch` — customer-name search workflow
- `ResetClientFilter` / `ClearHomePageSelection` — reset home-page filter selections
- `UpdatePeriodDropdown` / `UpdateFilterValueDropdown` — called from `Worksheet_Change` event handlers in the home sheet's class module

---

## Data Flow

```
┌──────────────────┐
│ Source: YYYY.xlsx│  ← Exported from agency management software
└────────┬─────────┘
         │ (optional: EmbedSourceSheets → __src_YYYY internal)
         ▼
   ┌─────────────┐
   │ BuildReview │  ← Button 1
   └──────┬──────┘
          │ produces
          ▼
   letipul_YYYY (review sheet, user marks Fix/Ignore/Transfer for review)
          │
          ▼   (user clicks "Done updating" → SendForReview → Outlook email)
          │
          ▼
┌─────────────────────────┐
│ ApplyCorrections...     │  ← Button 2
└────────┬────────────────┘
         │ produces
         ▼
   base_YYYY + comparison sheets (companies, branches, agents, tellers, months, main-branch) + summary
         │
         ▼
   ┌──────────────────┐
   │ BuildPresentation│  ← Button 3
   └────────┬─────────┘
            │ produces
            ▼
   PowerPoint deck (title + total + 5 charts × 5 categories + data tables + 2 excluded-agent variants)
```

---

## Sheets in the Workbook

| Sheet (Hebrew name) | Internal code reference | Role |
|--------------------|--------------------------|------|
| **דף הבית** | `CONTROL_SHEET_NAME` | Home: parameters + 3 buttons |
| **הגדרות** | `MANAGEMENT_SHEET_NAME` | Settings: parameters table (cols J/K), checked-fields table (cols E-H), helper translations (cols N/O), branch dictionary (cols A/B), message texts (col S), helper values for CF (cols Z) |
| **לטיפול_YYYY** | `REVIEW_SHEET_NAME` + year | Review sheet — exceptions found by Button 1 |
| **בסיס_YYYY** | `base_<year>` (Hebrew) | Cleaned base data produced by Button 2 |
| **חברות / ענפים / סוכנים / טלרים / חודשים / ענף מרכז** | various | Comparison sheets (current year vs. reference year) |
| **סיכום** | `SHEET_SUMMARY` | Summary sheet with row counts |
| **רשימות** | hidden | Unique filter values collected during BuildReview |
| **__src_YYYY** | embedded source | Optional: hidden internal copy of `YYYY.xlsx` |
| **חיפוש** | temp | Created on demand by `SearchClientName`, deleted after selection |

---

## Configuration (Settings Sheet)

All parameters are in `הגדרות` (the Management sheet), starting at **row 40** (rows 1–39 are reserved for navigation buttons).

### Tables

| Columns | Table | Description |
|---------|-------|-------------|
| **A:B** | Branch dictionary | Branch name (source) → main branch (consolidated). Branches not found here trigger `UNKNOWN_BRANCH` |
| **E:H** | Checked fields | E=display name, F=source column letter, G=`CHECK`/`SKIP`, H=key (e.g. `BRANCH_NAME`, `PREMIUM`) |
| **J:K** | Parameters | Key/value parameters (see below) |
| **N:O** | Helper translations | English code (e.g. `MISSING_BRANCH_NAME`) → Hebrew display text |
| **Q:R** | Period dropdown lookups | Internal lists used by the home-page period dropdowns |
| **S** | Message texts | UI message text rows 1–20 |
| **Z** | CF helper values | Z1=`שנתי`, Z2=`בחר/י` — used by conditional formatting formulas |

### Parameters (J:K)

| Key | Type | Default | Purpose |
|-----|------|---------|---------|
| `PREMIUM_THRESHOLD` | numeric | configurable | Premium values exceeding this trigger `PREMIUM_OVER_THRESHOLD` |
| `ERROR_EMAIL` | string | `zvi@gorentech.co.il` | Recipient for the SendForReview email |
| `FILES_FOLDER` | string | empty → fallback `C:\DEMO PROJECT\SOURCE\` | Where external `YYYY.xlsx` lives. Ignored when `__src_YYYY` exists |
| `REPORTS_FOLDER` | string | empty → `ThisWorkbook.Path` | Where PPTX/PDF are saved |
| `Agency_Name` | string | hardcoded fallback | Agency name shown on slide titles |
| `English_Name` | string | `Demo` | Short English name |
| `BACKUP_PATH` | string | `C:\DEMO PROJECT\BACKUPS` | (declared but not currently used in code) |

---

## Home Sheet Cells (`דף הבית`)

| Cell | Named Range | Purpose | Default / Notes |
|------|-------------|---------|-----------------|
| G3 | `rngBaseYear` | Reference year (e.g. 2019) | label F3: `שנת בסיס` |
| G4 | `rngCurrentYear` | Current year (e.g. 2020) | label F4: `שנה נוכחית` |
| G5 | `rngPeriodType` | Period type dropdown | `שנתי` / `חצי שנתי` / `רבעוני` / `חודשי` |
| G6 | `rngPeriodValue` | Period detail (dependent on G5) | Highlighted gold when G5 requires a value |
| G7 | `rngDateType` | Date column to use | `בורדרו` / `תחילת ביטוח` |
| G9 | `rngFilterType` | Report cross-filter type | `בחר/י` (none) or company / teller / agent / branch / main branch |
| G10 | `rngFilterValue` | Filter value (dependent on G9) | Highlighted gold when G9 is set |
| G12 | `rngClientName` | Customer-name filter | Set via the "חפש" button |
| J3:K4 | `rngDOLAR` etc. | Exchange rates (USD, EUR) | Fetched from Bank of Israel API; falls back to cached value |

---

## Worksheet Events (Home Sheet Class Module)

The home sheet's code module (not `modReports`) must contain a `Worksheet_Change` handler that calls back into `modReports`:

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells.Count > 1 Then Exit Sub
    If Target.Address = "$G$5" Then
        Application.EnableEvents = False
        UpdatePeriodDropdown
        Application.EnableEvents = True
    ElseIf Target.Address = "$G$9" Then
        Application.EnableEvents = False
        UpdateFilterValueDropdown
        Application.EnableEvents = True
    End If
End Sub
```

This wires G5 / G9 changes to the corresponding cascade-dropdown updaters in `modReports`.

---

## Exception Types Detected by `BuildReview`

| Reason code | Trigger | Default Hebrew text |
|-------------|---------|---------------------|
| `MISSING_<key>` | A `CHECK` field in the row is blank (one per missing field) | "חסר X" |
| `PREMIUM_OVER_THRESHOLD` | `\|premium × currency-rate\|` > `PREMIUM_THRESHOLD` | "פרמיה חריגה" |
| `PREMIUM_NOT_NUMERIC` | Premium cell is non-empty and non-numeric | "ערך פרמיה לא מספרי" |
| `UNKNOWN_BRANCH` | Branch name not found in `מילון_ענפים` (col A) | "ענף לא נמצא" |

Each exception is written as one row in `לטיפול_<year>` with the source row number in column A and the offending field values in subsequent columns.

---

## Review Sheet Workflow

In `לטיפול_<year>`, the user fills the **"פעולה"** column for each row:

| Action | Effect during Button 2 |
|--------|------------------------|
| **תקן** + fix value | The fix is applied to the row in `base_<year>` (NOT to source `__src_<year>`) |
| **התעלם** | Row is excluded from `base_<year>` |
| **העבר לבדיקה** | Row is excluded from `base_<year>`; also queued for SendForReview email |

Clicking **"סיימתי לעדכן"** (button at Q1 of the review sheet):
1. Counts rows marked "העבר לבדיקה"
2. Creates a temp Excel file with those rows
3. Opens Outlook with a pre-filled email (to `ERROR_EMAIL`, subject and body in Hebrew, attachment included)
4. Shows a Yes/No prompt: "נוצר מייל עם X שורות. האם שלחת?" (positioned top-left so it doesn't cover Outlook)
5. If No → second prompt: "שלח את המייל מחלון Outlook ולחץ אישור לאחר השליחה"
6. Final prompt: "טיפלת בכל החריגים? לעבור להמשך הפקה?" — Yes returns to the home sheet

---

## Currency Handling

The premium check converts USD to ILS using the exchange rate fetched once per run from the Bank of Israel public API (`https://boi.org.il/PublicApi/GetExchangeRates`). If the API is unreachable (5-second timeout), the cached value in `rngDOLAR` on the home sheet is used.

Detection: source column 27 (`RAW_CURRENCY`). `1` = USD, `0`/`90` = ILS, anything else assumed ILS.

---

## Comparison & Summary Sheets

Button 2 produces 7 sheets:

| Sheet | Group By | Tab Color |
|-------|----------|-----------|
| `חברות` | Insurance company | pastel blue |
| `ענפים` | Branch | pastel peach |
| `ענף מרכז` | Main branch (mapped via dictionary) | pastel pink |
| `טלרים` | Underwriter teller | pastel lavender |
| `סוכנים` | Agent | pastel green |
| `חודשים` | Month (1–12) | pastel yellow |
| `סיכום` | (summary, see below) | pastel orange |

Each comparison sheet has 16 columns: `שם` + 5 metric blocks of `[ref year, current year, change %]` for: premiums, documents, insured, policies, commissions. The totals row uses **bold gold** (`RGB(255, 217, 102)`).

Summary sheet rows:
- Period descriptor
- Row counts (ref year, current year, review count, corrections applied, "rows that didn't need handling", "rows transferred for review")

---

## Presentation Output

Each comparison sheet (months/companies/main-branch/tellers/agents) gets **5 chart slides** + **1 data-table slide**:

1. Premiums (yellow/blue)
2. Documents (purple/teal)
3. Policies (red/dark green)
4. Insured (cyan/dark blue)
5. Commissions (orange/green)
6. Data table (16 cols, RTL, gold totals row)

Plus:
- Title slide (gold background, agency name large)
- Total summary chart (sum of premiums + commissions across all months)
- 2 variant slides: "premiums / commissions excluding `Sokhen 1`"

All slide titles use a **3-line layout**: agency name (top, blue), report type (middle, large), parameter subtitle (bottom, small gray).

After build, PowerPoint stays in the foreground (via `SetForegroundWindow` API). PPTX is NOT auto-saved — user saves manually via File→Save in PowerPoint. PDF export is mentioned in the closing message as "in the production version this will also be produced as PDF".

---

## Hidden Sheets & Visibility Toggle

After Button 2 completes, all internal sheets (`הגדרות`, `לטיפול_*`, `בסיס_*`, `רשימות`, `__src_*`) are hidden as `xlSheetVeryHidden`. Only the home sheet + result sheets remain visible.

The **"הצג/הסתר גיליונות"** button on the home sheet calls `ToggleHiddenSheets`. Showing requires the password `961814` (stored in `Public Const SHEET_PROTECT_PWD`). Hiding does not require a password.

Both actions also toggle the Excel Ribbon (hidden when sheets are hidden, shown when sheets are shown).

---

## Win32 API Usage

The module uses a few Windows API calls (declared in the header):
- `MessageBoxW` — Unicode message boxes (Hebrew rendering)
- `SetWindowsHookEx` / `CallNextHookEx` / `UnhookWindowsHookEx` — CBT hook for positioning the next MsgBox to a specific (X, Y) location
- `SetWindowPos` — used inside the CBT callback
- `SetForegroundWindow` / `ShowWindow` — bring PowerPoint to front after `BuildPresentation`
- `GetCurrentThreadId` — for the CBT hook

All API declarations are conditional (`#If VBA7 Then ... #Else ... #End If`) for 32-bit/64-bit compatibility.

---

## Version Highlights

- **v7.34** — Refactored to Variant-array-based base build (~10× perf on large files)
- **v7.37** — Embedded `__src_<year>` workflow for self-contained workbooks
- **v7.42** — Presentation gained policies column + 4-chart-per-sheet layout
- **v7.47** — Module renamed `modLevav` → `modReports`; all agency-specific names scrubbed for external review
- **v7.50** — Search sheet always recreated; "Reset parameters" button widened
- **v7.53** — 3-line slide titles; wider % columns; gold totals row
- **v7.54** — PowerPoint stays in foreground after success message
- **v7.57** — CBT hook for MsgBox positioning; Ribbon hide; insured chart added
- **v7.60** — Summary sheet splits "transfer for review" from "ignore"; gold totals
- **v7.61** — Conditional formatting via `הגדרות!Z1`/`Z2` cell references
- **v7.62** — Removed home-sheet protection (conflicted with dropdown validation after session reopen)
- **v7.63** — Green background extended to row 31
- **v7.64** — `ShowHiddenSheets`/`HideWorkSheets` also toggle Ribbon
- **v7.65** — Forced screen refresh when search sheet first appears
- **v7.66** — Filtered `בחר/י` default from slide subtitle
- **v7.67** — PDF action dropped (was crashing Excel); replaced with informational message
- **v7.68** — Review sheet now RTL
- **v7.69** — "Done updating" button anchored to cell Q1
- **v7.70** — Fixed `rngCurrentYear` / `rngBaseYear` mapping (was inverted vs. home-page labels)

---

## Security

The VBA project is locked with password `961814` (in production, change this and keep the password elsewhere). Note that VBA-project passwords are easily bypassed by freely available tools; they only deter casual viewing.

The sheet-visibility password (also `961814`) is stored in plaintext as `Public Const SHEET_PROTECT_PWD`. Change it for production.

---

## Known Limitations

- Network call to Bank of Israel API is synchronous; up to ~5-second delay if offline before fallback kicks in
- The CBT hook for MsgBox positioning is a session-local Windows hook; if Excel crashes mid-hook it could theoretically leak (auto-cleaned on session end)
- PPTX is not auto-saved — user must save manually via PowerPoint
- Reference-year base (`בסיס_<refYear>`) is preserved across runs of Button 1; only rebuilt by Button 2 when missing
- All Hebrew text is constructed via `ChrW()` codes (not literals) because the source file is saved as Windows-1252 and Excel's literal-Hebrew handling in VBA strings is unreliable

---

*End of documentation*