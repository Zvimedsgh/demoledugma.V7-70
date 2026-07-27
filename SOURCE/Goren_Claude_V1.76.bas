Attribute VB_Name = "Goren_Claude1_76"
' ============================================================================
' MODULE: modLevav
' PURPOSE: Complete system - BuildReview + ApplyCorrectionsAndBuildReports
' VERSION: V1.76
' - UI: Moved Exit System button to A21 + 60 (shifted left) to avoid Copilot icon.
' - UI: Verified Credit text is in row 21 and centered.
' DATE: 2026-06-27 09:18:00
' ============================================================================
' CHANGES IN 1.51:
'   - BUGFIX: Fixed SetupSettingsMenu floating buttons overlapping horizontally.
' CHANGES IN 1.50:
'   - STABLE: Finalizing version 1.50 with robust Error 1004 protections and syntax fixes.
' CHANGES IN 1.49:
'   - BUGFIX: Fixed syntax error in SaveReportsToFolder EnableEvents.
' CHANGES IN 1.48:
'   - BUGFIX: Disabled Application.EnableEvents in SaveReportsToFolder to prevent Sheet14 Worksheet_Activate crash.
' CHANGES IN 1.47:
'   - BUGFIX: Replaced all Range.Select calls with Application.Goto to prevent error 1004.
' CHANGES IN 1.46:
'   - UI: Pushed credit text 25 spaces from the right inside column A.
' CHANGES IN 1.45:
'   - UI: Changed credit text alignment to right with an IndentLevel of 2 (~25 points from right).
' CHANGES IN 1.44:
'   - UI: Adjusted column A width to 43. Moved credit to A19 (centered) and Show/Hide button to L19.
' CHANGES IN 1.43:
'   - UI: Changed J3:J4 background color to match F3 (RGB 200, 230, 255).
' CHANGES IN 1.42:
'   - UI: Fixed typo in currency message (Nun Sofit).
'   - UI: Adjusted column A width to 60 for better centering. Expanded green area to AZ40.
' CHANGES IN 1.41:
'   - UI: Adjusted currency message to fit without newlines and increased font size to 11.
'   - UI: Tweaked column padding (A=24, L=45).
' CHANGES IN 1.39:
'   - UI: Fixed J5:K7 borders showing gridlines. Merging is now done AFTER writing value to avoid Error 91.
'   - UI: Made J5:K7 area light blue to differentiate from currency table.
'   - UI: Made Credit text in A19 bold, size 12, and black for better visibility.
' CHANGES IN 1.38:
'   - UI: Fixed merged cells issue in currency box (changed to CenterAcrossSelection).
'   - UI: Centered the main layout with equal padding on columns A and L.
'   - UI: Moved toggle sheets button from M21 to J19. Moved credit text from A21 to A19 and made it darker.
' CHANGES IN 1.37:
'   - Fixed: SetupMainSheet crashing on first run (Error 91) due to missing rngSection_PeriodLists and rngSection_Messages Named Ranges. SetupSectionNamedRanges is now called at the beginning of SetupMainSheet.
'   - Updated: Version numbering, VB_Name, and timestamp headers in code.
' CHANGES IN 1.36:
'   - Fixed: SetupSettingsMenu now freezes top 4 rows and places navigation buttons horizontally as a floating toolbar.
'   - Fixed: SetupSectionNamedRanges dynamically offsets ranges based on toolbar presence.
'   - Fixed: Renamed ApplyUserPermissions back to CheckUserPermissions to fix Workbook_Open macro error.
'   - Updated: Version numbering, VB_Name, and timestamp headers in code.
' CHANGES IN 9.28 (Modularization):
'   - Extracted LoadCorrectionsToDicts and BuildBaseSheet to reduce the size of ApplyCorrectionsAndBuildReports.
' CHANGES IN 9.27 (Merged with 7.54 logic):
'   - Added SafeDouble, SafeLong, and SafeString functions to prevent Error 13 Type Mismatch on Variant Array Error values.
'   - Converted ApplyCorrectionsAndBuildReports cell-by-cell loops to use in-memory Variant Arrays.
'   - Converted BuildComparisonSheet to use in-memory Variant Arrays (curData, refDataComp) for massive speedup.
' CHANGES IN 9.2:
'   - Fixed: Runtime error 1004 when rngSection_Messages not yet created
'   - Fixed: BuildReview now auto-runs SetupSectionNamedRanges if Named Ranges missing
'   - Fixed: All Dim declarations moved to top of BuildReview (before executable code)
'   - Fixed: Duplicate VBA line number 150 in LoadHelperDictionary
' CHANGES IN 9.10:
'   - New: SetupSectionNamedRanges sub - defines all rngSection_* Named Ranges for vertical layout
'   - New: EOD_MARKER constant - each section ends with "EOD" in column A for flexible row counts
'   - New: Navigation buttons in column I (SetupSettingsMenu) - scroll to sections instead of hiding columns
'   - New: SettingsScrollTo helper + NavSettings_* subs for vertical navigation
'   - Fixed: rngFILES_FOLDER now points to B176 (was K4) in hagdarot Params section
'   - Fixed: rngREPORTS_FOLDER now points to B177 (was K5) in hagdarot Params section
'   - Changed: All data reads (fields/params/reasons) use Named Range start + EOD scan
'   - Changed: System messages moved from column S to rngSection_Messages area (A222+)
'   - Changed: Period lists moved from columns Q/R to rngSection_PeriodLists area (A196+)
'   - Changed: UpdateClientList uses rngSection_Clients Named Range dynamically
'   - Removed: SETTINGS_COL_* constants, MANAGEMENT_START_ROW, SECTION_CLIENTS_ROW
'   - Removed: Old column-based SettingsShowOnly approach
' CHANGES IN 9.09:
'   - Fixed: rngDOLAR now points to K3 (rate value cell) instead of K4
'   - New: Home page layout: J3=$, K3=USD rate, J4=â‚¬, K4=EUR rate, J5:K7=info message
'   - New: FetchBOIMonthlyAvg(currencyKey, monthStr) - fetches monthly average rate for any currency
'   - New: GetCurrencyRate() - unified function supporting USD/EUR with matach lookup
'   - New: BuildMatachCache() - pre-fetches all needed rates (USD+EUR) by bordero month
'   - New: Matach sheet auto-created if missing (Col A=YYYY-MM, B=USD, C=EUR)
'   - Fixed: Initial GetDollarRate call uses skipBOI:=False (fetches once from API)
'   - Fixed: BuildReview no longer deletes matach sheet (added to exclusion list)
'   - Fixed: SetupMainSheet creates rngFILES_FOLDER (K4) and rngREPORTS_FOLDER (K5) on hagdarot
'   - Fixed: Border area extended to J3:K7 to include exchange rate message
' CHANGES IN 9.02:
'   - Fixed: Error handling rewrite in SetupMainSheet second half (proper ERR_HANDLER transitions)
'   - New: Settings navigation menu (Hide/Unhide) with colored buttons
'   - New: Client list moved to settings sheet column U, auto-updated in BuildReview
'   - New: SetAllTabColors - pastel tabs for ALL sheets including hidden
'   - New: Dollar rate displayed in K4 with BOI auto-fetch on setup
'   - Changed: Date moved to S1, version to A23, credit to R23 (bold dark blue)
'   - Changed: NewClients buttons widened to 160x28 with font size 12
'   - Fixed: Matach sheet hidden (xlSheetVeryHidden) with brown tab color
' CHANGES IN 8.06:
'   - Fixed: PrintNewClients removed invalid PageSetup.RightToLeft (Error 438)
'   - New: GetDollarRateByMonth reads rate from sheet "matach" by bordero month (YYYY-MM)
'   - New: Auto-update current month rate from BOI API if missing in matach sheet
'   - New: Date displayed in R1, version in A18, credit in R18 on home page
'   - New: Private Const APP_VERSION for version tracking
' CHANGES IN 8.04:
'   - Fixed: CheckUserPermissions now disables buttons (grey + no action) instead of hiding
'   - Fixed: Added Unprotect for wsMain and wsMgmt in SetupMainSheet
'   - Fixed: Environ(USERNAME) for automatic user identification
' CHANGES IN 8.03:
'   - Button 5 added: "View Reports" - opens Reports folder in Explorer
'   - Button 6 added: "New Clients" - shows clients with premium>0 in current year not in base year
'   - Exit System button moved to B16, Show/Hide moved to P16
'   - Fixed SetupMainSheet: named range refs replaced with direct cell refs (G5/G6/G7)
'   - Error handler now shows line number for easier debugging
' CHANGES IN 7.42:
'   - reportsFolder read from named range rngREPORTS_FOLDER (fixes SharePoint/OneDrive URL path issue)
'   - Button 4 tempPath also uses rngREPORTS_FOLDER instead of ThisWorkbook.Path
' CHANGES IN 7.41:
'   - PPTX/PDF save via TEMP folder then FileCopy to Reports (OneDrive fix)
'   - G15 processing message: merged G15:K15, full text visible, yellow background
'   - G6/G10 color reset to green after dropdown selection (was turning purple)
' CHANGES IN 7.53:
'   - Fixed file truncation issue to restore all original lines from V7.4
'   - Added SafeDouble, SafeLong, and SafeString functions to prevent Error 13 Type Mismatch on Variant Array Error values.
'   - Converted ApplyCorrectionsAndBuildReports cell-by-cell loops to use in-memory Variant Arrays.
' CHANGES IN 7.40:
'   - Border added for J5:K6 on home page
'   - Removed vertical line artifact in column L
'   - Processing message moved to G15 (buttons 2 & 3)
'   - Fixed off-by-one in review row count message (outRow-3 instead of outRow-2)
' CHANGES IN 7.39:
'   - Button "Done Updating" moved to row 1 (above headers); data starts at row 3
'   - Year format fixed: no comma in year display (2025 not 2,025)
'   - Presentation success message simplified (no file paths shown)
'   - Presentations and PDF saved to Reports subfolder
'   - Button 4 added: "Save Reports" - exports result sheets as XLSX to Reports folder
'   - Charts added for Documents and Insured persons in presentation (4 charts per sheet)
'   - Unknown branches: auto-added to settings A:D with BRANCH_KEY transliteration
'   - Branch fix from corrections: updates settings B (main branch) + D (main branch key)
'   - HebrewToKey transliteration function added
' CHANGES IN 7.38:
'   - FIX: "?????? ?????" button narrowed (110x28) and positioned at actionCol.Left (not hardcoded P1)
'     so it stays within the visible frame in RTL layout. Font 11, margins reduced.
' CHANGES IN 7.37:
'   - CRITICAL FIX: rngCurrentYear/rngBaseYear were swapped (G3<->G4) causing anomalies on wrong year
'   - PowerPoint crash prevention: uses GetObject for existing PP instance, graceful Quit
'   - "Siymti le'adken" button moved to column P on review sheet
'   - Column O width reduced in review sheet
'   - Enhanced messages: "X rows handled" when no rows to send; email send confirmation
'   - Totals row: darker background (30% blend) + top border for emphasis
'   - Home sheet: column widths and row heights adjusted for smaller screens
'   - Performance: ScreenUpdating=False during chart export in BuildPresentation
' CHANGES IN 7.36:
'   - Fixed HOME_GREEN constant in Sheet1 code (was 14024924=purple, now 14479580=green)
'   - ResetHomeDefaults verified: uses RGB(220,240,220) correctly for G6 and G10
' CHANGES IN 7.33:
'   - Added paramsSubtitle to BuildTotalSlideFromImage (was missing - subtitle only showed on title slide)
' CHANGES IN 7.32:
'   - SetupMainSheet now creates ALL named ranges (rngCurrentYear, rngBaseYear, rngPeriodType,
'     rngPeriodValue, rngDateType, rngFilterType, rngFilterValue) pointing to correct G column cells
'   - Fixes paramsSubtitle being empty in presentation (rngPeriodValue was pointing to old C column)
' CHANGES IN 7.31:
'   - After button 2 finishes: ensure home sheet visible + activate it (fixes hidden home sheet)
' CHANGES IN 7.30:
'   - Exclude MainBranch="????" from ALL reports, summaries and presentation
'   - SetupMainSheet creates home sheet automatically if not found (no more error 9001)
'   - Column A minimum width 20 in result sheets (ensures full names visible)
' CHANGES IN 7.29:
'   - Fixed 'Select method of Range class failed' in SetupMainSheet (added wsMain.Activate before Select)
' CHANGES IN 7.28:
'   - Zebra striping contrast increased: light 92%/dark 55% (was 85%/65%)
'   - Column A only AutoFit in result sheets
'   - Thin light-gray borders on all data cells in result sheets
'   - Totals row: medium tint background (50% blend) with BLACK font (visible on all colors)
'   - RTL enforced on each result sheet
'   - Summary sheet (sikum) RTL added
'   - Border shape added for J3:K4 (exchange rate cells) on home page
' CHANGES IN 7.23:
'   - Row 2 headers re-applied AFTER green pastel background (was being overwritten)
' CHANGES IN 7.22:
'   - Row 2 headers (F2,G2,J2,K2) background changed to dark green (0,100,0)
' CHANGES IN 7.21:
'   - Exchange rate msg: UnMerge J5:L7 first, then Merge J5:K7 only, font 10
' CHANGES IN 7.20:
'   - Exchange rate message rewritten: 3 centered lines, singular ("hasha'ar mit'adken")
'   - paramsSubtitle: removed duplicate year comparison (already shown in chart)
'   - Chart data labels font enlarged (total:10, comparison:9)
'   - Fixed duplicate line number 220 in BuildTitleSlide
'   - Long line (1095 chars) split to fix VBA syntax error
' CHANGES IN 7.17:
'   - Presentation slides: paramsSubtitle added to ALL slides (title, chart, table)
'     Shows: periodDesc | dateType | detailBy | clientName
'   - Borders added to each cell in F3:G12 on home page (dark blue, thin)
'   - B12 processing message now restores green pastel background after clearing
'   - Chart data labels rotated to 90 degrees (vertical/upward) to prevent overlap
'   - MsgBoxU now uses MB_SYSTEMMODAL so messages appear above PowerPoint window
' CHANGES IN 7.16:
'   - Border for exchange rate message: Shape rectangle (fixes merged cell border issue)
'   - Green pastel background for home page A1:U24
'   - Zebra striping on result sheets (companies, branch, mainbranch, tellers, agents, months)
'     based on each sheet's tab color (light/dark alternating rows)
' CHANGES IN 7.15:
'   - Exchange rate message moved to J5:K7 (under rates)
'   - Show/Hide button moved to R18
'   - Font size 14 set for ALL sheets
'   - Added "kulam" (all) button to reset client filter (G12 -> bachar/i)
' CHANGES IN 7.14:
'   - Show/Hide button moved to R20
'   - RTL set for ALL sheets in workbook
'   - Exchange rate info message added at K5:L7 with border
'   - Cursor goes to A1 at end of SetupMainSheet
' CHANGES IN 7.13:
'   - Buttons repositioned to Left=849.5 (near col M) with 50px spacing
'   - F12 (shem lakoach) font fixed: size 12, bold, blue (0,70,140)
' CHANGES IN 7.12:
'   - ToggleHiddenSheets: showing hidden sheets now requires password
' ============================================================================

' --- Windows API for Unicode MsgBox ---
#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

' --- General constants ---

Private Const APP_VERSION As String = "1.76"
Private Const DATA_SHEET_NAME As String = "TmpClientPolicyListEx"

' --- New Clients sheet names: NC_TEMP_SHEET_NAME() and NC_MASTER_SHEET_NAME() functions are below ---

' --- NIHUL field definition table (columns A-D, starts at rngSection_FieldMap+1, ends at EOD) ---
Private Const COL_FIELD_NAME_HE As Long = 1   ' Column A = field name in Hebrew
Private Const COL_FIELD_COLUMN As Long = 2    ' Column B = source column letter
Private Const COL_FIELD_CHECKING As Long = 3  ' Column C = checking flag
Private Const COL_FIELD_KEY As Long = 4       ' Column D = field key

' --- NIHUL parameter table (columns A-B, starts at rngSection_Params+1, ends at EOD) ---
Private Const COL_PARAM_NAME As Long = 1      ' Column A = parameter name
Private Const COL_PARAM_VALUE As Long = 2     ' Column B = parameter value

' --- NIHUL reason/helper table (columns A-B, starts at rngSection_ReasonCode+1, ends at EOD) ---
Private Const COL_HELPER_KEY As Long = 1      ' Column A = key
Private Const COL_HELPER_VALUE As Long = 2    ' Column B = value

' --- End-of-data marker: each section ends with "EOD" in column A ---
Private Const EOD_MARKER As String = "EOD"

Private Const PARAM_PREMIUM_THRESHOLD As String = "PREMIUM_THRESHOLD"
Private Const PARAM_ERROR_EMAIL As String = "ERROR_EMAIL"
Private Const KEY_BRANCH_NAME As String = "BRANCH_NAME"
Private Const KEY_PREMIUM As String = "PREMIUM"
Private Const HELPER_REVIEW_SOURCE_ROW_HEADER As String = "REVIEW_SOURCE_ROW_HEADER"
Private Const HELPER_REVIEW_REASON_HEADER As String = "REVIEW_REASON_HEADER"
Private Const HELPER_REVIEW_REASON_CODE_HEADER As String = "REVIEW_REASON_CODE_HEADER"

' --- Output sheet names (Hebrew via functions below) ---

' --- Raw source column mapping ---
Private Const RAW_CUSTOMER As Long = 1
Private Const RAW_CUSTNAME As Long = 2
Private Const RAW_POLICY As Long = 11
Private Const RAW_ADDENDUM As Long = 12
Private Const RAW_COMPNUM As Long = 13
Private Const RAW_COMPANY As Long = 14
Private Const RAW_BRANCHNUM As Long = 15
Private Const RAW_BRANCHNAME As Long = 16
Private Const RAW_INSURANCE_START As Long = 17
Private Const RAW_BORDEREU As Long = 19
Private Const RAW_AGENTNUM As Long = 20
Private Const RAW_AGENTNAME As Long = 21
Private Const RAW_TELLERNUM As Long = 24
Private Const RAW_TELLERNAME As Long = 25
Private Const RAW_PREMIUM As Long = 28
Private Const RAW_COMMISSION As Long = 32
Private Const RAW_CURRENCY As Long = 27
Private Const RAW_ACTIONCOL As Long = 39
Private Const RAW_IDNUMBER As Long = 45

' --- Base sheet columns ---
Private Const BASE_COL_ID As Long = 1
Private Const BASE_COL_YEAR As Long = 2
Private Const BASE_COL_MONTH As Long = 3
Private Const BASE_COL_IDENTITY As Long = 4
Private Const BASE_COL_CUSTOMER As Long = 5
Private Const BASE_COL_CUSTNAME As Long = 6
Private Const BASE_COL_POLICY As Long = 7
Private Const BASE_COL_ADDENDUM As Long = 8
Private Const BASE_COL_COMPANY As Long = 9
Private Const BASE_COL_COMPNUM As Long = 10
Private Const BASE_COL_BRANCHNAME As Long = 11
Private Const BASE_COL_BRANCHNUM As Long = 12
Private Const BASE_COL_MAINBRANCH As Long = 13
Private Const BASE_COL_AGENTNAME As Long = 14
Private Const BASE_COL_AGENTNUM As Long = 15
Private Const BASE_COL_TELLER As Long = 16
Private Const BASE_COL_TELLERNUM As Long = 17
Private Const BASE_COL_ACTION As Long = 18
Private Const BASE_COL_PREMIUM As Long = 19
Private Const BASE_COL_COMMISSION As Long = 20
Private Const BASE_COL_ISSUE As Long = 21
Private Const BASE_COL_TOFIX As Long = 22

Private Const MB_RTLREADING As Long = &H100000
Private Const MB_RIGHT As Long = &H80000
Private Const MB_SYSTEMMODAL As Long = &H1000
' --- Settings section rows (vertical layout - all in column A) ---
' Named Ranges used: rngSection_FieldMap, rngSection_BranchName, rngSection_Params,
' rngSection_ReasonCode, rngSection_PeriodLists, rngSection_Messages,
' rngSection_Permissions, rngSection_Clients
' These are defined by SetupSectionNamedRanges sub
' SECTION_CLIENTS_ROW removed - now uses rngSection_Clients Named Range dynamically

' ============================================================================
' HEBREW STRING HELPERS
' ============================================================================
Private Function H_AGENT() As String: H_AGENT = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503): End Function
Private Function H_TELLER() As String: H_TELLER = ChrW(1496) & ChrW(1500) & ChrW(1512): End Function
Private Function H_COMPANY() As String: H_COMPANY = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492): End Function
Private Function H_BRANCH() As String: H_BRANCH = ChrW(1506) & ChrW(1504) & ChrW(1507): End Function
Private Function H_PREMIUM() As String: H_PREMIUM = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1492): End Function
Private Function H_COMMISSION() As String: H_COMMISSION = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1514): End Function
Private Function H_IGNORE() As String: H_IGNORE = ChrW(1492) & ChrW(1514) & ChrW(1506) & ChrW(1500) & ChrW(1501): End Function
Private Function H_REVIEW() As String: H_REVIEW = ChrW(1492) & ChrW(1506) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1492): End Function
Private Function H_FIX() As String: H_FIX = ChrW(1514) & ChrW(1511) & ChrW(1503): End Function
Private Function H_UNKNOWN_BRANCH() As String: H_UNKNOWN_BRANCH = ChrW(1506) & ChrW(1504) & ChrW(1507) & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1502) & ChrW(1494) & ChrW(1493) & ChrW(1492) & ChrW(1492): End Function
Private Function H_BASE() As String: H_BASE = ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505): End Function
Private Function H_RESHIMOT() As String: H_RESHIMOT = ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1493) & ChrW(1514): End Function

' --- CRITICAL:: Sheet name functions (Hebrew via ChrW - must be near top) ---
Private Function CONTROL_SHEET_NAME() As String
    ' daf habait
    CONTROL_SHEET_NAME = ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)
End Function

Private Function MANAGEMENT_SHEET_NAME() As String
    ' hagdarot
    MANAGEMENT_SHEET_NAME = ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514)
End Function

Private Function REVIEW_SHEET_NAME() As String
    ' letipul
    REVIEW_SHEET_NAME = ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500)
End Function

Private Function MATACH_SHEET_NAME() As String
    ' matach (foreign currency sheet)
    MATACH_SHEET_NAME = ChrW(1502) & ChrW(1496) & ChrW(34) & ChrW(1495)
End Function

Private Function SOURCE_FOLDER() As String
    ' Read from named range rngFILES_FOLDER (defined in NIHUL sheet)
    ' Falls back to C:\?????? ???\SOURCE\ if named range not found
    On Error Resume Next
    SOURCE_FOLDER = Trim$(CStr(ThisWorkbook.Names("rngFILES_FOLDER").RefersToRange.Value2))
    On Error GoTo 0
    If SOURCE_FOLDER = "" Then
        SOURCE_FOLDER = "C:\" & ChrW(1508) & ChrW(1512) & ChrW(1493) & ChrW(1497) & ChrW(1511) & ChrW(1496) & " " & ChrW(1500) & ChrW(1489) & ChrW(1489) & "\SOURCE\"
    End If
    ' Ensure trailing backslash
    If Right$(SOURCE_FOLDER, 1) <> "\" Then SOURCE_FOLDER = SOURCE_FOLDER & "\"
End Function

Private Function REPORTS_FOLDER() As String
    ' Read from named range rngREPORTS_FOLDER (defined in NIHUL sheet)
    ' Falls back to ThisWorkbook.Path & "\Reports" if named range not found
    On Error Resume Next
    REPORTS_FOLDER = Trim$(CStr(ThisWorkbook.Names("rngREPORTS_FOLDER").RefersToRange.Value2))
    On Error GoTo 0
    If REPORTS_FOLDER = "" Then
        REPORTS_FOLDER = ThisWorkbook.Path & "\Reports"
    End If
    ' Ensure trailing backslash removed (we add \ when building paths)
    If Right$(REPORTS_FOLDER, 1) = "\" Then REPORTS_FOLDER = Left$(REPORTS_FOLDER, Len(REPORTS_FOLDER) - 1)
End Function

Private Function SHEET_COMPANIES() As String
    ' hevrot
    SHEET_COMPANIES = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1514)
End Function

Private Function SHEET_BRANCH() As String
    ' anafim
    SHEET_BRANCH = ChrW(1506) & ChrW(1504) & ChrW(1508) & ChrW(1497) & ChrW(1501)
End Function

Private Function SHEET_MAINBRANCH() As String
    ' anaf merkaz
    SHEET_MAINBRANCH = H_BRANCH() & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)
End Function

Private Function SHEET_TELLERS() As String
    ' tlerim
    SHEET_TELLERS = H_TELLER() & ChrW(1497) & ChrW(1501)
End Function

Private Function SHEET_AGENTS() As String
    ' sochnim
    SHEET_AGENTS = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501)
End Function

Private Function SHEET_MONTHS() As String
    ' hodshim
    SHEET_MONTHS = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
End Function

Private Function NC_TEMP_SHEET_NAME() As String
    NC_TEMP_SHEET_NAME = ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)  ' ??????_?????
End Function

Private Function NC_MASTER_SHEET_NAME() As String
    NC_MASTER_SHEET_NAME = ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500) & "_" & ChrW(1489) & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)  ' ?????_???????_?????
End Function

Private Function SHEET_SUMMARY() As String
    ' sikum
    SHEET_SUMMARY = ChrW(1505) & ChrW(1497) & ChrW(1499) & ChrW(1493) & ChrW(1501)
End Function

' ============================================================================
' HELPER: Unicode MsgBox wrapper (uses Windows API MessageBoxW)
' ============================================================================
Private Function MsgBoxU(ByVal sText As String, Optional ByVal uType As Long = 0, Optional ByVal sCaption As String = "") As Long
    MsgBoxU = MessageBoxW(0, StrPtr(sText), StrPtr(sCaption), uType Or MB_RTLREADING Or MB_RIGHT Or MB_SYSTEMMODAL)
End Function

' ============================================================================
' HELPER: Find source file - supports both .xlsx and .xls
' ============================================================================
Private Function FindSourceFile(ByVal yearVal As String) As String
10      Dim p As String
20      Dim fso As Object
30      Set fso = CreateObject("Scripting.FileSystemObject")
40      p = SOURCE_FOLDER() & yearVal & ".xlsx"
50      If fso.FileExists(p) Then
60          FindSourceFile = p
70          Exit Function
80      End If
90      p = SOURCE_FOLDER() & yearVal & ".xls"
100     If fso.FileExists(p) Then
110         FindSourceFile = p
120         Exit Function
130     End If
140     FindSourceFile = ""
End Function

' ============================================================================
' HELPER: Safely open workbook or attach if already open
' ============================================================================
Private Function SafeOpenWorkbook(ByVal fPath As String) As Workbook
    Dim fn As String
    fn = Mid$(fPath, InStrRev(fPath, "\") + 1)
    
    Dim pvw As Object
    For Each pvw In Application.ProtectedViewWindows
        If LCase$(pvw.SourceName) = LCase$(fn) Then
            On Error Resume Next
            pvw.Edit
            On Error GoTo 0
            Exit For
        End If
    Next pvw
    
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(fn)
    On Error GoTo 0
    
    If wb Is Nothing Then
        Application.DisplayAlerts = True
        On Error Resume Next
        Workbooks.Open fPath, ReadOnly:=True
        On Error GoTo 0
        Application.DisplayAlerts = False
    End If
    
    For Each pvw In Application.ProtectedViewWindows
        If LCase$(pvw.SourceName) = LCase$(fn) Then
            On Error Resume Next
            pvw.Edit
            On Error GoTo 0
            Exit For
        End If
    Next pvw
    
    On Error Resume Next
    Set wb = Workbooks(fn)
    On Error GoTo 0
    
    If wb Is Nothing Then
        MsgBoxU ChrW(1492) & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509) & " " & fn & " " & ChrW(1495) & ChrW(1505) & ChrW(1493) & ChrW(1501) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1497) & ChrW(1491) & ChrW(1497) & " " & ChrW(1488) & ChrW(1511) & ChrW(1505) & ChrW(1500) & " (" & ChrW(1502) & ChrW(1490) & ChrW(1503) & " / Protected View). " & ChrW(1508) & ChrW(1514) & ChrW(1495) & " " & ChrW(1488) & ChrW(1493) & ChrW(1514) & ChrW(1493) & " " & ChrW(1497) & ChrW(1491) & ChrW(1504) & ChrW(1497) & ChrW(1514) & " " & ChrW(1493) & ChrW(1488) & ChrW(1513) & ChrW(1512) & " Enable Editing.", vbCritical
        Err.Raise vbObjectError + 999, "OPEN_SRC", "Protected View blocked access."
    End If
    
    Set SafeOpenWorkbook = wb
End Function

' ============================================================================
' HELPER: Open data worksheet from source workbook
' ============================================================================
Private Function OpenDataSheet(ByVal wb As Workbook) As Worksheet
10      On Error Resume Next
20      Dim ws As Worksheet
30      Set ws = wb.Worksheets(DATA_SHEET_NAME)
        If Err.Number <> 0 Then
            Err.Clear
            Set ws = wb.Worksheets(1)
            If Err.Number <> 0 Then
                MsgBoxU ChrW(1492) & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509) & " " & wb.Name & " " & ChrW(1495) & ChrW(1505) & ChrW(1493) & ChrW(1501) & " (" & ChrW(1502) & ChrW(1490) & ChrW(1503) & "). " & ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1495) & " " & ChrW(1488) & ChrW(1493) & ChrW(1514) & ChrW(1493) & " " & ChrW(1497) & ChrW(1491) & ChrW(1504) & ChrW(1497) & ChrW(1514) & " " & ChrW(1493) & ChrW(1500) & ChrW(1488) & ChrW(1513) & ChrW(1512) & " Enable Editing.", vbCritical
                Err.Raise vbObjectError + 999, "OPEN_DATA_SHEET", "Protected View block."
            End If
        End If
40      On Error GoTo 0
50      If ws Is Nothing Then
60          Set ws = wb.Worksheets(1)
70      End If
80      Set OpenDataSheet = ws
End Function

' ============================================================================
' HELPER: Get month range for comparison period from B4+C4
' Returns minMonth and maxMonth via ByRef
' ============================================================================
Private Sub GetMonthRange(ByVal wsMain As Worksheet, ByRef minMonth As Long, ByRef maxMonth As Long)
10      Dim periodType As String
20      Dim periodDetail As String
        Dim wsMgmt As Worksheet
        Dim monthIdx As Long
        Dim monthName As String
30      periodType = Trim$(CStr(wsMain.Range("rngPeriodType").Value2))
40      periodDetail = Trim$(CStr(wsMain.Range("rngPeriodValue").Value2))

        ' Default: full year
50      minMonth = 1
60      maxMonth = 12

        ' "chodshi" = monthly
70      If InStr(1, periodType, ChrW$(1495) & ChrW$(1493) & ChrW$(1491) & ChrW$(1513) & ChrW$(1497), vbTextCompare) > 0 Then
            ' E4 contains Hebrew month name from NIHUL!R10:R21
            ' Match it against the month list to find month number
80          If periodDetail <> "" Then
90              Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
100             For monthIdx = 1 To 12
110                 monthName = Trim$(CStr(wsMgmt.Cells(ThisWorkbook.Names("rngSection_PeriodLists").RefersToRange.Row + 14 + monthIdx, 1).Value2))
120                 If StrComp(periodDetail, monthName, vbTextCompare) = 0 Then
130                     minMonth = monthIdx
140                     maxMonth = monthIdx
150                     Exit For
160                 End If
170             Next monthIdx
180         End If

        ' "riv'oni" = quarterly
190     ElseIf InStr(1, periodType, ChrW$(1512) & ChrW$(1489) & ChrW$(1506) & ChrW$(1493) & ChrW$(1504) & ChrW$(1497), vbTextCompare) > 0 Then
            ' E4 = riv'on rishon/sheni/shlishi/revi'i from NIHUL!R5:R8
            ' Match by checking which quarter keyword is in the detail
200         If InStr(1, periodDetail, ChrW$(1512) & ChrW$(1488) & ChrW$(1513) & ChrW$(1493) & ChrW$(1503), vbTextCompare) > 0 Then
210             minMonth = 1: maxMonth = 3
220         ElseIf InStr(1, periodDetail, ChrW$(1513) & ChrW$(1504) & ChrW$(1497), vbTextCompare) > 0 Then
230             minMonth = 4: maxMonth = 6
240         ElseIf InStr(1, periodDetail, ChrW$(1513) & ChrW$(1500) & ChrW$(1497) & ChrW$(1513) & ChrW$(1497), vbTextCompare) > 0 Then
250             minMonth = 7: maxMonth = 9
260         ElseIf InStr(1, periodDetail, ChrW$(1512) & ChrW$(1489) & ChrW$(1497) & ChrW$(1506) & ChrW$(1497), vbTextCompare) > 0 Then
270             minMonth = 10: maxMonth = 12
280         End If

        ' "chatzi shnati" = half yearly
290     ElseIf InStr(1, periodType, ChrW$(1495) & ChrW$(1510) & ChrW$(1497), vbTextCompare) > 0 Then
            ' E4 = machatzit rishona/shniya from NIHUL!R2:R3
300         If InStr(1, periodDetail, ChrW$(1512) & ChrW$(1488) & ChrW$(1513) & ChrW$(1493) & ChrW$(1504), vbTextCompare) > 0 Then
310             minMonth = 1: maxMonth = 6
320         ElseIf InStr(1, periodDetail, ChrW$(1513) & ChrW$(1504) & ChrW$(1497), vbTextCompare) > 0 Then
330             minMonth = 7: maxMonth = 12
340         End If

        ' "shnatit" or anything else = full year (already set as default)
350     End If
End Sub

' ============================================================================
' HELPER: Get date column based on B5 selection
' ============================================================================
Private Function GetDateColumn(ByVal wsMain As Worksheet) As Long
10      Dim v As String
20      v = Trim$(CStr(wsMain.Range("rngDateType").Value2))
        ' Hebrew: insurance start
30      If InStr(1, v, ChrW$(1514) & ChrW$(1495) & ChrW$(1497) & ChrW$(1500) & ChrW$(1514), vbTextCompare) > 0 Then
40          GetDateColumn = RAW_INSURANCE_START
50      Else
            ' Default: bordereu
60          GetDateColumn = RAW_BORDEREU
70      End If
End Function

' ============================================================================
' MACRO 1: BuildReview
' ============================================================================
Public Sub BuildReview()

10      Dim wsMgmt As Worksheet
20      Dim wsSrc As Worksheet
30      Dim wsRev As Worksheet
40      Dim wbSrc As Workbook
50      Dim dictHelper As Object
60      Dim dictFieldCol As Object
70      Dim dictFieldDisp As Object
80      Dim keys() As String
90      Dim cols() As Long
100     Dim disp() As String
110     Dim cnt As Long
120     Dim threshold As Double
130     Dim srcPath As String
140     Dim yearVal As String
150     Dim lastRow As Long
160     Dim r As Long
170     Dim i As Long
180     Dim outRow As Long
190     Dim reasonCode As String
200     Dim reasonText As String
210     Dim premiumVal As Variant
220     Dim premiumNum As Double
230     Dim prevScreenUpdating As Boolean
240     Dim prevDisplayAlerts As Boolean
250     Dim prevEnableEvents As Boolean
260     Dim prevCalculation As XlCalculation
270     Dim actionCol As Long
275     Dim j As Long
280     Dim rng As Range
        Dim revSheetName As String
        Dim msgBase As Long
        Dim confirmMsg As String
        Dim wsCleanup As Worksheet
        Dim iSheet As Long
        Dim wsMsgSrc As Worksheet
        Dim singleCode As String
        Dim singleText As String
        Dim currCode As Variant
        Dim dateColBR As Long
        Dim rowBordMonthBR As String
        Dim rowRateBR As Double
        Dim testNR As Variant

        ' --- Ensure SetupMainSheet has been run (Named Ranges exist) ---
        On Error Resume Next
        testNR = ThisWorkbook.Names("rngBaseYear").RefersToRange.Value2
        If Err.Number <> 0 Then
            On Error GoTo 0
            A00_SetupMainSheet
        End If
        Err.Clear
        ' Also ensure section Named Ranges exist
        testNR = ThisWorkbook.Names("rngSection_Messages").RefersToRange.Row
        If Err.Number <> 0 Then
            On Error GoTo 0
            SetupSectionNamedRanges
        End If
        On Error GoTo 0

        ' --- Check if data for these years was already processed ---
        Dim chkCurYear As String
        Dim chkBaseYear As String
        Dim chkSheetName As String
        chkCurYear = Trim$(CStr(ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("rngCurrentYear").Value2))
        chkBaseYear = Trim$(CStr(ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("rngBaseYear").Value2))
        chkSheetName = H_BASE() & "_" & chkBaseYear
        If SheetExists(chkSheetName) Or SheetExists("base_" & chkBaseYear) Then
            ' "netunim elu kvar nivdeku. im birtzoncha livdok shanim acherot, shne et ha'arachim shel shnat basis ve'shana nochekhit."
            Dim alreadyMsg As String
            ' "netunim elu kvar nivdeku. im birtzoncha livdok shanim acherot, shne et ha'arachim shel shnat basis veshana nochekhit. ha'im leharitz shuv?"
            alreadyMsg = ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1488) & ChrW(1500) & ChrW(1493) & " " & ChrW(1499) & ChrW(1489) & ChrW(1512) & " " & ChrW(1504) & ChrW(1489) & ChrW(1491) & ChrW(1511) & ChrW(1493) & "." & vbCrLf & vbCrLf & _
                ChrW(1488) & ChrW(1501) & " " & ChrW(1489) & ChrW(1512) & ChrW(1510) & ChrW(1493) & ChrW(1504) & ChrW(1498) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1493) & ChrW(1511) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1488) & ChrW(1495) & ChrW(1512) & ChrW(1493) & ChrW(1514) & "," & vbCrLf & _
                ChrW(1513) & ChrW(1504) & ChrW(1492) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1497) & ChrW(1501) & " " & ChrW(1513) & ChrW(1500) & " " & ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & H_BASE() & " " & ChrW(1493) & ChrW(1513) & ChrW(1504) & ChrW(1492) & " " & ChrW(1504) & ChrW(1493) & ChrW(1499) & ChrW(1495) & ChrW(1497) & ChrW(1514) & "." & vbCrLf & vbCrLf & _
                ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1500) & ChrW(1492) & ChrW(1512) & ChrW(1497) & ChrW(1509) & " " & ChrW(1513) & ChrW(1493) & ChrW(1489) & "?"
            If MsgBoxU(alreadyMsg, vbYesNo + vbQuestion) <> vbYes Then
                Exit Sub
            End If
        End If

        ' --- Pre-run confirmation message (read from Messages section) ---
        ' msgBase+20 stores "1" if user chose "don't show again"
        Set wsMsgSrc = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
        msgBase = ThisWorkbook.Names("rngSection_Messages").RefersToRange.Row
        If CStr(wsMsgSrc.Cells(msgBase + 20, 1).Value) <> "1" Then
            confirmMsg = CStr(wsMsgSrc.Cells(msgBase + 13, 1).Value) & vbNewLine & vbNewLine & CStr(wsMsgSrc.Cells(msgBase + 14, 1).Value) & vbNewLine & vbNewLine & CStr(wsMsgSrc.Cells(msgBase + 15, 1).Value) & vbNewLine & vbNewLine & CStr(wsMsgSrc.Cells(msgBase + 16, 1).Value)
            If MsgBoxU(confirmMsg, vbOKCancel + vbExclamation, CStr(wsMsgSrc.Cells(msgBase + 8, 1).Value)) <> vbOK Then
                Exit Sub
            End If
            ' Ask if user wants to keep showing this message
            If MsgBoxU(CStr(wsMsgSrc.Cells(msgBase + 17, 1).Value), vbYesNo + vbQuestion) = vbNo Then
                wsMsgSrc.Cells(msgBase + 20, 1).Value = "1"
            End If
        End If

290     On Error GoTo ERR_HANDLER
        ' Show prominent processing message on Main sheet
        Dim wsProgress1 As Worksheet
        Set wsProgress1 = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        wsProgress1.Unprotect "Z961814r"
        wsProgress1.Range("G15:K15").UnMerge
        wsProgress1.Range("G15:K15").Merge
        With wsProgress1.Range("G15")
            .Value = ChrW(1502) & ChrW(1506) & ChrW(1489) & ChrW(1491) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1488) & ChrW(1504) & ChrW(1488) & " " & ChrW(1492) & ChrW(1502) & ChrW(1514) & ChrW(1497) & ChrW(1504) & ChrW(1493) & "..."
            .Font.Size = 18
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
            .Interior.Color = RGB(255, 255, 200)
            .HorizontalAlignment = -4108
        End With
        wsProgress1.Activate
        wsProgress1.Protect UserInterfaceOnly:=True
        Application.ScreenUpdating = True
        DoEvents
        Application.ScreenUpdating = False
        ' Remove any leftover sheet protection
        Dim wsUp As Worksheet
        For Each wsUp In ThisWorkbook.Worksheets
            On Error Resume Next
            wsUp.Unprotect "Z961814r"
            On Error GoTo ERR_HANDLER
        Next wsUp

300     prevScreenUpdating = Application.ScreenUpdating
310     prevDisplayAlerts = Application.DisplayAlerts
320     prevEnableEvents = Application.EnableEvents
330     prevCalculation = Application.Calculation

340     Application.ScreenUpdating = False
350     Application.DisplayAlerts = False
360     Application.EnableEvents = False
370     Application.Calculation = xlCalculationManual

        ' --- Delete all sheets except daf habait, hagdarot, and basis_{refYear} ---
        Dim refBaseKeep As String
        Dim refBaseKeepOld As String
        Dim refYearStr As String
        refYearStr = Trim$(CStr(ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("rngBaseYear").Value2))
        refBaseKeep = H_BASE() & "_" & refYearStr
        refBaseKeepOld = "base_" & refYearStr
375     For iSheet = ThisWorkbook.Worksheets.Count To 1 Step -1
376         Set wsCleanup = ThisWorkbook.Worksheets(iSheet)
377         If wsCleanup.Name <> CONTROL_SHEET_NAME() And wsCleanup.Name <> MANAGEMENT_SHEET_NAME() And wsCleanup.Name <> MATACH_SHEET_NAME() And wsCleanup.Name <> refBaseKeep And wsCleanup.Name <> refBaseKeepOld And wsCleanup.Name <> NC_MASTER_SHEET_NAME() Then
                On Error Resume Next
                If wsCleanup.Visible <> xlSheetVisible Then wsCleanup.Visible = xlSheetVisible
378             wsCleanup.Delete
                On Error GoTo ERR_HANDLER
379         End If
380     Next iSheet
        ' Rename old English base name to Hebrew if needed
381     If SheetExists(refBaseKeepOld) Then ThisWorkbook.Worksheets(refBaseKeepOld).Name = refBaseKeep

382     Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
390     Set dictHelper = LoadHelperDictionary(wsMgmt)

400     ValidateHelperKey dictHelper, HELPER_REVIEW_SOURCE_ROW_HEADER
410     ValidateHelperKey dictHelper, HELPER_REVIEW_REASON_HEADER
420     ValidateHelperKey dictHelper, HELPER_REVIEW_REASON_CODE_HEADER

430     Set dictFieldCol = CreateObject("Scripting.Dictionary")
440     Set dictFieldDisp = CreateObject("Scripting.Dictionary")
450     dictFieldCol.CompareMode = vbTextCompare
460     dictFieldDisp.CompareMode = vbTextCompare
470     LoadCheckedFields wsMgmt, dictFieldCol, dictFieldDisp

480     BuildArrays dictFieldCol, dictFieldDisp, keys, cols, disp, cnt
490     If cnt = 0 Then Err.Raise vbObjectError + 1001, "BuildReview", "NO CHECKED FIELDS"

        Dim dictBranch As Object
495     Set dictBranch = LoadBranchMapping(wsMgmt)

        ' Dictionaries to collect unique values for filter lists
        Dim dictCompanies As Object, dictTellers As Object, dictAgents As Object
        Dim dictBranches As Object, dictMainBranches As Object
        Dim tmpVal As String, mbVal As String
        Dim wsLists As Worksheet
        Dim listsName As String
        Dim arrKeys As Variant, kk As Long
        Set dictCompanies = CreateObject("Scripting.Dictionary")
        Set dictTellers = CreateObject("Scripting.Dictionary")
        Set dictAgents = CreateObject("Scripting.Dictionary")
        Set dictBranches = CreateObject("Scripting.Dictionary")
        Set dictMainBranches = CreateObject("Scripting.Dictionary")
        dictCompanies.CompareMode = vbTextCompare
        dictTellers.CompareMode = vbTextCompare
        dictAgents.CompareMode = vbTextCompare
        dictBranches.CompareMode = vbTextCompare
        dictMainBranches.CompareMode = vbTextCompare

500     threshold = GetNumericParameter(wsMgmt, PARAM_PREMIUM_THRESHOLD)

        ' Load dollar exchange rate (from matach sheet / BOI API / rngDOLAR fallback)
        Dim dollarRate As Double
        Dim wsMain As Worksheet
502     Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
504     dollarRate = GetDollarRate(wsMain, , False)
        dateColBR = GetDateColumn(wsMain)

510     yearVal = Trim$(CStr(ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("rngCurrentYear").Value2))
520     If yearVal = "" Then Err.Raise vbObjectError + 1002, "BuildReview", "B3 IS EMPTY"

530     srcPath = FindSourceFile(yearVal)
540     If srcPath = "" Then Err.Raise vbObjectError + 1003, "BuildReview", "SOURCE FILE NOT FOUND FOR YEAR: " & yearVal

    Application.DisplayAlerts = True
550     Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True)
    Application.DisplayAlerts = False
560     Set wsSrc = OpenDataSheet(wbSrc)

570     lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
580     If lastRow < 2 Then Err.Raise vbObjectError + 1006, "BuildReview", "NO DATA ROWS IN SOURCE"

        ' Pre-fetch and cache all exchange rates (USD+EUR) for unique bordero months
585     Call BuildMatachCache(wsSrc, RAW_BORDEREU, lastRow)

        ' Use year-specific REVIEW sheet name
590     revSheetName = REVIEW_SHEET_NAME() & "_" & yearVal
591     DeleteSheetIfExists revSheetName

600     Set wsRev = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
610     wsRev.Name = revSheetName

        ' Row 1 is reserved for the button; headers start at row 2, data at row 3
620     wsRev.Cells(2, 1).Value = dictHelper(HELPER_REVIEW_SOURCE_ROW_HEADER)

630     For i = 1 To cnt
640         wsRev.Cells(2, i + 1).Value = disp(i)
650     Next i

660     wsRev.Cells(2, cnt + 2).Value = dictHelper(HELPER_REVIEW_REASON_HEADER)

        ' Action dropdown column header
680     wsRev.Cells(2, cnt + 3).Value = ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) & ChrW(1492)
        ' Fix text column header
690     wsRev.Cells(2, cnt + 4).Value = ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1503)

700     outRow = 3

710     For r = 2 To lastRow

720         If IsIgnorableRow(wsSrc, r, keys, cols, cnt, dictFieldCol) Then GoTo NextRow

            ' Check missing fields - write one REVIEW row per issue
740         For i = 1 To cnt
750             If IsBlankValue(wsSrc.Cells(r, cols(i)).Value2) Then
751                 singleCode = "MISSING_" & keys(i)
752                 If dictHelper.Exists(singleCode) Then
753                     singleText = dictHelper(singleCode)
754                 Else
755                     singleText = singleCode
756                 End If
760                 wsRev.Cells(outRow, 1).Value = r
761                 For j = 1 To cnt
762                     wsRev.Cells(outRow, j + 1).Value = wsSrc.Cells(r, cols(j)).Value2
763                 Next j
764                 wsRev.Cells(outRow, cnt + 2).Value = singleText
766                 outRow = outRow + 1
770             End If
780         Next i

            ' Check premium threshold - separate row (after currency conversion)
790         premiumVal = wsSrc.Cells(r, dictFieldCol(KEY_PREMIUM)).Value2
800         If Not IsBlankValue(premiumVal) Then
810             If TryParseVariantNumber(premiumVal, premiumNum) Then
                    ' Currency conversion: column AA (27) - 1=dollar, 0/90=ILS (per-month rate)
812                 currCode = wsSrc.Cells(r, RAW_CURRENCY).Value2
813                 rowBordMonthBR = GetBordMonth(wsSrc.Cells(r, dateColBR).Value2)
                    If rowBordMonthBR <> "" Then
                        rowRateBR = GetDollarRate(wsMain, rowBordMonthBR, True)
                    Else
                        rowRateBR = dollarRate
                    End If
814                 If IsNumeric(currCode) Then
816                     If CLng(currCode) = 1 Then premiumNum = premiumNum * rowRateBR
818                 End If
820                 If Abs(premiumNum) > threshold Then
821                     singleCode = "PREMIUM_OVER_THRESHOLD"
822                     If dictHelper.Exists(singleCode) Then
823                         singleText = dictHelper(singleCode)
824                     Else
825                         singleText = singleCode
826                     End If
830                     wsRev.Cells(outRow, 1).Value = r
831                     For j = 1 To cnt
832                         wsRev.Cells(outRow, j + 1).Value = wsSrc.Cells(r, cols(j)).Value2
833                     Next j
834                     wsRev.Cells(outRow, cnt + 2).Value = singleText
836                     outRow = outRow + 1
840                 End If
850             Else
851                 singleCode = "PREMIUM_NOT_NUMERIC"
852                 If dictHelper.Exists(singleCode) Then
853                     singleText = dictHelper(singleCode)
854                 Else
855                     singleText = singleCode
856                 End If
860                 wsRev.Cells(outRow, 1).Value = r
861                 For j = 1 To cnt
862                     wsRev.Cells(outRow, j + 1).Value = wsSrc.Cells(r, cols(j)).Value2
863                 Next j
864                 wsRev.Cells(outRow, cnt + 2).Value = singleText
866                 outRow = outRow + 1
870             End If
880         End If

            ' Check branch mapping - if branch not in translation table
890         Dim brKey As String
900         brKey = UCase$(Trim$(CStr(wsSrc.Cells(r, RAW_BRANCHNAME).Value2)))
910         If brKey <> "" Then
920             If Not dictBranch.Exists(brKey) Then
921                 singleCode = "UNKNOWN_BRANCH"
922                 If dictHelper.Exists(singleCode) Then
923                     singleText = dictHelper(singleCode)
924                 Else
925                     singleText = H_BRANCH() & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1502) & ChrW(1494) & ChrW(1493) & ChrW(1492) & ChrW(1492)
926                 End If
930                 wsRev.Cells(outRow, 1).Value = r
931                 For j = 1 To cnt
932                     wsRev.Cells(outRow, j + 1).Value = wsSrc.Cells(r, cols(j)).Value2
933                 Next j
934                 wsRev.Cells(outRow, cnt + 2).Value = singleText
936                 outRow = outRow + 1
                    ' Add new branch to settings A:D (if not already there)
                    Dim brOrigName As String
                    brOrigName = Trim$(CStr(wsSrc.Cells(r, RAW_BRANCHNAME).Value2))
                    Dim brLastRow As Long
                    brLastRow = wsMgmt.Cells(wsMgmt.Rows.Count, 1).End(xlUp).Row
                    Dim brExists As Boolean
                    brExists = False
                    Dim brScan As Long
                    For brScan = 3 To brLastRow
                        If StrComp(Trim$(CStr(wsMgmt.Cells(brScan, 1).Value2)), brOrigName, vbTextCompare) = 0 Then
                            brExists = True
                            Exit For
                        End If
                    Next brScan
                    If Not brExists Then
                        brLastRow = brLastRow + 1
                        wsMgmt.Cells(brLastRow, 1).Value = brOrigName
                        ' Column B (main branch) left empty - to be filled via correction
                        ' Column C: generate BRANCH_KEY via transliteration
                        wsMgmt.Cells(brLastRow, 3).Value = HebrewToKey(brOrigName)
                        ' Column D (main branch key) left empty
                    End If
                    ' Add to dictBranch so same branch isn't flagged again in this run
                    dictBranch(brKey) = ""
940             End If
950         End If

            ' Collect unique values for filter lists
960         tmpVal = Trim$(CStr(wsSrc.Cells(r, RAW_COMPANY).Value2))
            If tmpVal <> "" Then If Not dictCompanies.Exists(tmpVal) Then dictCompanies.Add tmpVal, 1
962         tmpVal = Trim$(CStr(wsSrc.Cells(r, RAW_TELLERNAME).Value2))
            If tmpVal <> "" Then If Not dictTellers.Exists(tmpVal) Then dictTellers.Add tmpVal, 1
964         tmpVal = Trim$(CStr(wsSrc.Cells(r, RAW_AGENTNAME).Value2))
            If tmpVal <> "" Then If Not dictAgents.Exists(tmpVal) Then dictAgents.Add tmpVal, 1
966         tmpVal = Trim$(CStr(wsSrc.Cells(r, RAW_BRANCHNAME).Value2))
            If tmpVal <> "" Then If Not dictBranches.Exists(tmpVal) Then dictBranches.Add tmpVal, 1
968         If tmpVal <> "" Then
969             If dictBranch.Exists(UCase$(tmpVal)) Then
971                 mbVal = dictBranch(UCase$(tmpVal))
972                 If Not dictMainBranches.Exists(mbVal) Then dictMainBranches.Add mbVal, 1
973             End If
974         End If

NextRow:
990     Next r

        ' ---- Write unique lists to hidden "reshimot" sheet ----
991     listsName = H_RESHIMOT()
992     DeleteSheetIfExists listsName
993     Set wsLists = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
994     wsLists.Name = listsName
995     wsLists.Visible = xlSheetVeryHidden

        ' Headers
996     wsLists.Cells(1, 1).Value = H_COMPANY()
        wsLists.Cells(1, 2).Value = H_TELLER()
        wsLists.Cells(1, 3).Value = H_AGENT()
        wsLists.Cells(1, 4).Value = H_BRANCH()
        wsLists.Cells(1, 5).Value = H_BRANCH() & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)

        ' Write data
997     If dictCompanies.Count > 0 Then
            arrKeys = dictCompanies.keys
            For kk = 0 To UBound(arrKeys): wsLists.Cells(kk + 2, 1).Value = arrKeys(kk): Next kk
        End If
998     If dictTellers.Count > 0 Then
            arrKeys = dictTellers.keys
            For kk = 0 To UBound(arrKeys): wsLists.Cells(kk + 2, 2).Value = arrKeys(kk): Next kk
        End If
999     If dictAgents.Count > 0 Then
            arrKeys = dictAgents.keys
            For kk = 0 To UBound(arrKeys): wsLists.Cells(kk + 2, 3).Value = arrKeys(kk): Next kk
        End If
        If dictBranches.Count > 0 Then
            arrKeys = dictBranches.keys
            For kk = 0 To UBound(arrKeys): wsLists.Cells(kk + 2, 4).Value = arrKeys(kk): Next kk
        End If
        If dictMainBranches.Count > 0 Then
            arrKeys = dictMainBranches.keys
            For kk = 0 To UBound(arrKeys): wsLists.Cells(kk + 2, 5).Value = arrKeys(kk): Next kk
        End If

        ' Add dropdown validation for action column
1000    actionCol = cnt + 3
1010    If outRow > 3 Then
1020        Set rng = wsRev.Range(wsRev.Cells(3, actionCol), wsRev.Cells(outRow - 1, actionCol))
            ' Set default value: "ha'aver livdika" = transfer for review
1025        rng.Value = H_REVIEW()
1030        On Error Resume Next
1035        rng.Validation.Delete
1040        rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=H_FIX() & "," & H_IGNORE() & "," & H_REVIEW()
1050        rng.Validation.InCellDropdown = True
1060        On Error GoTo ERR_HANDLER
1070    End If

1080    wbSrc.Close SaveChanges:=False
1090    Set wbSrc = Nothing

1100    wsRev.Rows(2).Font.Bold = True
1110    wsRev.Columns.AutoFit

        ' Format action and fix columns to be clearly visible
1112    wsRev.Cells(2, actionCol).Interior.Color = RGB(255, 165, 0)
1113    wsRev.Cells(2, actionCol + 1).Interior.Color = RGB(255, 165, 0)
1114    wsRev.Columns(actionCol).ColumnWidth = 15
1115    wsRev.Columns(actionCol + 1).ColumnWidth = 25
        ' Light yellow fill for data area of action/fix columns
1116    If outRow > 3 Then
1117        wsRev.Range(wsRev.Cells(3, actionCol), wsRev.Cells(outRow - 1, actionCol + 1)).Interior.Color = RGB(255, 255, 200)
1118    End If

        ' Add "Done Updating" button in ROW 1 (above headers) - positioned at column 1
1119    Dim shpBtn As Shape
        wsRev.Rows(1).RowHeight = 30
1120    Set shpBtn = wsRev.Shapes.AddShape(msoShapeRoundedRectangle, wsRev.Cells(1, 1).Left + 2, 2, 110, 26)
1121    shpBtn.Name = "btnSendForReview"
1122    shpBtn.Fill.ForeColor.RGB = RGB(180, 0, 0)
        ' "siymti le'adken" = done updating
1123    shpBtn.TextFrame2.TextRange.Text = ChrW(1505) & ChrW(1497) & ChrW(1497) & ChrW(1502) & ChrW(1514) & ChrW(1497) & " " & ChrW(1500) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1503)
1124    shpBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
1125    shpBtn.TextFrame2.TextRange.Font.Size = 11
1126    shpBtn.TextFrame2.TextRange.Font.Bold = msoTrue
1127    shpBtn.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shpBtn.TextFrame2.MarginLeft = 2
        shpBtn.TextFrame2.MarginRight = 2
1128    shpBtn.OnAction = "SendForReview"

CLEANUP:
1200    Application.ScreenUpdating = prevScreenUpdating
1210    Application.DisplayAlerts = prevDisplayAlerts
1220    Application.EnableEvents = prevEnableEvents
1230    Application.Calculation = prevCalculation

        ' Restore E4 dropdown after sheet cleanup
1235    UpdatePeriodDropdown

        ' Activate the review sheet so user sees it
1239    On Error Resume Next
        If Not wsRev Is Nothing Then wsRev.Activate
        On Error GoTo 0

1240    If outRow > 3 Then
            ' "Finished - found X issues" - stay on review sheet
            MsgBoxU wsMsgSrc.Cells(msgBase + 2, 1).Value & (outRow - 3) & wsMsgSrc.Cells(msgBase + 3, 1).Value, vbInformation
        Else
            ' No issues found - go straight to home
            MsgBoxU wsMsgSrc.Cells(msgBase + 4, 1).Value, vbInformation
            ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Activate
        End If

        ' Update client list in settings sheet (non-critical)
        On Error Resume Next
1245    UpdateClientList
        On Error GoTo ERR_HANDLER
        ' Clear processing message
        wsProgress1.Unprotect "Z961814r"
1246    wsProgress1.Range("G15:K15").UnMerge
        wsProgress1.Range("G15:K15").ClearContents
1247    wsProgress1.Range("G15:K15").Interior.Color = RGB(220, 240, 220)
        wsProgress1.Protect UserInterfaceOnly:=True
1250    Exit Sub

ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Dim errLine As Long
        Dim errNum As Long
        Dim errDesc As String
        Dim errSrc As String
        errLine = Erl
        errNum = Err.Number
        errDesc = Err.Description
        errSrc = Err.Source
1300    On Error Resume Next
1310    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
1320    Application.ScreenUpdating = True
1330    Application.DisplayAlerts = True
1340    Application.EnableEvents = True
1350    Application.Calculation = xlCalculationAutomatic
1360    MsgBoxU wsMsgSrc.Cells(msgBase + 5, 1).Value & errLine & vbCrLf & wsMsgSrc.Cells(msgBase + 6, 1).Value & errNum & vbCrLf & errSrc & vbCrLf & errDesc, vbCritical

End Sub

' ============================================================================
' MACRO 2: ApplyCorrectionsAndBuildReports
' ============================================================================
Public Sub ApplyCorrectionsAndBuildReports()

10      Dim wsMain As Worksheet
20      Dim wsMgmt As Worksheet
30      Dim wsRev As Worksheet
40      Dim wbSrc As Workbook
50      Dim wsSrc As Worksheet
60      Dim wbRef As Workbook
70      Dim wsRef As Worksheet
80      Dim wsBase As Worksheet
81      Dim wsBaseRef As Worksheet
90      Dim dictHelper As Object
100     Dim dictBranch As Object
110     Dim yearVal As String
120     Dim refYear As String
130     Dim srcPath As String
140     Dim refPath As String
150     Dim threshold As Double
160     Dim maxMonth As Long
        Dim minMonth As Long
        Dim dateCol As Long
170     Dim lastRow As Long
180     Dim r As Long
190     Dim outRow As Long
200     Dim baseSheetName As String
210     Dim refBaseSheetName As String
220     Dim corrCount As Long
230     Dim ignoreCount As Long
240     Dim unhandledCount As Long
250     Dim reviewCount As Long
260     Dim countRef As Long
270     Dim countCurrent As Long
280     Dim prevScreenUpdating As Boolean
290     Dim prevDisplayAlerts As Boolean
300     Dim prevEnableEvents As Boolean
310     Dim prevCalculation As XlCalculation
320     Dim debugStep As String
330     Dim revLastRow As Long
340     Dim revLastCol As Long
350     Dim actionText As String
360     Dim fixText As String
370     Dim srcRowNum As Long
380     Dim dictCorrections As Object
390     Dim dictIgnore As Object
400     Dim ans As VbMsgBoxResult
410     Dim monthVal As Long
420     Dim premVal As Double
        Dim commVal As Double
        Dim currCode As Variant
430     Dim bordereu As Variant
        Dim fixParts() As String
        Dim fp As Long
        Dim oneFix As String
        Dim fixPrem As Double
        Dim fixComm As Double
        Dim curRevName As String
        Dim actionColIdx As Long
        Dim fixColIdx As Long
        Dim hdrCol As Long
        Dim hdrActionText As String
        Dim hdrFixText As String
        Dim dictHasFix As Object
        Dim dictHasIgnore As Object
        Dim dictHasUnhandled As Object
        Dim srcKey As Variant
        Dim allKeys As Object
        Dim c As Long
        Dim refRevName As String
        Dim dictRefCorr As Object
        Dim dictRefIgnore As Object
        Dim wsRefRev As Worksheet
        Dim refRevLastRow As Long
        Dim refRevLastCol As Long
        Dim refActionColIdx As Long
        Dim refFixColIdx As Long
        Dim refHdrCol As Long
        Dim refSrcRowNum As Long
        Dim refActionText As String
        Dim refFixText As String
        Dim dictRefHasFix As Object
        Dim dictRefHasIgnore As Object
        Dim dictRefHasUnhandled As Object
        Dim refSrcKey As Variant
        Dim refAllKeys As Object
        Dim periodDesc As String
        Dim reasonColIdx As Long
        Dim hdrReasonText As String
        Dim reasonText As String
        Dim dictRowFixes As Object
        Dim fixKey As Variant
        Dim refReasonColIdx As Long
        Dim refReasonText As String
        Dim dictRefRowFixes As Object
        Dim refFixKey As Variant
        Dim rowBordMonth As String
        Dim rowRate As Double
        Dim dtStr As String
        Dim mPart As String
        Dim dollarRate As Double
        Dim clientFilter As String

440     On Error GoTo ERR_HANDLER

        ' Remove any leftover sheet protection
        Dim wsUp2 As Worksheet
        For Each wsUp2 In ThisWorkbook.Worksheets
            On Error Resume Next
            wsUp2.Unprotect "Z961814r"
            On Error GoTo ERR_HANDLER
        Next wsUp2

        ' --- Check if reports already exist for these parameters ---
        Dim wsChk2 As Worksheet
        Set wsChk2 = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        Dim chkSheets2 As String
        chkSheets2 = SHEET_COMPANIES()
        If SheetExists(chkSheets2) Then
            Dim alreadyMsg2 As String
            ' "dochot elu kvar keiyamim. nitan litzpot bahem bekaftor 5. ha'im leharitz shuv?"
            alreadyMsg2 = ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1488) & ChrW(1500) & ChrW(1493) & " " & ChrW(1499) & ChrW(1489) & ChrW(1512) & " " & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1502) & ChrW(1497) & ChrW(1501) & "." & vbCrLf & _
                ChrW(1504) & ChrW(1497) & ChrW(1514) & ChrW(1503) & " " & ChrW(1500) & ChrW(1510) & ChrW(1508) & ChrW(1493) & ChrW(1514) & " " & ChrW(1489) & ChrW(1492) & ChrW(1501) & " " & ChrW(1489) & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " 5." & vbCrLf & vbCrLf & _
                ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1500) & ChrW(1492) & ChrW(1512) & ChrW(1497) & ChrW(1509) & " " & ChrW(1513) & ChrW(1493) & ChrW(1489) & "?"
            If MsgBoxU(alreadyMsg2, vbYesNo + vbQuestion) <> vbYes Then
                Exit Sub
            End If
        End If

        ' --- Validate required dropdown selections ---
        Dim wsCheck As Worksheet
        Set wsCheck = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        Dim selText As String
        selText = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)  ' "bechar/i"
        
        clientFilter = Trim$(CStr(wsCheck.Range("rngClientName").Value2))
        If clientFilter = selText Or clientFilter = "" Then
            clientFilter = ""
        End If
        
        Dim ptVal As String
        ptVal = Trim$(CStr(wsCheck.Range("rngPeriodType").Value2))
        ' If period type requires a value (not shnatit/yearly), check G6
        If ptVal <> "" And ptVal <> selText Then
            ' Check if NOT shnatit (yearly doesn't need G6)
            ' Must check it's shnatit but NOT chatzi shnati (which contains shnatit as substring)
            Dim isYearly As Boolean
            isYearly = (InStr(1, ptVal, ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497), vbTextCompare) > 0) And _
                        (InStr(1, ptVal, ChrW(1495) & ChrW(1510) & ChrW(1497), vbTextCompare) = 0)
            If Not isYearly Then
                Dim pvVal As String
                pvVal = Trim$(CStr(wsCheck.Range("rngPeriodValue").Value2))
                If pvVal = "" Or pvVal = selText Then
                    MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1489) & ChrW(1495) & ChrW(1493) & ChrW(1512) & " " & ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492) & " " & ChrW(1489) & ChrW(1514) & ChrW(1488) & " G6", vbExclamation
                    Exit Sub
                End If
            End If
        End If
        
        Dim ftVal As String
        ftVal = Trim$(CStr(wsCheck.Range("rngFilterType").Value2))
        If ftVal <> "" And ftVal <> selText Then
            Dim fvVal As String
            fvVal = Trim$(CStr(wsCheck.Range("rngFilterValue").Value2))
            If fvVal = "" Or fvVal = selText Then
                MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1489) & ChrW(1495) & ChrW(1493) & ChrW(1512) & " " & ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1505) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1503) & " " & ChrW(1489) & ChrW(1514) & ChrW(1488) & " G10", vbExclamation
                Exit Sub
            End If
        End If
        ' --- End validation ---

        Dim wsMsgSrc2 As Worksheet
        Dim msgBase2 As Long
        Set wsMsgSrc2 = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
        msgBase2 = ThisWorkbook.Names("rngSection_Messages").RefersToRange.Row

        ' Show prominent processing message on Main sheet (below currency area)
        Dim wsProgress As Worksheet
        Set wsProgress = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        wsProgress.Unprotect "Z961814r"
470     wsProgress.Range("G15:K15").UnMerge
        wsProgress.Range("G15:K15").Merge
        With wsProgress.Range("G15")
            .Value = ChrW(1502) & ChrW(1506) & ChrW(1489) & ChrW(1491) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1488) & ChrW(1504) & ChrW(1488) & " " & ChrW(1492) & ChrW(1502) & ChrW(1514) & ChrW(1497) & ChrW(1504) & ChrW(1493) & "..."
            .Font.Size = 18
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
            .Interior.Color = RGB(255, 255, 200)
            .HorizontalAlignment = -4108
        End With
        wsProgress.Activate
        wsProgress.Protect UserInterfaceOnly:=True
        Application.ScreenUpdating = True
        DoEvents
        Application.ScreenUpdating = False

475     debugStep = "INIT"

480     prevScreenUpdating = Application.ScreenUpdating
490     prevDisplayAlerts = Application.DisplayAlerts
500     prevEnableEvents = Application.EnableEvents
510     prevCalculation = Application.Calculation

520     Application.ScreenUpdating = False
530     Application.DisplayAlerts = False
540     Application.EnableEvents = False
550     Application.Calculation = xlCalculationManual

560     debugStep = "LOAD_SHEETS"
570     Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
580     Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())

590     debugStep = "READ_YEARS"
600     yearVal = Trim$(CStr(wsMain.Range("rngCurrentYear").Value2))
610     refYear = Trim$(CStr(wsMain.Range("rngBaseYear").Value2))
620     If yearVal = "" Or refYear = "" Then Err.Raise vbObjectError + 2001, "ApplyCorrections", "B2 OR B3 IS EMPTY"

630     debugStep = "LOAD_HELPER"
640     Set dictHelper = LoadHelperDictionary(wsMgmt)

650     debugStep = "LOAD_BRANCH"
660     Set dictBranch = LoadBranchMapping(wsMgmt)

670     debugStep = "GET_THRESHOLD"
680     threshold = GetNumericParameter(wsMgmt, PARAM_PREMIUM_THRESHOLD)

        ' Load dollar exchange rate (from matach sheet / BOI API / rngDOLAR fallback)
682     dollarRate = GetDollarRate(wsMain, , False)

690     debugStep = "GET_PERIOD"
700     GetMonthRange wsMain, minMonth, maxMonth
        dateCol = GetDateColumn(wsMain)

        ' ---- Read corrections from REVIEW ----

    corrCount = 0
    ignoreCount = 0
    unhandledCount = 0
    reviewCount = 0

    LoadCorrectionsToDicts wsMgmt, yearVal, dictHelper, _
        dictCorrections, dictIgnore, dictHasFix, dictHasIgnore, dictHasUnhandled, dictRowFixes, _
        corrCount, ignoreCount, unhandledCount, reviewCount

    LoadCorrectionsToDicts wsMgmt, refYear, dictHelper, _
        dictRefCorr, dictRefIgnore, dictRefHasFix, dictRefHasIgnore, dictRefHasUnhandled, dictRefRowFixes, _
        corrCount, ignoreCount, unhandledCount, reviewCount

1040    debugStep = "CHECK_SRC"
1050    srcPath = FindSourceFile(yearVal)
1060    If srcPath = "" Then Err.Raise vbObjectError + 2002, "ApplyCorrections", "SOURCE NOT FOUND: " & yearVal

1070    debugStep = "OPEN_SRC"
1080    Set wbSrc = SafeOpenWorkbook(srcPath)
    If wbSrc Is Nothing Then Err.Raise 424, "OPEN_SRC", "Failed to open source file. Excel might be blocking it."
1090    Set wsSrc = OpenDataSheet(wbSrc)

    baseSheetName = H_BASE() & "_" & yearVal
    BuildBaseSheet wsSrc, wsBase, baseSheetName, _
        dictIgnore, dictCorrections, dictRowFixes, dictHasFix, dictHasUnhandled, _
        dictBranch, _
        minMonth, maxMonth, filterCol, filterVal, dollarRate, dateCol, threshold, wsMain, countCurrent

    refPath = FindSourceFile(refYear)
    If refPath = "" Then Err.Raise vbObjectError + 2003, "ApplyCorrections", "REF SOURCE NOT FOUND: " & refYear
    
    debugStep = "OPEN_REF"
    Set wbRef = SafeOpenWorkbook(refPath)
    If wbRef Is Nothing Then Err.Raise 424, "OPEN_REF", "Failed to open ref file. Excel might be blocking it."
    Set wsRef = OpenDataSheet(wbRef)

    refBaseSheetName = H_BASE() & "_" & refYear
    BuildBaseSheet wsRef, wsBaseRef, refBaseSheetName, _
        dictRefIgnore, dictRefCorr, dictRefRowFixes, dictRefHasFix, dictRefHasUnhandled, _
        dictBranch, _
        minMonth, maxMonth, filterCol, filterVal, dollarRate, dateCol, threshold, wsMain, countRef

2630    debugStep = "BUILD_COMPANIES"
2640    BuildComparisonSheet wsBase, wsBaseRef, SHEET_COMPANIES(), BASE_COL_COMPANY, minMonth, maxMonth, yearVal, refYear, filterCol, filterValue, hashvaatText & SHEET_COMPANIES() & " | " & titlePrefix, clientFilter
        ThisWorkbook.Worksheets(SHEET_COMPANIES()).Tab.Color = RGB(173, 216, 230)  ' pastel blue
        ApplyZebraStriping ThisWorkbook.Worksheets(SHEET_COMPANIES())

2650    debugStep = "BUILD_BRANCH"
2660    BuildComparisonSheet wsBase, wsBaseRef, SHEET_BRANCH(), BASE_COL_BRANCHNAME, minMonth, maxMonth, yearVal, refYear, filterCol, filterValue, hashvaatText & SHEET_BRANCH() & " | " & titlePrefix, clientFilter
        ThisWorkbook.Worksheets(SHEET_BRANCH()).Tab.Color = RGB(255, 218, 185)  ' pastel peach
        ApplyZebraStriping ThisWorkbook.Worksheets(SHEET_BRANCH())

2665    debugStep = "BUILD_MAINBRANCH"
2666    BuildComparisonSheet wsBase, wsBaseRef, SHEET_MAINBRANCH(), BASE_COL_MAINBRANCH, minMonth, maxMonth, yearVal, refYear, filterCol, filterValue, hashvaatText & SHEET_MAINBRANCH() & " | " & titlePrefix, clientFilter
        ThisWorkbook.Worksheets(SHEET_MAINBRANCH()).Tab.Color = RGB(255, 182, 193)  ' pastel pink
        ApplyZebraStriping ThisWorkbook.Worksheets(SHEET_MAINBRANCH())

2670    debugStep = "BUILD_TELLERS"
2680    BuildComparisonSheet wsBase, wsBaseRef, SHEET_TELLERS(), BASE_COL_TELLER, minMonth, maxMonth, yearVal, refYear, filterCol, filterValue, hashvaatText & SHEET_TELLERS() & " | " & titlePrefix, clientFilter
        ThisWorkbook.Worksheets(SHEET_TELLERS()).Tab.Color = RGB(204, 204, 255)  ' pastel lavender
        ApplyZebraStriping ThisWorkbook.Worksheets(SHEET_TELLERS())

2690    debugStep = "BUILD_AGENTS"
2700    BuildComparisonSheet wsBase, wsBaseRef, SHEET_AGENTS(), BASE_COL_AGENTNAME, minMonth, maxMonth, yearVal, refYear, filterCol, filterValue, hashvaatText & SHEET_AGENTS() & " | " & titlePrefix, clientFilter
        ThisWorkbook.Worksheets(SHEET_AGENTS()).Tab.Color = RGB(176, 226, 172)  ' pastel green
        ApplyZebraStriping ThisWorkbook.Worksheets(SHEET_AGENTS())

2710    debugStep = "BUILD_MONTHS"
2720    BuildComparisonSheet wsBase, wsBaseRef, SHEET_MONTHS(), BASE_COL_MONTH, minMonth, maxMonth, yearVal, refYear, filterCol, filterValue, hashvaatText & SHEET_MONTHS() & " | " & titlePrefix, clientFilter
        ThisWorkbook.Worksheets(SHEET_MONTHS()).Tab.Color = RGB(255, 255, 186)  ' pastel yellow
        ApplyZebraStriping ThisWorkbook.Worksheets(SHEET_MONTHS())

2730    debugStep = "BUILD_SUMMARY"
        periodDesc = Trim$(CStr(wsMain.Range("rngPeriodType").Value2))
        If Trim$(CStr(wsMain.Range("rngPeriodValue").Value2)) <> "" Then periodDesc = periodDesc & " " & Trim$(CStr(wsMain.Range("rngPeriodValue").Value2))
2740    BuildSummarySheet countRef, countCurrent, reviewCount, corrCount, ignoreCount, unhandledCount, yearVal, refYear, periodDesc, hashvaatText & SHEET_SUMMARY() & " | " & titlePrefix
        ThisWorkbook.Worksheets(SHEET_SUMMARY()).Tab.Color = RGB(255, 204, 153)  ' pastel orange

CLEANUP:
        On Error Resume Next
        If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
        If Not wbRef Is Nothing Then wbRef.Close SaveChanges:=False
        On Error GoTo 0
        
2750    Application.ScreenUpdating = prevScreenUpdating
2760    Application.DisplayAlerts = prevDisplayAlerts
2770    Application.EnableEvents = prevEnableEvents
2780    Application.Calculation = prevCalculation

        ' Clear processing message from Main sheet and restore green background
2785    On Error Resume Next
        ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Unprotect "Z961814r"
        ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("G15:K15").UnMerge
        With ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("G15:K15")
            .Value = ""
            .Interior.Color = RGB(220, 240, 220)
        End With
        ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Protect UserInterfaceOnly:=True
        On Error GoTo 0
2790    Application.StatusBar = False

        ' Hide internal sheets (reshimot, letipul, basis, hagdarot)
        On Error Resume Next
        Dim wsHide As Worksheet
        For Each wsHide In ThisWorkbook.Worksheets
            If wsHide.Name = MANAGEMENT_SHEET_NAME() Then
                wsHide.Visible = xlSheetVeryHidden
            ElseIf Left$(wsHide.Name, Len(REVIEW_SHEET_NAME())) = REVIEW_SHEET_NAME() Then
                wsHide.Visible = xlSheetVeryHidden
            ElseIf Left$(wsHide.Name, 5) = H_BASE() & "_" Then
                wsHide.Visible = xlSheetVeryHidden
            ElseIf wsHide.Name = H_RESHIMOT() Then
                wsHide.Visible = xlSheetVeryHidden
            End If
        Next wsHide
        On Error GoTo 0

        ' Ensure home sheet is visible and activate it
        ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Visible = xlSheetVisible
        ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Activate

2795    MsgBoxU ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501), vbInformation

2800    Exit Sub

ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
2810    Dim errLine As Long
2820    Dim errNum As Long
2830    Dim errSrc As String
2840    Dim errDesc As String
2850    errLine = Erl
2860    errNum = Err.Number
2870    errSrc = Err.Source
2880    errDesc = Err.Description

2890    On Error Resume Next
2900    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
2910    If Not wbRef Is Nothing Then wbRef.Close SaveChanges:=False
2920    Application.ScreenUpdating = True
2930    Application.DisplayAlerts = True
2940    Application.EnableEvents = True
2950    Application.Calculation = xlCalculationAutomatic
        On Error Resume Next
        ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Unprotect "Z961814r"
        ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("G15:K15").UnMerge
        With ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("G15:K15")
            .Value = ""
            .Interior.Color = RGB(220, 240, 220)
        End With
        ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Protect UserInterfaceOnly:=True
2955    Application.StatusBar = False

2960    MsgBoxU wsMsgSrc2.Cells(msgBase2 + 5, 1).Value & errLine & vbCrLf & wsMsgSrc2.Cells(msgBase2 + 6, 1).Value & errNum & vbCrLf & wsMsgSrc2.Cells(msgBase2 + 7, 1).Value & debugStep & vbCrLf & errSrc & vbCrLf & errDesc, vbCritical

End Sub

' ============================================================================
' HELPER: Build comparison sheet (companies, branch, tellers, agents, months)
' ============================================================================

' ============================================================================
' HELPER: LoadCorrectionsToDicts
' ============================================================================
Private Sub LoadCorrectionsToDicts(ByVal wsMgmt As Worksheet, ByVal yearStr As String, ByVal dictHelper As Object, _
    ByRef dictCorrections As Object, ByRef dictIgnore As Object, ByRef dictHasFix As Object, ByRef dictHasIgnore As Object, ByRef dictHasUnhandled As Object, ByRef dictRowFixes As Object, _
    ByRef corrCount As Long, ByRef ignoreCount As Long, ByRef unhandledCount As Long, ByRef reviewCount As Long)

    Dim revLastRow As Long, revLastCol As Long
    Dim c As Long, fp As Long
    Dim curRevName As String
    Dim actionColIdx As Long, fixColIdx As Long, reasonColIdx As Long
    Dim hdrActionText As String, hdrFixText As String, hdrReasonText As String
    Dim actionText As String, fixText As String, reasonText As String
    Dim srcRowNum As Long
    Dim fixParts() As String, oneFix As String, fixKey As Variant
    Dim srcKey As Variant

    Dim wsRefRev As Worksheet
    Dim refRevLastRow As Long, refRevLastCol As Long
    Dim refActionColIdx As Long, refFixColIdx As Long, refReasonColIdx As Long
    Dim refActionText As String, refFixText As String, refReasonText As String
    Dim refSrcRowNum As Long
    Dim refFixKey As Variant, refSrcKey As Variant
    Dim wsRev As Worksheet
    Dim hdrCol As Long, r As Long
    
710     debugStep = "READ_CORRECTIONS"
720     Set dictCorrections = CreateObject("Scripting.Dictionary")
730     Set dictIgnore = CreateObject("Scripting.Dictionary")
        Set dictHasFix = CreateObject("Scripting.Dictionary")
        Set dictHasIgnore = CreateObject("Scripting.Dictionary")
        Set dictHasUnhandled = CreateObject("Scripting.Dictionary")
        
        ' Read corrections from year-specific REVIEW sheet
780     curRevName = REVIEW_SHEET_NAME() & "_" & yearStr
781     If SheetExists(curRevName) Then
790         Set wsRev = ThisWorkbook.Worksheets(curRevName)
800         revLastRow = wsRev.Cells(wsRev.Rows.Count, 1).End(xlUp).Row

            ' Find action, fix, and reason columns by scanning header row (row 2, since row 1 has button)
815         hdrActionText = ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) & ChrW(1492)
816         hdrFixText = ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1503)
            ' Hebrew: "sibat hriga" = reason header from helper dictionary
817         hdrReasonText = dictHelper(HELPER_REVIEW_REASON_HEADER)
818         actionColIdx = 0
819         fixColIdx = 0
820         reasonColIdx = 0
821         revLastCol = wsRev.Cells(2, wsRev.Columns.Count).End(xlToLeft).Column
822         For hdrCol = 1 To revLastCol + 2
823             If StrComp(Trim$(CStr(wsRev.Cells(2, hdrCol).Value2)), hdrActionText, vbTextCompare) = 0 Then actionColIdx = hdrCol
824             If StrComp(Trim$(CStr(wsRev.Cells(2, hdrCol).Value2)), hdrFixText, vbTextCompare) = 0 Then fixColIdx = hdrCol
825             If StrComp(Trim$(CStr(wsRev.Cells(2, hdrCol).Value2)), hdrReasonText, vbTextCompare) = 0 Then reasonColIdx = hdrCol
826         Next hdrCol
827         If actionColIdx = 0 Then actionColIdx = revLastCol - 1
828         If fixColIdx = 0 Then fixColIdx = revLastCol
829         If reasonColIdx = 0 Then reasonColIdx = actionColIdx - 1

830         If revLastRow >= 3 Then
840             For r = 3 To revLastRow
850                 reviewCount = reviewCount + 1
860                 srcRowNum = CLng(wsRev.Cells(r, 1).Value2)
870                 actionText = Trim$(CStr(wsRev.Cells(r, actionColIdx).Value2))
880                 fixText = Trim$(CStr(wsRev.Cells(r, fixColIdx).Value2))
890                 reasonText = Trim$(CStr(wsRev.Cells(r, reasonColIdx).Value2))

900                 If InStr(1, actionText, H_FIX(), vbTextCompare) > 0 Then
                        ' Store reason->fix pair in a sub-dictionary per source row
910                     If Not dictHasFix.Exists(CStr(srcRowNum)) Then
920                         Set dictRowFixes = CreateObject("Scripting.Dictionary")
930                         Set dictHasFix(CStr(srcRowNum)) = dictRowFixes
935                     Else
937                         Set dictRowFixes = dictHasFix(CStr(srcRowNum))
940                     End If
950                     dictRowFixes(reasonText) = fixText
960                     corrCount = corrCount + 1
                        ' --- Special handling: if reason is UNKNOWN_BRANCH, update settings A:D ---
                        Dim unknownBrText As String
                        unknownBrText = ""
                        If dictHelper.Exists("UNKNOWN_BRANCH") Then unknownBrText = dictHelper("UNKNOWN_BRANCH")
                        If unknownBrText = "" Then unknownBrText = H_BRANCH() & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1502) & ChrW(1494) & ChrW(1493) & ChrW(1492) & ChrW(1492)
                        If InStr(1, reasonText, unknownBrText, vbTextCompare) > 0 And fixText <> "" Then
                            ' fixText = the main branch name the user entered
                            ' Find the branch row in settings where col B is empty
                            Dim brFixRow As Long
                            Dim brFixLast As Long
                            brFixLast = wsMgmt.Cells(wsMgmt.Rows.Count, 1).End(xlUp).Row
                            For brFixRow = 3 To brFixLast
                                If Trim$(CStr(wsMgmt.Cells(brFixRow, 2).Value2)) = "" Then
                                    ' This is a branch without a main branch assigned
                                    ' Write the main branch
                                    wsMgmt.Cells(brFixRow, 2).Value = fixText
                                    ' Find MAIN_BRANCH_KEY from existing rows with same main branch
                                    Dim mbKeyRow As Long
                                    For mbKeyRow = 3 To brFixLast
                                        If StrComp(Trim$(CStr(wsMgmt.Cells(mbKeyRow, 2).Value2)), fixText, vbTextCompare) = 0 And mbKeyRow <> brFixRow Then
                                            If Trim$(CStr(wsMgmt.Cells(mbKeyRow, 4).Value2)) <> "" Then
                                                wsMgmt.Cells(brFixRow, 4).Value = wsMgmt.Cells(mbKeyRow, 4).Value2
                                                Exit For
                                            End If
                                        End If
                                    Next mbKeyRow
                                    ' If no existing key found, generate from fixText
                                    If Trim$(CStr(wsMgmt.Cells(brFixRow, 4).Value2)) = "" Then
                                        wsMgmt.Cells(brFixRow, 4).Value = HebrewToKey(fixText)
                                    End If
                                    Exit For
                                End If
                            Next brFixRow
                        End If
970                 ElseIf InStr(1, actionText, H_IGNORE(), vbTextCompare) > 0 Then
980                     dictHasIgnore(CStr(srcRowNum)) = True
990                     ignoreCount = ignoreCount + 1
                        ' "ha'aver livdika" = transfer for review - treat as ignore (pending response)
992                 ElseIf InStr(1, actionText, H_REVIEW(), vbTextCompare) > 0 Then
994                     dictHasIgnore(CStr(srcRowNum)) = True
996                     ignoreCount = ignoreCount + 1
1000                Else
1010                    dictHasUnhandled(CStr(srcRowNum)) = True
1020                    unhandledCount = unhandledCount + 1
1030                End If
1191            Next r

                ' Build final dictionaries: fix wins; if all ignore then ignore; else unhandled
1192            Set allKeys = CreateObject("Scripting.Dictionary")
1193            For Each srcKey In dictHasFix.keys: allKeys(srcKey) = True: Next
1194            For Each srcKey In dictHasIgnore.keys: allKeys(srcKey) = True: Next
1195            For Each srcKey In dictHasUnhandled.keys: allKeys(srcKey) = True: Next

1196            For Each srcKey In allKeys.keys
1197                If dictHasFix.Exists(srcKey) Then
1198                    Set dictCorrections(srcKey) = dictHasFix(srcKey)
1199                ElseIf dictHasUnhandled.Exists(srcKey) Then
                        ' has unhandled rows - do not ignore
1206                ElseIf dictHasIgnore.Exists(srcKey) Then
1207                    dictIgnore(srcKey) = True
1208                End If
1209            Next srcKey
1212        End If
1213    End If

        ' ---- Build base for current year ----

End Sub

' ============================================================================
' HELPER: BuildBaseSheet
' ============================================================================
Private Sub BuildBaseSheet(ByVal wsSrc As Worksheet, ByRef wsBase As Worksheet, ByVal baseSheetName As String, _
    ByVal dictIgnore As Object, ByVal dictCorrections As Object, ByVal dictRowFixes As Object, ByVal dictHasFix As Object, ByVal dictHasUnhandled As Object, _
    ByVal dictBranch As Object, _
    ByVal minMonth As Long, ByVal maxMonth As Long, ByVal filterCol As Long, ByVal filterVal As String, _
    ByVal dollarRate As Double, ByVal dateCol As Long, ByVal threshold As Double, ByVal wsMain As Worksheet, _
    ByRef outCount As Long)

    Dim lastRow As Long, outRow As Long, r As Long
    Dim premVal As Double, commVal As Double, currCode As Variant, rowBordMonth As String, rowRate As Double
    Dim dtStr As String, mPart As String, monthVal As Long
    Dim fixPrem As Double, fixComm As Double, fp As Long
    Dim fixParts() As String, oneFix As String
    Dim actionText As String, fixText As String
    
    DeleteSheetIfExists baseSheetName
    Set wsBase = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsBase.Name = baseSheetName
    
    wsBase.Cells(1, BASE_COL_ID).Value = "ID"
1160    wsBase.Cells(1, BASE_COL_YEAR).Value = "Year"
1170    wsBase.Cells(1, BASE_COL_MONTH).Value = "Month"
1180    wsBase.Cells(1, BASE_COL_IDENTITY).Value = "Identity"
1190    wsBase.Cells(1, BASE_COL_CUSTOMER).Value = "Customer"
1200    wsBase.Cells(1, BASE_COL_CUSTNAME).Value = "CustName"
1210    wsBase.Cells(1, BASE_COL_POLICY).Value = "Policy"
1220    wsBase.Cells(1, BASE_COL_ADDENDUM).Value = "Addendum"
1230    wsBase.Cells(1, BASE_COL_COMPANY).Value = "Company"
1240    wsBase.Cells(1, BASE_COL_COMPNUM).Value = "CompNum"
1250    wsBase.Cells(1, BASE_COL_BRANCHNAME).Value = "BranchName"
1260    wsBase.Cells(1, BASE_COL_BRANCHNUM).Value = "BranchNum"
1270    wsBase.Cells(1, BASE_COL_MAINBRANCH).Value = "MainBranch"
1280    wsBase.Cells(1, BASE_COL_AGENTNAME).Value = "AgentName"
1290    wsBase.Cells(1, BASE_COL_AGENTNUM).Value = "AgentNum"
1300    wsBase.Cells(1, BASE_COL_TELLER).Value = "Teller"
1310    wsBase.Cells(1, BASE_COL_TELLERNUM).Value = "TellerNum"
1320    wsBase.Cells(1, BASE_COL_ACTION).Value = "Action"
1330    wsBase.Cells(1, BASE_COL_PREMIUM).Value = "Premium"
1340    wsBase.Cells(1, BASE_COL_COMMISSION).Value = "Commission"
1350    wsBase.Cells(1, BASE_COL_ISSUE).Value = "Issue"
1360    wsBase.Cells(1, BASE_COL_TOFIX).Value = "ToFix"

1370    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
1380    outRow = 2

    Dim dictReasonToCol As Object
    Set dictReasonToCol = CreateObject("Scripting.Dictionary")
    dictReasonToCol(H_AGENT()) = BASE_COL_AGENTNAME
    dictReasonToCol(H_TELLER()) = BASE_COL_TELLER
    dictReasonToCol(H_COMPANY()) = BASE_COL_COMPANY
    dictReasonToCol(H_BRANCH()) = BASE_COL_BRANCHNAME
    
    Dim reasonFound As Boolean
    Dim rKey As Variant

        Dim srcData As Variant
        Dim baseData As Variant
        If lastRow >= 2 Then
            srcData = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, 50)).Value2
            ReDim baseData(1 To lastRow, 1 To BASE_COL_TOFIX)
        Else
            ReDim baseData(1 To 1, 1 To BASE_COL_TOFIX)
        End If

1390    For r = 2 To lastRow
            ' Skip ignored rows
1400        If dictIgnore.Exists(CStr(r)) Then GoTo NextSrcRow

            ' Get premium value, apply currency conversion, and check threshold
1410        premVal = 0
1412        commVal = 0
1414        currCode = srcData(r, RAW_CURRENCY)
            ' Per-row dollar rate by bordero month
1417        rowBordMonth = GetBordMonth(srcData(r, dateCol))
1418        If rowBordMonth <> "" Then
1419            rowRate = GetDollarRate(wsMain, rowBordMonth, True)
            Else
                rowRate = dollarRate
            End If
1420        If Not IsBlankValue(srcData(r, RAW_PREMIUM)) Then
1430            If TryParseVariantNumber(srcData(r, RAW_PREMIUM), premVal) Then
                    ' Currency conversion: 1=dollar (using per-month rate)
1432                If IsNumeric(currCode) Then
1434                    If CLng(currCode) = 1 Then premVal = premVal * rowRate
1436                End If
1440                If Abs(premVal) > threshold And Not dictCorrections.Exists(CStr(r)) Then GoTo NextSrcRow
1450            End If
1460        End If
            ' Convert commission too (using per-month rate)
1462        If Not IsBlankValue(srcData(r, RAW_COMMISSION)) Then
1464            If TryParseVariantNumber(srcData(r, RAW_COMMISSION), commVal) Then
1466                If IsNumeric(currCode) Then
1468                    If CLng(currCode) = 1 Then commVal = commVal * rowRate
1470                End If
1472            End If
1474        End If

            ' Extract month from bordereu date
1476        monthVal = 0
1480        bordereu = srcData(r, dateCol)
1490        If IsDate(bordereu) Then
1500            monthVal = Month(CDate(bordereu))
1505        ElseIf IsNumeric(bordereu) Then
                ' Value2 returns serial date number for Date cells
1506            If CDbl(bordereu) > 1 Then monthVal = Month(CDate(CDbl(bordereu)))
1510        ElseIf Not IsBlankValue(bordereu) Then
1530            dtStr = CStr(bordereu)
1540            If Len(dtStr) >= 7 Then
1560                mPart = Mid$(dtStr, 6, 2)
1570                If IsNumeric(mPart) Then monthVal = CInt(mPart)
1580            End If
1590        End If

            ' Write base row
1600        baseData(outRow - 1, BASE_COL_ID) = r
1610        baseData(outRow - 1, BASE_COL_YEAR) = yearVal
1620        baseData(outRow - 1, BASE_COL_MONTH) = monthVal
1630        baseData(outRow - 1, BASE_COL_IDENTITY) = srcData(r, RAW_IDNUMBER)
1640        baseData(outRow - 1, BASE_COL_CUSTOMER) = srcData(r, RAW_CUSTOMER)
1650        baseData(outRow - 1, BASE_COL_CUSTNAME) = srcData(r, RAW_CUSTNAME)
1660        baseData(outRow - 1, BASE_COL_POLICY) = srcData(r, RAW_POLICY)
1670        baseData(outRow - 1, BASE_COL_ADDENDUM) = srcData(r, RAW_ADDENDUM)
1680        baseData(outRow - 1, BASE_COL_COMPANY) = srcData(r, RAW_COMPANY)
1690        baseData(outRow - 1, BASE_COL_COMPNUM) = srcData(r, RAW_COMPNUM)
1700        baseData(outRow - 1, BASE_COL_BRANCHNAME) = srcData(r, RAW_BRANCHNAME)
1710        baseData(outRow - 1, BASE_COL_BRANCHNUM) = srcData(r, RAW_BRANCHNUM)

            ' Main branch mapping
1720        Dim brKey As String
1730        brKey = UCase$(Trim$(CStr(srcData(r, RAW_BRANCHNAME))))
1740        If dictBranch.Exists(brKey) Then
1750            baseData(outRow - 1, BASE_COL_MAINBRANCH) = dictBranch(brKey)
1760        Else
1770            baseData(outRow - 1, BASE_COL_MAINBRANCH) = srcData(r, RAW_BRANCHNAME)
1780        End If

1790        baseData(outRow - 1, BASE_COL_AGENTNAME) = srcData(r, RAW_AGENTNAME)
1800        baseData(outRow - 1, BASE_COL_AGENTNUM) = srcData(r, RAW_AGENTNUM)
1810        baseData(outRow - 1, BASE_COL_TELLER) = srcData(r, RAW_TELLERNAME)
1820        baseData(outRow - 1, BASE_COL_TELLERNUM) = srcData(r, RAW_TELLERNUM)
1830        baseData(outRow - 1, BASE_COL_ACTION) = srcData(r, RAW_ACTIONCOL)
1840        baseData(outRow - 1, BASE_COL_PREMIUM) = premVal
1850        baseData(outRow - 1, BASE_COL_COMMISSION) = commVal

            ' Apply corrections using reason-based mapping from REVIEW
1860        If dictCorrections.Exists(CStr(r)) Then
1870            baseData(outRow - 1, BASE_COL_ISSUE) = "CORRECTED"
                ' Iterate each reason->fix pair and apply to the correct column
1872            Set dictRowFixes = dictCorrections(CStr(r))
1880            For Each fixKey In dictRowFixes.keys
1882                reasonText = CStr(fixKey)
1884                oneFix = Trim$(CStr(dictRowFixes(fixKey)))
1886                If oneFix <> "" Then
                    reasonFound = False
                    For Each rKey In dictReasonToCol.keys
                        If InStr(1, reasonText, rKey, vbTextCompare) > 0 Then
                            baseData(outRow - 1, dictReasonToCol(rKey)) = oneFix
                            reasonFound = True
                            Exit For
                        End If
                    Next rKey
                    
                    If Not reasonFound Then
                        If InStr(1, reasonText, H_PREMIUM(), vbTextCompare) > 0 Then
1906                        If TryParseVariantNumber(oneFix, fixPrem) Then
1908                            baseData(outRow - 1, BASE_COL_PREMIUM) = fixPrem
1910                            premVal = fixPrem
1912                        End If
                        ElseIf InStr(1, reasonText, H_COMMISSION(), vbTextCompare) > 0 Then
1916                        If TryParseVariantNumber(oneFix, fixComm) Then
1918                            baseData(outRow - 1, BASE_COL_COMMISSION) = fixComm
1920                        End If
                        End If
                    End If
1924                End If
1926            Next fixKey
1928        End If

1932        outRow = outRow + 1

NextSrcRow:
1934    Next r

        If outRow > 2 Then wsBase.Range(wsBase.Cells(2, 1), wsBase.Cells(outRow - 1, BASE_COL_TOFIX)).Value = baseData
1940    outCount = outRow - 2
        wsBase.Columns.AutoFit

        ' ---- Load or build base for reference year ----

End Sub

Private Sub BuildComparisonSheet(ByVal wsCurrent As Worksheet, ByVal wsRef As Worksheet, ByVal sheetName As String, ByVal groupCol As Long, ByVal minMonth As Long, ByVal maxMonth As Long, ByVal yearVal As String, ByVal refYear As String, Optional ByVal filterCol As Long = 0, Optional ByVal filterVal As String = "", Optional ByVal titleText As String = "", Optional ByVal clientFilterVal As String = "")

10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False

20      DeleteSheetIfExists sheetName
30      Dim wsOut As Worksheet
40      Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
50      wsOut.Name = sheetName

        ' Collect unique keys from both sheets
60      Dim dictKeys As Object
70      Set dictKeys = CreateObject("Scripting.Dictionary")
80      dictKeys.CompareMode = vbTextCompare

90      Dim lastRowCur As Long
100     Dim lastRowRef As Long
110     Dim r As Long
120     Dim k As String
130     Dim m As Long
        Dim sortedMonths() As String
        Dim mIdx As Long
        Dim mCnt As Long
        Dim mi As Long
        Dim dictGlobalCustCur As Object
        Dim dictGlobalCustRef As Object
        Dim dictGlobalPolCur As Object
        Dim dictGlobalPolRef As Object

140     lastRowCur = wsCurrent.Cells(wsCurrent.Rows.Count, 1).End(xlUp).Row
150     lastRowRef = wsRef.Cells(wsRef.Rows.Count, 1).End(xlUp).Row

        Dim curData As Variant
        If lastRowCur >= 2 Then curData = wsCurrent.Range(wsCurrent.Cells(1, 1), wsCurrent.Cells(lastRowCur, 22)).Value2
        Dim refDataComp As Variant
        If lastRowRef >= 2 Then refDataComp = wsRef.Range(wsRef.Cells(1, 1), wsRef.Cells(lastRowRef, 22)).Value2

160     For r = 2 To lastRowCur
170         m = 0
180         If Not IsBlankValue(curData(r, BASE_COL_MONTH)) Then
190             m = SafeLong(curData(r, BASE_COL_MONTH))
200         End If
210         If m >= minMonth And m <= maxMonth Then
                ' Exclude MainBranch = "????" from all reports
211             If StrComp(Trim$(SafeString(curData(r, BASE_COL_MAINBRANCH))), ChrW(1495) & ChrW(1493) & ChrW(1489) & ChrW(1492), vbTextCompare) = 0 Then GoTo SkipCurKey
                ' Cross-filter: if filterCol is set, only include rows matching filterVal
215             If filterCol > 0 And filterVal <> "" Then
216                 If StrComp(Trim$(SafeString(curData(r, filterCol))), filterVal, vbTextCompare) <> 0 Then GoTo SkipCurKey
217             End If
                ' Client filter: if clientFilterVal is set, only include rows matching client name
218             If clientFilterVal <> "" Then
219                 If StrComp(Trim$(SafeString(curData(r, BASE_COL_CUSTNAME))), clientFilterVal, vbTextCompare) <> 0 Then GoTo SkipCurKey
220             End If
221             k = Trim$(SafeString(curData(r, groupCol)))
230             If k <> "" And LCase$(k) <> "(empty)" Then
240                 If Not dictKeys.Exists(k) Then dictKeys(k) = True
250             End If
260         End If
SkipCurKey:
270     Next r

280     For r = 2 To lastRowRef
290         m = 0
300         If Not IsBlankValue(refDataComp(r, BASE_COL_MONTH)) Then
310             m = SafeLong(refDataComp(r, BASE_COL_MONTH))
320         End If
330         If m >= minMonth And m <= maxMonth Then
                ' Exclude MainBranch = "????" from all reports
331             If StrComp(Trim$(SafeString(refDataComp(r, BASE_COL_MAINBRANCH))), ChrW(1495) & ChrW(1493) & ChrW(1489) & ChrW(1492), vbTextCompare) = 0 Then GoTo SkipRefKey
                ' Cross-filter: if filterCol is set, only include rows matching filterVal
335             If filterCol > 0 And filterVal <> "" Then
336                 If StrComp(Trim$(SafeString(refDataComp(r, filterCol))), filterVal, vbTextCompare) <> 0 Then GoTo SkipRefKey
337             End If
                ' Client filter: if clientFilterVal is set, only include rows matching client name
338             If clientFilterVal <> "" Then
339                 If StrComp(Trim$(SafeString(refDataComp(r, BASE_COL_CUSTNAME))), clientFilterVal, vbTextCompare) <> 0 Then GoTo SkipRefKey
340             End If
341             k = Trim$(SafeString(refDataComp(r, groupCol)))
350             If k <> "" And LCase$(k) <> "(empty)" Then
360                 If Not dictKeys.Exists(k) Then dictKeys(k) = True
370             End If
380         End If
SkipRefKey:
390     Next r

        ' Write headers
400     WriteComparisonHeaders wsOut, yearVal, refYear, sheetName

        ' For each key, aggregate values
410     Dim outRow As Long
420     outRow = 3
430     Dim allKeys As Variant

        ' Sort keys for months sheet (numeric 1-12)
        If groupCol = BASE_COL_MONTH Then
            mCnt = 0
            ReDim sortedMonths(1 To 12)
            For mi = 1 To 12
                If dictKeys.Exists(CStr(mi)) Then
                    mCnt = mCnt + 1
                    sortedMonths(mCnt) = CStr(mi)
                End If
            Next mi
            If mCnt > 0 Then
                ReDim Preserve sortedMonths(1 To mCnt)
                ReDim allKeys(0 To mCnt - 1)
                For mi = 1 To mCnt
                    allKeys(mi - 1) = sortedMonths(mi)
                Next mi
            Else
                allKeys = dictKeys.keys
            End If
        Else
440         allKeys = dictKeys.keys
        End If

450     Dim idx As Long
460     Dim premCur As Double
470     Dim premRef As Double
480     Dim commCur As Double
490     Dim commRef As Double
500     Dim docsCur As Long
510     Dim docsRef As Long
520     Dim custCur As Long
530     Dim custRef As Long
540     Dim polCur As Long
550     Dim polRef As Long
560     Dim dictCustCur As Object
570     Dim dictCustRef As Object
580     Dim dictPolCur As Object
590     Dim dictPolRef As Object
600     Dim custKey As String
610     Dim polKey As String
615     Dim actValCur As String
616     Dim actValRef As String
617     Dim gKey As Variant

620     Dim totPremCur As Double
630     Dim totPremRef As Double
640     Dim totCommCur As Double
650     Dim totCommRef As Double
660     Dim totDocsCur As Long
670     Dim totDocsRef As Long
680     Dim totCustCur As Long
690     Dim totCustRef As Long
700     Dim totPolCur As Long
710     Dim totPolRef As Long

        ' Global unique dicts for correct totals
        Set dictGlobalCustCur = CreateObject("Scripting.Dictionary")
        Set dictGlobalCustRef = CreateObject("Scripting.Dictionary")
        Set dictGlobalPolCur = CreateObject("Scripting.Dictionary")
        Set dictGlobalPolRef = CreateObject("Scripting.Dictionary")

720     For idx = 0 To UBound(allKeys)
730         k = allKeys(idx)

740         premCur = 0: premRef = 0
750         commCur = 0: commRef = 0
760         docsCur = 0: docsRef = 0

770         Set dictCustCur = CreateObject("Scripting.Dictionary")
780         Set dictCustRef = CreateObject("Scripting.Dictionary")
790         Set dictPolCur = CreateObject("Scripting.Dictionary")
800         Set dictPolRef = CreateObject("Scripting.Dictionary")

            ' Aggregate current year
810         For r = 2 To lastRowCur
820             m = 0
830             If Not IsBlankValue(curData(r, BASE_COL_MONTH)) Then
840                 m = SafeLong(curData(r, BASE_COL_MONTH))
850             End If
860             If m < minMonth Or m > maxMonth Then GoTo NextCurRow
                ' Exclude MainBranch = "????" from all reports
861             If StrComp(Trim$(SafeString(curData(r, BASE_COL_MAINBRANCH))), ChrW(1495) & ChrW(1493) & ChrW(1489) & ChrW(1492), vbTextCompare) = 0 Then GoTo NextCurRow
                ' Cross-filter: skip rows not matching the filter
862             If filterCol > 0 And filterVal <> "" Then
864                 If StrComp(Trim$(SafeString(curData(r, filterCol))), filterVal, vbTextCompare) <> 0 Then GoTo NextCurRow
866             End If
                ' Client filter
867             If clientFilterVal <> "" Then
868                 If StrComp(Trim$(SafeString(curData(r, BASE_COL_CUSTNAME))), clientFilterVal, vbTextCompare) <> 0 Then GoTo NextCurRow
869             End If

870             If StrComp(Trim$(SafeString(curData(r, groupCol))), k, vbTextCompare) = 0 Then
880                 premCur = premCur + SafeDouble(curData(r, BASE_COL_PREMIUM))
890                 commCur = commCur + SafeDouble(curData(r, BASE_COL_COMMISSION))
900                 docsCur = docsCur + 1

                    ' Customer counting: unique customer numbers with at least one non-cancelled row
910                 custKey = Trim$(SafeString(curData(r, BASE_COL_CUSTOMER)))
920                 If custKey <> "" Then
                        actValCur = Trim$(SafeString(curData(r, BASE_COL_ACTION)))
                        ' Check if this row is NOT a cancellation
                        If InStr(1, actValCur, ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1500), vbTextCompare) = 0 Then
930                         If Not dictCustCur.Exists(custKey) Then dictCustCur(custKey) = True
                        End If
940                 End If

950                 polKey = Trim$(SafeString(curData(r, BASE_COL_POLICY)))
960                 If polKey <> "" Then
970                     If Not dictPolCur.Exists(polKey) Then dictPolCur(polKey) = True
980                 End If
990             End If
NextCurRow:
1000        Next r

            ' Aggregate reference year
1010        For r = 2 To lastRowRef
1020            m = 0
1030            If Not IsBlankValue(refDataComp(r, BASE_COL_MONTH)) Then
1040                m = SafeLong(refDataComp(r, BASE_COL_MONTH))
1050            End If
1060            If m < minMonth Or m > maxMonth Then GoTo NextRefRow2
                ' Exclude MainBranch = "????" from all reports
1061            If StrComp(Trim$(SafeString(refDataComp(r, BASE_COL_MAINBRANCH))), ChrW(1495) & ChrW(1493) & ChrW(1489) & ChrW(1492), vbTextCompare) = 0 Then GoTo NextRefRow2
                ' Cross-filter: skip rows not matching the filter
1062            If filterCol > 0 And filterVal <> "" Then
1064                If StrComp(Trim$(SafeString(refDataComp(r, filterCol))), filterVal, vbTextCompare) <> 0 Then GoTo NextRefRow2
1066            End If
                ' Client filter
1067            If clientFilterVal <> "" Then
1068                If StrComp(Trim$(SafeString(refDataComp(r, BASE_COL_CUSTNAME))), clientFilterVal, vbTextCompare) <> 0 Then GoTo NextRefRow2
1069            End If

1070            If StrComp(Trim$(SafeString(refDataComp(r, groupCol))), k, vbTextCompare) = 0 Then
1080                premRef = premRef + SafeDouble(refDataComp(r, BASE_COL_PREMIUM))
1090                commRef = commRef + SafeDouble(refDataComp(r, BASE_COL_COMMISSION))
1100                docsRef = docsRef + 1

                    ' Customer counting: unique customer numbers with at least one non-cancelled row
1110                custKey = Trim$(SafeString(refDataComp(r, BASE_COL_CUSTOMER)))
1120                If custKey <> "" Then
                        actValRef = Trim$(SafeString(refDataComp(r, BASE_COL_ACTION)))
                        ' Check if this row is NOT a cancellation
                        If InStr(1, actValRef, ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1500), vbTextCompare) = 0 Then
1130                        If Not dictCustRef.Exists(custKey) Then dictCustRef(custKey) = True
                        End If
1140                End If

1150                polKey = Trim$(SafeString(refDataComp(r, BASE_COL_POLICY)))
1160                If polKey <> "" Then
1170                    If Not dictPolRef.Exists(polKey) Then dictPolRef(polKey) = True
1180                End If
1190            End If
NextRefRow2:
1200        Next r

1210        custCur = dictCustCur.Count
1220        custRef = dictCustRef.Count
1230        polCur = dictPolCur.Count
1240        polRef = dictPolRef.Count

            ' Write row - for months, show Hebrew name
            If groupCol = BASE_COL_MONTH Then
1250            wsOut.Cells(outRow, 1).Value = HebrewMonthName(CLng(k))
            Else
1251            wsOut.Cells(outRow, 1).Value = k
            End If
1260        wsOut.Cells(outRow, 2).Value = premRef
1270        wsOut.Cells(outRow, 3).Value = premCur
1280        wsOut.Cells(outRow, 4).Value = SafePct(premCur, premRef)
1290        wsOut.Cells(outRow, 5).Value = docsRef
1300        wsOut.Cells(outRow, 6).Value = docsCur
1310        wsOut.Cells(outRow, 7).Value = SafePct(docsCur, docsRef)
1320        wsOut.Cells(outRow, 8).Value = custRef
1330        wsOut.Cells(outRow, 9).Value = custCur
1340        wsOut.Cells(outRow, 10).Value = SafePct(custCur, custRef)
1350        wsOut.Cells(outRow, 11).Value = polRef
1360        wsOut.Cells(outRow, 12).Value = polCur
1370        wsOut.Cells(outRow, 13).Value = SafePct(polCur, polRef)
1380        wsOut.Cells(outRow, 14).Value = commRef
1390        wsOut.Cells(outRow, 15).Value = commCur
1400        wsOut.Cells(outRow, 16).Value = SafePct(commCur, commRef)

            ' Accumulate totals (premium, commission, docs are additive)
1410        totPremCur = totPremCur + premCur
1420        totPremRef = totPremRef + premRef
1430        totCommCur = totCommCur + commCur
1440        totCommRef = totCommRef + commRef
1450        totDocsCur = totDocsCur + docsCur
1460        totDocsRef = totDocsRef + docsRef

            ' Merge per-group unique dicts into global dicts
            For Each gKey In dictCustCur.keys
                If Not dictGlobalCustCur.Exists(gKey) Then dictGlobalCustCur(gKey) = True
            Next gKey
            For Each gKey In dictCustRef.keys
                If Not dictGlobalCustRef.Exists(gKey) Then dictGlobalCustRef(gKey) = True
            Next gKey
            For Each gKey In dictPolCur.keys
                If Not dictGlobalPolCur.Exists(gKey) Then dictGlobalPolCur(gKey) = True
            Next gKey
            For Each gKey In dictPolRef.keys
                If Not dictGlobalPolRef.Exists(gKey) Then dictGlobalPolRef(gKey) = True
            Next gKey

1510        outRow = outRow + 1
1520    Next idx

        ' Compute global unique totals for customers and policies
        totCustCur = dictGlobalCustCur.Count
        totCustRef = dictGlobalCustRef.Count
        totPolCur = dictGlobalPolCur.Count
        totPolRef = dictGlobalPolRef.Count

        ' Write totals row
1530    wsOut.Cells(outRow, 1).Value = ChrW(1505) & ChrW(1492) & Chr(34) & ChrW(1499)
1540    wsOut.Cells(outRow, 2).Value = totPremRef
1550    wsOut.Cells(outRow, 3).Value = totPremCur
1560    wsOut.Cells(outRow, 4).Value = SafePct(totPremCur, totPremRef)
1570    wsOut.Cells(outRow, 5).Value = totDocsRef
1580    wsOut.Cells(outRow, 6).Value = totDocsCur
1590    wsOut.Cells(outRow, 7).Value = SafePct(totDocsCur, totDocsRef)
1600    wsOut.Cells(outRow, 8).Value = totCustRef
1610    wsOut.Cells(outRow, 9).Value = totCustCur
1620    wsOut.Cells(outRow, 10).Value = SafePct(totCustCur, totCustRef)
1630    wsOut.Cells(outRow, 11).Value = totPolRef
1640    wsOut.Cells(outRow, 12).Value = totPolCur
1650    wsOut.Cells(outRow, 13).Value = SafePct(totPolCur, totPolRef)
1660    wsOut.Cells(outRow, 14).Value = totCommRef
1670    wsOut.Cells(outRow, 15).Value = totCommCur
1680    wsOut.Cells(outRow, 16).Value = SafePct(totCommCur, totCommRef)
1690    wsOut.Rows(outRow).Font.Bold = True

        ' Format
1700    wsOut.Columns.AutoFit
1710    Dim col As Long
1720    For col = 2 To 16
1730        If col = 4 Or col = 7 Or col = 10 Or col = 13 Or col = 16 Then
1740            wsOut.Columns(col).NumberFormat = "0.0%"
1750        Else
1760            wsOut.Columns(col).NumberFormat = "#,##0"
1770        End If
1780    Next col

        ' --- Insert title row at top ---
        If titleText <> "" Then
            wsOut.Rows("1:1").Insert Shift:=xlDown
            wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, 16)).Merge
            wsOut.Range("A1").Value = titleText
            wsOut.Range("A1").Font.Bold = True
            wsOut.Range("A1").Font.Size = 14
            wsOut.Range("A1").Font.Color = RGB(0, 70, 140)
            wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, 16)).Interior.Color = RGB(255, 228, 225)  ' pastel pink
            wsOut.Range("A1").HorizontalAlignment = xlCenter
        End If

        ' ---- Zebra striping will be applied after Tab.Color is set (see ApplyZebraStriping) ----

1790    Exit Sub

ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
1800    Err.Raise Err.Number, "BuildComparisonSheet(" & sheetName & "):" & Erl, Err.Description
End Sub

' ============================================================================
' HELPER: Write comparison sheet headers (two-row merged layout)
' Row 1: category headers (merged, blue background, white bold text)
' Row 2: year sub-headers (light blue background, bold text)
' Data starts at row 3
' ============================================================================
Private Sub WriteComparisonHeaders(ByVal ws As Worksheet, ByVal yearVal As String, ByVal refYear As String, ByVal sheetName As String)
10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False

        Dim blueColor As Long
        Dim lightBlueColor As Long
        Dim pctLabel As String
20      blueColor = RGB(0, 70, 140)
30      lightBlueColor = RGB(155, 200, 235)

        ' --- Row 1: Category headers (merged cells) ---
        ' Col A: name header (merged rows 1-2)
40      ws.Range("A1:A2").Merge
50      ws.Cells(1, 1).Value = ChrW(1513) & ChrW(1501)
60      ws.Cells(1, 1).HorizontalAlignment = xlCenter
70      ws.Cells(1, 1).VerticalAlignment = xlCenter

        ' Cols B-D: production (merged)
80      ws.Range(ws.Cells(1, 2), ws.Cells(1, 4)).Merge
90      ws.Cells(1, 2).Value = ChrW(1508) & ChrW(1512) & ChrW(1493) & ChrW(1491) & ChrW(1493) & ChrW(1511) & ChrW(1510) & ChrW(1497) & ChrW(1492)
100     ws.Cells(1, 2).HorizontalAlignment = xlCenter

        ' Cols E-G: documents (merged)
110     ws.Range(ws.Cells(1, 5), ws.Cells(1, 7)).Merge
120     ws.Cells(1, 5).Value = ChrW(1502) & ChrW(1505) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1501)
130     ws.Cells(1, 5).HorizontalAlignment = xlCenter

        ' Cols H-J: insured (merged)
140     ws.Range(ws.Cells(1, 8), ws.Cells(1, 10)).Merge
150     ws.Cells(1, 8).Value = ChrW(1502) & ChrW(1489) & ChrW(1493) & ChrW(1496) & ChrW(1495) & ChrW(1497) & ChrW(1501)
160     ws.Cells(1, 8).HorizontalAlignment = xlCenter

        ' Cols K-M: policies (merged)
170     ws.Range(ws.Cells(1, 11), ws.Cells(1, 13)).Merge
180     ws.Cells(1, 11).Value = ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1505) & ChrW(1493) & ChrW(1514)
190     ws.Cells(1, 11).HorizontalAlignment = xlCenter

        ' Cols N-P: commission (merged)
200     ws.Range(ws.Cells(1, 14), ws.Cells(1, 16)).Merge
210     ws.Cells(1, 14).Value = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1492)
220     ws.Cells(1, 14).HorizontalAlignment = xlCenter

        ' --- Row 1 formatting: blue background, white bold ---
230     ws.Range(ws.Cells(1, 1), ws.Cells(1, 16)).Interior.Color = blueColor
240     ws.Range(ws.Cells(1, 1), ws.Cells(1, 16)).Font.Color = RGB(255, 255, 255)
250     ws.Range(ws.Cells(1, 1), ws.Cells(1, 16)).Font.Bold = True
260     ws.Range(ws.Cells(1, 1), ws.Cells(1, 16)).Font.Size = 12

        ' --- Row 2: Year sub-headers ---
        ' Hebrew: shinuy% = change%
270     pctLabel = ChrW(1513) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1497) & "%"

        ' Force row 2 to Text format BEFORE writing year values (prevents 2,025 display)
        ws.Range(ws.Cells(2, 1), ws.Cells(2, 16)).NumberFormat = "@"

280     ws.Cells(2, 1).Value = ""
290     ws.Cells(2, 2).Value = refYear
300     ws.Cells(2, 3).Value = yearVal
310     ws.Cells(2, 4).Value = pctLabel
320     ws.Cells(2, 5).Value = refYear
330     ws.Cells(2, 6).Value = yearVal
340     ws.Cells(2, 7).Value = pctLabel
350     ws.Cells(2, 8).Value = refYear
360     ws.Cells(2, 9).Value = yearVal
370     ws.Cells(2, 10).Value = pctLabel
380     ws.Cells(2, 11).Value = refYear
390     ws.Cells(2, 12).Value = yearVal
400     ws.Cells(2, 13).Value = pctLabel
410     ws.Cells(2, 14).Value = refYear
420     ws.Cells(2, 15).Value = yearVal
430     ws.Cells(2, 16).Value = pctLabel

        ' --- Row 2 formatting: light blue background, bold ---
440     ws.Range(ws.Cells(2, 1), ws.Cells(2, 16)).Interior.Color = lightBlueColor
450     ws.Range(ws.Cells(2, 1), ws.Cells(2, 16)).Font.Bold = True
460     ws.Range(ws.Cells(2, 1), ws.Cells(2, 16)).HorizontalAlignment = xlCenter

        ' --- Borders ---
470     ws.Range(ws.Cells(1, 1), ws.Cells(2, 16)).Borders.LineStyle = xlContinuous
480     ws.Range(ws.Cells(1, 1), ws.Cells(2, 16)).Borders.Weight = xlThin

490     Exit Sub

ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
500     Err.Raise Err.Number, "WriteComparisonHeaders:" & Erl, Err.Description
End Sub

' ============================================================================
' HELPER: Build summary sheet
' ============================================================================
Private Sub BuildSummarySheet(ByVal countRef As Long, ByVal countCurrent As Long, ByVal reviewCount As Long, ByVal corrCount As Long, ByVal ignoreCount As Long, ByVal unhandledCount As Long, ByVal yearVal As String, ByVal refYear As String, ByVal periodDesc As String, Optional ByVal titleText As String = "")

10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False
20      DeleteSheetIfExists SHEET_SUMMARY()
30      Dim wsOut As Worksheet
40      Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
50      wsOut.Name = SHEET_SUMMARY()

60      wsOut.Cells(1, 1).Value = ChrW(1505) & ChrW(1497) & ChrW(1499) & ChrW(1493) & ChrW(1501) & " " & ChrW(1514) & ChrW(1492) & ChrW(1500) & ChrW(1497) & ChrW(1498)
70      wsOut.Cells(1, 1).Font.Bold = True
80      wsOut.Cells(1, 1).Font.Size = 14

90      wsOut.Cells(3, 1).Value = ChrW(1508) & ChrW(1512) & ChrW(1496)
100     wsOut.Cells(3, 2).Value = ChrW(1506) & ChrW(1512) & ChrW(1498)
110     wsOut.Rows(3).Font.Bold = True

200     wsOut.Cells(4, 1).Value = ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
210     wsOut.Cells(4, 2).Value = periodDesc
220     wsOut.Cells(5, 1).Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & ChrW(1497) & ChrW(1497) & ChrW(1495) & ChrW(1493) & ChrW(1505)
230     wsOut.Cells(5, 2).Value = refYear
240     wsOut.Cells(6, 1).Value = ChrW(1513) & ChrW(1504) & ChrW(1492) & " " & ChrW(1504) & ChrW(1489) & ChrW(1491) & ChrW(1511) & ChrW(1514)
250     wsOut.Cells(6, 2).Value = yearVal
260     wsOut.Cells(7, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & refYear
270     wsOut.Cells(7, 2).Value = countRef
280     wsOut.Cells(8, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & yearVal & " (" & ChrW(1502) & ChrW(1514) & ChrW(1493) & ChrW(1511) & ChrW(1503) & ")"
290     wsOut.Cells(8, 2).Value = countCurrent
300     wsOut.Cells(9, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1489) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1512) & ChrW(1514)
310     wsOut.Cells(9, 2).Value = reviewCount
320     wsOut.Cells(10, 1).Value = ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1513) & ChrW(1497) & ChrW(1493) & ChrW(1513) & ChrW(1502) & ChrW(1493)
330     wsOut.Cells(10, 2).Value = corrCount
340     wsOut.Cells(11, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1513) & ChrW(1492) & ChrW(1493) & ChrW(1514) & ChrW(1506) & ChrW(1500) & ChrW(1502) & ChrW(1493)
350     wsOut.Cells(11, 2).Value = ignoreCount
360     wsOut.Cells(12, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1500) & ChrW(1488) & " " & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500)
370     wsOut.Cells(12, 2).Value = unhandledCount

380     wsOut.Columns(2).NumberFormat = "#,##0"
390     wsOut.Columns.AutoFit
        wsOut.DisplayRightToLeft = True
410     Exit Sub

ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
420     Err.Raise Err.Number, "BuildSummarySheet:" & Erl, Err.Description
End Sub

' ============================================================================
' HELPER: Safe percentage calculation
' ============================================================================
Private Function SafePct(ByVal newVal As Double, ByVal oldVal As Double) As Double
10      If oldVal = 0 Then
20          If newVal = 0 Then
30              SafePct = 0
40          Else
50              SafePct = 1
60          End If
70      Else
80          SafePct = (newVal - oldVal) / Abs(oldVal)
90      End If
End Function

' ============================================================================
' HELPER: Load reason codes dictionary from NIHUL (section rngSection_ReasonCode)
' ============================================================================
Private Function LoadHelperDictionary(ByVal ws As Worksheet) As Object
10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False
20      Dim dict As Object
30      Set dict = CreateObject("Scripting.Dictionary")
40      dict.CompareMode = vbTextCompare
50      Dim r As Long
60      Dim startRow As Long
70      Dim k As String
        ' Start from row after section header (rngSection_ReasonCode)
80      startRow = ThisWorkbook.Names("rngSection_ReasonCode").RefersToRange.Row + 1
90      r = startRow
100     Do While True
110         k = Trim$(CStr(ws.Cells(r, COL_HELPER_KEY).Value2))
120         If UCase$(k) = EOD_MARKER Or k = "" Then Exit Do
130         dict(k) = Trim$(CStr(ws.Cells(r, COL_HELPER_VALUE).Value2))
140         r = r + 1
145     Loop
        ' Add fallback translations for missing codes
150     If Not dict.Exists("MISSING_CUSTOMER_NUMBER") Then dict("MISSING_CUSTOMER_NUMBER") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)
152     If Not dict.Exists("MISSING_CUSTOMER_NAME") Then dict("MISSING_CUSTOMER_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)
154     If Not dict.Exists("MISSING_POLICY") Then dict("MISSING_POLICY") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1505) & ChrW(1492)
156     If Not dict.Exists("MISSING_ADDENDUM") Then dict("MISSING_ADDENDUM") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1514) & ChrW(1493) & ChrW(1505) & ChrW(1508) & ChrW(1514)
158     If Not dict.Exists("MISSING_COMPANY_NAME") Then dict("MISSING_COMPANY_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & H_COMPANY()
160     If Not dict.Exists("MISSING_BRANCH_NAME") Then dict("MISSING_BRANCH_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & H_BRANCH()
162     If Not dict.Exists("MISSING_AGENT_NAME") Then dict("MISSING_AGENT_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & H_AGENT()
164     If Not dict.Exists("MISSING_UNDERWRITER_TELLER_NAME") Then dict("MISSING_UNDERWRITER_TELLER_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & H_TELLER()
166     If Not dict.Exists("MISSING_CURRENCY") Then dict("MISSING_CURRENCY") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1502) & ChrW(1496) & ChrW(1489) & ChrW(1506)
168     If Not dict.Exists("MISSING_PREMIUM") Then dict("MISSING_PREMIUM") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & H_PREMIUM()
170     If Not dict.Exists("MISSING_COMPANY_COMMISSION") Then dict("MISSING_COMPANY_COMMISSION") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & H_COMMISSION() & " " & H_COMPANY()
172     If Not dict.Exists("PREMIUM_OVER_THRESHOLD") Then dict("PREMIUM_OVER_THRESHOLD") = H_PREMIUM() & " " & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1492)
174     If Not dict.Exists("PREMIUM_NOT_NUMERIC") Then dict("PREMIUM_NOT_NUMERIC") = ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & H_PREMIUM() & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & ChrW(1497)

180     Set LoadHelperDictionary = dict
190     Exit Function
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
200     Err.Raise Err.Number, "LoadHelperDictionary:" & Erl, Err.Description
End Function

' ============================================================================
' HELPER: Load branch mapping from NIHUL
' ============================================================================
Private Function LoadBranchMapping(ByVal ws As Worksheet) As Object
10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False
20      Dim dict As Object
30      Set dict = CreateObject("Scripting.Dictionary")
40      dict.CompareMode = vbTextCompare
50      Dim r As Long
60      Dim lastRow As Long
70      Dim brName As String
80      Dim mainBr As String
90      lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
100     For r = 3 To lastRow
110         brName = UCase$(Trim$(CStr(ws.Cells(r, 1).Value2)))
120         mainBr = Trim$(CStr(ws.Cells(r, 2).Value2))
130         If brName <> "" And mainBr <> "" Then
140             dict(brName) = mainBr
150         End If
160     Next r
170     Set LoadBranchMapping = dict
180     Exit Function
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
190     Err.Raise Err.Number, "LoadBranchMapping:" & Erl, Err.Description
End Function

' ============================================================================
' HELPER: Validate helper key exists
' ============================================================================
Private Sub ValidateHelperKey(ByVal dict As Object, ByVal key As String)
10      If Not dict.Exists(key) Then
20          Err.Raise vbObjectError + 2000, "ValidateHelperKey", "HELPER KEY NOT FOUND: " & key
30      End If
End Sub

' ============================================================================
' HELPER: Load checked fields from NIHUL
' ============================================================================
Private Sub LoadCheckedFields(ByVal ws As Worksheet, ByVal dictCol As Object, ByVal dictDisp As Object)
10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False
20      Dim r As Long
30      Dim startRow As Long
40      Dim fName As String
50      Dim fCol As String
60      Dim fCheck As String
70      Dim fKey As String
        ' Start from row after section header (rngSection_FieldMap)
80      startRow = ThisWorkbook.Names("rngSection_FieldMap").RefersToRange.Row + 1
90      r = startRow
100     Do While True
110         fName = Trim$(CStr(ws.Cells(r, COL_FIELD_NAME_HE).Value2))
120         If UCase$(fName) = EOD_MARKER Or fName = "" Then Exit Do
130         fCol = UCase$(Trim$(CStr(ws.Cells(r, COL_FIELD_COLUMN).Value2)))
140         fCheck = UCase$(Trim$(CStr(ws.Cells(r, COL_FIELD_CHECKING).Value2)))
150         fKey = Trim$(CStr(ws.Cells(r, COL_FIELD_KEY).Value2))
160         If fKey <> "" And fCol <> "" And fCheck = "CHECK" Then
170             dictCol(fKey) = ColumnLetterToNumber(fCol)
180             dictDisp(fKey) = fName
190         End If
200         r = r + 1
210     Loop
220     Exit Sub
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
230     Err.Raise Err.Number, "LoadCheckedFields:" & Erl, Err.Description
End Sub

' ============================================================================
' HELPER: Build arrays from dictionaries
' ============================================================================
Private Sub BuildArrays(ByVal dictCol As Object, ByVal dictDisp As Object, ByRef keys() As String, ByRef cols() As Long, ByRef disp() As String, ByRef cnt As Long)
10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False
20      cnt = dictCol.Count
30      If cnt = 0 Then Exit Sub
40      ReDim keys(1 To cnt)
50      ReDim cols(1 To cnt)
60      ReDim disp(1 To cnt)
70      Dim i As Long
80      Dim k As Variant
90      i = 0
100     For Each k In dictCol.keys
110         i = i + 1
120         keys(i) = CStr(k)
130         cols(i) = CLng(dictCol(k))
140         disp(i) = CStr(dictDisp(k))
150     Next k
160     Exit Sub
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
170     Err.Raise Err.Number, "BuildArrays:" & Erl, Err.Description
End Sub

' ============================================================================
' HELPER: Get string parameter from NIHUL
' ============================================================================
Private Function GetStringParameter(ByVal ws As Worksheet, ByVal paramName As String) As String
10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False
20      Dim r As Long
30      Dim startRow As Long
40      Dim nm As String
        ' Start from row after section header (rngSection_Params)
50      startRow = ThisWorkbook.Names("rngSection_Params").RefersToRange.Row + 1
60      r = startRow
70      Do While True
80          nm = UCase$(Trim$(CStr(ws.Cells(r, COL_PARAM_NAME).Value2)))
90          If nm = EOD_MARKER Or nm = "" Then Exit Do
100         If nm = UCase$(paramName) Then
110             GetStringParameter = Trim$(CStr(ws.Cells(r, COL_PARAM_VALUE).Value2))
120             Exit Function
130         End If
140         r = r + 1
150     Loop
160     GetStringParameter = ""
170     Exit Function
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
180     GetStringParameter = ""
End Function

' ============================================================================
' HELPER: Get numeric parameter from NIHUL
' ============================================================================
Private Function GetNumericParameter(ByVal ws As Worksheet, ByVal paramName As String) As Double
10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False
20      Dim r As Long
30      Dim startRow As Long
40      Dim nm As String
50      Dim v As Variant
60      Dim n As Double
        ' Start from row after section header (rngSection_Params)
70      startRow = ThisWorkbook.Names("rngSection_Params").RefersToRange.Row + 1
80      r = startRow
90      Do While True
100         nm = UCase$(Trim$(CStr(ws.Cells(r, COL_PARAM_NAME).Value2)))
110         If nm = EOD_MARKER Or nm = "" Then Exit Do
120         If nm = UCase$(paramName) Then
130             v = ws.Cells(r, COL_PARAM_VALUE).Value2
140             If TryParseVariantNumber(v, n) Then
150                 GetNumericParameter = n
160             Else
170                 Err.Raise vbObjectError + 3000, "GetNumericParameter", "PARAMETER NOT NUMERIC: " & paramName
180             End If
190             Exit Function
200         End If
210         r = r + 1
220     Loop
230     Err.Raise vbObjectError + 3001, "GetNumericParameter", "PARAMETER NOT FOUND: " & paramName
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
240     Err.Raise Err.Number, "GetNumericParameter:" & Erl, Err.Description
End Function

' ============================================================================
' HELPER: Column letter to number
' ============================================================================
Private Function ColumnLetterToNumber(ByVal col As String) As Long
10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False
20      Dim i As Long
30      Dim ch As String
40      For i = 1 To Len(col)
50          ch = Mid$(col, i, 1)
60          If ch < "A" Or ch > "Z" Then
70              Err.Raise vbObjectError + 4000, "ColumnLetterToNumber", "INVALID COLUMN LETTER: " & col
80          End If
90          ColumnLetterToNumber = ColumnLetterToNumber * 26 + (Asc(ch) - 64)
100     Next i
110     Exit Function
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
120     Err.Raise Err.Number, "ColumnLetterToNumber:" & Erl, Err.Description
End Function

' ============================================================================
' HELPER: Delete sheets
' ============================================================================
Private Sub DeleteReviewSheetIfExists()
10      On Error Resume Next
20      Application.DisplayAlerts = False
30      If SheetExists(REVIEW_SHEET_NAME()) Then
40          ThisWorkbook.Worksheets(REVIEW_SHEET_NAME()).Delete
50      End If
60      Application.DisplayAlerts = True
End Sub

Private Sub DeleteSheetIfExists(ByVal sName As String)
        Dim wsTarget As Worksheet
        Dim prevAlerts As Boolean
        Dim wsCtrl As Worksheet
10      If Not SheetExists(sName) Then Exit Sub

20      prevAlerts = Application.DisplayAlerts
30      Application.DisplayAlerts = False

        ' Ensure control sheet (daf habait) is visible so we always have at least one visible sheet
35      On Error Resume Next
36      Set wsCtrl = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
37      If Not wsCtrl Is Nothing Then wsCtrl.Visible = xlSheetVisible
38      On Error GoTo 0

        ' Unhide sheet if hidden (cannot delete very hidden sheets without unhiding first)
40      Set wsTarget = ThisWorkbook.Worksheets(sName)
50      On Error Resume Next
60      If wsTarget.Visible <> xlSheetVisible Then wsTarget.Visible = xlSheetVisible
70      On Error GoTo 0

        ' Delete the sheet
80      On Error Resume Next
90      wsTarget.Delete
100     On Error GoTo 0

        ' Verify deletion succeeded
110     If SheetExists(sName) Then
            ' Second attempt - force
120         On Error Resume Next
130         ThisWorkbook.Worksheets(sName).Visible = xlSheetVisible
140         ThisWorkbook.Worksheets(sName).Delete
150         On Error GoTo 0
160     End If

170     Application.DisplayAlerts = prevAlerts
End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
10      On Error GoTo NOT_FOUND
20      Dim ws As Worksheet
30      Set ws = ThisWorkbook.Worksheets(sheetName)
40      SheetExists = True
50      Exit Function
NOT_FOUND:
60      SheetExists = False
End Function

' ============================================================================
' HELPER: Strip common insurance company suffixes from name
' Removes: ???? ?????? ??"? / ????? ??"? / ??"?
' ============================================================================
Private Function ShortenCompanyName(ByVal sName As String) As String
    Dim sSuffix As String
    Dim pos As Long
    ' ???? ?????? ??"?
    sSuffix = H_COMPANY() & " " & ChrW(1500) & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495) & " " & ChrW(1489) & ChrW(1506) & ChrW(34) & ChrW(1502)
    pos = InStr(1, sName, sSuffix, vbTextCompare)
    If pos > 1 Then
        ShortenCompanyName = Trim$(Left$(sName, pos - 1))
        Exit Function
    End If
    ' ?????? ??"?
    sSuffix = ChrW(1500) & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495) & " " & ChrW(1489) & ChrW(1506) & ChrW(34) & ChrW(1502)
    pos = InStr(1, sName, sSuffix, vbTextCompare)
    If pos > 1 Then
        ShortenCompanyName = Trim$(Left$(sName, pos - 1))
        Exit Function
    End If
    ' ??"? alone at end
    sSuffix = " " & ChrW(1489) & ChrW(1506) & ChrW(34) & ChrW(1502)
    If Right$(sName, Len(sSuffix)) = sSuffix Then
        ShortenCompanyName = Trim$(Left$(sName, Len(sName) - Len(sSuffix)))
        Exit Function
    End If
    ShortenCompanyName = sName
End Function

' ============================================================================
' HELPER: Fetch monthly average exchange rate from Bank of Israel SDMX API
' currencyKey: "USD" or "EUR"
' monthStr: month in format "YYYY-MM" (e.g. "2025-01")
'           If empty, fetches most recent rate (last 7 days, returns last available)
' Uses BOI SDMX API: fetches all daily rates for the month and calculates average
' Returns the average rate for 1 unit in ILS. Returns 0 on failure.
' ============================================================================
Private Function FetchBOIMonthlyAvg(ByVal currencyKey As String, Optional ByVal monthStr As String = "") As Double
    On Error GoTo FAIL
    Dim xmlHttp As Object
    Dim url As String
    Dim responseText As String
    Dim lines() As String
    Dim fields() As String
    Dim i As Long
    Dim lineRate As Double
    Dim sumRate As Double
    Dim countRate As Long
    Dim startDate As String
    Dim endDate As String
    Dim lastValidRate As Double
    
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    xmlHttp.setTimeouts 5000, 5000, 5000, 10000
    
    If monthStr = "" Then
        ' Fetch most recent rate (last 7 days)
        startDate = Format$(Date - 7, "yyyy-mm-dd")
        endDate = Format$(Date, "yyyy-mm-dd")
    Else
        ' Fetch all daily rates for the entire month
        startDate = monthStr & "-01"
        endDate = monthStr & "-31"  ' API handles months with fewer days gracefully
    End If
    
    url = "https://edge.boi.gov.il/FusionEdgeServer/sdmx/v2/data/dataflow/BOI.STATISTICS/EXR/1.0/" & _
          "?c%5BBASE_CURRENCY%5D=" & UCase$(currencyKey) & _
          "&c%5BCOUNTER_CURRENCY%5D=ILS" & _
          "&c%5BDATA_TYPE%5D=OF00" & _
          "&startPeriod=" & startDate & _
          "&endPeriod=" & endDate & _
          "&format=csv"
    
    xmlHttp.Open "GET", url, False
    xmlHttp.send
    
    If xmlHttp.Status <> 200 Then GoTo FAIL
    
    responseText = xmlHttp.responseText
    If Len(responseText) < 20 Then GoTo FAIL
    
    ' Parse CSV response - calculate average of all daily rates
    lastValidRate = 0
    lines = Split(responseText, vbLf)
    sumRate = 0
    countRate = 0
    
    For i = 1 To UBound(lines)  ' Skip header row
        If Len(Trim$(lines(i))) > 10 Then
            fields = Split(lines(i), ",")
            If UBound(fields) >= 13 Then
                If IsNumeric(Trim$(fields(13))) Then
                    lineRate = CDbl(Trim$(fields(13)))
                    If lineRate > 0 Then
                        sumRate = sumRate + lineRate
                        countRate = countRate + 1
                        lastValidRate = lineRate
                    End If
                End If
            End If
        End If
    Next i
    
    If countRate > 0 Then
        If monthStr = "" Then
            FetchBOIMonthlyAvg = lastValidRate
        Else
            FetchBOIMonthlyAvg = Round(sumRate / countRate, 4)
        End If
    End If
    Exit Function
    
FAIL:
    FetchBOIMonthlyAvg = 0
End Function

' ============================================================================
' HELPER: Backward-compatible wrapper - fetch current USD rate (most recent)
' ============================================================================
Private Function FetchBOIRate() As Double
    FetchBOIRate = FetchBOIMonthlyAvg("USD", "")
End Function

' ============================================================================
' HELPER: Get currency exchange rate by bordero month from matach sheet
' currencyKey: "USD" (col B) or "EUR" (col C)
' Looks up YYYY-MM in column A of matach sheet, returns rate from col B (USD) or C (EUR)
' If month not found and skipBOI=False, fetches from BOI API for the 15th of that month
' Final fallback: rngDOLAR cell value (for USD) or 1 (for EUR if not found)
' ============================================================================
Private Function GetCurrencyRate(ByVal wsMain As Worksheet, Optional ByVal bordMonth As String = "", Optional ByVal skipBOI As Boolean = False, Optional ByVal currencyKey As String = "USD") As Double
    Dim wsMatach As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim cellVal As String
    Dim rate As Double
    Dim rateCol As Long
    Dim rateDate As String
    
    ' Determine column: B=USD, C=EUR
    If UCase$(currencyKey) = "EUR" Then
        rateCol = 3
    Else
        rateCol = 2
    End If
    
    ' Try to find matach sheet
    On Error Resume Next
    Set wsMatach = ThisWorkbook.Worksheets(MATACH_SHEET_NAME())
    On Error GoTo 0
    
    ' If matach sheet exists and bordMonth provided, look up rate
    If Not wsMatach Is Nothing And bordMonth <> "" Then
        lastRow = wsMatach.Cells(wsMatach.Rows.Count, 1).End(xlUp).Row
        For r = 2 To lastRow
            cellVal = Trim$(CStr(wsMatach.Cells(r, 1).Value2))
            If cellVal = bordMonth Then
                If IsNumeric(wsMatach.Cells(r, rateCol).Value2) Then
                    rate = CDbl(wsMatach.Cells(r, rateCol).Value2)
                    If rate > 0 Then
                        GetCurrencyRate = rate
                        Exit Function
                    End If
                End If
            End If
        Next r
    End If
    
    ' Fallback: fetch rate from BOI API for the 15th of the bordero month
    If Not skipBOI Then
        If bordMonth <> "" And Len(bordMonth) = 7 Then
            rateDate = bordMonth  ' YYYY-MM format
        Else
            rateDate = ""  ' current rate
        End If
        rate = FetchBOIMonthlyAvg(currencyKey, rateDate)
        If rate > 0 Then
            ' Store in rngDOLAR if USD and no specific month
            If UCase$(currencyKey) = "USD" And bordMonth = "" Then
                On Error Resume Next
                wsMain.Range("rngDOLAR").Value = rate
                On Error GoTo 0
            End If
            GetCurrencyRate = rate
            Exit Function
        End If
    End If
    
    ' Final fallback: read from rngDOLAR cell (USD only)
    If UCase$(currencyKey) = "USD" Then
        GetCurrencyRate = 1
        On Error Resume Next
        If Not IsBlankValue(wsMain.Range("rngDOLAR").Value2) Then
            If IsNumeric(wsMain.Range("rngDOLAR").Value2) Then GetCurrencyRate = CDbl(wsMain.Range("rngDOLAR").Value2)
        End If
        On Error GoTo 0
    Else
        GetCurrencyRate = 1
    End If
End Function

' ============================================================================
' HELPER: Backward-compatible wrapper for GetCurrencyRate (USD)
' ============================================================================
Private Function GetDollarRate(ByVal wsMain As Worksheet, Optional ByVal bordMonth As String = "", Optional ByVal skipBOI As Boolean = False) As Double
    GetDollarRate = GetCurrencyRate(wsMain, bordMonth, skipBOI, "USD")
End Function

' ============================================================================
' HELPER: Build/update matach cache sheet with USD+EUR rates for all needed months
' Scans source data for unique bordero months, fetches missing rates from BOI API
' Creates matach sheet if it doesn't exist
' Structure: Col A = YYYY-MM, Col B = USD rate, Col C = EUR rate
' ============================================================================
Private Sub BuildMatachCache(ByVal wsSource As Worksheet, ByVal dateCol As Long, ByVal lastDataRow As Long)
    Dim wsMatach As Worksheet
    Dim dictMonths As Object
    Dim r As Long
    Dim bordDate As Variant
    Dim monthKey As String
    Dim lastRow As Long
    Dim cellVal As String
    Dim rateDate As String
    Dim usdRate As Double
    Dim eurRate As Double
    Dim newRow As Long
    Dim vKey As Variant
    Dim found As Boolean
    
    ' Create dictionary of unique bordero months from source data
    Set dictMonths = CreateObject("Scripting.Dictionary")
    For r = 2 To lastDataRow
        bordDate = wsSource.Cells(r, dateCol).Value2
        monthKey = GetBordMonth(bordDate)
        If monthKey <> "" Then
            If Not dictMonths.Exists(monthKey) Then dictMonths.Add monthKey, True
        End If
    Next r
    
    If dictMonths.Count = 0 Then Exit Sub
    
    ' Get or create matach sheet
    On Error Resume Next
    Set wsMatach = ThisWorkbook.Worksheets(MATACH_SHEET_NAME())
    On Error GoTo 0
    
    If wsMatach Is Nothing Then
        ' Create matach sheet
        Set wsMatach = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsMatach.Name = MATACH_SHEET_NAME()
        wsMatach.Cells(1, 1).Value = "Month"
        wsMatach.Cells(1, 2).Value = "USD"
        wsMatach.Cells(1, 3).Value = "EUR"
        wsMatach.Visible = xlSheetVeryHidden
    End If
    
    ' --- For each unique month, check if already in matach; if not, fetch monthly avg from API ---
    For Each vKey In dictMonths.keys
        monthKey = CStr(vKey)
        ' Check if month already exists in matach with valid rates
        found = False
        lastRow = wsMatach.Cells(wsMatach.Rows.Count, 1).End(xlUp).Row
        For r = 2 To lastRow
            cellVal = Trim$(CStr(wsMatach.Cells(r, 1).Value2))
            If cellVal = monthKey Then
                ' Check if both USD and EUR have values > 0
                If IsNumeric(wsMatach.Cells(r, 2).Value2) And IsNumeric(wsMatach.Cells(r, 3).Value2) Then
                    If CDbl(wsMatach.Cells(r, 2).Value2) > 0 And CDbl(wsMatach.Cells(r, 3).Value2) > 0 Then
                        found = True
                        Exit For
                    End If
                End If
            End If
        Next r
        
        If Not found Then
            ' Month not in matach (future month or missing) - fetch monthly avg from BOI API
            usdRate = FetchBOIMonthlyAvg("USD", monthKey)
            eurRate = FetchBOIMonthlyAvg("EUR", monthKey)
            
            ' If API failed (future month with no data), try current rate as fallback
            If usdRate = 0 Then usdRate = FetchBOIMonthlyAvg("USD", "")
            If eurRate = 0 Then eurRate = FetchBOIMonthlyAvg("EUR", "")
            
            ' Write to matach sheet (only if we got a rate)
            If usdRate > 0 Or eurRate > 0 Then
                newRow = wsMatach.Cells(wsMatach.Rows.Count, 1).End(xlUp).Row + 1
                wsMatach.Cells(newRow, 1).Value = monthKey
                wsMatach.Cells(newRow, 2).Value = usdRate
                wsMatach.Cells(newRow, 3).Value = eurRate
            End If
        End If
    Next vKey
End Sub

' ============================================================================
' HELPER: Legacy wrapper - Update matach sheet with current month rate
' Adds a new row with YYYY-MM format and the BOI rate
' ============================================================================
Private Sub UpdateMatachCurrentMonth(ByVal wsMatach As Worksheet, ByVal boiRate As Double)
    Dim curMonth As String
    Dim lastRow As Long
    Dim r As Long
    Dim cellVal As String
    
    curMonth = Format$(Date, "yyyy-mm")
    lastRow = wsMatach.Cells(wsMatach.Rows.Count, 1).End(xlUp).Row
    
    ' Check if current month already exists
    For r = 2 To lastRow
        cellVal = Trim$(CStr(wsMatach.Cells(r, 1).Value2))
        If cellVal = curMonth Then Exit Sub  ' Already exists
    Next r
    
    ' Add new row
    wsMatach.Cells(lastRow + 1, 1).Value = curMonth
    wsMatach.Cells(lastRow + 1, 2).Value = boiRate
End Sub

' ============================================================================
' HELPER: Get bordero month as YYYY-MM string from a date value
' ============================================================================
Private Function GetBordMonth(ByVal bordereu As Variant) As String
    Dim dt As Date
    On Error GoTo FAIL
    If IsDate(bordereu) Then
        dt = CDate(bordereu)
    ElseIf IsNumeric(bordereu) Then
        If CDbl(bordereu) > 1 Then dt = CDate(CDbl(bordereu))
    Else
        GoTo FAIL
    End If
    GetBordMonth = Format$(dt, "yyyy-mm")
    Exit Function
FAIL:
    GetBordMonth = ""
End Function

' ============================================================================
' HELPER: Check if row should be ignored
' ============================================================================
Private Function IsIgnorableRow(ByVal ws As Worksheet, ByVal r As Long, ByRef keys() As String, ByRef cols() As Long, ByVal cnt As Long, ByVal dictFieldCol As Object) As Boolean
10      Dim allBlank As Boolean
20      Dim i As Long
30      allBlank = True
40      For i = 1 To cnt
50          If Not IsBlankValue(ws.Cells(r, cols(i)).Value2) Then
60              allBlank = False
70              Exit For
80          End If
90      Next i
100     IsIgnorableRow = allBlank
End Function

' ============================================================================
' HELPER: Check blank value
' ============================================================================
Private Function IsBlankValue(ByVal v As Variant) As Boolean
10      If IsEmpty(v) Then
20          IsBlankValue = True
30      ElseIf IsNull(v) Then
40          IsBlankValue = True
50      ElseIf VarType(v) = vbString Then
60          IsBlankValue = (Trim$(CStr(v)) = "")
70      Else
80          IsBlankValue = False
90      End If
End Function

' ============================================================================
' HELPER: Try parse variant to number
' ============================================================================
Private Function TryParseVariantNumber(ByVal v As Variant, ByRef result As Double) As Boolean
10      On Error GoTo FAIL
20      If IsNumeric(v) Then
30          result = CDbl(v)
40          TryParseVariantNumber = True
50      Else
60          TryParseVariantNumber = False
70      End If
80      Exit Function
FAIL:
90      TryParseVariantNumber = False
End Function

' ============================================================================
' HELPER: Add reason code
' ============================================================================
Private Function AddReason(ByVal existing As String, ByVal newReason As String) As String
10      If existing = "" Then
20          AddReason = newReason
30      Else
40          AddReason = existing & ", " & newReason
50      End If
End Function

' ============================================================================
' HELPER: Translate reason codes to Hebrew
' ============================================================================
Private Function TranslateReason(ByVal reasonCode As String, ByVal dictHelper As Object) As String
10      On Error Resume Next
20      Dim parts() As String
30      Dim i As Long
40      Dim translated As String
50      Dim part As String
60      parts = Split(reasonCode, ", ")
70      For i = 0 To UBound(parts)
80          part = Trim$(parts(i))
90          If dictHelper.Exists(part) Then
100             part = dictHelper(part)
110         End If
120         If translated = "" Then
130             translated = part
140         Else
150             translated = translated & ", " & part
160         End If
170     Next i
180     TranslateReason = translated
End Function

' ============================================================================
' HELPER: Return Hebrew month name for month number 1-12
' ============================================================================
Private Function HebrewMonthName(ByVal m As Long) As String
10      If m = 1 Then
20          HebrewMonthName = ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1488) & ChrW(1512)
30      ElseIf m = 2 Then
40          HebrewMonthName = ChrW(1508) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1488) & ChrW(1512)
50      ElseIf m = 3 Then
60          HebrewMonthName = ChrW(1502) & ChrW(1512) & ChrW(1509)
70      ElseIf m = 4 Then
80          HebrewMonthName = ChrW(1488) & ChrW(1508) & ChrW(1512) & ChrW(1497) & ChrW(1500)
90      ElseIf m = 5 Then
100         HebrewMonthName = ChrW(1502) & ChrW(1488) & ChrW(1497)
110     ElseIf m = 6 Then
120         HebrewMonthName = ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1497)
130     ElseIf m = 7 Then
140         HebrewMonthName = ChrW(1497) & ChrW(1493) & ChrW(1500) & ChrW(1497)
150     ElseIf m = 8 Then
160         HebrewMonthName = ChrW(1488) & ChrW(1493) & ChrW(1490) & ChrW(1493) & ChrW(1505) & ChrW(1496)
170     ElseIf m = 9 Then
180         HebrewMonthName = ChrW(1505) & ChrW(1508) & ChrW(1496) & ChrW(1502) & ChrW(1489) & ChrW(1512)
190     ElseIf m = 10 Then
200         HebrewMonthName = ChrW(1488) & ChrW(1493) & ChrW(1511) & ChrW(1496) & ChrW(1493) & ChrW(1489) & ChrW(1512)
210     ElseIf m = 11 Then
220         HebrewMonthName = ChrW(1504) & ChrW(1493) & ChrW(1489) & ChrW(1502) & ChrW(1489) & ChrW(1512)
230     ElseIf m = 12 Then
240         HebrewMonthName = ChrW(1491) & ChrW(1510) & ChrW(1502) & ChrW(1489) & ChrW(1512)
250     Else
260         HebrewMonthName = CStr(m)
270     End If
End Function

' ============================================================================
' SETUP: Create labels, dropdowns, period lists, and buttons on Main sheet
' Run this once to set up the control panel
' ============================================================================
Public Sub A00_SetupMainSheet()

    Dim wsMain As Worksheet
    Dim wsMgmt As Worksheet
    Dim shp As Shape
    Dim s As Shape
    Dim blueClr As Long
    Dim lblShnBasis As String
    Dim lblShnShotef As String
    Dim lblTkufa As String
    Dim lblPirutTkufa As String
    Dim lblSugTaarih As String
    Dim periodTypeList As String
    Dim halfList As String
    Dim quarterList As String
    Dim monthList As String
    Dim periodTypeFormula As String
    Dim dvC4 As String
    Dim btnSearch As Shape
    Dim btnAll As Shape
    Dim btnReset As Shape
    Dim hdrRng As Range
    Dim foundBackup As Boolean
    Dim filterTypeList As String
    Dim wsLoop As Worksheet
    Dim rngMsg As Range
    Dim msgTxt As String
    Dim borderShp As Shape
    Dim rEdge As Variant
    Dim cellRate As Range
    Dim cellBorder As Range
    Dim bEdge As Variant
    Dim wsMatachSetup As Worksheet
    Dim curRate As Double
    Dim eurRate As Double
    Dim matachMsg As String
    Dim existingPath As String
    Dim paramStartRow As Long
    Dim paramLastRow As Long
    Dim foundEmail As Boolean
    Dim pr As Long
    Dim msgBaseRow As Long
    Dim periodBaseRow As Long
    ' Disable events to prevent Worksheet_Change from firing during setupp
10  Application.EnableEvents = False
20  On Error GoTo ERR_HANDLER

    ' Try to find Main sheet by current Hebrew name or old English name
30  On Error Resume Next
40  Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
50  If wsMain Is Nothing Then Set wsMain = ThisWorkbook.Worksheets("Main")
60  On Error GoTo ERR_HANDLER
    ' If not found, create it as the first sheet
70  If wsMain Is Nothing Then
80  Set wsMain = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
90  wsMain.Name = CONTROL_SHEET_NAME()
100  End If
    ' Rename to Hebrew if still English
110  If wsMain.Name <> CONTROL_SHEET_NAME() Then wsMain.Name = CONTROL_SHEET_NAME()
    ' Unprotect home sheet in case it's protected
120  On Error Resume Next
130  wsMain.Unprotect "Z961814r"
140  Err.Clear
150  On Error GoTo ERR_HANDLER
    ' Try to find NIHUL/hagdarot sheet by current Hebrew name or old English name
160  On Error Resume Next
170  Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
180  If wsMgmt Is Nothing Then Set wsMgmt = ThisWorkbook.Worksheets("NIHUL")
190  Err.Clear
200  On Error GoTo ERR_HANDLER
210  If wsMgmt Is Nothing Then Err.Raise vbObjectError + 9002, "SetupMainSheet", "Cannot find NIHUL/hagdarot sheet"
    ' Rename to Hebrew if still English
220  If wsMgmt.Name <> MANAGEMENT_SHEET_NAME() Then wsMgmt.Name = MANAGEMENT_SHEET_NAME()
    ' Unprotect management sheet in case it's protected
230  On Error Resume Next
240  wsMgmt.Unprotect "Z961814r"
250  Err.Clear
260  On Error GoTo ERR_HANDLER
270  blueClr = RGB(0, 70, 140)

    ' ---- Hebrew label strings ----
    ' "shnat basis" = year of reference
280  lblShnBasis = ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & H_BASE()
    ' "shna shoteft" = current year
290  lblShnShotef = ChrW(1513) & ChrW(1504) & ChrW(1492) & " " & ChrW(1513) & ChrW(1493) & ChrW(1496) & ChrW(1508) & ChrW(1514)
    ' "tkufa" = period
300  lblTkufa = ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
    ' "pirot tkufa" = period detail
310  lblPirutTkufa = ChrW(1508) & ChrW(1497) & ChrW(1512) & ChrW(1493) & ChrW(1496) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
    ' "sug taarih" = date type
320  lblSugTaarih = ChrW(1505) & ChrW(1493) & ChrW(1490) & " " & ChrW(1514) & ChrW(1488) & ChrW(1512) & ChrW(1497) & ChrW(1498)

    ' ---- Clear old labels from A2:D5 (no longer used) ----
    wsMain.Range("A2:K20").ClearContents
    wsMain.Cells.Font.Size = 14
    wsMain.Cells.Interior.ColorIndex = xlNone
    wsMain.Range("A1:AZ40").Interior.Color = RGB(220, 240, 220)
    wsMain.Range("A2:K20").Borders.LineStyle = xlNone

    ' ---- ONE-TIME HARD RESET FOR LAYOUT (V1.64) ----
    Dim permRowCheck As Long
    permRowCheck = 300 + offsetRow
    
    If wsMgmt.Range("A" & permRowCheck).Value <> ChrW(1492) & ChrW(1512) & ChrW(1513) & ChrW(1488) & ChrW(1493) & ChrW(1514) Then
        ' 1. Wipe everything from 210+offsetRow to 1000 (preserving reason codes up to 209)
        wsMgmt.Range("A" & (210 + offsetRow) & ":F" & (1000 + offsetRow)).ClearContents
        wsMgmt.Range("A" & (210 + offsetRow) & ":F" & (1000 + offsetRow)).Font.Bold = False
        wsMgmt.Range("A" & (210 + offsetRow) & ":F" & (1000 + offsetRow)).Font.Size = 11
        
        ' 2. Recreate Permissions Header
        wsMgmt.Range("A" & permRowCheck).Value = ChrW(1492) & ChrW(1512) & ChrW(1513) & ChrW(1488) & ChrW(1493) & ChrW(1514)
        wsMgmt.Range("A" & permRowCheck).Font.Bold = True
        wsMgmt.Range("A" & permRowCheck).Font.Size = 12
        wsMgmt.Range("A" & permRowCheck + 1).Value = "EOD"
        
        ' 3. Recreate Clients Header
        Dim clientRowCheck As Long
        clientRowCheck = 350 + offsetRow
        wsMgmt.Range("A" & clientRowCheck).Value = ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1514) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
        wsMgmt.Range("A" & clientRowCheck).Font.Bold = True
        wsMgmt.Range("A" & clientRowCheck).Font.Size = 12
        wsMgmt.Range("A" & clientRowCheck + 1).Value = "EOD"
    End If

    ' ---- Write period lookup lists in PeriodLists section ----
    SetupSectionNamedRanges
    periodBaseRow = ThisWorkbook.Names("rngSection_PeriodLists").RefersToRange.Row
    wsMgmt.Range("A" & (periodBaseRow + 1) & ":C" & (periodBaseRow + 40)).ClearContents
    wsMgmt.Cells(periodBaseRow + 1, 1).Value = "PERIOD_TYPE"
    wsMgmt.Cells(periodBaseRow + 2, 1).Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 3, 1).Value = ChrW(1495) & ChrW(1510) & ChrW(1497) & " " & ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 4, 1).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 5, 1).Value = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 6, 1).Value = "HALF_YEAR"
    wsMgmt.Cells(periodBaseRow + 7, 1).Value = ChrW(1502) & ChrW(1495) & ChrW(1510) & ChrW(1497) & ChrW(1514) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1504) & ChrW(1492)
    wsMgmt.Cells(periodBaseRow + 8, 1).Value = ChrW(1502) & ChrW(1495) & ChrW(1510) & ChrW(1497) & ChrW(1514) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497) & ChrW(1492)
    
    On Error Resume Next
    ThisWorkbook.Names("lst_period_type").Delete
    ThisWorkbook.Names("lst_half_year").Delete
    ThisWorkbook.Names("lst_quarter").Delete
    ThisWorkbook.Names("lst_month").Delete
    Err.Clear
    On Error GoTo ERR_HANDLER
    ThisWorkbook.Names.Add "lst_period_type", wsMgmt.Range(wsMgmt.Cells(periodBaseRow + 2, 1), wsMgmt.Cells(periodBaseRow + 5, 1))
    ThisWorkbook.Names.Add "lst_half_year", wsMgmt.Range(wsMgmt.Cells(periodBaseRow + 7, 1), wsMgmt.Cells(periodBaseRow + 8, 1))
    
    wsMgmt.Cells(periodBaseRow + 10, 1).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1503)
    wsMgmt.Cells(periodBaseRow + 11, 1).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 12, 1).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1513) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 13, 1).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1489) & ChrW(1497) & ChrW(1506) & ChrW(1497)
    ThisWorkbook.Names.Add "lst_quarter", wsMgmt.Range(wsMgmt.Cells(periodBaseRow + 10, 1), wsMgmt.Cells(periodBaseRow + 13, 1))
    
    wsMgmt.Cells(periodBaseRow + 15, 1).Value = ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1488) & ChrW(1512)
    wsMgmt.Cells(periodBaseRow + 16, 1).Value = ChrW(1508) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1488) & ChrW(1512)
    wsMgmt.Cells(periodBaseRow + 17, 1).Value = ChrW(1502) & ChrW(1512) & ChrW(1509)
    wsMgmt.Cells(periodBaseRow + 18, 1).Value = ChrW(1488) & ChrW(1508) & ChrW(1512) & ChrW(1497) & ChrW(1500)
    wsMgmt.Cells(periodBaseRow + 19, 1).Value = ChrW(1502) & ChrW(1488) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 20, 1).Value = ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 21, 1).Value = ChrW(1497) & ChrW(1493) & ChrW(1500) & ChrW(1497)
    wsMgmt.Cells(periodBaseRow + 22, 1).Value = ChrW(1488) & ChrW(1493) & ChrW(1490) & ChrW(1493) & ChrW(1505) & ChrW(1496)
    wsMgmt.Cells(periodBaseRow + 23, 1).Value = ChrW(1505) & ChrW(1508) & ChrW(1496) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    wsMgmt.Cells(periodBaseRow + 24, 1).Value = ChrW(1488) & ChrW(1493) & ChrW(1511) & ChrW(1496) & ChrW(1493) & ChrW(1489) & ChrW(1512)
    wsMgmt.Cells(periodBaseRow + 25, 1).Value = ChrW(1504) & ChrW(1493) & ChrW(1489) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    wsMgmt.Cells(periodBaseRow + 26, 1).Value = ChrW(1491) & ChrW(1510) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    ThisWorkbook.Names.Add "lst_month", wsMgmt.Range(wsMgmt.Cells(periodBaseRow + 15, 1), wsMgmt.Cells(periodBaseRow + 26, 1))
    
    ' ---- Layout setup for Home Sheet ----
    ' Parameters Table (F2:G10)
    wsMain.Range("F2").Value = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512)
    wsMain.Range("G2").Value = ChrW(1506) & ChrW(1512) & ChrW(1498)
    Set hdrRng = wsMain.Range("F2:G2")
    hdrRng.Interior.Color = RGB(0, 100, 0)
    hdrRng.Font.Color = RGB(255, 255, 255)
    hdrRng.Font.Bold = True
    hdrRng.Font.Size = 12
    hdrRng.HorizontalAlignment = xlCenter
    
    wsMain.Range("F3").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505)
    wsMain.Range("F4").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1499) & ChrW(1495) & ChrW(1497) & ChrW(1514)
    wsMain.Range("F5").Value = ChrW(1505) & ChrW(1493) & ChrW(1490) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
    wsMain.Range("F6").Value = ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
    wsMain.Range("F7").Value = ChrW(1505) & ChrW(1493) & ChrW(1490) & " " & ChrW(1514) & ChrW(1488) & ChrW(1512) & ChrW(1497) & ChrW(1498)
    wsMain.Range("F8").Value = ChrW(1495) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1498) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497)
    wsMain.Range("F9").Value = ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1495) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1498)
    wsMain.Range("F10").Value = ChrW(1513) & ChrW(1501) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)
    
    wsMain.Range("F3:F10").Font.Color = RGB(0, 70, 140)
    wsMain.Range("F3:F10").Font.Bold = True
    wsMain.Range("F3:F10").Interior.Color = RGB(200, 230, 255)
    wsMain.Range("G3:G10").Interior.Color = RGB(255, 245, 230)
    wsMain.Range("F3:G10").HorizontalAlignment = xlCenter
    wsMain.Range("F3:F10").HorizontalAlignment = xlRight
    wsMain.Range("G3:G12").Locked = False
    
    On Error Resume Next
    ThisWorkbook.Names("rngCurrentYear").Delete
    ThisWorkbook.Names("rngBaseYear").Delete
    ThisWorkbook.Names("rngPeriodType").Delete
    ThisWorkbook.Names("rngPeriodValue").Delete
    ThisWorkbook.Names("rngDateType").Delete
    ThisWorkbook.Names("rngFilterType").Delete
    ThisWorkbook.Names("rngFilterValue").Delete
    ThisWorkbook.Names("rngClientName").Delete
    Err.Clear
    On Error GoTo ERR_HANDLER
    
    ThisWorkbook.Names.Add "rngBaseYear", wsMain.Range("G3")
    ThisWorkbook.Names.Add "rngCurrentYear", wsMain.Range("G4")
    ThisWorkbook.Names.Add "rngPeriodType", wsMain.Range("G5")
    ThisWorkbook.Names.Add "rngPeriodValue", wsMain.Range("G6")
    ThisWorkbook.Names.Add "rngDateType", wsMain.Range("G7")
    ThisWorkbook.Names.Add "rngFilterType", wsMain.Range("G8")
    ThisWorkbook.Names.Add "rngFilterValue", wsMain.Range("G9")
    ThisWorkbook.Names.Add "rngClientName", wsMain.Range("G10")
    
    ' Set defaults
    If IsEmpty(wsMain.Range("G3").Value) Then wsMain.Range("G3").Value = 2025
    If IsEmpty(wsMain.Range("G4").Value) Then wsMain.Range("G4").Value = 2026
    If IsEmpty(wsMain.Range("G5").Value) Then wsMain.Range("G5").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
    If IsEmpty(wsMain.Range("G8").Value) Then wsMain.Range("G8").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)
    If IsEmpty(wsMain.Range("G10").Value) Then wsMain.Range("G10").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)
    
    On Error Resume Next
    wsMain.Range("G5").Validation.Delete
    wsMain.Range("G5").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lst_period_type"
    wsMain.Range("G6").Validation.Delete
    wsMain.Range("G7").Validation.Delete
    dateTypeList = ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493) & "," & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & ChrW(1514) & " " & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495)
    wsMain.Range("G7").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=dateTypeList
    wsMain.Range("G7").Value = ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493)
    wsMain.Range("G8").Validation.Delete
    filterTypeList = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497) & "," & H_COMPANY() & "," & H_TELLER() & "," & H_AGENT() & "," & H_BRANCH() & "," & H_BRANCH() & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)
    wsMain.Range("G8").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=filterTypeList
    wsMain.Range("G9").Validation.Delete
    Err.Clear
    On Error GoTo ERR_HANDLER
    
    ' ---- Currency Table J2:K4 ----
    wsMain.Range("J2").Value = ChrW(1502) & ChrW(1496) & ChrW(1489) & ChrW(1506)
    wsMain.Range("K2").Value = ChrW(1513) & ChrW(1506) & ChrW(1512)
    wsMain.Range("J2:K2").Interior.Color = RGB(0, 100, 0)
    wsMain.Range("J2:K2").Font.Color = RGB(255, 255, 255)
    wsMain.Range("J2:K2").Font.Bold = True
    wsMain.Range("J2:K2").Font.Size = 12
    wsMain.Range("J2:K2").HorizontalAlignment = xlCenter
    wsMain.Range("J3").Value = "$" & " " & ChrW(1491) & ChrW(1493) & ChrW(1500) & ChrW(1512)
    wsMain.Range("J4").Value = ChrW(8364) & " " & ChrW(1488) & ChrW(1497) & ChrW(1512) & ChrW(1493)
    wsMain.Range("J3:J4").Font.Bold = True
    wsMain.Range("K3:K4").HorizontalAlignment = xlCenter
    wsMain.Range("J3:J4").Interior.Color = RGB(200, 230, 255)
    wsMain.Range("K3:K4").Interior.Color = RGB(255, 245, 230)
    
    On Error Resume Next
    ThisWorkbook.Names("rngDOLAR").Delete
    Err.Clear
    On Error GoTo ERR_HANDLER
    ThisWorkbook.Names.Add "rngDOLAR", wsMain.Range("K3")
    
    curRate = FetchBOIRate()
    If curRate > 0 Then
        wsMain.Range("K3").Value = curRate
    ElseIf IsEmpty(wsMain.Range("K3").Value) Then
        wsMain.Range("K3").Value = 3.6
    End If
    wsMain.Range("K3").NumberFormat = "0.0000"
    
    eurRate = FetchBOIMonthlyAvg("EUR", "")
    If eurRate > 0 Then
        wsMain.Range("K4").Value = eurRate
    ElseIf IsEmpty(wsMain.Range("K4").Value) Then
        wsMain.Range("K4").Value = 3.9
    End If
    wsMain.Range("K4").NumberFormat = "0.0000"
    wsMain.Range("K3:K4").HorizontalAlignment = xlCenter
    
    matachMsg = ChrW(1492) & ChrW(1513) & ChrW(1506) & ChrW(1512) & " " & _
        ChrW(1502) & ChrW(1514) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1504) & " " & _
        ChrW(1506) & ChrW(34) & ChrW(1508) & " " & _
        ChrW(1489) & ChrW(1504) & ChrW(1511) & " " & _
        ChrW(1497) & ChrW(1513) & ChrW(1512) & ChrW(1488) & ChrW(1500) & " " & _
        ChrW(1488) & ChrW(1501) & " " & _
        ChrW(1500) & ChrW(1488) & " " & _
        ChrW(1497) & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & _
        ChrW(1513) & ChrW(1506) & ChrW(1512) & " " & _
        ChrW(1497) & ChrW(1513) & ChrW(1502) & ChrW(1513) & " " & _
        ChrW(1492) & ChrW(1513) & ChrW(1506) & ChrW(1512) & " " & _
        ChrW(1492) & ChrW(1512) & ChrW(1513) & ChrW(1493) & ChrW(1501) & " " & _
        ChrW(1499) & ChrW(1488) & ChrW(1503)
    wsMain.Range("J5:L7").UnMerge
    wsMain.Range("J5:K7").ClearContents
    wsMain.Range("J5").Value = matachMsg
    wsMain.Range("J5:K7").Merge
    wsMain.Range("J5").WrapText = True
    wsMain.Range("J5").Font.Size = 11
    wsMain.Range("J5").Font.Color = RGB(0, 70, 140)
    wsMain.Range("J5").ShrinkToFit = False
    wsMain.Range("J5").VerticalAlignment = xlCenter
    wsMain.Range("J5").HorizontalAlignment = xlCenter
    wsMain.Range("J5:K7").Interior.Color = RGB(220, 230, 255)
    
    ' ---- Formatting Columns ----
    wsMain.Columns("A").ColumnWidth = 43
    wsMain.Columns("B").ColumnWidth = 8
    wsMain.Columns("C").ColumnWidth = 50
    wsMain.Columns("D").ColumnWidth = 8
    wsMain.Columns("E").ColumnWidth = 8
    wsMain.Columns("F").ColumnWidth = 14
    wsMain.Columns("G").ColumnWidth = 14
    wsMain.Columns("H").ColumnWidth = 13
    wsMain.Columns("I").ColumnWidth = 13
    wsMain.Columns("J").ColumnWidth = 12
    wsMain.Columns("K").ColumnWidth = 12
    wsMain.Columns("L").ColumnWidth = 45
    wsMain.Rows("1:24").RowHeight = 22
    wsMain.Rows("1").RowHeight = 60
    

    ' ---- Borders ----
    For Each cellBorder In wsMain.Range("F2:G10")
        For Each bEdge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
            cellBorder.Borders(bEdge).LineStyle = xlContinuous
            cellBorder.Borders(bEdge).Color = RGB(0, 70, 140)
            cellBorder.Borders(bEdge).Weight = xlThin
        Next bEdge
    Next cellBorder
    For Each cellBorder In wsMain.Range("J2:K4")
        For Each bEdge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
            cellBorder.Borders(bEdge).LineStyle = xlContinuous
            cellBorder.Borders(bEdge).Color = RGB(0, 70, 140)
            cellBorder.Borders(bEdge).Weight = xlThin
        Next bEdge
    Next cellBorder
    wsMain.Range("J5:K7").BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(0, 70, 140)
    
    ' ---- Remove ALL old buttons ----
    On Error Resume Next
    For Each s In wsMain.Shapes
        s.Delete
    Next s
    Err.Clear
    On Error GoTo ERR_HANDLER
    
    ' ---- Buttons Column C (1 to 6) ----
    Dim btnLeft As Double, btnW As Double, btnH As Double
    btnW = 140
    btnH = 35
    btnLeft = wsMain.Range("C2").Left + (wsMain.Range("C2").Width - btnW) / 2
    
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top, btnW, btnH)
    shp.Name = "btnBuildReview"
    shp.Fill.ForeColor.RGB = blueClr
    shp.TextFrame2.TextRange.Text = "1 - " & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1514) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "BuildReview"
    
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 44, btnW, btnH)
    shp.Name = "btnApplyCorrections"
    shp.Fill.ForeColor.RGB = RGB(0, 150, 80)
    shp.TextFrame2.TextRange.Text = "2 - " & ChrW(1497) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1501) & " " & ChrW(1493) & ChrW(1491) & ChrW(1493) & ChrW(34) & ChrW(1495) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "ApplyCorrectionsAndBuildReports"
    
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 88, btnW, btnH)
    shp.Name = "btnBuildPresentation"
    shp.Fill.ForeColor.RGB = RGB(200, 100, 0)
    shp.TextFrame2.TextRange.Text = "3 - " & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1493) & ChrW(1512) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "BuildPresentation"
    
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 132, btnW, btnH)
    shp.Name = "btnSaveReports"
    shp.Fill.ForeColor.RGB = blueClr
    shp.TextFrame2.TextRange.Text = "4 - " & ChrW(1513) & ChrW(1502) & ChrW(1497) & ChrW(1512) & ChrW(1514) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "SaveReportsToFolder"
    
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 176, btnW, btnH)
    shp.Name = "btnViewReports"
    shp.Fill.ForeColor.RGB = RGB(0, 150, 80)
    shp.TextFrame2.TextRange.Text = "5 - " & ChrW(1510) & ChrW(1508) & ChrW(1497) & ChrW(1497) & ChrW(1492) & " " & ChrW(1489) & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "ViewReportsFolder"
    
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 220, btnW, btnH)
    shp.Name = "btnNewClients"
    shp.Fill.ForeColor.RGB = RGB(200, 100, 0)
    shp.TextFrame2.TextRange.Text = "6 - " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NewClients"

    ' ---- Helper buttons ----
    Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F11").Left + (wsMain.Range("F11").Width - 60) / 2, wsMain.Range("F11").Top + 5, 60, 25)
    btnSearch.Name = "btnSearchClient"
    btnSearch.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513)
    btnSearch.TextFrame2.TextRange.Font.Size = 10
    btnSearch.TextFrame2.TextRange.Font.Bold = msoTrue
    btnSearch.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    btnSearch.Fill.ForeColor.RGB = RGB(0, 150, 80)
    btnSearch.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btnSearch.OnAction = "SearchClientName"
    
    Set btnAll = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G11").Left + (wsMain.Range("G11").Width - 60) / 2, wsMain.Range("G11").Top + 5, 60, 25)
    btnAll.Name = "btnAllClients"
    btnAll.TextFrame2.TextRange.Text = ChrW(1499) & ChrW(1493) & ChrW(1500) & ChrW(1501)
    btnAll.TextFrame2.TextRange.Font.Size = 10
    btnAll.TextFrame2.TextRange.Font.Bold = msoTrue
    btnAll.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    btnAll.Fill.ForeColor.RGB = blueClr
    btnAll.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btnAll.OnAction = "ResetClientFilter"
    
    Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnAll.Left + (btnSearch.Left + btnSearch.Width - btnAll.Left - 120) / 2, wsMain.Range("F11").Top + 32, 120, 25)
    btnReset.Name = "btnResetDefaults"
    btnReset.TextFrame2.TextRange.Text = ChrW(1488) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1505) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
    btnReset.TextFrame2.TextRange.Font.Size = 11
    btnReset.TextFrame2.TextRange.Font.Bold = msoTrue
    btnReset.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    btnReset.Fill.ForeColor.RGB = RGB(180, 50, 50)
    btnReset.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btnReset.OnAction = "ResetHomeDefaults"
    
    ' Show/Hide Sheets in L19
    Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("L19").Left, wsMain.Range("L19").Top, 160, 30)
    shp.Name = "btnShowHidden"
    shp.Fill.ForeColor.RGB = RGB(80, 80, 80)
    shp.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1510) & ChrW(1490) & "/" & ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1512) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "ToggleHiddenSheets"
    
    ' ---- Add Exit System Button (Row 21) ----
    On Error Resume Next
    wsMain.Shapes("btnNavExit").Delete
    On Error GoTo 0
    Dim shpExit As Shape
    Set shpExit = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("A21").Left + 60, wsMain.Range("A21").Top, 120, 25)
    shpExit.Name = "btnNavExit"
    shpExit.Fill.ForeColor.RGB = RGB(180, 0, 0)
    shpExit.TextFrame2.TextRange.Text = ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1502) & ChrW(1492) & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514)
    shpExit.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shpExit.TextFrame2.TextRange.Font.Size = 10
    shpExit.TextFrame2.TextRange.Font.Bold = msoTrue
    shpExit.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shpExit.OnAction = "ExitSystem"

    ' ---- Write Credit Text to F20 ----
    wsMain.Range("F20").Value = ChrW(1504) & ChrW(1489) & ChrW(1504) & ChrW(1492) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1497) & ChrW(1491) & ChrW(1497) & " " & ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & " 054-6677396 - v" & APP_VERSION
        wsMain.Range("F20").Font.Size = 10
    wsMain.Range("F20").Font.Color = RGB(0, 0, 0)
    wsMain.Range("F20").Font.Bold = True
    wsMain.Range("F20").HorizontalAlignment = -4108

    
    ' ---- rngFILES_FOLDER and rngREPORTS_FOLDER ----
    On Error Resume Next
    existingPath = Trim$(CStr(ThisWorkbook.Names("rngFILES_FOLDER").RefersToRange.Value2))
    If existingPath = "" Then
        ThisWorkbook.Names("rngFILES_FOLDER").Delete
        Err.Clear
        ThisWorkbook.Names.Add "rngFILES_FOLDER", wsMgmt.Range("B176")
    End If
    existingPath = Trim$(CStr(ThisWorkbook.Names("rngREPORTS_FOLDER").RefersToRange.Value2))
    If existingPath = "" Then
        ThisWorkbook.Names("rngREPORTS_FOLDER").Delete
        Err.Clear
        ThisWorkbook.Names.Add "rngREPORTS_FOLDER", wsMgmt.Range("B177")
    End If
    Err.Clear
    On Error GoTo ERR_HANDLER
    
    ' ---- Store Hebrew message texts in Messages section (A222+) ----
    ' Msg1=header, Msg2=done+found, Msg3=issues, Msg4=done ok, Msg5=line, Msg6=error
    ' Msg7=step, Msg8=confirm title, Msg9=btn2 confirm, Msg10=setup err, Msg11=dropdown err, Msg12=setup ok
    ' Msg13-Msg16=confirmation lines 1-4
    msgBaseRow = ThisWorkbook.Names("rngSection_Messages").RefersToRange.Row  ' = 221
1790  wsMgmt.Cells(msgBaseRow + 1, 1).Value = ChrW(1492) & ChrW(1493) & ChrW(1491) & ChrW(1506) & ChrW(1493) & ChrW(1514) & " " & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514)
1800  wsMgmt.Cells(msgBaseRow + 2, 1).Value = ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & "-" & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " "
1810  wsMgmt.Cells(msgBaseRow + 3, 1).Value = " " & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1493) & ChrW(1514)
1820  wsMgmt.Cells(msgBaseRow + 4, 1).Value = ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492)
1830  wsMgmt.Cells(msgBaseRow + 5, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1492) & " "
1840  wsMgmt.Cells(msgBaseRow + 6, 1).Value = ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " "
1850  wsMgmt.Cells(msgBaseRow + 7, 1).Value = ChrW(1513) & ChrW(1500) & ChrW(1489) & ":" & " "
1860  wsMgmt.Cells(msgBaseRow + 8, 1).Value = ChrW(1488) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1512) & " " & ChrW(1500) & ChrW(1508) & ChrW(1504) & ChrW(1497) & " " & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491)
1870  wsMgmt.Cells(msgBaseRow + 9, 1).Value = ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1500) & ChrW(1497) & ChrW(1497) & ChrW(1513) & ChrW(1501) & " " & ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1512) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & "?"
1880  wsMgmt.Cells(msgBaseRow + 10, 1).Value = ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1514) & " " & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514) & ":" & " "
1890  wsMgmt.Cells(msgBaseRow + 11, 1).Value = ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1493) & ChrW(1503) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492) & ":" & " "
1900  wsMgmt.Cells(msgBaseRow + 12, 1).Value = ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514) & " " & ChrW(1492) & ChrW(1493) & ChrW(1511) & ChrW(1501) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!"
1910  wsMgmt.Cells(msgBaseRow + 13, 1).Value = "1" & "." & " " & ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1492) & ChrW(1506) & ChrW(1500) & ChrW(1497) & ChrW(1514) & " " & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509) & " " & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1508) & ChrW(1492) & " " & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1504) & ChrW(1497) & "?"
1920  wsMgmt.Cells(msgBaseRow + 14, 1).Value = "2" & "." & " " & ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1513) & ChrW(1502) & ChrW(1512) & ChrW(1514) & " " & ChrW(1488) & ChrW(1514) & " " & _
    ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1488) & ChrW(1495) & ChrW(1512) & ChrW(1493) & ChrW(1503) & " " & "(" & ChrW(1492) & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & _
    ChrW(1489) & ChrW(1495) & ChrW(1493) & ChrW(1489) & ChrW(1512) & ChrW(1514) & " " & ChrW(1494) & ChrW(1493) & " " & ChrW(1497) & ChrW(1502) & ChrW(1495) & ChrW(1511) & ChrW(1493) & " " & _
    ChrW(1489) & ChrW(1514) & ChrW(1492) & ChrW(1500) & ChrW(1497) & ChrW(1498) & " " & ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & ")"
1930  wsMgmt.Cells(msgBaseRow + 15, 1).Value = "3" & "." & " " & ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(1514) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1508) & ChrW(1512) & ChrW(1496) & ChrW(1497) & " " & ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1504) & ChrW(1491) & ChrW(1512) & ChrW(1513) & " " & ChrW(1489) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514) & "?"
1940  wsMgmt.Cells(msgBaseRow + 16, 1).Value = "4" & "." & " " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " " & ChrW(1506) & ChrW(1500) & " " & "O" & "K" & " " & ChrW(1500) & ChrW(1492) & ChrW(1502) & ChrW(1513) & ChrW(1498)
    ' S17 = "don't show this message again?" text
1950  wsMgmt.Cells(msgBaseRow + 17, 1).Value = ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1500) & ChrW(1492) & ChrW(1510) & ChrW(1497) & ChrW(1490) & " " & ChrW(1492) & ChrW(1493) & ChrW(1491) & ChrW(1506) & ChrW(1492) & " " & ChrW(1494) & ChrW(1493) & " " & ChrW(1513) & ChrW(1493) & ChrW(1489) & "?"
    ' S20 = flag: "1" means don't show confirmation again (empty = show)
    ' Don't overwrite S20 if already set

    ' ---- Set Error_Email parameter if not exists ----
1960  paramStartRow = ThisWorkbook.Names("rngSection_Params").RefersToRange.Row + 1
    ' Find last param row (scan until EOD or empty)
1965  paramLastRow = paramStartRow
      Do While UCase$(Trim$(CStr(wsMgmt.Cells(paramLastRow, COL_PARAM_NAME).Value2))) <> EOD_MARKER _
             And Trim$(CStr(wsMgmt.Cells(paramLastRow, COL_PARAM_NAME).Value2)) <> ""
          paramLastRow = paramLastRow + 1
      Loop
      paramLastRow = paramLastRow - 1  ' last actual data row
1970  foundEmail = False
1980  For pr = paramStartRow To paramLastRow
1990  If UCase$(Trim$(CStr(wsMgmt.Cells(pr, COL_PARAM_NAME).Value2))) = "ERROR_EMAIL" Then foundEmail = True: Exit For
2000  Next pr
2010  If Not foundEmail Then
2020  wsMgmt.Cells(paramLastRow + 1, COL_PARAM_NAME).Value = "ERROR_EMAIL"
2030  wsMgmt.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = "zvi@gorentech.co.il"
2040  paramLastRow = paramLastRow + 1
2050  End If

    ' ---- Set BACKUP_PATH parameter if not exists ----
2060  foundBackup = False
2070  For pr = paramStartRow To paramLastRow
2080  If UCase$(Trim$(CStr(wsMgmt.Cells(pr, COL_PARAM_NAME).Value2))) = "BACKUP_PATH" Then foundBackup = True: Exit For
2090  Next pr
2100  If Not foundBackup Then
2110  wsMgmt.Cells(paramLastRow + 1, COL_PARAM_NAME).Value = "BACKUP_PATH"
2120  wsMgmt.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = "C:\LEVAV PROJECT\BACKUPS"
2130  End If

    ' ---- Set RTL and font size 14 for ALL sheets in workbook ----
3200  For Each wsLoop In ThisWorkbook.Worksheets
3210  wsLoop.DisplayRightToLeft = True
3220  If wsLoop.Name <> CONTROL_SHEET_NAME() Then wsLoop.Cells.Font.Size = 14
3230  Next wsLoop

    ' ---- Clean old versions formatting ----
    On Error Resume Next
    wsMain.Range("R1:T21").ClearContents
    ' Removed ClearContents to protect credit text in A23
    Err.Clear
    On Error GoTo ERR_HANDLER

    ' ---- Set matach sheet tab color (brown) and hide it ----
3960  On Error Resume Next
3970  Set wsMatachSetup = ThisWorkbook.Worksheets(MATACH_SHEET_NAME())
3980  Err.Clear
3990  On Error GoTo ERR_HANDLER
4000  If Not wsMatachSetup Is Nothing Then
4010  wsMatachSetup.Tab.Color = RGB(139, 90, 43)
4020  wsMatachSetup.Visible = xlSheetVeryHidden
4030  End If
    ' ---- Setup section Named Ranges (vertical layout) ----
4035  SetupSectionNamedRanges
    ' ---- Setup settings navigation menu ----
4040  SetupSettingsMenu
    ' ---- Set tab colors for ALL sheets ----
4050  Call SetAllTabColors
    ' ---- Navigate to A1 ----
4060  wsMain.Activate
4070  Application.Goto wsMain.Range("A1")
4080  Application.EnableEvents = True
4090  MsgBoxU wsMgmt.Cells(msgBaseRow + 12, 1).Value, vbInformation
4100  Exit Sub
4110 ERR_HANDLER:
4120  Application.EnableEvents = True
4130  If Err.Number <> 0 Then
4140  MsgBoxU wsMgmt.Cells(msgBaseRow + 10, 1).Value & Err.Description & " (Line: " & Erl & ")", vbCritical
4150  End If

    CheckUserPermissions

End Sub

' ============================================================================
' EVENT HANDLER: Goes in the Main sheet module (Sheet code)
' Call UpdatePeriodDropdown from Worksheet_Change when B4 changes
' This sub updates E4 validation based on B4 period type selection
' ============================================================================
Public Sub UpdatePeriodDropdown()

10      Dim wsMain As Worksheet
        Dim periodType As String
        Dim listName As String

20      On Error GoTo ERR_HANDLER

30      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())

40      periodType = Trim$(CStr(wsMain.Range("rngPeriodType").Value2))

        ' "bechar/i" default text
        Dim selectText As String
45      selectText = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)

        ' Clear rngPeriodValue
50      wsMain.Range("rngPeriodValue").Value = ""
60      On Error Resume Next
70      wsMain.Range("rngPeriodValue").Validation.Delete
80      On Error GoTo ERR_HANDLER

        ' Check chatzi shnati BEFORE shnatit (shnatit is substring of chatzi shnati)
        ' "chatzi shnati" = half yearly
90      If InStr(1, periodType, ChrW(1495) & ChrW(1510) & ChrW(1497), vbTextCompare) > 0 Then
100         listName = "lst_half_year"
        ' "riv'oni" = quarterly
110     ElseIf InStr(1, periodType, ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497), vbTextCompare) > 0 Then
120         listName = "lst_quarter"
        ' "chodshi" = monthly
130     ElseIf InStr(1, periodType, ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497), vbTextCompare) > 0 Then
140         listName = "lst_month"
        ' "shnatit" = yearly -> no second dropdown needed, jump to DateType
150     ElseIf InStr(1, periodType, ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497), vbTextCompare) > 0 Then
155         wsMain.Range("rngPeriodValue").Value = ""
            On Error Resume Next
160         wsMain.Activate
161         Application.Goto wsMain.Range("rngDateType")
            On Error GoTo ERR_HANDLER
            GoTo CLEAN_EXIT
170     Else
            GoTo CLEAN_EXIT
190     End If

200     wsMain.Range("rngPeriodValue").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & listName
        ' Set default value and jump cursor
205     wsMain.Range("rngPeriodValue").Value = selectText
        On Error Resume Next
206     wsMain.Activate
207     Application.Goto wsMain.Range("rngPeriodValue")
        On Error GoTo ERR_HANDLER

        ' Reset background to paste color
        On Error Resume Next
        wsMain.Range("rngPeriodValue").Interior.Color = RGB(255, 245, 230)
        On Error GoTo 0
CLEAN_EXIT:

210     Exit Sub

ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True

End Sub

' ============================================================================
' EVENT HANDLER: UpdateFilterValueDropdown
' Called from Worksheet_Change when G9 (rngFilterType) changes
' Reads unique values from hidden "reshimot" sheet and sets G10 validation
' ============================================================================
Public Sub UpdateFilterValueDropdown()

10      Dim wsMain As Worksheet
        Dim wsLists As Worksheet
        Dim filterType As String
        Dim listsName As String
        Dim targetCol As Long
        Dim lastR As Long
        Dim valList As String
        Dim i As Long
        Dim selectText As String

20      On Error GoTo ERR_HANDLER

30      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        On Error Resume Next
        wsMain.Unprotect "Z961814r"
        On Error GoTo ERR_HANDLER

40      filterType = Trim$(CStr(wsMain.Range("rngFilterType").Value2))

        ' Clear rngFilterValue
50      wsMain.Range("rngFilterValue").Value = ""
60      On Error Resume Next
70      wsMain.Range("rngFilterValue").Validation.Delete
80      On Error GoTo ERR_HANDLER

        ' "bechar/i" text
90      selectText = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)

        ' If empty or "bechar/i" - just clear and exit
100     If filterType = "" Or filterType = selectText Then GoTo CLEAN_EXIT

        ' Determine which column in the lists sheet to read
        ' chevra=1, teler=2, sochen=3, anaf=4, anaf merkaz=5
110     If filterType = H_COMPANY() Then
120         targetCol = 1
130     ElseIf filterType = H_TELLER() Then
140         targetCol = 2
150     ElseIf filterType = H_AGENT() Then
160         targetCol = 3
170     ElseIf filterType = H_BRANCH() Then
180         targetCol = 4
190     ElseIf filterType = H_BRANCH() & ChrW(32) & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494) Then
200         targetCol = 5
210     Else
220         GoTo CLEAN_EXIT
230     End If

        ' Find the lists sheet
240     listsName = H_RESHIMOT()
250     If Not SheetExists(listsName) Then
260         MsgBoxU ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & listsName & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ". " & ChrW(1492) & ChrW(1512) & ChrW(1509) & " " & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " 1 " & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & ChrW(1492) & ".", vbExclamation
270         GoTo CLEAN_EXIT
280     End If

290     Set wsLists = ThisWorkbook.Worksheets(listsName)
300     lastR = wsLists.Cells(wsLists.Rows.Count, targetCol).End(xlUp).Row

310     If lastR < 2 Then GoTo CLEAN_EXIT

        ' Build comma-separated list
320     valList = ""
330     For i = 2 To lastR
340         If Trim$(CStr(wsLists.Cells(i, targetCol).Value2)) <> "" Then
350             If valList <> "" Then valList = valList & ","
360             valList = valList & Trim$(CStr(wsLists.Cells(i, targetCol).Value2))
370         End If
380     Next i

390     If valList = "" Then GoTo CLEAN_EXIT

        ' Add validation list to G10
400     wsMain.Range("rngFilterValue").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=valList
        ' Set default value and jump cursor
405     wsMain.Range("rngFilterValue").Value = selectText
        On Error Resume Next
406     wsMain.Activate
407     Application.Goto wsMain.Range("rngFilterValue")
        On Error GoTo ERR_HANDLER

        ' Reset G10 background to green (prevent gold/purple artifact)
        On Error Resume Next
        wsMain.Range("rngFilterValue").Interior.Color = RGB(220, 240, 220)
        On Error GoTo 0

CLEAN_EXIT:
        On Error Resume Next
        If Not wsMain Is Nothing Then wsMain.Protect Password:="Z961814r", UserInterfaceOnly:=True
410     Exit Sub

ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
420     MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1493) & ChrW(1503) & " " & ChrW(1505) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1503) & ": " & Err.Description, vbCritical
        On Error Resume Next
        If Not wsMain Is Nothing Then wsMain.Protect Password:="Z961814r", UserInterfaceOnly:=True
End Sub

' ============================================================================
' MACRO 3: BuildPresentation
' Creates a PPTX management presentation from comparison sheets
' v5: Landscape, split charts/tables, page numbers, insured column
' Phase 1: Export charts as images (Excel only)
' Phase 2: Build PowerPoint slides from images + data
' ============================================================================
Public Sub BuildPresentation()

10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False

        ' Remove any leftover sheet protection
        Dim wsUp3 As Worksheet
        For Each wsUp3 In ThisWorkbook.Worksheets
            On Error Resume Next
            wsUp3.Unprotect "Z961814r"
            On Error GoTo ERR_HANDLER
        Next wsUp3

        Dim wsMain As Worksheet
20      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())

        Dim yearVal As String
        Dim refYear As String
        Dim periodType As String
        Dim periodDetail As String
        Dim periodDesc As String
40      yearVal = Trim$(CStr(wsMain.Range("rngCurrentYear").Value2))
50      refYear = Trim$(CStr(wsMain.Range("rngBaseYear").Value2))
60      periodType = Trim$(CStr(wsMain.Range("rngPeriodType").Value2))
70      periodDetail = Trim$(CStr(wsMain.Range("rngPeriodValue").Value2))
80      periodDesc = periodType
90      If periodDetail <> "" Then periodDesc = periodDetail

        ' Build parameters subtitle for all slides
        Dim dateType As String
        Dim detailBy As String
        Dim clientName As String
        Dim paramsSubtitle As String
        dateType = Trim$(CStr(wsMain.Range("G7").Value2))
        detailBy = Trim$(CStr(wsMain.Range("rngFilterValue").Value2))
        clientName = Trim$(CStr(wsMain.Range("rngClientName").Value2))
        ' Build subtitle: periodDesc | dateType | detailBy | clientName (no years - already in chart)
        Dim bacharI As String
        bacharI = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)  ' "???/?"
        paramsSubtitle = ""
        If periodDesc <> "" And periodDesc <> bacharI Then paramsSubtitle = periodDesc
        If dateType <> "" And dateType <> bacharI Then paramsSubtitle = paramsSubtitle & " | " & dateType
        If detailBy <> "" And detailBy <> bacharI Then paramsSubtitle = paramsSubtitle & " | " & detailBy
        If clientName <> "" And clientName <> bacharI Then paramsSubtitle = paramsSubtitle & " | " & clientName
        ' Remove leading " | " if first param was empty
        If Left$(paramsSubtitle, 3) = " | " Then paramsSubtitle = Mid$(paramsSubtitle, 4)

        ' --- Check if presentation already exists ---
        Dim reportsPath3 As String
        reportsPath3 = Trim$(CStr(Range("rngREPORTS_FOLDER").Value2))
        If Len(Dir(reportsPath3 & "\*" & yearVal & "*.pptx")) > 0 Then
            Dim alreadyMsg3 As String
            ' "hamitzget kvar keiyemet. nitan litzpot ba bekaftor 5. ha'im leharitz shuv?"
            alreadyMsg3 = ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1499) & ChrW(1489) & ChrW(1512) & " " & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1502) & ChrW(1514) & "." & vbCrLf & _
                ChrW(1504) & ChrW(1497) & ChrW(1514) & ChrW(1503) & " " & ChrW(1500) & ChrW(1510) & ChrW(1508) & ChrW(1493) & ChrW(1514) & " " & ChrW(1489) & ChrW(1492) & " " & ChrW(1489) & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " 5." & vbCrLf & vbCrLf & _
                ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1500) & ChrW(1492) & ChrW(1512) & ChrW(1497) & ChrW(1509) & " " & ChrW(1513) & ChrW(1493) & ChrW(1489) & "?"
            If MsgBoxU(alreadyMsg3, vbYesNo + vbQuestion) <> vbYes Then
                Exit Sub
            End If
        End If

        ' Validate that comparison sheets exist
100     If Not SheetExists(SHEET_COMPANIES()) Then
110         MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1492) & ChrW(1512) & ChrW(1497) & ChrW(1509) & " " & ChrW(1511) & ChrW(1493) & ChrW(1491) & ChrW(1501) & " " & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " 2", vbCritical
120         Exit Sub
130     End If

        ' Show processing message (below currency area)
        wsMain.Unprotect "Z961814r"
        wsMain.Range("G15:K15").UnMerge
        wsMain.Range("G15:K15").Merge
140     With wsMain.Range("G15")
150         .Value = ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1512) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1488) & ChrW(1504) & ChrW(1488) & " " & ChrW(1492) & ChrW(1502) & ChrW(1514) & ChrW(1497) & ChrW(1504) & ChrW(1493) & "..."
160         .Font.Size = 18
170         .Font.Bold = True
180         .Font.Color = RGB(200, 0, 0)
190         .Interior.Color = RGB(255, 255, 200)
            .HorizontalAlignment = -4108
200     End With
210     Application.ScreenUpdating = True
220     DoEvents
        Application.ScreenUpdating = False

        ' ================================================================
        ' PHASE 1: Create all chart images in Excel (NO PowerPoint yet)
        ' ================================================================
        ' Unhide all sheets before accessing data (in case some are VeryHidden)
        Dim wsTmp3 As Worksheet
        Dim hiddenSheets3() As String
        Dim hiddenCount3 As Long
        hiddenCount3 = 0
        ReDim hiddenSheets3(1 To ThisWorkbook.Worksheets.Count)
        For Each wsTmp3 In ThisWorkbook.Worksheets
            If wsTmp3.Visible <> xlSheetVisible Then
                hiddenCount3 = hiddenCount3 + 1
                hiddenSheets3(hiddenCount3) = wsTmp3.Name
                wsTmp3.Visible = xlSheetVisible
            End If
        Next wsTmp3
        
        Dim tmpPath As String
230     tmpPath = Environ$("TEMP") & "\"

        Dim imgTotal As String
240     imgTotal = tmpPath & "levav_total.gif"
250     ExportTotalChart imgTotal, yearVal, refYear

        ' Build list of sheets to process
        Dim sheetList(1 To 6) As String
        Dim sheetCount As Long
260     sheetCount = 0

270     If SheetExists(SHEET_MONTHS()) Then
280         sheetCount = sheetCount + 1
290         sheetList(sheetCount) = SHEET_MONTHS()
300     End If
310     If SheetExists(SHEET_COMPANIES()) Then
320         sheetCount = sheetCount + 1
330         sheetList(sheetCount) = SHEET_COMPANIES()
340     End If
350     If SheetExists(SHEET_MAINBRANCH()) Then
360         sheetCount = sheetCount + 1
370         sheetList(sheetCount) = SHEET_MAINBRANCH()
380     End If
390     If SheetExists(SHEET_TELLERS()) Then
400         sheetCount = sheetCount + 1
410         sheetList(sheetCount) = SHEET_TELLERS()
420     End If
430     If SheetExists(SHEET_AGENTS()) Then
440         sheetCount = sheetCount + 1
450         sheetList(sheetCount) = SHEET_AGENTS()
460     End If

        ' Export 4 charts per sheet (prem, comm, docs, insured)
        Dim si As Long
        Dim imgFiles() As String
470     ReDim imgFiles(1 To sheetCount * 4)
        Dim exportOK() As Boolean
480     ReDim exportOK(1 To sheetCount)
490     For si = 1 To sheetCount
500         imgFiles(si * 4 - 3) = tmpPath & "levav_prem_" & si & ".gif"
510         imgFiles(si * 4 - 2) = tmpPath & "levav_comm_" & si & ".gif"
            imgFiles(si * 4 - 1) = tmpPath & "levav_docs_" & si & ".gif"
            imgFiles(si * 4) = tmpPath & "levav_ins_" & si & ".gif"
520         On Error Resume Next
530         ExportCompCharts sheetList(si), imgFiles(si * 4 - 3), imgFiles(si * 4 - 2), yearVal, refYear, "", imgFiles(si * 4 - 1), imgFiles(si * 4)
540         If Err.Number = 0 Then
550             exportOK(si) = True
560         Else
570             exportOK(si) = False
572             Err.Clear
580         End If
590         On Error GoTo ERR_HANDLER
600     Next si

605     DoEvents

        ' ================================================================
        ' PHASE 2: Open PowerPoint and build slides
        ' ================================================================
        Dim ppApp As Object
        Dim ppPres As Object
        Dim ppSlide As Object
        Dim ppWeOwnApp As Boolean
        ppWeOwnApp = False
        ' Try to use existing PowerPoint instance first
610     On Error Resume Next
        Set ppApp = GetObject(, "PowerPoint.Application")
        On Error GoTo ERR_HANDLER
        If ppApp Is Nothing Then
            Set ppApp = CreateObject("PowerPoint.Application")
            ppWeOwnApp = True
        End If
615     ppApp.Visible = True
620     Set ppPres = ppApp.Presentations.Add

        ' Set LANDSCAPE slide size (13.33" x 7.5")
625     ppPres.PageSetup.SlideWidth = 960
630     ppPres.PageSetup.SlideHeight = 540

        Dim slideIdx As Long
        Dim slideW As Single
        Dim slideH As Single
635     slideIdx = 0
640     slideW = 960
645     slideH = 540

        ' Slide title names (Hebrew)
        Dim titleNames(1 To 6) As String
650     titleNames(1) = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
655     titleNames(2) = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1514)
660     titleNames(3) = H_BRANCH() & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)
665     titleNames(4) = H_TELLER() & ChrW(1497) & ChrW(1493) & ChrW(1514)
670     titleNames(5) = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501)

        ' SLIDE 1: Title
675     slideIdx = slideIdx + 1
680     Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
685     BuildTitleSlide ppSlide, yearVal, refYear, periodDesc, slideW, slideH, paramsSubtitle

        ' SLIDE 2: Total Summary chart
690     slideIdx = slideIdx + 1
695     Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
700     BuildTotalSlideFromImage ppSlide, imgTotal, yearVal, refYear, slideW, paramsSubtitle

        ' For each comparison sheet: 5 slides (prem, comm, docs, insured charts + table)
730     For si = 1 To sheetCount
740         If exportOK(si) Then
                ' Slide: Premium chart
750             slideIdx = slideIdx + 1
760             Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
770             BuildChartSlide ppSlide, imgFiles(si * 4 - 3), ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & titleNames(si), yearVal, refYear, slideW, paramsSubtitle
                ' Slide: Commission chart
780             slideIdx = slideIdx + 1
790             Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
800             BuildChartSlide ppSlide, imgFiles(si * 4 - 2), ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & titleNames(si), yearVal, refYear, slideW, paramsSubtitle
                ' Slide: Documents chart
                slideIdx = slideIdx + 1
                Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
                ' "mismachim lefi" = documents by
                BuildChartSlide ppSlide, imgFiles(si * 4 - 1), ChrW(1502) & ChrW(1505) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & titleNames(si), yearVal, refYear, slideW, paramsSubtitle
                ' Slide: Insured persons chart
                slideIdx = slideIdx + 1
                Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
                ' "mevutachim lefi" = insured by
                BuildChartSlide ppSlide, imgFiles(si * 4), ChrW(1502) & ChrW(1489) & ChrW(1493) & ChrW(1496) & ChrW(1495) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & titleNames(si), yearVal, refYear, slideW, paramsSubtitle
810         End If
            ' Slide: Data table (always, even if charts failed)
820         slideIdx = slideIdx + 1
830         Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
840         BuildTableSlide ppSlide, sheetList(si), titleNames(si), yearVal, refYear, slideW, slideH, paramsSubtitle
850     Next si

        ' ================================================================
        ' PHASE 3: "Agents without Levav" variant slides (2 chart slides)
        ' ================================================================
        If SheetExists(SHEET_AGENTS()) Then
            Dim imgNoLevavPrem As String
            Dim imgNoLevavComm As String
            Dim noLevavOK As Boolean
            ' levav = ChrW(1500) & ChrW(1489) & ChrW(1489)
            Dim levavName As String
            levavName = ChrW(1500) & ChrW(1489) & ChrW(1489)
            imgNoLevavPrem = tmpPath & "levav_nolev_prem.gif"
            imgNoLevavComm = tmpPath & "levav_nolev_comm.gif"
            noLevavOK = False
            On Error Resume Next
            ExportCompCharts SHEET_AGENTS(), imgNoLevavPrem, imgNoLevavComm, yearVal, refYear, levavName
            If Err.Number = 0 Then noLevavOK = True
            Err.Clear
            On Error GoTo ERR_HANDLER
            If noLevavOK Then
                ' Slide: Premiums without Levav
                slideIdx = slideIdx + 1
                Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
                ' title: premiot lefi sochnim lelo levav
                BuildChartSlide ppSlide, imgNoLevavPrem, ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1500) & ChrW(1488) & " " & ChrW(1500) & ChrW(1489) & ChrW(1489), yearVal, refYear, slideW, paramsSubtitle
                ' Slide: Commissions without Levav
                slideIdx = slideIdx + 1
                Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
                ' title: amlot lefi sochnim lelo levav
                BuildChartSlide ppSlide, imgNoLevavComm, ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1500) & ChrW(1488) & " " & ChrW(1500) & ChrW(1489) & ChrW(1489), yearVal, refYear, slideW, paramsSubtitle
            End If
            ' Cleanup temp images
            On Error Resume Next
            Kill imgNoLevavPrem
            Kill imgNoLevavComm
            On Error GoTo ERR_HANDLER
        End If

        ' Add page numbers to all slides
        Dim pg As Long
860     For pg = 1 To ppPres.Slides.Count
870         AddPageNumber ppPres.Slides(pg), pg, ppPres.Slides.Count, slideW, slideH
880     Next pg

        ' Save presentation to Reports folder
        Dim reportsFolder As String
        Dim fsoRpt As Object
        Set fsoRpt = CreateObject("Scripting.FileSystemObject")
        reportsFolder = REPORTS_FOLDER()
        If Not fsoRpt.FolderExists(reportsFolder) Then fsoRpt.CreateFolder reportsFolder
        
        Dim tmpSavePath As String
        Dim finalPptxPath As String
        Dim finalPdfPath As String
        Dim presFileName As String
        Dim bSaved As Boolean
        presFileName = ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1492) & ChrW(1504) & ChrW(1492) & ChrW(1500) & ChrW(1492) & " " & yearVal
        finalPptxPath = reportsFolder & "\" & presFileName & ".pptx"
        finalPdfPath = reportsFolder & "\" & presFileName & ".pdf"
        bSaved = False
        ' --- Attempt 1: Save directly to Reports folder ---
890     On Error Resume Next
        Err.Clear
900     ppPres.SaveAs finalPptxPath
        If Err.Number = 0 Then
            Err.Clear
906         ppPres.SaveAs finalPdfPath, 32
            bSaved = True
        End If
        ' --- Attempt 2: Save via TEMP if direct save failed ---
        If Not bSaved Then
            Err.Clear
            tmpSavePath = Environ$("TEMP") & "\" & "LevavTemp_" & yearVal & ".pptx"
            Kill tmpSavePath
            Err.Clear
            ppPres.SaveAs tmpSavePath
            If Err.Number = 0 Then
                Err.Clear
                FileCopy tmpSavePath, finalPptxPath
                If Err.Number = 0 Then bSaved = True
                Kill tmpSavePath
            End If
        End If
        On Error GoTo ERR_HANDLER        ' Start slideshow automatically
        On Error Resume Next
        If Not ppApp Is Nothing Then
            ppApp.Visible = True
            ppApp.Activate
            With ppPres.SlideShowSettings
                .ShowType = 1 ' ppShowTypeSpeaker (Full Screen)
                .ShowPresenterView = 0 ' msoFalse (Disable presenter view)
                .Run
            End With
        End If
        On Error GoTo ERR_HANDLER
910     Set ppPres = Nothing
920     Set ppApp = Nothing

        ' Cleanup temp images
950     On Error Resume Next
960     Kill imgTotal
970     For si = 1 To sheetCount * 4
980         Kill imgFiles(si)
990     Next si
1000    On Error GoTo ERR_HANDLER

        ' Re-hide sheets that were hidden before
        Dim hi3 As Long
        For hi3 = 1 To hiddenCount3
            ThisWorkbook.Worksheets(hiddenSheets3(hi3)).Visible = xlSheetVeryHidden
        Next hi3
        
        ' Restore screen updating
        Application.ScreenUpdating = True

        ' Clear processing message and restore green background
        wsMain.Unprotect "Z961814r"
        wsMain.Range("G15:K15").UnMerge
1010    With wsMain.Range("G15:K15")
1020        .Value = ""
1030        .Interior.Color = RGB(220, 240, 220)
1040    End With
        wsMain.Protect UserInterfaceOnly:=True

        ' Bring Excel back to home silently (no AppActivate so PP stays in front)
        On Error Resume Next
        wsMain.Activate
        Application.Goto wsMain.Range("A1")
        On Error GoTo ERR_HANDLER
        
        Dim askMsg As String
        If Not bSaved Then
            ' "hamatzget notzra behatzlacha. yesh lishmor yadanit."
            askMsg = ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "." & vbCrLf & _
                ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1513) & ChrW(1502) & ChrW(1493) & ChrW(1512) & " " & ChrW(1497) & ChrW(1491) & ChrW(1504) & ChrW(1497) & ChrW(1514) & "."
            MsgBoxU askMsg, vbOKOnly + vbInformation
        End If

1060    Exit Sub

ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Dim errLine As Long
        Dim errDesc As String
        Dim errNum As Long
1070    errLine = Erl
1072    errDesc = Err.Description
1074    errNum = Err.Number
1076    On Error Resume Next
        ' Re-hide sheets that were unhidden
        Dim hi3err As Long
        For hi3err = 1 To hiddenCount3
            ThisWorkbook.Worksheets(hiddenSheets3(hi3err)).Visible = xlSheetVeryHidden
        Next hi3err
        Application.ScreenUpdating = True
        wsMain.Unprotect "Z961814r"
        wsMain.Range("G15:K15").UnMerge
        With wsMain.Range("G15:K15")
            .Value = ""
            .Interior.Color = RGB(220, 240, 220)
        End With
        wsMain.Protect UserInterfaceOnly:=True
        If Not ppPres Is Nothing Then ppPres.Close
        If Not ppApp Is Nothing And ppWeOwnApp Then ppApp.Quit
        Kill imgTotal
        Dim ei As Long
        For ei = 1 To sheetCount * 4
            Kill imgFiles(ei)
        Next ei
        ' User-friendly error message
        Dim userMsg As String
        If InStr(1, errDesc, "SaveAs", vbTextCompare) > 0 Or InStr(1, errDesc, "access", vbTextCompare) > 0 Or errNum = -2147467259 Then
            ' File locked / PP open error
            userMsg = ChrW(1506) & ChrW(1500) & " " & ChrW(1502) & ChrW(1504) & ChrW(1514) & " " & _
                ChrW(1500) & ChrW(1489) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & _
                ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & vbCrLf & _
                ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1505) & ChrW(1490) & ChrW(1493) & ChrW(1512) & " " & _
                ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & _
                ChrW(1492) & ChrW(1511) & ChrW(1493) & ChrW(1491) & ChrW(1502) & ChrW(1514)
        Else
            userMsg = ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & _
                ChrW(1489) & ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1512) & ChrW(1514) & " " & _
                ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & vbCrLf & errDesc
        End If
1080    MsgBoxU userMsg, vbCritical

End Sub

' ============================================================================
' HELPER: Export Total Summary chart to image file
' ============================================================================
Private Sub ExportTotalChart(ByVal imgPath As String, ByVal yearVal As String, ByVal refYear As String)

10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False

        Dim ws As Worksheet
        Dim lastRow As Long
20      Set ws = ThisWorkbook.Worksheets(SHEET_MONTHS())
30      lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        Dim sumPR As Double
        Dim sumPC As Double
        Dim sumCR As Double
        Dim sumCC As Double
40      sumPR = CDbl(ws.Cells(lastRow, 2).Value2)
50      sumPC = CDbl(ws.Cells(lastRow, 3).Value2)
60      sumCR = CDbl(ws.Cells(lastRow, 14).Value2)
70      sumCC = CDbl(ws.Cells(lastRow, 15).Value2)

        Dim tmpWs As Worksheet
        Dim co As Object
        Dim xlCht As Object
80      Application.ScreenUpdating = False
90      Set tmpWs = ThisWorkbook.Worksheets.Add

100     tmpWs.Cells(1, 1).Value = ""
110     tmpWs.Cells(1, 2).Value = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & refYear   ' premiot refYear
120     tmpWs.Cells(1, 3).Value = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & yearVal   ' premiot yearVal
130     tmpWs.Cells(1, 4).Value = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " " & refYear   ' amlot refYear
140     tmpWs.Cells(1, 5).Value = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " " & yearVal   ' amlot yearVal
150     tmpWs.Cells(2, 1).Value = ChrW(1505) & ChrW(1499) & ChrW(1493) & ChrW(1501) & " " & ChrW(1499) & ChrW(1493) & ChrW(1500) & ChrW(1500)
160     tmpWs.Cells(2, 2).Value = sumPR
170     tmpWs.Cells(2, 3).Value = sumPC
180     tmpWs.Cells(2, 4).Value = sumCR
190     tmpWs.Cells(2, 5).Value = sumCC

200     Set co = tmpWs.ChartObjects.Add(10, 10, 600, 400)
210     Set xlCht = co.Chart
220     xlCht.ChartType = 51
230     xlCht.SetSourceData tmpWs.Range("A1:E2"), 2  ' xlColumns
240     xlCht.HasTitle = False
250     xlCht.HasLegend = True

        ' --- Colors: Yellow=ref prem, Blue=cur prem, Orange=ref comm, Green=cur comm ---
        xlCht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)     ' yellow/gold
        xlCht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(68, 114, 196)    ' blue
        xlCht.SeriesCollection(3).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)    ' orange
        xlCht.SeriesCollection(4).Format.Fill.ForeColor.RGB = RGB(112, 173, 71)    ' green

        ' --- Data labels (show in K, vertical/upward) ---
        Dim sTot As Long
        For sTot = 1 To 4
            xlCht.SeriesCollection(sTot).HasDataLabels = True
            xlCht.SeriesCollection(sTot).DataLabels.NumberFormat = "#,##0,""K"""
            xlCht.SeriesCollection(sTot).DataLabels.Font.Size = 10
            xlCht.SeriesCollection(sTot).DataLabels.Orientation = 90
        Next sTot

        ' --- Y-axis number format (in K) ---
        xlCht.Axes(2).TickLabels.NumberFormat = "#,##0,""K"""

260     xlCht.Export imgPath

270     Application.DisplayAlerts = False
280     tmpWs.Delete
290     Application.DisplayAlerts = True
300     Application.ScreenUpdating = True

310     Exit Sub
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
320     On Error Resume Next
        Application.DisplayAlerts = False
        If Not tmpWs Is Nothing Then tmpWs.Delete
        Application.DisplayAlerts = True
330     Err.Raise Err.Number, "ExportTotalChart:" & Erl, Err.Description
End Sub

' ============================================================================
' HELPER: Export Comparison charts (premiums, commissions, documents, insured)
' to 4 image files (docs and insured are optional - pass "" to skip)
' ============================================================================
Private Sub ExportCompCharts(ByVal sheetName As String, ByVal imgPrem As String, ByVal imgComm As String, ByVal yearVal As String, ByVal refYear As String, Optional ByVal excludeName As String = "", Optional ByVal imgDocs As String = "", Optional ByVal imgInsured As String = "")

10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False

        Dim ws As Worksheet
        Dim lastRow As Long
        Dim dataRows As Long
        Dim r As Long
        Dim tmpName As String
20      Set ws = ThisWorkbook.Worksheets(sheetName)
30      lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
40      dataRows = lastRow - 3

        Dim arrNames() As String
        Dim arrPremR() As Double
        Dim arrPremC() As Double
        Dim arrCommR() As Double
        Dim arrCommC() As Double
        Dim arrDocsR() As Double
        Dim arrDocsC() As Double
        Dim arrInsR() As Double
        Dim arrInsC() As Double
        Dim nItems As Long
50      nItems = dataRows - 1
60      If nItems < 1 Then Exit Sub

70      ReDim arrNames(1 To nItems)
80      ReDim arrPremR(1 To nItems)
82      ReDim arrPremC(1 To nItems)
90      ReDim arrCommR(1 To nItems)
92      ReDim arrCommC(1 To nItems)
        ReDim arrDocsR(1 To nItems)
        ReDim arrDocsC(1 To nItems)
        ReDim arrInsR(1 To nItems)
        ReDim arrInsC(1 To nItems)

        Dim idx As Long
100     idx = 0
110     For r = 4 To lastRow - 1
            tmpName = ShortenCompanyName(Trim$(CStr(ws.Cells(r, 1).Value2)))
            ' Skip excluded name (for "agents without Levav" variant)
            If Len(excludeName) > 0 Then
                If InStr(1, tmpName, excludeName, vbTextCompare) > 0 Then GoTo NEXT_ROW_ECC
            End If
120         idx = idx + 1
130         If idx > nItems Then Exit For
140         arrNames(idx) = tmpName
150         arrPremR(idx) = CDbl(ws.Cells(r, 2).Value2)
160         arrPremC(idx) = CDbl(ws.Cells(r, 3).Value2)
            arrDocsR(idx) = CDbl(ws.Cells(r, 5).Value2)
            arrDocsC(idx) = CDbl(ws.Cells(r, 6).Value2)
            arrInsR(idx) = CDbl(ws.Cells(r, 8).Value2)
            arrInsC(idx) = CDbl(ws.Cells(r, 9).Value2)
170         arrCommR(idx) = CDbl(ws.Cells(r, 14).Value2)
180         arrCommC(idx) = CDbl(ws.Cells(r, 15).Value2)
NEXT_ROW_ECC:
190     Next r
200     nItems = idx

        Dim chartItems As Long
210     chartItems = nItems
220     If chartItems > 15 Then chartItems = 15

        Dim tmpWs As Worksheet
        Dim co As Object
        Dim xlCht As Object
        Dim ci As Long
230     Application.ScreenUpdating = False
240     Set tmpWs = ThisWorkbook.Worksheets.Add

        ' ---- Chart 1: Premiums ----
250     tmpWs.Cells(1, 1).Value = ""
260     tmpWs.Cells(1, 2).Value = refYear
270     tmpWs.Cells(1, 3).Value = yearVal
280     For ci = 1 To chartItems
290         tmpWs.Cells(ci + 1, 1).Value = arrNames(ci)
300         tmpWs.Cells(ci + 1, 2).Value = arrPremR(ci)
310         tmpWs.Cells(ci + 1, 3).Value = arrPremC(ci)
320     Next ci

330     Set co = tmpWs.ChartObjects.Add(10, 10, 600, 350)
340     Set xlCht = co.Chart
350     xlCht.ChartType = 51
360     xlCht.SetSourceData tmpWs.Range(tmpWs.Cells(1, 1), tmpWs.Cells(chartItems + 1, 3)), 2  ' xlColumns
370     xlCht.HasTitle = False
380     xlCht.HasLegend = True

        ' --- Colors: Yellow=ref year, Blue=current year ---
        xlCht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 192, 0)     ' yellow/gold
        xlCht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(68, 114, 196)    ' blue

        ' --- Data labels (show in K, vertical/upward) ---
        Dim sP As Long
        For sP = 1 To 2
            xlCht.SeriesCollection(sP).HasDataLabels = True
            xlCht.SeriesCollection(sP).DataLabels.NumberFormat = "#,##0,""K"""
            xlCht.SeriesCollection(sP).DataLabels.Font.Size = 9
            xlCht.SeriesCollection(sP).DataLabels.Orientation = 90
        Next sP

        ' --- Y-axis number format (in K) ---
        xlCht.Axes(2).TickLabels.NumberFormat = "#,##0,""K"""

390     xlCht.Export imgPrem

400     tmpWs.ChartObjects.Delete
410     tmpWs.Cells.Clear

        ' ---- Chart 2: Commissions ----
420     tmpWs.Cells(1, 1).Value = ""
430     tmpWs.Cells(1, 2).Value = refYear
440     tmpWs.Cells(1, 3).Value = yearVal
450     For ci = 1 To chartItems
460         tmpWs.Cells(ci + 1, 1).Value = arrNames(ci)
470         tmpWs.Cells(ci + 1, 2).Value = arrCommR(ci)
480         tmpWs.Cells(ci + 1, 3).Value = arrCommC(ci)
490     Next ci

500     Set co = tmpWs.ChartObjects.Add(10, 10, 600, 350)
510     Set xlCht = co.Chart
520     xlCht.ChartType = 51
530     xlCht.SetSourceData tmpWs.Range(tmpWs.Cells(1, 1), tmpWs.Cells(chartItems + 1, 3)), 2  ' xlColumns
540     xlCht.HasTitle = False
550     xlCht.HasLegend = True

        ' --- Colors: Orange=ref year, Green=current year ---
        xlCht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)    ' orange
        xlCht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(112, 173, 71)    ' green

        ' --- Data labels (show in K, vertical/upward) ---
        Dim sC As Long
        For sC = 1 To 2
            xlCht.SeriesCollection(sC).HasDataLabels = True
            xlCht.SeriesCollection(sC).DataLabels.NumberFormat = "#,##0,""K"""
            xlCht.SeriesCollection(sC).DataLabels.Font.Size = 9
            xlCht.SeriesCollection(sC).DataLabels.Orientation = 90
        Next sC

        ' --- Y-axis number format (in K) ---
        xlCht.Axes(2).TickLabels.NumberFormat = "#,##0,""K"""

560     xlCht.Export imgComm

        ' ---- Chart 3: Documents (optional) ----
        If imgDocs <> "" Then
            tmpWs.ChartObjects.Delete
            tmpWs.Cells.Clear
            tmpWs.Cells(1, 1).Value = ""
            tmpWs.Cells(1, 2).Value = refYear
            tmpWs.Cells(1, 3).Value = yearVal
            For ci = 1 To chartItems
                tmpWs.Cells(ci + 1, 1).Value = arrNames(ci)
                tmpWs.Cells(ci + 1, 2).Value = arrDocsR(ci)
                tmpWs.Cells(ci + 1, 3).Value = arrDocsC(ci)
            Next ci
            Set co = tmpWs.ChartObjects.Add(10, 10, 600, 350)
            Set xlCht = co.Chart
            xlCht.ChartType = 51
            xlCht.SetSourceData tmpWs.Range(tmpWs.Cells(1, 1), tmpWs.Cells(chartItems + 1, 3)), 2
            xlCht.HasTitle = False
            xlCht.HasLegend = True
            xlCht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(180, 130, 70)    ' brown/tan
            xlCht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(91, 155, 213)    ' steel blue
            Dim sD As Long
            For sD = 1 To 2
                xlCht.SeriesCollection(sD).HasDataLabels = True
                xlCht.SeriesCollection(sD).DataLabels.NumberFormat = "#,##0"
                xlCht.SeriesCollection(sD).DataLabels.Font.Size = 9
                xlCht.SeriesCollection(sD).DataLabels.Orientation = 90
            Next sD
            xlCht.Axes(2).TickLabels.NumberFormat = "#,##0"
            xlCht.Export imgDocs
        End If

        ' ---- Chart 4: Insured persons (optional) ----
        If imgInsured <> "" Then
            tmpWs.ChartObjects.Delete
            tmpWs.Cells.Clear
            tmpWs.Cells(1, 1).Value = ""
            tmpWs.Cells(1, 2).Value = refYear
            tmpWs.Cells(1, 3).Value = yearVal
            For ci = 1 To chartItems
                tmpWs.Cells(ci + 1, 1).Value = arrNames(ci)
                tmpWs.Cells(ci + 1, 2).Value = arrInsR(ci)
                tmpWs.Cells(ci + 1, 3).Value = arrInsC(ci)
            Next ci
            Set co = tmpWs.ChartObjects.Add(10, 10, 600, 350)
            Set xlCht = co.Chart
            xlCht.ChartType = 51
            xlCht.SetSourceData tmpWs.Range(tmpWs.Cells(1, 1), tmpWs.Cells(chartItems + 1, 3)), 2
            xlCht.HasTitle = False
            xlCht.HasLegend = True
            xlCht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(128, 0, 128)     ' purple
            xlCht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)      ' green
            Dim si As Long
            For si = 1 To 2
                xlCht.SeriesCollection(si).HasDataLabels = True
                xlCht.SeriesCollection(si).DataLabels.NumberFormat = "#,##0"
                xlCht.SeriesCollection(si).DataLabels.Font.Size = 9
                xlCht.SeriesCollection(si).DataLabels.Orientation = 90
            Next si
            xlCht.Axes(2).TickLabels.NumberFormat = "#,##0"
            xlCht.Export imgInsured
        End If

570     Application.DisplayAlerts = False
580     tmpWs.Delete
590     Application.DisplayAlerts = True

610     Exit Sub
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
620     On Error Resume Next
        Application.DisplayAlerts = False
        If Not tmpWs Is Nothing Then tmpWs.Delete
        Application.DisplayAlerts = True
630     Err.Raise Err.Number, "ExportCompCharts(" & sheetName & "):" & Erl, Err.Description
End Sub

' ============================================================================
' HELPER: Build Title Slide (landscape)
' ============================================================================
Private Sub BuildTitleSlide(ByVal ppSlide As Object, ByVal yearVal As String, ByVal refYear As String, ByVal periodDesc As String, ByVal slideW As Single, ByVal slideH As Single, Optional ByVal paramsSubtitle As String = "")

10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False

        Dim shp As Object

        ' Yellow/gold background
20      Set shp = ppSlide.Shapes.AddShape(1, 0, 0, slideW, slideH)
30      shp.Fill.ForeColor.RGB = RGB(240, 190, 50)
40      shp.Line.Visible = False

        ' White center rectangle
        Dim wL As Single
        Dim wT As Single
        Dim wW As Single
        Dim wH As Single
50      wL = 50
52      wT = 40
54      wW = slideW - 100
56      wH = slideH - 80
60      Set shp = ppSlide.Shapes.AddShape(1, wL, wT, wW, wH)
70      shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
80      shp.Line.Visible = False

        ' Title: "matzget hanhala"
90      Set shp = ppSlide.Shapes.AddTextbox(1, wL + 30, wT + 50, wW - 60, 80)
100     shp.TextFrame.TextRange.Text = ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1492) & ChrW(1504) & ChrW(1492) & ChrW(1500) & ChrW(1492)
110     shp.TextFrame.TextRange.Font.Size = 40
120     shp.TextFrame.TextRange.Font.Bold = True
130     shp.TextFrame.TextRange.Font.Color.RGB = RGB(50, 50, 50)
140     shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
150     shp.TextFrame.WordWrap = True

        ' Period
160     Set shp = ppSlide.Shapes.AddTextbox(1, wL + 30, wT + 140, wW - 60, 60)
170     shp.TextFrame.TextRange.Text = refYear & " " & ChrW(1502) & ChrW(1493) & ChrW(1500) & " " & yearVal
180     shp.TextFrame.TextRange.Font.Size = 32
190     shp.TextFrame.TextRange.Font.Color.RGB = RGB(80, 80, 80)
200     shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
210     shp.TextFrame.WordWrap = True

        ' Parameters subtitle
        If paramsSubtitle <> "" Then
212         Set shp = ppSlide.Shapes.AddTextbox(1, wL + 30, wT + 210, wW - 60, 50)
214         shp.TextFrame.TextRange.Text = paramsSubtitle
216         shp.TextFrame.TextRange.Font.Size = 20
217         shp.TextFrame.TextRange.Font.Bold = False
218         shp.TextFrame.TextRange.Font.Color.RGB = RGB(120, 120, 120)
219         shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
220         shp.TextFrame.WordWrap = True
        End If

        ' Company name
222     Set shp = ppSlide.Shapes.AddTextbox(1, wL + 30, wT + wH - 100, wW - 60, 70)
224     shp.TextFrame.TextRange.Text = ChrW(1500) & ChrW(1489) & ChrW(1489) & " " & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495)
226     shp.TextFrame.TextRange.Font.Size = 34
228     shp.TextFrame.TextRange.Font.Bold = True
230     shp.TextFrame.TextRange.Font.Color.RGB = RGB(0, 130, 60)
232     shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
234     shp.TextFrame.WordWrap = True

        ' Bottom gold line
290     Set shp = ppSlide.Shapes.AddShape(1, 50, slideH - 35, slideW - 100, 5)
300     shp.Fill.ForeColor.RGB = RGB(200, 160, 30)
310     shp.Line.Visible = False

320     Exit Sub
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
330     Err.Raise Err.Number, "BuildTitleSlide:" & Erl, Err.Description
End Sub

' ============================================================================
' HELPER: Build Total Summary Slide from pre-exported image
' ============================================================================
Private Sub BuildTotalSlideFromImage(ByVal ppSlide As Object, ByVal imgPath As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single, Optional ByVal paramsSubtitle As String = "")

10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False

        Dim shp As Object

        ' Title textbox
20      Set shp = ppSlide.Shapes.AddTextbox(1, 20, 10, slideW - 40, 50)
30      shp.TextFrame.TextRange.Text = ChrW(1505) & ChrW(1492) & Chr(34) & ChrW(1499) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & ChrW(1493) & ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " - " & refYear & " " & ChrW(1502) & ChrW(1493) & ChrW(1500) & " " & yearVal
40      shp.TextFrame.TextRange.Font.Size = 24
50      shp.TextFrame.TextRange.Font.Bold = True
60      shp.TextFrame.TextRange.Font.Color.RGB = RGB(50, 50, 50)
70      shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
80      shp.TextFrame.WordWrap = True

        ' Subtitle (parameters)
        If paramsSubtitle <> "" Then
82          Set shp = ppSlide.Shapes.AddTextbox(1, 40, 52, slideW - 80, 22)
84          shp.TextFrame.TextRange.Text = paramsSubtitle
86          shp.TextFrame.TextRange.Font.Size = 12
87          shp.TextFrame.TextRange.Font.Bold = False
88          shp.TextFrame.TextRange.Font.Color.RGB = RGB(120, 120, 120)
89          shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        End If

        ' Insert chart image (landscape: wider)
90      ppSlide.Shapes.AddPicture imgPath, 0, 1, 80, 78, slideW - 160, 432

100     Exit Sub
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
110     Err.Raise Err.Number, "BuildTotalSlideFromImage:" & Erl, Err.Description
End Sub

' ============================================================================
' HELPER: Build a single chart slide (one chart image + title)
' ============================================================================
Private Sub BuildChartSlide(ByVal ppSlide As Object, ByVal imgPath As String, ByVal chartTitle As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single, Optional ByVal paramsSubtitle As String = "")

10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False

        Dim shp As Object

        ' Title
20      Set shp = ppSlide.Shapes.AddTextbox(1, 20, 8, slideW - 40, 36)
30      shp.TextFrame.TextRange.Text = chartTitle & " - " & refYear & " " & ChrW(1502) & ChrW(1493) & ChrW(1500) & " " & yearVal
40      shp.TextFrame.TextRange.Font.Size = 20
50      shp.TextFrame.TextRange.Font.Bold = True
60      shp.TextFrame.TextRange.Font.Color.RGB = RGB(50, 50, 50)
70      shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
80      shp.TextFrame.WordWrap = True

        ' Subtitle (parameters)
        If paramsSubtitle <> "" Then
82          Set shp = ppSlide.Shapes.AddTextbox(1, 40, 42, slideW - 80, 22)
84          shp.TextFrame.TextRange.Text = paramsSubtitle
86          shp.TextFrame.TextRange.Font.Size = 12
87          shp.TextFrame.TextRange.Font.Bold = False
88          shp.TextFrame.TextRange.Font.Color.RGB = RGB(120, 120, 120)
89          shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        End If

        ' Insert chart image
90      ppSlide.Shapes.AddPicture imgPath, 0, 1, 60, 68, slideW - 120, 448

100     Exit Sub
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
110     Err.Raise Err.Number, "BuildChartSlide:" & Erl, Err.Description
End Sub

' ============================================================================
' HELPER: Build a data table slide
' 13 cols: name | premRef | premCur | prem% | docsRef | docsCur | docs% |
'          insuredRef | insuredCur | ins% | commRef | commCur | comm%
' ============================================================================
Private Sub BuildTableSlide(ByVal ppSlide As Object, ByVal sheetName As String, ByVal slideTitle As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single, ByVal slideH As Single, Optional ByVal paramsSubtitle As String = "")

10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False

        Dim ws As Worksheet
        Dim lastRow As Long
        Dim r As Long
        Dim shp As Object
20      Set ws = ThisWorkbook.Worksheets(sheetName)
30      lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        Dim nItems As Long
40      nItems = lastRow - 4
50      If nItems < 1 Then Exit Sub

        ' Read data from sheet
        Dim arrNames() As String
        Dim arrPremR() As Double
        Dim arrPremC() As Double
        Dim arrDocR() As Long
        Dim arrDocC() As Long
        Dim arrInsR() As Long
        Dim arrInsC() As Long
        Dim arrCommR() As Double
        Dim arrCommC() As Double

60      ReDim arrNames(1 To nItems)
70      ReDim arrPremR(1 To nItems)
72      ReDim arrPremC(1 To nItems)
80      ReDim arrDocR(1 To nItems)
82      ReDim arrDocC(1 To nItems)
90      ReDim arrInsR(1 To nItems)
92      ReDim arrInsC(1 To nItems)
100     ReDim arrCommR(1 To nItems)
102     ReDim arrCommC(1 To nItems)

        Dim idx As Long
110     idx = 0
120     For r = 4 To lastRow - 1
130         idx = idx + 1
140         If idx > nItems Then Exit For
150         arrNames(idx) = ShortenCompanyName(Trim$(CStr(ws.Cells(r, 1).Value2)))
160         arrPremR(idx) = CDbl(ws.Cells(r, 2).Value2)
170         arrPremC(idx) = CDbl(ws.Cells(r, 3).Value2)
180         arrDocR(idx) = CLng(ws.Cells(r, 5).Value2)
190         arrDocC(idx) = CLng(ws.Cells(r, 6).Value2)
200         arrInsR(idx) = CLng(ws.Cells(r, 8).Value2)
210         arrInsC(idx) = CLng(ws.Cells(r, 9).Value2)
220         arrCommR(idx) = CDbl(ws.Cells(r, 14).Value2)
230         arrCommC(idx) = CDbl(ws.Cells(r, 15).Value2)
240     Next r
250     nItems = idx

        ' Title
260     Set shp = ppSlide.Shapes.AddTextbox(1, 20, 2, slideW - 40, 28)
270     shp.TextFrame.TextRange.Text = slideTitle & " - " & ChrW(1496) & ChrW(1489) & ChrW(1500) & ChrW(1514) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
280     shp.TextFrame.TextRange.Font.Size = 16
290     shp.TextFrame.TextRange.Font.Bold = True
300     shp.TextFrame.TextRange.Font.Color.RGB = RGB(50, 50, 50)
310     shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
320     shp.TextFrame.WordWrap = True

        ' Subtitle (parameters)
        If paramsSubtitle <> "" Then
322         Set shp = ppSlide.Shapes.AddTextbox(1, 40, 28, slideW - 80, 18)
324         shp.TextFrame.TextRange.Text = paramsSubtitle
326         shp.TextFrame.TextRange.Font.Size = 10
327         shp.TextFrame.TextRange.Font.Bold = False
328         shp.TextFrame.TextRange.Font.Color.RGB = RGB(120, 120, 120)
329         shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        End If

        ' Table: nItems + 3 rows (2 headers + data + total), 13 cols
        Dim tblRows As Long
        Dim tblCols As Long
        Dim tblTop As Single
        Dim tblLeft As Single
        Dim tblWidth As Single
        Dim tblHeight As Single
        Dim rowH As Single
        Dim ppTbl As Object
        Dim tbl As Object
        Dim c As Long
        Dim tblR As Long
        Dim blueClr As Long
        Dim lightBlue As Long
        Dim pctLabel As String

330     tblRows = nItems + 3
340     tblCols = 13
350     tblTop = 45
360     tblLeft = 10
370     tblWidth = slideW - 20
380     rowH = 18
390     If tblRows > 10 Then rowH = 16
392     If tblRows > 15 Then rowH = 14
394     If tblRows > 20 Then rowH = 12
400     tblHeight = tblRows * rowH

410     Set ppTbl = ppSlide.Shapes.AddTable(tblRows, tblCols, tblLeft, tblTop, tblWidth, tblHeight)
420     Set tbl = ppTbl.Table

        ' Column widths (13 cols)
430     tbl.Columns(1).Width = tblWidth * 0.08
440     tbl.Columns(2).Width = tblWidth * 0.105
450     tbl.Columns(3).Width = tblWidth * 0.105
460     tbl.Columns(4).Width = tblWidth * 0.06
470     tbl.Columns(5).Width = tblWidth * 0.07
480     tbl.Columns(6).Width = tblWidth * 0.07
490     tbl.Columns(7).Width = tblWidth * 0.06
500     tbl.Columns(8).Width = tblWidth * 0.07
510     tbl.Columns(9).Width = tblWidth * 0.07
520     tbl.Columns(10).Width = tblWidth * 0.06
530     tbl.Columns(11).Width = tblWidth * 0.09
540     tbl.Columns(12).Width = tblWidth * 0.09
550     tbl.Columns(13).Width = tblWidth * 0.07

560     blueClr = RGB(0, 100, 170)
570     lightBlue = RGB(180, 210, 240)
580     pctLabel = ChrW(1513) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1497)

        ' Header row 1 - category names
590     tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text = ""
        ' premiot
600     tbl.Cell(1, 2).Shape.TextFrame.TextRange.Text = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514)
        ' mismachim
610     tbl.Cell(1, 5).Shape.TextFrame.TextRange.Text = ChrW(1502) & ChrW(1505) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1501)
        ' mevutachim
620     tbl.Cell(1, 8).Shape.TextFrame.TextRange.Text = ChrW(1502) & ChrW(1489) & ChrW(1493) & ChrW(1496) & ChrW(1495) & ChrW(1497) & ChrW(1501)
        ' amulot
630     tbl.Cell(1, 11).Shape.TextFrame.TextRange.Text = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514)

        ' Header row 2 - year sub-headers
640     tbl.Cell(2, 1).Shape.TextFrame.TextRange.Text = ""
650     tbl.Cell(2, 2).Shape.TextFrame.TextRange.Text = refYear
660     tbl.Cell(2, 3).Shape.TextFrame.TextRange.Text = yearVal
670     tbl.Cell(2, 4).Shape.TextFrame.TextRange.Text = pctLabel
680     tbl.Cell(2, 5).Shape.TextFrame.TextRange.Text = refYear
690     tbl.Cell(2, 6).Shape.TextFrame.TextRange.Text = yearVal
700     tbl.Cell(2, 7).Shape.TextFrame.TextRange.Text = pctLabel
710     tbl.Cell(2, 8).Shape.TextFrame.TextRange.Text = refYear
720     tbl.Cell(2, 9).Shape.TextFrame.TextRange.Text = yearVal
730     tbl.Cell(2, 10).Shape.TextFrame.TextRange.Text = pctLabel
740     tbl.Cell(2, 11).Shape.TextFrame.TextRange.Text = refYear
750     tbl.Cell(2, 12).Shape.TextFrame.TextRange.Text = yearVal
760     tbl.Cell(2, 13).Shape.TextFrame.TextRange.Text = pctLabel

        ' Format header rows
770     For c = 1 To tblCols
780         tbl.Cell(1, c).Shape.TextFrame.TextRange.Font.Name = "Arial"
782         tbl.Cell(1, c).Shape.TextFrame.TextRange.Font.Size = 11
790         tbl.Cell(1, c).Shape.TextFrame.TextRange.Font.Bold = True
800         tbl.Cell(1, c).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
810         tbl.Cell(1, c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
820         tbl.Cell(1, c).Shape.Fill.ForeColor.RGB = blueClr
830         tbl.Cell(2, c).Shape.TextFrame.TextRange.Font.Name = "Arial"
832         tbl.Cell(2, c).Shape.TextFrame.TextRange.Font.Size = 11
840         tbl.Cell(2, c).Shape.TextFrame.TextRange.Font.Bold = True
850         tbl.Cell(2, c).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
860         tbl.Cell(2, c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
870         tbl.Cell(2, c).Shape.Fill.ForeColor.RGB = blueClr
880     Next c

        ' Data rows
        Dim chgVal As Double
890     For idx = 1 To nItems
900         tblR = idx + 2
910         tbl.Cell(tblR, 1).Shape.TextFrame.TextRange.Text = arrNames(idx)
            ' Premiums
920         tbl.Cell(tblR, 2).Shape.TextFrame.TextRange.Text = Format$(arrPremR(idx), "#,##0")
930         tbl.Cell(tblR, 3).Shape.TextFrame.TextRange.Text = Format$(arrPremC(idx), "#,##0")
940         If arrPremR(idx) <> 0 Then
950             chgVal = (arrPremC(idx) - arrPremR(idx)) / Abs(arrPremR(idx)) * 100
960             tbl.Cell(tblR, 4).Shape.TextFrame.TextRange.Text = Format$(chgVal, "0.0") & "%"
970         Else
980             tbl.Cell(tblR, 4).Shape.TextFrame.TextRange.Text = "-"
990         End If
            ' Documents
1000        tbl.Cell(tblR, 5).Shape.TextFrame.TextRange.Text = Format$(arrDocR(idx), "#,##0")
1010        tbl.Cell(tblR, 6).Shape.TextFrame.TextRange.Text = Format$(arrDocC(idx), "#,##0")
1020        If arrDocR(idx) <> 0 Then
1030            chgVal = (CDbl(arrDocC(idx)) - CDbl(arrDocR(idx))) / Abs(CDbl(arrDocR(idx))) * 100
1040            tbl.Cell(tblR, 7).Shape.TextFrame.TextRange.Text = Format$(chgVal, "0.0") & "%"
1050        Else
1060            tbl.Cell(tblR, 7).Shape.TextFrame.TextRange.Text = "-"
1070        End If
            ' Insured
1080        tbl.Cell(tblR, 8).Shape.TextFrame.TextRange.Text = Format$(arrInsR(idx), "#,##0")
1090        tbl.Cell(tblR, 9).Shape.TextFrame.TextRange.Text = Format$(arrInsC(idx), "#,##0")
1100        If arrInsR(idx) <> 0 Then
1110            chgVal = (CDbl(arrInsC(idx)) - CDbl(arrInsR(idx))) / Abs(CDbl(arrInsR(idx))) * 100
1120            tbl.Cell(tblR, 10).Shape.TextFrame.TextRange.Text = Format$(chgVal, "0.0") & "%"
1130        Else
1140            tbl.Cell(tblR, 10).Shape.TextFrame.TextRange.Text = "-"
1150        End If
            ' Commissions
1160        tbl.Cell(tblR, 11).Shape.TextFrame.TextRange.Text = Format$(arrCommR(idx), "#,##0")
1170        tbl.Cell(tblR, 12).Shape.TextFrame.TextRange.Text = Format$(arrCommC(idx), "#,##0")
1180        If arrCommR(idx) <> 0 Then
1190            chgVal = (arrCommC(idx) - arrCommR(idx)) / Abs(arrCommR(idx)) * 100
1200            tbl.Cell(tblR, 13).Shape.TextFrame.TextRange.Text = Format$(chgVal, "0.0") & "%"
1210        Else
1220            tbl.Cell(tblR, 13).Shape.TextFrame.TextRange.Text = "-"
1230        End If

            ' Format data cells
1240        For c = 1 To tblCols
1250            tbl.Cell(tblR, c).Shape.TextFrame.TextRange.Font.Name = "Arial"
1252            tbl.Cell(tblR, c).Shape.TextFrame.TextRange.Font.Size = 10
1260            tbl.Cell(tblR, c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
1270            If idx Mod 2 = 0 Then
1280                tbl.Cell(tblR, c).Shape.Fill.ForeColor.RGB = lightBlue
1290            Else
1300                tbl.Cell(tblR, c).Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
1310            End If
1320        Next c
1330    Next idx

        ' Total row
        Dim totR As Long
        Dim totPR As Double
        Dim totPC As Double
        Dim totCR As Double
        Dim totCC As Double
        Dim totDR As Long
        Dim totDC As Long
        Dim totIR As Long
        Dim totIC As Long
1340    totR = nItems + 3
1350    totPR = CDbl(ws.Cells(lastRow, 2).Value2)
1360    totPC = CDbl(ws.Cells(lastRow, 3).Value2)
1370    totDR = CLng(ws.Cells(lastRow, 5).Value2)
1380    totDC = CLng(ws.Cells(lastRow, 6).Value2)
1390    totIR = CLng(ws.Cells(lastRow, 8).Value2)
1400    totIC = CLng(ws.Cells(lastRow, 9).Value2)
1410    totCR = CDbl(ws.Cells(lastRow, 14).Value2)
1420    totCC = CDbl(ws.Cells(lastRow, 15).Value2)

        ' "sach hakol"
1430    tbl.Cell(totR, 1).Shape.TextFrame.TextRange.Text = ChrW(1505) & ChrW(1499) & ChrW(1493) & ChrW(1501) & " " & ChrW(1492) & ChrW(1499) & ChrW(1500)
1440    tbl.Cell(totR, 2).Shape.TextFrame.TextRange.Text = Format$(totPR, "#,##0")
1450    tbl.Cell(totR, 3).Shape.TextFrame.TextRange.Text = Format$(totPC, "#,##0")
1460    If totPR <> 0 Then
1470        tbl.Cell(totR, 4).Shape.TextFrame.TextRange.Text = Format$((totPC - totPR) / Abs(totPR) * 100, "0.0") & "%"
1480    Else
1490        tbl.Cell(totR, 4).Shape.TextFrame.TextRange.Text = "-"
1500    End If
1510    tbl.Cell(totR, 5).Shape.TextFrame.TextRange.Text = Format$(totDR, "#,##0")
1520    tbl.Cell(totR, 6).Shape.TextFrame.TextRange.Text = Format$(totDC, "#,##0")
1530    If totDR <> 0 Then
1540        tbl.Cell(totR, 7).Shape.TextFrame.TextRange.Text = Format$((CDbl(totDC) - CDbl(totDR)) / Abs(CDbl(totDR)) * 100, "0.0") & "%"
1550    Else
1560        tbl.Cell(totR, 7).Shape.TextFrame.TextRange.Text = "-"
1570    End If
1580    tbl.Cell(totR, 8).Shape.TextFrame.TextRange.Text = Format$(totIR, "#,##0")
1590    tbl.Cell(totR, 9).Shape.TextFrame.TextRange.Text = Format$(totIC, "#,##0")
1600    If totIR <> 0 Then
1610        tbl.Cell(totR, 10).Shape.TextFrame.TextRange.Text = Format$((CDbl(totIC) - CDbl(totIR)) / Abs(CDbl(totIR)) * 100, "0.0") & "%"
1620    Else
1630        tbl.Cell(totR, 10).Shape.TextFrame.TextRange.Text = "-"
1640    End If
1650    tbl.Cell(totR, 11).Shape.TextFrame.TextRange.Text = Format$(totCR, "#,##0")
1660    tbl.Cell(totR, 12).Shape.TextFrame.TextRange.Text = Format$(totCC, "#,##0")
1670    If totCR <> 0 Then
1680        tbl.Cell(totR, 13).Shape.TextFrame.TextRange.Text = Format$((totCC - totCR) / Abs(totCR) * 100, "0.0") & "%"
1690    Else
1700        tbl.Cell(totR, 13).Shape.TextFrame.TextRange.Text = "-"
1710    End If

        ' Format total row - dark background matching header for bold emphasis
1720    For c = 1 To tblCols
1730        tbl.Cell(totR, c).Shape.TextFrame.TextRange.Font.Name = "Arial"
1732        tbl.Cell(totR, c).Shape.TextFrame.TextRange.Font.Size = 11
1740        tbl.Cell(totR, c).Shape.TextFrame.TextRange.Font.Bold = True
1750        tbl.Cell(totR, c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
1755        tbl.Cell(totR, c).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
1760        tbl.Cell(totR, c).Shape.Fill.ForeColor.RGB = RGB(0, 60, 120)
1770    Next c

1780    Exit Sub
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
1790    Err.Raise Err.Number, "BuildTableSlide(" & sheetName & "):" & Erl, Err.Description
End Sub

' ============================================================================
' HELPER: Add page number to slide
' ============================================================================
Private Sub AddPageNumber(ByVal ppSlide As Object, ByVal pageNum As Long, ByVal totalPages As Long, ByVal slideW As Single, ByVal slideH As Single)

10      On Error Resume Next

        Dim shp As Object
20      Set shp = ppSlide.Shapes.AddTextbox(1, slideW - 120, 2, 110, 22)
30      shp.TextFrame.TextRange.Text = pageNum & " / " & totalPages
40      shp.TextFrame.TextRange.Font.Name = "Arial"
50      shp.TextFrame.TextRange.Font.Size = 10
60      shp.TextFrame.TextRange.Font.Color.RGB = RGB(120, 120, 120)
70      shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
80      shp.TextFrame.WordWrap = False
90      shp.TextFrame.MarginTop = 0
100     shp.TextFrame.MarginBottom = 0
110     shp.ZOrder 0

End Sub

' ============================================================================
' MACRO: SendForReview - collects rows marked "???? ??????" and sends via Outlook
' Called from the "?????? ?????" button on the review sheet
' ============================================================================
Public Sub SendForReview()

10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False

        ' Remove any leftover sheet protection
        Dim wsUp4 As Worksheet
        For Each wsUp4 In ThisWorkbook.Worksheets
            On Error Resume Next
            wsUp4.Unprotect "Z961814r"
            On Error GoTo ERR_HANDLER
        Next wsUp4

20      Dim wsRev As Worksheet
30      Set wsRev = ActiveSheet

        ' Verify we are on a review sheet (name starts with REVIEW_SHEET_NAME)
40      If InStr(1, wsRev.Name, REVIEW_SHEET_NAME(), vbTextCompare) = 0 Then
50          MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1492) & ChrW(1512) & ChrW(1497) & ChrW(1509) & " " & ChrW(1502) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500), vbExclamation
60          Exit Sub
70      End If

        ' Find action column (header in ROW 2 contains "peula" = ?????)
80      Dim lastCol As Long
90      lastCol = wsRev.Cells(2, wsRev.Columns.Count).End(xlToLeft).Column
100     Dim actionCol As Long
110     actionCol = 0
120     Dim c As Long
130     For c = 1 To lastCol
140         If InStr(1, CStr(wsRev.Cells(2, c).Value2), ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) & ChrW(1492), vbTextCompare) > 0 Then
150             actionCol = c
160             Exit For
170         End If
180     Next c
190     If actionCol = 0 Then
200         MsgBoxU ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1492) & " " & ChrW(1506) & ChrW(1502) & ChrW(1493) & ChrW(1491) & ChrW(1514) & " " & ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) & ChrW(1492), vbExclamation
210         Exit Sub
220     End If

        ' Count rows with "ha'aver livdika"
230     Dim lastRow As Long
240     lastRow = wsRev.Cells(wsRev.Rows.Count, 1).End(xlUp).Row
250     Dim sendCount As Long
260     sendCount = 0
270     Dim r As Long
        ' "ha'aver livdika" = ???? ??????
        Dim actionMatch As String
280     actionMatch = H_REVIEW()
290     For r = 3 To lastRow
300         If InStr(1, CStr(wsRev.Cells(r, actionCol).Value2), actionMatch, vbTextCompare) > 0 Then
310             sendCount = sendCount + 1
320         End If
330     Next r

340     If sendCount = 0 Then
            ' Count handled rows (rows with non-empty action column)
            Dim handledCount As Long
            handledCount = 0
            For r = 3 To lastRow
                If Trim$(CStr(wsRev.Cells(r, actionCol).Value2)) <> "" Then handledCount = handledCount + 1
            Next r
            ' "ein shurot le'ha'avara. X shurot tuplu" = no rows to transfer. X rows handled
350         MsgBoxU ChrW(1488) & ChrW(1497) & ChrW(1503) & " " & ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1492) & ChrW(1506) & ChrW(1489) & ChrW(1512) & ChrW(1492) & "." & vbCrLf & handledCount & " " & ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1496) & ChrW(1493) & ChrW(1508) & ChrW(1500) & ChrW(1493), vbInformation
360         Exit Sub
370     End If

        ' Get email address from parameters
380     Dim wsMgmt As Worksheet
390     Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
400     Dim emailAddr As String
410     emailAddr = GetStringParameter(wsMgmt, PARAM_ERROR_EMAIL)
420     If emailAddr = "" Then
            ' "lo hugdra ktovet email" = email address not defined
430         MsgBoxU ChrW(1500) & ChrW(1488) & " " & ChrW(1492) & ChrW(1493) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1492) & " " & ChrW(1499) & ChrW(1514) & ChrW(1493) & ChrW(1489) & ChrW(1514) & " " & ChrW(1488) & ChrW(1497) & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514), vbExclamation
440         Exit Sub
450     End If

        ' Create temp workbook with matching rows
460     Dim wbTemp As Workbook
470     Set wbTemp = Workbooks.Add
480     Dim wsTemp As Worksheet
490     Set wsTemp = wbTemp.Worksheets(1)

        ' Copy header row (from row 2 of review sheet)
500     Dim hdrCol As Long
510     For hdrCol = 1 To lastCol
520         wsTemp.Cells(1, hdrCol).Value = wsRev.Cells(2, hdrCol).Value
530     Next hdrCol
540     wsTemp.Rows(1).Font.Bold = True

        ' Copy matching rows (data starts at row 3)
550     Dim outRow As Long
560     outRow = 2
570     For r = 3 To lastRow
580         If InStr(1, CStr(wsRev.Cells(r, actionCol).Value2), actionMatch, vbTextCompare) > 0 Then
590             For hdrCol = 1 To lastCol
600                 wsTemp.Cells(outRow, hdrCol).Value = wsRev.Cells(r, hdrCol).Value
610             Next hdrCol
620             outRow = outRow + 1
630         End If
640     Next r
650     wsTemp.Columns.AutoFit

        ' Save temp file
660     Dim tempPath As String
670     tempPath = REPORTS_FOLDER() & "\" & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501) & "_" & wsRev.Name & ".xlsx"
680     Application.DisplayAlerts = False
690     wbTemp.SaveAs tempPath, xlOpenXMLWorkbook
700     wbTemp.Close SaveChanges:=False
710     Application.DisplayAlerts = True

        ' Build email body
        ' "hi lahav" = ?? ???
720     Dim bodyLine1 As String
730     bodyLine1 = ChrW(1492) & ChrW(1497) & " " & ChrW(1500) & ChrW(1492) & ChrW(1489)
        ' "likrat hachanat doch avurchem nimtze'u hachrigim haram" = ????? ???? ??? ?????? ????? ??????? ??"?
740     Dim bodyLine2 As String
750     bodyLine2 = ChrW(1500) & ChrW(1511) & ChrW(1512) & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1499) & ChrW(1504) & ChrW(1514) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & " " & ChrW(1506) & ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1499) & ChrW(1501) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & ChrW(1492) & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501) & " " & ChrW(1492) & ChrW(1512) & ChrW(34) & ChrW(1502)
        ' "al mnat lehafik et hadoch ani mevakeshet tguvatcha al mnat she'etaken beheta'am"
760     Dim bodyLine3 As String
770     bodyLine3 = ChrW(1506) & ChrW(1500) & " " & ChrW(1502) & ChrW(1504) & ChrW(1514) & " " & ChrW(1500) & ChrW(1492) & ChrW(1508) & ChrW(1497) & ChrW(1511) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1491) & ChrW(1493) & ChrW(1495) & " " & ChrW(1488) & ChrW(1504) & ChrW(1497) & " " & ChrW(1502) & ChrW(1489) & ChrW(1511) & ChrW(1513) & ChrW(1514) & " " & ChrW(1514) & ChrW(1490) & ChrW(1493) & ChrW(1489) & ChrW(1514) & ChrW(1498) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1502) & ChrW(1504) & ChrW(1514) & " " & ChrW(1513) & ChrW(1488) & H_FIX() & " " & ChrW(1489) & ChrW(1492) & ChrW(1514) & ChrW(1488) & ChrW(1501)
        ' toda = ????
780     Dim bodyLine4 As String
790     bodyLine4 = ChrW(1514) & ChrW(1493) & ChrW(1491) & ChrW(1492)
        ' orit = ?????
800     Dim bodyLine5 As String
810     bodyLine5 = ChrW(1488) & ChrW(1493) & ChrW(1512) & ChrW(1497) & ChrW(1514)

820     Dim emailBody As String
830     emailBody = bodyLine1 & vbCrLf & vbCrLf & bodyLine2 & vbCrLf & bodyLine3 & vbCrLf & vbCrLf & bodyLine4 & vbCrLf & bodyLine5

        ' Email subject: "charigim shenimtze'u letipulcha" = ?????? ?????? ???????
840     Dim emailSubject As String
850     emailSubject = ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501) & " " & ChrW(1513) & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1498)

        ' Create Outlook email (late binding)
860     Dim olApp As Object
870     Dim olMail As Object
880     Set olApp = CreateObject("Outlook.Application")
890     Set olMail = olApp.CreateItem(0)
900     olMail.To = emailAddr
910     olMail.Subject = emailSubject
920     olMail.Body = emailBody
930     olMail.Attachments.Add tempPath
940     olMail.Display

        ' Success message: "email huchan be'hatzlacha" = ???? ???? ??????
        ' Then ask: "ha'im shalacht et hamail?" = ??? ???? ?? ??????
950     Dim askSent As Long
        askSent = MsgBoxU( _
            ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & " " & ChrW(1492) & ChrW(1493) & ChrW(1499) & ChrW(1503) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & " " & ChrW(1506) & ChrW(1501) & " " & sendCount & " " & ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & "." & vbCrLf & vbCrLf & _
            ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & ChrW(1514) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & "?" & vbCrLf & _
            ChrW(1499) & ChrW(1503) & " = " & ChrW(1492) & ChrW(1502) & ChrW(1513) & ChrW(1498) & " | " & ChrW(1500) & ChrW(1488) & " = " & ChrW(1488) & ChrW(1504) & ChrW(1488) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " " & ChrW(1493) & ChrW(1488) & ChrW(1513) & ChrW(1512), _
            vbYesNo + vbQuestion)
        If askSent = vbYes Then
            ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Activate
        Else
            ' "ana shlach ve'asher" = ??? ??? ????
            MsgBoxU ChrW(1488) & ChrW(1504) & ChrW(1488) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & " " & ChrW(1493) & ChrW(1488) & ChrW(1513) & ChrW(1512), vbExclamation
        End If

960     Exit Sub

ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
970     MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1495) & ChrW(1514) & " " & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & ":" & vbCrLf & Err.Description, vbCritical

End Sub

' ============================================================================
' HIDE/SHOW SHEETS - Show/Hide toggle
' ============================================================================
Public Sub HideWorkSheets()
    ' Hides ALL sheets except home sheet (daf habait)
    Dim ws As Worksheet
    Dim ctrlName As String
    
    ctrlName = CONTROL_SHEET_NAME()
    
    ' Make sure home sheet is visible before hiding others
    ThisWorkbook.Worksheets(ctrlName).Visible = xlSheetVisible
    
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, ctrlName, vbTextCompare) <> 0 Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    
    ' Activate home sheet
    ThisWorkbook.Worksheets(ctrlName).Activate
    
End Sub

Public Sub ShowHiddenSheets()
    ' Shows all hidden sheets (no password needed since protection removed)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    Next ws
    
    ' Stay on home sheet
    On Error Resume Next
    ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Activate
    On Error GoTo 0
    
    MsgBoxU ChrW(1499) & ChrW(1500) & " " & ChrW(1492) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1502) & ChrW(1493) & ChrW(1510) & ChrW(1490) & ChrW(1497) & ChrW(1501), vbInformation
End Sub

' ============================================================================
' TOGGLE HIDDEN SHEETS - checks if sheets are hidden, then shows or hides
' ============================================================================
Public Sub ToggleHiddenSheets()
    ' Check if any non-home sheet is visible
    Dim ws As Worksheet
    Dim ctrlName As String
    Dim anyVisible As Boolean
    
    ctrlName = CONTROL_SHEET_NAME()
    anyVisible = False
    
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, ctrlName, vbTextCompare) <> 0 Then
            If ws.Visible = xlSheetVisible Then
                anyVisible = True
                Exit For
            End If
        End If
    Next ws
    
    If anyVisible Then
        ' Currently visible - hide them
        HideWorkSheets
    Else
        ' Currently hidden - ask for password before showing
        Dim pwd As String
        pwd = InputBox("Password:", "Show Sheets")
        If pwd <> "Z961814r" Then
            If pwd <> "" Then MsgBoxU ChrW(1505) & ChrW(1497) & ChrW(1505) & ChrW(1502) & ChrW(1492) & " " & ChrW(1513) & ChrW(1490) & ChrW(1493) & ChrW(1497) & ChrW(1492), vbExclamation  ' "????? ?????"
            Exit Sub
        End If
        ShowHiddenSheets
    End If
End Sub

' ============================================================================
' SETUP SETTINGS SHEET - Add cover page with navigation buttons
' ============================================================================
Public Sub SetupSettingsSheet()
    On Error GoTo ERR_HANDLER
    
    Dim wsMgmt As Worksheet
    Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    
    ' Delete existing navigation buttons
    Dim s As Shape
    On Error Resume Next
    For Each s In wsMgmt.Shapes
        If Left$(s.Name, 3) = "nav" Then s.Delete
    Next s
    On Error GoTo ERR_HANDLER
    
    ' ---- Title bar ----
    Dim shp As Shape
    Dim btnTop As Single
    
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, 200, 5, 350, 40)
    shp.Name = "navTitle"
    shp.Fill.ForeColor.RGB = RGB(0, 70, 130)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514) & " " & ChrW(1500) & ChrW(1489) & ChrW(1489)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 18
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    
    ' ---- Navigation buttons ----
    btnTop = 55
    
    ' Button 1: Milon Anafim (Branch Dictionary) -> A1
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, 200, btnTop, 170, 32)
    shp.Name = "navBranch"
    shp.Fill.ForeColor.RGB = RGB(0, 120, 60)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = ChrW(1502) & ChrW(1497) & ChrW(1500) & ChrW(1493) & ChrW(1503) & " " & ChrW(1506) & ChrW(1504) & ChrW(1508) & ChrW(1497) & ChrW(1501)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavToBranch"
    
    btnTop = btnTop + 40
    
    ' Button 2: Index Makor (Source Index) -> E1
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, 200, btnTop, 170, 32)
    shp.Name = "navIndex"
    shp.Fill.ForeColor.RGB = RGB(0, 100, 170)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = ChrW(1488) & ChrW(1497) & ChrW(1504) & ChrW(1491) & ChrW(1511) & ChrW(1505) & " " & ChrW(1502) & ChrW(1511) & ChrW(1493) & ChrW(1512)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavToIndex"
    
    btnTop = btnTop + 40
    
    ' Button 3: Parametrim (Parameters) -> J1
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, 200, btnTop, 170, 32)
    shp.Name = "navParams"
    shp.Fill.ForeColor.RGB = RGB(160, 80, 0)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512) & ChrW(1497) & ChrW(1501)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavToParams"
    
    btnTop = btnTop + 40
    
    ' Button 4: Reshimat Shgiot (Error List) -> N1
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, 200, btnTop, 170, 32)
    shp.Name = "navErrors"
    shp.Fill.ForeColor.RGB = RGB(180, 30, 30)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1514) & " " & ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavToErrors"
    
    MsgBoxU ChrW(1491) & ChrW(1507) & " " & ChrW(1513) & ChrW(1506) & ChrW(1512) & " " & ChrW(1492) & ChrW(1493) & ChrW(1490) & ChrW(1491) & ChrW(1512) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492), vbInformation
    Exit Sub
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
    MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1514) & " " & ChrW(1491) & ChrW(1507) & " " & ChrW(1513) & ChrW(1506) & ChrW(1512) & ":" & vbCrLf & Err.Description, vbCritical
End Sub

' ---- Navigation macros for settings sheet buttons ----
Public Sub NavToBranch()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    ws.Activate
    Application.Goto ws.Range("A1")
    Application.GoTo ws.Range("A1"), True
End Sub

Public Sub NavToIndex()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    ws.Visible = xlSheetVisible
    ws.Activate
    Application.Goto ws.Range("A1")
End Sub

Public Sub NavToParams()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    ws.Activate
    Application.Goto ws.Range("J1")
    Application.GoTo ws.Range("J1"), True
End Sub

Public Sub NavToErrors()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    ws.Activate
    Application.Goto ws.Range("N1")
    Application.GoTo ws.Range("N1"), True
End Sub

' ============================================================================
' PUBLIC: Search client name - opens search sheet with cell-based input
' Called from the "Search" button on the home sheet
' ============================================================================
Public Sub SearchClientName()
    On Error GoTo ERR_HANDLER
    
    ' Create/clear temp search sheet
    Dim SEARCH_SHEET_NAME As String
    SEARCH_SHEET_NAME = ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513)  ' "?????"
    
    Dim wsSearch As Worksheet
    On Error Resume Next
    Set wsSearch = ThisWorkbook.Worksheets(SEARCH_SHEET_NAME)
    On Error GoTo ERR_HANDLER
    
    If wsSearch Is Nothing Then
        Set wsSearch = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsSearch.Name = SEARCH_SHEET_NAME
    Else
        wsSearch.Cells.Clear
        Dim shp As Shape
        For Each shp In wsSearch.Shapes
            shp.Delete
        Next shp
    End If
    
    ' Set RTL
    wsSearch.DisplayRightToLeft = True
    
    ' Row 1: Instructions
    wsSearch.Cells(1, 1).Value = ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & " " & ChrW(1489) & ChrW(1514) & ChrW(1488) & " " & ChrW(1492) & ChrW(1510) & ChrW(1492) & ChrW(1493) & ChrW(1489) & " " & ChrW(1493) & ChrW(1500) & ChrW(1495) & ChrW(1509) & " " & ChrW(1495) & ChrW(1508) & ChrW(1513)  ' "???? ?? ???? ??? ????? ???? ???"
    wsSearch.Cells(1, 1).Font.Bold = True
    wsSearch.Cells(1, 1).Font.Size = 13
    
    ' Row 2: Yellow search cell
    wsSearch.Cells(2, 1).Interior.Color = RGB(255, 255, 200)
    wsSearch.Cells(2, 1).Font.Size = 14
    wsSearch.Cells(2, 1).Borders.LineStyle = xlContinuous
    wsSearch.Columns(1).ColumnWidth = 40
    
    ' Add "Search" button
    Dim btnLeft As Double, btnTop As Double
    btnLeft = wsSearch.Cells(2, 2).Left + 10
    btnTop = wsSearch.Cells(2, 1).Top
    Dim shpSearch As Shape
    Set shpSearch = wsSearch.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, 80, 25)
    With shpSearch
        .Name = "btnDoSearch"
        .TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513)  ' "???"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Fill.ForeColor.RGB = RGB(0, 176, 80)
        .Line.Visible = msoFalse
        .OnAction = "DoClientSearch"
    End With
    
    ' Add "Cancel" button
    Dim shpCancel As Shape
    Set shpCancel = wsSearch.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft + 90, btnTop, 80, 25)
    With shpCancel
        .Name = "btnCancelSearch"
        .TextFrame2.TextRange.Text = ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1500)  ' "?????"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Fill.ForeColor.RGB = RGB(192, 0, 0)
        .Line.Visible = msoFalse
        .OnAction = "CancelClientSearch"
    End With
    
    ' Activate the search sheet and put cursor in search cell
    wsSearch.Activate
    Application.Goto wsSearch.Cells(2, 1)
    Exit Sub
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
    MsgBoxU "Error: " & Err.Description, vbCritical
End Sub

' ============================================================================
' PUBLIC: Perform the actual search - reads text from A2 on search sheet
' Called from the "Search" button on the search sheet
' ============================================================================
Public Sub DoClientSearch()
    On Error GoTo ERR_HANDLER
    
    Dim SEARCH_SHEET_NAME As String
    SEARCH_SHEET_NAME = ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513)  ' "?????"
    
    Dim wsSearch As Worksheet
    Set wsSearch = ThisWorkbook.Worksheets(SEARCH_SHEET_NAME)
    
    ' Read search text from cell A2
    Dim searchText As String
    searchText = Trim$(CStr(wsSearch.Cells(2, 1).Value))
    If searchText = "" Then
        MsgBoxU ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " " & ChrW(1496) & ChrW(1511) & ChrW(1505) & ChrW(1496) & " " & ChrW(1500) & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513), vbExclamation  ' "???? ???? ??????"
        Application.Goto wsSearch.Cells(2, 1)
        Exit Sub
    End If
    
    ' Collect MATCHING unique client names from base sheets
    Dim ws As Worksheet
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Dim foundBase As Boolean
    foundBase = False
    Dim lastRow As Long, r As Long
    Dim cName As String
    
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, H_BASE(), vbTextCompare) > 0 Then
            foundBase = True
            lastRow = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
            For r = 2 To lastRow
                cName = Trim$(CStr(ws.Cells(r, 6).Value2))
                If cName <> "" Then
                    If InStr(1, cName, searchText, vbTextCompare) > 0 Then
                        If Not dict.Exists(cName) Then dict.Add cName, 1
                    End If
                End If
            Next r
        End If
    Next ws
    
    If Not foundBase Then
        MsgBoxU ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & H_BASE(), vbExclamation
        Exit Sub
    End If
    
    If dict.Count = 0 Then
        MsgBoxU ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1502) & ChrW(1514) & ChrW(1488) & ChrW(1497) & ChrW(1502) & ChrW(1497) & ChrW(1501), vbInformation  ' "?? ????? ?????? ???????"
        Application.Goto wsSearch.Cells(2, 1)
        Exit Sub
    End If
    
    ' If only 1 result, select it directly
    If dict.Count = 1 Then
        Dim wsMain As Worksheet
        Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        wsMain.Range("rngClientName").Value = dict.keys()(0)
        ' Clean up search sheet
        Application.DisplayAlerts = False
        wsSearch.Delete
        Application.DisplayAlerts = True
        wsMain.Activate
        MsgBoxU ChrW(1504) & ChrW(1489) & ChrW(1495) & ChrW(1512) & ": " & dict.keys()(0), vbInformation  ' "????: [name]"
        Exit Sub
    End If
    
    ' Sort matching names alphabetically
    Dim arrAll As Variant
    arrAll = dict.keys
    Dim i As Long, j As Long, tmp As String
    For i = 0 To UBound(arrAll) - 1
        For j = i + 1 To UBound(arrAll)
            If arrAll(i) > arrAll(j) Then
                tmp = arrAll(i): arrAll(i) = arrAll(j): arrAll(j) = tmp
            End If
        Next j
    Next i
    
    ' Clear results area (row 4 onwards) but keep search cell and buttons
    Dim clearLastRow As Long
    clearLastRow = wsSearch.Cells(wsSearch.Rows.Count, 1).End(xlUp).Row
    If clearLastRow >= 4 Then wsSearch.Range(wsSearch.Cells(4, 1), wsSearch.Cells(clearLastRow, 1)).Clear
    
    ' Write header with result count in row 3
    wsSearch.Cells(3, 1).Value = ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & dict.Count & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " - " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1493) & ChrW(1500) & ChrW(1495) & ChrW(1509) & " " & ChrW(1489) & ChrW(1495) & ChrW(1512)  ' "????? X ?????? - ??? ?? ?? ???? ???"
    wsSearch.Cells(3, 1).Font.Bold = True
    wsSearch.Cells(3, 1).Font.Size = 12
    wsSearch.Cells(3, 1).Font.Color = RGB(0, 112, 192)
    
    ' Write filtered client names starting row 4
    For i = 0 To UBound(arrAll)
        wsSearch.Cells(i + 4, 1).Value = arrAll(i)
        wsSearch.Cells(i + 4, 1).Font.Size = 12
    Next i
    
    ' Add "Select" button in column B row 3 (for selecting after clicking a name)
    Dim btnLeft As Double, btnTop2 As Double
    btnLeft = wsSearch.Cells(3, 2).Left + 10
    btnTop2 = wsSearch.Cells(3, 2).Top
    ' Remove old select button if exists
    On Error Resume Next
    wsSearch.Shapes("btnSelectClient").Delete
    On Error GoTo ERR_HANDLER
    Dim shpSelect As Shape
    Set shpSelect = wsSearch.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop2, 80, 25)
    With shpSelect
        .Name = "btnSelectClient"
        .TextFrame2.TextRange.Text = ChrW(1489) & ChrW(1495) & ChrW(1512)  ' "???"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Fill.ForeColor.RGB = RGB(0, 112, 192)
        .Line.Visible = msoFalse
        .OnAction = "ConfirmClientSelection"
    End With
    
    ' Select first result
    Application.Goto wsSearch.Cells(4, 1)
    Exit Sub
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
    MsgBoxU "Error: " & Err.Description, vbCritical
End Sub

' ============================================================================
' PUBLIC: Confirm client selection - reads active cell from search sheet
' Called from the "Select" button on the search sheet
' ============================================================================
Public Sub ConfirmClientSelection()
    On Error Resume Next
    
    Dim selectedName As String
    selectedName = Trim$(CStr(ActiveCell.Value))
    
    If selectedName = "" Then
        MsgBoxU ChrW(1489) & ChrW(1495) & ChrW(1512) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & " " & ChrW(1502) & ChrW(1492) & ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1492), vbExclamation  ' "??? ???? ???????"
        Exit Sub
    End If
    
    ' Write to G12 on home sheet
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    wsMain.Range("rngClientName").Value = selectedName
    
    ' Delete search sheet and go back to home
    Dim SEARCH_SHEET_NAME As String
    SEARCH_SHEET_NAME = ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513)  ' "?????"
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(SEARCH_SHEET_NAME).Delete
    Application.DisplayAlerts = True
    
    wsMain.Activate
End Sub

' ============================================================================
' PUBLIC: Cancel client search - delete search sheet, go back to home
' ============================================================================
Public Sub CancelClientSearch()
    On Error Resume Next
    
    Dim SEARCH_SHEET_NAME As String
    SEARCH_SHEET_NAME = ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513)  ' "?????"
    
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(SEARCH_SHEET_NAME).Delete
    Application.DisplayAlerts = True
    
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    wsMain.Activate
End Sub

' ============================================================================
' PUBLIC: Clear client filter - resets G12 to default
' ============================================================================
Public Sub ClearClientFilter()
    On Error Resume Next
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    wsMain.Range("rngClientName").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)  ' "???/?"
End Sub

' ============================================================================
' HELPER: Convert English keyboard input to Hebrew characters
' Maps QWERTY keys to their Hebrew equivalents on standard Israeli keyboard
' ============================================================================
Private Function ConvertEngToHeb(ByVal txt As String) As String
    ' English keys (lowercase) -> Hebrew ChrW codes
    ' Standard Israeli keyboard layout mapping
    Dim engKeys As String
    Dim hebCodes() As Long
    Dim i As Long
    
    ' Map: t=? e=? r=? a=? w=? y=? u=? i=? o=? p=?
    '      s=? d=? f=? g=? h=? j=? k=? l=?
    '      z=? x=? c=? v=? b=? n=? m=?
    '      ,=? .=? ;=? '=,
    
    ' Full mapping array - index by Asc of lowercase English letter
    Dim mapArr(0 To 127) As Long
    For i = 0 To 127: mapArr(i) = 0: Next i
    
    mapArr(116) = 1488  ' t -> ?
    mapArr(99) = 1489   ' c -> ?
    mapArr(100) = 1490  ' d -> ?
    mapArr(115) = 1491  ' s -> ?
    mapArr(118) = 1492  ' v -> ?
    mapArr(117) = 1493  ' u -> ?
    mapArr(122) = 1494  ' z -> ?
    mapArr(106) = 1495  ' j -> ?
    mapArr(121) = 1496  ' y -> ?
    mapArr(104) = 1497  ' h -> ?
    mapArr(108) = 1498  ' l -> ?
    mapArr(102) = 1499  ' f -> ?
    mapArr(107) = 1500  ' k -> ?
    mapArr(110) = 1502  ' n -> ?
    mapArr(98) = 1504   ' b -> ?
    mapArr(120) = 1505  ' x -> ?
    mapArr(103) = 1506  ' g -> ?
    mapArr(112) = 1508  ' p -> ?
    mapArr(109) = 1510  ' m -> ?
    mapArr(101) = 1511  ' e -> ?
    mapArr(114) = 1512  ' r -> ?
    mapArr(97) = 1513   ' a -> ?
    mapArr(44) = 1514   ' , -> ?
    mapArr(111) = 1501  ' o -> ?
    mapArr(105) = 1503  ' i -> ?
    mapArr(59) = 1507   ' ; -> ?
    mapArr(46) = 1509   ' . -> ?
    mapArr(119) = 1506  ' w -> ? (same as g in some layouts)
    
    Dim result As String
    result = ""
    Dim ch As String
    Dim code As Long
    
    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        code = AscW(ch)
        If code >= 0 And code <= 127 Then
            ' Convert uppercase to lowercase for lookup
            If code >= 65 And code <= 90 Then code = code + 32
            If mapArr(code) > 0 Then
                result = result & ChrW(mapArr(code))
            Else
                result = result & ch  ' Keep as-is (space, digits, etc.)
            End If
        Else
            result = result & ch  ' Already non-ASCII, keep as-is
        End If
    Next i
    
    ConvertEngToHeb = result
End Function

' ============================================================================
' RESET CLIENT FILTER - sets G12 back to "bachar/i" for full reports
' ============================================================================
Public Sub ResetClientFilter()
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    wsMain.Range("rngClientName").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)  ' "???/?"
End Sub

' ============================================================================
' RESET HOME DEFAULTS: Reset G5/G6/G9/G10 to default values
' Called from "???? ??????? ????" button at G13
' ============================================================================
Public Sub ResetHomeDefaults()
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    
    On Error Resume Next
    wsMain.Unprotect "Z961814r"
    On Error GoTo 0
    
    Application.EnableEvents = True
    Application.EnableEvents = False
    
    ' G5 = ????
    wsMain.Range("G5").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
    ' G6 = empty, restore green
    wsMain.Range("G6").Value = ""
    wsMain.Range("G6").Interior.Color = RGB(220, 240, 220)
    On Error Resume Next
    wsMain.Range("G6").Validation.Delete
    On Error GoTo 0
    wsMain.Range("G7").Value = ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493)
    ' G9 = ???/?
    wsMain.Range("G9").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497)
    ' G10 = ???/?, no gold, remove validation
    On Error Resume Next
    wsMain.Range("G10").Validation.Delete
    On Error GoTo 0
    wsMain.Range("G10").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497)
    wsMain.Range("G10").Interior.Color = RGB(220, 240, 220)
    
    Application.EnableEvents = True
    
    On Error Resume Next
    wsMain.Protect Password:="Z961814r", UserInterfaceOnly:=True
    On Error GoTo 0
    
    Application.Goto wsMain.Range("A1")
End Sub

' ============================================================================
' HELPER: Apply zebra striping to a result sheet based on its tab color
' ============================================================================
Private Sub ApplyZebraStriping(ByVal ws As Worksheet)
    On Error Resume Next
    Dim tabClr As Long
    tabClr = ws.Tab.Color
    If tabClr = 0 Then Exit Sub  ' no tab color set
    
    Dim rC As Long, gC As Long, bC As Long
    rC = tabClr Mod 256
    gC = (tabClr \ 256) Mod 256
    bC = (tabClr \ 65536) Mod 256
    
    ' Light: blend 92% toward white (almost white with hint of color)
    Dim zebraLight As Long
    zebraLight = RGB(rC + (255 - rC) * 0.92, gC + (255 - gC) * 0.92, bC + (255 - bC) * 0.92)
    ' Dark: blend 55% toward white (noticeably colored)
    Dim zebraDark As Long
    zebraDark = RGB(rC + (255 - rC) * 0.55, gC + (255 - gC) * 0.55, bC + (255 - bC) * 0.55)
    
    ' Determine data start row: if row 1 is merged (title), data starts at row 4; otherwise row 3
    Dim dataStart As Long
    If ws.Range("A1").MergeCells Then
        dataStart = 4  ' title=1, headers=2-3, data from 4
    Else
        dataStart = 3  ' headers=1-2, data from 3
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' AutoFit column A to fit text, ensure minimum width of 20
    ws.Columns(1).AutoFit
    If ws.Columns(1).ColumnWidth < 20 Then ws.Columns(1).ColumnWidth = 20

    ' Add thin light-gray borders to the data content area
    Dim dataRng As Range
    Set dataRng = ws.Range(ws.Cells(dataStart, 1), ws.Cells(lastRow, 16))
    dataRng.Borders(xlEdgeLeft).LineStyle = xlContinuous
    dataRng.Borders(xlEdgeLeft).Weight = xlHairline
    dataRng.Borders(xlEdgeLeft).Color = RGB(180, 180, 180)
    dataRng.Borders(xlEdgeRight).LineStyle = xlContinuous
    dataRng.Borders(xlEdgeRight).Weight = xlHairline
    dataRng.Borders(xlEdgeRight).Color = RGB(180, 180, 180)
    dataRng.Borders(xlEdgeTop).LineStyle = xlContinuous
    dataRng.Borders(xlEdgeTop).Weight = xlHairline
    dataRng.Borders(xlEdgeTop).Color = RGB(180, 180, 180)
    dataRng.Borders(xlEdgeBottom).LineStyle = xlContinuous
    dataRng.Borders(xlEdgeBottom).Weight = xlHairline
    dataRng.Borders(xlEdgeBottom).Color = RGB(180, 180, 180)
    dataRng.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    dataRng.Borders(xlInsideHorizontal).Weight = xlHairline
    dataRng.Borders(xlInsideHorizontal).Color = RGB(180, 180, 180)
    dataRng.Borders(xlInsideVertical).LineStyle = xlContinuous
    dataRng.Borders(xlInsideVertical).Weight = xlHairline
    dataRng.Borders(xlInsideVertical).Color = RGB(180, 180, 180)

    Dim zr As Long
    For zr = dataStart To lastRow
        If (zr - dataStart) Mod 2 = 0 Then
            ws.Range(ws.Cells(zr, 1), ws.Cells(zr, 16)).Interior.Color = zebraLight
        Else
            ws.Range(ws.Cells(zr, 1), ws.Cells(zr, 16)).Interior.Color = zebraDark
        End If
    Next zr

    ' Distinct background for totals row (last row with bold font)
    If lastRow >= dataStart Then
        If ws.Cells(lastRow, 1).Font.Bold Then
            ' Background: use the actual tab color (dark) for very bold emphasis
            ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 16)).Interior.Color = RGB(rC * 0.5, gC * 0.5, bC * 0.5)
            ' Font: white for guaranteed visibility on darker background
            ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 16)).Font.Color = RGB(255, 255, 255)
            ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 16)).Font.Bold = True
            ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 16)).Font.Size = 13
            ' Top border for separation
            ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 16)).Borders(xlEdgeTop).LineStyle = xlContinuous
            ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 16)).Borders(xlEdgeTop).Weight = xlMedium
            ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 16)).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
        End If
    End If

    ' Ensure RTL display
    ws.DisplayRightToLeft = True

    On Error GoTo 0
End Sub

' ============================================================================
' MACRO: SaveReportsToFolder - saves all result sheets as a single XLSX file
' to the Reports subfolder. Called from Button 4 on home page.
' ============================================================================
Public Sub SaveReportsToFolder()

10      On Error GoTo ERR_HANDLER
        Application.EnableEvents = False
        Application.ScreenUpdating = False

        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        ' Determine Reports folder path
        Dim reportsFolder As String
        reportsFolder = REPORTS_FOLDER()
        If Not fso.FolderExists(reportsFolder) Then fso.CreateFolder reportsFolder
        
        ' Get year from home page
        Dim yearVal As String
        yearVal = Trim$(CStr(ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("rngCurrentYear").Value2))
        
        ' Build list of result sheets to export
        Dim sheetNames() As String
        Dim sheetCount As Long
        sheetCount = 0
        ReDim sheetNames(1 To 7)
        
        If SheetExists(SHEET_SUMMARY()) Then sheetCount = sheetCount + 1: sheetNames(sheetCount) = SHEET_SUMMARY()
        If SheetExists(SHEET_COMPANIES()) Then sheetCount = sheetCount + 1: sheetNames(sheetCount) = SHEET_COMPANIES()
        If SheetExists(SHEET_BRANCH()) Then sheetCount = sheetCount + 1: sheetNames(sheetCount) = SHEET_BRANCH()
        If SheetExists(SHEET_MAINBRANCH()) Then sheetCount = sheetCount + 1: sheetNames(sheetCount) = SHEET_MAINBRANCH()
        If SheetExists(SHEET_TELLERS()) Then sheetCount = sheetCount + 1: sheetNames(sheetCount) = SHEET_TELLERS()
        If SheetExists(SHEET_AGENTS()) Then sheetCount = sheetCount + 1: sheetNames(sheetCount) = SHEET_AGENTS()
        If SheetExists(SHEET_MONTHS()) Then sheetCount = sheetCount + 1: sheetNames(sheetCount) = SHEET_MONTHS()
        
        If sheetCount = 0 Then
            ' "ein gilyonot dochot leshmira - haretz kodem kftor 2" = No report sheets to save - run Button 2 first
            MsgBoxU ChrW(1488) & ChrW(1497) & ChrW(1503) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1513) & ChrW(1502) & ChrW(1497) & ChrW(1512) & ChrW(1492) & " - " & ChrW(1492) & ChrW(1512) & ChrW(1509) & " " & ChrW(1511) & ChrW(1493) & ChrW(1491) & ChrW(1501) & " " & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " 2", vbExclamation
            Exit Sub
        End If
        
        ' Create array of sheet names for Copy
        Dim arrSheets() As String
        ReDim arrSheets(1 To sheetCount)
        Dim si As Long
        For si = 1 To sheetCount
            arrSheets(si) = sheetNames(si)
        Next si
        
        ' Unhide all sheets before Copy (in case some are VeryHidden)
        Dim wsTemp As Worksheet
        Dim hiddenSheets() As String
        Dim hiddenCount As Long
        hiddenCount = 0
        ReDim hiddenSheets(1 To ThisWorkbook.Worksheets.Count)
        For Each wsTemp In ThisWorkbook.Worksheets
            If wsTemp.Visible <> xlSheetVisible Then
                hiddenCount = hiddenCount + 1
                hiddenSheets(hiddenCount) = wsTemp.Name
                wsTemp.Visible = xlSheetVisible
            End If
        Next wsTemp
        
        ' Copy sheets to new workbook
        ThisWorkbook.Worksheets(arrSheets(1)).Copy
        Dim wbNew As Workbook
        Set wbNew = ActiveWorkbook
        
        ' Copy remaining sheets
        If sheetCount > 1 Then
            Dim si2 As Long
            For si2 = 2 To sheetCount
                ThisWorkbook.Worksheets(arrSheets(si2)).Copy After:=wbNew.Worksheets(wbNew.Worksheets.Count)
            Next si2
        End If
        
        ' Re-hide sheets that were hidden before
        Dim hi As Long
        For hi = 1 To hiddenCount
            ThisWorkbook.Worksheets(hiddenSheets(hi)).Visible = xlSheetVeryHidden
        Next hi
        
        ' Save as XLSX (no macros)
        Dim xlsxPath As String
        ' "dochot hashvaa" = comparison reports
        xlsxPath = reportsFolder & "\" & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1492) & ChrW(1513) & ChrW(1493) & ChrW(1493) & ChrW(1488) & ChrW(1492) & " " & yearVal & ".xlsx"
        
        Application.DisplayAlerts = False
        wbNew.SaveAs xlsxPath, xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        ' "hadochot nishm'ru behatzlacha!" = Reports saved successfully!
        MsgBoxU ChrW(1492) & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1504) & ChrW(1513) & ChrW(1502) & ChrW(1512) & ChrW(1493) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!", vbInformation
        
        ' Return to home page
        ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Activate
        Application.Goto ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("A1")
        
        Exit Sub
        
ERR_HANDLER:
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        On Error Resume Next
        ' Re-hide sheets that were unhidden
        Dim hiErr As Long
        For hiErr = 1 To hiddenCount
            ThisWorkbook.Worksheets(hiddenSheets(hiErr)).Visible = xlSheetVeryHidden
        Next hiErr
        MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1513) & ChrW(1502) & ChrW(1497) & ChrW(1512) & ChrW(1514) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & ": " & Err.Description, vbCritical
End Sub

' ============================================================================
' HELPER: HebrewToKey - Transliterate Hebrew text to uppercase Latin key
' Replaces spaces with underscore, maps Hebrew letters to Latin equivalents
' Example: "????? ???? ???" -> "ALVMH_LBYT_ASK"
' ============================================================================
Private Function HebrewToKey(ByVal hebText As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As Long
    Dim mapped As String
    
    hebText = Trim$(hebText)
    result = ""
    
    For i = 1 To Len(hebText)
        ch = AscW(Mid$(hebText, i, 1))
        Select Case ch
            Case 1488: mapped = "A"   ' ? alef
            Case 1489: mapped = "B"   ' ? bet
            Case 1490: mapped = "G"   ' ? gimel
            Case 1491: mapped = "D"   ' ? dalet
            Case 1492: mapped = "H"   ' ? he
            Case 1493: mapped = "V"   ' ? vav
            Case 1494: mapped = "Z"   ' ? zayin
            Case 1495: mapped = "CH"  ' ? chet
            Case 1496: mapped = "T"   ' ? tet
            Case 1497: mapped = "Y"   ' ? yod
            Case 1498: mapped = "K"   ' ? kaf sofit
            Case 1499: mapped = "K"   ' ? kaf
            Case 1500: mapped = "L"   ' ? lamed
            Case 1501: mapped = "M"   ' ? mem sofit
            Case 1502: mapped = "M"   ' ? mem
            Case 1503: mapped = "N"   ' ? nun sofit
            Case 1504: mapped = "N"   ' ? nun
            Case 1505: mapped = "S"   ' ? samech
            Case 1506: mapped = "A"   ' ? ayin
            Case 1507: mapped = "P"   ' ? pe sofit
            Case 1508: mapped = "P"   ' ? pe
            Case 1509: mapped = "TZ"  ' ? tsade sofit
            Case 1510: mapped = "TZ"  ' ? tsade
            Case 1511: mapped = "K"   ' ? kof
            Case 1512: mapped = "R"   ' ? resh
            Case 1513: mapped = "SH"  ' ? shin
            Case 1514: mapped = "T"   ' ? tav
            Case 32:   mapped = "_"   ' space -> underscore
            Case Else
                ' Keep ASCII letters/digits as-is, skip others
                If (ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Or (ch >= 48 And ch <= 57) Then
                    mapped = UCase$(Chr$(ch))
                Else
                    mapped = ""
                End If
        End Select
        result = result & mapped
    Next i
    
    ' Remove trailing/leading underscores and double underscores
    Do While InStr(result, "__") > 0
        result = Replace(result, "__", "_")
    Loop
    If Left$(result, 1) = "_" Then result = Mid$(result, 2)
    If Right$(result, 1) = "_" Then result = Left$(result, Len(result) - 1)
    
    HebrewToKey = result
End Function

' ============================================================================
' VIEW REPORTS FOLDER: Open the Reports folder in Windows Explorer
' Called from "View Reports" button on home page (button 5)
' ============================================================================
Public Sub ViewReportsFolder()
    Dim reportsPath As String
    reportsPath = REPORTS_FOLDER()
    If Dir(reportsPath, vbDirectory) = "" Then
        MsgBoxU ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1514) & " " & ChrW(1492) & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1492) & ".", vbExclamation
        Exit Sub
    End If
    Shell "explorer.exe " & Chr(34) & reportsPath & Chr(34), vbNormalFocus
End Sub

' ============================================================================
' NEW CLIENTS: Show clients with premium>0 in current year that don't exist
'              (or have premium=0) in the base year.
'              Maintains history - previously reported clients are marked with date.
'              New clients since last run are highlighted in yellow.
' Called from "New Clients" button on home page (button 6)
' ============================================================================
Public Sub NewClients()
    Dim wsMain As Worksheet
    Dim yearVal As String
    Dim refYear As String
    Dim curSheetName As String
    Dim refSheetName As String
    Dim wsCur As Worksheet
    Dim wsRef As Worksheet
    Dim wsOut As Worksheet
    Dim outSheetName As String
    Dim lastRowCur As Long
    Dim lastRowRef As Long
    Dim lastRowOut As Long
    Dim r As Long
    Dim outRow As Long
    Dim custKey As String
    Dim dictRefCust As Object
    Dim dictExisting As Object
    Dim premVal As Double
    Dim sheetExisted As Boolean
    Dim newCount As Long
    Dim todayDate As Date
    
    On Error GoTo NC_ERR
    
    todayDate = Date
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    yearVal = Trim$(CStr(wsMain.Range("rngCurrentYear").Value2))
    refYear = Trim$(CStr(wsMain.Range("rngBaseYear").Value2))
    
    If yearVal = "" Or refYear = "" Then
        MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1497) & ChrW(1512) & " " & ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & H_BASE() & " " & ChrW(1493) & ChrW(1513) & ChrW(1504) & ChrW(1492) & " " & ChrW(1504) & ChrW(1493) & ChrW(1499) & ChrW(1495) & ChrW(1497) & ChrW(1514) & ".", vbExclamation
        Exit Sub
    End If
    
    ' --- Check if master sheet already exists (already processed) ---
    If SheetExists(NC_MASTER_SHEET_NAME()) Then
        Dim alreadyMsg6 As String
        ' "lekochot chadashim kvar nivdeku letkufa zo. im birtzoncha litzpot bareshima lchatz ken."
        alreadyMsg6 = ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501) & " " & ChrW(1499) & ChrW(1489) & ChrW(1512) & " " & ChrW(1504) & ChrW(1489) & ChrW(1491) & ChrW(1511) & ChrW(1493) & " " & ChrW(1500) & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492) & " " & ChrW(1494) & ChrW(1493) & "." & vbCrLf & vbCrLf & _
            ChrW(1488) & ChrW(1501) & " " & ChrW(1489) & ChrW(1512) & ChrW(1510) & ChrW(1493) & ChrW(1504) & ChrW(1498) & " " & ChrW(1500) & ChrW(1510) & ChrW(1508) & ChrW(1493) & ChrW(1514) & " " & ChrW(1489) & ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1492) & " " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " " & ChrW(1499) & ChrW(1503) & "."
        If MsgBoxU(alreadyMsg6, vbYesNo + vbQuestion) = vbYes Then
            ' Show the master sheet
            ThisWorkbook.Worksheets(NC_MASTER_SHEET_NAME()).Visible = xlSheetVisible
            ThisWorkbook.Worksheets(NC_MASTER_SHEET_NAME()).Activate
        End If
        Exit Sub
    End If
    
    ' Find basis sheets (Hebrew name: ????_YYYY)
    curSheetName = H_BASE() & "_" & yearVal
    refSheetName = H_BASE() & "_" & refYear
    
    ' Check if basis sheets exist
    If Not SheetExists(curSheetName) Then
        If SheetExists("base_" & yearVal) Then
            curSheetName = "base_" & yearVal
        Else
            MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1492) & ChrW(1512) & ChrW(1497) & ChrW(1509) & " " & ChrW(1511) & ChrW(1493) & ChrW(1491) & ChrW(1501) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " 2 (" & ChrW(1497) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1501) & " " & ChrW(1493) & ChrW(1491) & ChrW(1493) & Chr(34) & ChrW(1495) & ChrW(1493) & ChrW(1514) & ").", vbExclamation
            Exit Sub
        End If
    End If
    If Not SheetExists(refSheetName) Then
        If SheetExists("base_" & refYear) Then
            refSheetName = "base_" & refYear
        Else
            MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1492) & ChrW(1512) & ChrW(1497) & ChrW(1509) & " " & ChrW(1511) & ChrW(1493) & ChrW(1491) & ChrW(1501) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " 2 (" & ChrW(1497) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1501) & " " & ChrW(1493) & ChrW(1491) & ChrW(1493) & Chr(34) & ChrW(1495) & ChrW(1493) & ChrW(1514) & ").", vbExclamation
            Exit Sub
        End If
    End If
    
    Set wsCur = ThisWorkbook.Worksheets(curSheetName)
    Set wsRef = ThisWorkbook.Worksheets(refSheetName)
    
    Application.ScreenUpdating = False
    
    ' ---- Build dictionary of all customers in reference year (with premium > 0) ----
    Set dictRefCust = CreateObject("Scripting.Dictionary")
    lastRowRef = wsRef.Cells(wsRef.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRowRef
        custKey = Trim$(CStr(wsRef.Cells(r, BASE_COL_CUSTOMER).Value2))
        If custKey <> "" Then
            premVal = 0
            On Error Resume Next
            premVal = CDbl(wsRef.Cells(r, BASE_COL_PREMIUM).Value2)
            On Error GoTo NC_ERR
            If premVal > 0 Then
                If Not dictRefCust.Exists(custKey) Then dictRefCust(custKey) = True
            End If
        End If
    Next r
    
    ' ---- Create temp output sheet (??????_?????) - always fresh ----
    outSheetName = NC_TEMP_SHEET_NAME()
    If SheetExists(outSheetName) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(outSheetName).Delete
        Application.DisplayAlerts = True
    End If
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = outSheetName
    wsOut.DisplayRightToLeft = True
    
    ' Row 2: Headers for temp sheet
    wsOut.Cells(2, 1).Value = ChrW(1513) & ChrW(1501) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)  ' ?? ????
    wsOut.Cells(2, 2).Value = ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)  ' ???? ????
    wsOut.Cells(2, 3).Value = H_COMPANY()  ' ????
    wsOut.Cells(2, 4).Value = H_BRANCH()  ' ???
    wsOut.Cells(2, 5).Value = H_PREMIUM()  ' ?????
    wsOut.Cells(2, 6).Value = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1492)  ' ????
    wsOut.Cells(2, 7).Value = ChrW(1514) & ChrW(1488) & ChrW(1512) & ChrW(1497) & ChrW(1498) & " " & ChrW(1491) & ChrW(1497) & ChrW(1493) & ChrW(1493) & ChrW(1495)  ' ????? ?????
    wsOut.Cells(2, 8).Value = ChrW(1505) & ChrW(1496) & ChrW(1496) & ChrW(1493) & ChrW(1505)  ' ?????
    With wsOut.Range("A2:H2")
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(0, 100, 0)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    lastRowOut = 2
    
    ' ---- Ensure master sheet (?????_???????_?????) exists ----
    Dim masterSheetName As String
    Dim wsMaster As Worksheet
    Dim masterExisted As Boolean
    Dim lastRowMaster As Long
    masterSheetName = NC_MASTER_SHEET_NAME()
    masterExisted = SheetExists(masterSheetName)
    
    If masterExisted Then
        Set wsMaster = ThisWorkbook.Worksheets(masterSheetName)
    Else
        Set wsMaster = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsMaster.Name = masterSheetName
        wsMaster.DisplayRightToLeft = True
        ' Row 2: Headers
        wsMaster.Cells(2, 1).Value = ChrW(1513) & ChrW(1501) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)  ' ?? ????
        wsMaster.Cells(2, 2).Value = ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)  ' ???? ????
        wsMaster.Cells(2, 3).Value = H_COMPANY()  ' ????
        wsMaster.Cells(2, 4).Value = H_BRANCH()  ' ???
        wsMaster.Cells(2, 5).Value = H_PREMIUM()  ' ?????
        wsMaster.Cells(2, 6).Value = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1492)  ' ????
        wsMaster.Cells(2, 7).Value = ChrW(1514) & ChrW(1488) & ChrW(1512) & ChrW(1497) & ChrW(1498) & " " & ChrW(1491) & ChrW(1497) & ChrW(1493) & ChrW(1493) & ChrW(1495)  ' ????? ?????
        wsMaster.Cells(2, 8).Value = ChrW(1505) & ChrW(1496) & ChrW(1496) & ChrW(1493) & ChrW(1505)  ' ?????
        With wsMaster.Range("A2:H2")
            .Font.Bold = True
            .Font.Size = 12
            .Interior.Color = RGB(0, 100, 0)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
    End If
    
    ' Build dictionary of already-existing customers in master sheet
    Set dictExisting = CreateObject("Scripting.Dictionary")
    lastRowMaster = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row
    If lastRowMaster >= 3 Then
        For r = 3 To lastRowMaster
            custKey = Trim$(CStr(wsMaster.Cells(r, 2).Value2))  ' Column B = customer number
            If custKey <> "" Then dictExisting(custKey) = True
        Next r
    End If
    
    ' ---- Scan current year - find new clients not in ref year and not already reported ----
    Dim dictCustPrem As Object
    Dim dictCustComm As Object
    Dim dictCustName As Object
    Dim dictCustComp As Object
    Dim dictCustBranch As Object
    Set dictCustPrem = CreateObject("Scripting.Dictionary")
    Set dictCustComm = CreateObject("Scripting.Dictionary")
    Set dictCustName = CreateObject("Scripting.Dictionary")
    Set dictCustComp = CreateObject("Scripting.Dictionary")
    Set dictCustBranch = CreateObject("Scripting.Dictionary")
    
    lastRowCur = wsCur.Cells(wsCur.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRowCur
        custKey = Trim$(CStr(wsCur.Cells(r, BASE_COL_CUSTOMER).Value2))
        If custKey = "" Then GoTo NextNCRow
        
        ' Skip if customer exists in reference year
        If dictRefCust.Exists(custKey) Then GoTo NextNCRow
        
        ' dictExisting is checked later when adding to master (not here - temp sheet gets all)
        
        premVal = 0
        On Error Resume Next
        premVal = CDbl(wsCur.Cells(r, BASE_COL_PREMIUM).Value2)
        On Error GoTo NC_ERR
        
        If premVal <= 0 Then GoTo NextNCRow
        
        ' Aggregate premium and commission per customer
        If Not dictCustPrem.Exists(custKey) Then
            dictCustPrem(custKey) = premVal
            On Error Resume Next
            dictCustComm(custKey) = CDbl(wsCur.Cells(r, BASE_COL_COMMISSION).Value2)
            On Error GoTo NC_ERR
            dictCustName(custKey) = Trim$(CStr(wsCur.Cells(r, BASE_COL_CUSTNAME).Value2))
            dictCustComp(custKey) = Trim$(CStr(wsCur.Cells(r, BASE_COL_COMPANY).Value2))
            dictCustBranch(custKey) = Trim$(CStr(wsCur.Cells(r, BASE_COL_BRANCHNAME).Value2))
        Else
            dictCustPrem(custKey) = dictCustPrem(custKey) + premVal
            On Error Resume Next
            dictCustComm(custKey) = dictCustComm(custKey) + CDbl(wsCur.Cells(r, BASE_COL_COMMISSION).Value2)
            On Error GoTo NC_ERR
        End If
NextNCRow:
    Next r
    
    ' ---- Write new clients to temp sheet ----
    outRow = lastRowOut + 1
    newCount = 0
    Dim k As Variant
    Dim statusFormula As String
    statusFormula = ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1495) & "," & ChrW(1500) & ChrW(1495) & ChrW(1494) & ChrW(1493) & ChrW(1512) & "," & ChrW(1489) & ChrW(1493) & ChrW(1510) & ChrW(1506) & ChrW(1492) & " " & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1512) & ChrW(1492) & "," & ChrW(1500) & ChrW(1488) & " " & ChrW(1502) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1497) & ChrW(1503)
    
    For Each k In dictCustPrem.keys
        ' Write to temp sheet
        wsOut.Cells(outRow, 1).Value = dictCustName(k)
        wsOut.Cells(outRow, 2).Value = k
        wsOut.Cells(outRow, 3).Value = dictCustComp(k)
        wsOut.Cells(outRow, 4).Value = dictCustBranch(k)
        wsOut.Cells(outRow, 5).Value = dictCustPrem(k)
        wsOut.Cells(outRow, 6).Value = dictCustComm(k)
        wsOut.Cells(outRow, 7).Value = todayDate
        wsOut.Cells(outRow, 8).Value = ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1495)  ' ????
        
        ' Also add to master sheet if not already there
        If Not dictExisting.Exists(CStr(k)) Then
            lastRowMaster = lastRowMaster + 1
            wsMaster.Cells(lastRowMaster, 1).Value = dictCustName(k)
            wsMaster.Cells(lastRowMaster, 2).Value = k
            wsMaster.Cells(lastRowMaster, 3).Value = dictCustComp(k)
            wsMaster.Cells(lastRowMaster, 4).Value = dictCustBranch(k)
            wsMaster.Cells(lastRowMaster, 5).Value = dictCustPrem(k)
            wsMaster.Cells(lastRowMaster, 6).Value = dictCustComm(k)
            wsMaster.Cells(lastRowMaster, 7).Value = todayDate
            wsMaster.Cells(lastRowMaster, 8).Value = ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1495)  ' ????
            With wsMaster.Cells(lastRowMaster, 8).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=statusFormula
            End With
            wsMaster.Range("A" & lastRowMaster & ":H" & lastRowMaster).Interior.Color = RGB(255, 255, 200)
            newCount = newCount + 1
        End If
        
        outRow = outRow + 1
    Next k
    
    ' Format temp sheet
    Dim finalRow As Long
    finalRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    If finalRow >= 3 Then
        wsOut.Range("E3:E" & finalRow).NumberFormat = "#,##0"
        wsOut.Range("F3:F" & finalRow).NumberFormat = "#,##0"
        wsOut.Range("G3:G" & finalRow).NumberFormat = "dd/mm/yyyy"
    End If
    wsOut.Columns("A:H").AutoFit
    
    ' Format master sheet
    Dim masterFinalRow As Long
    masterFinalRow = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row
    If masterFinalRow >= 3 Then
        wsMaster.Range("E3:E" & masterFinalRow).NumberFormat = "#,##0"
        wsMaster.Range("F3:F" & masterFinalRow).NumberFormat = "#,##0"
        wsMaster.Range("G3:G" & masterFinalRow).NumberFormat = "dd/mm/yyyy"
    End If
    wsMaster.Columns("A:H").AutoFit
    wsMaster.Tab.Color = RGB(255, 165, 0)  ' orange
    
    ' ---- Create/update buttons in column L on MASTER sheet ----
    Dim s As Shape
    On Error Resume Next
    For Each s In wsMaster.Shapes
        If Left$(s.Name, 3) = "btn" Then s.Delete
    Next s
    On Error GoTo NC_ERR
    
    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btnW As Double
    Dim btnH As Double
    Dim btnGap As Double
    
    btnLeft = wsMaster.Range("L1").Left
    btnTop = wsMaster.Range("L2").Top
    btnW = 160
    btnH = 28
    btnGap = btnH + 5  ' button height + spacing
    
    ' 1. Show All (light blue)
    Dim shpAll As Shape
    Set shpAll = wsMaster.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
    shpAll.Name = "btnShowAll"
    shpAll.Fill.ForeColor.RGB = RGB(70, 130, 180)
    shpAll.TextFrame.Characters.Text = ChrW(1492) & ChrW(1510) & ChrW(1490) & " " & ChrW(1492) & ChrW(1499) & ChrW(1500)  ' ??? ???
    shpAll.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shpAll.TextFrame.Characters.Font.Size = 12
    shpAll.TextFrame.Characters.Font.Bold = True
    shpAll.TextFrame.HorizontalAlignment = xlCenter
    shpAll.TextFrame.VerticalAlignment = xlCenter
    shpAll.Placement = xlFreeFloating
    shpAll.OnAction = "FilterNC_All"
    btnTop = btnTop + btnGap
    
    ' 2. Show Open (green)
    Dim shpOpen As Shape
    Set shpOpen = wsMaster.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
    shpOpen.Name = "btnShowOpen"
    shpOpen.Fill.ForeColor.RGB = RGB(34, 139, 34)
    shpOpen.TextFrame.Characters.Text = ChrW(1492) & ChrW(1510) & ChrW(1490) & " " & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1495) & ChrW(1497) & ChrW(1501)  ' ??? ??????
    shpOpen.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shpOpen.TextFrame.Characters.Font.Size = 12
    shpOpen.TextFrame.Characters.Font.Bold = True
    shpOpen.TextFrame.HorizontalAlignment = xlCenter
    shpOpen.TextFrame.VerticalAlignment = xlCenter
    shpOpen.Placement = xlFreeFloating
    shpOpen.OnAction = "FilterNC_Open"
    btnTop = btnTop + btnGap
    
    ' 3. Show Sold (purple)
    Dim shpSold As Shape
    Set shpSold = wsMaster.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
    shpSold.Name = "btnShowSold"
    shpSold.Fill.ForeColor.RGB = RGB(128, 0, 128)
    shpSold.TextFrame.Characters.Text = ChrW(1492) & ChrW(1510) & ChrW(1490) & " " & ChrW(1504) & ChrW(1502) & ChrW(1499) & ChrW(1512) & ChrW(1493)  ' ??? ?????
    shpSold.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shpSold.TextFrame.Characters.Font.Size = 12
    shpSold.TextFrame.Characters.Font.Bold = True
    shpSold.TextFrame.HorizontalAlignment = xlCenter
    shpSold.TextFrame.VerticalAlignment = xlCenter
    shpSold.Placement = xlFreeFloating
    shpSold.OnAction = "FilterNC_Sold"
    btnTop = btnTop + btnGap
    
    ' 4. Show Refused (red)
    Dim shpRefused As Shape
    Set shpRefused = wsMaster.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
    shpRefused.Name = "btnShowRefused"
    shpRefused.Fill.ForeColor.RGB = RGB(180, 30, 30)
    shpRefused.TextFrame.Characters.Text = ChrW(1492) & ChrW(1510) & ChrW(1490) & " " & ChrW(1505) & ChrW(1497) & ChrW(1512) & ChrW(1489) & ChrW(1493)  ' ??? ?????
    shpRefused.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shpRefused.TextFrame.Characters.Font.Size = 12
    shpRefused.TextFrame.Characters.Font.Bold = True
    shpRefused.TextFrame.HorizontalAlignment = xlCenter
    shpRefused.TextFrame.VerticalAlignment = xlCenter
    shpRefused.Placement = xlFreeFloating
    shpRefused.OnAction = "FilterNC_Refused"
    btnTop = btnTop + btnGap * 2  ' extra gap before utility buttons
    
    ' 5. Email (orange)
    Dim shpEmail As Shape
    Set shpEmail = wsMaster.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
    shpEmail.Name = "btnEmailNC"
    shpEmail.Fill.ForeColor.RGB = RGB(230, 130, 0)
    shpEmail.TextFrame.Characters.Text = ChrW(1513) & ChrW(1500) & ChrW(1495) & " " & ChrW(1489) & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500)  ' ??? ?????
    shpEmail.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shpEmail.TextFrame.Characters.Font.Size = 12
    shpEmail.TextFrame.Characters.Font.Bold = True
    shpEmail.TextFrame.HorizontalAlignment = xlCenter
    shpEmail.TextFrame.VerticalAlignment = xlCenter
    shpEmail.Placement = xlFreeFloating
    shpEmail.OnAction = "EmailNewClients"
    btnTop = btnTop + btnGap
    
    ' 6. Print (dark blue)
    Dim shpPrint As Shape
    Set shpPrint = wsMaster.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
    shpPrint.Name = "btnPrintNC"
    shpPrint.Fill.ForeColor.RGB = RGB(0, 51, 102)
    shpPrint.TextFrame.Characters.Text = ChrW(1492) & ChrW(1491) & ChrW(1508) & ChrW(1505)  ' ????
    shpPrint.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shpPrint.TextFrame.Characters.Font.Size = 12
    shpPrint.TextFrame.Characters.Font.Bold = True
    shpPrint.TextFrame.HorizontalAlignment = xlCenter
    shpPrint.TextFrame.VerticalAlignment = xlCenter
    shpPrint.Placement = xlFreeFloating
    shpPrint.OnAction = "PrintNewClients"
    btnTop = btnTop + btnGap
    
    ' 7. Back to home (dark grey)
    Dim shpHome As Shape
    Set shpHome = wsMaster.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
    shpHome.Name = "btnBackHome"
    shpHome.Fill.ForeColor.RGB = RGB(60, 60, 60)
    shpHome.TextFrame.Characters.Text = ChrW(1495) & ChrW(1494) & ChrW(1512) & ChrW(1492) & " " & ChrW(1500) & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)  ' ???? ??? ????
    shpHome.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shpHome.TextFrame.Characters.Font.Size = 12
    shpHome.TextFrame.Characters.Font.Bold = True
    shpHome.TextFrame.HorizontalAlignment = xlCenter
    shpHome.TextFrame.VerticalAlignment = xlCenter
    shpHome.Placement = xlFreeFloating
    shpHome.OnAction = "NavToIndex"
    
    Application.ScreenUpdating = True
    
    ' Show result and navigate to MASTER sheet
    wsMaster.Activate
    Application.Goto wsMaster.Range("A2")
    ' Calculate total in master
    Dim totalInMaster As Long
    totalInMaster = 0
    If lastRowMaster >= 3 Then totalInMaster = lastRowMaster - 2
    
    If newCount = 0 Then
        ' ?? ????? ?????? ????? ?????. ??"? ???????: X
        MsgBoxU ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501) & " " & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501) & "." & vbCrLf & ChrW(1505) & ChrW(1492) & Chr(34) & ChrW(1499) & " " & ChrW(1489) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & ": " & totalInMaster, vbInformation
    Else
        ' ????? X ?????? ????? ?????. ??"? ???????: Y
        MsgBoxU ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & newCount & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501) & " " & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501) & "." & vbCrLf & ChrW(1505) & ChrW(1492) & Chr(34) & ChrW(1499) & " " & ChrW(1489) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & ": " & totalInMaster, vbInformation
    End If
    
    Exit Sub
    
NC_ERR:
    Application.ScreenUpdating = True
    MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501) & ": " & Err.Description, vbCritical
End Sub

' ============================================================================
' FILTER NEW CLIENTS: Individual filter subs for each status view
' Called from filter buttons in column L of ??????_????? sheet
' ============================================================================
Private Sub FilterNC_ByStatus(ParamArray allowedStatuses() As Variant)
    Dim wsOut As Worksheet
    Dim outSheetName As String
    Dim lastRow As Long
    Dim r As Long
    Dim statusVal As String
    Dim showRow As Boolean
    Dim i As Long
    
    outSheetName = NC_MASTER_SHEET_NAME()
    If Not SheetExists(outSheetName) Then Exit Sub
    
    Set wsOut = ThisWorkbook.Worksheets(outSheetName)
    
    Application.ScreenUpdating = False
    
    ' First unhide all rows so End(xlUp) finds the real last row
    wsOut.Rows.Hidden = False
    lastRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' If no statuses passed, show all (already unhidden above)
    If UBound(allowedStatuses) < LBound(allowedStatuses) Then
        ' Already all visible
    Else
        For r = 3 To lastRow
            statusVal = Trim$(CStr(wsOut.Cells(r, 8).Value2))
            showRow = False
            For i = LBound(allowedStatuses) To UBound(allowedStatuses)
                If statusVal = CStr(allowedStatuses(i)) Then
                    showRow = True
                    Exit For
                End If
            Next i
            wsOut.Rows(r).Hidden = Not showRow
        Next r
    End If
    
    Application.ScreenUpdating = True
End Sub

Public Sub FilterNC_All()
    ' Show all rows - unhide everything from row 3 to end of used range
    Dim wsOut As Worksheet
    Dim outSheetName As String
    Dim lastRow As Long
    
    outSheetName = NC_MASTER_SHEET_NAME()
    If Not SheetExists(outSheetName) Then Exit Sub
    
    Set wsOut = ThisWorkbook.Worksheets(outSheetName)
    
    Application.ScreenUpdating = False
    ' First unhide all rows so End(xlUp) works correctly
    wsOut.Rows.Hidden = False
    Application.ScreenUpdating = True
End Sub

Public Sub FilterNC_Open()
    ' Show only "????" and "?????"
    Dim openStatus As String
    Dim returnStatus As String
    openStatus = ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1495)
    returnStatus = ChrW(1500) & ChrW(1495) & ChrW(1494) & ChrW(1493) & ChrW(1512)
    FilterNC_ByStatus openStatus, returnStatus
End Sub

Public Sub FilterNC_Sold()
    ' Show only "????? ?????"
    Dim soldStatus As String
    soldStatus = ChrW(1489) & ChrW(1493) & ChrW(1510) & ChrW(1506) & ChrW(1492) & " " & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1512) & ChrW(1492)
    FilterNC_ByStatus soldStatus
End Sub

Public Sub FilterNC_Refused()
    ' Show only "?? ???????"
    Dim refusedStatus As String
    refusedStatus = ChrW(1500) & ChrW(1488) & " " & ChrW(1502) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1497) & ChrW(1503)
    FilterNC_ByStatus refusedStatus
End Sub

' ============================================================================
' ASSIGN BUTTON MACROS: Fix OnAction for all main buttons
' Called from Workbook_Open to ensure buttons work without running SetupMainSheet
' ============================================================================
Public Sub AssignButtonMacros()
    Dim wsMain As Worksheet
    Dim shp As Shape
    On Error Resume Next
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    If wsMain Is Nothing Then Exit Sub
    For Each shp In wsMain.Shapes
        Select Case shp.Name
            Case "btnBuildReview": shp.OnAction = "BuildReview"
            Case "btnApplyCorrections": shp.OnAction = "ApplyCorrectionsAndBuildReports"
            Case "btnBuildPresentation": shp.OnAction = "BuildPresentation"
            Case "btnSaveReports": shp.OnAction = "SaveReportsToFolder"
            Case "btnViewReports": shp.OnAction = "ViewReportsFolder"
            Case "btnNewClients": shp.OnAction = "NewClients"
        End Select
    Next shp
    On Error GoTo 0
End Sub

' ============================================================================
' CHECK USER PERMISSIONS: Show/hide sheets based on Windows username
' Called from Workbook_Open (ThisWorkbook)
' User list in permissions section (rngSection_Permissions, A=username, B=access level)
' ============================================================================
Public Sub CheckUserPermissions()
    Dim wsMgmt As Worksheet
    Dim wsMain As Worksheet
    Dim winUser As String
    Dim accessLevel As String
    Dim lastRow As Long
    Dim r As Long
    Dim found As Boolean
    Dim ws As Worksheet
    Dim ncSheetName As String
    Dim homeSheetName As String
    Dim mgmtSheetName As String
    Dim fullAccess As String
    Dim limitedAccess As String
    Dim shp As Shape
    Dim permStartRow As Long
    Dim cellVal As String
    Dim creditVersionCell As String
    Dim dateCell As String
    Dim creditText As String
    
    On Error GoTo PERM_ERR
    
    fullAccess = ChrW(1502) & ChrW(1500) & ChrW(1488)  ' ???
    limitedAccess = ChrW(1502) & ChrW(1493) & ChrW(1490) & ChrW(1489) & ChrW(1500)  ' ?????
    ncSheetName = NC_MASTER_SHEET_NAME()  ' ?????_???????_????? (master sheet for limited users)
    homeSheetName = CONTROL_SHEET_NAME()
    mgmtSheetName = MANAGEMENT_SHEET_NAME()
    
    ' Get Windows username automatically
    winUser = LCase$(Trim$(Environ("USERNAME")))
    If Len(winUser) = 0 Then Exit Sub  ' cannot identify user
    
    ' Find user in permissions list (rngSection_Permissions+1 downward, A=username, B=access level)
    On Error Resume Next
    Set wsMgmt = ThisWorkbook.Names("rngSection_Permissions").RefersToRange.Parent
    permStartRow = ThisWorkbook.Names("rngSection_Permissions").RefersToRange.Row + 1
    If Err.Number <> 0 Then
        On Error GoTo PERM_ERR
        Exit Sub  ' Named Range not defined yet
    End If
    On Error GoTo PERM_ERR
    
    found = False
    r = permStartRow
    Do While True
        cellVal = Trim$(CStr(wsMgmt.Cells(r, 1).Value2))
        If UCase$(cellVal) = "EOD" Or cellVal = "" Then Exit Do
        If LCase$(cellVal) = winUser Then
            accessLevel = Trim$(CStr(wsMgmt.Cells(r, 2).Value2))
            creditVersionCell = Trim$(CStr(wsMgmt.Cells(r, 3).Value2))
            dateCell = Trim$(CStr(wsMgmt.Cells(r, 4).Value2))
            found = True
            Exit Do
        End If
        r = r + 1
    Loop
    
    ' If user not found in list, default to limited
    If Not found Then accessLevel = limitedAccess
    
    ' Apply permissions
    Application.ScreenUpdating = False
    
    Set wsMain = ThisWorkbook.Worksheets(homeSheetName)
    
    If accessLevel = fullAccess Then
        ' Full access - show all sheets, restore all buttons
        For Each ws In ThisWorkbook.Worksheets
            ws.Visible = xlSheetVisible
        Next ws
        ' Restore buttons to full functionality (re-assign OnAction and original colors)
        On Error Resume Next
        For Each shp In wsMain.Shapes
            Select Case shp.Name
                Case "btnBuildReview"
                    shp.OnAction = "BuildReview"
                    shp.Fill.ForeColor.RGB = RGB(0, 70, 140)  ' blue
                Case "btnApplyCorrections"
                    shp.OnAction = "ApplyCorrectionsAndBuildReports"
                    shp.Fill.ForeColor.RGB = RGB(0, 120, 60)  ' green
                Case "btnBuildPresentation"
                    shp.OnAction = "BuildPresentation"
                    shp.Fill.ForeColor.RGB = RGB(160, 80, 0)  ' orange
                Case "btnSaveReports"
                    shp.OnAction = "SaveReportsToFolder"
                    shp.Fill.ForeColor.RGB = RGB(100, 60, 140)  ' purple
                Case "btnViewReports"
                    shp.OnAction = "ViewReportsFolder"
                    shp.Fill.ForeColor.RGB = RGB(70, 180, 220)  ' light blue
            End Select
        Next shp
        On Error GoTo PERM_ERR
    Else
        ' Limited access - show only home + new clients, hide everything else
        ThisWorkbook.Worksheets(homeSheetName).Visible = xlSheetVisible
        If SheetExists(ncSheetName) Then
            ThisWorkbook.Worksheets(ncSheetName).Visible = xlSheetVisible
        End If
        
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> homeSheetName And ws.Name <> ncSheetName Then
                ws.Visible = xlSheetVeryHidden
            End If
        Next ws
        
        ' Disable buttons 1-5 (remove OnAction + grey them out) - don't hide!
        On Error Resume Next
        For Each shp In wsMain.Shapes
            Select Case shp.Name
                Case "btnBuildReview", "btnApplyCorrections", "btnBuildPresentation", "btnSaveReports", "btnViewReports"
                    shp.OnAction = ""  ' disable click
                    shp.Fill.ForeColor.RGB = RGB(180, 180, 180)  ' grey
            End Select
        Next shp
        On Error GoTo PERM_ERR
    End If
    
    ' Write credit+version to Column C cell, date to Column D cell
    creditText = ChrW(1504) & ChrW(1489) & ChrW(1504) & ChrW(1492) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1497) & ChrW(1491) & ChrW(1497) & " " & ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & " 054-6677396"
    On Error Resume Next
    If Len(creditVersionCell) > 0 And creditVersionCell <> "A23" Then
        On Error Resume Next
        wsMain.Range(creditVersionCell).ClearContents
        Err.Clear
    End If
    If Len(dateCell) > 0 Then
        wsMain.Range(dateCell).Value = Format$(Date, "dd/mm/yyyy")
        wsMain.Range(dateCell).Font.Size = 9
        wsMain.Range(dateCell).Font.Color = RGB(0, 0, 102)
        wsMain.Range(dateCell).Font.Bold = True
    End If
    On Error GoTo PERM_ERR
    
    ' Navigate to home page
    ThisWorkbook.Worksheets(homeSheetName).Activate
    
    Application.ScreenUpdating = True
    Exit Sub
    
PERM_ERR:
    Application.ScreenUpdating = True
    ' On error, default to full access so system is usable
End Sub

' ============================================================================
' EMAIL NEW CLIENTS: Export visible rows to temp Excel file and send via Outlook
' Called from "Send Email" button in row 1 of ??????_????? sheet
' ============================================================================
Public Sub EmailNewClients()
    Dim wsOut As Worksheet
    Dim outSheetName As String
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim outR As Long
    Dim c As Long
    Dim tempPath As String
    Dim olApp As Object
    Dim olMail As Object
    
    On Error GoTo EMAIL_ERR
    
    outSheetName = NC_MASTER_SHEET_NAME()
    If Not SheetExists(outSheetName) Then Exit Sub
    
    Set wsOut = ThisWorkbook.Worksheets(outSheetName)
    lastRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    lastCol = 8  ' A through H
    
    If lastRow < 3 Then
        MsgBoxU ChrW(1488) & ChrW(1497) & ChrW(1503) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1495) & ChrW(1492) & ".", vbExclamation
        Exit Sub
    End If
    
    ' Create temp workbook with visible rows only
    Set wbNew = Workbooks.Add(xlWBATWorksheet)
    Set wsNew = wbNew.Worksheets(1)
    wsNew.DisplayRightToLeft = True
    
    ' Copy headers (row 2)
    For c = 1 To lastCol
        wsNew.Cells(1, c).Value = wsOut.Cells(2, c).Value
    Next c
    wsNew.Range("A1:H1").Font.Bold = True
    
    ' Copy visible data rows
    outR = 2
    For r = 3 To lastRow
        If Not wsOut.Rows(r).Hidden Then
            For c = 1 To lastCol
                wsNew.Cells(outR, c).Value = wsOut.Cells(r, c).Value
            Next c
            outR = outR + 1
        End If
    Next r
    
    wsNew.Columns("A:H").AutoFit
    wsNew.Range("E2:E" & outR - 1).NumberFormat = "#,##0"
    wsNew.Range("F2:F" & outR - 1).NumberFormat = "#,##0"
    wsNew.Range("G2:G" & outR - 1).NumberFormat = "dd/mm/yyyy"
    
    ' Add status dropdown validation to column H (same list as source sheet)
    If outR > 2 Then
        Dim statusList As String
        statusList = ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1495) & "," & ChrW(1500) & ChrW(1495) & ChrW(1494) & ChrW(1493) & ChrW(1512) & "," & ChrW(1489) & ChrW(1493) & ChrW(1510) & ChrW(1506) & ChrW(1492) & " " & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1512) & ChrW(1492) & "," & ChrW(1500) & ChrW(1488) & " " & ChrW(1502) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1497) & ChrW(1503)
        With wsNew.Range("H2:H" & outR - 1).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=statusList
            .InCellDropdown = True
        End With
    End If
    
    ' Save temp file
    tempPath = Environ$("TEMP") & "\" & outSheetName & "_" & Format$(Date, "yyyy-mm-dd") & ".xlsx"
    Application.DisplayAlerts = False
    wbNew.SaveAs tempPath, xlOpenXMLWorkbook
    wbNew.Close False
    Application.DisplayAlerts = True
    
    ' Create email via Outlook
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    With olMail
        .Subject = outSheetName & " - " & Format$(Date, "dd/mm/yyyy")
        .Body = ChrW(1502) & ChrW(1510) & ChrW(1493) & ChrW(1512) & ChrW(1507) & " " & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501) & "."  ' ????? ???? ?????? ?????.
        .Attachments.Add tempPath
        .Display
    End With
    
    Set olMail = Nothing
    Set olApp = Nothing
    Exit Sub
    
EMAIL_ERR:
    Application.DisplayAlerts = True
    MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1495) & ChrW(1514) & " " & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & ": " & Err.Description, vbCritical
End Sub

' ============================================================================
' PRINT NEW CLIENTS: Print the visible rows of the new clients sheet
' Called from "Print" button in row 1 of ??????_????? sheet
' ============================================================================
Public Sub PrintNewClients()
    Dim wsOut As Worksheet
    Dim outSheetName As String
    Dim lastRow As Long
    
    On Error GoTo PRINT_ERR
    
    outSheetName = NC_MASTER_SHEET_NAME()
    If Not SheetExists(outSheetName) Then Exit Sub
    
    Set wsOut = ThisWorkbook.Worksheets(outSheetName)
    lastRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 3 Then
        MsgBoxU ChrW(1488) & ChrW(1497) & ChrW(1503) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1492) & ChrW(1491) & ChrW(1508) & ChrW(1505) & ChrW(1492) & ".", vbExclamation
        Exit Sub
    End If
    
    ' Set print area (row 2 = headers, row 3+ = data, columns A:H)
    wsOut.PageSetup.PrintArea = "A2:H" & lastRow
    wsOut.PageSetup.Orientation = xlLandscape
    wsOut.PageSetup.Zoom = False
    wsOut.PageSetup.FitToPagesWide = 1
    wsOut.PageSetup.FitToPagesTall = False
    ' RTL is set via sheet property (DisplayRightToLeft), not PageSetup
    
    ' Print with dialog so user can choose printer/settings
    wsOut.PrintOut Preview:=True
    
    Exit Sub
    
PRINT_ERR:
    MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1491) & ChrW(1508) & ChrW(1505) & ChrW(1492) & ": " & Err.Description, vbCritical
End Sub

' ============================================================================
' EXIT SYSTEM: Close the workbook/Excel
' Called from "Exit System" button on home page (B16)
' ============================================================================
Public Sub ExitSystem()
    If Application.Workbooks.Count > 1 Then
        Dim closeAll As Long
        closeAll = MsgBoxU(ChrW(1497) & ChrW(1513) & " " & ChrW(1511) & ChrW(1489) & ChrW(1510) & ChrW(1497) & ChrW(1501) & " " & ChrW(1504) & ChrW(1493) & ChrW(1505) & ChrW(1508) & ChrW(1497) & ChrW(1501) & " " & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1495) & ChrW(1497) & ChrW(1501) & ". " & ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1500) & ChrW(1505) & ChrW(1490) & ChrW(1493) & ChrW(1512) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1499) & ChrW(1500) & " Excel?", vbYesNo + vbQuestion)
        If closeAll = vbYes Then
            Application.Quit
        End If
    Else
        Application.Quit
    End If
End Sub

' ============================================================================
' SETTINGS NAVIGATION: Menu-based navigation for the settings sheet
' All sections stacked vertically in column A. Navigation scrolls to row.
' ============================================================================

' Vertical layout on settings sheet (??????):
'   A1   = ????? ???? (rngSection_FieldMap)
'   A57  = ????? (rngSection_BranchName)
'   A173 = ??????? (rngSection_Params) - includes B176=FILES, B177=REPORTS
'   A181 = ????? (rngSection_ReasonCode)
'   A196 = ?????? ????? (rngSection_PeriodLists)
'   A221 = ?????? (rngSection_Messages)
'   A240 = ?????? (rngSection_Permissions)
'   A249 = ?????? (rngSection_Clients)

' ============================================================================
' SETUP SECTION NAMED RANGES: Deletes old/misspelled names, creates correct ones
' Called from SetupMainSheet
' ============================================================================
Public Sub SetupSectionNamedRanges()
    Dim wsMgmt As Worksheet
    Dim nm As Name
    
    On Error Resume Next
    Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    On Error GoTo 0
    If wsMgmt Is Nothing Then Exit Sub
    
    ' --- Delete ALL existing rngSection_* names (including typos) ---
    On Error Resume Next
    For Each nm In ThisWorkbook.Names
        If LCase$(Left$(nm.Name, 10)) = "rngsection" Then
            nm.Delete
        End If
        ' Also catch old misspelled names
        If LCase$(Left$(nm.Name, 10)) = "rngsrction" Then
            nm.Delete
        End If
    Next nm
    
    ' Also delete old rngFILES_FOLDER and rngREPORTS_FOLDER (will be recreated with correct refs)
    ThisWorkbook.Names("rngFILES_FOLDER").Delete
    ThisWorkbook.Names("rngREPORTS_FOLDER").Delete
    Err.Clear
    On Error GoTo 0
    
    Dim offsetRow As Long
    offsetRow = 0
    If wsMgmt.Range("A1").Value = "MENU_AREA" Then offsetRow = 4

    ' --- Create section Named Ranges ---
    ThisWorkbook.Names.Add "rngSection_FieldMap", wsMgmt.Range("A" & 1 + offsetRow)
    ThisWorkbook.Names.Add "rngSection_BranchName", wsMgmt.Range("A" & 57 + offsetRow)
    ThisWorkbook.Names.Add "rngSection_Params", wsMgmt.Range("A" & 173 + offsetRow)
    ThisWorkbook.Names.Add "rngSection_ReasonCode", wsMgmt.Range("A" & 181 + offsetRow)
    ThisWorkbook.Names.Add "rngSection_PeriodLists", wsMgmt.Range("A" & 210 + offsetRow)
      ThisWorkbook.Names.Add "rngSection_Messages", wsMgmt.Range("A" & 250 + offsetRow)
      ThisWorkbook.Names.Add "rngSection_Permissions", wsMgmt.Range("A" & 300 + offsetRow)
      ThisWorkbook.Names.Add "rngSection_Clients", wsMgmt.Range("A" & 350 + offsetRow)
    
    ' --- Create folder path Named Ranges (Params section) ---
    ThisWorkbook.Names.Add "rngFILES_FOLDER", wsMgmt.Range("B" & 176 + offsetRow)
    ThisWorkbook.Names.Add "rngREPORTS_FOLDER", wsMgmt.Range("B" & 177 + offsetRow)
End Sub

' ----------------------------------------------------------------------------
' HELPER: Scroll settings sheet to a Named Range (section start row)
' ----------------------------------------------------------------------------
Private Sub SettingsScrollTo(ByVal sectionName As String)
10  Dim wsMgmt As Worksheet
20  On Error Resume Next
30  Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
40  On Error GoTo 0
50  If wsMgmt Is Nothing Then Exit Sub
    
60  On Error Resume Next
70  Application.GoTo Reference:=ThisWorkbook.Names(sectionName).RefersToRange, Scroll:=True
80  On Error GoTo 0
End Sub

' ----------------------------------------------------------------------------
' PUBLIC: Scroll to top of settings sheet ("Menu" / home of settings)
' ----------------------------------------------------------------------------
Public Sub NavSettings_Menu()
10  Dim wsMgmt As Worksheet
20  On Error Resume Next
30  Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
40  On Error GoTo 0
50  If wsMgmt Is Nothing Then Exit Sub
    
60  wsMgmt.Activate
70  Application.GoTo Reference:=wsMgmt.Range("A1"), Scroll:=True
End Sub

' ----------------------------------------------------------------------------
' PUBLIC: Scroll to FIELD MAPPING section
' ----------------------------------------------------------------------------
Public Sub NavSettings_FieldMap()
10  SettingsScrollTo "rngSection_FieldMap"
End Sub

' ----------------------------------------------------------------------------
' PUBLIC: Scroll to BRANCH NAMES section
' ----------------------------------------------------------------------------
Public Sub NavSettings_BranchName()
10  SettingsScrollTo "rngSection_BranchName"
End Sub

' ----------------------------------------------------------------------------
' PUBLIC: Scroll to PARAMETERS section
' ----------------------------------------------------------------------------
Public Sub NavSettings_Params()
10  SettingsScrollTo "rngSection_Params"
End Sub

' ----------------------------------------------------------------------------
' PUBLIC: Scroll to REASON CODES section
' ----------------------------------------------------------------------------
Public Sub NavSettings_ReasonCode()
10  SettingsScrollTo "rngSection_ReasonCode"
End Sub

' ----------------------------------------------------------------------------
' PUBLIC: Scroll to PERIOD LISTS section
' ----------------------------------------------------------------------------
Public Sub NavSettings_PeriodType()
10  SettingsScrollTo "rngSection_PeriodLists"
End Sub

' ----------------------------------------------------------------------------
' PUBLIC: Scroll to SYSTEM MESSAGES section
' ----------------------------------------------------------------------------
Public Sub NavSettings_Messages()
10  SettingsScrollTo "rngSection_Messages"
End Sub

' ----------------------------------------------------------------------------
' PUBLIC: Scroll to PERMISSIONS section
' ----------------------------------------------------------------------------
Public Sub NavSettings_Permissions()
10  Dim wsMgmt As Worksheet
20  On Error Resume Next
30  Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
40  On Error GoTo 0
50  If Not wsMgmt Is Nothing Then
        Application.GoTo Reference:=wsMgmt.Range("A300"), Scroll:=True
    End If
End Sub

' ----------------------------------------------------------------------------
' PUBLIC: Scroll to CLIENT LIST section
' ----------------------------------------------------------------------------
Public Sub NavSettings_Clients()
10  SettingsScrollTo "rngSection_Clients"
End Sub

' ----------------------------------------------------------------------------
' PUBLIC: Scroll to top (show all - no hiding needed in vertical layout)
' ----------------------------------------------------------------------------
Public Sub NavSettings_ShowAll()
10  Dim wsMgmt As Worksheet
20  On Error Resume Next
30  Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
40  On Error GoTo 0
50  If wsMgmt Is Nothing Then Exit Sub
60  wsMgmt.Activate
70  Application.GoTo Reference:=wsMgmt.Range("A1"), Scroll:=True
End Sub

' ----------------------------------------------------------------------------
' SETUP: Create navigation menu buttons on settings sheet (column I - always visible)
' Called from SetupMainSheet
Public Sub SetupSettingsMenu()
    Dim wsMgmt As Worksheet
    Dim shp As Shape
    Dim s As Shape
    Dim btnTop As Double
    Dim btnLeft As Double
    Dim btnW As Double
    Dim btnH As Double
    Dim btnGapH As Double
    Dim btnTop1 As Double
    Dim btnTop2 As Double
    Dim btnTop3 As Double
    Dim blockWidth As Double
    Dim startLeft As Double
    Dim initialLeft As Double
    
    On Error Resume Next
    Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    On Error GoTo 0
    If wsMgmt Is Nothing Then Exit Sub
    
    Application.ScreenUpdating = False

    ' 1. Auto-insert 4 blank rows for the floating menu if they don't exist!
    If wsMgmt.Range("A1").Value <> "MENU_AREA" Then
        wsMgmt.Rows("1:4").Insert Shift:=-4121 ' xlDown
        wsMgmt.Range("A1").Value = "MENU_AREA"
        wsMgmt.Rows("1:4").Interior.ColorIndex = -4142 ' xlNone
        wsMgmt.Range("A1").Font.ColorIndex = 2 ' White font to hide the text
        ' We must also update named ranges since rows shifted
        Call SetupSectionNamedRanges
    End If

    ' Make top rows larger to hold floating menu (3 rows of buttons)
    wsMgmt.Rows("1:4").RowHeight = 28
    
    ' Remove old menu buttons
    On Error Resume Next
    For Each s In wsMgmt.Shapes
        If Left(s.Name, 6) = "btnNav" Then s.Delete
    Next s
    On Error GoTo 0
    
    ' Safely freeze panes at row 5 so menu is always visible and no columns are frozen!
    wsMgmt.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    wsMgmt.Range("A5").Select
    ActiveWindow.FreezePanes = True
    
    ' Horizontal Floating Menu Dimensions (3x3 grid)
    btnW = 135
    btnH = 24
    btnGapH = btnW + 10
    btnTop1 = 5
    btnTop2 = 34
    btnTop3 = 63
    
    ' Calculate the rightmost edge of Column A
    startLeft = wsMgmt.Range("A1").Left + wsMgmt.Range("A1").Width
    
    ' Calculate the total width of 3 buttons to anchor the block to the right margin
    blockWidth = (3 * btnW) + (2 * 10)
    
    ' This is the leftmost starting position, ensuring the whole block sits against the right margin
    initialLeft = startLeft - blockWidth - 10
    
    ' CRITICAL PRECAUTION: Never allow buttons to be drawn at negative coordinates!
    If initialLeft < 5 Then initialLeft = 5
    
    ' =========================================================
    ' ROW 1
    ' =========================================================
    btnLeft = initialLeft
    
    ' 1. Field Mapping button (blue)
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop1, btnW, btnH)
    shp.Name = "btnNavFieldMap"
    shp.Fill.ForeColor.RGB = RGB(0, 70, 140)
    shp.TextFrame2.TextRange.Text = ChrW(1502) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1497) & " " & ChrW(1513) & ChrW(1491) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavSettings_FieldMap"
    btnLeft = btnLeft + btnGapH
    
    ' 2. Branch Names button (dark cyan)
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop1, btnW, btnH)
    shp.Name = "btnNavBranchName"
    shp.Fill.ForeColor.RGB = RGB(0, 100, 100)
    shp.TextFrame2.TextRange.Text = ChrW(1506) & ChrW(1504) & ChrW(1508) & ChrW(1497) & ChrW(1501)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavSettings_BranchName"
    btnLeft = btnLeft + btnGapH
    
    ' 3. Parameters button (dark green)
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop1, btnW, btnH)
    shp.Name = "btnNavParams"
    shp.Fill.ForeColor.RGB = RGB(0, 100, 0)
    shp.TextFrame2.TextRange.Text = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512) & ChrW(1497) & ChrW(1501)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavSettings_Params"
    
    ' =========================================================
    ' ROW 2
    ' =========================================================
    btnLeft = initialLeft
    
    ' 4. Reason Codes button (purple)
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop2, btnW, btnH)
    shp.Name = "btnNavReasonCode"
    shp.Fill.ForeColor.RGB = RGB(100, 60, 140)
    shp.TextFrame2.TextRange.Text = ChrW(1505) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavSettings_ReasonCode"
    btnLeft = btnLeft + btnGapH
    
    ' 5. Period Lists button (orange)
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop2, btnW, btnH)
    shp.Name = "btnNavPeriodType"
    shp.Fill.ForeColor.RGB = RGB(200, 100, 0)
    shp.TextFrame2.TextRange.Text = H_RESHIMOT() & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavSettings_PeriodType"
    btnLeft = btnLeft + btnGapH
    
    ' 6. System Messages button (teal)
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop2, btnW, btnH)
    shp.Name = "btnNavMessages"
    shp.Fill.ForeColor.RGB = RGB(0, 128, 128)
    shp.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1493) & ChrW(1491) & ChrW(1506) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavSettings_Messages"
    
    ' =========================================================
    ' ROW 3
    ' =========================================================
    btnLeft = initialLeft
    
    ' 7. Permissions button (dark brown)
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop3, btnW, btnH)
    shp.Name = "btnNavPermissions"
    shp.Fill.ForeColor.RGB = RGB(120, 60, 0)
    shp.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1512) & ChrW(1513) & ChrW(1488) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavSettings_Permissions"
    btnLeft = btnLeft + btnGapH
    
    ' 8. Client List button (steel blue)
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop3, btnW, btnH)
    shp.Name = "btnNavClients"
    shp.Fill.ForeColor.RGB = RGB(70, 130, 180)
    shp.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavSettings_Clients"
    btnLeft = btnLeft + btnGapH
    
    ' 9. Back to Home button (red)
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop3, btnW, btnH)
    shp.Name = "btnNavHome"
    shp.Fill.ForeColor.RGB = RGB(180, 0, 0)
    shp.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1494) & ChrW(1512) & ChrW(1492) & " " & ChrW(1500) & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 10
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavToIndex"
    
    Application.ScreenUpdating = True
End Sub
' ============================================================================
' HELPER: Populate client list on settings sheet column A starting at row 249
' Reads unique client names from all "basis" sheets
' Called from BuildReview (button 1) to keep list current
' ============================================================================
Public Sub UpdateClientList()
10  Dim wsMgmt As Worksheet
20  Dim ws As Worksheet
30  Dim dict As Object
40  Dim lastRow As Long
50  Dim r As Long
60  Dim cName As String
70  Dim arrAll As Variant
80  Dim i As Long, j As Long, tmp As String
    
90  On Error Resume Next
100 Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
110 On Error GoTo 0
120 If wsMgmt Is Nothing Then Exit Sub
    
130 Set dict = CreateObject("Scripting.Dictionary")
140 dict.CompareMode = vbTextCompare
    
    ' Collect unique client names from all "basis" sheets
150 For Each ws In ThisWorkbook.Worksheets
160     If InStr(1, ws.Name, H_BASE(), vbTextCompare) > 0 Then
170         lastRow = ws.Cells(ws.Rows.Count, BASE_COL_CUSTNAME).End(xlUp).Row
180         For r = 2 To lastRow
190             cName = Trim$(CStr(ws.Cells(r, BASE_COL_CUSTNAME).Value2))
200             If cName <> "" Then
210                 If Not dict.Exists(cName) Then dict.Add cName, 1
220             End If
230         Next r
240     End If
250 Next ws
    
260 If dict.Count = 0 Then Exit Sub
    
    ' Sort alphabetically
270 arrAll = dict.keys
280 For i = 0 To UBound(arrAll) - 1
290     For j = i + 1 To UBound(arrAll)
300         If arrAll(i) > arrAll(j) Then
310             tmp = arrAll(i): arrAll(i) = arrAll(j): arrAll(j) = tmp
320         End If
330     Next j
340 Next i
    
    ' Write to column A starting at rngSection_Clients on settings sheet
    Dim startRow As Long
350 startRow = ThisWorkbook.Names("rngSection_Clients").RefersToRange.Row
360 wsMgmt.Cells(startRow, 1).Value = ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1514) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)  ' "????? ??????"
370 wsMgmt.Cells(startRow, 1).Font.Bold = True
380 wsMgmt.Cells(startRow, 1).Font.Size = 12
    ' Clear old list (rows startRow+1 onwards in column A)
390 Dim clearLast As Long
400 clearLast = wsMgmt.Cells(wsMgmt.Rows.Count, 1).End(xlUp).Row
410 If clearLast >= startRow + 1 Then wsMgmt.Range(wsMgmt.Cells(startRow + 1, 1), wsMgmt.Cells(clearLast, 1)).ClearContents
    ' Write sorted names
420 For i = 0 To UBound(arrAll)
430     wsMgmt.Cells(startRow + 1 + i, 1).Value = arrAll(i)
440 Next i
    
    ' Update named range for client list (used by search)
450 On Error Resume Next
460 ThisWorkbook.Names("lst_clients").Delete
470 On Error GoTo 0
480 If UBound(arrAll) >= 0 Then
490     ThisWorkbook.Names.Add "lst_clients", wsMgmt.Range(wsMgmt.Cells(startRow + 1, 1), wsMgmt.Cells(startRow + 1 + UBound(arrAll), 1))
500 End If
End Sub

' ============================================================================
' SET TAB COLORS: Assign pastel colors to ALL sheets (including hidden)
' Called from SetupMainSheet
' ============================================================================
Public Sub SetAllTabColors()
    Dim ws As Worksheet
    Dim i As Long
    Dim colors() As Variant
    
    ' Pastel color palette (12 distinct colors)
    colors = Array( _
        RGB(173, 216, 230), _
        RGB(176, 226, 172), _
        RGB(255, 218, 185), _
        RGB(255, 182, 193), _
        RGB(204, 204, 255), _
        RGB(255, 255, 186), _
        RGB(255, 204, 153), _
        RGB(200, 230, 255), _
        RGB(220, 255, 220), _
        RGB(255, 230, 210), _
        RGB(230, 200, 255), _
        RGB(255, 240, 200))
    
    ' Apply cycling colors to ALL sheets (including hidden)
    i = 0
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Tab.Color = colors(i Mod 12)
        i = i + 1
    Next ws
    
    ' Override specific sheets with distinctive colors
    ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Tab.Color = RGB(0, 70, 140)
    ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME()).Tab.Color = RGB(128, 128, 128)
    ThisWorkbook.Worksheets(MATACH_SHEET_NAME()).Tab.Color = RGB(0, 80, 120)
    On Error GoTo 0
End Sub

' ============================================================================
' HELPER: Safe Conversion Functions to prevent Type Mismatch
' ============================================================================
Private Function SafeDouble(ByVal valInput As Variant) As Double
    On Error Resume Next
    If IsNumeric(valInput) Then SafeDouble = CDbl(valInput) Else SafeDouble = 0
End Function

Private Function SafeLong(ByVal valInput As Variant) As Long
    On Error Resume Next
    If IsNumeric(valInput) Then SafeLong = CLng(valInput) Else SafeLong = 0
End Function

Private Function SafeString(ByVal valInput As Variant) As String
    On Error Resume Next
    If IsError(valInput) Then
        SafeString = ""
    ElseIf IsNull(valInput) Then
        SafeString = ""
    Else
        SafeString = CStr(valInput)
    End If
End Function












' ============================================================================
' ONE-TIME MIGRATION: Fixes overlap between PeriodLists and Messages (V1.57)
' ============================================================================





