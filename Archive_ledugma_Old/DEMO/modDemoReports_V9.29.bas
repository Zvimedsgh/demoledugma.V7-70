Attribute VB_Name = "Module1"
'Attribute VB_Name = "modReports"
' ============================================================================
' MODULE: modReports' VERSION: 9.28 (2026-06-24 11:16:58)
' PURPOSE: Complete reporting system - BuildReview + ApplyCorrectionsAndBuildReports
'          VERSION 9.28
' ============================================================================
' CHANGES IN 9.28:
'   - Added ApplyDemoRestrictions to instantly apply UI blocks
'   - Added Data Validation to block manual typing via keyboard in G3:G4
' CHANGES IN 9.27:
'   - Fixed BuildReview missing year override
'   - Added Date and Time to version header
'   - Fixed syntax error with backslash in comments
'   - Blocked UI years (2024/2025) and forced backend processing to 2019/2020
' CHANGES IN 7.93:
'   - API fetch now gets both USD and EUR and updates K3:K4 accordingly
'   - Unique pastel colors for buttons 4, 5, 6 corresponding to 1, 2, 3
'   - Parameter column pastel color, values pastel color
'   - Currency rates colorful background
' ============================================================================
' CHANGES IN 7.75:
'   - Added AutoFitHomeToScreen to automatically zoom the UI.
'   - BuildPresentation: Forced yearVal=2025 and refYear=2024 for demo output.
' ============================================================================
' CHANGES IN 7.50:
'   - SearchClientName: search sheet is now always deleted-and-recreated
'     (was: reused if present). Fixes the case where the sheet was
'     hidden by the post-report cleanup, making the green Search button
'     appear unresponsive.
'   - Home-page button text "Clear selection" -> "Reset parameters".
'     Button widened so the longer label fits.
'   - New Public Sub RebuildHomeButtons: re-creates the 3 client-row
'     buttons (Search / All / Reset). Run from Macros dialog when an
'     existing workbook is missing one of these buttons.
' ----------------------------------------------------------------------------
' To lock this code so others cannot view it, in the VBA editor:
'   Tools -> VBAProject Properties -> Protection tab
'   tick "Lock project for viewing", set a password, OK
'   Save the workbook, close it, reopen. Code will be hidden until the
'   password is entered. (Note: VBA project passwords can be bypassed by
'   freely available tools, so do not rely on them for sensitive secrets;
'   they only deter casual viewing.)
' ============================================================================
' CHANGES IN 7.47:
'   - Module renamed and all internal references to the original agency
'     name removed (identifiers, temp-file names, and inline comments) so
'     this file can be shared with an external software house for review.
'   - Earlier version history was removed for the same reason.
' ============================================================================

' --- Windows API for Unicode MsgBox ---
#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

' --- v7.57: extra Win32 API for CBT hook + window positioning ---
#If VBA7 Then
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hMod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
#Else
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
#End If

Private Const HCBT_ACTIVATE As Long = 5
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SW_SHOW As Long = 5

#If VBA7 Then
    Private mHook As LongPtr
#Else
    Private mHook As Long
#End If
Private mHookX As Long
Private mHookY As Long




' --- General constants ---
Private Const MANAGEMENT_START_ROW As Long = 2
Private Const DATA_SHEET_NAME As String = "TmpClientPolicyListEx"

' --- NIHUL field definition table ---
Private Const COL_FIELD_NAME_HE As Long = 5
Private Const COL_FIELD_COLUMN As Long = 6
Private Const COL_FIELD_CHECKING As Long = 7
Private Const COL_FIELD_KEY As Long = 8

' --- NIHUL parameter table ---
Private Const COL_PARAM_NAME As Long = 10
Private Const COL_PARAM_VALUE As Long = 11

' --- NIHUL helper translation table ---
Private Const COL_HELPER_KEY As Long = 14
Private Const COL_HELPER_VALUE As Long = 15

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

' --- 7.33: sheet protection password (was duplicated in 5 places) ---
Public Const SHEET_PROTECT_PWD As String = "961814"

' --- 7.33: message text rows in NIHUL/hagdarot column S ---
Private Const MSG_COL As Long = 19
Private Const MSG_ROW_DONE_FOUND As Long = 2
Private Const MSG_ROW_ISSUES_TAIL As Long = 3
Private Const MSG_ROW_DONE_OK As Long = 4
Private Const MSG_ROW_LINE As Long = 5
Private Const MSG_ROW_ERROR As Long = 6
Private Const MSG_ROW_STEP As Long = 7
Private Const MSG_ROW_CONFIRM_TITLE As Long = 8
Private Const MSG_ROW_CONFIRM_LINE_1 As Long = 13
Private Const MSG_ROW_CONFIRM_LINE_4 As Long = 16
Private Const MSG_ROW_ASK_AGAIN As Long = 17
Private Const MSG_ROW_DONT_SHOW_FLAG As Long = 20

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

Private Function SOURCE_FOLDER() As String
    ' 7.35: priority chain: FILES_FOLDER param -> rngFILES_FOLDER named range -> hardcoded
    Dim p As String
    p = GetMgmtParam("FILES_FOLDER")
    If p = "" Then
        On Error Resume Next
        p = Trim$(CStr(ThisWorkbook.names("rngFILES_FOLDER").RefersToRange.Value2))
        On Error GoTo 0
    End If
    If p = "" Then
        ' Final fallback: "C:\???????? ???\SOURCE\"
        p = "C:\DEMO PROJECT\SOURCE\"
    End If
    If Right$(p, 1) <> "\" Then p = p & "\"
    SOURCE_FOLDER = p
End Function

' ============================================================================
' HELPER (7.35): Read a string parameter from the management sheet (J/K columns).
' Returns "" on any failure (sheet missing, parameter missing, etc.).
' ============================================================================
Private Function GetMgmtParam(ByVal paramName As String) As String
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    On Error GoTo 0
    If ws Is Nothing Then
        GetMgmtParam = ""
        Exit Function
    End If
    GetMgmtParam = GetStringParameter(ws, paramName)
End Function

' ============================================================================
' HELPER (7.35): Reports output folder (PPTX/PDF). Falls back to ThisWorkbook.Path.
' ============================================================================
Private Function REPORTS_FOLDER() As String
    Dim p As String
    p = GetMgmtParam("REPORTS_FOLDER")
    If p = "" Then p = ThisWorkbook.Path
    If Right$(p, 1) <> "\" Then p = p & "\"
    REPORTS_FOLDER = p
End Function

' ============================================================================
' HELPER (7.35): Agency display name (used in slide titles).
' Falls back to the legacy hardcoded "?"? ?????? ??????".
' ============================================================================
Private Function AGENCY_NAME() As String
    Dim p As String
    p = GetMgmtParam("Agency_Name")
    If p = "" Then
        ' Fallback: legacy hardcoded "?"? ?????? ??????"
        p = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1492) & ChrW(1491) & ChrW(1490) & ChrW(1502) & ChrW(1492)
    End If
    AGENCY_NAME = p
End Function

' ============================================================================
' HELPER (7.35): English short name (used to derive paths via formula in K4-K6).
' Falls back to "Demo".
' ============================================================================
Private Function ENGLISH_NAME() As String
    Dim p As String
    p = GetMgmtParam("English_Name")
    If p = "" Then p = "Demo"
    ENGLISH_NAME = p
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
    SHEET_MAINBRANCH = ChrW(1506) & ChrW(1504) & ChrW(1507) & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)
End Function

Private Function SHEET_TELLERS() As String
    ' tlerim
    SHEET_TELLERS = ChrW(1496) & ChrW(1500) & ChrW(1512) & ChrW(1497) & ChrW(1501)
End Function

Private Function SHEET_AGENTS() As String
    ' sochnim
    SHEET_AGENTS = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501)
End Function

Private Function SHEET_MONTHS() As String
    ' hodshim
    SHEET_MONTHS = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
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

' 7.33: "bachar/i" sentinel string (was inline ChrW chain in 7+ places)
Private Function TXT_BACHAR_I() As String
    TXT_BACHAR_I = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)
End Function

' 7.33: "????" placeholder for an unmapped main branch (excluded from reports)
Private Function TXT_INVALID_BRANCH() As String
    TXT_INVALID_BRANCH = ChrW(1495) & ChrW(1493) & ChrW(1489) & ChrW(1492)
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
' HELPER (7.37): Resolve a year to a worksheet, preferring an internal embedded
' sheet ("__src_<year>") over the external file. Caller must close wbExternal
' if it is non-Nothing on return.
' Returns Nothing if neither internal nor external source can be located.
' ============================================================================
Private Function OpenSourceFor(ByVal yearVal As String, ByRef wbExternal As Workbook) As Worksheet
        Set wbExternal = Nothing

        ' --- Try internal embedded sheet first ---
        Dim internalName As String
10      internalName = "__src_" & yearVal
20      On Error Resume Next
30      Set OpenSourceFor = ThisWorkbook.Worksheets(internalName)
40      On Error GoTo 0
50      If Not OpenSourceFor Is Nothing Then Exit Function

        ' --- Fall back to external YYYY.xlsx / YYYY.xls ---
        Dim srcPath As String
60      srcPath = FindSourceFile(yearVal)
70      If srcPath = "" Then
80          Set OpenSourceFor = Nothing
90          Exit Function
100     End If
110     Set wbExternal = Workbooks.Open(srcPath, ReadOnly:=True, UpdateLinks:=0)
120     Set OpenSourceFor = OpenDataSheet(wbExternal)
End Function


' ============================================================================
' HELPER: Open data worksheet from source workbook
' ============================================================================
Private Function OpenDataSheet(ByVal wb As Workbook) As Worksheet
10      On Error Resume Next
20      Dim ws As Worksheet
30      Set ws = wb.Worksheets(DATA_SHEET_NAME)
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
70      If InStr(1, periodType, ChrW$(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497), vbTextCompare) > 0 Then
            ' E4 contains Hebrew month name from NIHUL!R10:R21
            ' Match it against the month list to find month number
80          If periodDetail <> "" Then
90              Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
100             For monthIdx = 1 To 12
110                 monthName = Trim$(CStr(wsMgmt.Cells(9 + monthIdx, 18).Value2))
120                 If StrComp(periodDetail, monthName, vbTextCompare) = 0 Then
130                     minMonth = monthIdx
140                     maxMonth = monthIdx
150                     Exit For
160                 End If
170             Next monthIdx
180         End If

        ' "riv'oni" = quarterly
190     ElseIf InStr(1, periodType, ChrW$(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497), vbTextCompare) > 0 Then
            ' E4 = riv'on rishon/sheni/shlishi/revi'i from NIHUL!R5:R8
            ' Match by checking which quarter keyword is in the detail
200         If InStr(1, periodDetail, ChrW$(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1503), vbTextCompare) > 0 Then
210             minMonth = 1: maxMonth = 3
220         ElseIf InStr(1, periodDetail, ChrW(1513) & ChrW(1504) & ChrW(1497), vbTextCompare) > 0 Then
230             minMonth = 4: maxMonth = 6
240         ElseIf InStr(1, periodDetail, ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1513) & ChrW(1497), vbTextCompare) > 0 Then
250             minMonth = 7: maxMonth = 9
260         ElseIf InStr(1, periodDetail, ChrW(1512) & ChrW(1489) & ChrW(1497) & ChrW(1506) & ChrW(1497), vbTextCompare) > 0 Then
270             minMonth = 10: maxMonth = 12
280         End If

        ' "chatzi shnati" = half yearly
290     ElseIf InStr(1, periodType, ChrW$(1495) & ChrW(1510) & ChrW(1497), vbTextCompare) > 0 Then
            ' E4 = machatzit rishona/shniya from NIHUL!R2:R3
300         If InStr(1, periodDetail, ChrW$(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1504), vbTextCompare) > 0 Then
310             minMonth = 1: maxMonth = 6
320         ElseIf InStr(1, periodDetail, ChrW$(1513) & ChrW(1504) & ChrW(1497), vbTextCompare) > 0 Then
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
30      If InStr(1, v, ChrW$(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & ChrW(1514), vbTextCompare) > 0 Then
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
        Dim singleCode As String
        Dim singleText As String
        Dim currCode As Variant   ' 7.33: was implicit Variant

        ' --- Pre-run confirmation message (read from hagdarot S13-S16, title from S8) ---
        ' S20 stores "1" if user chose "don't show again"
        Dim confirmMsg As String
        Dim wsCleanup As Worksheet
        Dim iSheet As Long
        Dim wsMsgSrc As Worksheet
        Set wsMsgSrc = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
        If CStr(wsMsgSrc.Cells(20, 19).Value) <> "1" Then
            confirmMsg = CStr(wsMsgSrc.Cells(13, 19).Value) & vbNewLine & vbNewLine & CStr(wsMsgSrc.Cells(14, 19).Value) & vbNewLine & vbNewLine & CStr(wsMsgSrc.Cells(15, 19).Value) & vbNewLine & vbNewLine & CStr(wsMsgSrc.Cells(16, 19).Value)
            If MsgBoxU(confirmMsg, vbOKCancel + vbExclamation, CStr(wsMsgSrc.Cells(8, 19).Value)) <> vbOK Then
                Exit Sub
            End If
            ' Ask if user wants to keep showing this message
            ' S17 = "show this message again?" text
            If MsgBoxU(CStr(wsMsgSrc.Cells(17, 19).Value), vbYesNo + vbQuestion) = vbNo Then
                wsMsgSrc.Cells(20, 19).Value = "1"
            End If
        End If

290     On Error GoTo ERR_HANDLER

        ' Remove any leftover sheet protection
        Dim wsUp As Worksheet
        For Each wsUp In ThisWorkbook.Worksheets
            On Error Resume Next
            wsUp.Unprotect SHEET_PROTECT_PWD
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
        refBaseKeep = ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505) & "_" & refYearStr
        refBaseKeepOld = "base_" & refYearStr
375     For iSheet = ThisWorkbook.Worksheets.count To 1 Step -1
376         Set wsCleanup = ThisWorkbook.Worksheets(iSheet)
377         If wsCleanup.Name <> CONTROL_SHEET_NAME() And wsCleanup.Name <> MANAGEMENT_SHEET_NAME() And wsCleanup.Name <> refBaseKeep And wsCleanup.Name <> refBaseKeepOld And Left$(wsCleanup.Name, 6) <> "__src_" Then
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

        ' Load dollar exchange rate from Bank of Israel API (fallback to rngDOLAR cell)
        Dim dollarRate As Double
        Dim wsMain As Worksheet
502     Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
504     dollarRate = GetDollarRate(wsMain)

510     yearVal = "2020" ' Forced for demo
        ' yearVal = Trim$(CStr(ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("rngCurrentYear").Value2))
520     If yearVal = "" Then Err.Raise vbObjectError + 1002, "BuildReview", "B3 IS EMPTY"

        ' 7.37: try internal embedded sheet first, fall back to external file
545     Set wsSrc = OpenSourceFor(yearVal, wbSrc)
548     If wsSrc Is Nothing Then Err.Raise vbObjectError + 1003, "BuildReview", "SOURCE NOT FOUND FOR YEAR: " & yearVal

570     lastRow = wsSrc.Cells(wsSrc.Rows.count, 2).End(xlUp).Row
580     If lastRow < 2 Then Err.Raise vbObjectError + 1006, "BuildReview", "NO DATA ROWS IN SOURCE"

        ' Use year-specific REVIEW sheet name
590     revSheetName = REVIEW_SHEET_NAME() & "_" & yearVal
591     DeleteSheetIfExists revSheetName

600     Set wsRev = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
610     wsRev.Name = revSheetName
        wsRev.DisplayRightToLeft = True   ' v7.68: RTL like all other sheets

620     wsRev.Cells(1, 1).Value = dictHelper(HELPER_REVIEW_SOURCE_ROW_HEADER)

630     For i = 1 To cnt
640         wsRev.Cells(1, i + 1).Value = disp(i)
650     Next i

660     wsRev.Cells(1, cnt + 2).Value = dictHelper(HELPER_REVIEW_REASON_HEADER)

        ' Action dropdown column header
680     wsRev.Cells(1, cnt + 3).Value = ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) & ChrW(1492)
        ' Fix text column header
690     wsRev.Cells(1, cnt + 4).Value = ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1503)

700     outRow = 2


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
                    ' Currency conversion: column AA (27) - 1=dollar, 0/90=ILS
812                 currCode = wsSrc.Cells(r, RAW_CURRENCY).Value2
814                 If IsNumeric(currCode) Then
816                     If CLng(currCode) = 1 Then premiumNum = premiumNum * dollarRate
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
925                     singleText = ChrW(1506) & ChrW(1504) & ChrW(1507) & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1502) & ChrW(1494) & ChrW(1493) & ChrW(1492) & ChrW(1492)
926                 End If
930                 wsRev.Cells(outRow, 1).Value = r
931                 For j = 1 To cnt
932                     wsRev.Cells(outRow, j + 1).Value = wsSrc.Cells(r, cols(j)).Value2
933                 Next j
934                 wsRev.Cells(outRow, cnt + 2).Value = singleText
936                 outRow = outRow + 1
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
991     listsName = ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1493) & ChrW(1514)
992     DeleteSheetIfExists listsName
993     Set wsLists = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
994     wsLists.Name = listsName
995     wsLists.Visible = xlSheetVeryHidden

        ' Headers
996     wsLists.Cells(1, 1).Value = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492)
        wsLists.Cells(1, 2).Value = ChrW(1496) & ChrW(1500) & ChrW(1512)
        wsLists.Cells(1, 3).Value = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503)
        wsLists.Cells(1, 4).Value = ChrW(1506) & ChrW(1504) & ChrW(1507)
        wsLists.Cells(1, 5).Value = ChrW(1506) & ChrW(1504) & ChrW(1507) & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)

        ' Write data
997     If dictCompanies.count > 0 Then
            arrKeys = dictCompanies.keys
            For kk = 0 To UBound(arrKeys): wsLists.Cells(kk + 2, 1).Value = arrKeys(kk): Next kk
        End If
998     If dictTellers.count > 0 Then
            arrKeys = dictTellers.keys
            For kk = 0 To UBound(arrKeys): wsLists.Cells(kk + 2, 2).Value = arrKeys(kk): Next kk
        End If
999     If dictAgents.count > 0 Then
            arrKeys = dictAgents.keys
            For kk = 0 To UBound(arrKeys): wsLists.Cells(kk + 2, 3).Value = arrKeys(kk): Next kk
        End If
        If dictBranches.count > 0 Then
            arrKeys = dictBranches.keys
            For kk = 0 To UBound(arrKeys): wsLists.Cells(kk + 2, 4).Value = arrKeys(kk): Next kk
        End If
        If dictMainBranches.count > 0 Then
            arrKeys = dictMainBranches.keys
            For kk = 0 To UBound(arrKeys): wsLists.Cells(kk + 2, 5).Value = arrKeys(kk): Next kk
        End If

        ' Add dropdown validation for action column
1000    actionCol = cnt + 3
1010    If outRow > 2 Then
1020        Set rng = wsRev.Range(wsRev.Cells(2, actionCol), wsRev.Cells(outRow - 1, actionCol))
            ' Set default value: "ha'aver livdika" = transfer for review
1025        rng.Value = ChrW(1492) & ChrW(1506) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1492)
1030        On Error Resume Next
1035        rng.Validation.Delete
1040        rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=ChrW(1514) & ChrW(1511) & ChrW(1503) & "," & ChrW(1492) & ChrW(1514) & ChrW(1506) & ChrW(1500) & ChrW(1501) & "," & ChrW(1492) & ChrW(1506) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1492)
1050        rng.Validation.InCellDropdown = True
1060        On Error GoTo ERR_HANDLER
1070    End If

        ' 7.37: only close if external (internal wsSrc has no Workbook to close)
1080    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
1090    Set wbSrc = Nothing

1100    wsRev.Rows(1).Font.Bold = True
1110    wsRev.Columns.AutoFit

        ' Format action and fix columns to be clearly visible
1112    wsRev.Cells(1, actionCol).Interior.Color = RGB(255, 165, 0)
1113    wsRev.Cells(1, actionCol + 1).Interior.Color = RGB(255, 165, 0)
1114    wsRev.Columns(actionCol).ColumnWidth = 15
1115    wsRev.Columns(actionCol + 1).ColumnWidth = 30
        ' Light yellow fill for data area of action/fix columns
1116    If outRow > 2 Then
1117        wsRev.Range(wsRev.Cells(2, actionCol), wsRev.Cells(outRow - 1, actionCol + 1)).Interior.Color = RGB(255, 255, 200)
1118    End If

        ' Add "Done Updating" button on the review sheet
1119    Dim shpBtn As Shape
1120    Set shpBtn = wsRev.Shapes.AddShape(msoShapeRoundedRectangle, wsRev.Range("Q1").Left, wsRev.Range("Q1").Top, 160, 30)  ' v7.69: anchored to Q1
1121    shpBtn.Name = "btnSendForReview"
1122    shpBtn.Fill.ForeColor.RGB = RGB(180, 0, 0)
        ' "siymti le'adken" = done updating
1123    shpBtn.TextFrame2.TextRange.Text = ChrW(1505) & ChrW(1497) & ChrW(1497) & ChrW(1502) & ChrW(1514) & ChrW(1497) & " " & ChrW(1500) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1503)
1124    shpBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
1125    shpBtn.TextFrame2.TextRange.Font.Size = 12
1126    shpBtn.TextFrame2.TextRange.Font.Bold = msoTrue
1127    shpBtn.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
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

1240    If outRow > 2 Then
            ' "Finished - found X issues" - stay on review sheet
            MsgBoxU wsMsgSrc.Cells(2, 19).Value & (outRow - 2) & wsMsgSrc.Cells(3, 19).Value, vbInformation
        Else
            ' No issues found - go straight to home
            MsgBoxU wsMsgSrc.Cells(4, 19).Value, vbInformation
            ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Activate
        End If

1250    Exit Sub

ERR_HANDLER:
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
1360    MsgBoxU wsMsgSrc.Cells(5, 19).Value & errLine & vbCrLf & wsMsgSrc.Cells(6, 19).Value & errNum & vbCrLf & errSrc & vbCrLf & errDesc, vbCritical

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
        Dim transferredCount As Long   ' v7.60: separate count for "transfer for review"

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

440     On Error GoTo ERR_HANDLER

        ' Remove any leftover sheet protection
        Dim wsUp2 As Worksheet
        For Each wsUp2 In ThisWorkbook.Worksheets
            On Error Resume Next
            wsUp2.Unprotect SHEET_PROTECT_PWD
            On Error GoTo ERR_HANDLER
        Next wsUp2

        ' --- Validate required dropdown selections ---
        Dim wsCheck As Worksheet
        Set wsCheck = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        Dim selText As String
        selText = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)  ' "bechar/i"
        
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
        Set wsMsgSrc2 = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())

        ' Show prominent processing message on Main sheet (below currency area)
        Dim wsProgress As Worksheet
        Set wsProgress = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
470     With wsProgress.Range("B12")
            .Value = ChrW(1502) & ChrW(1506) & ChrW(1489) & ChrW(1491) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & "," & " " & ChrW(1488) & ChrW(1504) & ChrW(1488) & " " & ChrW(1492) & ChrW(1502) & ChrW(1514) & ChrW(1503) & "/" & ChrW(1497) & "." & "." & "."
            .Font.Size = 20
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
            .Interior.Color = RGB(255, 255, 200)
        End With
        wsProgress.Activate
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
600     yearVal = "2020" ' Forced for demo
        ' yearVal = Trim$(CStr(wsMain.Range("rngCurrentYear").Value2))
610     refYear = "2019" ' Forced for demo
        ' refYear = Trim$(CStr(wsMain.Range("rngBaseYear").Value2))
620     If yearVal = "" Or refYear = "" Then Err.Raise vbObjectError + 2001, "ApplyCorrections", "B2 OR B3 IS EMPTY"

630     debugStep = "LOAD_HELPER"
640     Set dictHelper = LoadHelperDictionary(wsMgmt)

650     debugStep = "LOAD_BRANCH"
660     Set dictBranch = LoadBranchMapping(wsMgmt)

670     debugStep = "GET_THRESHOLD"
680     threshold = GetNumericParameter(wsMgmt, PARAM_PREMIUM_THRESHOLD)

        ' Load dollar exchange rate from Bank of Israel API (fallback to rngDOLAR cell)
        Dim dollarRate As Double
682     dollarRate = GetDollarRate(wsMain)

690     debugStep = "GET_PERIOD"
700     GetMonthRange wsMain, minMonth, maxMonth
        dateCol = GetDateColumn(wsMain)

        ' ---- Read corrections from REVIEW ----
710     debugStep = "READ_CORRECTIONS"
720     Set dictCorrections = CreateObject("Scripting.Dictionary")
730     Set dictIgnore = CreateObject("Scripting.Dictionary")
740     corrCount = 0
750     ignoreCount = 0
760     unhandledCount = 0
770     reviewCount = 0
        transferredCount = 0   ' v7.60

        ' Read corrections from year-specific REVIEW sheet
780     curRevName = REVIEW_SHEET_NAME() & "_" & yearVal
781     If SheetExists(curRevName) Then
790         Set wsRev = ThisWorkbook.Worksheets(curRevName)
800         revLastRow = wsRev.Cells(wsRev.Rows.count, 1).End(xlUp).Row

            ' Find action, fix, and reason columns by scanning header row
815         hdrActionText = ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) & ChrW(1492)
816         hdrFixText = ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1503)
            ' Hebrew: "sibat hriga" = reason header from helper dictionary
817         hdrReasonText = dictHelper(HELPER_REVIEW_REASON_HEADER)
818         actionColIdx = 0
819         fixColIdx = 0
820         reasonColIdx = 0
821         revLastCol = wsRev.Cells(1, wsRev.Columns.count).End(xlToLeft).Column
822         For hdrCol = 1 To revLastCol + 2
823             If StrComp(Trim$(CStr(wsRev.Cells(1, hdrCol).Value2)), hdrActionText, vbTextCompare) = 0 Then actionColIdx = hdrCol
824             If StrComp(Trim$(CStr(wsRev.Cells(1, hdrCol).Value2)), hdrFixText, vbTextCompare) = 0 Then fixColIdx = hdrCol
825             If StrComp(Trim$(CStr(wsRev.Cells(1, hdrCol).Value2)), hdrReasonText, vbTextCompare) = 0 Then reasonColIdx = hdrCol
826         Next hdrCol
827         If actionColIdx = 0 Then actionColIdx = revLastCol - 1
828         If fixColIdx = 0 Then fixColIdx = revLastCol
829         If reasonColIdx = 0 Then reasonColIdx = actionColIdx - 1

830         If revLastRow >= 2 Then
            ' Track per-source-row: dictHasFix stores a Dictionary of reason->fix per row
            Set dictHasFix = CreateObject("Scripting.Dictionary")
            Set dictHasIgnore = CreateObject("Scripting.Dictionary")
            Set dictHasUnhandled = CreateObject("Scripting.Dictionary")

840             For r = 2 To revLastRow
850                 reviewCount = reviewCount + 1
860                 srcRowNum = CLng(wsRev.Cells(r, 1).Value2)
870                 actionText = Trim$(CStr(wsRev.Cells(r, actionColIdx).Value2))
880                 fixText = Trim$(CStr(wsRev.Cells(r, fixColIdx).Value2))
890                 reasonText = Trim$(CStr(wsRev.Cells(r, reasonColIdx).Value2))

900                 If InStr(1, actionText, ChrW(1514) & ChrW(1511) & ChrW(1503), vbTextCompare) > 0 Then
                        ' Store reason->fix pair in a sub-dictionary per source row
910                     If Not dictHasFix.Exists(CStr(srcRowNum)) Then
920                         Set dictRowFixes = CreateObject("Scripting.Dictionary")
930                         Set dictHasFix(CStr(srcRowNum)) = dictRowFixes
935                     Else
937                         Set dictRowFixes = dictHasFix(CStr(srcRowNum))
940                     End If
950                     dictRowFixes(reasonText) = fixText
960                     corrCount = corrCount + 1
970                 ElseIf InStr(1, actionText, ChrW(1492) & ChrW(1514) & ChrW(1506) & ChrW(1500) & ChrW(1501), vbTextCompare) > 0 Then
980                     dictHasIgnore(CStr(srcRowNum)) = True
990                     ignoreCount = ignoreCount + 1
                        ' "ha'aver livdika" = transfer for review (v7.60: own counter)
992                 ElseIf InStr(1, actionText, ChrW(1492) & ChrW(1506) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1492), vbTextCompare) > 0 Then
994                     dictHasIgnore(CStr(srcRowNum)) = True
996                     transferredCount = transferredCount + 1
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
1040    debugStep = "OPEN_SRC"
        ' 7.37: try internal embedded sheet first, fall back to external file
1080    Set wsSrc = OpenSourceFor(yearVal, wbSrc)
1085    If wsSrc Is Nothing Then Err.Raise vbObjectError + 2002, "ApplyCorrections", "SOURCE NOT FOUND: " & yearVal

1100    debugStep = "BUILD_BASE_CURRENT"
1110    baseSheetName = ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505) & "_" & yearVal
        ' 7.34: was ~150 lines of cell-by-cell base build; now one call
1130    Set wsBase = BuildBaseSheet(wsSrc, baseSheetName, yearVal, dateCol, dollarRate, threshold, dictBranch, dictCorrections, dictIgnore, countCurrent)

        ' 7.37: only close if external
1950    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
1960    Set wbSrc = Nothing

        ' ---- Load or build base for reference year ----
1970    debugStep = "BUILD_BASE_REF"
1980    refBaseSheetName = ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505) & "_" & refYear

        ' If ref base already exists (with manual corrections), reuse it
        ' Check both Hebrew and old English name
1990    If SheetExists(refBaseSheetName) Then
1995        Set wsBaseRef = ThisWorkbook.Worksheets(refBaseSheetName)
        ' v7.44: count rows in existing sheet so summary doesn't show 0 for ref year
1996        countRef = wsBaseRef.Cells(wsBaseRef.Rows.count, 1).End(xlUp).Row - 1
1997        If countRef < 0 Then countRef = 0
2000        GoTo BUILD_COMPARISONS
2002    ElseIf SheetExists("base_" & refYear) Then
2003        Set wsBaseRef = ThisWorkbook.Worksheets("base_" & refYear)
2004        wsBaseRef.Name = refBaseSheetName
        ' v7.44: count rows for legacy English name path too
2006        countRef = wsBaseRef.Cells(wsBaseRef.Rows.count, 1).End(xlUp).Row - 1
2007        If countRef < 0 Then countRef = 0
2008        GoTo BUILD_COMPARISONS
        End If

        ' Build ref base from source (7.37: internal first, external fallback)
2030    Set wsRef = OpenSourceFor(refYear, wbRef)
2035    If wsRef Is Nothing Then Err.Raise vbObjectError + 2003, "ApplyCorrections", "REF SOURCE NOT FOUND: " & refYear

        ' 7.34: was ~80 lines of duplicate cell-by-cell ref base build
' Read corrections from ref year REVIEW sheet if exists
2112    Set dictRefCorr = CreateObject("Scripting.Dictionary")
2113    Set dictRefIgnore = CreateObject("Scripting.Dictionary")
2114    refRevName = REVIEW_SHEET_NAME() & "_" & refYear
    If SheetExists(refRevName) Then
        Set wsRefRev = ThisWorkbook.Worksheets(refRevName)
2117    refRevLastRow = wsRefRev.Cells(refRevLastRow, 1).End(xlUp).Row
2118    If refRevLastRow >= 2 Then
                ' Find action, fix, and reason columns
2119            refActionColIdx = 0: refFixColIdx = 0: refReasonColIdx = 0
2120            refRevLastCol = wsRefRev.Cells(1, wsRefRev.Columns.count).End(xlToLeft).Column
2121            For refHdrCol = 1 To refRevLastCol + 2
2122                If StrComp(Trim$(CStr(wsRefRev.Cells(1, refHdrCol).Value2)), hdrActionText, vbTextCompare) = 0 Then refActionColIdx = refHdrCol
2123                If StrComp(Trim$(CStr(wsRefRev.Cells(1, refHdrCol).Value2)), hdrFixText, vbTextCompare) = 0 Then refFixColIdx = refHdrCol
2124                If StrComp(Trim$(CStr(wsRefRev.Cells(1, refHdrCol).Value2)), hdrReasonText, vbTextCompare) = 0 Then refReasonColIdx = refHdrCol
2125            Next refHdrCol
2126            If refActionColIdx = 0 Then refActionColIdx = refRevLastCol - 1
2127            If refFixColIdx = 0 Then refFixColIdx = refRevLastCol
2128            If refReasonColIdx = 0 Then refReasonColIdx = refActionColIdx - 1
                ' Build ref correction and ignore dicts (reason-based)
2130            Set dictRefHasFix = CreateObject("Scripting.Dictionary")
2131            Set dictRefHasIgnore = CreateObject("Scripting.Dictionary")
2132            Set dictRefHasUnhandled = CreateObject("Scripting.Dictionary")
2134            For r = 2 To refRevLastRow
2136                refSrcRowNum = CLng(wsRefRev.Cells(r, 1).Value2)
2138                refActionText = Trim$(CStr(wsRefRev.Cells(r, refActionColIdx).Value2))
2140                refFixText = Trim$(CStr(wsRefRev.Cells(r, refFixColIdx).Value2))
2142                refReasonText = Trim$(CStr(wsRefRev.Cells(r, refReasonColIdx).Value2))
2144                If InStr(1, refActionText, ChrW(1514) & ChrW(1511) & ChrW(1503), vbTextCompare) > 0 Then
2146                    If Not dictRefHasFix.Exists(CStr(refSrcRowNum)) Then
2148                        Set dictRefRowFixes = CreateObject("Scripting.Dictionary")
2150                        Set dictRefHasFix(CStr(refSrcRowNum)) = dictRefRowFixes
2151                    Else
2153                        Set dictRefRowFixes = dictRefHasFix(CStr(refSrcRowNum))
2155                    End If
2157                    dictRefRowFixes(refReasonText) = refFixText
2159                ElseIf InStr(1, refActionText, ChrW(1492) & ChrW(1514) & ChrW(1506) & ChrW(1500) & ChrW(1501), vbTextCompare) > 0 Then
2161                    dictRefHasIgnore(CStr(refSrcRowNum)) = True
                        ' "ha'aver livdika" = transfer for review - treat as ignore
2162                ElseIf InStr(1, refActionText, ChrW(1492) & ChrW(1506) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1492), vbTextCompare) > 0 Then
2163                    dictRefHasIgnore(CStr(refSrcRowNum)) = True
2164                Else
2165                    dictRefHasUnhandled(CStr(refSrcRowNum)) = True
2167                End If
2169            Next r
                ' Merge: fix wins
2171            Set refAllKeys = CreateObject("Scripting.Dictionary")
2173            For Each refSrcKey In dictRefHasFix.keys: refAllKeys(refSrcKey) = True: Next
2175            For Each refSrcKey In dictRefHasIgnore.keys: refAllKeys(refSrcKey) = True: Next
2177            For Each refSrcKey In dictRefHasUnhandled.keys: refAllKeys(refSrcKey) = True: Next
2179            For Each refSrcKey In refAllKeys.keys
2181                If dictRefHasFix.Exists(refSrcKey) Then
2183                    Set dictRefCorr(refSrcKey) = dictRefHasFix(refSrcKey)
2185                ElseIf Not dictRefHasUnhandled.Exists(refSrcKey) Then
2187                    If dictRefHasIgnore.Exists(refSrcKey) Then dictRefIgnore(refSrcKey) = True
2189                End If
2191            Next refSrcKey
2193        End If
2195    End If
        ' 7.34: ref base built via unified helper (Variant arrays)
3170    Set wsBaseRef = BuildBaseSheet(wsRef, refBaseSheetName, refYear, dateCol, dollarRate, threshold, dictBranch, dictRefCorr, dictRefIgnore, countRef)

        ' 7.37: only close if external
3620    If Not wbRef Is Nothing Then wbRef.Close SaveChanges:=False
3630    Set wbRef = Nothing

BUILD_COMPARISONS:
        ' ---- Build comparison sheets (with optional filter from rngFilterType/rngFilterValue) ----
        Dim filterType As String
        Dim filterValue As String
        filterType = ""
        filterValue = ""
        On Error Resume Next
        filterType = Trim$(CStr(wsMain.Range("rngFilterType").Value2))
        filterValue = Trim$(CStr(wsMain.Range("rngFilterValue").Value2))
        On Error GoTo ERR_HANDLER

        ' ---- Client name filter from G12/rngClientName ----
        Dim clientFilter As String
        clientFilter = ""
        On Error Resume Next
        clientFilter = Trim$(CStr(wsMain.Range("rngClientName").Value2))
        On Error GoTo ERR_HANDLER
        ' If default text or empty, no client filter
        If clientFilter = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497) Or clientFilter = "" Then
            clientFilter = ""
        End If

        ' Resolve filterType to a base column for cross-filtering
        Dim filterCol As Long
        Dim selectText As String
        selectText = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)
        filterCol = 0
        If filterType <> "" And filterType <> selectText And filterValue <> "" And filterValue <> selectText Then
            ' Map Hebrew filter type to BASE_COL
            If StrComp(filterType, ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492), vbTextCompare) = 0 Then
                filterCol = BASE_COL_COMPANY
            ElseIf StrComp(filterType, ChrW(1496) & ChrW(1500) & ChrW(1512), vbTextCompare) = 0 Then
                filterCol = BASE_COL_TELLER
            ElseIf StrComp(filterType, ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503), vbTextCompare) = 0 Then
                filterCol = BASE_COL_AGENTNAME
            ElseIf StrComp(filterType, ChrW(1506) & ChrW(1504) & ChrW(1507) & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494), vbTextCompare) = 0 Then
                filterCol = BASE_COL_MAINBRANCH
            ElseIf StrComp(filterType, ChrW(1506) & ChrW(1504) & ChrW(1507), vbTextCompare) = 0 Then
                filterCol = BASE_COL_BRANCHNAME
            End If
        End If

        ' Build title text from home page parameters
        ' Format: "hashvaat [sheet] | [refYear] mul [yearVal] | [periodType] - [periodValue] | [filterType]: [filterValue]"
        Dim titlePrefix As String
        Dim mulText As String
        mulText = ChrW(1502) & ChrW(1493) & ChrW(1500)  ' "mul"
        titlePrefix = refYear & " " & mulText & " " & yearVal
        
        Dim prdType As String
        Dim prdValue As String
        prdType = Trim$(CStr(wsMain.Range("rngPeriodType").Value2))
        prdValue = ""
        On Error Resume Next
        prdValue = Trim$(CStr(wsMain.Range("rngPeriodValue").Value2))
        On Error GoTo ERR_HANDLER
        If prdValue <> "" And prdValue <> selectText Then
            titlePrefix = titlePrefix & " | " & prdValue
        ElseIf prdType <> "" And prdType <> selectText Then
            titlePrefix = titlePrefix & " | " & prdType
        End If
        
        If filterCol > 0 And filterValue <> "" And filterValue <> selectText Then
            titlePrefix = titlePrefix & " | " & filterType & ": " & filterValue
        End If
        
        ' Add client name to title if filtered
        If clientFilter <> "" Then
            ' "lakoach" = ????
            titlePrefix = titlePrefix & " | " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ": " & clientFilter
        End If
        
        ' Hebrew: "hashvaat" = comparison of
        Dim hashvaatText As String
        hashvaatText = ChrW(1492) & ChrW(1513) & ChrW(1493) & ChrW(1493) & ChrW(1488) & ChrW(1514) & " "

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
2740    BuildSummarySheet countRef, countCurrent, reviewCount, corrCount, ignoreCount, transferredCount, yearVal, refYear, periodDesc, hashvaatText & SHEET_SUMMARY() & " | " & titlePrefix
        ThisWorkbook.Worksheets(SHEET_SUMMARY()).Tab.Color = RGB(255, 204, 153)  ' pastel orange

CLEANUP:
2750    Application.ScreenUpdating = prevScreenUpdating
2760    Application.DisplayAlerts = prevDisplayAlerts
2770    Application.EnableEvents = prevEnableEvents
2780    Application.Calculation = prevCalculation

        ' Clear processing message from Main sheet and restore green background
2785    On Error Resume Next
        With ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("B12")
            .Value = ""
            .Interior.Color = RGB(220, 240, 220)
        End With
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
            ElseIf Left$(wsHide.Name, 5) = ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505) & "_" Then
                wsHide.Visible = xlSheetVeryHidden
            ElseIf wsHide.Name = ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1493) & ChrW(1514) Then
                wsHide.Visible = xlSheetVeryHidden
            ElseIf Left$(wsHide.Name, 6) = "__src_" Then
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
        With ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("B12")
            .Value = ""
            .Interior.Color = RGB(220, 240, 220)
        End With
2955    Application.StatusBar = False

2960    MsgBoxU wsMsgSrc2.Cells(5, 19).Value & errLine & vbCrLf & wsMsgSrc2.Cells(6, 19).Value & errNum & vbCrLf & wsMsgSrc2.Cells(7, 19).Value & debugStep & vbCrLf & errSrc & vbCrLf & errDesc, vbCritical

End Sub



' ============================================================================
' HELPER (7.34): Build base sheet from a source worksheet using Variant arrays
'
' Replaces the two near-duplicate ~150-line loops in v7.33 ApplyCorrections.
'
' Behaviour identical to v7.33:
'   - Skips rows in dictIgnore
'   - Drops rows where |premium| > threshold AND no correction exists
'   - Currency conversion when RAW_CURRENCY = 1 (USD) using dollarRate
'   - Maps RAW_BRANCHNAME -> dictBranch for MainBranch column
'   - Applies reason-based fixes from dictCorrections via ApplyRowFixes
'
' Returns the worksheet object (caller assigns to wsBase / wsBaseRef).
' ============================================================================
Private Function BuildBaseSheet(ByVal wsSrc As Worksheet, ByVal baseSheetName As String, _
        ByVal yearVal As String, ByVal dateCol As Long, _
        ByVal dollarRate As Double, ByVal threshold As Double, _
        ByVal dictBranch As Object, ByVal dictCorrections As Object, _
        ByVal dictIgnore As Object, ByRef countOut As Long) As Worksheet

10      On Error GoTo ERR_HANDLER

20      DeleteSheetIfExists baseSheetName
30      Dim wsBase As Worksheet
40      Set wsBase = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
50      wsBase.Name = baseSheetName
60      WriteBaseHeaders wsBase

        Dim lastRow As Long
70      lastRow = wsSrc.Cells(wsSrc.Rows.count, 2).End(xlUp).Row
80      If lastRow < 2 Then
90          countOut = 0
100         wsBase.Columns.AutoFit
110         Set BuildBaseSheet = wsBase
120         Exit Function
130     End If

        ' --- Bulk read source: 50 cols covers all RAW_* indexes (max RAW_IDNUMBER=45) ---
        Dim srcArr As Variant
140     srcArr = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, 50)).Value2
        Dim nRows As Long
150     nRows = UBound(srcArr, 1)

        ' --- Allocate output (over-allocated; Excel only writes outRow rows) ---
        Dim outArr() As Variant
160     ReDim outArr(1 To nRows, 1 To BASE_COL_TOFIX)

        Dim r As Long, outRow As Long, srcRow As Long
        Dim premVal As Double, commVal As Double
        Dim currCode As Variant
        Dim monthVal As Long, bordereu As Variant
        Dim dtStr As String, mPart As String
        Dim brKey As String

170     outRow = 0
180     For r = 1 To nRows
190         srcRow = r + 1   ' source sheet row (skipping header row 1)

200         If dictIgnore.Exists(CStr(srcRow)) Then GoTo NextRow

            ' --- Premium with currency + threshold ---
210         premVal = 0
220         commVal = 0
230         currCode = srcArr(r, RAW_CURRENCY)
240         If Not IsBlankValue(srcArr(r, RAW_PREMIUM)) Then
250             If TryParseVariantNumber(srcArr(r, RAW_PREMIUM), premVal) Then
260                 If IsNumeric(currCode) Then
270                     If CLng(currCode) = 1 Then premVal = premVal * dollarRate
280                 End If
290                 If Abs(premVal) > threshold And Not dictCorrections.Exists(CStr(srcRow)) Then GoTo NextRow
300             End If
310         End If

            ' --- Commission with currency conversion ---
320         If Not IsBlankValue(srcArr(r, RAW_COMMISSION)) Then
330             If TryParseVariantNumber(srcArr(r, RAW_COMMISSION), commVal) Then
340                 If IsNumeric(currCode) Then
350                     If CLng(currCode) = 1 Then commVal = commVal * dollarRate
360                 End If
370             End If
380         End If

            ' --- Month from date column ---
390         monthVal = 0
400         bordereu = srcArr(r, dateCol)
410         If IsDate(bordereu) Then
420             monthVal = Month(CDate(bordereu))
430         ElseIf IsNumeric(bordereu) Then
440             If CDbl(bordereu) > 1 Then monthVal = Month(CDate(CDbl(bordereu)))
450         ElseIf Not IsBlankValue(bordereu) Then
460             dtStr = CStr(bordereu)
470             If Len(dtStr) >= 7 Then
480                 mPart = Mid$(dtStr, 6, 2)
490                 If IsNumeric(mPart) Then monthVal = CInt(mPart)
500             End If
510         End If

            ' --- Write to output array (in-memory; no COM calls) ---
520         outRow = outRow + 1
530         outArr(outRow, BASE_COL_ID) = srcRow
540         outArr(outRow, BASE_COL_YEAR) = yearVal
550         outArr(outRow, BASE_COL_MONTH) = monthVal
560         outArr(outRow, BASE_COL_IDENTITY) = srcArr(r, RAW_IDNUMBER)
570         outArr(outRow, BASE_COL_CUSTOMER) = srcArr(r, RAW_CUSTOMER)
580         outArr(outRow, BASE_COL_CUSTNAME) = srcArr(r, RAW_CUSTNAME)
590         outArr(outRow, BASE_COL_POLICY) = srcArr(r, RAW_POLICY)
600         outArr(outRow, BASE_COL_ADDENDUM) = srcArr(r, RAW_ADDENDUM)
610         outArr(outRow, BASE_COL_COMPANY) = srcArr(r, RAW_COMPANY)
620         outArr(outRow, BASE_COL_COMPNUM) = srcArr(r, RAW_COMPNUM)
630         outArr(outRow, BASE_COL_BRANCHNAME) = srcArr(r, RAW_BRANCHNAME)
640         outArr(outRow, BASE_COL_BRANCHNUM) = srcArr(r, RAW_BRANCHNUM)

            ' Main branch mapping
650         brKey = UCase$(Trim$(CStr(srcArr(r, RAW_BRANCHNAME))))
660         If dictBranch.Exists(brKey) Then
670             outArr(outRow, BASE_COL_MAINBRANCH) = dictBranch(brKey)
680         Else
690             outArr(outRow, BASE_COL_MAINBRANCH) = srcArr(r, RAW_BRANCHNAME)
700         End If

710         outArr(outRow, BASE_COL_AGENTNAME) = srcArr(r, RAW_AGENTNAME)
720         outArr(outRow, BASE_COL_AGENTNUM) = srcArr(r, RAW_AGENTNUM)
730         outArr(outRow, BASE_COL_TELLER) = srcArr(r, RAW_TELLERNAME)
740         outArr(outRow, BASE_COL_TELLERNUM) = srcArr(r, RAW_TELLERNUM)
750         outArr(outRow, BASE_COL_ACTION) = srcArr(r, RAW_ACTIONCOL)
760         outArr(outRow, BASE_COL_PREMIUM) = premVal
770         outArr(outRow, BASE_COL_COMMISSION) = commVal

            ' Apply reason-based corrections
780         If dictCorrections.Exists(CStr(srcRow)) Then
790             outArr(outRow, BASE_COL_ISSUE) = "CORRECTED"
800             ApplyRowFixes outArr, outRow, dictCorrections(CStr(srcRow))
810         End If

NextRow:
820     Next r

        ' --- Bulk write output (Excel reads only outRow rows from over-allocated array) ---
830     If outRow > 0 Then
840         wsBase.Range(wsBase.Cells(2, 1), wsBase.Cells(outRow + 1, BASE_COL_TOFIX)).Value = outArr
850     End If

860     wsBase.Columns.AutoFit
870     countOut = outRow
880     Set BuildBaseSheet = wsBase
890     Exit Function

ERR_HANDLER:
900     Err.Raise Err.Number, "BuildBaseSheet:" & Erl, Err.Description
End Function


' ============================================================================
' HELPER (7.34): Apply reason-based fixes to a row in the output array
' Same Hebrew-substring matching as v7.33 (no behavioural change).
' ============================================================================
Private Sub ApplyRowFixes(ByRef outArr As Variant, ByVal outRow As Long, ByVal dictRowFixes As Object)
        Dim fixKey As Variant
        Dim reasonText As String, oneFix As String
        Dim parsedNum As Double

10      For Each fixKey In dictRowFixes.keys
20          reasonText = CStr(fixKey)
30          oneFix = Trim$(CStr(dictRowFixes(fixKey)))
40          If oneFix = "" Then GoTo NextFix

            ' Map Hebrew reason text to base column (matches v7.33 behaviour exactly)
50          If InStr(1, reasonText, ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503), vbTextCompare) > 0 Then
60              outArr(outRow, BASE_COL_AGENTNAME) = oneFix
70          ElseIf InStr(1, reasonText, ChrW(1496) & ChrW(1500) & ChrW(1512), vbTextCompare) > 0 Then
80              outArr(outRow, BASE_COL_TELLER) = oneFix
90          ElseIf InStr(1, reasonText, ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492), vbTextCompare) > 0 Then
100             outArr(outRow, BASE_COL_COMPANY) = oneFix
110         ElseIf InStr(1, reasonText, ChrW(1506) & ChrW(1504) & ChrW(1507), vbTextCompare) > 0 Then
120             outArr(outRow, BASE_COL_BRANCHNAME) = oneFix
130         ElseIf InStr(1, reasonText, ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1492), vbTextCompare) > 0 Then
140             If TryParseVariantNumber(oneFix, parsedNum) Then outArr(outRow, BASE_COL_PREMIUM) = parsedNum
150         ElseIf InStr(1, reasonText, ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1514), vbTextCompare) > 0 Then
160             If TryParseVariantNumber(oneFix, parsedNum) Then outArr(outRow, BASE_COL_COMMISSION) = parsedNum
170         End If
NextFix:
180     Next fixKey
End Sub


' ============================================================================
' HELPER (7.34): Write the 22-column base sheet header row
' Centralised so v7.33's two duplicate inline blocks become one definition.
' ============================================================================
Private Sub WriteBaseHeaders(ByVal wsBase As Worksheet)
10      wsBase.Cells(1, BASE_COL_ID).Value = "ID"
20      wsBase.Cells(1, BASE_COL_YEAR).Value = "Year"
30      wsBase.Cells(1, BASE_COL_MONTH).Value = "Month"
40      wsBase.Cells(1, BASE_COL_IDENTITY).Value = "Identity"
50      wsBase.Cells(1, BASE_COL_CUSTOMER).Value = "Customer"
60      wsBase.Cells(1, BASE_COL_CUSTNAME).Value = "CustName"
70      wsBase.Cells(1, BASE_COL_POLICY).Value = "Policy"
80      wsBase.Cells(1, BASE_COL_ADDENDUM).Value = "Addendum"
90      wsBase.Cells(1, BASE_COL_COMPANY).Value = "Company"
100     wsBase.Cells(1, BASE_COL_COMPNUM).Value = "CompNum"
110     wsBase.Cells(1, BASE_COL_BRANCHNAME).Value = "BranchName"
120     wsBase.Cells(1, BASE_COL_BRANCHNUM).Value = "BranchNum"
130     wsBase.Cells(1, BASE_COL_MAINBRANCH).Value = "MainBranch"
140     wsBase.Cells(1, BASE_COL_AGENTNAME).Value = "AgentName"
150     wsBase.Cells(1, BASE_COL_AGENTNUM).Value = "AgentNum"
160     wsBase.Cells(1, BASE_COL_TELLER).Value = "Teller"
170     wsBase.Cells(1, BASE_COL_TELLERNUM).Value = "TellerNum"
180     wsBase.Cells(1, BASE_COL_ACTION).Value = "Action"
190     wsBase.Cells(1, BASE_COL_PREMIUM).Value = "Premium"
200     wsBase.Cells(1, BASE_COL_COMMISSION).Value = "Commission"
210     wsBase.Cells(1, BASE_COL_ISSUE).Value = "Issue"
220     wsBase.Cells(1, BASE_COL_TOFIX).Value = "ToFix"
End Sub


' ============================================================================
' HELPER: Build comparison sheet (companies, branch, tellers, agents, months)
' ============================================================================
Private Sub BuildComparisonSheet(ByVal wsCurrent As Worksheet, ByVal wsRef As Worksheet, ByVal sheetName As String, ByVal groupCol As Long, ByVal minMonth As Long, ByVal maxMonth As Long, ByVal yearVal As String, ByVal refYear As String, Optional ByVal filterCol As Long = 0, Optional ByVal filterVal As String = "", Optional ByVal titleText As String = "", Optional ByVal clientFilterVal As String = "")

10      On Error GoTo ERR_HANDLER

20      DeleteSheetIfExists sheetName
30      Dim wsOut As Worksheet
40      Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
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


140     lastRowCur = wsCurrent.Cells(wsCurrent.Rows.count, 1).End(xlUp).Row
150     lastRowRef = wsRef.Cells(wsRef.Rows.count, 1).End(xlUp).Row

160     For r = 2 To lastRowCur
170         m = 0
180         If Not IsBlankValue(wsCurrent.Cells(r, BASE_COL_MONTH).Value2) Then
190             m = CLng(wsCurrent.Cells(r, BASE_COL_MONTH).Value2)
200         End If
210         If m >= minMonth And m <= maxMonth Then
                ' Exclude MainBranch = "????" from all reports
211             If StrComp(Trim$(CStr(wsCurrent.Cells(r, BASE_COL_MAINBRANCH).Value2)), ChrW(1495) & ChrW(1493) & ChrW(1489) & ChrW(1492), vbTextCompare) = 0 Then GoTo SkipCurKey
                ' Cross-filter: if filterCol is set, only include rows matching filterVal
215             If filterCol > 0 And filterVal <> "" Then
216                 If StrComp(Trim$(CStr(wsCurrent.Cells(r, filterCol).Value2)), filterVal, vbTextCompare) <> 0 Then GoTo SkipCurKey
217             End If
                ' Client filter: if clientFilterVal is set, only include rows matching client name
218             If clientFilterVal <> "" Then
219                 If StrComp(Trim$(CStr(wsCurrent.Cells(r, BASE_COL_CUSTNAME).Value2)), clientFilterVal, vbTextCompare) <> 0 Then GoTo SkipCurKey
220             End If
221             k = Trim$(CStr(wsCurrent.Cells(r, groupCol).Value2))
230             If k <> "" And LCase$(k) <> "(empty)" Then
240                 If Not dictKeys.Exists(k) Then dictKeys(k) = True
250             End If
260         End If
SkipCurKey:
270     Next r

280     For r = 2 To lastRowRef
290         m = 0
300         If Not IsBlankValue(wsRef.Cells(r, BASE_COL_MONTH).Value2) Then
310             m = CLng(wsRef.Cells(r, BASE_COL_MONTH).Value2)
320         End If
330         If m >= minMonth And m <= maxMonth Then
                ' Exclude MainBranch = "????" from all reports
331             If StrComp(Trim$(CStr(wsRef.Cells(r, BASE_COL_MAINBRANCH).Value2)), ChrW(1495) & ChrW(1493) & ChrW(1489) & ChrW(1492), vbTextCompare) = 0 Then GoTo SkipRefKey
                ' Cross-filter: if filterCol is set, only include rows matching filterVal
335             If filterCol > 0 And filterVal <> "" Then
336                 If StrComp(Trim$(CStr(wsRef.Cells(r, filterCol).Value2)), filterVal, vbTextCompare) <> 0 Then GoTo SkipRefKey
337             End If
                ' Client filter: if clientFilterVal is set, only include rows matching client name
338             If clientFilterVal <> "" Then
339                 If StrComp(Trim$(CStr(wsRef.Cells(r, BASE_COL_CUSTNAME).Value2)), clientFilterVal, vbTextCompare) <> 0 Then GoTo SkipRefKey
340             End If
341             k = Trim$(CStr(wsRef.Cells(r, groupCol).Value2))
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
830             If Not IsBlankValue(wsCurrent.Cells(r, BASE_COL_MONTH).Value2) Then
840                 m = CLng(wsCurrent.Cells(r, BASE_COL_MONTH).Value2)
850             End If
860             If m < minMonth Or m > maxMonth Then GoTo NextCurRow
                ' Exclude MainBranch = "????" from all reports
861             If StrComp(Trim$(CStr(wsCurrent.Cells(r, BASE_COL_MAINBRANCH).Value2)), ChrW(1495) & ChrW(1493) & ChrW(1489) & ChrW(1492), vbTextCompare) = 0 Then GoTo NextCurRow
                ' Cross-filter: skip rows not matching the filter
862             If filterCol > 0 And filterVal <> "" Then
864                 If StrComp(Trim$(CStr(wsCurrent.Cells(r, filterCol).Value2)), filterVal, vbTextCompare) <> 0 Then GoTo NextCurRow
866             End If
                ' Client filter
867             If clientFilterVal <> "" Then
868                 If StrComp(Trim$(CStr(wsCurrent.Cells(r, BASE_COL_CUSTNAME).Value2)), clientFilterVal, vbTextCompare) <> 0 Then GoTo NextCurRow
869             End If

870             If StrComp(Trim$(CStr(wsCurrent.Cells(r, groupCol).Value2)), k, vbTextCompare) = 0 Then
880                 premCur = premCur + CDbl(wsCurrent.Cells(r, BASE_COL_PREMIUM).Value2)
890                 commCur = commCur + CDbl(wsCurrent.Cells(r, BASE_COL_COMMISSION).Value2)
900                 docsCur = docsCur + 1

                    ' Customer counting: unique customer numbers with at least one non-cancelled row
910                 custKey = Trim$(CStr(wsCurrent.Cells(r, BASE_COL_CUSTOMER).Value2))
920                 If custKey <> "" Then
                        actValCur = Trim$(CStr(wsCurrent.Cells(r, BASE_COL_ACTION).Value2))
                        ' Check if this row is NOT a cancellation
                        If InStr(1, actValCur, ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1500), vbTextCompare) = 0 Then
930                         If Not dictCustCur.Exists(custKey) Then dictCustCur(custKey) = True
                        End If
940                 End If

950                 polKey = Trim$(CStr(wsCurrent.Cells(r, BASE_COL_POLICY).Value2))
960                 If polKey <> "" Then
970                     If Not dictPolCur.Exists(polKey) Then dictPolCur(polKey) = True
980                 End If
990             End If
NextCurRow:
1000        Next r

            ' Aggregate reference year
1010        For r = 2 To lastRowRef
1020            m = 0
1030            If Not IsBlankValue(wsRef.Cells(r, BASE_COL_MONTH).Value2) Then
1040                m = CLng(wsRef.Cells(r, BASE_COL_MONTH).Value2)
1050            End If
1060            If m < minMonth Or m > maxMonth Then GoTo NextRefRow2
                ' Exclude MainBranch = "????" from all reports
1061            If StrComp(Trim$(CStr(wsRef.Cells(r, BASE_COL_MAINBRANCH).Value2)), ChrW(1495) & ChrW(1493) & ChrW(1489) & ChrW(1492), vbTextCompare) = 0 Then GoTo NextRefRow2
                ' Cross-filter: skip rows not matching the filter
1062            If filterCol > 0 And filterVal <> "" Then
1064                If StrComp(Trim$(CStr(wsRef.Cells(r, filterCol).Value2)), filterVal, vbTextCompare) <> 0 Then GoTo NextRefRow2
1066            End If
                ' Client filter
1067            If clientFilterVal <> "" Then
1068                If StrComp(Trim$(CStr(wsRef.Cells(r, BASE_COL_CUSTNAME).Value2)), clientFilterVal, vbTextCompare) <> 0 Then GoTo NextRefRow2
1069            End If

1070            If StrComp(Trim$(CStr(wsRef.Cells(r, groupCol).Value2)), k, vbTextCompare) = 0 Then
1080                premRef = premRef + CDbl(wsRef.Cells(r, BASE_COL_PREMIUM).Value2)
1090                commRef = commRef + CDbl(wsRef.Cells(r, BASE_COL_COMMISSION).Value2)
1100                docsRef = docsRef + 1

                    ' Customer counting: unique customer numbers with at least one non-cancelled row
1110                custKey = Trim$(CStr(wsRef.Cells(r, BASE_COL_CUSTOMER).Value2))
1120                If custKey <> "" Then
                        actValRef = Trim$(CStr(wsRef.Cells(r, BASE_COL_ACTION).Value2))
                        ' Check if this row is NOT a cancellation
                        If InStr(1, actValRef, ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1500), vbTextCompare) = 0 Then
1130                        If Not dictCustRef.Exists(custKey) Then dictCustRef(custKey) = True
                        End If
1140                End If

1150                polKey = Trim$(CStr(wsRef.Cells(r, BASE_COL_POLICY).Value2))
1160                If polKey <> "" Then
1170                    If Not dictPolRef.Exists(polKey) Then dictPolRef(polKey) = True
1180                End If
1190            End If
NextRefRow2:
1200        Next r

1210        custCur = dictCustCur.count
1220        custRef = dictCustRef.count
1230        polCur = dictPolCur.count
1240        polRef = dictPolRef.count

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
        totCustCur = dictGlobalCustCur.count
        totCustRef = dictGlobalCustRef.count
        totPolCur = dictGlobalPolCur.count
        totPolRef = dictGlobalPolRef.count

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
500     Err.Raise Err.Number, "WriteComparisonHeaders:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Build summary sheet
' ============================================================================
Private Sub BuildSummarySheet(ByVal countRef As Long, ByVal countCurrent As Long, ByVal reviewCount As Long, ByVal corrCount As Long, ByVal ignoreCount As Long, ByVal transferredCount As Long, ByVal yearVal As String, ByVal refYear As String, ByVal periodDesc As String, Optional ByVal titleText As String = "")

10      On Error GoTo ERR_HANDLER
20      DeleteSheetIfExists SHEET_SUMMARY()
30      Dim wsOut As Worksheet
40      Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
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
340     wsOut.Cells(11, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1500) & ChrW(1488) & " " & ChrW(1510) & ChrW(1493) & ChrW(1512) & ChrW(1498) & " " & ChrW(1489) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500)
350     wsOut.Cells(11, 2).Value = ignoreCount
        ' v7.60: row 12 was "shurot lelo tipul" (unhandledCount), now "shurot shehu'avru livdika"
360     wsOut.Cells(12, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1513) & ChrW(1492) & ChrW(1493) & ChrW(1506) & ChrW(1489) & ChrW(1512) & ChrW(1493) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1492)
370     wsOut.Cells(12, 2).Value = transferredCount

380     wsOut.Columns(2).NumberFormat = "#,##0"
390     wsOut.Columns.AutoFit
        wsOut.DisplayRightToLeft = True
410     Exit Sub

ERR_HANDLER:
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
' HELPER: Load helper dictionary from NIHUL columns N-O
' ============================================================================
Private Function LoadHelperDictionary(ByVal ws As Worksheet) As Object
10      On Error GoTo ERR_HANDLER
20      Dim dict As Object
30      Set dict = CreateObject("Scripting.Dictionary")
40      dict.CompareMode = vbTextCompare
50      Dim r As Long
60      Dim lastRow As Long
70      Dim k As String
80      lastRow = ws.Cells(ws.Rows.count, COL_HELPER_KEY).End(xlUp).Row
90      For r = 1 To lastRow
100         k = Trim$(CStr(ws.Cells(r, COL_HELPER_KEY).Value2))
110         If k <> "" Then
120             dict(k) = Trim$(CStr(ws.Cells(r, COL_HELPER_VALUE).Value2))
130         End If
140     Next r
        ' Add fallback translations for missing codes
150     If Not dict.Exists("MISSING_CUSTOMER_NUMBER") Then dict("MISSING_CUSTOMER_NUMBER") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)
151     If Not dict.Exists("MISSING_CUSTOMER_NAME") Then dict("MISSING_CUSTOMER_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)
152     If Not dict.Exists("MISSING_POLICY") Then dict("MISSING_POLICY") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1505) & ChrW(1492)
153     If Not dict.Exists("MISSING_ADDENDUM") Then dict("MISSING_ADDENDUM") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1514) & ChrW(1493) & ChrW(1505) & ChrW(1508) & ChrW(1514)
154     If Not dict.Exists("MISSING_COMPANY_NAME") Then dict("MISSING_COMPANY_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492)
155     If Not dict.Exists("MISSING_BRANCH_NAME") Then dict("MISSING_BRANCH_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1506) & ChrW(1504) & ChrW(1507)
156     If Not dict.Exists("MISSING_AGENT_NAME") Then dict("MISSING_AGENT_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503)
157     If Not dict.Exists("MISSING_UNDERWRITER_TELLER_NAME") Then dict("MISSING_UNDERWRITER_TELLER_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1496) & ChrW(1500) & ChrW(1512)
158     If Not dict.Exists("MISSING_CURRENCY") Then dict("MISSING_CURRENCY") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1502) & ChrW(1496) & ChrW(1489) & ChrW(1506)
159     If Not dict.Exists("MISSING_PREMIUM") Then dict("MISSING_PREMIUM") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1492)
160     If Not dict.Exists("MISSING_COMPANY_COMMISSION") Then dict("MISSING_COMPANY_COMMISSION") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1514) & " " & ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492)
161     If Not dict.Exists("PREMIUM_OVER_THRESHOLD") Then dict("PREMIUM_OVER_THRESHOLD") = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1492) & " " & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1492)
162     If Not dict.Exists("PREMIUM_NOT_NUMERIC") Then dict("PREMIUM_NOT_NUMERIC") = ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1492) & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & ChrW(1497)

170     Set LoadHelperDictionary = dict
180     Exit Function
ERR_HANDLER:
190     Err.Raise Err.Number, "LoadHelperDictionary:" & Erl, Err.Description
End Function


' ============================================================================
' HELPER: Load branch mapping from NIHUL
' ============================================================================
Private Function LoadBranchMapping(ByVal ws As Worksheet) As Object
10      On Error GoTo ERR_HANDLER
20      Dim dict As Object
30      Set dict = CreateObject("Scripting.Dictionary")
40      dict.CompareMode = vbTextCompare
50      Dim r As Long
60      Dim lastRow As Long
70      Dim brName As String
80      Dim mainBr As String
90      lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
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
20      Dim r As Long
30      Dim lastRow As Long
40      Dim fName As String
50      Dim fCol As String
60      Dim fCheck As String
70      Dim fKey As String
80      lastRow = ws.Cells(ws.Rows.count, COL_FIELD_NAME_HE).End(xlUp).Row
90      For r = MANAGEMENT_START_ROW To lastRow
100         fName = Trim$(CStr(ws.Cells(r, COL_FIELD_NAME_HE).Value2))
110         fCol = UCase$(Trim$(CStr(ws.Cells(r, COL_FIELD_COLUMN).Value2)))
120         fCheck = UCase$(Trim$(CStr(ws.Cells(r, COL_FIELD_CHECKING).Value2)))
130         fKey = Trim$(CStr(ws.Cells(r, COL_FIELD_KEY).Value2))
140         If fKey <> "" And fCol <> "" And fCheck = "CHECK" Then
150             dictCol(fKey) = ColumnLetterToNumber(fCol)
160             dictDisp(fKey) = fName
170         End If
180     Next r
190     Exit Sub
ERR_HANDLER:
200     Err.Raise Err.Number, "LoadCheckedFields:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Build arrays from dictionaries
' ============================================================================
Private Sub BuildArrays(ByVal dictCol As Object, ByVal dictDisp As Object, ByRef keys() As String, ByRef cols() As Long, ByRef disp() As String, ByRef cnt As Long)
10      On Error GoTo ERR_HANDLER
20      cnt = dictCol.count
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
170     Err.Raise Err.Number, "BuildArrays:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Get string parameter from NIHUL
' ============================================================================
Private Function GetStringParameter(ByVal ws As Worksheet, ByVal paramName As String) As String
10      On Error GoTo ERR_HANDLER
20      Dim r As Long
30      Dim lastRow As Long
40      Dim nm As String
50      lastRow = ws.Cells(ws.Rows.count, COL_PARAM_NAME).End(xlUp).Row
60      For r = 1 To lastRow
70          nm = UCase$(Trim$(CStr(ws.Cells(r, COL_PARAM_NAME).Value2)))
80          If nm = UCase$(paramName) Then
90              GetStringParameter = Trim$(CStr(ws.Cells(r, COL_PARAM_VALUE).Value2))
100             Exit Function
110         End If
120     Next r
130     GetStringParameter = ""
140     Exit Function
ERR_HANDLER:
150     GetStringParameter = ""
End Function


' ============================================================================
' HELPER: Get numeric parameter from NIHUL
' ============================================================================
Private Function GetNumericParameter(ByVal ws As Worksheet, ByVal paramName As String) As Double
10      On Error GoTo ERR_HANDLER
20      Dim r As Long
30      Dim lastRow As Long
40      Dim nm As String
50      Dim v As Variant
60      Dim n As Double
70      lastRow = ws.Cells(ws.Rows.count, COL_PARAM_NAME).End(xlUp).Row
80      For r = 1 To lastRow
90          nm = UCase$(Trim$(CStr(ws.Cells(r, COL_PARAM_NAME).Value2)))
100         If nm = UCase$(paramName) Then
110             v = ws.Cells(r, COL_PARAM_VALUE).Value2
120             If TryParseVariantNumber(v, n) Then
130                 GetNumericParameter = n
140             Else
150                 Err.Raise vbObjectError + 3000, "GetNumericParameter", "PARAMETER NOT NUMERIC: " & paramName
160             End If
170             Exit Function
180         End If
190     Next r
200     Err.Raise vbObjectError + 3001, "GetNumericParameter", "PARAMETER NOT FOUND: " & paramName
ERR_HANDLER:
210     Err.Raise Err.Number, "GetNumericParameter:" & Erl, Err.Description
End Function


' ============================================================================
' HELPER: Column letter to number
' ============================================================================
Private Function ColumnLetterToNumber(ByVal col As String) As Long
10      On Error GoTo ERR_HANDLER
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
120     Err.Raise Err.Number, "ColumnLetterToNumber:" & Erl, Err.Description
End Function


' ============================================================================
' HELPER: Delete sheets
' ============================================================================
' --- 7.33: removed dead procedure DeleteReviewSheetIfExists ---

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
    sSuffix = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492) & " " & ChrW(1500) & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495) & " " & ChrW(1489) & ChrW(1506) & ChrW(34) & ChrW(1502)
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
' HELPER: Fetch latest exchange rates from Bank of Israel API
' Returns True if successful, and populates outUSD and outEUR
' ============================================================================
Private Function FetchBOIRates(ByRef outUSD As Double, ByRef outEUR As Double) As Boolean
    On Error GoTo FAIL
    Dim xmlHttp As Object
    Dim xmlDoc As Object
    Dim nodes As Object
    Dim node As Object
    Dim nodeKey As String
    Dim i As Long
    Dim found As Long
    
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP")
    ' Set timeouts: resolve=5s, connect=5s, send=5s, receive=5s
    xmlHttp.setTimeouts 5000, 5000, 5000, 5000
    xmlHttp.Open "GET", "https://boi.org.il/PublicApi/GetExchangeRates?asXML=true", False
    xmlHttp.send
    
    If xmlHttp.Status <> 200 Then GoTo FAIL
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.LoadXML xmlHttp.responseText
    xmlDoc.SetProperty "SelectionNamespaces", "xmlns:d='http://schemas.datacontract.org/2004/07/BOI.Core.Models.HotData'"
    
    Set nodes = xmlDoc.SelectNodes("//d:ExchangeRateResponseDTO")
    found = 0
    For i = 0 To nodes.Length - 1
        Set node = nodes.Item(i)
        nodeKey = UCase$(node.SelectSingleNode("d:Key").Text)
        If nodeKey = "USD" Then
            outUSD = CDbl(node.SelectSingleNode("d:CurrentExchangeRate").Text)
            found = found + 1
        ElseIf nodeKey = "EUR" Then
            outEUR = CDbl(node.SelectSingleNode("d:CurrentExchangeRate").Text)
            found = found + 1
        End If
        If found = 2 Then Exit For
    Next i
    
    FetchBOIRates = (found > 0)
    Exit Function
FAIL:
    FetchBOIRates = False
End Function

' ============================================================================
' HELPER: Get dollar exchange rate - BOI API first, fallback to K3 cell
' Writes fetched rates (USD and EUR) back to home sheet for user visibility
' ============================================================================
Private Function GetDollarRate(ByVal wsMain As Worksheet) As Double
    Dim rateUSD As Double, rateEUR As Double
    If FetchBOIRates(rateUSD, rateEUR) Then
        ' Write back to home sheet for user visibility
        On Error Resume Next
        If rateUSD > 0 Then wsMain.Range("K3").Value = rateUSD
        If rateEUR > 0 Then wsMain.Range("K4").Value = rateEUR
        On Error GoTo 0
        GetDollarRate = rateUSD
    Else
        ' Fallback: read from cell
        GetDollarRate = 1
        On Error Resume Next
        If Not IsBlankValue(wsMain.Range("K3").Value2) Then
            If IsNumeric(wsMain.Range("K3").Value2) Then GetDollarRate = CDbl(wsMain.Range("K3").Value2)
        End If
        On Error GoTo 0
    End If
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
55          result = 0          ' 7.33: deterministic ByRef value
60          TryParseVariantNumber = False
70      End If
80      Exit Function
FAIL:
85      result = 0              ' 7.33: deterministic ByRef value
90      TryParseVariantNumber = False
End Function


' --- 7.33: removed dead function AddReason ---

' --- 7.33: removed dead function TranslateReason ---

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
Public Sub SetupMainSheet()

10      Dim wsMain As Worksheet
20      Dim wsMgmt As Worksheet
30      Dim shp As Shape
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
        Dim dateTypeList As String   ' 7.33: was implicit Variant

40      On Error GoTo ERR_HANDLER

        ' Try to find Main sheet by current Hebrew name or old English name
50      On Error Resume Next
51      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
52      If wsMain Is Nothing Then Set wsMain = ThisWorkbook.Worksheets("Main")
53      On Error GoTo ERR_HANDLER
        ' If not found, create it as the first sheet
54      If wsMain Is Nothing Then
            Set wsMain = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
            wsMain.Name = CONTROL_SHEET_NAME()
        End If
        ' Rename to Hebrew if still English
55      If wsMain.Name <> CONTROL_SHEET_NAME() Then wsMain.Name = CONTROL_SHEET_NAME()
        ' Try to find NIHUL/hagdarot sheet by current Hebrew name or old English name
56      On Error Resume Next
57      Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
58      If wsMgmt Is Nothing Then Set wsMgmt = ThisWorkbook.Worksheets("NIHUL")
59      On Error GoTo ERR_HANDLER
60      If wsMgmt Is Nothing Then Err.Raise vbObjectError + 9002, "SetupMainSheet", "Cannot find NIHUL/hagdarot sheet"
        ' Rename to Hebrew if still English
62      If wsMgmt.Name <> MANAGEMENT_SHEET_NAME() Then wsMgmt.Name = MANAGEMENT_SHEET_NAME()
70      blueClr = RGB(0, 70, 140)
        ' ---- Layout Spacing for Physical Centering ----
        ' Instead of relying purely on Zoom (which leaves a massive empty left side on wide screens),
        ' we physically spread out the elements.
        wsMain.Columns("A").ColumnWidth = 50     ' Right margin (User requested 50)
        wsMain.Columns("B").ColumnWidth = 15     ' Buttons (User requested)
        wsMain.Columns("C:D").ColumnWidth = 5    ' Spacer
        wsMain.Columns("E").ColumnWidth = 13     ' User requested 13
        wsMain.Columns("F:G").ColumnWidth = 16   ' Parameter Table (reverted to F:G)
        wsMain.Columns("H:I").ColumnWidth = 13   ' User requested 13
        wsMain.Columns("J:K").ColumnWidth = 12   ' Exchange Rate box (reverted to J:K)
        wsMain.Columns("L:M").ColumnWidth = 5    ' Spacer
        wsMain.Columns("N:O").ColumnWidth = 12   ' Spacer
        wsMain.Columns("P:T").ColumnWidth = 5    ' Left margins
        
        ' ---- Vertical Centering: Push everything down by making top rows even taller ----
        wsMain.Rows("1:3").RowHeight = 45

        ' ---- CLEANUP OF PREVIOUS VERSIONS ARTIFACTS ----
        ' Versions 7.76 and 7.78 incorrectly drew borders and text in columns H:I and L:N.
        ' This cleans up the duplicates so the grid looks clean like 7.75 again.
        On Error Resume Next
        wsMain.Range("H2:I17").Clear
        wsMain.Range("L2:O17").Clear
        On Error GoTo ERR_HANDLER

        ' "shnat basis" = year of reference
80      lblShnBasis = ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505)
        ' "shna shoteft" = current year
90      lblShnShotef = ChrW(1513) & ChrW(1504) & ChrW(1492) & " " & ChrW(1513) & ChrW(1493) & ChrW(1496) & ChrW(1508) & ChrW(1514)
        ' "tkufa" = period
100     lblTkufa = ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
        ' "pirot tkufa" = period detail
110     lblPirutTkufa = ChrW(1508) & ChrW(1497) & ChrW(1512) & ChrW(1493) & ChrW(1496) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
        ' "sug taarih" = date type
120     lblSugTaarih = ChrW(1505) & ChrW(1493) & ChrW(1490) & " " & ChrW(1514) & ChrW(1488) & ChrW(1512) & ChrW(1497) & ChrW(1498)

        ' ---- Clear old labels from A2:D5 (no longer used) ----
130     wsMain.Range("A2:D5").ClearContents

        ' ---- Style row 2 headers (F2, G2, J2, K2) ----
        Dim hdrRng As Range
        Set hdrRng = Union(wsMain.Range("F2"), wsMain.Range("G2"), wsMain.Range("J2"), wsMain.Range("K2"))
135     hdrRng.Interior.Color = RGB(0, 100, 0)
136     hdrRng.Font.Color = RGB(255, 255, 255)
137     hdrRng.Font.Bold = True
138     hdrRng.Font.Size = 12
139     hdrRng.HorizontalAlignment = xlCenter

        ' ---- Write period lookup lists on NIHUL column Q ----
        ' Q1 header: "period_type"
250     wsMgmt.Cells(1, 17).Value = "PERIOD_TYPE"
        ' Q2-Q5: the four period types
        ' shnatit = yearly
260     wsMgmt.Cells(2, 17).Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
        ' chatzi shnati = half yearly
270     wsMgmt.Cells(3, 17).Value = ChrW(1495) & ChrW(1510) & ChrW(1497) & " " & ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
        ' riv'oni = quarterly
280     wsMgmt.Cells(4, 17).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497)
        ' chodshi = monthly
290     wsMgmt.Cells(5, 17).Value = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497)

        ' R1 header: "HALF_YEAR"
300     wsMgmt.Cells(1, 18).Value = "HALF_YEAR"
        ' R2: machatzit rishona = first half
310     wsMgmt.Cells(2, 18).Value = ChrW(1502) & ChrW(1495) & ChrW(1510) & ChrW(1497) & ChrW(1514) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1504) & ChrW(1492)
        ' R3: machatzit shniya = second half
320     wsMgmt.Cells(3, 18).Value = ChrW(1502) & ChrW(1495) & ChrW(1510) & ChrW(1497) & ChrW(1514) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497) & ChrW(1492)

        ' Define named ranges on NIHUL for dropdown sources
        ' Period types: NIHUL!Q2:Q5
330     On Error Resume Next
340     ThisWorkbook.names("lst_period_type").Delete
350     ThisWorkbook.names("lst_half_year").Delete
360     ThisWorkbook.names("lst_quarter").Delete
370     ThisWorkbook.names("lst_month").Delete
380     On Error GoTo ERR_HANDLER

390     ThisWorkbook.names.Add Name:="lst_period_type", RefersTo:="=" & MANAGEMENT_SHEET_NAME() & "!$Q$2:$Q$5"
400     ThisWorkbook.names.Add Name:="lst_half_year", RefersTo:="=" & MANAGEMENT_SHEET_NAME() & "!$R$2:$R$3"

        ' Quarter list in NIHUL column R rows 5-8
        ' riv'on rishon = Q1
410     wsMgmt.Cells(5, 18).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1503)
        ' riv'on sheni = Q2
420     wsMgmt.Cells(6, 18).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1504) & ChrW(1497)
        ' riv'on shlishi = Q3
430     wsMgmt.Cells(7, 18).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1513) & ChrW(1497)
        ' riv'on revi'i = Q4
440     wsMgmt.Cells(8, 18).Value = ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1503) & " " & ChrW(1512) & ChrW(1489) & ChrW(1497) & ChrW(1506) & ChrW(1497)
450     ThisWorkbook.names.Add Name:="lst_quarter", RefersTo:="=" & MANAGEMENT_SHEET_NAME() & "!$R$5:$R$8"

        ' Month list in NIHUL column R rows 10-21 (Hebrew month names)
460     wsMgmt.Cells(10, 18).Value = ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1488) & ChrW(1512)
470     wsMgmt.Cells(11, 18).Value = ChrW(1508) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1488) & ChrW(1512)
480     wsMgmt.Cells(12, 18).Value = ChrW(1502) & ChrW(1512) & ChrW(1509)
490     wsMgmt.Cells(13, 18).Value = ChrW(1488) & ChrW(1508) & ChrW(1512) & ChrW(1497) & ChrW(1500)
500     wsMgmt.Cells(14, 18).Value = ChrW(1502) & ChrW(1488) & ChrW(1497)
510     wsMgmt.Cells(15, 18).Value = ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1497)
520     wsMgmt.Cells(16, 18).Value = ChrW(1497) & ChrW(1493) & ChrW(1500) & ChrW(1497)
530     wsMgmt.Cells(17, 18).Value = ChrW(1488) & ChrW(1493) & ChrW(1490) & ChrW(1493) & ChrW(1505) & ChrW(1496)
540     wsMgmt.Cells(18, 18).Value = ChrW(1505) & ChrW(1508) & ChrW(1496) & ChrW(1502) & ChrW(1489) & ChrW(1512)
550     wsMgmt.Cells(19, 18).Value = ChrW(1488) & ChrW(1493) & ChrW(1511) & ChrW(1496) & ChrW(1493) & ChrW(1489) & ChrW(1512)
560     wsMgmt.Cells(20, 18).Value = ChrW(1504) & ChrW(1493) & ChrW(1489) & ChrW(1502) & ChrW(1489) & ChrW(1512)
570     wsMgmt.Cells(21, 18).Value = ChrW(1491) & ChrW(1510) & ChrW(1502) & ChrW(1489) & ChrW(1512)
580     ThisWorkbook.names.Add Name:="lst_month", RefersTo:="=" & MANAGEMENT_SHEET_NAME() & "!$R$10:$R$21"

        ' ---- rngPeriodType dropdown: period type ----
590     wsMain.Range("rngPeriodType").Validation.Delete
600     wsMain.Range("rngPeriodType").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=lst_period_type"

        ' ---- rngDateType dropdown: date type ----
        ' bordereu / thilat bituah
610     wsMain.Range("rngDateType").Validation.Delete

620     dateTypeList = ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493) & "," & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & ChrW(1514) & " " & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495)
630     wsMain.Range("rngDateType").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=dateTypeList

        ' ---- C4 dependent dropdown: uses INDIRECT based on B4 value ----
        ' We use a formula approach: map B4 value to named range
        ' INDIRECT formula: =IF(B4=shnatit,"",IF(B4=chatzi,lst_half_year,IF(B4=riv'oni,lst_quarter,IF(B4=chodshi,lst_month,""))))
        ' Since INDIRECT with dynamic named ranges is complex, we use Worksheet_Change event instead
        ' For now, set C4 validation to allow any list - it will be updated by the event macro
640     wsMain.Range("rngPeriodValue").Validation.Delete

        ' ---- Remove ALL old buttons ----
650     On Error Resume Next
        For Each s In wsMain.Shapes
660         s.Delete
670     Next s
680     On Error GoTo ERR_HANDLER

        ' ---- Button 1: BuildReview ----
        ' "1 - bdikat netunim" = Data Review
        ' ---- Pre-calculate button heights to match Clear button's bottom ----
        Dim btnSearchTop As Single
        Dim btnClearTop As Single
        Dim btnClearBottom As Single
        btnSearchTop = wsMain.Range("F12").Top + wsMain.Range("F12").Height + 5
        btnClearTop = btnSearchTop + 25 + 5
        btnClearBottom = btnClearTop + 25

        Dim startTop As Single
        Dim totalAvailHeight As Single
        Dim btnGap As Single
        Dim calcBtnHeight As Single
        Dim btnTop As Single
        
        startTop = wsMain.Range("B2").Top
        btnGap = 3
        totalAvailHeight = btnClearBottom - startTop
        calcBtnHeight = (totalAvailHeight - (5 * btnGap)) / 6
        btnTop = startTop

        ' ---- Button 1: BuildReview ----
        Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("B2").Left + 2, btnTop, 80, calcBtnHeight)
        btnTop = shp.Top + shp.Height + btnGap
        shp.Name = "btnBuildReview"
        shp.Fill.ForeColor.RGB = blueClr
        shp.TextFrame2.TextRange.Text = "1" & vbCrLf & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1514) & vbCrLf & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.TextFrame2.TextRange.Font.Size = 10
        shp.TextFrame2.TextRange.Font.Bold = msoTrue
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.OnAction = "BuildReview"

        ' ---- Button 2: ApplyCorrectionsAndBuildReports ----
        Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("B2").Left + 2, btnTop, 80, calcBtnHeight)
        btnTop = shp.Top + shp.Height + btnGap
        shp.Name = "btnApplyCorrections"
        shp.Fill.ForeColor.RGB = RGB(0, 120, 60)
        shp.TextFrame2.TextRange.Text = "2" & vbCrLf & ChrW(1497) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1501) & vbCrLf & ChrW(1493) & ChrW(1491) & ChrW(1493) & ChrW(34) & ChrW(1495) & ChrW(1493) & ChrW(1514)
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.TextFrame2.TextRange.Font.Size = 10
        shp.TextFrame2.TextRange.Font.Bold = msoTrue
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.OnAction = "ApplyCorrectionsAndBuildReports"

        ' ---- Button 3: BuildPresentation ----
        Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("B2").Left + 2, btnTop, 80, calcBtnHeight)
        btnTop = shp.Top + shp.Height + btnGap
        shp.Name = "btnBuildPresentation"
        shp.Fill.ForeColor.RGB = RGB(160, 80, 0)
        shp.TextFrame2.TextRange.Text = "3" & vbCrLf & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1493) & ChrW(1512) & vbCrLf & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514)
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.TextFrame2.TextRange.Font.Size = 10
        shp.TextFrame2.TextRange.Font.Bold = msoTrue
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.OnAction = "BuildPresentation"

        ' ---- Button 4: Save Reports (Pro Only) ----
        Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("B2").Left + 2, btnTop, 80, calcBtnHeight)
        btnTop = shp.Top + shp.Height + btnGap
        shp.Name = "btnSaveReports"
        shp.Fill.ForeColor.RGB = RGB(50, 120, 190)
        shp.TextFrame2.TextRange.Text = "4" & vbCrLf & ChrW(1513) & ChrW(1502) & ChrW(1497) & ChrW(1512) & ChrW(1514) & vbCrLf & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.TextFrame2.TextRange.Font.Size = 10
        shp.TextFrame2.TextRange.Font.Bold = msoTrue
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.OnAction = "ProVersionOnly"

        ' ---- Button 5: View Reports (Pro Only) ----
        Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("B2").Left + 2, btnTop, 80, calcBtnHeight)
        btnTop = shp.Top + shp.Height + btnGap
        shp.Name = "btnViewReports"
        shp.Fill.ForeColor.RGB = RGB(50, 170, 110)
        shp.TextFrame2.TextRange.Text = "5" & vbCrLf & ChrW(1510) & ChrW(1508) & ChrW(1497) & ChrW(1497) & ChrW(1492) & vbCrLf & ChrW(1489) & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.TextFrame2.TextRange.Font.Size = 10
        shp.TextFrame2.TextRange.Font.Bold = msoTrue
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.OnAction = "ProVersionOnly"

        ' ---- Button 6: New Clients? (Pro Only) ----
        Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("B2").Left + 2, btnTop, 80, calcBtnHeight)
        btnTop = shp.Top + shp.Height + 10 ' 10 points before show sheets
        shp.Name = "btnNewClients"
        shp.Fill.ForeColor.RGB = RGB(210, 130, 50)
        shp.TextFrame2.TextRange.Text = "6" & vbCrLf & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & vbCrLf & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501) & "?"
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.TextFrame2.TextRange.Font.Size = 10
        shp.TextFrame2.TextRange.Font.Bold = msoTrue
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.OnAction = "ProVersionOnly"

        ' Button: Show/Hide hidden sheets - placed below buttons
        Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("L16").Left, wsMain.Range("L16").Top, 160, 30)
        shp.Name = "btnShowHidden"
        shp.Fill.ForeColor.RGB = RGB(80, 80, 80)
        shp.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1510) & ChrW(1490) & "/" & ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1512) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & ChrW(1514)
        shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.TextFrame2.TextRange.Font.Size = 10
        shp.TextFrame2.TextRange.Font.Bold = msoTrue
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.OnAction = "ToggleHiddenSheets"
        
        ' ---- Restore Credit Text in A16 aligned left (near B) ----
        On Error Resume Next
        wsMain.Range("A18:F18").UnMerge
        wsMain.Range("A18:F18").Clear
        wsMain.Range("F16:G16").UnMerge
        wsMain.Range("F16:G16").Clear
        wsMain.Range("B16:C16").UnMerge
        wsMain.Range("B16:C16").Clear
        wsMain.Range("A16").Value = ChrW(1508) & ChrW(1497) & ChrW(1514) & ChrW(1493) & ChrW(1495) & " " & ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & " 054-6677396"
        wsMain.Range("A16").Font.Size = 12
        wsMain.Range("A16").Font.Bold = True
        wsMain.Range("A16").Font.Color = RGB(0, 100, 0)
        wsMain.Range("A16").HorizontalAlignment = xlLeft
        wsMain.Range("A16").VerticalAlignment = xlCenter
        On Error GoTo ERR_HANDLER

        ' ---- Store Hebrew message texts in column S of hagdarot ----
        ' S1=header, S2=done+found, S3=issues, S4=done ok, S5=line, S6=error
        ' S7=step, S8=confirm title, S9=btn2 confirm, S10=setup err, S11=dropdown err, S12=setup ok
        ' S13-S16=confirmation lines 1-4
865     wsMgmt.Cells(1, 19).Value = ChrW(1492) & ChrW(1493) & ChrW(1491) & ChrW(1506) & ChrW(1493) & ChrW(1514) & " " & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514)
866     wsMgmt.Cells(2, 19).Value = ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & "-" & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " "
867     wsMgmt.Cells(3, 19).Value = " " & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1493) & ChrW(1514)
868     wsMgmt.Cells(4, 19).Value = ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492)
869     wsMgmt.Cells(5, 19).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1492) & " "
        wsMgmt.Cells(6, 19).Value = ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " "
        wsMgmt.Cells(7, 19).Value = ChrW(1513) & ChrW(1500) & ChrW(1489) & ":" & " "
        wsMgmt.Cells(8, 19).Value = ChrW(1488) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1512) & " " & ChrW(1500) & ChrW(1508) & ChrW(1504) & ChrW(1497) & " " & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491)
        wsMgmt.Cells(9, 19).Value = ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1500) & ChrW(1497) & ChrW(1497) & ChrW(1513) & ChrW(1501) & " " & ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1512) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & "?"
        wsMgmt.Cells(10, 19).Value = ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1514) & " " & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514) & ":" & " "
        wsMgmt.Cells(11, 19).Value = ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1493) & ChrW(1503) & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492) & ":" & " "
        wsMgmt.Cells(12, 19).Value = ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514) & " " & ChrW(1492) & ChrW(1493) & ChrW(1511) & ChrW(1501) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!"
        wsMgmt.Cells(13, 19).Value = "1" & "." & " " & ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1492) & ChrW(1506) & ChrW(1500) & ChrW(1497) & ChrW(1514) & " " & ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509) & " " & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1508) & ChrW(1492) & " " & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1504) & ChrW(1497) & "?"
        wsMgmt.Cells(14, 19).Value = "2" & "." & " " & ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1513) & ChrW(1502) & ChrW(1512) & ChrW(1514) & " " & ChrW(1488) & ChrW(1514) & " " & _
            ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1488) & ChrW(1495) & ChrW(1512) & ChrW(1493) & ChrW(1503) & " " & "(" & ChrW(1492) & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & _
            ChrW(1489) & ChrW(1495) & ChrW(1493) & ChrW(1489) & ChrW(1512) & ChrW(1514) & " " & ChrW(1494) & ChrW(1493) & " " & ChrW(1497) & ChrW(1502) & ChrW(1495) & ChrW(1511) & ChrW(1493) & " " & _
            ChrW(1489) & ChrW(1514) & ChrW(1492) & ChrW(1500) & ChrW(1497) & ChrW(1498) & " " & ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & ")"
        wsMgmt.Cells(15, 19).Value = "3" & "." & " " & ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(1514) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1508) & ChrW(1512) & ChrW(1496) & ChrW(1497) & " " & ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1504) & ChrW(1491) & ChrW(1512) & ChrW(1513) & " " & ChrW(1489) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514) & "?"
        wsMgmt.Cells(16, 19).Value = "4" & "." & " " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " " & ChrW(1506) & ChrW(1500) & " " & "O" & "K" & " " & ChrW(1500) & ChrW(1492) & ChrW(1502) & ChrW(1513) & ChrW(1498)
        ' S17 = "don't show this message again?" text
        wsMgmt.Cells(17, 19).Value = ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1500) & ChrW(1492) & ChrW(1510) & ChrW(1497) & ChrW(1490) & " " & ChrW(1492) & ChrW(1493) & ChrW(1491) & ChrW(1506) & ChrW(1492) & " " & ChrW(1494) & ChrW(1493) & " " & ChrW(1513) & ChrW(1493) & ChrW(1489) & "?"
        ' S20 = flag: "1" means don't show confirmation again (empty = show)
        ' Don't overwrite S20 if already set

        ' ---- Set Error_Email parameter if not exists ----
895     Dim paramLastRow As Long
896     paramLastRow = wsMgmt.Cells(wsMgmt.Rows.count, COL_PARAM_NAME).End(xlUp).Row
897     Dim foundEmail As Boolean
898     foundEmail = False
899     Dim pr As Long
        For pr = 1 To paramLastRow
            If UCase$(Trim$(CStr(wsMgmt.Cells(pr, COL_PARAM_NAME).Value2))) = "ERROR_EMAIL" Then foundEmail = True: Exit For
        Next pr
        If Not foundEmail Then
            wsMgmt.Cells(paramLastRow + 1, COL_PARAM_NAME).Value = "ERROR_EMAIL"
            wsMgmt.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = "zvi@gorentech.co.il"
            paramLastRow = paramLastRow + 1
        End If

        ' ---- Set BACKUP_PATH parameter if not exists ----
        Dim foundBackup As Boolean
        foundBackup = False
        For pr = 1 To paramLastRow
            If UCase$(Trim$(CStr(wsMgmt.Cells(pr, COL_PARAM_NAME).Value2))) = "BACKUP_PATH" Then foundBackup = True: Exit For
        Next pr
        If Not foundBackup Then
            wsMgmt.Cells(paramLastRow + 1, COL_PARAM_NAME).Value = "BACKUP_PATH"
            wsMgmt.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = "C:\DEMO PROJECT\BACKUPS"
        End If

        ' ---- Set default values if empty ----
900     If IsEmpty(wsMain.Range("rngPeriodType").Value) Or wsMain.Range("rngPeriodType").Value = "" Then
902         wsMain.Range("rngPeriodType").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
904     End If

        ' ---- Set default value for rngFilterType if empty ----
        If IsEmpty(wsMain.Range("G9").Value) Or wsMain.Range("G9").Value = "" Then
            wsMain.Range("G9").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)
        End If

        ' ---- rngFilterType dropdown: report filter type ----
        ' Options: bechar/i, chevra, teler, sochen, anaf, anaf merkaz
        Dim filterTypeList As String
905     filterTypeList = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497) & "," & ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492) & "," & ChrW(1496) & ChrW(1500) & ChrW(1512) & "," & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503) & "," & ChrW(1506) & ChrW(1504) & ChrW(1507) & "," & ChrW(1506) & ChrW(1504) & ChrW(1507) & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)
        ' Use direct cell reference G9 in case named range not yet defined
906     On Error Resume Next
        wsMain.Range("G9").Validation.Delete
907     wsMain.Range("G9").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=filterTypeList
        wsMain.Range("G10").Validation.Delete
908     On Error GoTo ERR_HANDLER
        ' Create/update ALL named ranges for home sheet cells
        On Error Resume Next
        ThisWorkbook.names("rngCurrentYear").Delete
        ThisWorkbook.names("rngBaseYear").Delete
        ThisWorkbook.names("rngPeriodType").Delete
        ThisWorkbook.names("rngPeriodValue").Delete
        ThisWorkbook.names("rngDateType").Delete
        ThisWorkbook.names("rngFilterType").Delete
        ThisWorkbook.names("rngFilterValue").Delete
        On Error GoTo ERR_HANDLER
        ' v7.70: G3 is labeled "shnat basis" (base year), G4 is labeled "shna shoteft" (current year).
        ' Previous mapping was inverted, causing BuildReview to read the wrong year.
        ThisWorkbook.names.Add Name:="rngBaseYear", RefersTo:="='" & wsMain.Name & "'!$G$3"
        ThisWorkbook.names.Add Name:="rngCurrentYear", RefersTo:="='" & wsMain.Name & "'!$G$4"
        ThisWorkbook.names.Add Name:="rngPeriodType", RefersTo:="='" & wsMain.Name & "'!$G$5"
        ThisWorkbook.names.Add Name:="rngPeriodValue", RefersTo:="='" & wsMain.Name & "'!$G$6"
        ThisWorkbook.names.Add Name:="rngDateType", RefersTo:="='" & wsMain.Name & "'!$G$7"
        ThisWorkbook.names.Add Name:="rngFilterType", RefersTo:="='" & wsMain.Name & "'!$G$9"
        ThisWorkbook.names.Add Name:="rngFilterValue", RefersTo:="='" & wsMain.Name & "'!$G$10"

        ' ---- Client name filter (F12 label, G12 Data Validation dropdown) ----
        ' "shem lakoach" = ?? ????
        wsMain.Range("F12").Value = ChrW(1513) & ChrW(1501) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)
        wsMain.Range("F12").Font.Bold = True
        wsMain.Range("F12").Font.Size = 12
        wsMain.Range("F12").Font.Color = RGB(0, 70, 140)

        ' Create/update named range for G12
        On Error Resume Next
        ThisWorkbook.names("rngClientName").Delete
        On Error GoTo ERR_HANDLER
        ThisWorkbook.names.Add Name:="rngClientName", RefersTo:="='" & wsMain.Name & "'!$G$12"

        ' Set default value "bechar/i" = ???/?
        If IsEmpty(wsMain.Range("G12").Value) Or wsMain.Range("G12").Value = "" Then
            wsMain.Range("G12").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)
        End If

        ' Remove old ComboBox if exists (from previous versions)
        On Error Resume Next
        wsMain.OLEObjects("cmbClientName").Delete
        On Error GoTo ERR_HANDLER

        ' Remove Data Validation from G12 (from previous versions)
        On Error Resume Next
        wsMain.Range("G12").Validation.Delete
        On Error GoTo ERR_HANDLER

        ' Add search button (tight under table F)
        Dim btnSearch As Shape
        On Error Resume Next
        wsMain.Shapes("btnSearchClient").Delete
        On Error GoTo ERR_HANDLER
        Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, _
            wsMain.Range("F13").Left + (wsMain.Range("F13").Width - 50) / 2, _
            btnSearchTop, 50, 25)
        btnSearch.Name = "btnSearchClient"
        btnSearch.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513)  ' "???"
        btnSearch.TextFrame2.TextRange.Font.Size = 9
        btnSearch.TextFrame2.TextRange.Font.Bold = msoTrue
        btnSearch.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        btnSearch.Fill.ForeColor.RGB = RGB(70, 130, 180)
        btnSearch.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        btnSearch.OnAction = "SearchClientName"

        ' Add "kulam" (all) button (tight under table G)
        Dim btnAll As Shape
        On Error Resume Next
        wsMain.Shapes("btnAllClients").Delete
        On Error GoTo ERR_HANDLER
        Set btnAll = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, _
            wsMain.Range("G13").Left + (wsMain.Range("G13").Width - 50) / 2, _
            btnSearchTop, 50, 25)
        btnAll.Name = "btnAllClients"
        btnAll.TextFrame2.TextRange.Text = ChrW(1499) & ChrW(1493) & ChrW(1500) & ChrW(1501)  ' "????"
        btnAll.TextFrame2.TextRange.Font.Size = 9
        btnAll.TextFrame2.TextRange.Font.Bold = msoTrue
        btnAll.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        btnAll.Fill.ForeColor.RGB = RGB(60, 160, 60)
        btnAll.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        btnAll.OnAction = "ResetClientFilter"

        ' clear selection button tight under search buttons
        Dim btnClearSel As Shape
        On Error Resume Next
        wsMain.Shapes("btnClearSelection").Delete
        On Error GoTo ERR_HANDLER
        Set btnClearSel = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, _
            wsMain.Range("F14:G14").Left + (wsMain.Range("F14:G14").Width - 105) / 2, _
            btnClearTop, 105, 25)
        btnClearSel.Name = "btnClearSelection"
        ' "naki bechira" = clear selection
        btnClearSel.TextFrame2.TextRange.Text = ChrW(1488) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1505) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512) & ChrW(1497) & ChrW(1501)
        btnClearSel.TextFrame2.TextRange.Font.Size = 9
        btnClearSel.TextFrame2.TextRange.Font.Bold = msoTrue
        btnClearSel.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        btnClearSel.Fill.ForeColor.RGB = RGB(160, 60, 60)
        btnClearSel.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        btnClearSel.OnAction = "ClearSelection"


        ' ---- Set RTL and font size 14 for ALL sheets in workbook ----
        Dim wsLoop As Worksheet
        For Each wsLoop In ThisWorkbook.Worksheets
            wsLoop.DisplayRightToLeft = True
            wsLoop.Cells.Font.Size = 14
        Next wsLoop

        ' ---- Exchange rate info message at J5:K7 ----
        Dim rngMsg As Range
        ' Clear any old wider merge (e.g. J5:L7 from previous versions)
        wsMain.Range("J5:L7").UnMerge
        Set rngMsg = wsMain.Range("J5:K7")
        rngMsg.Merge
        ' Message: line1="???? ?????? ?"? ??? ?????" line2="?? ?? ????? ???" line3="???? ???? ????? ???"
        Dim msgTxt As String
        ' Line 1: hasha'ar mit'adken al pi bank israel
        msgTxt = ChrW(1492) & ChrW(1513) & ChrW(1506) & ChrW(1512) & " " & _
            ChrW(1502) & ChrW(1514) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1503) & " " & _
            ChrW(1506) & ChrW(34) & ChrW(1508) & " " & _
            ChrW(1489) & ChrW(1504) & ChrW(1511) & " " & _
            ChrW(1497) & ChrW(1513) & ChrW(1512) & ChrW(1488) & ChrW(1500) & vbLf
        ' Line 2: im lo yimatze sha'ar
        msgTxt = msgTxt & ChrW(1488) & ChrW(1501) & " " & _
            ChrW(1500) & ChrW(1488) & " " & _
            ChrW(1497) & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & _
            ChrW(1513) & ChrW(1506) & ChrW(1512) & vbLf
        ' Line 3: yeshamesh hasha'ar harashum kan
        msgTxt = msgTxt & ChrW(1497) & ChrW(1513) & ChrW(1502) & ChrW(1513) & " " & _
            ChrW(1492) & ChrW(1513) & ChrW(1506) & ChrW(1512) & " " & _
            ChrW(1492) & ChrW(1512) & ChrW(1513) & ChrW(1493) & ChrW(1501) & " " & _
            ChrW(1499) & ChrW(1488) & ChrW(1503)
        rngMsg.Value = msgTxt
        rngMsg.Font.Size = 10
        rngMsg.Font.Color = RGB(0, 70, 180)
        rngMsg.Font.Italic = True
        rngMsg.WrapText = True
        rngMsg.VerticalAlignment = xlCenter
        rngMsg.HorizontalAlignment = xlCenter
        ' Border as Shape (rectangle) around J5:K7 - works reliably with merged cells
        Dim borderShp As Shape
        On Error Resume Next
        wsMain.Shapes("shpMsgBorder").Delete
        On Error GoTo ERR_HANDLER
        Set borderShp = wsMain.Shapes.AddShape(msoShapeRectangle, _
            rngMsg.Left, rngMsg.Top, rngMsg.Width, rngMsg.Height)
        borderShp.Name = "shpMsgBorder"
        borderShp.Fill.Visible = msoFalse
        borderShp.Line.ForeColor.RGB = RGB(0, 0, 120)
        borderShp.Line.Weight = 1.5
        borderShp.Placement = xlMoveAndSize

        ' ---- Border for exchange rate cells J3:K4 using native cell borders instead of a shape (fixes weird leftover lines) ----
        Dim rngRates As Range
        Set rngRates = wsMain.Range("J3:K4")
        On Error Resume Next
        wsMain.Shapes("shpRateBorder").Delete
        On Error GoTo ERR_HANDLER
        rngRates.BorderAround xlContinuous, xlMedium, , RGB(0, 0, 120)
        
        ' ---- Center J3:K4 horizontally and vertically ----
        rngRates.HorizontalAlignment = xlCenter
        rngRates.VerticalAlignment = xlCenter

        ' ---- Populate Currency Rates ----
        wsMain.Range("J3").Value = ChrW(1491) & ChrW(1493) & ChrW(1500) & ChrW(1512) ' Dollar
        wsMain.Range("J4").Value = ChrW(1488) & ChrW(1497) & ChrW(1512) & ChrW(1493) ' Euro
        wsMain.Range("K3").Value = 2.991
        wsMain.Range("K4").Value = 3.4114

        ' ---- 7.33: borders for F3:G12 (single Range.Borders block) ----
        With wsMain.Range("F3:G12").Borders
            .LineStyle = xlContinuous
            .Color = RGB(0, 70, 140)
            .Weight = xlThin
        End With

        ' ---- Green pastel background for home page (Extended to AZ to cover shrinking) ----
        wsMain.Range("A1:AZ50").Interior.Color = RGB(220, 240, 220)
        
        ' ---- Pastel coloring for parameters table & currencies ----
        wsMain.Range("G3:G12").Interior.Color = RGB(255, 245, 225) ' Pastel orange/yellow for parameters
        wsMain.Range("F3:F12").Interior.Color = RGB(225, 245, 255) ' Pastel blue for values
        wsMain.Range("J3:K4").Interior.Color = RGB(255, 255, 200) ' Pastel yellow for currency rates

        ' ---- Re-apply row 2 header colors (overwritten by pastel above) ----
        hdrRng.Interior.Color = RGB(0, 100, 0)

        ' ---- Navigate to A1 ----
        wsMain.Activate
        wsMain.Range("A1").Select

910     MsgBoxU wsMgmt.Cells(12, 19).Value, vbInformation

        ' ---- v7.61: conditional formatting on G6 and G10 ----
        ' Helper values in NIHUL!Z1 (shnati) and NIHUL!Z2 (bachar/i),
        ' formula references those cells (avoids Hebrew literals).
        On Error Resume Next
        wsMgmt.Range("Z1").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
        wsMgmt.Range("Z2").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)
        Dim mgmtRef As String
        mgmtRef = "'" & MANAGEMENT_SHEET_NAME() & "'"
        wsMain.Range("G6").FormatConditions.Delete
        wsMain.Range("G6").FormatConditions.Add Type:=xlExpression, _
            Formula1:="=AND(LEN($G$5)>0,$G$5<>" & mgmtRef & "!$Z$1,$G$5<>" & mgmtRef & "!$Z$2,OR($G$6="""",$G$6=" & mgmtRef & "!$Z$2))"
        wsMain.Range("G6").FormatConditions(wsMain.Range("G6").FormatConditions.count).Interior.Color = RGB(255, 215, 0)
        wsMain.Range("G10").FormatConditions.Delete
        wsMain.Range("G10").FormatConditions.Add Type:=xlExpression, _
            Formula1:="=AND(LEN($G$9)>0,$G$9<>" & mgmtRef & "!$Z$2,OR($G$10="""",$G$10=" & mgmtRef & "!$Z$2))"
        wsMain.Range("G10").FormatConditions(wsMain.Range("G10").FormatConditions.count).Interior.Color = RGB(255, 215, 0)
        On Error GoTo ERR_HANDLER

        ' ---- Override years to 2024/2025 and block UI with transparent shape ----
        wsMain.Range("G3").Value = 2024
        wsMain.Range("G4").Value = 2025
        On Error Resume Next
        wsMain.Shapes("shpProYears").Delete
        On Error GoTo ERR_HANDLER
        Dim rngYearsBlock As Range
        Set rngYearsBlock = wsMain.Range("G3:G4")
        Dim shpYears As Shape
        Set shpYears = wsMain.Shapes.AddShape(msoShapeRectangle, rngYearsBlock.Left, rngYearsBlock.Top, rngYearsBlock.Width, rngYearsBlock.Height)
        shpYears.Name = "shpProYears"
        shpYears.Fill.Visible = msoFalse
        shpYears.Line.Visible = msoFalse
        shpYears.OnAction = "ProVersionOnly"

        ' ---- Remove Page Break lines (which can look like random lines between columns) ----
        wsMain.DisplayPageBreaks = False

        ' v7.62: removed sheet protection ? it broke Validation.Add for G6/G10 dropdowns
        ' after session reopen (UserInterfaceOnly does not persist). Trade-off:
        ' 2-clicks-needed on buttons is acceptable per user.
        ' HideRibbon                              ' v7.57 - user explicitly requested to keep Ribbon
        ShowRibbon                              ' Ensure ribbon is visible

920     Exit Sub

ERR_HANDLER:
930     MsgBoxU wsMgmt.Cells(10, 19).Value & Err.Description, vbCritical

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
        ' "shnatit" = yearly -> no second dropdown needed, jump to G7
150     ElseIf InStr(1, periodType, ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497), vbTextCompare) > 0 Then
155         wsMain.Range("rngPeriodValue").Value = ""
            On Error Resume Next
160         wsMain.Activate
161         wsMain.Range("G7").Select
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
207     wsMain.Range("rngPeriodValue").Select
        On Error GoTo ERR_HANDLER

CLEAN_EXIT:

210     Exit Sub

ERR_HANDLER:

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
110     If filterType = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492) Then
120         targetCol = 1
130     ElseIf filterType = ChrW(1496) & ChrW(1500) & ChrW(1512) Then
140         targetCol = 2
150     ElseIf filterType = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503) Then
160         targetCol = 3
170     ElseIf filterType = ChrW(1506) & ChrW(1504) & ChrW(1507) Then
180         targetCol = 4
190     ElseIf filterType = ChrW(1506) & ChrW(1504) & ChrW(1507) & ChrW(32) & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494) Then
200         targetCol = 5
210     Else
220         GoTo CLEAN_EXIT
230     End If

        ' Find the lists sheet
240     listsName = ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1493) & ChrW(1514)
250     If Not SheetExists(listsName) Then
260         MsgBoxU ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & listsName & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ". " & ChrW(1492) & ChrW(1512) & ChrW(1509) & " " & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " 1 " & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & ChrW(1492) & ".", vbExclamation
270         GoTo CLEAN_EXIT
280     End If

290     Set wsLists = ThisWorkbook.Worksheets(listsName)
300     lastR = wsLists.Cells(wsLists.Rows.count, targetCol).End(xlUp).Row

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
407     wsMain.Range("rngFilterValue").Select
        On Error GoTo ERR_HANDLER

CLEAN_EXIT:

410     Exit Sub

ERR_HANDLER:
420     MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1493) & ChrW(1503) & " " & ChrW(1505) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1503) & ": " & Err.Description, vbCritical

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

        ' 7.33: save Application state for clean restoration
        Dim prevCalcBP As XlCalculation
        Dim prevSU_BP As Boolean
        prevCalcBP = Application.Calculation
        prevSU_BP = Application.ScreenUpdating
        Application.Calculation = xlCalculationManual

        ' Remove any leftover sheet protection
        Dim wsUp3 As Worksheet
        For Each wsUp3 In ThisWorkbook.Worksheets
            On Error Resume Next
            wsUp3.Unprotect SHEET_PROTECT_PWD
            On Error GoTo ERR_HANDLER
        Next wsUp3

        Dim wsMain As Worksheet
20      Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())

        Dim yearVal As String
        Dim refYear As String
        Dim periodType As String
        Dim periodDetail As String
        Dim periodDesc As String
40      yearVal = "2020" ' Forced for demo
        ' yearVal = Trim$(CStr(wsMain.Range("rngCurrentYear").Value2))
50      refYear = "2019" ' Forced for demo
        ' refYear = Trim$(CStr(wsMain.Range("rngBaseYear").Value2))
        
        ' FORCE YEARS FOR DEMO OUTPUT (Requested by user)
        yearVal = "2020"
        refYear = "2019"
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
        detailBy = Trim$(CStr(wsMain.Range("G10").Value2))
        clientName = Trim$(CStr(wsMain.Range("G12").Value2))
        ' Build subtitle: periodDesc | dateType | detailBy | clientName (no years - already in chart)
        ' v7.66: skip any value that is empty OR equals "bachar/i" default
        Dim bacharI As String
        bacharI = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)
        paramsSubtitle = ""
        If periodDesc <> "" And periodDesc <> bacharI Then paramsSubtitle = periodDesc
        If dateType <> "" And dateType <> bacharI Then paramsSubtitle = paramsSubtitle & " | " & dateType
        If detailBy <> "" And detailBy <> bacharI Then paramsSubtitle = paramsSubtitle & " | " & detailBy
        If clientName <> "" And clientName <> bacharI Then paramsSubtitle = paramsSubtitle & " | " & clientName
        ' Remove leading " | " if periodDesc was empty
        If Left$(paramsSubtitle, 3) = " | " Then paramsSubtitle = Mid$(paramsSubtitle, 4)

        ' Validate that comparison sheets exist
100     If Not SheetExists(SHEET_COMPANIES()) Then
110         MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1492) & ChrW(1512) & ChrW(1497) & ChrW(1509) & " " & ChrW(1511) & ChrW(1493) & ChrW(1491) & ChrW(1501) & " " & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & " 2", vbCritical
120         Exit Sub
130     End If

        ' Show processing message (below currency area)
140     With wsMain.Range("B12")
150         .Value = ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1512) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & "..."
160         .Font.Size = 16
170         .Font.Bold = True
180         .Font.Color = RGB(200, 0, 0)
190         .Interior.Color = RGB(255, 255, 200)
200     End With
210     Application.ScreenUpdating = True
220     DoEvents

        ' ================================================================
        ' PHASE 1: Create all chart images in Excel (NO PowerPoint yet)
        ' ================================================================
        Dim tmpPath As String
230     tmpPath = Environ$("TEMP") & "\"

        Dim imgTotal As String
240     imgTotal = tmpPath & "chart_total.gif"
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

        ' v7.43: Export 4 charts per sheet (premium, docs, policies, commission)
        Dim si As Long
        Dim imgFiles() As String
470     ReDim imgFiles(1 To sheetCount * 5)  ' v7.57: 5 charts per sheet
        Dim exportOK() As Boolean
480     ReDim exportOK(1 To sheetCount)
490     For si = 1 To sheetCount
500         imgFiles(si * 5 - 4) = tmpPath & "chart_prem_" & si & ".gif"
505         imgFiles(si * 5 - 3) = tmpPath & "chart_doc_" & si & ".gif"
510         imgFiles(si * 5 - 2) = tmpPath & "chart_pol_" & si & ".gif"
512         imgFiles(si * 5 - 1) = tmpPath & "chart_ins_" & si & ".gif"
515         imgFiles(si * 5) = tmpPath & "chart_comm_" & si & ".gif"
520         On Error Resume Next
530         ExportCompCharts sheetList(si), imgFiles(si * 5 - 4), imgFiles(si * 5 - 3), imgFiles(si * 5 - 2), imgFiles(si * 5 - 1), imgFiles(si * 5), yearVal, refYear
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
610     Set ppApp = CreateObject("PowerPoint.Application")
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
660     titleNames(3) = ChrW(1506) & ChrW(1504) & ChrW(1507) & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)
665     titleNames(4) = ChrW(1496) & ChrW(1500) & ChrW(1512) & ChrW(1497) & ChrW(1493) & ChrW(1514)
670     titleNames(5) = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501)

        ' SLIDE 1: Title
675     slideIdx = slideIdx + 1
680     Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
685     BuildTitleSlide ppSlide, yearVal, refYear, periodDesc, slideW, slideH, paramsSubtitle

        ' SLIDE 2: Total Summary chart
690     slideIdx = slideIdx + 1
695     Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
700     BuildTotalSlideFromImage ppSlide, imgTotal, yearVal, refYear, slideW, paramsSubtitle

        ' v7.43: For each comparison sheet: 5 slides (prem, docs, policies, comm, table)
730     For si = 1 To sheetCount
740         If exportOK(si) Then
                ' Slide: Premium chart
750             slideIdx = slideIdx + 1
751             Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
752             BuildChartSlide ppSlide, imgFiles(si * 5 - 4), ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & titleNames(si), yearVal, refYear, slideW, paramsSubtitle
                ' Slide: Documents chart
753             slideIdx = slideIdx + 1
754             Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
755             BuildChartSlide ppSlide, imgFiles(si * 5 - 3), ChrW(1502) & ChrW(1505) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & titleNames(si), yearVal, refYear, slideW, paramsSubtitle
                ' Slide: Policies chart
756             slideIdx = slideIdx + 1
757             Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
758             BuildChartSlide ppSlide, imgFiles(si * 5 - 2), ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1505) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & titleNames(si), yearVal, refYear, slideW, paramsSubtitle
                ' Slide (v7.57): Insured chart
                slideIdx = slideIdx + 1
                Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
                BuildChartSlide ppSlide, imgFiles(si * 5 - 1), ChrW(1502) & ChrW(1489) & ChrW(1493) & ChrW(1496) & ChrW(1495) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & titleNames(si), yearVal, refYear, slideW, paramsSubtitle
                ' Slide: Commission chart
759             slideIdx = slideIdx + 1
760             Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
761             BuildChartSlide ppSlide, imgFiles(si * 5), ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & titleNames(si), yearVal, refYear, slideW, paramsSubtitle
810         End If
            ' Slide: Data table
820         slideIdx = slideIdx + 1
830         Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
840         BuildTableSlide ppSlide, sheetList(si), titleNames(si), yearVal, refYear, slideW, slideH, paramsSubtitle
850     Next si

        ' ================================================================
        ' PHASE 3: variant slides excluding one specified agent
        ' ================================================================
        If SheetExists(SHEET_AGENTS()) Then
            Dim imgNoExclPrem As String
            Dim imgNoExclComm As String
            Dim noExclOK As Boolean
            Dim exclAgentName As String
            exclAgentName = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503) & " 1"
            imgNoExclPrem = tmpPath & "chart_excl_prem.gif"
            imgNoExclComm = tmpPath & "chart_excl_comm.gif"
            noExclOK = False
            On Error Resume Next
            ' v7.43: now passing 4 image paths (premium, docs, policies, commission)
            Dim imgNoExclDoc As String, imgNoExclPol As String, imgNoExclIns As String
            imgNoExclDoc = tmpPath & "chart_excl_doc.gif"
            imgNoExclPol = tmpPath & "chart_excl_pol.gif"
            imgNoExclIns = tmpPath & "chart_excl_ins.gif"
            ExportCompCharts SHEET_AGENTS(), imgNoExclPrem, imgNoExclDoc, imgNoExclPol, imgNoExclIns, imgNoExclComm, yearVal, refYear, exclAgentName
            If Err.Number = 0 Then noExclOK = True
            Err.Clear
            On Error GoTo ERR_HANDLER
            If noExclOK Then
                ' Slide: Premiums excluding the specified agent
                slideIdx = slideIdx + 1
                Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
                ' title: premiums by agent, excluding the specified agent
                BuildChartSlide ppSlide, imgNoExclPrem, ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1500) & ChrW(1488) & " " & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503) & " 1", yearVal, refYear, slideW, paramsSubtitle
                ' Slide: Commissions excluding the specified agent
                slideIdx = slideIdx + 1
                Set ppSlide = ppPres.Slides.Add(slideIdx, 12)
                ' title: commissions by agent, excluding the specified agent
                BuildChartSlide ppSlide, imgNoExclComm, ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1508) & ChrW(1497) & " " & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1500) & ChrW(1500) & ChrW(1488) & " " & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503) & " 1", yearVal, refYear, slideW, paramsSubtitle
            End If
            ' Cleanup temp images
            On Error Resume Next
            Kill imgNoExclPrem
            Kill imgNoExclDoc
            Kill imgNoExclPol
            Kill imgNoExclIns
            Kill imgNoExclComm
            On Error GoTo ERR_HANDLER
        End If

        ' Add page numbers to all slides
        Dim pg As Long
860     For pg = 1 To ppPres.Slides.count
870         AddPageNumber ppPres.Slides(pg), pg, ppPres.Slides.count, slideW, slideH
880     Next pg

        ' v7.57: do NOT auto-save the PPTX or PDF.  The deck stays open in
        ' PowerPoint as an unsaved presentation; PDF is offered later via
        ' Yes/No prompt.
        Dim savePath As String
890     savePath = REPORTS_FOLDER() & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1492) & ChrW(1504) & ChrW(1492) & ChrW(1500) & ChrW(1492) & " " & yearVal & ".pptx"
        Dim pdfPath As String
905     pdfPath = Replace(savePath, ".pptx", ".pdf")

        ' v7.54: keep ppApp / ppPres references alive until after the success
        ' MsgBox so we can re-activate PowerPoint at the end.
        ' Cleanup temp images
950     On Error Resume Next
960     Kill imgTotal
970     For si = 1 To sheetCount * 5
980         Kill imgFiles(si)
990     Next si
1000    On Error GoTo ERR_HANDLER

        ' Clear processing message and restore green background
1010    With wsMain.Range("B12")
1020        .Value = ""
1030        .Interior.Color = RGB(220, 240, 220)
1040    End With

        ' v7.57: bring PowerPoint to the foreground BEFORE the question.
1051    On Error Resume Next
        Dim ppHwnd As LongPtr
        ppHwnd = ppApp.ActiveWindow.hWnd
        If ppHwnd <> 0 Then
            ShowWindow ppHwnd, SW_SHOW
            SetForegroundWindow ppHwnd
        End If
1052    ppApp.Activate
1053    AppActivate ppApp.Caption
1054    On Error GoTo ERR_HANDLER

        ' v7.67: PDF action disabled for demo (caused Excel crash).
        ' In production version, the presentation will also be produced as PDF.
        PositionNextMsgBox 40, 80
        ' "Hamatzget nutzra behatzlacha. Be'girsat amet tufak gam ke-PDF"
1050    MsgBoxU ChrW(1492) & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "." & vbCrLf & ChrW(1489) & ChrW(1490) & ChrW(1512) & ChrW(1505) & ChrW(1514) & " " & ChrW(1488) & ChrW(1502) & ChrW(1514) & " " & ChrW(1514) & ChrW(1493) & ChrW(1508) & ChrW(1511) & " " & ChrW(1490) & ChrW(1501) & " " & ChrW(1499) & "-PDF", vbInformation

        ' Release references (PowerPoint stays open with the deck visible)
        Set ppPres = Nothing
        Set ppApp = Nothing

        ' 7.33: restore Application state
1055    Application.Calculation = prevCalcBP
1056    Application.ScreenUpdating = prevSU_BP

1060    Exit Sub

ERR_HANDLER:
        Dim errLine As Long
        Dim errDesc As String
        Dim errNum As Long
1070    errLine = Erl
1072    errDesc = Err.Description
1074    errNum = Err.Number
1076    On Error Resume Next
        With wsMain.Range("B12")
            .Value = ""
            .Interior.Color = RGB(220, 240, 220)
        End With
        If Not ppPres Is Nothing Then ppPres.Close
        If Not ppApp Is Nothing Then ppApp.Quit
        Kill imgTotal
        Dim ei As Long
        For ei = 1 To sheetCount * 5
            Kill imgFiles(ei)
        Next ei
1080    MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1497) & ChrW(1510) & ChrW(1497) & ChrW(1512) & ChrW(1514) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & ":" & vbCrLf & "Line: " & errLine & vbCrLf & "Err #" & errNum & ": " & errDesc, vbCritical
        ' 7.33: best-effort restore Application state
1090    On Error Resume Next
1091    Application.Calculation = prevCalcBP
1092    Application.ScreenUpdating = prevSU_BP

End Sub


' ============================================================================
' HELPER: Export Total Summary chart to image file
' ============================================================================
Private Sub ExportTotalChart(ByVal imgPath As String, ByVal yearVal As String, ByVal refYear As String)

10      On Error GoTo ERR_HANDLER

        Dim ws As Worksheet
        Dim lastRow As Long
20      Set ws = ThisWorkbook.Worksheets(SHEET_MONTHS())
30      lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

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
320     On Error Resume Next
        Application.DisplayAlerts = False
        If Not tmpWs Is Nothing Then tmpWs.Delete
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
330     Err.Raise Err.Number, "ExportTotalChart:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Export Comparison charts (premiums + commissions) to 2 image files
' ============================================================================
Private Sub ExportCompCharts(ByVal sheetName As String, ByVal imgPrem As String, ByVal imgDoc As String, ByVal imgPol As String, ByVal imgIns As String, ByVal imgComm As String, ByVal yearVal As String, ByVal refYear As String, Optional ByVal excludeName As String = "")

10      On Error GoTo ERR_HANDLER

        Dim ws As Worksheet
        Dim lastRow As Long
        Dim dataRows As Long
        Dim r As Long
        Dim tmpName As String
20      Set ws = ThisWorkbook.Worksheets(sheetName)
30      lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
40      dataRows = lastRow - 3

        Dim arrNames() As String
        Dim arrPremR() As Double, arrPremC() As Double
        Dim arrDocR() As Double, arrDocC() As Double
        Dim arrPolR() As Double, arrPolC() As Double
        Dim arrInsR() As Double, arrInsC() As Double
        Dim arrCommR() As Double, arrCommC() As Double
        Dim nItems As Long
50      nItems = dataRows - 1
60      If nItems < 1 Then Exit Sub

70      ReDim arrNames(1 To nItems)
80      ReDim arrPremR(1 To nItems), arrPremC(1 To nItems)
82      ReDim arrDocR(1 To nItems), arrDocC(1 To nItems)
84      ReDim arrPolR(1 To nItems), arrPolC(1 To nItems)
86      ReDim arrInsR(1 To nItems), arrInsC(1 To nItems)
90      ReDim arrCommR(1 To nItems), arrCommC(1 To nItems)

        Dim idx As Long
100     idx = 0
110     For r = 4 To lastRow - 1
            tmpName = ShortenCompanyName(Trim$(CStr(ws.Cells(r, 1).Value2)))
            ' v7.40: exact match (was substring) so 'Sokhen 1' won't accidentally match 'Sokhen 10' etc.
            If Len(excludeName) > 0 Then
                If StrComp(Trim$(tmpName), excludeName, vbTextCompare) = 0 Then GoTo NEXT_ROW_ECC
            End If
120         idx = idx + 1
130         If idx > nItems Then Exit For
140         arrNames(idx) = tmpName
150         arrPremR(idx) = CDbl(ws.Cells(r, 2).Value2)
160         arrPremC(idx) = CDbl(ws.Cells(r, 3).Value2)
162         arrDocR(idx) = CDbl(ws.Cells(r, 5).Value2)
164         arrDocC(idx) = CDbl(ws.Cells(r, 6).Value2)
166         arrPolR(idx) = CDbl(ws.Cells(r, 11).Value2)
168         arrPolC(idx) = CDbl(ws.Cells(r, 12).Value2)
169         arrInsR(idx) = CDbl(ws.Cells(r, 8).Value2)
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
        Dim s As Long
230     Application.ScreenUpdating = False

        ' ============================================================
        ' Helper: build a single 2-series column chart
        ' Reuses one tmpWs across all 4 charts to avoid clutter
        ' ============================================================
240     Set tmpWs = ThisWorkbook.Worksheets.Add

        ' ---- Chart 1: Premiums (yellow=ref, blue=cur) ----
245     CallExportOneChart tmpWs, chartItems, arrNames, arrPremR, arrPremC, refYear, yearVal, RGB(255, 192, 0), RGB(68, 114, 196), imgPrem

        ' ---- Chart 2: Documents (purple=ref, teal=cur) ----
250     CallExportOneChart tmpWs, chartItems, arrNames, arrDocR, arrDocC, refYear, yearVal, RGB(128, 100, 162), RGB(75, 172, 198), imgDoc, "#,##0"

        ' ---- Chart 3: Policies (red=ref, dark green=cur) ----
255     CallExportOneChart tmpWs, chartItems, arrNames, arrPolR, arrPolC, refYear, yearVal, RGB(192, 80, 77), RGB(0, 130, 70), imgPol, "#,##0"

        ' ---- Chart 4 (v7.57): Insured (cyan=ref, dark blue=cur) ----
257     CallExportOneChart tmpWs, chartItems, arrNames, arrInsR, arrInsC, refYear, yearVal, RGB(91, 155, 213), RGB(31, 78, 121), imgIns, "#,##0"

        ' ---- Chart 5: Commissions (orange=ref, green=cur) ----
260     CallExportOneChart tmpWs, chartItems, arrNames, arrCommR, arrCommC, refYear, yearVal, RGB(237, 125, 49), RGB(112, 173, 71), imgComm

        ' Cleanup
570     Application.DisplayAlerts = False
580     tmpWs.Delete
590     Application.DisplayAlerts = True
600     Application.ScreenUpdating = True

610     Exit Sub
ERR_HANDLER:
620     On Error Resume Next
        Application.DisplayAlerts = False
        If Not tmpWs Is Nothing Then tmpWs.Delete
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
630     Err.Raise Err.Number, "ExportCompCharts(" & sheetName & "):" & Erl, Err.Description
End Sub

' ============================================================================
' HELPER (7.43): Build one column chart on tmpWs and export to file
' ============================================================================
Private Sub CallExportOneChart(ByVal tmpWs As Worksheet, ByVal chartItems As Long, _
        ByRef arrNames() As String, ByRef arrRef() As Double, ByRef arrCur() As Double, _
        ByVal refYear As String, ByVal yearVal As String, _
        ByVal colorRef As Long, ByVal colorCur As Long, _
        ByVal imgPath As String, Optional ByVal numFmt As String = "")

10      On Error GoTo ERR_HANDLER

        ' Clear sheet
20      tmpWs.Cells.Clear
        Dim co As Object
30      For Each co In tmpWs.ChartObjects: co.Delete: Next co

        ' Headers
40      tmpWs.Cells(1, 1).Value = ""
50      tmpWs.Cells(1, 2).Value = refYear
60      tmpWs.Cells(1, 3).Value = yearVal

        Dim ci As Long
70      For ci = 1 To chartItems
80          tmpWs.Cells(ci + 1, 1).Value = arrNames(ci)
90          tmpWs.Cells(ci + 1, 2).Value = arrRef(ci)
100         tmpWs.Cells(ci + 1, 3).Value = arrCur(ci)
110     Next ci

        Dim chtObj As Object
        Dim xlCht As Object
120     Set chtObj = tmpWs.ChartObjects.Add(10, 10, 600, 350)
130     Set xlCht = chtObj.Chart
140     xlCht.ChartType = 51   ' xlColumnClustered
150     xlCht.SetSourceData tmpWs.Range(tmpWs.Cells(1, 1), tmpWs.Cells(chartItems + 1, 3)), 2  ' xlColumns
160     xlCht.HasTitle = False
170     xlCht.HasLegend = True

        ' Colors
180     xlCht.SeriesCollection(1).Format.Fill.ForeColor.RGB = colorRef
190     xlCht.SeriesCollection(2).Format.Fill.ForeColor.RGB = colorCur

        ' Data labels
        Dim s As Long
        Dim effFmt As String
        effFmt = numFmt
        If effFmt = "" Then effFmt = "#,##0,""K"""  ' default: K-thousands
200     For s = 1 To 2
210         xlCht.SeriesCollection(s).HasDataLabels = True
220         xlCht.SeriesCollection(s).DataLabels.NumberFormat = effFmt
230         xlCht.SeriesCollection(s).DataLabels.Font.Size = 9
240         xlCht.SeriesCollection(s).DataLabels.Orientation = 90
250     Next s

260     xlCht.Axes(2).TickLabels.NumberFormat = effFmt

        ' Export
270     xlCht.Export imgPath

        ' Cleanup
280     chtObj.Delete

290     Exit Sub
ERR_HANDLER:
300     Err.Raise Err.Number, "CallExportOneChart:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Build Title Slide (landscape)
' ============================================================================
Private Sub BuildTitleSlide(ByVal ppSlide As Object, ByVal yearVal As String, ByVal refYear As String, ByVal periodDesc As String, ByVal slideW As Single, ByVal slideH As Single, Optional ByVal paramsSubtitle As String = "")

10      On Error GoTo ERR_HANDLER

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
        ' 7.35: was hardcoded; now reads Agency_Name parameter
224     shp.TextFrame.TextRange.Text = AGENCY_NAME()
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
330     Err.Raise Err.Number, "BuildTitleSlide:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Build Total Summary Slide from pre-exported image
' ============================================================================
Private Sub BuildTotalSlideFromImage(ByVal ppSlide As Object, ByVal imgPath As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single, Optional ByVal paramsSubtitle As String = "")

10      On Error GoTo ERR_HANDLER

        Dim shp As Object

        ' v7.53: 3-line title (agency / type / params)
        Set shp = ppSlide.Shapes.AddTextbox(1, 20, 4, slideW - 40, 26)
        shp.TextFrame.TextRange.Text = AGENCY_NAME()
        shp.TextFrame.TextRange.Font.Size = 18
        shp.TextFrame.TextRange.Font.Bold = True
        shp.TextFrame.TextRange.Font.Color.RGB = RGB(0, 70, 140)
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
20      Set shp = ppSlide.Shapes.AddTextbox(1, 20, 30, slideW - 40, 28)
30      shp.TextFrame.TextRange.Text = ChrW(1505) & ChrW(1492) & Chr(34) & ChrW(1499) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514) & " " & ChrW(1493) & ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514) & " - " & refYear & " " & ChrW(1502) & ChrW(1493) & ChrW(1500) & " " & yearVal
40      shp.TextFrame.TextRange.Font.Size = 20
50      shp.TextFrame.TextRange.Font.Bold = True
60      shp.TextFrame.TextRange.Font.Color.RGB = RGB(50, 50, 50)
70      shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        If paramsSubtitle <> "" Then
            Set shp = ppSlide.Shapes.AddTextbox(1, 40, 58, slideW - 80, 20)
            shp.TextFrame.TextRange.Text = paramsSubtitle
            shp.TextFrame.TextRange.Font.Size = 11
            shp.TextFrame.TextRange.Font.Color.RGB = RGB(120, 120, 120)
            shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        End If
80      shp.TextFrame.WordWrap = True

        ' Insert chart image (landscape: wider)
90      ppSlide.Shapes.AddPicture imgPath, 0, 1, 80, 88, slideW - 160, 425

100     Exit Sub
ERR_HANDLER:
110     Err.Raise Err.Number, "BuildTotalSlideFromImage:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Build a single chart slide (one chart image + title)
' ============================================================================
Private Sub BuildChartSlide(ByVal ppSlide As Object, ByVal imgPath As String, ByVal chartTitle As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single, Optional ByVal paramsSubtitle As String = "")

10      On Error GoTo ERR_HANDLER

        Dim shp As Object

        ' v7.53: 3-line title
        Set shp = ppSlide.Shapes.AddTextbox(1, 20, 4, slideW - 40, 24)
        shp.TextFrame.TextRange.Text = AGENCY_NAME()
        shp.TextFrame.TextRange.Font.Size = 16
        shp.TextFrame.TextRange.Font.Bold = True
        shp.TextFrame.TextRange.Font.Color.RGB = RGB(0, 70, 140)
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
20      Set shp = ppSlide.Shapes.AddTextbox(1, 20, 28, slideW - 40, 26)
30      shp.TextFrame.TextRange.Text = chartTitle & " - " & refYear & " " & ChrW(1502) & ChrW(1493) & ChrW(1500) & " " & yearVal
40      shp.TextFrame.TextRange.Font.Size = 18
50      shp.TextFrame.TextRange.Font.Bold = True
60      shp.TextFrame.TextRange.Font.Color.RGB = RGB(50, 50, 50)
70      shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
80      shp.TextFrame.WordWrap = True

        ' Subtitle (parameters)
        If paramsSubtitle <> "" Then
82          Set shp = ppSlide.Shapes.AddTextbox(1, 40, 54, slideW - 80, 18)
84          shp.TextFrame.TextRange.Text = paramsSubtitle
86          shp.TextFrame.TextRange.Font.Size = 11
87          shp.TextFrame.TextRange.Font.Bold = False
88          shp.TextFrame.TextRange.Font.Color.RGB = RGB(120, 120, 120)
89          shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        End If

        ' Insert chart image
90      ppSlide.Shapes.AddPicture imgPath, 0, 1, 60, 80, slideW - 120, 436

100     Exit Sub
ERR_HANDLER:
110     Err.Raise Err.Number, "BuildChartSlide:" & Erl, Err.Description
End Sub


' ============================================================================
' HELPER: Build a data table slide
' 13 cols: name | premRef | premCur | prem% | docsRef | docsCur | docs% |
'          insuredRef | insuredCur | ins% | commRef | commCur | comm%
' ============================================================================
Private Sub BuildTableSlide(ByVal ppSlide As Object, ByVal sheetName As String, ByVal slideTitle As String, ByVal yearVal As String, ByVal refYear As String, ByVal slideW As Single, ByVal slideH As Single, Optional ByVal paramsSubtitle As String = "")

10      On Error GoTo ERR_HANDLER

        Dim ws As Worksheet
        Dim lastRow As Long
        Dim r As Long
        Dim shp As Object
20      Set ws = ThisWorkbook.Worksheets(sheetName)
30      lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

        Dim nItems As Long
40      nItems = lastRow - 4
50      If nItems < 1 Then Exit Sub

        ' Read data: name, premR, premC, docR, docC, insR, insC, polR, polC, commR, commC
        Dim arrNames() As String
        Dim arrPremR() As Double, arrPremC() As Double
        Dim arrDocR() As Long, arrDocC() As Long
        Dim arrInsR() As Long, arrInsC() As Long
        Dim arrPolR() As Long, arrPolC() As Long  ' v7.42: policies
        Dim arrCommR() As Double, arrCommC() As Double

        ReDim arrNames(1 To nItems)
        ReDim arrPremR(1 To nItems), arrPremC(1 To nItems)
        ReDim arrDocR(1 To nItems), arrDocC(1 To nItems)
        ReDim arrInsR(1 To nItems), arrInsC(1 To nItems)
        ReDim arrPolR(1 To nItems), arrPolC(1 To nItems)
        ReDim arrCommR(1 To nItems), arrCommC(1 To nItems)

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
            ' v7.42: policies (cols 11-12 of comparison sheet)
212         arrPolR(idx) = CLng(ws.Cells(r, 11).Value2)
213         arrPolC(idx) = CLng(ws.Cells(r, 12).Value2)
220         arrCommR(idx) = CDbl(ws.Cells(r, 14).Value2)
230         arrCommC(idx) = CDbl(ws.Cells(r, 15).Value2)
240     Next r
250     nItems = idx

        ' --- v7.53: 3-line title ---
        Set shp = ppSlide.Shapes.AddTextbox(1, 20, 2, slideW - 40, 22)
        shp.TextFrame.TextRange.Text = AGENCY_NAME()
        shp.TextFrame.TextRange.Font.Size = 14
        shp.TextFrame.TextRange.Font.Bold = True
        shp.TextFrame.TextRange.Font.Color.RGB = RGB(0, 70, 140)
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
260     Set shp = ppSlide.Shapes.AddTextbox(1, 20, 22, slideW - 40, 22)
270     shp.TextFrame.TextRange.Text = slideTitle & " - " & ChrW(1496) & ChrW(1489) & ChrW(1500) & ChrW(1514) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
280     shp.TextFrame.TextRange.Font.Size = 14
290     shp.TextFrame.TextRange.Font.Bold = True
300     shp.TextFrame.TextRange.Font.Color.RGB = RGB(50, 50, 50)
310     shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
320     shp.TextFrame.WordWrap = True

        ' --- Subtitle ---
        If paramsSubtitle <> "" Then
322         Set shp = ppSlide.Shapes.AddTextbox(1, 40, 42, slideW - 80, 16)
324         shp.TextFrame.TextRange.Text = paramsSubtitle
326         shp.TextFrame.TextRange.Font.Size = 9
328         shp.TextFrame.TextRange.Font.Color.RGB = RGB(120, 120, 120)
329         shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        End If

        ' --- Table: 16 cols, RTL layout ---
        ' Col 1 (leftmost) = Commission %, ..., Col 16 (rightmost) = Name
        ' Categories left-to-right: Commission(1-3), Policies(4-6), Insured(7-9), Docs(10-12), Premium(13-15), Name(16)
        Dim tblRows As Long, tblCols As Long
        Dim tblTop As Single, tblLeft As Single, tblWidth As Single, tblHeight As Single, rowH As Single
        Dim ppTbl As Object, tbl As Object
        Dim c As Long, tblR As Long
        Dim blueClr As Long, lightBlue As Long, pctLabel As String

330     tblRows = nItems + 3
340     tblCols = 16
350     tblTop = 62
360     tblLeft = 10
370     tblWidth = slideW - 20
        ' v7.52: taller rows so wider value cells breathe
380     rowH = 18
390     If tblRows > 10 Then rowH = 16
392     If tblRows > 15 Then rowH = 14
394     If tblRows > 20 Then rowH = 12
400     tblHeight = tblRows * rowH

410     Set ppTbl = ppSlide.Shapes.AddTable(tblRows, tblCols, tblLeft, tblTop, tblWidth, tblHeight)
420     Set tbl = ppTbl.Table
        ' v7.44: lock row heights so cells with negative-% don't auto-grow
422     Dim rowI As Long
424     For rowI = 1 To tblRows
426         tbl.Rows(rowI).Height = rowH
428     Next rowI

        ' v7.52: wider value cols at the expense of the names column.
        ' Total = 0.08 + 5*0.04 + 10*0.072 = 1.00.
430     tbl.Columns(1).Width = tblWidth * 0.05    ' comm %
432     tbl.Columns(2).Width = tblWidth * 0.068   ' comm cur
434     tbl.Columns(3).Width = tblWidth * 0.068   ' comm ref
436     tbl.Columns(4).Width = tblWidth * 0.05    ' pol %
438     tbl.Columns(5).Width = tblWidth * 0.068   ' pol cur
440     tbl.Columns(6).Width = tblWidth * 0.068   ' pol ref
442     tbl.Columns(7).Width = tblWidth * 0.05    ' ins %
444     tbl.Columns(8).Width = tblWidth * 0.068   ' ins cur
446     tbl.Columns(9).Width = tblWidth * 0.068   ' ins ref
448     tbl.Columns(10).Width = tblWidth * 0.05   ' doc %
450     tbl.Columns(11).Width = tblWidth * 0.068  ' doc cur
452     tbl.Columns(12).Width = tblWidth * 0.068  ' doc ref
454     tbl.Columns(13).Width = tblWidth * 0.05   ' prem %
456     tbl.Columns(14).Width = tblWidth * 0.068  ' prem cur
458     tbl.Columns(15).Width = tblWidth * 0.068  ' prem ref
460     tbl.Columns(16).Width = tblWidth * 0.07   ' name (v7.53: was 0.08)

470     blueClr = RGB(0, 100, 170)
480     lightBlue = RGB(180, 210, 240)
490     pctLabel = ChrW(1513) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1497)

        ' --- Header row 1: merge 3 cells per category, set category names ---
        ' Category names (Hebrew):
        Dim hdrComm As String, hdrPol As String, hdrIns As String, hdrDoc As String, hdrPrem As String
500     hdrComm = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1493) & ChrW(1514)         ' amulot
510     hdrPol = ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1505) & ChrW(1493) & ChrW(1514)  ' polisot
520     hdrIns = ChrW(1502) & ChrW(1489) & ChrW(1493) & ChrW(1496) & ChrW(1495) & ChrW(1497) & ChrW(1501)  ' mevutachim
530     hdrDoc = ChrW(1502) & ChrW(1505) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1501)              ' mismachim
540     hdrPrem = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1514)             ' premiot

        ' Merge cells for each category and set text
        On Error Resume Next
545     tbl.Cell(1, 1).Merge tbl.Cell(1, 3)
546     tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text = hdrComm
547     tbl.Cell(1, 4).Merge tbl.Cell(1, 6)
548     tbl.Cell(1, 4).Shape.TextFrame.TextRange.Text = hdrPol
549     tbl.Cell(1, 7).Merge tbl.Cell(1, 9)
550     tbl.Cell(1, 7).Shape.TextFrame.TextRange.Text = hdrIns
551     tbl.Cell(1, 10).Merge tbl.Cell(1, 12)
552     tbl.Cell(1, 10).Shape.TextFrame.TextRange.Text = hdrDoc
553     tbl.Cell(1, 13).Merge tbl.Cell(1, 15)
554     tbl.Cell(1, 13).Shape.TextFrame.TextRange.Text = hdrPrem
        ' Col 16 = name placeholder
555     tbl.Cell(1, 16).Shape.TextFrame.TextRange.Text = ""
        On Error GoTo ERR_HANDLER

        ' Header row 2: sub-headers (% / yearVal / refYear) for each category
        ' Note: indexes refer to original (pre-merge) cell positions; after merge, sub-header writes go directly to row 2 cells
560     tbl.Cell(2, 1).Shape.TextFrame.TextRange.Text = pctLabel
561     tbl.Cell(2, 2).Shape.TextFrame.TextRange.Text = yearVal
562     tbl.Cell(2, 3).Shape.TextFrame.TextRange.Text = refYear
563     tbl.Cell(2, 4).Shape.TextFrame.TextRange.Text = pctLabel
564     tbl.Cell(2, 5).Shape.TextFrame.TextRange.Text = yearVal
565     tbl.Cell(2, 6).Shape.TextFrame.TextRange.Text = refYear
566     tbl.Cell(2, 7).Shape.TextFrame.TextRange.Text = pctLabel
567     tbl.Cell(2, 8).Shape.TextFrame.TextRange.Text = yearVal
568     tbl.Cell(2, 9).Shape.TextFrame.TextRange.Text = refYear
569     tbl.Cell(2, 10).Shape.TextFrame.TextRange.Text = pctLabel
570     tbl.Cell(2, 11).Shape.TextFrame.TextRange.Text = yearVal
571     tbl.Cell(2, 12).Shape.TextFrame.TextRange.Text = refYear
572     tbl.Cell(2, 13).Shape.TextFrame.TextRange.Text = pctLabel
573     tbl.Cell(2, 14).Shape.TextFrame.TextRange.Text = yearVal
574     tbl.Cell(2, 15).Shape.TextFrame.TextRange.Text = refYear
575     tbl.Cell(2, 16).Shape.TextFrame.TextRange.Text = ""

        ' --- Format header rows ---
580     For c = 1 To tblCols
582         On Error Resume Next
584         tbl.Cell(1, c).Shape.TextFrame.TextRange.Font.Name = "Arial"
586         tbl.Cell(1, c).Shape.TextFrame.TextRange.Font.Size = 11
588         tbl.Cell(1, c).Shape.TextFrame.TextRange.Font.Bold = True
590         tbl.Cell(1, c).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
592         tbl.Cell(1, c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
594         tbl.Cell(1, c).Shape.Fill.ForeColor.RGB = blueClr
596         tbl.Cell(2, c).Shape.TextFrame.TextRange.Font.Name = "Arial"
598         tbl.Cell(2, c).Shape.TextFrame.TextRange.Font.Size = 10
600         tbl.Cell(2, c).Shape.TextFrame.TextRange.Font.Bold = True
602         tbl.Cell(2, c).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
604         tbl.Cell(2, c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
606         tbl.Cell(2, c).Shape.Fill.ForeColor.RGB = blueClr
608         On Error GoTo ERR_HANDLER
610     Next c

        ' --- Data rows ---
        Dim chgVal As Double
620     For idx = 1 To nItems
630         tblR = idx + 2

            ' Col 16 = name (rightmost)
640         tbl.Cell(tblR, 16).Shape.TextFrame.TextRange.Text = arrNames(idx)

            ' Premium block: cols 13-15 (15=ref, 14=cur, 13=%)
650         tbl.Cell(tblR, 15).Shape.TextFrame.TextRange.Text = Format$(arrPremR(idx), "#,##0")
651         tbl.Cell(tblR, 14).Shape.TextFrame.TextRange.Text = Format$(arrPremC(idx), "#,##0")
652         If arrPremR(idx) <> 0 Then
653             chgVal = (arrPremC(idx) - arrPremR(idx)) / Abs(arrPremR(idx)) * 100
654             tbl.Cell(tblR, 13).Shape.TextFrame.TextRange.Text = Format$(chgVal, "0.0") & "%"
655         Else
656             tbl.Cell(tblR, 13).Shape.TextFrame.TextRange.Text = "-"
657         End If

            ' Docs block: cols 10-12 (12=ref, 11=cur, 10=%)
660         tbl.Cell(tblR, 12).Shape.TextFrame.TextRange.Text = Format$(arrDocR(idx), "#,##0")
661         tbl.Cell(tblR, 11).Shape.TextFrame.TextRange.Text = Format$(arrDocC(idx), "#,##0")
662         If arrDocR(idx) <> 0 Then
663             chgVal = (CDbl(arrDocC(idx)) - CDbl(arrDocR(idx))) / Abs(CDbl(arrDocR(idx))) * 100
664             tbl.Cell(tblR, 10).Shape.TextFrame.TextRange.Text = Format$(chgVal, "0.0") & "%"
665         Else
666             tbl.Cell(tblR, 10).Shape.TextFrame.TextRange.Text = "-"
667         End If

            ' Insured block: cols 7-9 (9=ref, 8=cur, 7=%)
670         tbl.Cell(tblR, 9).Shape.TextFrame.TextRange.Text = Format$(arrInsR(idx), "#,##0")
671         tbl.Cell(tblR, 8).Shape.TextFrame.TextRange.Text = Format$(arrInsC(idx), "#,##0")
672         If arrInsR(idx) <> 0 Then
673             chgVal = (CDbl(arrInsC(idx)) - CDbl(arrInsR(idx))) / Abs(CDbl(arrInsR(idx))) * 100
674             tbl.Cell(tblR, 7).Shape.TextFrame.TextRange.Text = Format$(chgVal, "0.0") & "%"
675         Else
676             tbl.Cell(tblR, 7).Shape.TextFrame.TextRange.Text = "-"
677         End If

            ' Policies block: cols 4-6 (6=ref, 5=cur, 4=%) - v7.42 NEW
680         tbl.Cell(tblR, 6).Shape.TextFrame.TextRange.Text = Format$(arrPolR(idx), "#,##0")
681         tbl.Cell(tblR, 5).Shape.TextFrame.TextRange.Text = Format$(arrPolC(idx), "#,##0")
682         If arrPolR(idx) <> 0 Then
683             chgVal = (CDbl(arrPolC(idx)) - CDbl(arrPolR(idx))) / Abs(CDbl(arrPolR(idx))) * 100
684             tbl.Cell(tblR, 4).Shape.TextFrame.TextRange.Text = Format$(chgVal, "0.0") & "%"
685         Else
686             tbl.Cell(tblR, 4).Shape.TextFrame.TextRange.Text = "-"
687         End If

            ' Commission block: cols 1-3 (3=ref, 2=cur, 1=%)
690         tbl.Cell(tblR, 3).Shape.TextFrame.TextRange.Text = Format$(arrCommR(idx), "#,##0")
691         tbl.Cell(tblR, 2).Shape.TextFrame.TextRange.Text = Format$(arrCommC(idx), "#,##0")
692         If arrCommR(idx) <> 0 Then
693             chgVal = (arrCommC(idx) - arrCommR(idx)) / Abs(arrCommR(idx)) * 100
694             tbl.Cell(tblR, 1).Shape.TextFrame.TextRange.Text = Format$(chgVal, "0.0") & "%"
695         Else
696             tbl.Cell(tblR, 1).Shape.TextFrame.TextRange.Text = "-"
697         End If

            ' Format data row cells
700         For c = 1 To tblCols
710             tbl.Cell(tblR, c).Shape.TextFrame.TextRange.Font.Name = "Arial"
711             tbl.Cell(tblR, c).Shape.TextFrame.TextRange.Font.Size = 9
712             tbl.Cell(tblR, c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
713             If idx Mod 2 = 0 Then
714                 tbl.Cell(tblR, c).Shape.Fill.ForeColor.RGB = lightBlue
715             Else
716                 tbl.Cell(tblR, c).Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
717             End If
720         Next c
730     Next idx

        ' --- Total row ---
        Dim totR As Long
        Dim totPR As Double, totPC As Double, totCR As Double, totCC As Double
        Dim totDR As Long, totDC As Long, totIR As Long, totIC As Long
        Dim totPolR As Long, totPolC As Long
740     totR = nItems + 3
750     totPR = CDbl(ws.Cells(lastRow, 2).Value2)
760     totPC = CDbl(ws.Cells(lastRow, 3).Value2)
770     totDR = CLng(ws.Cells(lastRow, 5).Value2)
780     totDC = CLng(ws.Cells(lastRow, 6).Value2)
790     totIR = CLng(ws.Cells(lastRow, 8).Value2)
800     totIC = CLng(ws.Cells(lastRow, 9).Value2)
802     totPolR = CLng(ws.Cells(lastRow, 11).Value2)  ' v7.42
804     totPolC = CLng(ws.Cells(lastRow, 12).Value2)  ' v7.42
810     totCR = CDbl(ws.Cells(lastRow, 14).Value2)
820     totCC = CDbl(ws.Cells(lastRow, 15).Value2)

        ' Col 16 = "sach hakol"
830     tbl.Cell(totR, 16).Shape.TextFrame.TextRange.Text = ChrW(1505) & ChrW(1492) & Chr(34) & ChrW(1499)
        ' Premium total
840     tbl.Cell(totR, 15).Shape.TextFrame.TextRange.Text = Format$(totPR, "#,##0")
841     tbl.Cell(totR, 14).Shape.TextFrame.TextRange.Text = Format$(totPC, "#,##0")
842     If totPR <> 0 Then
843         tbl.Cell(totR, 13).Shape.TextFrame.TextRange.Text = Format$((totPC - totPR) / Abs(totPR) * 100, "0.0") & "%"
844     Else
845         tbl.Cell(totR, 13).Shape.TextFrame.TextRange.Text = "-"
846     End If
        ' Docs total
850     tbl.Cell(totR, 12).Shape.TextFrame.TextRange.Text = Format$(totDR, "#,##0")
851     tbl.Cell(totR, 11).Shape.TextFrame.TextRange.Text = Format$(totDC, "#,##0")
852     If totDR <> 0 Then
853         tbl.Cell(totR, 10).Shape.TextFrame.TextRange.Text = Format$((CDbl(totDC) - CDbl(totDR)) / Abs(CDbl(totDR)) * 100, "0.0") & "%"
854     Else
855         tbl.Cell(totR, 10).Shape.TextFrame.TextRange.Text = "-"
856     End If
        ' Insured total
860     tbl.Cell(totR, 9).Shape.TextFrame.TextRange.Text = Format$(totIR, "#,##0")
861     tbl.Cell(totR, 8).Shape.TextFrame.TextRange.Text = Format$(totIC, "#,##0")
862     If totIR <> 0 Then
863         tbl.Cell(totR, 7).Shape.TextFrame.TextRange.Text = Format$((CDbl(totIC) - CDbl(totIR)) / Abs(CDbl(totIR)) * 100, "0.0") & "%"
864     Else
865         tbl.Cell(totR, 7).Shape.TextFrame.TextRange.Text = "-"
866     End If
        ' Policies total
870     tbl.Cell(totR, 6).Shape.TextFrame.TextRange.Text = Format$(totPolR, "#,##0")
871     tbl.Cell(totR, 5).Shape.TextFrame.TextRange.Text = Format$(totPolC, "#,##0")
872     If totPolR <> 0 Then
873         tbl.Cell(totR, 4).Shape.TextFrame.TextRange.Text = Format$((CDbl(totPolC) - CDbl(totPolR)) / Abs(CDbl(totPolR)) * 100, "0.0") & "%"
874     Else
875         tbl.Cell(totR, 4).Shape.TextFrame.TextRange.Text = "-"
876     End If
        ' Commission total
880     tbl.Cell(totR, 3).Shape.TextFrame.TextRange.Text = Format$(totCR, "#,##0")
881     tbl.Cell(totR, 2).Shape.TextFrame.TextRange.Text = Format$(totCC, "#,##0")
882     If totCR <> 0 Then
883         tbl.Cell(totR, 1).Shape.TextFrame.TextRange.Text = Format$((totCC - totCR) / Abs(totCR) * 100, "0.0") & "%"
884     Else
885         tbl.Cell(totR, 1).Shape.TextFrame.TextRange.Text = "-"
886     End If

        ' Format total row
890     For c = 1 To tblCols
891         tbl.Cell(totR, c).Shape.TextFrame.TextRange.Font.Name = "Arial"
892         tbl.Cell(totR, c).Shape.TextFrame.TextRange.Font.Size = 10
893         tbl.Cell(totR, c).Shape.TextFrame.TextRange.Font.Bold = True
894         tbl.Cell(totR, c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
895         tbl.Cell(totR, c).Shape.Fill.ForeColor.RGB = RGB(255, 217, 102)
896     Next c

        ' v7.44: re-apply row heights after content written
897     For rowI = 1 To tblRows
898         tbl.Rows(rowI).Height = rowH
899     Next rowI
900     Exit Sub
ERR_HANDLER:
910     Err.Raise Err.Number, "BuildTableSlide(" & sheetName & "):" & Erl, Err.Description
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

        ' Remove any leftover sheet protection
        Dim wsUp4 As Worksheet
        For Each wsUp4 In ThisWorkbook.Worksheets
            On Error Resume Next
            wsUp4.Unprotect SHEET_PROTECT_PWD
            On Error GoTo ERR_HANDLER
        Next wsUp4

20      Dim wsRev As Worksheet
30      Set wsRev = ActiveSheet

        ' Verify we are on a review sheet (name starts with REVIEW_SHEET_NAME)
40      If InStr(1, wsRev.Name, REVIEW_SHEET_NAME(), vbTextCompare) = 0 Then
50          MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1492) & ChrW(1512) & ChrW(1497) & ChrW(1509) & " " & ChrW(1502) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500), vbExclamation
60          Exit Sub
70      End If

        ' Find action column (header contains "peula" = ?????)
80      Dim lastCol As Long
90      lastCol = wsRev.Cells(1, wsRev.Columns.count).End(xlToLeft).Column
100     Dim actionCol As Long
110     actionCol = 0
120     Dim c As Long
130     For c = 1 To lastCol
140         If InStr(1, CStr(wsRev.Cells(1, c).Value2), ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) & ChrW(1492), vbTextCompare) > 0 Then
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
240     lastRow = wsRev.Cells(wsRev.Rows.count, 1).End(xlUp).Row
250     Dim sendCount As Long
260     sendCount = 0
270     Dim r As Long
        ' "ha'aver livdika" = ???? ??????
        Dim actionMatch As String
280     actionMatch = ChrW(1492) & ChrW(1506) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1492)
290     For r = 2 To lastRow
300         If InStr(1, CStr(wsRev.Cells(r, actionCol).Value2), actionMatch, vbTextCompare) > 0 Then
310             sendCount = sendCount + 1
320         End If
330     Next r

340     If sendCount = 0 Then
            ' "ein shurot le'ha'avara livdika" = no rows to transfer for review
350         MsgBoxU ChrW(1488) & ChrW(1497) & ChrW(1503) & " " & ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1492) & ChrW(1506) & ChrW(1489) & ChrW(1512) & ChrW(1492) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1492), vbInformation
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

        ' Copy header row
500     Dim hdrCol As Long
510     For hdrCol = 1 To lastCol
520         wsTemp.Cells(1, hdrCol).Value = wsRev.Cells(1, hdrCol).Value2  ' 7.33: was .Value
530     Next hdrCol
540     wsTemp.Rows(1).Font.Bold = True

        ' Copy matching rows
550     Dim outRow As Long
560     outRow = 2
570     For r = 2 To lastRow
580         If InStr(1, CStr(wsRev.Cells(r, actionCol).Value2), actionMatch, vbTextCompare) > 0 Then
590             For hdrCol = 1 To lastCol
600                 wsTemp.Cells(outRow, hdrCol).Value = wsRev.Cells(r, hdrCol).Value2  ' 7.33: was .Value
610             Next hdrCol
620             outRow = outRow + 1
630         End If
640     Next r
650     wsTemp.Columns.AutoFit

        ' Save temp file
660     Dim tempPath As String
670     tempPath = ThisWorkbook.Path & "\" & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501) & "_" & wsRev.Name & ".xlsx"
680     Application.DisplayAlerts = False
690     wbTemp.SaveAs tempPath, xlOpenXMLWorkbook
700     wbTemp.Close SaveChanges:=False
710     Application.DisplayAlerts = True

        ' Build email body
        ' "hi lahav" = ?? ???
720     Dim bodyLine1 As String
730     bodyLine1 = ChrW(1492) & ChrW(1497) & " " & ChrW(1502) & ChrW(1488) & ChrW(1497) & ChrW(1512)  ' v7.55: was Lehav, now Meir
        ' "likrat hachanat doch avurchem nimtze'u hachrigim haram" = ????? ???? ??? ?????? ????? ??????? ??"?
740     Dim bodyLine2 As String
750     bodyLine2 = ChrW(1500) & ChrW(1511) & ChrW(1512) & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1499) & ChrW(1504) & ChrW(1514) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & " " & ChrW(1506) & ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1499) & ChrW(1501) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & ChrW(1492) & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501) & " " & ChrW(1492) & ChrW(1512) & ChrW(34) & ChrW(1502)
        ' "al mnat lehafik et hadoch ani mevakeshet tguvatcha al mnat she'etaken beheta'am"
760     Dim bodyLine3 As String
770     bodyLine3 = ChrW(1506) & ChrW(1500) & " " & ChrW(1502) & ChrW(1504) & ChrW(1514) & " " & ChrW(1500) & ChrW(1492) & ChrW(1508) & ChrW(1497) & ChrW(1511) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1491) & ChrW(1493) & ChrW(1495) & " " & ChrW(1488) & ChrW(1504) & ChrW(1497) & " " & ChrW(1502) & ChrW(1489) & ChrW(1511) & ChrW(1513) & ChrW(1514) & " " & ChrW(1514) & ChrW(1490) & ChrW(1493) & ChrW(1489) & ChrW(1514) & ChrW(1498) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1502) & ChrW(1504) & ChrW(1514) & " " & ChrW(1513) & ChrW(1488) & ChrW(1514) & ChrW(1511) & ChrW(1503) & " " & ChrW(1489) & ChrW(1492) & ChrW(1514) & ChrW(1488) & ChrW(1501)
        ' toda = ????
780     Dim bodyLine4 As String
790     bodyLine4 = ChrW(1514) & ChrW(1493) & ChrW(1491) & ChrW(1492)
        ' orit = ?????
800     Dim bodyLine5 As String
810     bodyLine5 = ChrW(1510) & ChrW(1489) & ChrW(1497)  ' v7.55: was Orit, now Zvi

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

        ' Success message: "email huchan be'hatzlacha im X shurot" = ???? ???? ?????? ?? X ?????
        ' v7.56: ask whether the user actually clicked Send in Outlook;
        ' if not, show a follow-up that waits for them to do it.
        ' v7.57: position the popup to the upper-left so it does not
        ' cover the Outlook compose window centred on screen.
        PositionNextMsgBox 30, 80
        Dim sentRes As Long
950     sentRes = MsgBoxU(ChrW(1504) & ChrW(1493) & ChrW(1510) & ChrW(1512) & " " & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & " " & ChrW(1506) & ChrW(1501) & " " & sendCount & " " & ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & "." & vbCrLf & vbCrLf & ChrW(1492) & ChrW(1488) & ChrW(1501) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & ChrW(1514) & "?", vbYesNo + vbQuestion)
        If sentRes = vbNo Then
            MsgBoxU ChrW(1513) & ChrW(1500) & ChrW(1495) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & " " & ChrW(1502) & ChrW(1495) & ChrW(1500) & ChrW(1493) & ChrW(1503) & " Outlook " & ChrW(1493) & ChrW(1500) & ChrW(1495) & ChrW(1509) & " " & ChrW(1488) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1512) & " " & ChrW(1500) & ChrW(1488) & ChrW(1495) & ChrW(1512) & " " & ChrW(1492) & ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1495) & ChrW(1492), vbInformation
        End If

        ' Ask: "tipalt bekhol hahrigim? avar lehamshekh hahafaka?"
        ' ????? ??? ???????? ???? ????? ?????
        Dim askRes As Long
        askRes = MsgBoxU( _
            ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1500) & ChrW(1514) & " " & ChrW(1489) & ChrW(1499) & ChrW(1500) & " " & ChrW(1492) & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501) & "?" & vbCrLf & _
            ChrW(1506) & ChrW(1489) & ChrW(1493) & ChrW(1512) & " " & ChrW(1500) & ChrW(1492) & ChrW(1502) & ChrW(1513) & ChrW(1498) & " " & ChrW(1492) & ChrW(1492) & ChrW(1508) & ChrW(1511) & ChrW(1492), _
            vbYesNo + vbQuestion)
        If askRes = vbYes Then
            ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Activate
        End If

960     Exit Sub

ERR_HANDLER:
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
    HideRibbon   ' v7.64: hide ribbon together with sheets
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
    ShowRibbon   ' v7.64: show ribbon together with sheets
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
        If pwd <> SHEET_PROTECT_PWD Then
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
    ' 7.36: layout assumes data tables start at row 40 (rows 1-39 are free
    '       for buttons). Buttons sit in cols A:D so K4-K8 (the new
    '       parameter cells) stay visible on the right.
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

    Dim shp As Shape
    Dim rng As Range

    ' ---- Title bar (cols A:H, rows 2-3) ----
    Set rng = wsMgmt.Range("A2:H3")
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    shp.Name = "navTitle"
    shp.Fill.ForeColor.RGB = RGB(0, 70, 130)
    shp.Line.Visible = msoFalse
    ' Title: "hagdarot ma'arechet <ENGLISH_NAME>"
    shp.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514) & " " & ENGLISH_NAME()
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 18
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

    ' ---- Button 1: Milon Anafim (Branch Dictionary) -> A40 ----
    Set rng = wsMgmt.Range("A5:D7")
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    shp.Name = "navBranch"
    shp.Fill.ForeColor.RGB = RGB(0, 120, 60)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = ChrW(1502) & ChrW(1497) & ChrW(1500) & ChrW(1493) & ChrW(1503) & " " & ChrW(1506) & ChrW(1504) & ChrW(1508) & ChrW(1497) & ChrW(1501)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 13
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavToBranch"

    ' ---- Button 2: Index Makor (Source Index) -> E40 ----
    Set rng = wsMgmt.Range("A9:D11")
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    shp.Name = "navIndex"
    shp.Fill.ForeColor.RGB = RGB(0, 100, 170)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = ChrW(1488) & ChrW(1497) & ChrW(1504) & ChrW(1491) & ChrW(1511) & ChrW(1505) & " " & ChrW(1502) & ChrW(1511) & ChrW(1493) & ChrW(1512)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 13
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavToIndex"

    ' ---- Button 3: Parametrim (Parameters) -> J40 ----
    Set rng = wsMgmt.Range("A13:D15")
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    shp.Name = "navParams"
    shp.Fill.ForeColor.RGB = RGB(160, 80, 0)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512) & ChrW(1497) & ChrW(1501)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 13
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavToParams"

    ' ---- Button 4: Reshimat Shgiot (Helper translations) -> N40 ----
    Set rng = wsMgmt.Range("A17:D19")
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    shp.Name = "navErrors"
    shp.Fill.ForeColor.RGB = RGB(180, 30, 30)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = ChrW(1512) & ChrW(1513) & ChrW(1497) & ChrW(1502) & ChrW(1514) & " " & ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 13
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavToErrors"

    ' ---- Button 5 (NEW): Chazara Lerosh / Back to top -> A1 ----
    Set rng = wsMgmt.Range("A21:D23")
    Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    shp.Name = "navTop"
    shp.Fill.ForeColor.RGB = RGB(80, 80, 80)
    shp.Line.Visible = msoFalse
    ' "chazara lerosh" = back to top
    shp.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1494) & ChrW(1512) & ChrW(1492) & " " & ChrW(1500) & ChrW(1512) & ChrW(1488) & ChrW(1513)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 13
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "NavToTop"

    MsgBoxU ChrW(1491) & ChrW(1507) & " " & ChrW(1513) & ChrW(1506) & ChrW(1512) & " " & ChrW(1492) & ChrW(1493) & ChrW(1490) & ChrW(1491) & ChrW(1512) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492), vbInformation
    Exit Sub
ERR_HANDLER:
    MsgBoxU ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1514) & " " & ChrW(1491) & ChrW(1507) & " " & ChrW(1513) & ChrW(1506) & ChrW(1512) & ":" & vbCrLf & Err.Description, vbCritical
End Sub

' ---- Navigation macros for settings sheet buttons ----
Public Sub NavToBranch()
    ' 7.36: data tables moved to row 40
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    ws.Activate
    ws.Range("A40").Select
    Application.Goto ws.Range("A40"), True
End Sub

Public Sub NavToIndex()
    ' 7.36: data tables moved to row 40
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    ws.Activate
    ws.Range("E40").Select
    Application.Goto ws.Range("E40"), True
End Sub

Public Sub NavToParams()
    ' 7.36: data tables moved to row 40
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    ws.Activate
    ws.Range("J40").Select
    Application.Goto ws.Range("J40"), True
End Sub

Public Sub NavToErrors()
    ' 7.36: data tables moved to row 40
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    ws.Activate
    ws.Range("N40").Select
    Application.Goto ws.Range("N40"), True
End Sub

Public Sub NavToTop()
    ' 7.36: new - return to top of settings sheet (row 1)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
    ws.Activate
    ws.Range("A1").Select
    Application.Goto ws.Range("A1"), True
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
    
    ' v7.50: always delete-and-recreate so a previously-hidden sheet does
    ' not silently swallow the Activate call (which made the green Search
    ' button appear unresponsive).
    Dim wsSearch As Worksheet
    On Error Resume Next
    Set wsSearch = ThisWorkbook.Worksheets(SEARCH_SHEET_NAME)
    If Not wsSearch Is Nothing Then
        wsSearch.Visible = xlSheetVisible
        Application.DisplayAlerts = False
        wsSearch.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo ERR_HANDLER

    Set wsSearch = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    wsSearch.Name = SEARCH_SHEET_NAME
    wsSearch.Visible = xlSheetVisible
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
    
    ' v7.65: force screen refresh so shapes + content appear on first show
    ' (without this, user sometimes sees an empty sheet until switching tabs)
    Application.ScreenUpdating = True
    wsSearch.Activate
    DoEvents
    Application.Goto wsSearch.Cells(2, 1), True
    wsSearch.Cells(2, 1).Select
    
    Exit Sub
ERR_HANDLER:
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
        wsSearch.Cells(2, 1).Select
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
    
    ' v7.41: collect candidate customer names from all base sheets first,
    ' then do word-boundary match. Fall back to substring if no exact matches.
    Dim allNames As Object: Set allNames = CreateObject("Scripting.Dictionary")
    allNames.CompareMode = vbTextCompare
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505), vbTextCompare) > 0 Then
            foundBase = True
            lastRow = ws.Cells(ws.Rows.count, 6).End(xlUp).Row
            For r = 2 To lastRow
                cName = Trim$(CStr(ws.Cells(r, 6).Value2))
                If cName <> "" Then
                    If Not allNames.Exists(cName) Then allNames.Add cName, 1
                End If
            Next r
        End If
    Next ws

    ' Pass 1: word-boundary match (exact token)
    Dim nameKey As Variant, parts() As String, k As Long
    For Each nameKey In allNames.keys
        cName = CStr(nameKey)
        ' Check if cName exactly equals searchText
        If StrComp(cName, searchText, vbTextCompare) = 0 Then
            If Not dict.Exists(cName) Then dict.Add cName, 1
        Else
            ' Check if any space-separated token equals searchText
            parts = Split(cName, " ")
            For k = 0 To UBound(parts)
                If StrComp(Trim$(parts(k)), searchText, vbTextCompare) = 0 Then
                    If Not dict.Exists(cName) Then dict.Add cName, 1
                    Exit For
                End If
            Next k
        End If
    Next nameKey

    ' Pass 2: if no word-boundary matches, fall back to substring
    If dict.count = 0 Then
        For Each nameKey In allNames.keys
            cName = CStr(nameKey)
            If InStr(1, cName, searchText, vbTextCompare) > 0 Then
                If Not dict.Exists(cName) Then dict.Add cName, 1
            End If
        Next nameKey
    End If
    
    If Not foundBase Then
        MsgBoxU ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505), vbExclamation
        Exit Sub
    End If
    
    If dict.count = 0 Then
        MsgBoxU ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1502) & ChrW(1514) & ChrW(1488) & ChrW(1497) & ChrW(1502) & ChrW(1497) & ChrW(1501), vbInformation  ' "?? ????? ?????? ???????"
        wsSearch.Cells(2, 1).Select
        Exit Sub
    End If
    
    ' If only 1 result, select it directly
    If dict.count = 1 Then
        Dim wsMain As Worksheet
        Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        wsMain.Range("G12").Value = dict.keys()(0)
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
    clearLastRow = wsSearch.Cells(wsSearch.Rows.count, 1).End(xlUp).Row
    If clearLastRow >= 4 Then wsSearch.Range(wsSearch.Cells(4, 1), wsSearch.Cells(clearLastRow, 1)).Clear
    
    ' Write header with result count in row 3
    wsSearch.Cells(3, 1).Value = ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & dict.count & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " - " & ChrW(1500) & ChrW(1495) & ChrW(1509) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1493) & ChrW(1500) & ChrW(1495) & ChrW(1509) & " " & ChrW(1489) & ChrW(1495) & ChrW(1512)  ' "????? X ?????? - ??? ?? ?? ???? ???"
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
    wsSearch.Cells(4, 1).Select
    
    Exit Sub
ERR_HANDLER:
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
    wsMain.Range("G12").Value = selectedName
    
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
    
    ' 7.33: only attempt delete if sheet exists
    If SheetExists(SEARCH_SHEET_NAME) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(SEARCH_SHEET_NAME).Delete
        Application.DisplayAlerts = True
    End If
    
    If SheetExists(CONTROL_SHEET_NAME()) Then
        Dim wsMain As Worksheet
        Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
        wsMain.Activate
    End If
End Sub


' ============================================================================
' PUBLIC: Clear client filter - resets G12 to default
' ============================================================================
Public Sub ClearClientFilter()
    On Error Resume Next
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    wsMain.Range("G12").Value = TXT_BACHAR_I()  ' 7.33: was inline ChrW chain
End Sub


' --- 7.33: removed dead function ConvertEngToHeb ---

' ============================================================================
' RESET CLIENT FILTER - sets G12 back to "bachar/i" for full reports
' ============================================================================
Public Sub ResetClientFilter()
    ' 7.33: existence guard + helper
    If Not SheetExists(CONTROL_SHEET_NAME()) Then Exit Sub
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    wsMain.Range("G12").Value = TXT_BACHAR_I()  ' 7.33: was inline ChrW chain
End Sub

' ============================================================================
' PUBLIC (7.40): ClearHomePageSelection - reset G5-G12 to defaults
'   sug tkupa = shnati (yearly)
'   period value = empty
'   date type = bordereu
'   filter type = bechar/i
'   filter value = bechar/i
'   client name = bechar/i
' ============================================================================
Public Sub ClearHomePageSelection()
    On Error Resume Next
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    If wsMain Is Nothing Then Exit Sub

    ' G5: sug tkupa = shnati
    wsMain.Range("G5").Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497)
    ' G6: period value = empty
    wsMain.Range("G6").Value = ""
    ' G7: date type = bordereu
    wsMain.Range("G7").Value = ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1512) & ChrW(1493)
    ' G9: filter type = bechar/i
    wsMain.Range("G9").Value = TXT_BACHAR_I()
    ' G10: filter value = bechar/i
    wsMain.Range("G10").Value = TXT_BACHAR_I()
    ' G12: client name = bechar/i
    wsMain.Range("G12").Value = TXT_BACHAR_I()

    ' Refresh dependent dropdowns
    On Error Resume Next
    UpdatePeriodDropdown
    On Error GoTo 0
End Sub


' ============================================================================
' PUBLIC (7.50): RebuildHomeButtons - re-creates the 3 client-row buttons
' on the home page (Search / All / Reset parameters).
' Use this when an existing workbook is missing one of these buttons,
' so you don't have to re-run the full SetupMainSheet.
' ============================================================================
Public Sub RebuildHomeButtons()
    On Error GoTo ERR_HANDLER
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    If wsMain Is Nothing Then
        MsgBoxU "Home sheet not found.", vbExclamation
        Exit Sub
    End If

    Dim names As Variant: names = Array("btnSearchClient", "btnAllClients", "btnClearSelection")
    Dim nm As Variant
    For Each nm In names
        On Error Resume Next
        wsMain.Shapes(CStr(nm)).Delete
        On Error GoTo ERR_HANDLER
    Next nm

    Dim baseLeft As Double, baseTop As Double, h As Double
    baseLeft = wsMain.Range("H12").Left
    baseTop = wsMain.Range("H12").Top + 1
    h = wsMain.Range("H12").Height - 2

    ' Search button (chapas)
    Dim sh As Shape
    Set sh = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, baseLeft + 2, baseTop, 50, h)
    sh.Name = "btnSearchClient"
    sh.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513)
    sh.TextFrame2.TextRange.Font.Size = 9
    sh.TextFrame2.TextRange.Font.Bold = msoTrue
    sh.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    sh.Fill.ForeColor.RGB = RGB(70, 130, 180)
    sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    sh.OnAction = "SearchClientName"

    ' All button (kulam) - resets G12 to bechar/i
    Set sh = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, baseLeft + 55, baseTop, 50, h)
    sh.Name = "btnAllClients"
    sh.TextFrame2.TextRange.Text = ChrW(1499) & ChrW(1493) & ChrW(1500) & ChrW(1501)
    sh.TextFrame2.TextRange.Font.Size = 9
    sh.TextFrame2.TextRange.Font.Bold = msoTrue
    sh.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    sh.Fill.ForeColor.RGB = RGB(60, 160, 60)
    sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    sh.OnAction = "ResetClientFilter"

    ' Reset parameters button (ipus parametrim)
    Set sh = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, baseLeft + 110, baseTop, 130, h)
    sh.Name = "btnClearSelection"
    sh.TextFrame2.TextRange.Text = ChrW(1488) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1505) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512) & ChrW(1497) & ChrW(1501)
    sh.TextFrame2.TextRange.Font.Size = 9
    sh.TextFrame2.TextRange.Font.Bold = msoTrue
    sh.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    sh.Fill.ForeColor.RGB = RGB(160, 80, 0)
    sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    sh.OnAction = "ClearHomePageSelection"

    MsgBoxU ChrW(1492) & ChrW(1499) & ChrW(1508) & ChrW(1514) & ChrW(1493) & ChrW(1512) & ChrW(1497) & ChrW(1501) & " " & ChrW(1513) & ChrW(1493) & ChrW(1495) & ChrW(1494) & ChrW(1512) & ChrW(1493)  '' "Buttons restored"
    Exit Sub
ERR_HANDLER:
    MsgBoxU "Error: " & Err.Description, vbCritical
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
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    ' AutoFit column A to fit text, ensure minimum width of 20
    ws.Columns(1).AutoFit
    If ws.Columns(1).ColumnWidth < 20 Then ws.Columns(1).ColumnWidth = 20

    ' Add thin light-gray borders to the data content area (7.33: collapsed)
    Dim dataRng As Range
    Set dataRng = ws.Range(ws.Cells(dataStart, 1), ws.Cells(lastRow, 16))
    With dataRng.Borders
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = RGB(180, 180, 180)
    End With

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
            ' v7.60: bold gold matching the presentation total row
            ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 16)).Interior.Color = RGB(255, 217, 102)
            ' Font: black for guaranteed visibility on any color
            ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 16)).Font.Color = RGB(0, 0, 0)
            ws.Range(ws.Cells(lastRow, 1), ws.Cells(lastRow, 16)).Font.Bold = True
        End If
    End If

    ' Ensure RTL display
    ws.DisplayRightToLeft = True


    On Error GoTo 0
End Sub


' ============================================================================
' PUBLIC (7.37): EmbedSourceSheets
' Scans a folder for YYYY.xlsx files and embeds each one as a hidden
' internal sheet named "__src_<year>". After running, the workbook is
' self-contained for the embedded years -- BuildReview and ApplyCorrections
' will read from the internal sheets and never touch the external files.
'
' Use this when you want to ship a portable demo workbook to someone who
' should not need to install or copy any source folders.
' ============================================================================
Public Sub EmbedSourceSheets()
        On Error GoTo ERR_HANDLER

        Dim folder As String
10      folder = SOURCE_FOLDER()
20      folder = InputBox("Source folder containing YYYY.xlsx files:" & vbCrLf & _
                "Each YYYY.xlsx will be embedded as a hidden sheet named __src_YYYY", _
                "EmbedSourceSheets", folder)
30      If folder = "" Then Exit Sub
40      If Right$(folder, 1) <> "\" Then folder = folder & "\"

        Dim fso As Object
50      Set fso = CreateObject("Scripting.FileSystemObject")
60      If Not fso.FolderExists(folder) Then
70          MsgBoxU "Folder not found: " & folder, vbCritical
80          Exit Sub
90      End If

        ' --- Find YYYY.xlsx / YYYY.xls files ---
        Dim foundYears As Object
100     Set foundYears = CreateObject("Scripting.Dictionary")

        Dim file As Object, baseName As String, ext As String
110     For Each file In fso.GetFolder(folder).Files
120         baseName = fso.GetBaseName(file.Name)
130         ext = LCase$(fso.GetExtensionName(file.Name))
140         If (ext = "xlsx" Or ext = "xls") And Len(baseName) = 4 And IsNumeric(baseName) Then
150             If Not foundYears.Exists(baseName) Then foundYears.Add baseName, True
160         End If
170     Next file

180     If foundYears.count = 0 Then
190         MsgBoxU "No YYYY.xlsx files in: " & folder, vbExclamation
200         Exit Sub
210     End If

        ' --- Confirm with user ---
        Dim listMsg As String, yKey As Variant
220     For Each yKey In foundYears.keys
230         listMsg = listMsg & "  " & yKey & vbCrLf
240     Next yKey
250     If MsgBox("Embed these years as hidden internal sheets?" & vbCrLf & vbCrLf & listMsg, _
                vbYesNo + vbQuestion, "EmbedSourceSheets") <> vbYes Then Exit Sub

        ' --- App state ---
        Dim prevSU As Boolean, prevDA As Boolean, prevCalc As XlCalculation
260     prevSU = Application.ScreenUpdating
270     prevDA = Application.DisplayAlerts
280     prevCalc = Application.Calculation
290     Application.ScreenUpdating = False
300     Application.DisplayAlerts = False
310     Application.Calculation = xlCalculationManual

        ' --- Embed each ---
        Dim count As Long
320     count = 0
330     For Each yKey In foundYears.keys
340         EmbedOneYear CStr(yKey), folder
350         count = count + 1
360     Next yKey

370     Application.ScreenUpdating = prevSU
380     Application.DisplayAlerts = prevDA
390     Application.Calculation = prevCalc

400     MsgBoxU "Embedded " & count & " source files as hidden internal sheets." & vbCrLf & vbCrLf & _
                listMsg & vbCrLf & _
                "The workbook is now self-contained for these years.", vbInformation
410     Exit Sub

ERR_HANDLER:
500     On Error Resume Next
510     Application.ScreenUpdating = True
520     Application.DisplayAlerts = True
530     Application.Calculation = xlCalculationAutomatic
540     MsgBoxU "Error in EmbedSourceSheets:" & vbCrLf & Err.Description, vbCritical
End Sub

' Helper for EmbedSourceSheets - embed a single YYYY file
Private Sub EmbedOneYear(ByVal yearVal As String, ByVal folder As String)
        Dim internalName As String
10      internalName = "__src_" & yearVal

        ' Delete existing internal sheet if present
20      DeleteSheetIfExists internalName

        ' Locate external file
        Dim fso As Object
30      Set fso = CreateObject("Scripting.FileSystemObject")
        Dim srcPath As String
40      srcPath = folder & yearVal & ".xlsx"
50      If Not fso.FileExists(srcPath) Then srcPath = folder & yearVal & ".xls"
60      If Not fso.FileExists(srcPath) Then Exit Sub

        ' Open source
        Dim wbSrc As Workbook
70      Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True, UpdateLinks:=0)
        Dim wsSrc As Worksheet
80      Set wsSrc = OpenDataSheet(wbSrc)

        ' Copy the data sheet into ThisWorkbook
90      wsSrc.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
        Dim wsNew As Worksheet
100     Set wsNew = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
110     wsNew.Name = internalName
120     wsNew.Visible = xlSheetVeryHidden

130     wbSrc.Close SaveChanges:=False
End Sub


' ============================================================================
' v7.58: moved here from earlier in the file so module-level declarations
' come before any Sub/Function (fixes 'Only comments may appear after
' End Sub' compile error).
' ============================================================================
#If VBA7 Then
Private Function CBTProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Private Function CBTProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
    On Error Resume Next
    If nCode = HCBT_ACTIVATE Then
        SetWindowPos wParam, 0, mHookX, mHookY, 0, 0, SWP_NOSIZE Or SWP_NOZORDER
        UnhookWindowsHookEx mHook
        mHook = 0
    End If
    CBTProc = CallNextHookEx(mHook, nCode, wParam, lParam)
End Function

' v7.57: install a one-shot CBT hook that moves the next MsgBox to (X, Y).
' The hook auto-uninstalls itself on first activation.
Private Sub PositionNextMsgBox(ByVal X As Long, ByVal Y As Long)
    On Error Resume Next
    mHookX = X
    mHookY = Y
    mHook = SetWindowsHookEx(5, AddressOf CBTProc, 0, GetCurrentThreadId())  ' WH_CBT = 5
End Sub

' v7.57: hide / show the Excel Ribbon
Public Sub HideRibbon()
    On Error Resume Next
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",False)"
End Sub
Public Sub ShowRibbon()
    On Error Resume Next
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",True)"
End Sub

' ============================================================================
' PUBLIC (7.71): AutoFitHomeToScreen - Adjusts zoom to fit the dashboard to screen
' ============================================================================
Public Sub AutoFitHomeToScreen()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    If ws Is Nothing Then Exit Sub
    
    ws.Activate
    ' Select columns A through T to include the balanced left and right margins
    ws.Columns("A:T").Select
    ActiveWindow.Zoom = True
    
    ' Return focus to a safe cell
    ws.Range("D1").Select
End Sub

' ============================================================================
' PUBLIC: Macro for buttons that are only active in the full version
' ============================================================================
Public Sub ProVersionOnly()
    MsgBox ChrW(1508) & ChrW(1506) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & ChrW(1490) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " " & ChrW(1492) & ChrW(1502) & ChrW(1500) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1500) & ChrW(1489) & ChrW(1491), vbInformation, ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514)
End Sub


' ============================================================================
' MACRO: ApplyDemoRestrictions
' Instantly updates the UI to show 2024/2025, blocks typing, and blocks clicks
' ============================================================================
Public Sub ApplyDemoRestrictions()
    On Error GoTo ERR_HANDLER
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets(ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514))
    
    ' Set the values
    wsMain.Range("G3").Value = "2024"
    wsMain.Range("G4").Value = "2025"
    
    ' Block typing via Data Validation
    With wsMain.Range("G3:G4")
        .Validation.Delete
        .Validation.Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="=FALSE"
        .Validation.ShowError = True
        .Validation.ErrorMessage = ChrW(1508) & ChrW(1506) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & ChrW(1490) & ChrW(1512) & ChrW(1505) & ChrW(1492) & " " & ChrW(1492) & ChrW(1502) & ChrW(1500) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & ChrW(1500) & ChrW(1489) & ChrW(1491)
    End With
    
    ' Block clicking via transparent shape
    On Error Resume Next
    wsMain.Shapes("shpProYears").Delete
    On Error GoTo ERR_HANDLER
    
    Dim rngYearsBlock As Range
    Set rngYearsBlock = wsMain.Range("G3:G4")
    Dim shpYears As Shape
    Set shpYears = wsMain.Shapes.AddShape(msoShapeRectangle, rngYearsBlock.Left, rngYearsBlock.Top, rngYearsBlock.Width, rngYearsBlock.Height)
    shpYears.Name = "shpProYears"
    shpYears.Fill.Visible = msoFalse
    shpYears.Line.Visible = msoFalse
    shpYears.OnAction = "ProVersionOnly"
    
    MsgBox "UI updated to Demo restrictions (2024/2025).", vbInformation
    Exit Sub
    
ERR_HANDLER:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
