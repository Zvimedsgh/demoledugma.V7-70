Attribute VB_Name = "ModLLevav"
'Attribute VB_Name = "modLevav"
' ============================================================================
' MODULE: modLevav         VERSION: v1.0 (LevavClaude)
' Built 2026-05-02 — based on Manus v6.8 + Option A workflow
'
' Architecture (Option A):
'   - 2025_fixed = the living "source of truth" sheet inside the workbook
'   - BuildReview: builds 2025_fixed from SOURCE if missing, then scans it
'   - SendForReview ("סיימתי לעדכן"):
'       1. Rebuilds 2025_fixed fresh from SOURCE
'       2. Applies "תקן" rows ? writes fix to the right column
'       3. Applies "התעלם" rows ? marks _IGNORED=1
'       4. Sends Outlook email for "העבר לבדיקה" rows
'       5. Leaves the לטיפול sheet untouched
'   - ApplyCorrectionsAndBuildReports: reads from 2025_fixed (skip _IGNORED=1),
'     builds comparison sheets
'
' Parameters on דף הבית:
'   G3 = שנת בסיס (refYear)
'   G4 = שנה נוכחית (yearVal)
'   G5 = סוג תקופה (period type)
'   G6 = ערך תקופה (period detail)
'   G7 = סוג תאריך (date type — bordereu / insurance start)
'
' Buttons on דף הבית (created by SetupMainSheet):
'   B3:C4 = כפתור 1 (BuildReview)         — blue
'   B5:C6 = כפתור 2 (ApplyCorrections...)  — green
'   B7:C8 = כפתור 3 (BuildPresentation)    — brown
' ============================================================================

#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

' --- Settings sheet column indexes ---
Private Const MANAGEMENT_START_ROW As Long = 2
Private Const COL_FIELD_NAME_HE As Long = 5
Private Const COL_FIELD_COLUMN As Long = 6
Private Const COL_FIELD_CHECKING As Long = 7
Private Const COL_FIELD_KEY As Long = 8
Private Const COL_PARAM_NAME As Long = 10
Private Const COL_PARAM_VALUE As Long = 11
Private Const COL_HELPER_KEY As Long = 14
Private Const COL_HELPER_VALUE As Long = 15

' --- Parameter names ---
Private Const PARAM_PREMIUM_THRESHOLD As String = "PREMIUM_THRESHOLD"
Private Const PARAM_ERROR_EMAIL As String = "ERROR_EMAIL"
Private Const KEY_BRANCH_NAME As String = "BRANCH_NAME"
Private Const KEY_PREMIUM As String = "PREMIUM"
Private Const HELPER_REVIEW_SOURCE_ROW_HEADER As String = "REVIEW_SOURCE_ROW_HEADER"
Private Const HELPER_REVIEW_REASON_HEADER As String = "REVIEW_REASON_HEADER"
Private Const HELPER_REVIEW_REASON_CODE_HEADER As String = "REVIEW_REASON_CODE_HEADER"

' --- Source workbook layout (TmpClientPolicyListEx) ---
Private Const DATA_SHEET_NAME As String = "TmpClientPolicyListEx"
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
Private Const RAW_ACTIONCOL As Long = 39
Private Const RAW_IDNUMBER As Long = 45

' --- base_YEAR sheet column layout (used by ApplyCorrections + reports) ---
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

' --- _IGNORED column header text written into 2025_fixed (column 54 by default) ---
Private Const IGNORED_HEADER As String = "_IGNORED"

' --- Unicode MsgBox flags (RTL Hebrew alignment) ---
Private Const MB_RTLREADING As Long = &H100000
Private Const MB_RIGHT As Long = &H80000

' ============================================================================
' Unicode MessageBox wrapper (Hebrew RTL)
' ============================================================================
Private Function MsgBoxU(ByVal sText As String, Optional ByVal uType As Long = 0, Optional ByVal sCaption As String = "") As Long
    MsgBoxU = MessageBoxW(0, StrPtr(sText), StrPtr(sCaption), uType Or MB_RTLREADING Or MB_RIGHT)
End Function

' ============================================================================
' LOW-LEVEL HELPERS
' ============================================================================

' Find SOURCE/yearVal.xlsx (or .xls) — returns full path or "" if not found
Private Function FindSourceFile(ByVal yearVal As String) As String
    Dim p As String, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    p = SOURCE_FOLDER() & yearVal & ".xlsx"
    If fso.FileExists(p) Then
        FindSourceFile = p
        Exit Function
    End If
    p = SOURCE_FOLDER() & yearVal & ".xls"
    If fso.FileExists(p) Then
        FindSourceFile = p
        Exit Function
    End If
    FindSourceFile = ""
End Function

' Open the data worksheet inside a source workbook (TmpClientPolicyListEx or first)
Private Function OpenDataSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(DATA_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = wb.Worksheets(1)
    Set OpenDataSheet = ws
End Function

' Excel column letter (e.g. "AB") ? number (28)
Private Function ColumnLetterToNumber(ByVal col As String) As Long
    Dim i As Long, ch As String
    For i = 1 To Len(col)
        ch = Mid$(col, i, 1)
        If ch < "A" Or ch > "Z" Then
            Err.Raise vbObjectError + 4000, , "Invalid column letter: " & col
        End If
        ColumnLetterToNumber = ColumnLetterToNumber * 26 + (Asc(ch) - 64)
    Next i
End Function

Private Function SheetExists(ByVal sheetName As String) As Boolean
    On Error GoTo NOT_FOUND
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = True
    Exit Function
NOT_FOUND:
    SheetExists = False
End Function

Private Sub DeleteSheetIfExists(ByVal sName As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    If SheetExists(sName) Then ThisWorkbook.Worksheets(sName).Delete
    Application.DisplayAlerts = True
End Sub

Private Function IsBlankValue(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsBlankValue = True
    ElseIf IsNull(v) Then
        IsBlankValue = True
    ElseIf VarType(v) = vbString Then
        IsBlankValue = (Trim$(CStr(v)) = "")
    Else
        IsBlankValue = False
    End If
End Function

Private Function TryParseVariantNumber(ByVal v As Variant, ByRef result As Double) As Boolean
    On Error GoTo FAIL
    If IsNumeric(v) Then
        result = CDbl(v)
        TryParseVariantNumber = True
    Else
        TryParseVariantNumber = False
    End If
    Exit Function
FAIL:
    TryParseVariantNumber = False
End Function

Private Function IsIgnorableRow(ByVal ws As Worksheet, ByVal r As Long, ByRef keys() As String, ByRef cols() As Long, ByVal cnt As Long, ByVal dictFieldCol As Object) As Boolean
    Dim allBlank As Boolean, i As Long
    allBlank = True
    For i = 1 To cnt
        If Not IsBlankValue(ws.Cells(r, cols(i)).Value2) Then
            allBlank = False
            Exit For
        End If
    Next i
    IsIgnorableRow = allBlank
End Function

' Find the column index of the _IGNORED marker in 2025_fixed (0 if not present)
Private Function FindIgnoredColumn(ByVal wsFixed As Worksheet) As Long
    Dim lastCol As Long, j As Long
    lastCol = wsFixed.Cells(1, wsFixed.Columns.Count).End(xlToLeft).Column
    For j = 1 To lastCol
        If StrComp(Trim$(CStr(wsFixed.Cells(1, j).Value2)), IGNORED_HEADER, vbTextCompare) = 0 Then
            FindIgnoredColumn = j
            Exit Function
        End If
    Next j
    FindIgnoredColumn = 0
End Function

' Hebrew month name (1..12)
Private Function HebrewMonthName(ByVal m As Long) As String
    Select Case m
        Case 1:  HebrewMonthName = ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1488) & ChrW(1512)
        Case 2:  HebrewMonthName = ChrW(1508) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1488) & ChrW(1512)
        Case 3:  HebrewMonthName = ChrW(1502) & ChrW(1512) & ChrW(1509)
        Case 4:  HebrewMonthName = ChrW(1488) & ChrW(1508) & ChrW(1512) & ChrW(1497) & ChrW(1500)
        Case 5:  HebrewMonthName = ChrW(1502) & ChrW(1488) & ChrW(1497)
        Case 6:  HebrewMonthName = ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1497)
        Case 7:  HebrewMonthName = ChrW(1497) & ChrW(1493) & ChrW(1500) & ChrW(1497)
        Case 8:  HebrewMonthName = ChrW(1488) & ChrW(1493) & ChrW(1490) & ChrW(1493) & ChrW(1505) & ChrW(1496)
        Case 9:  HebrewMonthName = ChrW(1505) & ChrW(1508) & ChrW(1496) & ChrW(1502) & ChrW(1489) & ChrW(1512)
        Case 10: HebrewMonthName = ChrW(1488) & ChrW(1493) & ChrW(1511) & ChrW(1496) & ChrW(1493) & ChrW(1489) & ChrW(1512)
        Case 11: HebrewMonthName = ChrW(1504) & ChrW(1493) & ChrW(1489) & ChrW(1502) & ChrW(1489) & ChrW(1512)
        Case 12: HebrewMonthName = ChrW(1491) & ChrW(1510) & ChrW(1502) & ChrW(1489) & ChrW(1512)
        Case Else: HebrewMonthName = CStr(m)
    End Select
End Function

' Period filter: returns minMonth/maxMonth based on G5 (period type) + G6 (period detail)
Private Sub GetMonthRange(ByVal wsMain As Worksheet, ByRef minMonth As Long, ByRef maxMonth As Long)
    Dim periodType As String, periodDetail As String
    Dim wsMgmt As Worksheet, monthIdx As Long, monthName As String
    periodType = Trim$(CStr(wsMain.Range("G5").Value2))
    periodDetail = Trim$(CStr(wsMain.Range("G6").Value2))
    minMonth = 1
    maxMonth = 12
    ' חודשי
    If InStr(1, periodType, ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497), vbTextCompare) > 0 Then
        If periodDetail <> "" Then
            Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
            For monthIdx = 1 To 12
                monthName = Trim$(CStr(wsMgmt.Cells(9 + monthIdx, 18).Value2))
                If StrComp(periodDetail, monthName, vbTextCompare) = 0 Then
                    minMonth = monthIdx
                    maxMonth = monthIdx
                    Exit For
                End If
            Next monthIdx
        End If
    ' רבעוני
    ElseIf InStr(1, periodType, ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497), vbTextCompare) > 0 Then
        If InStr(1, periodDetail, ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1503), vbTextCompare) > 0 Then
            minMonth = 1: maxMonth = 3
        ElseIf InStr(1, periodDetail, ChrW(1513) & ChrW(1504) & ChrW(1497), vbTextCompare) > 0 Then
            minMonth = 4: maxMonth = 6
        ElseIf InStr(1, periodDetail, ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1513) & ChrW(1497), vbTextCompare) > 0 Then
            minMonth = 7: maxMonth = 9
        ElseIf InStr(1, periodDetail, ChrW(1512) & ChrW(1489) & ChrW(1497) & ChrW(1506) & ChrW(1497), vbTextCompare) > 0 Then
            minMonth = 10: maxMonth = 12
        End If
    ' חצי שנתי
    ElseIf InStr(1, periodType, ChrW(1495) & ChrW(1510) & ChrW(1497), vbTextCompare) > 0 Then
        If InStr(1, periodDetail, ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1504), vbTextCompare) > 0 Then
            minMonth = 1: maxMonth = 6
        ElseIf InStr(1, periodDetail, ChrW(1513) & ChrW(1504) & ChrW(1497), vbTextCompare) > 0 Then
            minMonth = 7: maxMonth = 12
        End If
    End If
    ' default = full year (already 1..12)
End Sub

' Date column to use for period filtering (G7)
Private Function GetDateColumn(ByVal wsMain As Worksheet) As Long
    Dim v As String
    v = Trim$(CStr(wsMain.Range("G7").Value2))
    ' "תחילת ביטוח"
    If InStr(1, v, ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & ChrW(1514), vbTextCompare) > 0 Then
        GetDateColumn = RAW_INSURANCE_START
    Else
        GetDateColumn = RAW_BORDEREU
    End If
End Function

' ============================================================================
' SETTINGS HELPERS — read from "הגדרות" sheet
' ============================================================================

' Load REASON_CODE ? Hebrew translation table from N/O columns
Private Function LoadHelperDictionary(ByVal ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Dim r As Long, lastRow As Long, k As String
    lastRow = ws.Cells(ws.Rows.Count, COL_HELPER_KEY).End(xlUp).Row
    For r = 1 To lastRow
        k = Trim$(CStr(ws.Cells(r, COL_HELPER_KEY).Value2))
        If k <> "" Then
            dict(k) = Trim$(CStr(ws.Cells(r, COL_HELPER_VALUE).Value2))
        End If
    Next r
    ' Fallback Hebrew translations for codes that may not be in the table
    If Not dict.Exists("MISSING_CUSTOMER_NUMBER") Then dict("MISSING_CUSTOMER_NUMBER") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)
    If Not dict.Exists("MISSING_CUSTOMER_NAME") Then dict("MISSING_CUSTOMER_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495)
    If Not dict.Exists("MISSING_POLICY") Then dict("MISSING_POLICY") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1505) & ChrW(1492)
    If Not dict.Exists("MISSING_ADDENDUM") Then dict("MISSING_ADDENDUM") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1514) & ChrW(1493) & ChrW(1505) & ChrW(1508) & ChrW(1514)
    If Not dict.Exists("MISSING_COMPANY_NAME") Then dict("MISSING_COMPANY_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492)
    If Not dict.Exists("MISSING_BRANCH_NAME") Then dict("MISSING_BRANCH_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1506) & ChrW(1504) & ChrW(1507)
    If Not dict.Exists("MISSING_AGENT_NAME") Then dict("MISSING_AGENT_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503)
    If Not dict.Exists("MISSING_UNDERWRITER_TELLER_NAME") Then dict("MISSING_UNDERWRITER_TELLER_NAME") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1513) & ChrW(1501) & " " & ChrW(1496) & ChrW(1500) & ChrW(1512)
    If Not dict.Exists("MISSING_CURRENCY") Then dict("MISSING_CURRENCY") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1502) & ChrW(1496) & ChrW(1489) & ChrW(1506)
    If Not dict.Exists("MISSING_PREMIUM") Then dict("MISSING_PREMIUM") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1492)
    If Not dict.Exists("MISSING_COMPANY_COMMISSION") Then dict("MISSING_COMPANY_COMMISSION") = ChrW(1495) & ChrW(1505) & ChrW(1512) & " " & ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1514) & " " & ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492)
    If Not dict.Exists("PREMIUM_OVER_THRESHOLD") Then dict("PREMIUM_OVER_THRESHOLD") = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1492) & " " & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1492)
    If Not dict.Exists("PREMIUM_NOT_NUMERIC") Then dict("PREMIUM_NOT_NUMERIC") = ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1497) & ChrW(1492) & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1502) & ChrW(1505) & ChrW(1508) & ChrW(1512) & ChrW(1497)
    If Not dict.Exists(HELPER_REVIEW_SOURCE_ROW_HEADER) Then dict(HELPER_REVIEW_SOURCE_ROW_HEADER) = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1514) & " " & ChrW(1502) & ChrW(1511) & ChrW(1493) & ChrW(1512)
    If Not dict.Exists(HELPER_REVIEW_REASON_HEADER) Then dict(HELPER_REVIEW_REASON_HEADER) = ChrW(1505) & ChrW(1497) & ChrW(1489) & ChrW(1514) & " " & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1492)
    If Not dict.Exists(HELPER_REVIEW_REASON_CODE_HEADER) Then dict(HELPER_REVIEW_REASON_CODE_HEADER) = ChrW(1511) & ChrW(1493) & ChrW(1491) & " " & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1492)
    Set LoadHelperDictionary = dict
End Function

' Load BRANCH_NAME ? MAIN_BRANCH mapping from columns A/B (rows 3..N)
Private Function LoadBranchMapping(ByVal ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Dim r As Long, lastRow As Long, brName As String, mainBr As String
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 3 To lastRow
        brName = UCase$(Trim$(CStr(ws.Cells(r, 1).Value2)))
        mainBr = Trim$(CStr(ws.Cells(r, 2).Value2))
        If brName <> "" And mainBr <> "" Then dict(brName) = mainBr
    Next r
    Set LoadBranchMapping = dict
End Function

Private Sub ValidateHelperKey(ByVal dict As Object, ByVal key As String)
    If Not dict.Exists(key) Then
        Err.Raise vbObjectError + 2000, "ValidateHelperKey", "Helper key not found in הגדרות N column: " & key
    End If
End Sub

' Load FIELD_KEY ? column number for all fields marked CHECK in the settings table
Private Sub LoadCheckedFields(ByVal ws As Worksheet, ByVal dictCol As Object, ByVal dictDisp As Object)
    Dim r As Long, lastRow As Long
    Dim fName As String, fCol As String, fCheck As String, fKey As String
    lastRow = ws.Cells(ws.Rows.Count, COL_FIELD_NAME_HE).End(xlUp).Row
    For r = MANAGEMENT_START_ROW To lastRow
        fName = Trim$(CStr(ws.Cells(r, COL_FIELD_NAME_HE).Value2))
        fCol = UCase$(Trim$(CStr(ws.Cells(r, COL_FIELD_COLUMN).Value2)))
        fCheck = UCase$(Trim$(CStr(ws.Cells(r, COL_FIELD_CHECKING).Value2)))
        fKey = Trim$(CStr(ws.Cells(r, COL_FIELD_KEY).Value2))
        If fKey <> "" And fCol <> "" And fCheck = "CHECK" Then
            dictCol(fKey) = ColumnLetterToNumber(fCol)
            dictDisp(fKey) = fName
        End If
    Next r
End Sub

' Load full FIELD_KEY ? column number table (CHECK + SKIP + everything with a column)
' Used by SendForReview to know which column to fix for ANY reason — even ones we don't
' currently scan for.
Private Function LoadAllFieldColumns(ByVal ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Dim r As Long, lastRow As Long, fCol As String, fKey As String
    lastRow = ws.Cells(ws.Rows.Count, COL_FIELD_NAME_HE).End(xlUp).Row
    For r = MANAGEMENT_START_ROW To lastRow
        fCol = UCase$(Trim$(CStr(ws.Cells(r, COL_FIELD_COLUMN).Value2)))
        fKey = Trim$(CStr(ws.Cells(r, COL_FIELD_KEY).Value2))
        If fKey <> "" And fCol <> "" Then
            dict(fKey) = ColumnLetterToNumber(fCol)
        End If
    Next r
    Set LoadAllFieldColumns = dict
End Function

Private Sub BuildArrays(ByVal dictCol As Object, ByVal dictDisp As Object, ByRef keys() As String, ByRef cols() As Long, ByRef disp() As String, ByRef cnt As Long)
    cnt = dictCol.Count
    If cnt = 0 Then Exit Sub
    ReDim keys(1 To cnt)
    ReDim cols(1 To cnt)
    ReDim disp(1 To cnt)
    Dim i As Long, k As Variant
    i = 0
    For Each k In dictCol.keys
        i = i + 1
        keys(i) = CStr(k)
        cols(i) = CLng(dictCol(k))
        disp(i) = CStr(dictDisp(k))
    Next k
End Sub

Private Function GetStringParameter(ByVal ws As Worksheet, ByVal paramName As String) As String
    On Error GoTo FAIL
    Dim r As Long, lastRow As Long, nm As String
    lastRow = ws.Cells(ws.Rows.Count, COL_PARAM_NAME).End(xlUp).Row
    For r = 1 To lastRow
        nm = UCase$(Trim$(CStr(ws.Cells(r, COL_PARAM_NAME).Value2)))
        If nm = UCase$(paramName) Then
            GetStringParameter = Trim$(CStr(ws.Cells(r, COL_PARAM_VALUE).Value2))
            Exit Function
        End If
    Next r
    GetStringParameter = ""
    Exit Function
FAIL:
    GetStringParameter = ""
End Function

Private Function GetNumericParameter(ByVal ws As Worksheet, ByVal paramName As String) As Double
    Dim r As Long, lastRow As Long, nm As String, v As Variant, n As Double
    lastRow = ws.Cells(ws.Rows.Count, COL_PARAM_NAME).End(xlUp).Row
    For r = 1 To lastRow
        nm = UCase$(Trim$(CStr(ws.Cells(r, COL_PARAM_NAME).Value2)))
        If nm = UCase$(paramName) Then
            v = ws.Cells(r, COL_PARAM_VALUE).Value2
            If TryParseVariantNumber(v, n) Then
                GetNumericParameter = n
                Exit Function
            Else
                Err.Raise vbObjectError + 3000, , "Parameter not numeric: " & paramName
            End If
        End If
    Next r
    Err.Raise vbObjectError + 3001, , "Parameter not found: " & paramName
End Function

' Reverse-lookup: reason text (e.g. "חסר שם חברה") ? field column in 2025_fixed.
' Uses dictHelper (REASON_CODE ? translated text) and dictAllCols (FIELD_KEY ? column).
Private Function GetFixedColumnFromReason(ByVal reasonText As String, ByVal dictHelper As Object, ByVal dictAllCols As Object) As Long
    Dim k As Variant, foundCode As String
    foundCode = ""
    For Each k In dictHelper.keys
        If StrComp(CStr(dictHelper(k)), reasonText, vbTextCompare) = 0 Then
            If Left$(CStr(k), 8) = "MISSING_" Then
                foundCode = Mid$(CStr(k), 9)
                Exit For
            ElseIf CStr(k) = "PREMIUM_OVER_THRESHOLD" Or CStr(k) = "PREMIUM_NOT_NUMERIC" Then
                foundCode = "PREMIUM"
                Exit For
            End If
        End If
    Next k
    If foundCode <> "" And dictAllCols.Exists(foundCode) Then
        GetFixedColumnFromReason = CLng(dictAllCols(foundCode))
    Else
        GetFixedColumnFromReason = 0
    End If
End Function

' ============================================================================
' BUTTON 1 — BuildReview (Option A workflow)
' ============================================================================
' If 2025_fixed and לטיפול_2025 don't exist  ?  build both fresh from SOURCE
' If they exist                             ?  ask user before replacing
' ============================================================================

Public Sub BuildReview()
    Dim wsMgmt As Worksheet, wsRev As Worksheet, wsFixed As Worksheet
    Dim yearVal As String, refYearStr As String
    Dim revSheetName As String, fixedSheetName As String
    Dim ans As VbMsgBoxResult
    Dim issuesFound As Long
    Dim prevScreenUpdating As Boolean, prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean, prevCalculation As XlCalculation

    On Error GoTo ERR_HANDLER

    yearVal = Trim$(CStr(ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("G4").Value2))
    refYearStr = Trim$(CStr(ThisWorkbook.Worksheets(CONTROL_SHEET_NAME()).Range("G3").Value2))
    If yearVal = "" Then Err.Raise vbObjectError + 1002, "BuildReview", "G4 (שנה נוכחית) ריק"

    revSheetName = REVIEW_SHEET_NAME() & "_" & yearVal
    fixedSheetName = yearVal & "_fixed"

    ' Ask before overwriting if the user has work in flight
    If SheetExists(revSheetName) Or SheetExists(fixedSheetName) Then
        Dim warnMsg As String
        warnMsg = ChrW(1499) & ChrW(1489) & ChrW(1512) & " " & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1502) & ChrW(1497) & ChrW(1501) & " " & _
                  ChrW(1492) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & _
                  fixedSheetName & " " & ChrW(1493) & "/" & ChrW(1488) & ChrW(1493) & " " & revSheetName & "." & vbCrLf & _
                  ChrW(1500) & ChrW(1489) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1502) & ChrW(1495) & ChrW(1491) & ChrW(1513) & " " & _
                  ChrW(1502) & ChrW(1492) & ChrW(1502) & ChrW(1511) & ChrW(1493) & ChrW(1512) & "?" & vbCrLf & _
                  "(" & ChrW(1499) & ChrW(1500) & " " & ChrW(1492) & ChrW(1506) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1493) & ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1497) & ChrW(1488) & ChrW(1489) & ChrW(1491) & ChrW(1493) & ")"
        ans = MsgBoxU(warnMsg, vbYesNo + vbExclamation)
        If ans <> vbYes Then Exit Sub
    End If

    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())

    ' Build fresh year_fixed from SOURCE
    Set wsFixed = BuildFreshFixedFromSource(yearVal)

    ' Scan year_fixed ? לטיפול_year
    issuesFound = BuildReviewSheetFromFixed(wsFixed, wsMgmt, yearVal, wsRev)

    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalculation

    MsgBoxU ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501) & "." & vbCrLf & _
            ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & issuesFound & " " & _
            ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1493) & ChrW(1514) & " " & _
            ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500), vbInformation
    Exit Sub

ERR_HANDLER:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBoxU "BuildReview error (line " & Erl & "): " & vbCrLf & Err.Description, vbCritical
End Sub

' ============================================================================
' Helper: build a fresh `yearVal_fixed` sheet from SOURCE/yearVal.xlsx
' Adds an "_IGNORED" column at the right edge (initialized to 0 for all rows).
' ============================================================================
Public Function BuildFreshFixedFromSource(ByVal yearVal As String) As Worksheet
    Dim wbSrc As Workbook, wsSrc As Worksheet, wsFixed As Worksheet
    Dim srcPath As String
    Dim lastRow As Long, lastCol As Long, r As Long
    Dim fixedSheetName As String

    On Error GoTo ERR_HANDLER

    fixedSheetName = yearVal & "_fixed"

    srcPath = FindSourceFile(yearVal)
    If srcPath = "" Then
        Err.Raise vbObjectError + 1003, "BuildFreshFixedFromSource", _
            "Source file not found: " & SOURCE_FOLDER() & yearVal & ".xlsx (or .xls)"
    End If

    DeleteSheetIfExists fixedSheetName

    Set wsFixed = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsFixed.Name = fixedSheetName

    Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True)
    Set wsSrc = OpenDataSheet(wbSrc)

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    If lastRow >= 1 And lastCol >= 1 Then
        wsFixed.Range(wsFixed.Cells(1, 1), wsFixed.Cells(lastRow, lastCol)).Value = _
            wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).Value
    End If

    ' _IGNORED column to the right of the last data column
    wsFixed.Cells(1, lastCol + 1).Value = IGNORED_HEADER
    If lastRow >= 2 Then
        wsFixed.Range(wsFixed.Cells(2, lastCol + 1), wsFixed.Cells(lastRow, lastCol + 1)).Value = 0
    End If

    wbSrc.Close SaveChanges:=False
    Set wbSrc = Nothing

    wsFixed.Rows(1).Font.Bold = True
    wsFixed.DisplayRightToLeft = True

    Set BuildFreshFixedFromSource = wsFixed
    Exit Function

ERR_HANDLER:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Err.Raise Err.Number, "BuildFreshFixedFromSource", Err.Description
End Function

' ============================================================================
' Helper: scan a `_fixed` sheet for issues, build a לטיפול sheet
' Returns the number of issues found and sets wsRev (ByRef) to the new sheet.
' Skips rows where _IGNORED = 1.
' ============================================================================
Public Function BuildReviewSheetFromFixed(ByVal wsFixed As Worksheet, ByVal wsMgmt As Worksheet, _
                                           ByVal yearVal As String, ByRef wsRev As Worksheet) As Long
    Dim revSheetName As String
    Dim dictHelper As Object, dictFieldCol As Object, dictFieldDisp As Object
    Dim keys() As String, cols() As Long, disp() As String
    Dim cnt As Long, r As Long, lastRow As Long, j As Long, i As Long
    Dim outRow As Long
    Dim threshold As Double
    Dim singleCode As String, singleText As String
    Dim premiumVal As Variant, premiumNum As Double
    Dim ignoredCol As Long
    Dim actionCol As Long, rng As Range
    Dim defaultAction As String, listFormula As String

    revSheetName = REVIEW_SHEET_NAME() & "_" & yearVal

    Set dictHelper = LoadHelperDictionary(wsMgmt)
    ValidateHelperKey dictHelper, HELPER_REVIEW_SOURCE_ROW_HEADER
    ValidateHelperKey dictHelper, HELPER_REVIEW_REASON_HEADER

    Set dictFieldCol = CreateObject("Scripting.Dictionary")
    Set dictFieldDisp = CreateObject("Scripting.Dictionary")
    dictFieldCol.CompareMode = vbTextCompare
    dictFieldDisp.CompareMode = vbTextCompare
    LoadCheckedFields wsMgmt, dictFieldCol, dictFieldDisp
    BuildArrays dictFieldCol, dictFieldDisp, keys, cols, disp, cnt
    If cnt = 0 Then Err.Raise vbObjectError + 1001, , "אין שדות מסומנים כ-CHECK בהגדרות"

    threshold = GetNumericParameter(wsMgmt, PARAM_PREMIUM_THRESHOLD)

    ignoredCol = FindIgnoredColumn(wsFixed)
    lastRow = wsFixed.Cells(wsFixed.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 1006, , "אין נתונים ב-" & wsFixed.Name

    DeleteSheetIfExists revSheetName
    Set wsRev = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsRev.Name = revSheetName

    ' Headers: source row | field 1..cnt | reason | פעולה | תיקון
    wsRev.Cells(1, 1).Value = dictHelper(HELPER_REVIEW_SOURCE_ROW_HEADER)
    For i = 1 To cnt
        wsRev.Cells(1, i + 1).Value = disp(i)
    Next i
    wsRev.Cells(1, cnt + 2).Value = dictHelper(HELPER_REVIEW_REASON_HEADER)
    wsRev.Cells(1, cnt + 3).Value = ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) & ChrW(1492)
    wsRev.Cells(1, cnt + 4).Value = ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1503)

    outRow = 2
    For r = 2 To lastRow
        ' Skip _IGNORED rows
        If ignoredCol > 0 Then
            If CLng(Val(wsFixed.Cells(r, ignoredCol).Value2)) = 1 Then GoTo NextRow
        End If

        ' Skip rows where ALL CHECK fields are blank
        If IsIgnorableRow(wsFixed, r, keys, cols, cnt, dictFieldCol) Then GoTo NextRow

        ' Missing required fields ? one row per missing field
        For i = 1 To cnt
            If IsBlankValue(wsFixed.Cells(r, cols(i)).Value2) Then
                singleCode = "MISSING_" & keys(i)
                If dictHelper.Exists(singleCode) Then
                    singleText = dictHelper(singleCode)
                Else
                    singleText = singleCode
                End If
                wsRev.Cells(outRow, 1).Value = r
                For j = 1 To cnt
                    wsRev.Cells(outRow, j + 1).Value = wsFixed.Cells(r, cols(j)).Value2
                Next j
                wsRev.Cells(outRow, cnt + 2).Value = singleText
                outRow = outRow + 1
            End If
        Next i

        ' Premium threshold check
        premiumVal = wsFixed.Cells(r, dictFieldCol(KEY_PREMIUM)).Value2
        If Not IsBlankValue(premiumVal) Then
            If TryParseVariantNumber(premiumVal, premiumNum) Then
                If Abs(premiumNum) > threshold Then
                    singleCode = "PREMIUM_OVER_THRESHOLD"
                    If dictHelper.Exists(singleCode) Then
                        singleText = dictHelper(singleCode)
                    Else
                        singleText = singleCode
                    End If
                    wsRev.Cells(outRow, 1).Value = r
                    For j = 1 To cnt
                        wsRev.Cells(outRow, j + 1).Value = wsFixed.Cells(r, cols(j)).Value2
                    Next j
                    wsRev.Cells(outRow, cnt + 2).Value = singleText
                    outRow = outRow + 1
                End If
            Else
                singleCode = "PREMIUM_NOT_NUMERIC"
                If dictHelper.Exists(singleCode) Then
                    singleText = dictHelper(singleCode)
                Else
                    singleText = singleCode
                End If
                wsRev.Cells(outRow, 1).Value = r
                For j = 1 To cnt
                    wsRev.Cells(outRow, j + 1).Value = wsFixed.Cells(r, cols(j)).Value2
                Next j
                wsRev.Cells(outRow, cnt + 2).Value = singleText
                outRow = outRow + 1
            End If
        End If
NextRow:
    Next r

    actionCol = cnt + 3

    ' Default value + dropdown for action column
    If outRow > 2 Then
        ' "העבר לבדיקה"
        defaultAction = ChrW(1492) & ChrW(1506) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1492)
        ' "תקן,התעלם,העבר לבדיקה"
        listFormula = ChrW(1514) & ChrW(1511) & ChrW(1503) & "," & _
                      ChrW(1492) & ChrW(1514) & ChrW(1506) & ChrW(1500) & ChrW(1501) & "," & defaultAction
        Set rng = wsRev.Range(wsRev.Cells(2, actionCol), wsRev.Cells(outRow - 1, actionCol))
        rng.Value = defaultAction
        On Error Resume Next
        rng.Validation.Delete
        rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listFormula
        rng.Validation.IgnoreBlank = True
        rng.Validation.InCellDropdown = True
        On Error GoTo 0
    End If

    wsRev.Rows(1).Font.Bold = True
    wsRev.Columns.AutoFit
    wsRev.DisplayRightToLeft = True
    wsRev.Cells(1, actionCol).Interior.Color = RGB(255, 165, 0)
    wsRev.Cells(1, actionCol + 1).Interior.Color = RGB(255, 165, 0)
    wsRev.Columns(actionCol).ColumnWidth = 15
    wsRev.Columns(actionCol + 1).ColumnWidth = 30
    If outRow > 2 Then
        wsRev.Range(wsRev.Cells(2, actionCol), wsRev.Cells(outRow - 1, actionCol + 1)).Interior.Color = RGB(255, 255, 200)
    End If

    ' Red "סיימתי לעדכן" button at the top-left of the sheet
    Dim shpBtn As Shape
    Set shpBtn = wsRev.Shapes.AddShape(msoShapeRoundedRectangle, 10, 2, 160, 30)
    shpBtn.Name = "btnSendForReview"
    shpBtn.Fill.ForeColor.RGB = RGB(180, 0, 0)
    shpBtn.Line.Visible = msoFalse
    shpBtn.TextFrame2.TextRange.Text = ChrW(1505) & ChrW(1497) & ChrW(1497) & ChrW(1502) & ChrW(1514) & ChrW(1497) & " " & ChrW(1500) & ChrW(1506) & ChrW(1491) & ChrW(1499) & ChrW(1503)
    shpBtn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shpBtn.TextFrame2.TextRange.Font.Size = 12
    shpBtn.TextFrame2.TextRange.Font.Bold = msoTrue
    shpBtn.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shpBtn.OnAction = "SendForReview"

    BuildReviewSheetFromFixed = outRow - 2
End Function

' ============================================================================
' SendForReview — the "סיימתי לעדכן" button on לטיפול
' ============================================================================
' On click:
'   1. Rebuild yearVal_fixed fresh from SOURCE
'   2. For each row in the לטיפול sheet:
'        "תקן"          ? write fixText into the right column of yearVal_fixed
'        "התעלם"        ? set _IGNORED = 1 in yearVal_fixed
'        "העבר לבדיקה"  ? collect into the email batch
'   3. If any "העבר לבדיקה" rows: open Outlook with attached .xlsx
'   4. לטיפול sheet is left untouched
' ============================================================================

Public Sub SendForReview()
    Dim wsRev As Worksheet
    Dim wsMgmt As Worksheet
    Dim wsFixed As Worksheet
    Dim yearVal As String
    Dim dictHelper As Object, dictAllCols As Object
    Dim revLastRow As Long, revLastCol As Long
    Dim hdrCol As Long
    Dim hdrActionText As String, hdrFixText As String, hdrReasonText As String
    Dim actionColIdx As Long, fixColIdx As Long, reasonColIdx As Long
    Dim r As Long, sourceRow As Long
    Dim actionText As String, fixText As String, reasonText As String
    Dim actionFix As String, actionIgnore As String, actionSend As String
    Dim countFix As Long, countIgnore As Long, countSend As Long
    Dim sendRows As Object
    Dim ignoredCol As Long
    Dim targetCol As Long, numVal As Double
    Dim prevScreenUpdating As Boolean, prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean, prevCalculation As XlCalculation

    On Error GoTo ERR_HANDLER

    Set wsRev = ActiveSheet
    If InStr(1, wsRev.Name, REVIEW_SHEET_NAME(), vbTextCompare) = 0 Then
        MsgBoxU ChrW(1497) & ChrW(1513) & " " & ChrW(1500) & ChrW(1492) & ChrW(1512) & ChrW(1497) & ChrW(1509) & " " & _
                ChrW(1502) & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & _
                ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500), vbExclamation
        Exit Sub
    End If

    yearVal = Mid$(wsRev.Name, InStr(1, wsRev.Name, "_") + 1)
    If Not IsNumeric(yearVal) Then
        MsgBoxU "Sheet name parse failed: " & wsRev.Name, vbExclamation
        Exit Sub
    End If

    Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())

    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Step 1: rebuild yearVal_fixed fresh from SOURCE
    Set wsFixed = BuildFreshFixedFromSource(yearVal)
    ignoredCol = FindIgnoredColumn(wsFixed)

    ' Load dictionaries needed for action/fix logic
    Set dictHelper = LoadHelperDictionary(wsMgmt)
    Set dictAllCols = LoadAllFieldColumns(wsMgmt)

    ' Find action / fix / reason columns by header
    revLastRow = wsRev.Cells(wsRev.Rows.Count, 1).End(xlUp).Row
    revLastCol = wsRev.Cells(1, wsRev.Columns.Count).End(xlToLeft).Column
    hdrActionText = ChrW(1508) & ChrW(1506) & ChrW(1493) & ChrW(1500) & ChrW(1492)
    hdrFixText = ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1493) & ChrW(1503)
    hdrReasonText = dictHelper(HELPER_REVIEW_REASON_HEADER)
    actionColIdx = 0: fixColIdx = 0: reasonColIdx = 0
    For hdrCol = 1 To revLastCol + 2
        If StrComp(Trim$(CStr(wsRev.Cells(1, hdrCol).Value2)), hdrActionText, vbTextCompare) = 0 Then actionColIdx = hdrCol
        If StrComp(Trim$(CStr(wsRev.Cells(1, hdrCol).Value2)), hdrFixText, vbTextCompare) = 0 Then fixColIdx = hdrCol
        If StrComp(Trim$(CStr(wsRev.Cells(1, hdrCol).Value2)), hdrReasonText, vbTextCompare) = 0 Then reasonColIdx = hdrCol
    Next hdrCol
    If actionColIdx = 0 Then actionColIdx = revLastCol - 1
    If fixColIdx = 0 Then fixColIdx = revLastCol
    If reasonColIdx = 0 Then reasonColIdx = actionColIdx - 1

    ' Action text constants (Hebrew)
    actionFix = ChrW(1514) & ChrW(1511) & ChrW(1503)                                                   ' תקן
    actionIgnore = ChrW(1492) & ChrW(1514) & ChrW(1506) & ChrW(1500) & ChrW(1501)                       ' התעלם
    actionSend = ChrW(1492) & ChrW(1506) & ChrW(1489) & ChrW(1512) & " " & ChrW(1500) & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1492)  ' העבר לבדיקה

    countFix = 0: countIgnore = 0: countSend = 0
    Set sendRows = CreateObject("Scripting.Dictionary")  ' key = sourceRow, value = first review row idx

    ' Process every review row
    For r = 2 To revLastRow
        sourceRow = CLng(Val(wsRev.Cells(r, 1).Value2))
        If sourceRow < 2 Then GoTo NextR

        actionText = Trim$(CStr(wsRev.Cells(r, actionColIdx).Value2))
        fixText = CStr(wsRev.Cells(r, fixColIdx).Value2)
        reasonText = Trim$(CStr(wsRev.Cells(r, reasonColIdx).Value2))

        If StrComp(actionText, actionFix, vbTextCompare) = 0 Then
            ' Apply correction
            targetCol = GetFixedColumnFromReason(reasonText, dictHelper, dictAllCols)
            If targetCol > 0 And Trim$(fixText) <> "" Then
                If TryParseVariantNumber(fixText, numVal) Then
                    wsFixed.Cells(sourceRow, targetCol).Value = numVal
                Else
                    wsFixed.Cells(sourceRow, targetCol).Value = fixText
                End If
                countFix = countFix + 1
            End If

        ElseIf StrComp(actionText, actionIgnore, vbTextCompare) = 0 Then
            ' Mark row as ignored
            If ignoredCol > 0 Then
                wsFixed.Cells(sourceRow, ignoredCol).Value = 1
                countIgnore = countIgnore + 1
            End If

        ElseIf StrComp(actionText, actionSend, vbTextCompare) = 0 Then
            ' Collect for email (one entry per source row, in case of multiple issues)
            If Not sendRows.Exists(CStr(sourceRow)) Then
                sendRows.Add CStr(sourceRow), r
                countSend = countSend + 1
            End If
        End If
NextR:
    Next r

    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalculation

    ' Step 3: send email if there are any "העבר לבדיקה" rows
    If countSend > 0 Then
        SendOpenIssuesEmail wsRev, sendRows, wsMgmt, yearVal
    End If

    ' Step 4: summary
    Dim msg As String
    msg = ChrW(1505) & ChrW(1497) & ChrW(1499) & ChrW(1493) & ChrW(1501) & " " & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & ":" & vbCrLf & vbCrLf & _
          ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501) & " " & ChrW(1513) & ChrW(1492) & ChrW(1493) & ChrW(1495) & ChrW(1500) & ChrW(1493) & ":  " & countFix & vbCrLf & _
          ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1513) & ChrW(1492) & ChrW(1493) & ChrW(1514) & ChrW(1506) & ChrW(1500) & ChrW(1502) & ChrW(1493) & ":  " & countIgnore & vbCrLf & _
          ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1513) & ChrW(1504) & ChrW(1513) & ChrW(1500) & ChrW(1495) & ChrW(1493) & " " & ChrW(1489) & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & ":  " & countSend
    MsgBoxU msg, vbInformation
    Exit Sub

ERR_HANDLER:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBoxU "SendForReview error (line " & Erl & "): " & vbCrLf & Err.Description, vbCritical
End Sub

' ============================================================================
' Build & open Outlook email with the "העבר לבדיקה" rows attached as .xlsx
' ============================================================================
Private Sub SendOpenIssuesEmail(ByVal wsRev As Worksheet, ByVal sendRows As Object, _
                                  ByVal wsMgmt As Worksheet, ByVal yearVal As String)
    Dim emailAddr As String
    emailAddr = GetStringParameter(wsMgmt, PARAM_ERROR_EMAIL)
    If emailAddr = "" Then
        MsgBoxU ChrW(1500) & ChrW(1488) & " " & ChrW(1492) & ChrW(1493) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1492) & " " & _
                ChrW(1499) & ChrW(1514) & ChrW(1493) & ChrW(1489) & ChrW(1514) & " " & ChrW(1488) & ChrW(1497) & ChrW(1502) & ChrW(1497) & ChrW(1497) & ChrW(1500) & " " & _
                "(" & PARAM_ERROR_EMAIL & ")" & " " & ChrW(1489) & ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514), vbExclamation
        Exit Sub
    End If

    Dim revLastCol As Long, hdrCol As Long
    revLastCol = wsRev.Cells(1, wsRev.Columns.Count).End(xlToLeft).Column

    ' Build temp xlsx with the headers + selected rows
    Dim wbTemp As Workbook, wsTemp As Worksheet
    Set wbTemp = Workbooks.Add
    Set wsTemp = wbTemp.Worksheets(1)
    For hdrCol = 1 To revLastCol
        wsTemp.Cells(1, hdrCol).Value = wsRev.Cells(1, hdrCol).Value
    Next hdrCol
    wsTemp.Rows(1).Font.Bold = True

    Dim outRow As Long, k As Variant, srcReviewRow As Long
    outRow = 2
    For Each k In sendRows.keys
        srcReviewRow = CLng(sendRows(k))
        For hdrCol = 1 To revLastCol
            wsTemp.Cells(outRow, hdrCol).Value = wsRev.Cells(srcReviewRow, hdrCol).Value
        Next hdrCol
        outRow = outRow + 1
    Next k
    wsTemp.Columns.AutoFit
    wsTemp.DisplayRightToLeft = True

    Dim tempPath As String
    tempPath = ThisWorkbook.Path & "\" & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501) & "_" & wsRev.Name & ".xlsx"
    Application.DisplayAlerts = False
    wbTemp.SaveAs tempPath, xlOpenXMLWorkbook
    wbTemp.Close SaveChanges:=False
    Application.DisplayAlerts = True

    ' Email body
    Dim emailBody As String
    emailBody = ChrW(1492) & ChrW(1497) & " " & ChrW(1500) & ChrW(1492) & ChrW(1489) & "," & vbCrLf & vbCrLf & _
                ChrW(1500) & ChrW(1511) & ChrW(1512) & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1499) & ChrW(1504) & ChrW(1514) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & " " & ChrW(1506) & ChrW(1489) & ChrW(1493) & ChrW(1512) & ChrW(1499) & ChrW(1501) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & ChrW(1492) & ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501) & " " & ChrW(1492) & ChrW(1512) & ChrW(34) & ChrW(1502) & "." & vbCrLf & _
                ChrW(1506) & ChrW(1500) & " " & ChrW(1502) & ChrW(1504) & ChrW(1514) & " " & ChrW(1500) & ChrW(1492) & ChrW(1508) & ChrW(1497) & ChrW(1511) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1491) & ChrW(1493) & ChrW(1495) & " " & ChrW(1488) & ChrW(1504) & ChrW(1497) & " " & ChrW(1502) & ChrW(1489) & ChrW(1511) & ChrW(1513) & ChrW(1514) & " " & ChrW(1514) & ChrW(1490) & ChrW(1493) & ChrW(1489) & ChrW(1514) & ChrW(1498) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1502) & ChrW(1504) & ChrW(1514) & " " & ChrW(1513) & ChrW(1488) & ChrW(1514) & ChrW(1511) & ChrW(1503) & " " & ChrW(1489) & ChrW(1492) & ChrW(1514) & ChrW(1488) & ChrW(1501) & "." & vbCrLf & vbCrLf & _
                ChrW(1514) & ChrW(1493) & ChrW(1491) & ChrW(1492) & "," & vbCrLf & _
                ChrW(1488) & ChrW(1493) & ChrW(1512) & ChrW(1497) & ChrW(1514)

    Dim emailSubject As String
    emailSubject = ChrW(1495) & ChrW(1512) & ChrW(1497) & ChrW(1490) & ChrW(1497) & ChrW(1501) & " " & ChrW(1513) & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ChrW(1493) & " " & ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1498) & " - " & yearVal

    ' Outlook (late binding)
    Dim olApp As Object, olMail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    olMail.To = emailAddr
    olMail.Subject = emailSubject
    olMail.Body = emailBody
    olMail.Attachments.Add tempPath
    olMail.Display
End Sub

' ============================================================================
' BUTTON 2 — ApplyCorrectionsAndBuildReports (Option A workflow)
' ============================================================================
' Reads from yearVal_fixed (corrections already applied via SendForReview)
' Skips rows where _IGNORED = 1.
' Falls back to SOURCE only if yearVal_fixed doesn't exist (e.g. ref year).
' Builds base_yearVal + comparison sheets + summary.
' ============================================================================

Public Sub ApplyCorrectionsAndBuildReports()
    Dim wsMain As Worksheet, wsMgmt As Worksheet
    Dim wsSrcCur As Worksheet, wsSrcRef As Worksheet
    Dim wsBase As Worksheet, wsBaseRef As Worksheet
    Dim wbSrcExternal As Workbook, wbRefExternal As Workbook
    Dim curIsFixed As Boolean, refIsFixed As Boolean
    Dim ignoredColCur As Long, ignoredColRef As Long
    Dim dictHelper As Object, dictBranch As Object
    Dim yearVal As String, refYear As String
    Dim threshold As Double
    Dim minMonth As Long, maxMonth As Long, dateCol As Long
    Dim lastRow As Long, r As Long, outRow As Long
    Dim baseSheetName As String, refBaseSheetName As String
    Dim countCurrent As Long, countRef As Long
    Dim periodDesc As String
    Dim premVal As Double, monthVal As Long, bordereu As Variant
    Dim brKey As String, dtStr As String, mPart As String
    Dim debugStep As String
    Dim prevScreenUpdating As Boolean, prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean, prevCalculation As XlCalculation

    On Error GoTo ERR_HANDLER

    debugStep = "INIT"
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())

    debugStep = "READ_YEARS"
    yearVal = Trim$(CStr(wsMain.Range("G4").Value2))
    refYear = Trim$(CStr(wsMain.Range("G3").Value2))
    If yearVal = "" Or refYear = "" Then Err.Raise vbObjectError + 2001, , "G3 או G4 ריקים"

    debugStep = "LOAD_HELPERS"
    Set dictHelper = LoadHelperDictionary(wsMgmt)
    Set dictBranch = LoadBranchMapping(wsMgmt)
    threshold = GetNumericParameter(wsMgmt, PARAM_PREMIUM_THRESHOLD)

    debugStep = "GET_PERIOD"
    GetMonthRange wsMain, minMonth, maxMonth
    dateCol = GetDateColumn(wsMain)

    ' --- Cleanup: keep control, settings, _fixed sheets, base_REF, and לטיפול sheets ---
    debugStep = "CLEANUP"
    Dim iSheet As Long, ws As Worksheet, keepIt As Boolean
    For iSheet = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(iSheet)
        keepIt = False
        If ws.Name = CONTROL_SHEET_NAME() Then keepIt = True
        If ws.Name = MANAGEMENT_SHEET_NAME() Then keepIt = True
        If ws.Name = (yearVal & "_fixed") Then keepIt = True
        If ws.Name = (refYear & "_fixed") Then keepIt = True
        If ws.Name = (ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505) & "_" & refYear) Then keepIt = True
        If InStr(1, ws.Name, REVIEW_SHEET_NAME(), vbTextCompare) > 0 Then keepIt = True
        If LCase$(ws.Name) = "דוח_template" Then keepIt = True
        If Not keepIt Then ws.Delete
    Next iSheet

    ' --- Determine source for current year: prefer YearVal_fixed; fallback to SOURCE ---
    debugStep = "OPEN_CUR_DATA"
    If SheetExists(yearVal & "_fixed") Then
        Set wsSrcCur = ThisWorkbook.Worksheets(yearVal & "_fixed")
        curIsFixed = True
        ignoredColCur = FindIgnoredColumn(wsSrcCur)
    Else
        Dim curPath As String
        curPath = FindSourceFile(yearVal)
        If curPath = "" Then Err.Raise vbObjectError + 2002, , "Source not found for current year: " & yearVal
        Set wbSrcExternal = Workbooks.Open(curPath, ReadOnly:=True)
        Set wsSrcCur = OpenDataSheet(wbSrcExternal)
        curIsFixed = False
        ignoredColCur = 0
    End If

    ' --- Build base sheet for current year ---
    debugStep = "BUILD_BASE_CUR"
    baseSheetName = ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505) & "_" & yearVal
    DeleteSheetIfExists baseSheetName
    Set wsBase = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsBase.Name = baseSheetName
    WriteBaseHeaders wsBase

    lastRow = wsSrcCur.Cells(wsSrcCur.Rows.Count, 1).End(xlUp).Row
    outRow = 2
    For r = 2 To lastRow
        ' Skip rows marked _IGNORED
        If ignoredColCur > 0 Then
            If CLng(Val(wsSrcCur.Cells(r, ignoredColCur).Value2)) = 1 Then GoTo NextCurRow
        End If

        ' Premium threshold filter (only when reading raw SOURCE, not fixed)
        If Not curIsFixed Then
            premVal = 0
            If Not IsBlankValue(wsSrcCur.Cells(r, RAW_PREMIUM).Value2) Then
                If TryParseVariantNumber(wsSrcCur.Cells(r, RAW_PREMIUM).Value2, premVal) Then
                    If Abs(premVal) > threshold Then GoTo NextCurRow
                End If
            End If
        End If

        ' Compute the month from the configured date column
        monthVal = 0
        bordereu = wsSrcCur.Cells(r, dateCol).Value2
        If IsDate(bordereu) Then
            monthVal = Month(CDate(bordereu))
        ElseIf IsNumeric(bordereu) Then
            If CDbl(bordereu) > 1 Then monthVal = Month(CDate(CDbl(bordereu)))
        ElseIf Not IsBlankValue(bordereu) Then
            dtStr = CStr(bordereu)
            If Len(dtStr) >= 7 Then
                mPart = Mid$(dtStr, 6, 2)
                If IsNumeric(mPart) Then monthVal = CInt(mPart)
            End If
        End If

        ' Read premium safely
        premVal = 0
        If TryParseVariantNumber(wsSrcCur.Cells(r, RAW_PREMIUM).Value2, premVal) Then
            ' ok
        Else
            premVal = 0
        End If

        ' Write base row
        wsBase.Cells(outRow, BASE_COL_ID).Value = r
        wsBase.Cells(outRow, BASE_COL_YEAR).Value = yearVal
        wsBase.Cells(outRow, BASE_COL_MONTH).Value = monthVal
        wsBase.Cells(outRow, BASE_COL_IDENTITY).Value = wsSrcCur.Cells(r, RAW_IDNUMBER).Value2
        wsBase.Cells(outRow, BASE_COL_CUSTOMER).Value = wsSrcCur.Cells(r, RAW_CUSTOMER).Value2
        wsBase.Cells(outRow, BASE_COL_CUSTNAME).Value = wsSrcCur.Cells(r, RAW_CUSTNAME).Value2
        wsBase.Cells(outRow, BASE_COL_POLICY).Value = wsSrcCur.Cells(r, RAW_POLICY).Value2
        wsBase.Cells(outRow, BASE_COL_ADDENDUM).Value = wsSrcCur.Cells(r, RAW_ADDENDUM).Value2
        wsBase.Cells(outRow, BASE_COL_COMPANY).Value = wsSrcCur.Cells(r, RAW_COMPANY).Value2
        wsBase.Cells(outRow, BASE_COL_COMPNUM).Value = wsSrcCur.Cells(r, RAW_COMPNUM).Value2
        wsBase.Cells(outRow, BASE_COL_BRANCHNAME).Value = wsSrcCur.Cells(r, RAW_BRANCHNAME).Value2
        wsBase.Cells(outRow, BASE_COL_BRANCHNUM).Value = wsSrcCur.Cells(r, RAW_BRANCHNUM).Value2
        brKey = UCase$(Trim$(CStr(wsSrcCur.Cells(r, RAW_BRANCHNAME).Value2)))
        If dictBranch.Exists(brKey) Then
            wsBase.Cells(outRow, BASE_COL_MAINBRANCH).Value = dictBranch(brKey)
        Else
            wsBase.Cells(outRow, BASE_COL_MAINBRANCH).Value = wsSrcCur.Cells(r, RAW_BRANCHNAME).Value2
        End If
        wsBase.Cells(outRow, BASE_COL_AGENTNAME).Value = wsSrcCur.Cells(r, RAW_AGENTNAME).Value2
        wsBase.Cells(outRow, BASE_COL_AGENTNUM).Value = wsSrcCur.Cells(r, RAW_AGENTNUM).Value2
        wsBase.Cells(outRow, BASE_COL_TELLER).Value = wsSrcCur.Cells(r, RAW_TELLERNAME).Value2
        wsBase.Cells(outRow, BASE_COL_TELLERNUM).Value = wsSrcCur.Cells(r, RAW_TELLERNUM).Value2
        wsBase.Cells(outRow, BASE_COL_ACTION).Value = wsSrcCur.Cells(r, RAW_ACTIONCOL).Value2
        wsBase.Cells(outRow, BASE_COL_PREMIUM).Value = premVal
        wsBase.Cells(outRow, BASE_COL_COMMISSION).Value = wsSrcCur.Cells(r, RAW_COMMISSION).Value2
        outRow = outRow + 1
NextCurRow:
    Next r
    countCurrent = outRow - 2
    wsBase.Columns.AutoFit

    ' Close external workbook if we used one
    If Not wbSrcExternal Is Nothing Then
        wbSrcExternal.Close SaveChanges:=False
        Set wbSrcExternal = Nothing
    End If

    ' --- Reference year base ---
    debugStep = "BUILD_BASE_REF"
    refBaseSheetName = ChrW(1489) & ChrW(1505) & ChrW(1497) & ChrW(1505) & "_" & refYear
    If SheetExists(refBaseSheetName) Then
        Set wsBaseRef = ThisWorkbook.Worksheets(refBaseSheetName)
        GoTo BUILD_COMPARISONS
    End If

    If SheetExists(refYear & "_fixed") Then
        Set wsSrcRef = ThisWorkbook.Worksheets(refYear & "_fixed")
        refIsFixed = True
        ignoredColRef = FindIgnoredColumn(wsSrcRef)
    Else
        Dim refPath As String
        refPath = FindSourceFile(refYear)
        If refPath = "" Then Err.Raise vbObjectError + 2003, , "Source not found for ref year: " & refYear
        Set wbRefExternal = Workbooks.Open(refPath, ReadOnly:=True)
        Set wsSrcRef = OpenDataSheet(wbRefExternal)
        refIsFixed = False
        ignoredColRef = 0
    End If

    Set wsBaseRef = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsBaseRef.Name = refBaseSheetName
    WriteBaseHeaders wsBaseRef

    lastRow = wsSrcRef.Cells(wsSrcRef.Rows.Count, 1).End(xlUp).Row
    outRow = 2
    For r = 2 To lastRow
        If ignoredColRef > 0 Then
            If CLng(Val(wsSrcRef.Cells(r, ignoredColRef).Value2)) = 1 Then GoTo NextRefRow
        End If

        If Not refIsFixed Then
            premVal = 0
            If Not IsBlankValue(wsSrcRef.Cells(r, RAW_PREMIUM).Value2) Then
                If TryParseVariantNumber(wsSrcRef.Cells(r, RAW_PREMIUM).Value2, premVal) Then
                    If Abs(premVal) > threshold Then GoTo NextRefRow
                End If
            End If
        End If

        monthVal = 0
        bordereu = wsSrcRef.Cells(r, dateCol).Value2
        If IsDate(bordereu) Then
            monthVal = Month(CDate(bordereu))
        ElseIf IsNumeric(bordereu) Then
            If CDbl(bordereu) > 1 Then monthVal = Month(CDate(CDbl(bordereu)))
        ElseIf Not IsBlankValue(bordereu) Then
            dtStr = CStr(bordereu)
            If Len(dtStr) >= 7 Then
                mPart = Mid$(dtStr, 6, 2)
                If IsNumeric(mPart) Then monthVal = CInt(mPart)
            End If
        End If

        premVal = 0
        If TryParseVariantNumber(wsSrcRef.Cells(r, RAW_PREMIUM).Value2, premVal) Then
        Else
            premVal = 0
        End If

        wsBaseRef.Cells(outRow, BASE_COL_ID).Value = r
        wsBaseRef.Cells(outRow, BASE_COL_YEAR).Value = refYear
        wsBaseRef.Cells(outRow, BASE_COL_MONTH).Value = monthVal
        wsBaseRef.Cells(outRow, BASE_COL_IDENTITY).Value = wsSrcRef.Cells(r, RAW_IDNUMBER).Value2
        wsBaseRef.Cells(outRow, BASE_COL_CUSTOMER).Value = wsSrcRef.Cells(r, RAW_CUSTOMER).Value2
        wsBaseRef.Cells(outRow, BASE_COL_CUSTNAME).Value = wsSrcRef.Cells(r, RAW_CUSTNAME).Value2
        wsBaseRef.Cells(outRow, BASE_COL_POLICY).Value = wsSrcRef.Cells(r, RAW_POLICY).Value2
        wsBaseRef.Cells(outRow, BASE_COL_ADDENDUM).Value = wsSrcRef.Cells(r, RAW_ADDENDUM).Value2
        wsBaseRef.Cells(outRow, BASE_COL_COMPANY).Value = wsSrcRef.Cells(r, RAW_COMPANY).Value2
        wsBaseRef.Cells(outRow, BASE_COL_COMPNUM).Value = wsSrcRef.Cells(r, RAW_COMPNUM).Value2
        wsBaseRef.Cells(outRow, BASE_COL_BRANCHNAME).Value = wsSrcRef.Cells(r, RAW_BRANCHNAME).Value2
        wsBaseRef.Cells(outRow, BASE_COL_BRANCHNUM).Value = wsSrcRef.Cells(r, RAW_BRANCHNUM).Value2
        brKey = UCase$(Trim$(CStr(wsSrcRef.Cells(r, RAW_BRANCHNAME).Value2)))
        If dictBranch.Exists(brKey) Then
            wsBaseRef.Cells(outRow, BASE_COL_MAINBRANCH).Value = dictBranch(brKey)
        Else
            wsBaseRef.Cells(outRow, BASE_COL_MAINBRANCH).Value = wsSrcRef.Cells(r, RAW_BRANCHNAME).Value2
        End If
        wsBaseRef.Cells(outRow, BASE_COL_AGENTNAME).Value = wsSrcRef.Cells(r, RAW_AGENTNAME).Value2
        wsBaseRef.Cells(outRow, BASE_COL_AGENTNUM).Value = wsSrcRef.Cells(r, RAW_AGENTNUM).Value2
        wsBaseRef.Cells(outRow, BASE_COL_TELLER).Value = wsSrcRef.Cells(r, RAW_TELLERNAME).Value2
        wsBaseRef.Cells(outRow, BASE_COL_TELLERNUM).Value = wsSrcRef.Cells(r, RAW_TELLERNUM).Value2
        wsBaseRef.Cells(outRow, BASE_COL_ACTION).Value = wsSrcRef.Cells(r, RAW_ACTIONCOL).Value2
        wsBaseRef.Cells(outRow, BASE_COL_PREMIUM).Value = premVal
        wsBaseRef.Cells(outRow, BASE_COL_COMMISSION).Value = wsSrcRef.Cells(r, RAW_COMMISSION).Value2
        outRow = outRow + 1
NextRefRow:
    Next r
    countRef = outRow - 2
    wsBaseRef.Columns.AutoFit

    If Not wbRefExternal Is Nothing Then
        wbRefExternal.Close SaveChanges:=False
        Set wbRefExternal = Nothing
    End If

BUILD_COMPARISONS:
    debugStep = "BUILD_COMPARISONS"
    BuildComparisonSheet wsBase, wsBaseRef, SHEET_COMPANIES(), BASE_COL_COMPANY, minMonth, maxMonth, yearVal, refYear
    BuildComparisonSheet wsBase, wsBaseRef, SHEET_BRANCH(), BASE_COL_BRANCHNAME, minMonth, maxMonth, yearVal, refYear
    BuildComparisonSheet wsBase, wsBaseRef, SHEET_MAINBRANCH(), BASE_COL_MAINBRANCH, minMonth, maxMonth, yearVal, refYear
    BuildComparisonSheet wsBase, wsBaseRef, SHEET_TELLERS(), BASE_COL_TELLER, minMonth, maxMonth, yearVal, refYear
    BuildComparisonSheet wsBase, wsBaseRef, SHEET_AGENTS(), BASE_COL_AGENTNAME, minMonth, maxMonth, yearVal, refYear
    BuildComparisonSheet wsBase, wsBaseRef, SHEET_MONTHS(), BASE_COL_MONTH, minMonth, maxMonth, yearVal, refYear

    debugStep = "BUILD_SUMMARY"
    periodDesc = Trim$(CStr(wsMain.Range("G5").Value2))
    If Trim$(CStr(wsMain.Range("G6").Value2)) <> "" Then periodDesc = periodDesc & " " & Trim$(CStr(wsMain.Range("G6").Value2))
    BuildSummarySheet countRef, countCurrent, yearVal, refYear, periodDesc

    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalculation

    MsgBoxU ChrW(1492) & ChrW(1506) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1491) & " " & ChrW(1492) & ChrW(1505) & ChrW(1514) & ChrW(1497) & ChrW(1497) & ChrW(1501), vbInformation
    Exit Sub

ERR_HANDLER:
    On Error Resume Next
    If Not wbSrcExternal Is Nothing Then wbSrcExternal.Close SaveChanges:=False
    If Not wbRefExternal Is Nothing Then wbRefExternal.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    MsgBoxU "ApplyCorrections error (line " & Erl & ", step " & debugStep & "):" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub WriteBaseHeaders(ByVal ws As Worksheet)
    ws.Cells(1, BASE_COL_ID).Value = "ID"
    ws.Cells(1, BASE_COL_YEAR).Value = "Year"
    ws.Cells(1, BASE_COL_MONTH).Value = "Month"
    ws.Cells(1, BASE_COL_IDENTITY).Value = "Identity"
    ws.Cells(1, BASE_COL_CUSTOMER).Value = "Customer"
    ws.Cells(1, BASE_COL_CUSTNAME).Value = "CustName"
    ws.Cells(1, BASE_COL_POLICY).Value = "Policy"
    ws.Cells(1, BASE_COL_ADDENDUM).Value = "Addendum"
    ws.Cells(1, BASE_COL_COMPANY).Value = "Company"
    ws.Cells(1, BASE_COL_COMPNUM).Value = "CompNum"
    ws.Cells(1, BASE_COL_BRANCHNAME).Value = "BranchName"
    ws.Cells(1, BASE_COL_BRANCHNUM).Value = "BranchNum"
    ws.Cells(1, BASE_COL_MAINBRANCH).Value = "MainBranch"
    ws.Cells(1, BASE_COL_AGENTNAME).Value = "AgentName"
    ws.Cells(1, BASE_COL_AGENTNUM).Value = "AgentNum"
    ws.Cells(1, BASE_COL_TELLER).Value = "Teller"
    ws.Cells(1, BASE_COL_TELLERNUM).Value = "TellerNum"
    ws.Cells(1, BASE_COL_ACTION).Value = "Action"
    ws.Cells(1, BASE_COL_PREMIUM).Value = "Premium"
    ws.Cells(1, BASE_COL_COMMISSION).Value = "Commission"
    ws.Cells(1, BASE_COL_ISSUE).Value = "Issue"
    ws.Cells(1, BASE_COL_TOFIX).Value = "ToFix"
    ws.Rows(1).Font.Bold = True
End Sub

' ============================================================================
' BuildComparisonSheet — aggregates premiums/docs/customers/policies/commissions
' per group (company / branch / main branch / teller / agent / month) and writes
' a side-by-side year comparison sheet.
' ============================================================================
Private Sub BuildComparisonSheet(ByVal wsCurrent As Worksheet, ByVal wsRef As Worksheet, _
                                  ByVal sheetName As String, ByVal groupCol As Long, _
                                  ByVal minMonth As Long, ByVal maxMonth As Long, _
                                  ByVal yearVal As String, ByVal refYear As String)
    On Error GoTo ERR_HANDLER
    DeleteSheetIfExists sheetName
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = sheetName

    Dim dictKeys As Object
    Set dictKeys = CreateObject("Scripting.Dictionary")
    dictKeys.CompareMode = vbTextCompare

    Dim lastRowCur As Long, lastRowRef As Long, r As Long, k As String, m As Long
    Dim sortedMonths() As String, mIdx As Long, mCnt As Long, mi As Long
    Dim dictGlobalCustCur As Object, dictGlobalCustRef As Object
    Dim dictGlobalPolCur As Object, dictGlobalPolRef As Object

    lastRowCur = wsCurrent.Cells(wsCurrent.Rows.Count, 1).End(xlUp).Row
    lastRowRef = wsRef.Cells(wsRef.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRowCur
        m = 0
        If Not IsBlankValue(wsCurrent.Cells(r, BASE_COL_MONTH).Value2) Then
            m = CLng(wsCurrent.Cells(r, BASE_COL_MONTH).Value2)
        End If
        If m >= minMonth And m <= maxMonth Then
            k = Trim$(CStr(wsCurrent.Cells(r, groupCol).Value2))
            If k <> "" And LCase$(k) <> "(empty)" Then
                If Not dictKeys.Exists(k) Then dictKeys(k) = True
            End If
        End If
    Next r
    For r = 2 To lastRowRef
        m = 0
        If Not IsBlankValue(wsRef.Cells(r, BASE_COL_MONTH).Value2) Then
            m = CLng(wsRef.Cells(r, BASE_COL_MONTH).Value2)
        End If
        If m >= minMonth And m <= maxMonth Then
            k = Trim$(CStr(wsRef.Cells(r, groupCol).Value2))
            If k <> "" And LCase$(k) <> "(empty)" Then
                If Not dictKeys.Exists(k) Then dictKeys(k) = True
            End If
        End If
    Next r

    WriteComparisonHeaders wsOut, yearVal, refYear, sheetName

    Dim outRow As Long
    outRow = 3
    Dim allKeys As Variant

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
        allKeys = dictKeys.keys
    End If

    Dim idx As Long
    Dim premCur As Double, premRef As Double, commCur As Double, commRef As Double
    Dim docsCur As Long, docsRef As Long, custCur As Long, custRef As Long, polCur As Long, polRef As Long
    Dim dictCustCur As Object, dictCustRef As Object, dictPolCur As Object, dictPolRef As Object
    Dim custKey As String, polKey As String, actValCur As String, actValRef As String, gKey As Variant
    Dim totPremCur As Double, totPremRef As Double, totCommCur As Double, totCommRef As Double
    Dim totDocsCur As Long, totDocsRef As Long, totCustCur As Long, totCustRef As Long
    Dim totPolCur As Long, totPolRef As Long

    Set dictGlobalCustCur = CreateObject("Scripting.Dictionary")
    Set dictGlobalCustRef = CreateObject("Scripting.Dictionary")
    Set dictGlobalPolCur = CreateObject("Scripting.Dictionary")
    Set dictGlobalPolRef = CreateObject("Scripting.Dictionary")

    For idx = 0 To UBound(allKeys)
        k = allKeys(idx)
        premCur = 0: premRef = 0
        commCur = 0: commRef = 0
        docsCur = 0: docsRef = 0
        Set dictCustCur = CreateObject("Scripting.Dictionary")
        Set dictCustRef = CreateObject("Scripting.Dictionary")
        Set dictPolCur = CreateObject("Scripting.Dictionary")
        Set dictPolRef = CreateObject("Scripting.Dictionary")

        For r = 2 To lastRowCur
            m = 0
            If Not IsBlankValue(wsCurrent.Cells(r, BASE_COL_MONTH).Value2) Then
                m = CLng(wsCurrent.Cells(r, BASE_COL_MONTH).Value2)
            End If
            If m < minMonth Or m > maxMonth Then GoTo NextCurR
            If StrComp(Trim$(CStr(wsCurrent.Cells(r, groupCol).Value2)), k, vbTextCompare) = 0 Then
                premCur = premCur + CDbl(wsCurrent.Cells(r, BASE_COL_PREMIUM).Value2)
                commCur = commCur + CDbl(wsCurrent.Cells(r, BASE_COL_COMMISSION).Value2)
                docsCur = docsCur + 1
                custKey = Trim$(CStr(wsCurrent.Cells(r, BASE_COL_CUSTOMER).Value2))
                If custKey <> "" Then
                    actValCur = Trim$(CStr(wsCurrent.Cells(r, BASE_COL_ACTION).Value2))
                    If InStr(1, actValCur, ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1500), vbTextCompare) = 0 Then
                        If Not dictCustCur.Exists(custKey) Then dictCustCur(custKey) = True
                    End If
                End If
                polKey = Trim$(CStr(wsCurrent.Cells(r, BASE_COL_POLICY).Value2))
                If polKey <> "" Then
                    If Not dictPolCur.Exists(polKey) Then dictPolCur(polKey) = True
                End If
            End If
NextCurR:
        Next r

        For r = 2 To lastRowRef
            m = 0
            If Not IsBlankValue(wsRef.Cells(r, BASE_COL_MONTH).Value2) Then
                m = CLng(wsRef.Cells(r, BASE_COL_MONTH).Value2)
            End If
            If m < minMonth Or m > maxMonth Then GoTo NextRefR
            If StrComp(Trim$(CStr(wsRef.Cells(r, groupCol).Value2)), k, vbTextCompare) = 0 Then
                premRef = premRef + CDbl(wsRef.Cells(r, BASE_COL_PREMIUM).Value2)
                commRef = commRef + CDbl(wsRef.Cells(r, BASE_COL_COMMISSION).Value2)
                docsRef = docsRef + 1
                custKey = Trim$(CStr(wsRef.Cells(r, BASE_COL_CUSTOMER).Value2))
                If custKey <> "" Then
                    actValRef = Trim$(CStr(wsRef.Cells(r, BASE_COL_ACTION).Value2))
                    If InStr(1, actValRef, ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1500), vbTextCompare) = 0 Then
                        If Not dictCustRef.Exists(custKey) Then dictCustRef(custKey) = True
                    End If
                End If
                polKey = Trim$(CStr(wsRef.Cells(r, BASE_COL_POLICY).Value2))
                If polKey <> "" Then
                    If Not dictPolRef.Exists(polKey) Then dictPolRef(polKey) = True
                End If
            End If
NextRefR:
        Next r

        custCur = dictCustCur.Count
        custRef = dictCustRef.Count
        polCur = dictPolCur.Count
        polRef = dictPolRef.Count

        If groupCol = BASE_COL_MONTH Then
            wsOut.Cells(outRow, 1).Value = HebrewMonthName(CLng(k))
        Else
            wsOut.Cells(outRow, 1).Value = k
        End If
        wsOut.Cells(outRow, 2).Value = premRef
        wsOut.Cells(outRow, 3).Value = premCur
        wsOut.Cells(outRow, 4).Value = SafePct(premCur, premRef)
        wsOut.Cells(outRow, 5).Value = docsRef
        wsOut.Cells(outRow, 6).Value = docsCur
        wsOut.Cells(outRow, 7).Value = SafePct(docsCur, docsRef)
        wsOut.Cells(outRow, 8).Value = custRef
        wsOut.Cells(outRow, 9).Value = custCur
        wsOut.Cells(outRow, 10).Value = SafePct(custCur, custRef)
        wsOut.Cells(outRow, 11).Value = polRef
        wsOut.Cells(outRow, 12).Value = polCur
        wsOut.Cells(outRow, 13).Value = SafePct(polCur, polRef)
        wsOut.Cells(outRow, 14).Value = commRef
        wsOut.Cells(outRow, 15).Value = commCur
        wsOut.Cells(outRow, 16).Value = SafePct(commCur, commRef)

        totPremCur = totPremCur + premCur
        totPremRef = totPremRef + premRef
        totCommCur = totCommCur + commCur
        totCommRef = totCommRef + commRef
        totDocsCur = totDocsCur + docsCur
        totDocsRef = totDocsRef + docsRef
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

        outRow = outRow + 1
    Next idx

    totCustCur = dictGlobalCustCur.Count
    totCustRef = dictGlobalCustRef.Count
    totPolCur = dictGlobalPolCur.Count
    totPolRef = dictGlobalPolRef.Count

    ' "סה""כ"
    wsOut.Cells(outRow, 1).Value = ChrW(1505) & ChrW(1492) & Chr(34) & ChrW(1499)
    wsOut.Cells(outRow, 2).Value = totPremRef
    wsOut.Cells(outRow, 3).Value = totPremCur
    wsOut.Cells(outRow, 4).Value = SafePct(totPremCur, totPremRef)
    wsOut.Cells(outRow, 5).Value = totDocsRef
    wsOut.Cells(outRow, 6).Value = totDocsCur
    wsOut.Cells(outRow, 7).Value = SafePct(totDocsCur, totDocsRef)
    wsOut.Cells(outRow, 8).Value = totCustRef
    wsOut.Cells(outRow, 9).Value = totCustCur
    wsOut.Cells(outRow, 10).Value = SafePct(totCustCur, totCustRef)
    wsOut.Cells(outRow, 11).Value = totPolRef
    wsOut.Cells(outRow, 12).Value = totPolCur
    wsOut.Cells(outRow, 13).Value = SafePct(totPolCur, totPolRef)
    wsOut.Cells(outRow, 14).Value = totCommRef
    wsOut.Cells(outRow, 15).Value = totCommCur
    wsOut.Cells(outRow, 16).Value = SafePct(totCommCur, totCommRef)
    wsOut.Rows(outRow).Font.Bold = True

    wsOut.Columns.AutoFit
    wsOut.DisplayRightToLeft = True

    Dim col As Long
    For col = 2 To 16
        If col = 4 Or col = 7 Or col = 10 Or col = 13 Or col = 16 Then
            wsOut.Columns(col).NumberFormat = "0.0%"
        Else
            wsOut.Columns(col).NumberFormat = "#,##0"
        End If
    Next col
    Exit Sub

ERR_HANDLER:
    Err.Raise Err.Number, "BuildComparisonSheet(" & sheetName & "):" & Erl, Err.Description
End Sub

Private Sub WriteComparisonHeaders(ByVal ws As Worksheet, ByVal yearVal As String, ByVal refYear As String, ByVal sheetName As String)
    Dim blueColor As Long, lightBlueColor As Long, pctLabel As String
    blueColor = RGB(0, 70, 140)
    lightBlueColor = RGB(155, 200, 235)
    ws.Range("A1:A2").Merge
    ws.Cells(1, 1).Value = ChrW(1513) & ChrW(1501)
    ws.Cells(1, 1).HorizontalAlignment = xlCenter
    ws.Cells(1, 1).VerticalAlignment = xlCenter
    ws.Range(ws.Cells(1, 2), ws.Cells(1, 4)).Merge
    ws.Cells(1, 2).Value = ChrW(1508) & ChrW(1512) & ChrW(1493) & ChrW(1491) & ChrW(1493) & ChrW(1511) & ChrW(1510) & ChrW(1497) & ChrW(1492)
    ws.Cells(1, 2).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 5), ws.Cells(1, 7)).Merge
    ws.Cells(1, 5).Value = ChrW(1502) & ChrW(1505) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1501)
    ws.Cells(1, 5).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 8), ws.Cells(1, 10)).Merge
    ws.Cells(1, 8).Value = ChrW(1502) & ChrW(1489) & ChrW(1493) & ChrW(1496) & ChrW(1495) & ChrW(1497) & ChrW(1501)
    ws.Cells(1, 8).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 11), ws.Cells(1, 13)).Merge
    ws.Cells(1, 11).Value = ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1505) & ChrW(1493) & ChrW(1514)
    ws.Cells(1, 11).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 14), ws.Cells(1, 16)).Merge
    ws.Cells(1, 14).Value = ChrW(1506) & ChrW(1502) & ChrW(1500) & ChrW(1492)
    ws.Cells(1, 14).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 16)).Interior.Color = blueColor
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 16)).Font.Color = RGB(255, 255, 255)
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 16)).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 16)).Font.Size = 12
    pctLabel = ChrW(1513) & ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1497) & "%"
    ws.Cells(2, 1).Value = ""
    ws.Cells(2, 2).Value = refYear: ws.Cells(2, 3).Value = yearVal: ws.Cells(2, 4).Value = pctLabel
    ws.Cells(2, 5).Value = refYear: ws.Cells(2, 6).Value = yearVal: ws.Cells(2, 7).Value = pctLabel
    ws.Cells(2, 8).Value = refYear: ws.Cells(2, 9).Value = yearVal: ws.Cells(2, 10).Value = pctLabel
    ws.Cells(2, 11).Value = refYear: ws.Cells(2, 12).Value = yearVal: ws.Cells(2, 13).Value = pctLabel
    ws.Cells(2, 14).Value = refYear: ws.Cells(2, 15).Value = yearVal: ws.Cells(2, 16).Value = pctLabel
    ws.Range(ws.Cells(2, 1), ws.Cells(2, 16)).Interior.Color = lightBlueColor
    ws.Range(ws.Cells(2, 1), ws.Cells(2, 16)).Font.Bold = True
    ws.Range(ws.Cells(2, 1), ws.Cells(2, 16)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, 1), ws.Cells(2, 16)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(1, 1), ws.Cells(2, 16)).Borders.Weight = xlThin
End Sub

Private Function SafePct(ByVal newVal As Double, ByVal oldVal As Double) As Double
    If oldVal = 0 Then
        If newVal = 0 Then SafePct = 0 Else SafePct = 1
    Else
        SafePct = (newVal - oldVal) / Abs(oldVal)
    End If
End Function

Private Sub BuildSummarySheet(ByVal countRef As Long, ByVal countCurrent As Long, _
                              ByVal yearVal As String, ByVal refYear As String, ByVal periodDesc As String)
    DeleteSheetIfExists SHEET_SUMMARY()
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = SHEET_SUMMARY()
    wsOut.Cells(1, 1).Value = ChrW(1505) & ChrW(1497) & ChrW(1499) & ChrW(1493) & ChrW(1501) & " " & ChrW(1514) & ChrW(1492) & ChrW(1500) & ChrW(1497) & ChrW(1498)
    wsOut.Cells(1, 1).Font.Bold = True
    wsOut.Cells(1, 1).Font.Size = 14
    wsOut.Cells(3, 1).Value = ChrW(1508) & ChrW(1512) & ChrW(1496)
    wsOut.Cells(3, 2).Value = ChrW(1506) & ChrW(1512) & ChrW(1498)
    wsOut.Rows(3).Font.Bold = True
    wsOut.Cells(4, 1).Value = ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
    wsOut.Cells(4, 2).Value = periodDesc
    wsOut.Cells(5, 1).Value = ChrW(1513) & ChrW(1504) & ChrW(1514) & " " & ChrW(1497) & ChrW(1497) & ChrW(1495) & ChrW(1493) & ChrW(1505)
    wsOut.Cells(5, 2).Value = refYear
    wsOut.Cells(6, 1).Value = ChrW(1513) & ChrW(1504) & ChrW(1492) & " " & ChrW(1504) & ChrW(1489) & ChrW(1491) & ChrW(1511) & ChrW(1514)
    wsOut.Cells(6, 2).Value = yearVal
    wsOut.Cells(7, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & refYear
    wsOut.Cells(7, 2).Value = countRef
    wsOut.Cells(8, 1).Value = ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & yearVal
    wsOut.Cells(8, 2).Value = countCurrent
    wsOut.Columns(2).NumberFormat = "#,##0"
    wsOut.Columns.AutoFit
    wsOut.DisplayRightToLeft = True
End Sub

' ============================================================================
' SetupMainSheet — buttons starting from B3 (LevavClaude v1.0 layout)
' ============================================================================
' - Cleans up leftover labels in A2:A5 / D4 from older runs
' - Does NOT touch column G (preserves user's existing dropdowns at G5/G6/G7)
' - Creates 3 stacked buttons: B3:C4, B5:C6, B7:C8
' - Ensures ERROR_EMAIL parameter exists in הגדרות J/K
' - Writes Hebrew system messages in column S of הגדרות
' ============================================================================

Public Sub SetupMainSheet()
    Dim wsMain As Worksheet, wsMgmt As Worksheet
    Dim shp As Shape, s As Shape

    On Error GoTo ERR_HANDLER

    Set wsMain = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())

    ' Cleanup leftover column-A / D4 labels (from older v6.8 SetupMainSheet runs)
    On Error Resume Next
    wsMain.Range("A2:A5").ClearContents
    wsMain.Range("D4").ClearContents
    On Error GoTo ERR_HANDLER

    ' Remove old buttons if they exist
    On Error Resume Next
    For Each s In wsMain.Shapes
        If s.Name = "btnBuildReview" Or s.Name = "btnApplyCorrections" Or s.Name = "btnBuildPresentation" Then s.Delete
    Next s
    On Error GoTo ERR_HANDLER

    ' Button 1 — BuildReview (blue) at B3:C4
    With wsMain.Range("B3:C4")
        Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, .Left, .Top, .Width, .Height)
    End With
    shp.Name = "btnBuildReview"
    shp.Fill.ForeColor.RGB = RGB(0, 70, 140)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = "1 - " & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1514) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 14
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "BuildReview"

    ' Button 2 — ApplyCorrectionsAndBuildReports (green) at B5:C6
    With wsMain.Range("B5:C6")
        Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, .Left, .Top, .Width, .Height)
    End With
    shp.Name = "btnApplyCorrections"
    shp.Fill.ForeColor.RGB = RGB(0, 120, 60)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = "2 - " & ChrW(1497) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1501) & " " & ChrW(1493) & ChrW(1491) & ChrW(1493) & ChrW(34) & ChrW(1495) & ChrW(1493) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 14
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "ApplyCorrectionsAndBuildReports"

    ' Button 3 — BuildPresentation (brown) at B7:C8
    With wsMain.Range("B7:C8")
        Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, .Left, .Top, .Width, .Height)
    End With
    shp.Name = "btnBuildPresentation"
    shp.Fill.ForeColor.RGB = RGB(160, 80, 0)
    shp.Line.Visible = msoFalse
    shp.TextFrame2.TextRange.Text = "3 - " & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1493) & ChrW(1512) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.TextFrame2.TextRange.Font.Size = 14
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    shp.OnAction = "BuildPresentation"

    ' Ensure ERROR_EMAIL parameter exists in J/K
    Dim paramLastRow As Long, foundEmail As Boolean, pr As Long
    paramLastRow = wsMgmt.Cells(wsMgmt.Rows.Count, COL_PARAM_NAME).End(xlUp).Row
    foundEmail = False
    For pr = 1 To paramLastRow
        If UCase$(Trim$(CStr(wsMgmt.Cells(pr, COL_PARAM_NAME).Value2))) = "ERROR_EMAIL" Then
            foundEmail = True
            Exit For
        End If
    Next pr
    If Not foundEmail Then
        wsMgmt.Cells(paramLastRow + 1, COL_PARAM_NAME).Value = "ERROR_EMAIL"
        wsMgmt.Cells(paramLastRow + 1, COL_PARAM_VALUE).Value = "zvi@gorentech.co.il"
    End If

    MsgBoxU ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514) & " " & ChrW(1492) & ChrW(1493) & ChrW(1511) & ChrW(1501) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & "!", vbInformation
    Exit Sub

ERR_HANDLER:
    On Error Resume Next
    MsgBoxU "SetupMainSheet error: " & Err.Description, vbCritical
End Sub

' ============================================================================
' UpdatePeriodDropdown — disabled stub
' Period detail dropdown (G6) is managed by Sheet1.Worksheet_Change event
' (driven by named ranges rngPeriodType / rngPeriodValue). This stub exists
' so any leftover call to UpdatePeriodDropdown won't fail.
' ============================================================================
Public Sub UpdatePeriodDropdown()
    ' intentionally empty
End Sub

' ============================================================================
' BuildPresentation — stub (not yet ported to v1.0). Returns an info message.
' Will be ported in a later version.
' ============================================================================
Public Sub BuildPresentation()
    MsgBoxU ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1493) & ChrW(1512) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514) & " " & ChrW(1506) & ChrW(1491) & ChrW(1497) & ChrW(1497) & ChrW(1503) & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1502) & ChrW(1497) & ChrW(1493) & ChrW(1513) & ChrW(1501) & " " & ChrW(1489) & "-" & "v1.0" & "." & vbCrLf & "Build Presentation is not yet ported in this version.", vbInformation
End Sub

' ============================================================================
' SHEET NAME FUNCTIONS (Hebrew via ChrW so VBA Editor's Unicode quirks
' don't corrupt the strings)
' ============================================================================

' C:\פרויקט לבב\SOURCE\
Private Function SOURCE_FOLDER() As String
    SOURCE_FOLDER = "C:\" & ChrW(1508) & ChrW(1512) & ChrW(1493) & ChrW(1497) & ChrW(1511) & ChrW(1496) & " " & ChrW(1500) & ChrW(1489) & ChrW(1489) & "\SOURCE\"
End Function

' "דף הבית"
Private Function CONTROL_SHEET_NAME() As String
    CONTROL_SHEET_NAME = ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)
End Function

' "הגדרות"
Private Function MANAGEMENT_SHEET_NAME() As String
    MANAGEMENT_SHEET_NAME = ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514)
End Function

' "לטיפול"
Private Function REVIEW_SHEET_NAME() As String
    REVIEW_SHEET_NAME = ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500)
End Function

' "חברות"
Private Function SHEET_COMPANIES() As String
    SHEET_COMPANIES = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1514)
End Function

' "ענפים"
Private Function SHEET_BRANCH() As String
    SHEET_BRANCH = ChrW(1506) & ChrW(1504) & ChrW(1508) & ChrW(1497) & ChrW(1501)
End Function

' "ענף מרכז"
Private Function SHEET_MAINBRANCH() As String
    SHEET_MAINBRANCH = ChrW(1506) & ChrW(1504) & ChrW(1507) & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)
End Function

' "טלרים"
Private Function SHEET_TELLERS() As String
    SHEET_TELLERS = ChrW(1496) & ChrW(1500) & ChrW(1512) & ChrW(1497) & ChrW(1501)
End Function

' "סוכנים"
Private Function SHEET_AGENTS() As String
    SHEET_AGENTS = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501)
End Function

' "חודשים"
Private Function SHEET_MONTHS() As String
    SHEET_MONTHS = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
End Function

' "סיכום"
Private Function SHEET_SUMMARY() As String
    SHEET_SUMMARY = ChrW(1505) & ChrW(1497) & ChrW(1499) & ChrW(1493) & ChrW(1501)
End Function

