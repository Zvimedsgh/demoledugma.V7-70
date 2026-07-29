--- Windows API for Unicode MsgBox ---
#If VBA7 Then
    Public Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
#Else
    Public Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
#End If

' --- General constants ---
Public Const MANAGEMENT_START_ROW As Long = 2
Public Const DATA_SHEET_NAME As String = "TmpClientPolicyListEx"

' --- NIHUL field definition table ---
Public Const COL_FIELD_NAME_HE As Long = 5
Public Const COL_FIELD_COLUMN As Long = 6
Public Const COL_FIELD_CHECKING As Long = 7
Public Const COL_FIELD_KEY As Long = 8

' --- NIHUL parameter table ---
Public Const COL_PARAM_NAME As Long = 10
Public Const COL_PARAM_VALUE As Long = 11

' --- NIHUL helper translation table ---
Public Const COL_HELPER_KEY As Long = 14
Public Const COL_HELPER_VALUE As Long = 15

Public Const PARAM_PREMIUM_THRESHOLD As String = "PREMIUM_THRESHOLD"
Public Const PARAM_ERROR_EMAIL As String = "ERROR_EMAIL"
Public Const KEY_BRANCH_NAME As String = "BRANCH_NAME"
Public Const KEY_PREMIUM As String = "PREMIUM"
Public Const HELPER_REVIEW_SOURCE_ROW_HEADER As String = "REVIEW_SOURCE_ROW_HEADER"
Public Const HELPER_REVIEW_REASON_HEADER As String = "REVIEW_REASON_HEADER"
Public Const HELPER_REVIEW_REASON_CODE_HEADER As String = "REVIEW_REASON_CODE_HEADER"

' --- Raw source column mapping ---
Public Const RAW_CUSTOMER As Long = 1
Public Const RAW_CUSTNAME As Long = 2
Public Const RAW_POLICY As Long = 11
Public Const RAW_ADDENDUM As Long = 12
Public Const RAW_COMPNUM As Long = 13
Public Const RAW_COMPANY As Long = 14
Public Const RAW_BRANCHNUM As Long = 15
Public Const RAW_BRANCHNAME As Long = 16
Public Const RAW_INSURANCE_START As Long = 17
Public Const RAW_BORDEREU As Long = 19
Public Const RAW_AGENTNUM As Long = 20
Public Const RAW_AGENTNAME As Long = 21
Public Const RAW_TELLERNUM As Long = 24
Public Const RAW_TELLERNAME As Long = 25
Public Const RAW_PREMIUM As Long = 28
Public Const RAW_COMMISSION As Long = 32
Public Const RAW_CURRENCY As Long = 27
Public Const RAW_ACTIONCOL As Long = 39
Public Const RAW_IDNUMBER As Long = 45

' --- Base sheet columns ---
Public Const BASE_COL_ID As Long = 1
Public Const BASE_COL_YEAR As Long = 2
Public Const BASE_COL_MONTH As Long = 3
Public Const BASE_COL_IDENTITY As Long = 4
Public Const BASE_COL_CUSTOMER As Long = 5
Public Const BASE_COL_CUSTNAME As Long = 6
Public Const BASE_COL_POLICY As Long = 7
Public Const BASE_COL_ADDENDUM As Long = 8
Public Const BASE_COL_COMPANY As Long = 9
Public Const BASE_COL_COMPNUM As Long = 10
Public Const BASE_COL_BRANCHNAME As Long = 11
Public Const BASE_COL_BRANCHNUM As Long = 12
Public Const BASE_COL_MAINBRANCH As Long = 13
Public Const BASE_COL_AGENTNAME As Long = 14
Public Const BASE_COL_AGENTNUM As Long = 15
Public Const BASE_COL_TELLER As Long = 16
Public Const BASE_COL_TELLERNUM As Long = 17
Public Const BASE_COL_ACTION As Long = 18
Public Const BASE_COL_PREMIUM As Long = 19
Public Const BASE_COL_COMMISSION As Long = 20
Public Const BASE_COL_ISSUE As Long = 21
Public Const BASE_COL_TOFIX As Long = 22

Public Const MB_RTLREADING As Long = &H100000
Public Const MB_RIGHT As Long = &H80000
Public Const MB_SYSTEMMODAL As Long = &H1000

' --- Sheet name functions ---
Public Function CONTROL_SHEET_NAME() As String
    CONTROL_SHEET_NAME = ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)
End Function

Public Function MANAGEMENT_SHEET_NAME() As String
    MANAGEMENT_SHEET_NAME = ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514)
End Function

Public Function REVIEW_SHEET_NAME() As String
    REVIEW_SHEET_NAME = ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500)
End Function

Public Function SOURCE_FOLDER() As String
    On Error Resume Next
    SOURCE_FOLDER = Trim$(CStr(ThisWorkbook.Names("rngFILES_FOLDER").RefersToRange.Value2))
    On Error GoTo 0
    If SOURCE_FOLDER = "" Then
        SOURCE_FOLDER = "C:\" & ChrW(1508) & ChrW(1512) & ChrW(1493) & ChrW(1497) & ChrW(1511) & ChrW(1496) & " " & ChrW(1500) & ChrW(1489) & ChrW(1489) & "\SOURCE\"
    End If
    If Right$(SOURCE_FOLDER, 1) <> "\" Then SOURCE_FOLDER = SOURCE_FOLDER & "\"
End Function

Public Function SHEET_COMPANIES() As String
    SHEET_COMPANIES = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1514)
End Function

Public Function SHEET_BRANCH() As String
    SHEET_BRANCH = ChrW(1506) & ChrW(1504) & ChrW(1508) & ChrW(1497) & ChrW(1501)
End Function

Public Function SHEET_MAINBRANCH() As String
    SHEET_MAINBRANCH = ChrW(1506) & ChrW(1504) & ChrW(1507) & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)
End Function

Public Function SHEET_TELLERS() As String
    SHEET_TELLERS = ChrW(1496) & ChrW(1500) & ChrW(1512) & ChrW(1497) & ChrW(1501)
End Function

Public Function SHEET_AGENTS() As String
    SHEET_AGENTS = ChrW(1505) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
End Function

Public Function SHEET_MONTHS() As String
    SHEET_MONTHS = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
End Function

Public Function SHEET_SUMMARY() As String
    SHEET_SUMMARY = ChrW(1505) & ChrW(1497) & ChrW(1499) & ChrW(1493) & ChrW(1501)
End Function

' --- Core helpers ---
Public Function MsgBoxU(ByVal sText As String, Optional ByVal uType As Long = 0, Optional ByVal sCaption As String = "") As Long
    MsgBoxU = MessageBoxW(0, StrPtr(sText), StrPtr(sCaption), uType Or MB_RTLREADING Or MB_RIGHT Or MB_SYSTEMMODAL)
End Function

Public Function SheetExists(ByVal sheetName As String) As Boolean
    On Error GoTo NOT_FOUND
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = True
    Exit Function
NOT_FOUND:
    SheetExists = False
End Function

Public Sub DeleteSheetIfExists(ByVal sName As String)
    Dim wsTarget As Worksheet
    Dim prevAlerts As Boolean
    Dim wsCtrl As Worksheet
    If Not SheetExists(sName) Then Exit Sub
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wsCtrl = ThisWorkbook.Worksheets(CONTROL_SHEET_NAME())
    If Not wsCtrl Is Nothing Then wsCtrl.Visible = xlSheetVisible
    On Error GoTo 0
    Set wsTarget = ThisWorkbook.Worksheets(sName)
    On Error Resume Next
    If wsTarget.Visible <> xlSheetVisible Then wsTarget.Visible = xlSheetVisible
    On Error GoTo 0
    On Error Resume Next
    wsTarget.Delete
    On Error GoTo 0
    If SheetExists(sName) Then
        On Error Resume Next
        ThisWorkbook.Worksheets(sName).Visible = xlSheetVisible
        ThisWorkbook.Worksheets(sName).Delete
        On Error GoTo 0
    End If
    Application.DisplayAlerts = prevAlerts
End Sub

Public Function HebrewMonthName(ByVal m As Long) As String
    If m = 1 Then
        HebrewMonthName = ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1488) & ChrW(1512)
    ElseIf m = 2 Then
        HebrewMonthName = ChrW(1508) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1488) & ChrW(1512)
    ElseIf m = 3 Then
        HebrewMonthName = ChrW(1502) & ChrW(1512) & ChrW(1509)
    ElseIf m = 4 Then
        HebrewMonthName = ChrW(1488) & ChrW(1508) & ChrW(1512) & ChrW(1497) & ChrW(1500)
    ElseIf m = 5 Then
        HebrewMonthName = ChrW(1502) & ChrW(1488) & ChrW(1497)
    ElseIf m = 6 Then
        HebrewMonthName = ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1497)
    ElseIf m = 7 Then
        HebrewMonthName = ChrW(1497) & ChrW(1493) & ChrW(1500) & ChrW(1497)
    ElseIf m = 8 Then
        HebrewMonthName = ChrW(1488) & ChrW(1493) & ChrW(1490) & ChrW(1505) & ChrW(1496)
    ElseIf m = 9 Then
        HebrewMonthName = ChrW(1505) & ChrW(1508) & ChrW(1496) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    ElseIf m = 10 Then
        HebrewMonthName = ChrW(1488) & ChrW(1493) & ChrW(1511) & ChrW(1496) & ChrW(1493) & ChrW(1489) & ChrW(1512)
    ElseIf m = 11 Then
        HebrewMonthName = ChrW(1504) & ChrW(1493) & ChrW(1489) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    ElseIf m = 12 Then
        HebrewMonthName = ChrW(1491) & ChrW(1510) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    Else
        HebrewMonthName = CStr(m)
    End If
End Function

Public Function IsBlankValue(ByVal v As Variant) As Boolean
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

Public Function TryParseVariantNumber(ByVal v As Variant, ByRef result As Double) As Boolean
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

Public Function GetNumericParameter(ByVal ws As Worksheet, ByVal paramName As String) As Double
    On Error GoTo ERR_HANDLER
    Dim r As Long, lastRow As Long, nm As String, v As Variant, n As Double
    lastRow = ws.Cells(ws.Rows.Count, COL_PARAM_NAME).End(xlUp).Row
    For r = 1 To lastRow
        nm = UCase$(Trim$(CStr(ws.Cells(r, COL_PARAM_NAME).Value2)))
        If nm = UCase$(paramName) Then
            v = ws.Cells(r, COL_PARAM_VALUE).Value2
            If TryParseVariantNumber(v, n) Then
                GetNumericParameter = n
            Else
                Err.Raise vbObjectError + 3000, "GetNumericParameter", "PARAMETER NOT NUMERIC: " & paramName
            End If
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 3001, "GetNumericParameter", "PARAMETER NOT FOUND: " & paramName
ERR_HANDLER:
    Err.Raise Err.Number, "GetNumericParameter:" & Erl, Err.Description
End Function

Public Function GetStringParameter(ByVal ws As Worksheet, ByVal paramName As String) As String
    On Error GoTo ERR_HANDLER
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
ERR_HANDLER:
    GetStringParameter = ""
End Function

Public Function ColumnLetterToNumber(ByVal col As String) As Long
    On Error GoTo ERR_HANDLER
    Dim i As Long, ch As String
    For i = 1 To Len(col)
        ch = Mid$(col, i, 1)
        If ch < "A" Or ch > "Z" Then
            Err.Raise vbObjectError + 4000, "ColumnLetterToNumber", "INVALID COLUMN LETTER: " & col
        End If
        ColumnLetterToNumber = ColumnLetterToNumber * 26 + (Asc(ch) - 64)
    Next i
    Exit Function
ERR_HANDLER:
    Err.Raise Err.Number, "ColumnLetterToNumber:" & Erl, Err.Description
End Function

Public Function HebrewToKey(ByVal hebText As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As Long
    Dim mapped As String
    
    hebText = Trim$(hebText)
    result = ""
    
    For i = 1 To Len(hebText)
        ch = AscW(Mid$(hebText, i, 1))
        Select Case ch
            Case 1488: mapped = "A"   ' א alef
            Case 1489: mapped = "B"   ' ב bet
            Case 1490: mapped = "G"   ' ג gimel
            Case 1491: mapped = "D"   ' ד dalet
            Case 1492: mapped = "H"   ' ה he
            Case 1493: mapped = "V"   ' ו vav
            Case 1494: mapped = "Z"   ' ז zayin
            Case 1495: mapped = "CH"  ' ח chet
            Case 1496: mapped = "T"   ' ט tet
            Case 1497: mapped = "Y"   ' י yod
            Case 1498: mapped = "K"   ' ך kaf sofit
            Case 1499: mapped = "K"   ' כ kaf
            Case 1500: mapped = "L"   ' ל lamed
            Case 1501: mapped = "M"   ' ם mem sofit
            Case 1502: mapped = "M"   ' מ mem
            Case 1503: mapped = "N"   ' ן nun sofit
            Case 1504: mapped = "N"   ' נ nun
            Case 1505: mapped = "S"   ' ס samech
            Case 1506: mapped = "A"   ' ע ayin
            Case 1507: mapped = "P"   ' ף pe sofit
            Case 1508: mapped = "P"   ' פ pe
            Case 1509: mapped = "TZ"  ' ץ tsade sofit
            Case 1510: mapped = "TZ"  ' צ tsade
            Case 1511: mapped = "K"   ' ק kof
            Case 1512: mapped = "R"   ' ר resh
            Case 1513: mapped = "SH"  ' ש shin
            Case 1514: mapped = "T"   ' ת tav
            Case 32:   mapped = "_"   ' space -> underscore
            Case Else
                If (ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Or (ch >= 48 And ch <= 57) Then
                    mapped = UCase$(Chr$(ch))
                Else
                    mapped = ""
                End If
        End Select
        result = result & mapped
    Next i
    
    Do While InStr(result, "__") > 0
        result = Replace(result, "__", "_")
    Loop
    If Left$(result, 1) = "_" Then result = Mid$(result, 2)
    If Right$(result, 1) = "_" Then result = Left$(result, Len(result) - 1)
    
    HebrewToKey = result
End Function

Public Function ShortenCompanyName(ByVal sName As String) As String
    Dim sSuffix As String
    Dim pos As Long
    ' חברה לביטוח בע"מ
    sSuffix = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1492) & " " & ChrW(1500) & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495) & " " & ChrW(1489) & ChrW(1506) & ChrW(34) & ChrW(1502)
    pos = InStr(1, sName, sSuffix, vbTextCompare)
    If pos > 1 Then
        ShortenCompanyName = Trim$(Left$(sName, pos - 1))
        Exit Function
    End If
    ' לביטוח בע"מ
    sSuffix = ChrW(1500) & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495) & " " & ChrW(1489) & ChrW(1506) & ChrW(34) & ChrW(1502)
    pos = InStr(1, sName, sSuffix, vbTextCompare)
    If pos > 1 Then
        ShortenCompanyName = Trim$(Left$(sName, pos - 1))
        Exit Function
    End If
    ' בע"מ alone at end
    sSuffix = " " & ChrW(1489) & ChrW(1506) & ChrW(34) & ChrW(1502)
    If Right$(sName, Len(sSuffix)) = sSuffix Then
        ShortenCompanyName = Trim$(Left$(sName, Len(sName) - Len(sSuffix)))
        Exit Function
    End If
    ShortenCompanyName = sName
End Function


' ============================================================================
' UTILITY: Find source file for a given year
' Looks in the current workbook's directory
' ============================================================================
Public Function FindSourceFile(ByVal yr As String) As String
    Dim fName As String
    Dim fPath As String
    fName = yr & ".xlsx"
    fPath = ThisWorkbook.Path & "\" & fName
    If CreateObject("Scripting.FileSystemObject").FileExists(fPath) Then
        FindSourceFile = fPath
    Else
        FindSourceFile = ""
    End If
End Function


' ============================================================================
' UTILITY: Open the data sheet from the source workbook
' ============================================================================
Public Function OpenDataSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set OpenDataSheet = wb.Worksheets(1)
End Function


' ============================================================================
' UTILITY: Get month range from Main sheet selection (rngPeriodValue)
' ============================================================================
Public Sub GetMonthRange(ByVal wsMain As Worksheet, ByRef minM As Long, ByRef maxM As Long)
    Dim pVal As String
    Dim pType As String
    pType = Trim$(CStr(wsMain.Range("rngPeriodType").Value2))
    pVal = Trim$(CStr(wsMain.Range("rngPeriodValue").Value2))
    
    ' Default: all months
    minM = 1
    maxM = 12
    
    If pType = "" Or pType = ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497) Then Exit Sub
    
    ' Yearly
    If InStr(1, pType, ChrW(1513) & ChrW(1504) & ChrW(1514) & ChrW(1497), vbTextCompare) > 0 And _
       InStr(1, pType, ChrW(1495) & ChrW(1510) & ChrW(1497), vbTextCompare) = 0 Then
        Exit Sub
    End If
    
    ' Half yearly
    If InStr(1, pType, ChrW(1495) & ChrW(1510) & ChrW(1497), vbTextCompare) > 0 Then
        If InStr(1, pVal, ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1504) & ChrW(1492), vbTextCompare) > 0 Then
            maxM = 6
        Else
            minM = 7
        End If
        Exit Sub
    End If
    
    ' Quarterly
    If InStr(1, pType, ChrW(1512) & ChrW(1489) & ChrW(1506) & ChrW(1493) & ChrW(1504) & ChrW(1497), vbTextCompare) > 0 Then
        If InStr(1, pVal, ChrW(1512) & ChrW(1488) & ChrW(1513) & ChrW(1493) & ChrW(1503), vbTextCompare) > 0 Then
            maxM = 3
        ElseIf InStr(1, pVal, ChrW(1513) & ChrW(1504) & ChrW(1497), vbTextCompare) > 0 Then
            minM = 4: maxM = 6
        ElseIf InStr(1, pVal, ChrW(1513) & ChrW(1500) & ChrW(1497) & ChrW(1513) & ChrW(1497), vbTextCompare) > 0 Then
            minM = 7: maxM = 9
        Else
            minM = 10: maxM = 12
        End If
        Exit Sub
    End If
    
    ' Monthly
    If InStr(1, pType, ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497), vbTextCompare) > 0 Then
        minM = MonthNumberFromHebrew(pVal)
        maxM = minM
    End If
End Sub


' ============================================================================
' UTILITY: Get date column number from Main sheet (rngDateType)
' ============================================================================
Public Function GetDateColumn(ByVal wsMain As Worksheet) As Long
    Dim dType As String
    dType = Trim$(CStr(wsMain.Range("rngDateType").Value2))
    ' Default: RAW_DATE_BORDEREU (6)
    GetDateColumn = 6
    ' "taarih hafaka" = 8
    If InStr(1, dType, ChrW(1492) & ChrW(1508) & ChrW(1511) & ChrW(1492), vbTextCompare) > 0 Then
        GetDateColumn = 8
    End If
End Function


' ============================================================================
' UTILITY: Hebrew month name to number
' ============================================================================
Private Function MonthNumberFromHebrew(ByVal mName As String) As Long
    If InStr(1, mName, ChrW(1497) & ChrW(1504) & ChrW(1493), vbTextCompare) > 0 Then MonthNumberFromHebrew = 1
    If InStr(1, mName, ChrW(1508) & ChrW(1489) & ChrW(1512) & ChrW(1493), vbTextCompare) > 0 Then MonthNumberFromHebrew = 2
    If InStr(1, mName, ChrW(1502) & ChrW(1512) & ChrW(1509), vbTextCompare) > 0 Then MonthNumberFromHebrew = 3
    If InStr(1, mName, ChrW(1488) & ChrW(1508) & ChrW(1512) & ChrW(1497), vbTextCompare) > 0 Then MonthNumberFromHebrew = 4
    If InStr(1, mName, ChrW(1502) & ChrW(1488) & ChrW(1497), vbTextCompare) > 0 Then MonthNumberFromHebrew = 5
    If InStr(1, mName, ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1497), vbTextCompare) > 0 Then MonthNumberFromHebrew = 6
    If InStr(1, mName, ChrW(1497) & ChrW(1493) & ChrW(1500) & ChrW(1497), vbTextCompare) > 0 Then MonthNumberFromHebrew = 7
    If InStr(1, mName, ChrW(1488) & ChrW(1493) & ChrW(1490) & ChrW(1493), vbTextCompare) > 0 Then MonthNumberFromHebrew = 8
    If InStr(1, mName, ChrW(1505) & ChrW(1508) & ChrW(1496) & ChrW(1502), vbTextCompare) > 0 Then MonthNumberFromHebrew = 9
    If InStr(1, mName, ChrW(1488) & ChrW(1493) & ChrW(1511) & ChrW(1496) & ChrW(1493), vbTextCompare) > 0 Then MonthNumberFromHebrew = 10
    If InStr(1, mName, ChrW(1504) & ChrW(1493) & ChrW(1489) & ChrW(1502), vbTextCompare) > 0 Then MonthNumberFromHebrew = 11
    If InStr(1, mName, ChrW(1491) & ChrW(1510) & ChrW(1502) & ChrW(1489), vbTextCompare) > 0 Then MonthNumberFromHebrew = 12
End Function


' ============================================================================
' UTILITY: Return Hebrew month name for month number 1-12
' ============================================================================
Public Function HebrewMonthName(ByVal m As Long) As String
    If m = 1 Then
        HebrewMonthName = ChrW(1497) & ChrW(1504) & ChrW(1493) & ChrW(1488) & ChrW(1512)
    ElseIf m = 2 Then
        HebrewMonthName = ChrW(1508) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1488) & ChrW(1512)
    ElseIf m = 3 Then
        HebrewMonthName = ChrW(1502) & ChrW(1512) & ChrW(1509)
    ElseIf m = 4 Then
        HebrewMonthName = ChrW(1488) & ChrW(1508) & ChrW(1512) & ChrW(1497) & ChrW(1500)
    ElseIf m = 5 Then
        HebrewMonthName = ChrW(1502) & ChrW(1488) & ChrW(1497)
    ElseIf m = 6 Then
        HebrewMonthName = ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1497)
    ElseIf m = 7 Then
        HebrewMonthName = ChrW(1497) & ChrW(1493) & ChrW(1500) & ChrW(1497)
    ElseIf m = 8 Then
        HebrewMonthName = ChrW(1488) & ChrW(1493) & ChrW(1490) & ChrW(1493) & ChrW(1505) & ChrW(1496)
    ElseIf m = 9 Then
        HebrewMonthName = ChrW(1505) & ChrW(1508) & ChrW(1496) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    ElseIf m = 10 Then
        HebrewMonthName = ChrW(1488) & ChrW(1493) & ChrW(1511) & ChrW(1496) & ChrW(1493) & ChrW(1489) & ChrW(1512)
    ElseIf m = 11 Then
        HebrewMonthName = ChrW(1504) & ChrW(1493) & ChrW(1489) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    ElseIf m = 12 Then
        HebrewMonthName = ChrW(1491) & ChrW(1510) & ChrW(1502) & ChrW(1489) & ChrW(1512)
    Else
        HebrewMonthName = CStr(m)
    End If
End Function


