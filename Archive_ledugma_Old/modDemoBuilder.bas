Attribute VB_Name = "modDemoBuilder"
' ============================================================================
' MODULE: modDemoBuilder
' PURPOSE: ONE-TIME UTILITY to build anonymized demo source files (.xlsx)
'          from real source files. Independent of modLevav -- duplicates the
'          column constants it needs.
' VERSION: 1.0
' DATE: 2026-05-07
' ============================================================================
' USAGE:
'   1. Import this module into the workbook (Alt+F11 -> File -> Import File)
'   2. Make sure FILES_FOLDER points to your REAL source (e.g., K8 = "Levav")
'   3. Run BuildDemoSources from Alt+F8
'   4. Two new files appear in the demo folder: 2016.xlsx and 2017.xlsx
'   5. REMOVE this module from the workbook (right-click -> Remove)
'   6. Set K8 = "Demo" to switch the production code to read demo data
' ============================================================================
' WHAT IT ANONYMIZES (per the agreed plan):
'   - Agent name + number     -> "souchen 1, 2, ..." (consistent per real number, both years)
'   - Teller name + number    -> "tler 1, 2, ..." (consistent per real number, both years)
'   - Customer name           -> "lakouach 1, 2, ..." (consistent per real customer ID)
'   - Customer ID number      -> blanked
'   - Policy number           -> "POL-000001, 000002, ..." (consistent across both years)
'   - Premium + Commission    -> multiplied by per-agent factor in [0.85, 1.30]
'   - Bordereu date           -> shifted: 2024 -> 2016, 2025 -> 2017
'   - Insurance start date    -> shifted: 2024 -> 2016, 2025 -> 2017
'   - Companies + branches    -> UNCHANGED
' ============================================================================

' --- Source column indexes (must match modLevav RAW_* constants) ---
Private Const D_RAW_CUSTOMER As Long = 1
Private Const D_RAW_CUSTNAME As Long = 2
Private Const D_RAW_POLICY As Long = 11
Private Const D_RAW_BRANCHNAME As Long = 16
Private Const D_RAW_INSURANCE_START As Long = 17
Private Const D_RAW_BORDEREU As Long = 19
Private Const D_RAW_AGENTNUM As Long = 20
Private Const D_RAW_AGENTNAME As Long = 21
Private Const D_RAW_TELLERNUM As Long = 24
Private Const D_RAW_TELLERNAME As Long = 25
Private Const D_RAW_PREMIUM As Long = 28
Private Const D_RAW_COMMISSION As Long = 32
Private Const D_RAW_IDNUMBER As Long = 45

Private Const D_DATA_SHEET_NAME As String = "TmpClientPolicyListEx"
Private Const D_PARAM_NAME_COL As Long = 10
Private Const D_PARAM_VALUE_COL As Long = 11

' --- Year mapping (real -> demo) ---
Private Const D_REAL_YEAR_1 As String = "2024"
Private Const D_DEMO_YEAR_1 As String = "2016"
Private Const D_REAL_YEAR_2 As String = "2025"
Private Const D_DEMO_YEAR_2 As String = "2017"


' ============================================================================
' MAIN ENTRY: BuildDemoSources
' ============================================================================
Public Sub BuildDemoSources()

    Dim realFolder As String, demoFolder As String

    ' --- Suggest source folder from FILES_FOLDER parameter (read live) ---
    realFolder = ReadParam("FILES_FOLDER")
    If realFolder = "" Then realFolder = "C:\LEVAV PROJECT\SOURCE\"

    Dim defaultTarget As String
    defaultTarget = "C:\DEMO PROJECT\SOURCE\"

    ' --- Confirm both paths via InputBox ---
    Dim ans As String
    ans = InputBox(ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1514) & " " & ChrW(1502) & ChrW(1511) & ChrW(1493) & ChrW(1512) & " " & "(" & ChrW(1488) & ChrW(1502) & ChrW(1497) & ChrW(1514) & ChrW(1497) & ChrW(1514) & "):" & vbCrLf & _
        ChrW(1502) & ChrW(1492) & ChrW(1497) & ChrW(1499) & ChrW(1503) & " " & ChrW(1500) & ChrW(1511) & ChrW(1512) & ChrW(1488) & " " & "2024.xlsx" & " " & ChrW(1493) & "-" & " " & "2025.xlsx", _
        "BuildDemoSources - " & ChrW(1502) & ChrW(1511) & ChrW(1493) & ChrW(1512), realFolder)
    If ans = "" Then Exit Sub
    realFolder = EnsureTrailingSlash(ans)

    ans = InputBox(ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1514) & " " & ChrW(1497) & ChrW(1506) & ChrW(1491) & " " & "(" & ChrW(1491) & ChrW(1502) & ChrW(1493) & "):" & vbCrLf & _
        ChrW(1500) & ChrW(1488) & ChrW(1503) & " " & ChrW(1500) & ChrW(1499) & ChrW(1514) & ChrW(1493) & ChrW(1489) & " " & "2016.xlsx" & " " & ChrW(1493) & "-" & " " & "2017.xlsx", _
        "BuildDemoSources - " & ChrW(1497) & ChrW(1506) & ChrW(1491), defaultTarget)
    If ans = "" Then Exit Sub
    demoFolder = EnsureTrailingSlash(ans)

    If StrComp(realFolder, demoFolder, vbTextCompare) = 0 Then
        MsgBox ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1514) & " " & ChrW(1502) & ChrW(1511) & ChrW(1493) & ChrW(1512) & " " & ChrW(1493) & ChrW(1497) & ChrW(1506) & ChrW(1491) & " " & ChrW(1494) & ChrW(1492) & ChrW(1493) & ChrW(1514) & ".", vbCritical
        Exit Sub
    End If

    ' --- Ensure target folder exists ---
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(demoFolder) Then
        On Error Resume Next
        ' Try to create (only works if parent exists)
        fso.CreateFolder demoFolder
        On Error GoTo 0
        If Not fso.FolderExists(demoFolder) Then
            MsgBox ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1497) & ChrW(1514) & ChrW(1503) & " " & ChrW(1500) & ChrW(1497) & ChrW(1510) & ChrW(1493) & ChrW(1512) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1514) & " " & ChrW(1492) & ChrW(1497) & ChrW(1506) & ChrW(1491) & ":" & vbCrLf & demoFolder & vbCrLf & vbCrLf & _
                ChrW(1497) & ChrW(1510) & ChrW(1512) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1514) & ChrW(1497) & ChrW(1511) & ChrW(1497) & ChrW(1497) & ChrW(1514) & " " & ChrW(1502) & ChrW(1497) & ChrW(1491) & ChrW(1504) & ChrW(1497) & ChrW(1514) & " " & ChrW(1493) & ChrW(1492) & ChrW(1512) & ChrW(1509) & " " & ChrW(1513) & ChrW(1493) & ChrW(1489) & ".", vbCritical
            Exit Sub
        End If
    End If

    On Error GoTo ERR_HANDLER

    ' --- Save Application state ---
    Dim prevSU As Boolean, prevDA As Boolean, prevCalc As XlCalculation
    prevSU = Application.ScreenUpdating
    prevDA = Application.DisplayAlerts
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' --- Shared anonymization dicts (used across both years -> consistent labels) ---
    Dim dictAgents As Object: Set dictAgents = CreateObject("Scripting.Dictionary")
    Dim dictTellers As Object: Set dictTellers = CreateObject("Scripting.Dictionary")
    Dim dictCustomers As Object: Set dictCustomers = CreateObject("Scripting.Dictionary")
    Dim dictPolicies As Object: Set dictPolicies = CreateObject("Scripting.Dictionary")
    Dim dictAgentFactor As Object: Set dictAgentFactor = CreateObject("Scripting.Dictionary")

    Dim totalRows As Long
    totalRows = 0
    totalRows = totalRows + AnonymizeOne(D_REAL_YEAR_1, D_DEMO_YEAR_1, realFolder, demoFolder, _
                                          dictAgents, dictTellers, dictCustomers, dictPolicies, dictAgentFactor)
    totalRows = totalRows + AnonymizeOne(D_REAL_YEAR_2, D_DEMO_YEAR_2, realFolder, demoFolder, _
                                          dictAgents, dictTellers, dictCustomers, dictPolicies, dictAgentFactor)

    ' --- Restore application state ---
    Application.ScreenUpdating = prevSU
    Application.DisplayAlerts = prevDA
    Application.Calculation = prevCalc

    MsgBox "BuildDemoSources " & ChrW(1505) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1489) & ChrW(1492) & ChrW(1510) & ChrW(1500) & ChrW(1495) & ChrW(1492) & vbCrLf & vbCrLf & _
           ChrW(1497) & ChrW(1497) & ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1497) & ChrW(1497) & ChrW(1501) & " " & ChrW(1489) & ChrW(1497) & ChrW(1503) & " " & ChrW(1513) & ChrW(1514) & ChrW(1497) & " " & ChrW(1492) & ChrW(1513) & ChrW(1504) & ChrW(1497) & ChrW(1501) & ":" & vbCrLf & _
           ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501) & ": " & dictAgents.Count & vbCrLf & _
           ChrW(1496) & ChrW(1500) & ChrW(1512) & ChrW(1497) & ChrW(1501) & ": " & dictTellers.Count & vbCrLf & _
           ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & ": " & dictCustomers.Count & vbCrLf & _
           ChrW(1508) & ChrW(1493) & ChrW(1500) & ChrW(1497) & ChrW(1505) & ChrW(1493) & ChrW(1514) & ": " & dictPolicies.Count & vbCrLf & vbCrLf & _
           ChrW(1505) & ChrW(1499) & ChrW(34) & ChrW(1499) & " " & ChrW(1513) & ChrW(1493) & ChrW(1512) & ChrW(1493) & ChrW(1514) & " " & ChrW(1513) & ChrW(1504) & ChrW(1499) & ChrW(1514) & ChrW(1489) & ChrW(1493) & ": " & totalRows & vbCrLf & vbCrLf & _
           ChrW(1499) & ChrW(1506) & ChrW(1514) & " " & ChrW(1504) & ChrW(1497) & ChrW(1514) & ChrW(1503) & " " & ChrW(1500) & ChrW(1492) & ChrW(1505) & ChrW(1497) & ChrW(1512) & " " & ChrW(1488) & ChrW(1514) & " " & ChrW(1492) & ChrW(1502) & ChrW(1493) & ChrW(1491) & ChrW(1493) & ChrW(1500) & " " & "modDemoBuilder" & ".", vbInformation

    Exit Sub

ERR_HANDLER:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox ChrW(1513) & ChrW(1490) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1489) & "BuildDemoSources" & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub


' ============================================================================
' Process one source file: read real, anonymize in memory, write demo file.
' Returns the number of data rows written (lastRow - 1).
' ============================================================================
Private Function AnonymizeOne(ByVal realYear As String, ByVal demoYear As String, _
        ByVal realFolder As String, ByVal demoFolder As String, _
        ByVal dictAgents As Object, ByVal dictTellers As Object, _
        ByVal dictCustomers As Object, ByVal dictPolicies As Object, _
        ByVal dictAgentFactor As Object) As Long

    On Error GoTo ERR_HANDLER

    ' --- Locate source file (.xlsx, fallback .xls) ---
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim srcPath As String
    srcPath = realFolder & realYear & ".xlsx"
    If Not fso.FileExists(srcPath) Then srcPath = realFolder & realYear & ".xls"
    If Not fso.FileExists(srcPath) Then
        MsgBox ChrW(1511) & ChrW(1493) & ChrW(1489) & ChrW(1509) & " " & ChrW(1502) & ChrW(1511) & ChrW(1493) & ChrW(1512) & " " & ChrW(1500) & ChrW(1488) & " " & ChrW(1504) & ChrW(1502) & ChrW(1510) & ChrW(1488) & ":" & vbCrLf & realFolder & realYear & ".xlsx", vbExclamation
        AnonymizeOne = 0
        Exit Function
    End If

    ' --- Open source ---
    Dim wbSrc As Workbook
    Set wbSrc = Workbooks.Open(srcPath, ReadOnly:=True, UpdateLinks:=0)
    Dim wsSrc As Worksheet
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(D_DATA_SHEET_NAME)
    On Error GoTo 0
    If wsSrc Is Nothing Then Set wsSrc = wbSrc.Worksheets(1)

    Dim lastRow As Long, lastCol As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    If lastCol < D_RAW_IDNUMBER Then lastCol = D_RAW_IDNUMBER

    If lastRow < 2 Then
        wbSrc.Close SaveChanges:=False
        AnonymizeOne = 0
        Exit Function
    End If

    ' --- Bulk read entire range (header + data) ---
    Dim arr As Variant
    arr = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).Value2

    wbSrc.Close SaveChanges:=False  ' release real source

    ' --- Year shift in years ---
    Dim yearShift As Long
    yearShift = CLng(demoYear) - CLng(realYear)

    Dim r As Long
    Dim agentNum As String, tellerNum As String, custNum As String, polNum As String
    Dim factor As Double

    For r = 2 To lastRow
        ' --- Agent ---
        agentNum = Trim$(CStr(arr(r, D_RAW_AGENTNUM)))
        If agentNum <> "" Then
            If Not dictAgents.Exists(agentNum) Then
                dictAgents(agentNum) = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1503) & " " & (dictAgents.Count + 1)
                dictAgentFactor(agentNum) = 0.85 + ((HashStr(agentNum) Mod 46) / 100)
            End If
            arr(r, D_RAW_AGENTNAME) = dictAgents(agentNum)
        End If
        factor = 1
        If dictAgentFactor.Exists(agentNum) Then factor = dictAgentFactor(agentNum)

        ' --- Teller ---
        tellerNum = Trim$(CStr(arr(r, D_RAW_TELLERNUM)))
        If tellerNum <> "" Then
            If Not dictTellers.Exists(tellerNum) Then
                dictTellers(tellerNum) = ChrW(1496) & ChrW(1500) & ChrW(1512) & " " & (dictTellers.Count + 1)
            End If
            arr(r, D_RAW_TELLERNAME) = dictTellers(tellerNum)
        End If

        ' --- Customer ---
        custNum = Trim$(CStr(arr(r, D_RAW_CUSTOMER)))
        If custNum <> "" Then
            If Not dictCustomers.Exists(custNum) Then
                dictCustomers(custNum) = ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & " " & (dictCustomers.Count + 1)
            End If
            arr(r, D_RAW_CUSTNAME) = dictCustomers(custNum)
        End If
        ' Wipe ID number entirely
        arr(r, D_RAW_IDNUMBER) = ""

        ' --- Policy ---
        polNum = Trim$(CStr(arr(r, D_RAW_POLICY)))
        If polNum <> "" Then
            If Not dictPolicies.Exists(polNum) Then
                dictPolicies(polNum) = "POL-" & Format$(dictPolicies.Count + 1, "000000")
            End If
            arr(r, D_RAW_POLICY) = dictPolicies(polNum)
        End If

        ' --- Premium / Commission scrambling (per-agent factor) ---
        If Not IsBlankV(arr(r, D_RAW_PREMIUM)) Then
            If IsNumeric(arr(r, D_RAW_PREMIUM)) Then
                arr(r, D_RAW_PREMIUM) = CDbl(arr(r, D_RAW_PREMIUM)) * factor
            End If
        End If
        If Not IsBlankV(arr(r, D_RAW_COMMISSION)) Then
            If IsNumeric(arr(r, D_RAW_COMMISSION)) Then
                arr(r, D_RAW_COMMISSION) = CDbl(arr(r, D_RAW_COMMISSION)) * factor
            End If
        End If

        ' --- Date shifts ---
        ShiftDate arr, r, D_RAW_BORDEREU, yearShift
        ShiftDate arr, r, D_RAW_INSURANCE_START, yearShift
    Next r

    ' --- Write to brand-new workbook ---
    Dim wbOut As Workbook
    Set wbOut = Workbooks.Add
    Dim wsOut As Worksheet
    Set wsOut = wbOut.Worksheets(1)
    wsOut.Name = D_DATA_SHEET_NAME

    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(lastRow, lastCol)).Value = arr

    Dim outPath As String
    outPath = demoFolder & demoYear & ".xlsx"
    Application.DisplayAlerts = False
    wbOut.SaveAs outPath, xlOpenXMLWorkbook
    wbOut.Close
    Application.DisplayAlerts = True

    AnonymizeOne = lastRow - 1
    Exit Function

ERR_HANDLER:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Err.Raise Err.Number, "AnonymizeOne(" & realYear & ")", Err.Description
End Function


' ============================================================================
' Helpers
' ============================================================================

' Polynomial hash (deterministic, stable across runs)
Private Function HashStr(ByVal s As String) As Long
    Dim h As Long, i As Long, ch As Long
    h = 0
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        h = ((h * 31) + ch) Mod 1000003
        If h < 0 Then h = h + 1000003
    Next i
    HashStr = h
End Function

Private Sub ShiftDate(ByRef arr As Variant, ByVal r As Long, ByVal col As Long, ByVal yearShift As Long)
    Dim v As Variant
    v = arr(r, col)
    If IsDate(v) Then
        arr(r, col) = DateAdd("yyyy", yearShift, CDate(v))
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 1 Then arr(r, col) = DateAdd("yyyy", yearShift, CDate(CDbl(v)))
    End If
End Sub

Private Function IsBlankV(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsBlankV = True
    ElseIf IsNull(v) Then
        IsBlankV = True
    ElseIf VarType(v) = vbString Then
        IsBlankV = (Trim$(CStr(v)) = "")
    Else
        IsBlankV = False
    End If
End Function

Private Function ReadParam(ByVal paramName As String) As String
    On Error GoTo FAIL
    Dim ws As Worksheet
    ' "hagdarot" sheet name (Hebrew)
    Set ws = ThisWorkbook.Worksheets(ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514))
    If ws Is Nothing Then GoTo FAIL

    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, D_PARAM_NAME_COL).End(xlUp).Row
    For r = 1 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, D_PARAM_NAME_COL).Value2))) = UCase$(paramName) Then
            ReadParam = Trim$(CStr(ws.Cells(r, D_PARAM_VALUE_COL).Value2))
            Exit Function
        End If
    Next r
FAIL:
    ReadParam = ""
End Function

Private Function EnsureTrailingSlash(ByVal p As String) As String
    If Right$(p, 1) <> "\" Then p = p & "\"
    EnsureTrailingSlash = p
End Function
