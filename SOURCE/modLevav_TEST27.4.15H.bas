Attribute VB_Name = "modLevav_TEST"
Option Explicit

' ============================================================
' LevavChat v7 - FINAL TESTABLE CORE
' One clean module. No Hebrew sheet names in code.
'
' Required named ranges:
'   tblParams          : parameters table, first col = key, second col = value
'   tblColumnMapping   : column mapping table, contains keys and column letters
'   tblBranchMapping   : branch mapping table, col1 = branch name, col2 = main branch
'   tblSystemMessages  : messages table, col1 = display message, col2 = key
'
' Required params:
'   DATA_PATH          : folder path ending with backslash, e.g. C:\LevavData\
'
' Required named cells:
'   rngBaseYear
'   rngCurrentYear
'
' Normalized structure:
'   D = ID_NUMBER
'   G = POLICY
'   I = COMPANY_NAME
'   K = BRANCH_NAME
'   N = AGENT_NAME
'   R = ACTION
'   S = PREMIUM
'   T = COMPANY_COMMISSION
' ============================================================

Private gParams As Object
Private gMessages As Object
Private gColumnMap As Object
Private gMainBranchMap As Object


Public Sub RunLevav_FINAL()

    On Error GoTo EH

    Dim baseYear As String
    Dim currentYear As String
    Dim dataPath As String

    Dim wsBase As Worksheet
    Dim wsCurrentFixed As Worksheet

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = "LevavChat: starting..."

    T_LoadAllLookups

    baseYear = T_GetNamedValue("rngBaseYear")
    currentYear = T_GetNamedValue("rngCurrentYear")
    dataPath = T_GetParam("DATA_PATH")

    If Len(baseYear) = 0 Then Err.Raise vbObjectError + 1001, , "rngBaseYear is empty"
    If Len(currentYear) = 0 Then Err.Raise vbObjectError + 1002, , "rngCurrentYear is empty"
    If Len(dataPath) = 0 Then Err.Raise vbObjectError + 1003, , "DATA_PATH is missing in tblParams"

    If Right$(dataPath, 1) <> "\" Then dataPath = dataPath & "\"

    Application.StatusBar = "LevavChat: normalizing current year..."
    T_NormalizeRawYear currentYear, dataPath

    Set wsBase = T_GetBaseNormalizedSheet(baseYear)
    Set wsCurrentFixed = ThisWorkbook.Worksheets(currentYear & "_fixed")

    Application.StatusBar = "LevavChat: building comparisons..."
    T_BuildComparisons wsBase, wsCurrentFixed, baseYear, currentYear

    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "RunLevav_FINAL completed successfully", vbInformation
    Exit Sub

EH:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "RunLevav_FINAL failed:" & vbCrLf & _
           "Error " & Err.Number & vbCrLf & _
           Err.Description, vbCritical

End Sub


Private Sub T_LoadAllLookups()

    T_LoadParams
    T_LoadMessages
    T_LoadColumnMap
    T_LoadMainBranchMap

End Sub


Private Sub T_LoadParams()

    Dim rng As Range
    Dim r As Range
    Dim keyText As String
    Dim valueText As String

    Set gParams = CreateObject("Scripting.Dictionary")
    Set rng = ThisWorkbook.Names("tblParams").RefersToRange

    For Each r In rng.Rows
        keyText = Trim(CStr(r.Cells(1, 1).Value))
        valueText = Trim(CStr(r.Cells(1, 2).Value))

        If Len(keyText) > 0 Then
            gParams(keyText) = valueText
        End If
    Next r

End Sub


Private Sub T_LoadMessages()

    Dim rng As Range
    Dim r As Range
    Dim msgText As String
    Dim keyText As String

    Set gMessages = CreateObject("Scripting.Dictionary")
    Set rng = ThisWorkbook.Names("tblSystemMessages").RefersToRange

    For Each r In rng.Rows
        msgText = Trim(CStr(r.Cells(1, 1).Value))
        keyText = Trim(CStr(r.Cells(1, 2).Value))

        If Len(keyText) > 0 Then
            gMessages(keyText) = msgText
        End If
    Next r

End Sub


Private Sub T_LoadColumnMap()

    Dim rng As Range
    Dim rowRange As Range
    Dim cell As Range
    Dim keyText As String
    Dim candidate As String
    Dim colNumber As Long

    Set gColumnMap = CreateObject("Scripting.Dictionary")
    Set rng = ThisWorkbook.Names("tblColumnMapping").RefersToRange

    For Each rowRange In rng.Rows

        keyText = ""

        For Each cell In rowRange.Cells
            candidate = Trim(CStr(cell.Value))

            If T_IsKnownFieldKey(candidate) Then
                keyText = UCase$(candidate)
                Exit For
            End If
        Next cell

        If Len(keyText) > 0 Then
            colNumber = 0

            For Each cell In rowRange.Cells
                candidate = Trim(CStr(cell.Value))

                If Len(candidate) > 0 Then
                    If IsNumeric(candidate) Then
                        If CLng(candidate) > 0 Then
                            colNumber = CLng(candidate)
                            Exit For
                        End If
                    ElseIf T_IsColumnLetter(candidate) Then
                        colNumber = T_ColumnLetterToNumber(candidate)
                        Exit For
                    End If
                End If
            Next cell

            If colNumber > 0 Then
                gColumnMap(keyText) = colNumber
            End If
        End If

    Next rowRange

End Sub


Private Sub T_LoadMainBranchMap()

    Dim rng As Range
    Dim r As Range
    Dim branchName As String
    Dim mainBranch As String

    Set gMainBranchMap = CreateObject("Scripting.Dictionary")
    Set rng = ThisWorkbook.Names("tblBranchMapping").RefersToRange

    For Each r In rng.Rows
        branchName = Trim(CStr(r.Cells(1, 1).Value))
        mainBranch = Trim(CStr(r.Cells(1, 2).Value))

        If Len(branchName) > 0 Then
            gMainBranchMap(branchName) = mainBranch
        End If
    Next r

End Sub


Private Function T_GetParam(ByVal keyText As String) As String

    If Not gParams Is Nothing Then
        If gParams.Exists(keyText) Then
            T_GetParam = Trim(CStr(gParams(keyText)))
            Exit Function
        End If
    End If

    T_GetParam = ""

End Function


Private Function T_GetMessage(ByVal keyText As String) As String

    If Not gMessages Is Nothing Then
        If gMessages.Exists(keyText) Then
            T_GetMessage = Trim(CStr(gMessages(keyText)))
            Exit Function
        End If
    End If

    T_GetMessage = keyText

End Function


Private Function T_GetNamedValue(ByVal nameText As String) As String

    T_GetNamedValue = Trim(CStr(ThisWorkbook.Names(nameText).RefersToRange.Value))

End Function


Private Function T_GetMappedColumn(ByVal fieldKey As String) As Long

    fieldKey = UCase$(Trim(fieldKey))

    If Not gColumnMap Is Nothing Then
        If gColumnMap.Exists(fieldKey) Then
            T_GetMappedColumn = CLng(gColumnMap(fieldKey))
            Exit Function
        End If
    End If

    Err.Raise vbObjectError + 2001, , "Mapped column not found: " & fieldKey

End Function


Private Function T_GetMappedColumnOptional(ByVal primaryKey As String, ByVal fallbackKey As String) As Long

    primaryKey = UCase$(Trim(primaryKey))
    fallbackKey = UCase$(Trim(fallbackKey))

    If Not gColumnMap Is Nothing Then
        If gColumnMap.Exists(primaryKey) Then
            T_GetMappedColumnOptional = CLng(gColumnMap(primaryKey))
            Exit Function
        End If

        If gColumnMap.Exists(fallbackKey) Then
            T_GetMappedColumnOptional = CLng(gColumnMap(fallbackKey))
            Exit Function
        End If
    End If

    Err.Raise vbObjectError + 2002, , "Mapped column not found: " & primaryKey & " / " & fallbackKey

End Function


Private Sub T_NormalizeRawYear(ByVal yearText As String, ByVal dataPath As String)

    Dim wbRaw As Workbook
    Dim wsRaw As Worksheet
    Dim wsFixed As Worksheet

    Dim filePath As String
    Dim openedExternal As Boolean

    Dim lastRow As Long
    Dim r As Long

    Dim colID As Long
    Dim colPolicy As Long
    Dim colCompany As Long
    Dim colBranch As Long
    Dim colAgent As Long
    Dim colAction As Long
    Dim colPremium As Long
    Dim colCommission As Long

    filePath = dataPath & yearText & ".xlsx"

    If Dir(filePath) <> "" Then
        Set wbRaw = Workbooks.Open(filePath, ReadOnly:=True)
        Set wsRaw = wbRaw.Worksheets(1)
        openedExternal = True
    ElseIf T_SheetExists(yearText) Then
        Set wsRaw = ThisWorkbook.Worksheets(yearText)
        openedExternal = False
    Else
        Err.Raise vbObjectError + 3001, , "Raw year source not found. Expected file or sheet: " & yearText
    End If

    Set wsFixed = T_RecreateSheet(yearText & "_fixed")

    wsFixed.Cells(1, "D").Value = "ID_NUMBER"
    wsFixed.Cells(1, "G").Value = "POLICY"
    wsFixed.Cells(1, "I").Value = "COMPANY_NAME"
    wsFixed.Cells(1, "K").Value = "BRANCH_NAME"
    wsFixed.Cells(1, "N").Value = "AGENT_NAME"
    wsFixed.Cells(1, "R").Value = "ACTION"
    wsFixed.Cells(1, "S").Value = "PREMIUM"
    wsFixed.Cells(1, "T").Value = "COMPANY_COMMISSION"

    colID = T_GetMappedColumn("ID_NUMBER")
    colPolicy = T_GetMappedColumn("POLICY")
    colCompany = T_GetMappedColumn("COMPANY_NAME")
    colBranch = T_GetMappedColumnOptional("BRANCH_NAME", "BRANCH")
    colAgent = T_GetMappedColumnOptional("AGENT_NAME", "AGENT")
    colAction = T_GetMappedColumn("ACTION")
    colPremium = T_GetMappedColumn("PREMIUM")
    colCommission = T_GetMappedColumn("COMPANY_COMMISSION")

    lastRow = wsRaw.Cells(wsRaw.Rows.Count, colCompany).End(xlUp).Row

    For r = 2 To lastRow

        wsFixed.Cells(r, "D").Value = wsRaw.Cells(r, colID).Value
        wsFixed.Cells(r, "G").Value = wsRaw.Cells(r, colPolicy).Value
        wsFixed.Cells(r, "I").Value = wsRaw.Cells(r, colCompany).Value
        wsFixed.Cells(r, "K").Value = wsRaw.Cells(r, colBranch).Value
        wsFixed.Cells(r, "N").Value = wsRaw.Cells(r, colAgent).Value
        wsFixed.Cells(r, "R").Value = wsRaw.Cells(r, colAction).Value
        wsFixed.Cells(r, "S").Value = wsRaw.Cells(r, colPremium).Value
        wsFixed.Cells(r, "T").Value = wsRaw.Cells(r, colCommission).Value

        If r Mod 5000 = 0 Then
            Application.StatusBar = "LevavChat: normalizing " & yearText & " row " & r & " of " & lastRow
        End If

    Next r

    wsFixed.DisplayRightToLeft = True
    wsFixed.Columns.AutoFit

    If openedExternal Then
        wbRaw.Close SaveChanges:=False
    End If

End Sub


Private Function T_GetBaseNormalizedSheet(ByVal baseYear As String) As Worksheet

    If T_SheetExists(baseYear) Then
        Set T_GetBaseNormalizedSheet = ThisWorkbook.Worksheets(baseYear)
        Exit Function
    End If

    If T_SheetExists(baseYear & "_fixed") Then
        Set T_GetBaseNormalizedSheet = ThisWorkbook.Worksheets(baseYear & "_fixed")
        Exit Function
    End If

    Err.Raise vbObjectError + 4001, , "Base normalized sheet not found: " & baseYear

End Function


Private Sub T_BuildComparisons( _
    ByVal wsBase As Worksheet, _
    ByVal wsCurrent As Worksheet, _
    ByVal baseYear As String, _
    ByVal currentYear As String)

    Dim companyDict As Object
    Dim branchDict As Object
    Dim mainBranchDict As Object

    Set companyDict = CreateObject("Scripting.Dictionary")
    Set branchDict = CreateObject("Scripting.Dictionary")
    Set mainBranchDict = CreateObject("Scripting.Dictionary")

    T_ProcessGroupSheet wsBase, companyDict, 9, 19, 20, 4, 7, 18, 0, False
    T_ProcessGroupSheet wsCurrent, companyDict, 9, 19, 20, 4, 7, 18, 1, False
    T_WriteComparisonOutput companyDict, baseYear, currentYear, "companies_comparison", T_H_Name()

    T_ProcessGroupSheet wsBase, branchDict, 11, 19, 20, 4, 7, 18, 0, False
    T_ProcessGroupSheet wsCurrent, branchDict, 11, 19, 20, 4, 7, 18, 1, False
    T_WriteComparisonOutput branchDict, baseYear, currentYear, "branches_comparison", T_H_Branch()

    T_ProcessGroupSheet wsBase, mainBranchDict, 11, 19, 20, 4, 7, 18, 0, True
    T_ProcessGroupSheet wsCurrent, mainBranchDict, 11, 19, 20, 4, 7, 18, 1, True
    T_WriteComparisonOutput mainBranchDict, baseYear, currentYear, "main_branch_comparison", T_H_MainBranch()

End Sub


Private Sub T_ProcessGroupSheet( _
    ByVal ws As Worksheet, _
    ByVal dict As Object, _
    ByVal colGroup As Long, _
    ByVal colPremium As Long, _
    ByVal colCommission As Long, _
    ByVal colIdNumber As Long, _
    ByVal colPolicy As Long, _
    ByVal colAction As Long, _
    ByVal yearSide As Long, _
    ByVal convertToMainBranch As Boolean)

    Dim lastRow As Long
    Dim r As Long

    Dim groupKey As String
    Dim idValue As String
    Dim policyValue As String
    Dim actionValue As String

    Dim arr As Variant
    Dim insuredDict As Object
    Dim policyState As Object

    Dim fullIdKey As String
    Dim fullPolicyKey As String

    Dim st As Variant
    Dim k As Variant
    Dim groupFromKey As String

    Set insuredDict = CreateObject("Scripting.Dictionary")
    Set policyState = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.Count, colGroup).End(xlUp).Row

    For r = 2 To lastRow

        groupKey = Trim(CStr(ws.Cells(r, colGroup).Value))

        If convertToMainBranch Then
            groupKey = T_GetMainBranchName(groupKey)
        End If

        If Len(groupKey) = 0 Then GoTo NextRow

        If Not dict.Exists(groupKey) Then
            dict.Add groupKey, Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
        End If

        arr = dict(groupKey)

        If yearSide = 0 Then
            arr(0) = arr(0) + T_ToNumber(ws.Cells(r, colPremium).Value)
            arr(2) = arr(2) + 1
            arr(8) = arr(8) + T_ToNumber(ws.Cells(r, colCommission).Value)
        Else
            arr(1) = arr(1) + T_ToNumber(ws.Cells(r, colPremium).Value)
            arr(3) = arr(3) + 1
            arr(9) = arr(9) + T_ToNumber(ws.Cells(r, colCommission).Value)
        End If

        idValue = Trim(CStr(ws.Cells(r, colIdNumber).Value))
        policyValue = Trim(CStr(ws.Cells(r, colPolicy).Value))
        actionValue = Trim(CStr(ws.Cells(r, colAction).Value))

        If Not T_IsCancelAction(actionValue) Then
            If Len(idValue) > 0 Then
                fullIdKey = groupKey & Chr(30) & idValue

                If Not insuredDict.Exists(fullIdKey) Then
                    insuredDict.Add fullIdKey, True

                    If yearSide = 0 Then
                        arr(4) = arr(4) + 1
                    Else
                        arr(5) = arr(5) + 1
                    End If
                End If
            End If
        End If

        If Len(policyValue) > 0 Then

            fullPolicyKey = groupKey & Chr(30) & policyValue

            If Not policyState.Exists(fullPolicyKey) Then
                policyState.Add fullPolicyKey, Array(False, False, False)
            End If

            st = policyState(fullPolicyKey)

            If T_IsPrimaryPolicyAction(actionValue) Then st(0) = True
            If T_IsAdditionAction(actionValue) Then st(1) = True
            If T_IsCancelAction(actionValue) Then st(2) = True

            policyState(fullPolicyKey) = st

        End If

        dict(groupKey) = arr

NextRow:
    Next r

    For Each k In policyState.keys

        st = policyState(k)
        groupFromKey = Split(CStr(k), Chr(30))(0)

        If dict.Exists(groupFromKey) Then

            arr = dict(groupFromKey)

            If st(2) = False Then
                If st(0) = True Or st(1) = True Then
                    If yearSide = 0 Then
                        arr(6) = arr(6) + 1
                    Else
                        arr(7) = arr(7) + 1
                    End If
                End If
            End If

            dict(groupFromKey) = arr

        End If

    Next k

End Sub


Private Function T_GetMainBranchName(ByVal branchName As String) As String

    Dim cleanBranch As String
    Dim rng As Range
    Dim ws As Worksheet
    Dim newRow As Long
    Dim firstCol As Long

    cleanBranch = Trim(CStr(branchName))

    If Len(cleanBranch) = 0 Then
        T_GetMainBranchName = T_GetMessage("MISSING_MAIN_BRANCH")
        Exit Function
    End If

    If gMainBranchMap.Exists(cleanBranch) Then
        If Len(Trim(CStr(gMainBranchMap(cleanBranch)))) > 0 Then
            T_GetMainBranchName = Trim(CStr(gMainBranchMap(cleanBranch)))
        Else
            T_GetMainBranchName = T_GetMessage("MISSING_MAIN_BRANCH")
        End If
        Exit Function
    End If

    Set rng = ThisWorkbook.Names("tblBranchMapping").RefersToRange
    Set ws = rng.Worksheet
    firstCol = rng.Columns(1).Column

    newRow = ws.Cells(ws.Rows.Count, firstCol).End(xlUp).Row + 1

    ws.Cells(newRow, firstCol).Value = cleanBranch
    ws.Cells(newRow, firstCol + 1).Value = ""
    ws.Cells(newRow, firstCol + 2).Value = "BRANCH_NAME"
    ws.Cells(newRow, firstCol + 3).Value = "MAIN_BRANCH"

    gMainBranchMap(cleanBranch) = ""

    T_GetMainBranchName = T_GetMessage("MISSING_MAIN_BRANCH")

End Function


Private Sub T_WriteComparisonOutput( _
    ByVal dict As Object, _
    ByVal baseYear As Variant, _
    ByVal currentYear As Variant, _
    ByVal outputSheetName As String, _
    ByVal firstHeader As String)

    Dim wsOut As Worksheet
    Dim outRow As Long
    Dim totalRow As Long
    Dim k As Variant
    Dim arr As Variant

    Set wsOut = T_RecreateSheet(outputSheetName)

    wsOut.DisplayRightToLeft = True

    wsOut.Range("A1:A2").Merge
    wsOut.Range("A1").Value = firstHeader

    T_WriteMetricHeader wsOut, 2, T_H_Production(), baseYear, currentYear
    T_WriteMetricHeader wsOut, 5, T_H_Documents(), baseYear, currentYear
    T_WriteMetricHeader wsOut, 8, T_H_Insureds(), baseYear, currentYear
    T_WriteMetricHeader wsOut, 11, T_H_Policies(), baseYear, currentYear
    T_WriteMetricHeader wsOut, 14, T_H_Commission(), baseYear, currentYear

    outRow = 3

    For Each k In dict.keys

        arr = dict(k)

        wsOut.Cells(outRow, 1).Value = k

        T_WriteMetricValues wsOut, outRow, 2, arr(0), arr(1)
        T_WriteMetricValues wsOut, outRow, 5, arr(2), arr(3)
        T_WriteMetricValues wsOut, outRow, 8, arr(4), arr(5)
        T_WriteMetricValues wsOut, outRow, 11, arr(6), arr(7)
        T_WriteMetricValues wsOut, outRow, 14, arr(8), arr(9)

        outRow = outRow + 1

    Next k

    totalRow = outRow
    wsOut.Cells(totalRow, 1).Value = T_H_Total()

    T_WriteMetricTotals wsOut, totalRow, 2
    T_WriteMetricTotals wsOut, totalRow, 5
    T_WriteMetricTotals wsOut, totalRow, 8
    T_WriteMetricTotals wsOut, totalRow, 11
    T_WriteMetricTotals wsOut, totalRow, 14

    T_FormatComparisonSheet wsOut, totalRow

End Sub


Private Sub T_WriteMetricHeader( _
    ByVal ws As Worksheet, _
    ByVal startCol As Long, _
    ByVal titleText As String, _
    ByVal baseYear As Variant, _
    ByVal currentYear As Variant)

    ws.Range(ws.Cells(1, startCol), ws.Cells(1, startCol + 2)).Merge
    ws.Cells(1, startCol).Value = titleText

    ws.Cells(2, startCol).Value = baseYear
    ws.Cells(2, startCol + 1).Value = currentYear
    ws.Cells(2, startCol + 2).Value = T_H_ChangePercent()

End Sub


Private Sub T_WriteMetricValues( _
    ByVal ws As Worksheet, _
    ByVal rowIndex As Long, _
    ByVal startCol As Long, _
    ByVal baseValue As Double, _
    ByVal currentValue As Double)

    ws.Cells(rowIndex, startCol).Value = baseValue
    ws.Cells(rowIndex, startCol + 1).Value = currentValue
    ws.Cells(rowIndex, startCol + 2).Formula = "=IF(" & _
        ws.Cells(rowIndex, startCol).Address(False, False) & "=0,""""," & _
        ws.Cells(rowIndex, startCol + 1).Address(False, False) & "/" & _
        ws.Cells(rowIndex, startCol).Address(False, False) & "-1)"

End Sub


Private Sub T_WriteMetricTotals( _
    ByVal ws As Worksheet, _
    ByVal totalRow As Long, _
    ByVal startCol As Long)

    If totalRow <= 3 Then Exit Sub

    ws.Cells(totalRow, startCol).Formula = "=SUM(" & ws.Cells(3, startCol).Address(False, False) & ":" & ws.Cells(totalRow - 1, startCol).Address(False, False) & ")"
    ws.Cells(totalRow, startCol + 1).Formula = "=SUM(" & ws.Cells(3, startCol + 1).Address(False, False) & ":" & ws.Cells(totalRow - 1, startCol + 1).Address(False, False) & ")"
    ws.Cells(totalRow, startCol + 2).Formula = "=IF(" & _
        ws.Cells(totalRow, startCol).Address(False, False) & "=0,""""," & _
        ws.Cells(totalRow, startCol + 1).Address(False, False) & "/" & _
        ws.Cells(totalRow, startCol).Address(False, False) & "-1)"

End Sub


Private Sub T_FormatComparisonSheet(ByVal ws As Worksheet, ByVal totalRow As Long)

    With ws.Range("A1:P2")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With

    If totalRow >= 3 Then
        ws.Range("A" & totalRow & ":P" & totalRow).Font.Bold = True
        ws.Range("A" & totalRow & ":P" & totalRow).Interior.Color = RGB(217, 225, 242)
    End If

    ws.Range("B:P").NumberFormat = "#,##0"
    ws.Range("D:D,G:G,J:J,M:M,P:P").NumberFormat = "0.0%"
    ws.Columns("A:P").AutoFit

End Sub


Private Function T_RecreateSheet(ByVal sheetName As String) As Worksheet

    Dim ws As Worksheet
    Dim safeName As String

    safeName = Left$(sheetName, 31)

    Application.DisplayAlerts = False

    On Error Resume Next
    ThisWorkbook.Worksheets(safeName).Delete
    On Error GoTo 0

    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = safeName

    Set T_RecreateSheet = ws

End Function


Private Function T_SheetExists(ByVal sheetName As String) As Boolean

    On Error Resume Next
    T_SheetExists = Not ThisWorkbook.Worksheets(sheetName) Is Nothing
    On Error GoTo 0

End Function


Private Function T_ToNumber(ByVal v As Variant) As Double

    Dim s As String

    If IsError(v) Then Exit Function

    If IsNumeric(v) Then
        T_ToNumber = CDbl(v)
        Exit Function
    End If

    s = Trim(CStr(v))
    If Len(s) = 0 Then Exit Function

    s = Replace(s, "?", "")
    s = Replace(s, ",", "")
    s = Replace(s, " ", "")
    s = Replace(s, Chr(160), "")

    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        s = "-" & Mid$(s, 2, Len(s) - 2)
    End If

    If IsNumeric(s) Then
        T_ToNumber = CDbl(s)
    Else
        T_ToNumber = Val(s)
    End If

End Function


Private Function T_IsKnownFieldKey(ByVal s As String) As Boolean

    s = UCase$(Trim(CStr(s)))

    Select Case s
        Case "COMPANY_NAME", "PREMIUM", "COMPANY_COMMISSION", _
             "ID_NUMBER", "POLICY", "ACTION", _
             "BRANCH", "BRANCH_NAME", _
             "AGENT", "AGENT_NAME"
            T_IsKnownFieldKey = True
        Case Else
            T_IsKnownFieldKey = False
    End Select

End Function


Private Function T_IsColumnLetter(ByVal s As String) As Boolean

    Dim i As Long
    Dim ch As String

    s = UCase$(Trim(s))

    If Len(s) < 1 Or Len(s) > 3 Then Exit Function

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "A" Or ch > "Z" Then Exit Function
    Next i

    T_IsColumnLetter = True

End Function


Private Function T_ColumnLetterToNumber(ByVal colLetter As String) As Long

    Dim i As Long
    Dim result As Long
    Dim ch As Integer

    colLetter = UCase$(Trim(colLetter))

    For i = 1 To Len(colLetter)
        ch = Asc(Mid$(colLetter, i, 1))
        If ch < 65 Or ch > 90 Then
            T_ColumnLetterToNumber = 0
            Exit Function
        End If

        result = result * 26 + (ch - 64)
    Next i

    T_ColumnLetterToNumber = result

End Function


Private Function T_IsPrimaryPolicyAction(ByVal actionValue As String) As Boolean

    T_IsPrimaryPolicyAction = _
        (StrComp(Trim(actionValue), T_H_New(), vbTextCompare) = 0 Or _
         StrComp(Trim(actionValue), T_H_Renewal(), vbTextCompare) = 0)

End Function


Private Function T_IsCancelAction(ByVal actionValue As String) As Boolean

    T_IsCancelAction = (StrComp(Trim(actionValue), T_H_Cancel(), vbTextCompare) = 0)

End Function


Private Function T_IsAdditionAction(ByVal actionValue As String) As Boolean

    T_IsAdditionAction = (StrComp(Trim(actionValue), T_H_Addition(), vbTextCompare) = 0)

End Function


Private Function T_H_New() As String
    T_H_New = ChrW(&H5D7) & ChrW(&H5D3) & ChrW(&H5E9)
End Function

Private Function T_H_Renewal() As String
    T_H_Renewal = ChrW(&H5D7) & ChrW(&H5D9) & ChrW(&H5D3) & ChrW(&H5D5) & ChrW(&H5E9)
End Function

Private Function T_H_Cancel() As String
    T_H_Cancel = ChrW(&H5D1) & ChrW(&H5D9) & ChrW(&H5D8) & ChrW(&H5D5) & ChrW(&H5DC)
End Function

Private Function T_H_Addition() As String
    T_H_Addition = ChrW(&H5EA) & ChrW(&H5D5) & ChrW(&H5E1) & ChrW(&H5E4) & ChrW(&H5EA)
End Function

Private Function T_H_Name() As String
    T_H_Name = ChrW(&H5E9) & ChrW(&H5DD)
End Function

Private Function T_H_Branch() As String
    T_H_Branch = ChrW(&H5E2) & ChrW(&H5E0) & ChrW(&H5E3)
End Function

Private Function T_H_MainBranch() As String
    T_H_MainBranch = ChrW(&H5E2) & ChrW(&H5E0) & ChrW(&H5E3) & " " & ChrW(&H5DE) & ChrW(&H5E8) & ChrW(&H5DB) & ChrW(&H5D6)
End Function

Private Function T_H_Production() As String
    T_H_Production = ChrW(&H5E4) & ChrW(&H5E8) & ChrW(&H5D5) & ChrW(&H5D3) & ChrW(&H5D5) & ChrW(&H5E7) & ChrW(&H5E6) & ChrW(&H5D9) & ChrW(&H5D4)
End Function

Private Function T_H_Documents() As String
    T_H_Documents = ChrW(&H5DE) & ChrW(&H5E1) & ChrW(&H5DE) & ChrW(&H5DB) & ChrW(&H5D9) & ChrW(&H5DD)
End Function

Private Function T_H_Insureds() As String
    T_H_Insureds = ChrW(&H5DE) & ChrW(&H5D1) & ChrW(&H5D5) & ChrW(&H5D8) & ChrW(&H5D7) & ChrW(&H5D9) & ChrW(&H5DD)
End Function

Private Function T_H_Policies() As String
    T_H_Policies = ChrW(&H5E4) & ChrW(&H5D5) & ChrW(&H5DC) & ChrW(&H5D9) & ChrW(&H5E1) & ChrW(&H5D5) & ChrW(&H5EA)
End Function

Private Function T_H_Commission() As String
    T_H_Commission = ChrW(&H5E2) & ChrW(&H5DE) & ChrW(&H5DC) & ChrW(&H5D4)
End Function

Private Function T_H_ChangePercent() As String
    T_H_ChangePercent = ChrW(&H5E9) & ChrW(&H5D9) & ChrW(&H5E0) & ChrW(&H5D5) & ChrW(&H5D9) & "%"
End Function

Private Function T_H_Total() As String
    T_H_Total = ChrW(&H5E1) & ChrW(&H5D4) & ChrW(&H22) & ChrW(&H5DB)
End Function

