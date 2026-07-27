Attribute VB_Name = "modLevav_TEST"
Option Explicit

Public Sub RunLevav_TEST()
    On Error GoTo EH

    Dim baseYear As String, currentYear As String
    Dim wsBase As Worksheet, wsCurrentFixed As Worksheet
    Dim dict As Object

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    baseYear = T_GetNamedValue("rngBaseYear")
    currentYear = T_GetNamedValue("rngCurrentYear")

    If Len(baseYear) = 0 Then Err.Raise vbObjectError + 1001, , "rngBaseYear is empty"
    If Len(currentYear) = 0 Then Err.Raise vbObjectError + 1002, , "rngCurrentYear is empty"

    ' Step 1: normalize raw current year, e.g. 2025 -> 2025_fixed
    T_NormalizeCurrentYear currentYear

    ' Step 2: compare normalized base year vs normalized current year
    Set wsBase = ThisWorkbook.Worksheets(baseYear)
    Set wsCurrentFixed = ThisWorkbook.Worksheets(currentYear & "_fixed")

    Set dict = CreateObject("Scripting.Dictionary")

    ' Fixed normalized structure:
    ' D = ID, G = Policy, I = Company, R = Action, S = Premium, T = Commission
    T_ProcessCompanySheet wsBase, dict, 9, 19, 20, 4, 7, 18, 0
    T_ProcessCompanySheet wsCurrentFixed, dict, 9, 19, 20, 4, 7, 18, 1

    T_WriteCompaniesOutput dict, baseYear, currentYear

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "RunLevav_TEST completed successfully", vbInformation
    Exit Sub

EH:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "RunLevav_TEST failed:" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub


Private Sub T_NormalizeCurrentYear(ByVal currentYear As String)
    Dim wsRaw As Worksheet, wsFixed As Worksheet
    Dim lastRow As Long, r As Long

    Set wsRaw = ThisWorkbook.Worksheets(currentYear)
    Set wsFixed = T_RecreateSheet(currentYear & "_fixed")

    wsFixed.Cells(1, "D").Value = "ID_NUMBER"
    wsFixed.Cells(1, "G").Value = "POLICY"
    wsFixed.Cells(1, "I").Value = "COMPANY_NAME"
    wsFixed.Cells(1, "R").Value = "ACTION"
    wsFixed.Cells(1, "S").Value = "PREMIUM"
    wsFixed.Cells(1, "T").Value = "COMPANY_COMMISSION"

    lastRow = wsRaw.Cells(wsRaw.Rows.Count, T_GetMappedColumnIndex("COMPANY_NAME")).End(xlUp).Row

    For r = 2 To lastRow
        wsFixed.Cells(r, "D").Value = wsRaw.Cells(r, T_GetMappedColumnIndex("ID_NUMBER")).Value
        wsFixed.Cells(r, "G").Value = wsRaw.Cells(r, T_GetMappedColumnIndex("POLICY")).Value
        wsFixed.Cells(r, "I").Value = wsRaw.Cells(r, T_GetMappedColumnIndex("COMPANY_NAME")).Value
        wsFixed.Cells(r, "R").Value = wsRaw.Cells(r, T_GetMappedColumnIndex("ACTION")).Value
        wsFixed.Cells(r, "S").Value = wsRaw.Cells(r, T_GetMappedColumnIndex("PREMIUM")).Value
        wsFixed.Cells(r, "T").Value = wsRaw.Cells(r, T_GetMappedColumnIndex("COMPANY_COMMISSION")).Value
    Next r

    wsFixed.DisplayRightToLeft = True
    wsFixed.Columns.AutoFit
End Sub


Private Sub T_ProcessCompanySheet( _
    ByVal ws As Worksheet, _
    ByVal dict As Object, _
    ByVal colCompany As Long, _
    ByVal colPremium As Long, _
    ByVal colCommission As Long, _
    ByVal colIdNumber As Long, _
    ByVal colPolicy As Long, _
    ByVal colAction As Long, _
    ByVal yearSide As Long)

    Dim lastRow As Long, r As Long
    Dim companyKey As String, idValue As String, policyValue As String, actionValue As String
    Dim arr As Variant
    Dim insuredDict As Object, policyState As Object
    Dim fullIdKey As String, fullPolicyKey As String
    Dim st As Variant, k As Variant, companyFromKey As String

    Set insuredDict = CreateObject("Scripting.Dictionary")
    Set policyState = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.Count, colCompany).End(xlUp).Row

    For r = 2 To lastRow

        companyKey = Trim(CStr(ws.Cells(r, colCompany).Value))
        If Len(companyKey) = 0 Then GoTo NextRow

        If Not dict.Exists(companyKey) Then
            dict.Add companyKey, Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
        End If

        arr = dict(companyKey)

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
                fullIdKey = companyKey & Chr(30) & idValue
                If Not insuredDict.Exists(fullIdKey) Then
                    insuredDict.Add fullIdKey, True
                    If yearSide = 0 Then arr(4) = arr(4) + 1 Else arr(5) = arr(5) + 1
                End If
            End If
        End If

        If Len(policyValue) > 0 Then
            fullPolicyKey = companyKey & Chr(30) & policyValue

            If Not policyState.Exists(fullPolicyKey) Then
                policyState.Add fullPolicyKey, Array(False, False, False)
            End If

            st = policyState(fullPolicyKey)

            If T_IsPrimaryPolicyAction(actionValue) Then st(0) = True
            If T_IsAdditionAction(actionValue) Then st(1) = True
            If T_IsCancelAction(actionValue) Then st(2) = True

            policyState(fullPolicyKey) = st
        End If

        dict(companyKey) = arr

NextRow:
    Next r

    For Each k In policyState.keys
        st = policyState(k)
        companyFromKey = Split(CStr(k), Chr(30))(0)

        If dict.Exists(companyFromKey) Then
            arr = dict(companyFromKey)

            If st(2) = False Then
                If st(0) = True Or st(1) = True Then
                    If yearSide = 0 Then arr(6) = arr(6) + 1 Else arr(7) = arr(7) + 1
                End If
            End If

            dict(companyFromKey) = arr
        End If
    Next k
End Sub


Private Sub T_WriteCompaniesOutput(ByVal dict As Object, ByVal baseYear As Variant, ByVal currentYear As Variant)
    Dim wsOut As Worksheet
    Dim outRow As Long, totalRow As Long
    Dim k As Variant, arr As Variant

    Set wsOut = T_RecreateSheet("companies_comparison")
    wsOut.DisplayRightToLeft = True

    wsOut.Range("A1:A2").Merge
    wsOut.Range("A1").Value = T_H_Name()

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

    T_FormatCompaniesSheet wsOut, totalRow
End Sub


Private Sub T_WriteMetricHeader(ByVal ws As Worksheet, ByVal startCol As Long, ByVal titleText As String, ByVal baseYear As Variant, ByVal currentYear As Variant)
    ws.Range(ws.Cells(1, startCol), ws.Cells(1, startCol + 2)).Merge
    ws.Cells(1, startCol).Value = titleText
    ws.Cells(2, startCol).Value = baseYear
    ws.Cells(2, startCol + 1).Value = currentYear
    ws.Cells(2, startCol + 2).Value = T_H_ChangePercent()
End Sub


Private Sub T_WriteMetricValues(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal startCol As Long, ByVal baseValue As Double, ByVal currentValue As Double)
    ws.Cells(rowIndex, startCol).Value = baseValue
    ws.Cells(rowIndex, startCol + 1).Value = currentValue
    ws.Cells(rowIndex, startCol + 2).Formula = "=IF(" & ws.Cells(rowIndex, startCol).Address(False, False) & "=0,""""," & ws.Cells(rowIndex, startCol + 1).Address(False, False) & "/" & ws.Cells(rowIndex, startCol).Address(False, False) & "-1)"
End Sub


Private Sub T_WriteMetricTotals(ByVal ws As Worksheet, ByVal totalRow As Long, ByVal startCol As Long)
    ws.Cells(totalRow, startCol).Formula = "=SUM(" & ws.Cells(3, startCol).Address(False, False) & ":" & ws.Cells(totalRow - 1, startCol).Address(False, False) & ")"
    ws.Cells(totalRow, startCol + 1).Formula = "=SUM(" & ws.Cells(3, startCol + 1).Address(False, False) & ":" & ws.Cells(totalRow - 1, startCol + 1).Address(False, False) & ")"
    ws.Cells(totalRow, startCol + 2).Formula = "=IF(" & ws.Cells(totalRow, startCol).Address(False, False) & "=0,""""," & ws.Cells(totalRow, startCol + 1).Address(False, False) & "/" & ws.Cells(totalRow, startCol).Address(False, False) & "-1)"
End Sub


Private Sub T_FormatCompaniesSheet(ByVal ws As Worksheet, ByVal totalRow As Long)
    With ws.Range("A1:P2")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With

    ws.Range("A" & totalRow & ":P" & totalRow).Font.Bold = True
    ws.Range("A" & totalRow & ":P" & totalRow).Interior.Color = RGB(217, 225, 242)

    ws.Range("B:P").NumberFormat = "#,##0"
    ws.Range("D:D,G:G,J:J,M:M,P:P").NumberFormat = "0.0%"
    ws.Columns("A:P").AutoFit
End Sub


Private Function T_GetNamedValue(ByVal nm As String) As String
    T_GetNamedValue = Trim(CStr(ThisWorkbook.Names(nm).RefersToRange.Value))
End Function


Private Function T_GetMappedColumnIndex(ByVal fieldKey As String) As Long
    Dim ws As Worksheet
    Dim r As Long, c As Long, cc As Long
    Dim v As String, candidate As String

    Set ws = T_GetSettingsSheet()
    fieldKey = UCase$(Trim$(fieldKey))

    For r = 1 To 300
        For c = 5 To 8
            v = UCase$(Trim$(CStr(ws.Cells(r, c).Value)))

            If v = fieldKey Then
                For cc = 5 To 8
                    candidate = Trim$(CStr(ws.Cells(r, cc).Value))

                    If Len(candidate) > 0 Then
                        If IsNumeric(candidate) Then
                            T_GetMappedColumnIndex = CLng(candidate)
                            Exit Function
                        ElseIf T_IsColumnLetter(candidate) Then
                            T_GetMappedColumnIndex = T_ColumnLetterToNumber(candidate)
                            Exit Function
                        End If
                    End If
                Next cc
            End If
        Next c
    Next r

    Err.Raise vbObjectError + 3001, , "Mapped field not found in settings E:H: " & fieldKey
End Function


Private Function T_GetSettingsSheet() As Worksheet
    Dim ws As Worksheet
    Dim r As Long, c As Long

    For Each ws In ThisWorkbook.Worksheets
        For r = 1 To 300
            For c = 5 To 8
                If UCase$(Trim$(CStr(ws.Cells(r, c).Value))) = "COMPANY_NAME" Then
                    Set T_GetSettingsSheet = ws
                    Exit Function
                End If
            Next c
        Next r
    Next ws

    Err.Raise vbObjectError + 3002, , "Settings sheet with COMPANY_NAME in E:H was not found"
End Function


Private Function T_IsColumnLetter(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    s = UCase$(Trim$(s))

    If Len(s) < 1 Or Len(s) > 3 Then Exit Function

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "A" Or ch > "Z" Then Exit Function
    Next i

    T_IsColumnLetter = True
End Function


Private Function T_ColumnLetterToNumber(ByVal colLetter As String) As Long
    Dim i As Long, result As Long, ch As Integer

    colLetter = UCase$(Trim$(colLetter))

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


Private Function T_RecreateSheet(ByVal sheetName As String) As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set T_RecreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    T_RecreateSheet.Name = sheetName
End Function


Private Function T_ToNumber(ByVal v As Variant) As Double
    Dim s As String

    If IsError(v) Then Exit Function
    If IsNumeric(v) Then
        T_ToNumber = CDbl(v)
        Exit Function
    End If

    s = Trim(CStr(v))
    s = Replace(s, "?", "")
    s = Replace(s, ",", "")
    s = Replace(s, " ", "")
    s = Replace(s, Chr(160), "")

    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        s = "-" & Mid$(s, 2, Len(s) - 2)
    End If

    If IsNumeric(s) Then T_ToNumber = CDbl(s)
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

