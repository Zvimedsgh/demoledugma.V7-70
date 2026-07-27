Attribute VB_Name = "modLevav_v7"
Option Explicit

' ============================================================
' LEVAVCHAT V7 - CLEAN RUNNER + REVIEW + COMPANIES COMPARISON
' Reads parameters from Named Ranges:
'   rngBaseYear
'   rngCurrentYear
'
' Reads column mapping from Settings sheet columns E:H:
'   COMPANY_NAME
'   PREMIUM
'   COMPANY_COMMISSION
'   ID_NUMBER
'   POLICY
'   ACTION
' ============================================================

Public Sub RunLevav()
    On Error GoTo EH
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Call BuildReview
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "RunLevav completed successfully", vbInformation
    Exit Sub

EH:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "RunLevav error:" & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical
End Sub


Public Sub BuildReview()
    

    Dim baseYear As String
    Dim currentYear As String
    
    Dim wsBase As Worksheet
    Dim wsCurrent As Worksheet
    
    Dim colCompany As Long
    Dim colPremium As Long
    Dim colCommission As Long
    Dim colIdNumber As Long
    Dim colPolicy As Long
    Dim colAction As Long
    
    ' Read years only from Home named ranges
    baseYear = GetNamedValue("rngBaseYear")
    currentYear = GetNamedValue("rngCurrentYear")
    
    If Len(baseYear) = 0 Then Err.Raise vbObjectError + 1001, , "rngBaseYear is empty"
    If Len(currentYear) = 0 Then Err.Raise vbObjectError + 1002, , "rngCurrentYear is empty"
    
    ' Find sheets by convention:
    ' Base year:     בסיס_2024 or 2024
    ' Current year:  נוכחית_2025 or 2025
    Set wsBase = FindYearSheet(baseYear, "בסיס")
    If wsBase Is Nothing Then
        Err.Raise vbObjectError + 1003, , "Base sheet not found. Tried: בסיס_" & baseYear & " or " & baseYear
    End If
    
    Set wsCurrent = FindYearSheet(currentYear, "נוכחית")
    If wsCurrent Is Nothing Then
        Err.Raise vbObjectError + 1004, , "Current sheet not found. Tried: נוכחית_" & currentYear & " or " & currentYear
    End If
    
    colCompany = GetMappedColumnIndex("COMPANY_NAME")
    colPremium = GetMappedColumnIndex("PREMIUM")
    colCommission = GetMappedColumnIndex("COMPANY_COMMISSION")
    colIdNumber = GetMappedColumnIndex("ID_NUMBER")
    colPolicy = GetMappedColumnIndex("POLICY")
    colAction = GetMappedColumnIndex("ACTION")
    
    Call BuildCompaniesComparisonSheet( _
        wsBase, wsCurrent, _
        colCompany, colPremium, colCommission, colIdNumber, colPolicy, colAction, _
        baseYear, currentYear)
    
    MsgBox "BuildReview completed successfully" & vbCrLf & _
           "Base sheet: " & wsBase.Name & vbCrLf & _
           "Current sheet: " & wsCurrent.Name, vbInformation
    End Sub


Private Function GetMappedColumnIndex(ByVal fieldKey As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim keyValue As String
    Dim colLetter As String

    For Each ws In ThisWorkbook.Worksheets

        lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

        For r = 1 To lastRow
            keyValue = Trim(CStr(ws.Cells(r, "H").Value))

            If StrComp(keyValue, fieldKey, vbTextCompare) = 0 Then
                colLetter = Trim(CStr(ws.Cells(r, "F").Value))
                GetMappedColumnIndex = ColumnLetterToNumber(colLetter)
                Exit Function
            End If
        Next r

    Next ws

    GetMappedColumnIndex = 0

End Function



Private Sub BuildCompaniesComparisonSheet( _
    ByVal wsBase As Worksheet, _
    ByVal wsCurrent As Worksheet, _
    ByVal colCompany As Long, _
    ByVal colPremium As Long, _
    ByVal colCommission As Long, _
    ByVal colIdNumber As Long, _
    ByVal colPolicy As Long, _
    ByVal colAction As Long, _
    ByVal baseYear As Variant, _
    ByVal currentYear As Variant)

    Dim wsOut As Worksheet
    Dim dict As Object
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    Call ProcessCompanySheet(wsBase, dict, colCompany, colPremium, colCommission, colIdNumber, colPolicy, colAction, 0)
    Call ProcessCompanySheet(wsCurrent, dict, colCompany, colPremium, colCommission, colIdNumber, colPolicy, colAction, 1)
    
    Set wsOut = RecreateSheet("companies_comparison")
    
    Call WriteCompaniesOutput(wsOut, dict, baseYear, currentYear)
End Sub


Public Sub ProcessCompanySheet( _
    ByVal ws As Worksheet, _
    ByVal dict As Object, _
    ByVal colCompany As Long, _
    ByVal colPremium As Long, _
    ByVal colCommission As Long, _
    ByVal colIdNumber As Long, _
    ByVal colPolicy As Long, _
    ByVal colAction As Long, _
    ByVal yearSide As Long)

    Dim lastRow As Long
    Dim r As Long
    
    Dim companyKey As String
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
    Dim companyFromKey As String
    
    Set insuredDict = CreateObject("Scripting.Dictionary")
    Set policyState = CreateObject("Scripting.Dictionary")
    
    lastRow = ws.Cells(ws.Rows.Count, colCompany).End(xlUp).Row
    
    For r = 2 To lastRow
        
        companyKey = Trim(CStr(ws.Cells(r, colCompany).Value))
        If Len(companyKey) = 0 Then GoTo NextRow
        
        If Not dict.Exists(companyKey) Then
            ' 0 Premium base
            ' 1 Premium current
            ' 2 Documents base
            ' 3 Documents current
            ' 4 Insured base
            ' 5 Insured current
            ' 6 Active policies base
            ' 7 Active policies current
            ' 8 Commission base
            ' 9 Commission current
            dict.Add companyKey, Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
        End If
        
        arr = dict(companyKey)
        
        If yearSide = 0 Then
            arr(0) = arr(0) + ToNumber(ws.Cells(r, colPremium).Value)
            arr(2) = arr(2) + 1
            arr(8) = arr(8) + ToNumber(ws.Cells(r, colCommission).Value)
        Else
            arr(1) = arr(1) + ToNumber(ws.Cells(r, colPremium).Value)
            arr(3) = arr(3) + 1
            arr(9) = arr(9) + ToNumber(ws.Cells(r, colCommission).Value)
        End If
        
        idValue = Trim(CStr(ws.Cells(r, colIdNumber).Value))
        policyValue = Trim(CStr(ws.Cells(r, colPolicy).Value))
        actionValue = Trim(CStr(ws.Cells(r, colAction).Value))
        
        If Not IsCancelAction(actionValue) Then
            If Len(idValue) > 0 Then
                fullIdKey = companyKey & Chr(30) & idValue
                
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
            fullPolicyKey = companyKey & Chr(30) & policyValue
            
            If Not policyState.Exists(fullPolicyKey) Then
                ' 0 = new / renewal
                ' 1 = addition
                ' 2 = cancellation
                policyState.Add fullPolicyKey, Array(False, False, False)
            End If
            
            st = policyState(fullPolicyKey)
            
            If IsPrimaryPolicyAction(actionValue) Then st(0) = True
            If IsAdditionAction(actionValue) Then st(1) = True
            If IsCancelAction(actionValue) Then st(2) = True
            
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
                    If yearSide = 0 Then
                        arr(6) = arr(6) + 1
                    Else
                        arr(7) = arr(7) + 1
                    End If
                End If
            End If
            
            dict(companyFromKey) = arr
        End If
    Next k
End Sub


Private Sub WriteCompaniesOutput( _
    ByVal ws As Worksheet, _
    ByVal dict As Object, _
    ByVal baseYear As Variant, _
    ByVal currentYear As Variant)

    Dim headers As Variant
    Dim r As Long
    Dim k As Variant
    Dim arr As Variant
    
    headers = Array( _
        "חברה", _
        "פרמיה " & CStr(baseYear), _
        "פרמיה " & CStr(currentYear), _
        "פרמיה שינוי %", _
        "מסמכים " & CStr(baseYear), _
        "מסמכים " & CStr(currentYear), _
        "מסמכים שינוי %", _
        "מבוטחים " & CStr(baseYear), _
        "מבוטחים " & CStr(currentYear), _
        "מבוטחים שינוי %", _
        "פוליסות פעילות " & CStr(baseYear), _
        "פוליסות פעילות " & CStr(currentYear), _
        "פוליסות שינוי %", _
        "עמלה " & CStr(baseYear), _
        "עמלה " & CStr(currentYear), _
        "עמלה שינוי %")
    
    ws.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    
    r = 2
    
    For Each k In dict.keys
        arr = dict(k)
        
        ws.Cells(r, 1).Value = CStr(k)
        
        ws.Cells(r, 2).Value = arr(0)
        ws.Cells(r, 3).Value = arr(1)
        ws.Cells(r, 4).Value = PercentChange(arr(0), arr(1))
        
        ws.Cells(r, 5).Value = arr(2)
        ws.Cells(r, 6).Value = arr(3)
        ws.Cells(r, 7).Value = PercentChange(arr(2), arr(3))
        
        ws.Cells(r, 8).Value = arr(4)
        ws.Cells(r, 9).Value = arr(5)
        ws.Cells(r, 10).Value = PercentChange(arr(4), arr(5))
        
        ws.Cells(r, 11).Value = arr(6)
        ws.Cells(r, 12).Value = arr(7)
        ws.Cells(r, 13).Value = PercentChange(arr(6), arr(7))
        
        ws.Cells(r, 14).Value = arr(8)
        ws.Cells(r, 15).Value = arr(9)
        ws.Cells(r, 16).Value = PercentChange(arr(8), arr(9))
        
        r = r + 1
    Next k
    
    If r > 2 Then
        ws.Cells(r, 1).Value = "סהכ"
        ws.Cells(r, 2).Formula = "=SUM(B2:B" & r - 1 & ")"
        ws.Cells(r, 3).Formula = "=SUM(C2:C" & r - 1 & ")"
        ws.Cells(r, 4).Formula = "=IF(B" & r & "=0,"""",C" & r & "/B" & r & ")"
        
        ws.Cells(r, 5).Formula = "=SUM(E2:E" & r - 1 & ")"
        ws.Cells(r, 6).Formula = "=SUM(F2:F" & r - 1 & ")"
        ws.Cells(r, 7).Formula = "=IF(E" & r & "=0,"""",F" & r & "/E" & r & ")"
        
        ws.Cells(r, 8).Formula = "=SUM(H2:H" & r - 1 & ")"
        ws.Cells(r, 9).Formula = "=SUM(I2:I" & r - 1 & ")"
        ws.Cells(r, 10).Formula = "=IF(H" & r & "=0,"""",I" & r & "/H" & r & ")"
        
        ws.Cells(r, 11).Formula = "=SUM(K2:K" & r - 1 & ")"
        ws.Cells(r, 12).Formula = "=SUM(L2:L" & r - 1 & ")"
        ws.Cells(r, 13).Formula = "=IF(K" & r & "=0,"""",L" & r & "/K" & r & ")"
        
        ws.Cells(r, 14).Formula = "=SUM(N2:N" & r - 1 & ")"
        ws.Cells(r, 15).Formula = "=SUM(O2:O" & r - 1 & ")"
        ws.Cells(r, 16).Formula = "=IF(N" & r & "=0,"""",O" & r & "/N" & r & ")"
    End If
    
    With ws
        .DisplayRightToLeft = True
        .Rows(1).Font.Bold = True
        .Columns("A:P").AutoFit
        .Range("B:C,N:O").NumberFormat = "#,##0"
        .Range("E:F,H:I,K:L").NumberFormat = "#,##0"
        .Range("D:D,G:G,J:J,M:M,P:P").NumberFormat = "0%"
        
        If r > 2 Then
            .Rows(r).Font.Bold = True
            .Rows(r).Interior.Color = RGB(255, 230, 153)
        End If
    End With
End Sub


Private Function GetNamedValue(ByVal nm As String) As String
    On Error GoTo EH
    GetNamedValue = Trim(CStr(ThisWorkbook.Names(nm).RefersToRange.Value))
    Exit Function

EH:
    Err.Raise vbObjectError + 2001, , "Named range not found or invalid: " & nm
End Function


Private Function SheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function


Private Function RecreateSheet(ByVal sheetName As String) As Worksheet
    Application.DisplayAlerts = False
    
    If SheetExists(sheetName) Then
        ThisWorkbook.Worksheets(sheetName).Delete
    End If
    
    Application.DisplayAlerts = True
    
    Set RecreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    RecreateSheet.Name = sheetName
End Function


Private Function GetSettingsSheet() As Worksheet
    Dim ws As Worksheet
    Dim r As Long
    Dim c As Long
    
    For Each ws In ThisWorkbook.Worksheets
        For r = 1 To 300
            For c = 5 To 8
                If UCase$(Trim$(CStr(ws.Cells(r, c).Value))) = "COMPANY_NAME" Then
                    Set GetSettingsSheet = ws
                    Exit Function
                End If
            Next c
        Next r
    Next ws
    
    Err.Raise vbObjectError + 3001, , "Settings sheet with COMPANY_NAME in columns E:H was not found"
End Function


'Private Function GetMappedColumnIndexClean(ByVal fieldKey As String) As Long    Dim ws As Worksheet
    Dim r As Long
    Dim c As Long
    Dim cc As Long
    Dim v As String
    Dim candidate As String
    
    Set ws = GetSettingsSheet()
    fieldKey = UCase$(Trim$(fieldKey))
    
    For r = 1 To 300
        For c = 5 To 8
            v = UCase$(Trim$(CStr(ws.Cells(r, c).Value)))
            
            If v = fieldKey Then
                For cc = 5 To 8
                    candidate = Trim$(CStr(ws.Cells(r, cc).Value))
                    
                    If Len(candidate) > 0 Then
                        If IsNumeric(candidate) Then
                            If CLng(candidate) > 0 Then
                                GetMappedColumnIndex = CLng(candidate)
                                Exit Function
                            End If
                        ElseIf IsColumnLetter(candidate) Then
                            GetMappedColumnIndex = ColumnLetterToNumber(candidate)
                            Exit Function
                        End If
                    End If
                Next cc
                
                Err.Raise vbObjectError + 3002, , "Field found but no column number/letter in E:H: " & fieldKey
            End If
        Next c
    Next r
    
    Err.Raise vbObjectError + 3003, , "Mapped field not found in settings E:H: " & fieldKey
End Function


Private Function IsColumnLetter(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    
    s = UCase$(Trim$(s))
    If Len(s) < 1 Or Len(s) > 3 Then Exit Function
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "A" Or ch > "Z" Then Exit Function
    Next i
    
    IsColumnLetter = True
End Function


Private Function ColumnLetterToNumber(ByVal colLetter As String) As Long
    Dim i As Long
    Dim result As Long
    Dim ch As Integer
    
    colLetter = UCase$(Trim$(colLetter))
    
    For i = 1 To Len(colLetter)
        ch = Asc(Mid$(colLetter, i, 1))
        If ch < 65 Or ch > 90 Then
            ColumnLetterToNumber = 0
            Exit Function
        End If
        
        result = result * 26 + (ch - 64)
    Next i
    
    ColumnLetterToNumber = result
End Function


Private Function ToNumber(ByVal v As Variant) As Double
    Dim s As String
    
    If IsError(v) Then
        ToNumber = 0
        Exit Function
    End If
    
    If IsNumeric(v) Then
        ToNumber = CDbl(v)
        Exit Function
    End If
    
    s = Trim(CStr(v))
    If Len(s) = 0 Then
        ToNumber = 0
        Exit Function
    End If
    
    s = Replace(s, "₪", "")
    s = Replace(s, ",", "")
    s = Replace(s, " ", "")
    s = Replace(s, Chr(160), "")
    
    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        s = "-" & Mid$(s, 2, Len(s) - 2)
    End If
    
    If IsNumeric(s) Then
        ToNumber = CDbl(s)
    Else
        ToNumber = Val(s)
    End If
End Function


Private Function PercentChange(ByVal oldValue As Double, ByVal newValue As Double) As Variant
    If oldValue = 0 Then
        PercentChange = vbNullString
    Else
        PercentChange = newValue / oldValue
    End If
End Function


Private Function NormalizeAction(ByVal actionText As String) As String
    NormalizeAction = Trim(LCase$(CStr(actionText)))
End Function


Private Function IsCancelAction(ByVal actionText As String) As Boolean
    Dim s As String
    s = NormalizeAction(actionText)
    
    IsCancelAction = _
        (InStr(1, s, "ביטול", vbTextCompare) > 0) Or _
        (InStr(1, s, "cancel", vbTextCompare) > 0)
End Function
Private Function GetSheetByYearOrName(ByVal valueFromParam As String, ByVal prefixText As String) As Worksheet
    Dim ws As Worksheet
    Dim target1 As String
    Dim target2 As String
    Dim target3 As String
    
    valueFromParam = Trim(CStr(valueFromParam))
    
    target1 = valueFromParam
    target2 = prefixText & "_" & valueFromParam
    target3 = prefixText & " " & valueFromParam
    
    For Each ws In ThisWorkbook.Worksheets
        If Trim(ws.Name) = target1 _
        Or Trim(ws.Name) = target2 _
        Or Trim(ws.Name) = target3 Then
            Set GetSheetByYearOrName = ws
            Exit Function
        End If
    Next ws
    
    Set GetSheetByYearOrName = Nothing
End Function

Private Function IsAdditionAction(ByVal actionText As String) As Boolean
    Dim s As String
    s = NormalizeAction(actionText)
    
    IsAdditionAction = _
        (InStr(1, s, "תוספת", vbTextCompare) > 0) Or _
        (InStr(1, s, "addition", vbTextCompare) > 0)
End Function


Private Function IsPrimaryPolicyAction(ByVal actionText As String) As Boolean
    Dim s As String
    s = NormalizeAction(actionText)
    
    IsPrimaryPolicyAction = _
        (InStr(1, s, "חדש", vbTextCompare) > 0) Or _
        (InStr(1, s, "חידוש", vbTextCompare) > 0) Or _
        (InStr(1, s, "new", vbTextCompare) > 0) Or _
        (InStr(1, s, "renew", vbTextCompare) > 0)
End Function
Private Function FindYearSheet(ByVal yearText As String, ByVal rolePrefix As String) As Worksheet
    Dim ws As Worksheet
    Dim name1 As String
    Dim name2 As String
    Dim name3 As String

    yearText = Trim(CStr(yearText))
    rolePrefix = Trim(CStr(rolePrefix))

    name1 = rolePrefix & "_" & yearText
    name2 = rolePrefix & " " & yearText
    name3 = yearText

    For Each ws In ThisWorkbook.Worksheets
        If Trim(ws.Name) = name1 _
        Or Trim(ws.Name) = name2 _
        Or Trim(ws.Name) = name3 Then
            
            Set FindYearSheet = ws
            Exit Function
        End If
    Next ws

    Set FindYearSheet = Nothing
End Function
