Attribute VB_Name = "Module1"
'Attribute VB_Name = "modLevav_clean"
Option Explicit

Public Sub RunLevav()
    NormalizeCurrentYear
    BuildIssuesReview
End Sub

Private Function GetMappedColumnIndex(fieldName As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim colLetter As String

    Set ws = ThisWorkbook.Worksheets(SettingsSheetName())
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

    For r = 2 To lastRow

        If UCase$(Trim$(CStr(ws.Cells(r, "H").Value))) = UCase$(fieldName) Then

            colLetter = Trim$(CStr(ws.Cells(r, "F").Value))

            If colLetter = "" Then
                Err.Raise vbObjectError + 2101, , "Empty column for field: " & fieldName
            End If

            ' המרה מאות עמודה למספר
          GetMappedColumnIndex = Range(colLetter & "1").Column
            Exit Function

        End If

    Next r

    Err.Raise vbObjectError + 2000, , "Field not found: " & fieldName

End Function

Private Function GetMappedColumn(fieldName As String) As Long

    Dim ws As Worksheet
    Dim r As Long
    Dim colLetter As String

    ' גיליון ההגדרות לפי פונקציה קיימת אצלך
    Set ws = ThisWorkbook.Worksheets(SettingsSheetName())

    For r = 2 To ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

        If UCase$(Trim$(CStr(ws.Cells(r, "H").Value))) = UCase$(fieldName) Then

            colLetter = Trim$(CStr(ws.Cells(r, "F").Value))

            ' ניקוי מקרים כמו I:I או $I
            colLetter = Replace(colLetter, ":", "")
            colLetter = Replace(colLetter, "$", "")

            GetMappedColumn = ws.Columns(colLetter).Column
            Exit Function

        End If

    Next r

    Err.Raise vbObjectError + 1000, , "Field not found: " & fieldName

End Function

Public Sub BuildIssuesReview()

    Dim currentYear As String
    Dim wsFixed As Worksheet
    Dim wsIssues As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim outRow As Long
    Dim lastCol As Long

    currentYear = Trim(CStr(ThisWorkbook.Names("rngCurrentYear").RefersToRange.Value))

    Set wsFixed = ThisWorkbook.Worksheets(currentYear & "_fixed")
    Set wsIssues = RecreateSheet(IssuesSheetName(currentYear))

    lastCol = wsFixed.Cells(1, wsFixed.Columns.Count).End(xlToLeft).Column

    wsFixed.Range(wsFixed.Cells(1, 1), wsFixed.Cells(1, lastCol)).Copy wsIssues.Cells(1, 1)

    wsIssues.Cells(1, lastCol + 1).Value = "ISSUE_REASON"
    wsIssues.Cells(1, lastCol + 2).Value = "ISSUE_VALUE"
    wsIssues.Cells(1, lastCol + 3).Value = "REQUIRED_ACTION"
    wsIssues.Cells(1, lastCol + 4).Value = "REQUIRED_VALUE"

    outRow = 2
    lastRow = wsFixed.Cells(wsFixed.Rows.Count, "I").End(xlUp).Row

    For r = 2 To lastRow

        If Trim(CStr(wsFixed.Cells(r, "I").Value)) = "" Then
            AddIssueRow wsFixed, wsIssues, r, outRow, lastCol, "MISSING_COMPANY_NAME", wsFixed.Cells(r, "I").Value

        ElseIf Trim(CStr(wsFixed.Cells(r, "K").Value)) = "" Then
            AddIssueRow wsFixed, wsIssues, r, outRow, lastCol, "MISSING_BRANCH", wsFixed.Cells(r, "K").Value

        ElseIf Trim(CStr(wsFixed.Cells(r, "N").Value)) = "" Then
            AddIssueRow wsFixed, wsIssues, r, outRow, lastCol, "MISSING_AGENT", wsFixed.Cells(r, "N").Value

        ElseIf Trim(CStr(wsFixed.Cells(r, "S").Value)) = "" Then
            AddIssueRow wsFixed, wsIssues, r, outRow, lastCol, "MISSING_PREMIUM", wsFixed.Cells(r, "S").Value

        ElseIf Trim(CStr(wsFixed.Cells(r, "T").Value)) = "" Then
            AddIssueRow wsFixed, wsIssues, r, outRow, lastCol, "MISSING_COMMISSION", wsFixed.Cells(r, "T").Value

        ElseIf Val(wsFixed.Cells(r, "V").Value) = 0 Then
            AddIssueRow wsFixed, wsIssues, r, outRow, lastCol, "NON_ILS_CURRENCY", wsFixed.Cells(r, "U").Value
        End If

    Next r

    wsIssues.DisplayRightToLeft = True
    wsIssues.Columns.AutoFit

End Sub


Private Sub AddIssueRow(ByVal wsFixed As Worksheet, ByVal wsIssues As Worksheet, _
                        ByVal sourceRow As Long, ByRef outRow As Long, _
                        ByVal lastCol As Long, ByVal reason As String, ByVal issueValue As Variant)

    wsIssues.Range(wsIssues.Cells(outRow, 1), wsIssues.Cells(outRow, lastCol)).Value = _
        wsFixed.Range(wsFixed.Cells(sourceRow, 1), wsFixed.Cells(sourceRow, lastCol)).Value

    wsIssues.Cells(outRow, lastCol + 1).Value = reason
    wsIssues.Cells(outRow, lastCol + 2).Value = issueValue
    wsIssues.Cells(outRow, lastCol + 3).Value = ""
    wsIssues.Cells(outRow, lastCol + 4).Value = ""

    outRow = outRow + 1

End Sub

Private Function IssuesSheetName(ByVal yearText As String) As String
    IssuesSheetName = ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500) & "_" & yearText
End Function
Public Sub NormalizeCurrentYear()
    Dim rawYear As String
    Dim wsRaw As Worksheet, wsFixed As Worksheet
    Dim lastRow As Long, r As Long
    
    rawYear = Trim(CStr(ThisWorkbook.Names("rngCurrentYear").RefersToRange.Value)) ' 2025
    Set wsRaw = ThisWorkbook.Worksheets(rawYear)
    Set wsFixed = RecreateSheet(rawYear & "_fixed")
    
    ' fixed structure matching 2024
    wsFixed.Cells(1, "D").Value = "ID_NUMBER"
    wsFixed.Cells(1, "G").Value = "POLICY"
    wsFixed.Cells(1, "I").Value = "COMPANY_NAME"
    wsFixed.Cells(1, "K").Value = "BRANCH"
    wsFixed.Cells(1, "N").Value = "AGENT"
    wsFixed.Cells(1, "R").Value = "ACTION"
    wsFixed.Cells(1, "S").Value = "PREMIUM"
    wsFixed.Cells(1, "T").Value = "COMPANY_COMMISSION"
    
    
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, GetMappedColumnIndex("COMPANY_NAME")).End(xlUp).Row
    
    For r = 2 To lastRow
        wsFixed.Cells(r, "D").Value = wsRaw.Cells(r, GetMappedColumn("ID_NUMBER")).Value
        wsFixed.Cells(r, "G").Value = wsRaw.Cells(r, GetMappedColumn("POLICY")).Value
        wsFixed.Cells(r, "I").Value = wsRaw.Cells(r, GetMappedColumn("COMPANY_NAME")).Value
        wsFixed.Cells(r, "R").Value = wsRaw.Cells(r, GetMappedColumn("ACTION")).Value
        wsFixed.Cells(r, "K").Value = wsRaw.Cells(r, GetMappedColumnIndex("BRANCH_NAME")).Value
        wsFixed.Cells(r, "N").Value = wsRaw.Cells(r, GetMappedColumnIndex("AGENT_NAME")).Value
        Dim cur As String
Dim fx As Double
Dim prem As Double
Dim comm As Double

cur = CStr(wsRaw.Cells(r, GetMappedColumnIndex("CURRENCY")).Value)

prem = Val(wsRaw.Cells(r, GetMappedColumnIndex("PREMIUM")).Value)
comm = Val(wsRaw.Cells(r, GetMappedColumnIndex("COMPANY_COMMISSION")).Value)

fx = GetFxRate(cur)

If fx > 0 Then
    wsFixed.Cells(r, "S").Value = Round(prem * fx, 0)
wsFixed.Cells(r, "T").Value = Round(comm * fx, 0)
Else
    wsFixed.Cells(r, "S").Value = prem
    wsFixed.Cells(r, "T").Value = comm
End If
    Next r
    
    wsFixed.DisplayRightToLeft = True
    wsFixed.Columns.AutoFit
    wsFixed.Columns("S:T").NumberFormat = "#,##0"
End Sub



Private Sub ProcessCompanySheet(ws As Worksheet, dict As Object, _
    colCompany As Long, colPremium As Long, colCommission As Long, _
    colID As Long, colPolicy As Long, colAction As Long, _
    isCurrent As Integer, Optional isFixed As Integer = 0)

    Dim lastRow As Long, i As Long
    Dim key As String
    Dim arr

    lastRow = ws.Cells(ws.Rows.Count, colCompany).End(xlUp).Row

    For i = 2 To lastRow
        key = Trim(CStr(ws.Cells(i, colCompany).Value))

        If key <> "" Then
            If Not dict.exists(key) Then
                ' arr(0)=base premium, arr(1)=current premium
                ' arr(2)=base commission, arr(3)=current commission
                dict.Add key, Array(0#, 0#, 0#, 0#)
            End If

            arr = dict(key)

            If isCurrent = 0 Then
                arr(0) = arr(0) + CDbl(ws.Cells(i, colPremium).Value)
                arr(2) = arr(2) + CDbl(ws.Cells(i, colCommission).Value)
            Else
                arr(1) = arr(1) + CDbl(ws.Cells(i, colPremium).Value)
                arr(3) = arr(3) + CDbl(ws.Cells(i, colCommission).Value)
            End If

            dict(key) = arr
        End If
    Next i

End Sub


Public Sub BuildReview_hold()

    Dim baseYear As String, currentYear As String
    Dim wsBase As Worksheet, wsCurrent As Worksheet
    Dim dict As Object

    baseYear = Trim(CStr(ThisWorkbook.Names("rngBaseYear").RefersToRange.Value))
    currentYear = Trim(CStr(ThisWorkbook.Names("rngCurrentYear").RefersToRange.Value))

    Set wsBase = ThisWorkbook.Worksheets(baseYear)
    Set wsCurrent = ThisWorkbook.Worksheets(currentYear & "_fixed")

    Set dict = CreateObject("Scripting.Dictionary")

    ProcessCompanySheet wsBase, dict, 9, 19, 20, 4, 7, 18, 0
    ProcessCompanySheet wsCurrent, dict, 9, 19, 20, 4, 7, 18, 1

    WriteCompaniesOutput RecreateSheet("companies_comparison"), dict, baseYear, currentYear

End Sub


Private Sub WriteCompaniesOutput(ByVal wsOut As Worksheet, ByVal dict As Object, ByVal baseYear As String, ByVal currentYear As String)

    Dim r As Long
    Dim k As Variant
    Dim arr As Variant

    wsOut.Cells.Clear

    wsOut.Cells(1, 1).Value = "COMPANY_NAME"
    wsOut.Cells(1, 2).Value = "PREMIUM_" & baseYear
    wsOut.Cells(1, 3).Value = "PREMIUM_" & currentYear
    wsOut.Cells(1, 4).Value = "COMMISSION_" & baseYear
    wsOut.Cells(1, 5).Value = "COMMISSION_" & currentYear

    r = 2

    For Each k In dict.keys
        arr = dict(k)

        wsOut.Cells(r, 1).Value = k
        wsOut.Cells(r, 2).Value = arr(0)
        wsOut.Cells(r, 3).Value = arr(1)
        wsOut.Cells(r, 4).Value = arr(2)
        wsOut.Cells(r, 5).Value = arr(3)

        r = r + 1
    Next k

    wsOut.DisplayRightToLeft = True
    wsOut.Columns.AutoFit

End Sub


    
Private Function RecreateSheet(ByVal sheetName As String) As Worksheet

    MsgBox "Creating sheet: [" & sheetName & "]"

    If sheetName = "" Then
        Err.Raise 1001, , "Sheet name is EMPTY"
    End If

    Application.DisplayAlerts = False

    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Delete
    On Error GoTo 0

    Application.DisplayAlerts = True

    Set RecreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    RecreateSheet.Name = sheetName

End Function
Private Function SettingsSheetName() As String
    SettingsSheetName = ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514)
End Function


Private Function GetFxRate(ByVal currencyCode As Variant) As Double

    Select Case CLng(Val(currencyCode))

        Case 0, 90 ' ILS
            GetFxRate = 1#

        Case 1 ' USD
           GetFxRate = Val(ThisWorkbook.Names("rngDOLAR").RefersToRange.Value)
        Case Else
            GetFxRate = 0 ' ייכנס ללטיפול בהמשך

    End Select

End Function
