Attribute VB_Name = "modLevavV8Controller"
Option Explicit

' =========================
' PUBLIC ENTRY
' =========================
Public Sub BuildReportsAndCharts()

    On Error GoTo ERR_HANDLER
    

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' יצירת כל הגיליונות מתוך TEMPLATE בלבד
    CreateReportFromTemplate SheetCompaniesName()
    CreateReportFromTemplate SheetMonthsName()
    CreateReportFromTemplate SheetAgentsName()
    CreateReportFromTemplate SheetTellersName()
    CreateReportFromTemplate SheetBranchName()
    CreateReportFromTemplate SheetMainBranchName()
    'FillCompaniesData

    ' הסתרת התבניות
    HideTemplateSheets

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Template reports created.", vbInformation
    
    Exit Sub

ERR_HANDLER:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "BuildReports failed: " & Err.Description, vbExclamation

End Sub

Public Sub DebugDataSheets()

    Dim ws As Worksheet
    Dim msg As String

    For Each ws In ThisWorkbook.Worksheets
        msg = msg & ws.Name & " | LastRow A=" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row & vbCrLf
    Next ws

    MsgBox msg, vbInformation

End Sub

Private Function GetCurrentYear() As String
    GetCurrentYear = "2025"
End Function

Private Sub FillCompaniesData()

    Dim wsOut As Worksheet
    Dim wsData As Worksheet
    Dim dictPremium As Object
    Dim dictCommission As Object

    Dim lastRow As Long
    Dim i As Long
    Dim outRow As Long
    Dim key As String

    Dim colCompany As Long
    Dim colPremium As Long
    Dim colCommission As Long

    Dim vPremium As Variant
    Dim vCommission As Variant

    Set wsOut = ThisWorkbook.Worksheets(SheetCompaniesName())
    Set wsData = ThisWorkbook.Worksheets("2025")

    colCompany = GetFieldColumn("COMPANY_NAME")
    colPremium = GetFieldColumn("PREMIUM")
    colCommission = GetFieldColumn("COMPANY_COMMISSION")

    Set dictPremium = CreateObject("Scripting.Dictionary")
    Set dictCommission = CreateObject("Scripting.Dictionary")

    lastRow = wsData.Cells(wsData.Rows.Count, colCompany).End(xlUp).Row

    For i = 2 To lastRow

        key = Trim$(CStr(wsData.Cells(i, colCompany).Value))

        If key <> "" Then

            If Not dictPremium.Exists(key) Then
                dictPremium.Add key, 0#
                dictCommission.Add key, 0#
            End If

            vPremium = wsData.Cells(i, colPremium).Value
            vCommission = wsData.Cells(i, colCommission).Value

            If Not IsError(vPremium) Then
                If IsNumeric(vPremium) Then
                    dictPremium(key) = dictPremium(key) + CDbl(vPremium)
                End If
            End If

            If Not IsError(vCommission) Then
                If IsNumeric(vCommission) Then
                    dictCommission(key) = dictCommission(key) + CDbl(vCommission)
                End If
            End If

        End If

    Next i

    outRow = 3

    Dim k As Variant
    For Each k In dictPremium.keys
        wsOut.Cells(outRow, 1).Value = k
        wsOut.Cells(outRow, 3).Value = dictPremium(k)
        wsOut.Cells(outRow, 7).Value = dictCommission(k)
        outRow = outRow + 1
    Next k

End Sub
' =========================
' TEMPLATE ENGINE
' =========================
Private Sub CreateReportFromTemplate(ByVal reportName As String)

    Dim wsTemplate As Worksheet
    Dim wsNew As Worksheet
    Dim templateName As String

    templateName = ReportTemplateName()

    If Not SheetExists(templateName) Then
        MsgBox "Missing template: " & templateName, vbCritical
        Exit Sub
    End If

    Set wsTemplate = ThisWorkbook.Worksheets(templateName)

    If wsTemplate.Visible <> xlSheetVisible Then
        wsTemplate.Visible = xlSheetVisible
    End If

    Application.DisplayAlerts = False
    If SheetExists(reportName) Then
        ThisWorkbook.Worksheets(reportName).Delete
    End If
    Application.DisplayAlerts = True

    wsTemplate.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)

    Set wsNew = ActiveSheet
    wsNew.Name = reportName
    wsNew.DisplayRightToLeft = True

    ClearReportDataRows wsNew

End Sub

Private Function ReportTemplateName() As String
    ReportTemplateName = ChrW(1491) & ChrW(1493) & ChrW(1495) & "_TEMPLATE"
End Function


' =========================
' CLEAN DATA
' =========================
Private Sub ClearReportDataRows(ByVal ws As Worksheet)

    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow >= 3 Then
        ws.Rows("3:" & ws.Rows.Count).ClearContents
    End If

End Sub


' =========================
' HIDE TEMPLATES
' =========================
Private Sub HideTemplateSheets()

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, ws.Name, "_TEMPLATE", vbTextCompare) > 0 Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

End Sub

Private Function CDblSafe(ByVal v As Variant) As Double

    If IsError(v) Then
        CDblSafe = 0
    ElseIf IsEmpty(v) Then
        CDblSafe = 0
    ElseIf Trim$(CStr(v)) = "" Then
        CDblSafe = 0
    ElseIf IsNumeric(v) Then
        CDblSafe = CDbl(v)
    Else
        CDblSafe = 0
    End If

End Function
' =========================
' UTIL
' =========================
Private Function SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    SheetExists = Not ws Is Nothing

End Function


' =========================
' SHEET NAMES (NO HEBREW IN CODE)
' =========================
Private Function SheetCompaniesName() As String
    SheetCompaniesName = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1514)
End Function

Private Function SheetMonthsName() As String
    SheetMonthsName = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
End Function

Private Function SheetAgentsName() As String
    SheetAgentsName = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1497) & ChrW(1501)
End Function

Private Function SheetTellersName() As String
    SheetTellersName = ChrW(1496) & ChrW(1500) & ChrW(1512) & ChrW(1497) & ChrW(1501)
End Function

Private Function SheetBranchName() As String
    SheetBranchName = ChrW(1506) & ChrW(1504) & ChrW(1508) & ChrW(1497) & ChrW(1501)
End Function

Private Function SheetMainBranchName() As String
    SheetMainBranchName = ChrW(1506) & ChrW(1504) & ChrW(1507) & " " & ChrW(1502) & ChrW(1512) & ChrW(1499) & ChrW(1494)
End Function
Private Function GetFieldColumn(ByVal fieldName As String) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim colLetter As String

    Set ws = ThisWorkbook.Worksheets(SettingsSheetName())
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row

    For r = 2 To lastRow

        If UCase$(Trim$(CStr(ws.Cells(r, "H").Value))) = UCase$(fieldName) Then

            colLetter = Trim$(CStr(ws.Cells(r, "E").Value))
            colLetter = Replace(colLetter, ":", "")
            colLetter = Replace(colLetter, "$", "")

            If colLetter = "" Then
                Err.Raise vbObjectError + 2101, , "Empty column letter for field: " & fieldName
            End If

            GetFieldColumn = ThisWorkbook.Worksheets("2025").Columns(colLetter).Column
            Exit Function

        End If

    Next r

    Err.Raise vbObjectError + 2100, , "Field not found in settings H: " & fieldName

End Function
Private Function SettingsSheetName() As String
    SettingsSheetName = ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514)
End Function
