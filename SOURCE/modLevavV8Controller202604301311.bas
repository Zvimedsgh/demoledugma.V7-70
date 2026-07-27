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
    FillCompaniesData

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

Public Sub TestCompaniesMinimal()

    Dim ws As Worksheet
    Dim baseYear As String
    Dim currentYear As String
    
    Dim colCompany As Long
    Dim colPremium As Long
    Dim colCommission As Long
    
    Dim i As Long
    
    ' שנים מדף הבית
    baseYear = CStr(ThisWorkbook.Names("rngBaseYear").RefersToRange.Value)
    currentYear = CStr(ThisWorkbook.Names("rngCurrentYear").RefersToRange.Value)
    
    Set ws = ThisWorkbook.Worksheets(currentYear)
    
    ' עמודות מהגדרות
    colCompany = GetFieldColumn("COMPANY_NAME")
    colPremium = GetFieldColumn("PREMIUM")
    colCommission = GetFieldColumn("COMPANY_COMMISSION")
    
    MsgBox "Company Col=" & colCompany & vbCrLf & _
           "Premium Col=" & colPremium & vbCrLf & _
           "Commission Col=" & colCommission
    
    ' הדפסת 10 שורות ראשונות
    For i = 2 To 11
        
        Debug.Print _
            "Row " & i & " | " & _
            ws.Cells(i, colCompany).Value & " | " & _
            ws.Cells(i, colPremium).Value & " | " & _
            ws.Cells(i, colCommission).Value
            
    Next i
    
    MsgBox "Check Immediate Window (Ctrl+G)", vbInformation

End Sub

Public Sub TestColumns()

    Dim colCompany As Long
    Dim colPremium As Long
    Dim colCommission As Long

    colCompany = GetFieldColumn("COMPANY_NAME")
    colPremium = GetFieldColumn("PREMIUM")
    colCommission = GetFieldColumn("COMPANY_COMMISSION")

    MsgBox "Company=" & colCompany & vbCrLf & _
           "Premium=" & colPremium & vbCrLf & _
           "Commission=" & colCommission

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
    Dim wsBase As Worksheet
    Dim wsCur As Worksheet
    Dim baseYear As String
    Dim currentYear As String

    Dim dictBasePrem As Object
    Dim dictCurPrem As Object
    Dim dictBaseComm As Object
    Dim dictCurComm As Object

    Dim keys As Object
    Dim k As Variant
    Dim outRow As Long

    Dim colCompany As Long
    Dim colPremium As Long
    Dim colCommission As Long

    baseYear = CStr(ThisWorkbook.Names("rngBaseYear").RefersToRange.Value)
    currentYear = CStr(ThisWorkbook.Names("rngCurrentYear").RefersToRange.Value)

    Set wsBase = ThisWorkbook.Worksheets(baseYear)
    Set wsCur = ThisWorkbook.Worksheets(currentYear)
    Set wsOut = ThisWorkbook.Worksheets(SheetCompaniesName())

    colCompany = GetFieldColumn("COMPANY_NAME")
    colPremium = GetFieldColumn("PREMIUM")
    colCommission = GetFieldColumn("COMPANY_COMMISSION")

    Set dictBasePrem = CreateObject("Scripting.Dictionary")
    Set dictCurPrem = CreateObject("Scripting.Dictionary")
    Set dictBaseComm = CreateObject("Scripting.Dictionary")
    Set dictCurComm = CreateObject("Scripting.Dictionary")
    Set keys = CreateObject("Scripting.Dictionary")

    AggregateCompanyMoney wsBase, colCompany, colPremium, colCommission, dictBasePrem, dictBaseComm, keys
    AggregateCompanyMoney wsCur, colCompany, colPremium, colCommission, dictCurPrem, dictCurComm, keys

    outRow = 3

    For Each k In keys.keys
        wsOut.Cells(outRow, 1).Value = k

        wsOut.Cells(outRow, 2).Value = GetDictNumber(dictBasePrem, k)
        wsOut.Cells(outRow, 3).Value = GetDictNumber(dictCurPrem, k)
        wsOut.Cells(outRow, 4).Value = PercentChange(GetDictNumber(dictBasePrem, k), GetDictNumber(dictCurPrem, k))

        wsOut.Cells(outRow, 8).Value = GetDictNumber(dictBaseComm, k)
        wsOut.Cells(outRow, 9).Value = GetDictNumber(dictCurComm, k)
        wsOut.Cells(outRow, 10).Value = PercentChange(GetDictNumber(dictBaseComm, k), GetDictNumber(dictCurComm, k))

        outRow = outRow + 1
    Next k

End Sub

Private Sub AggregateCompanyMoney( _
    ByVal wsData As Worksheet, _
    ByVal colCompany As Long, _
    ByVal colPremium As Long, _
    ByVal colCommission As Long, _
    ByVal dictPrem As Object, _
    ByVal dictComm As Object, _
    ByVal keys As Object)

    Dim lastRow As Long
    Dim r As Long
    Dim companyName As String

    lastRow = wsData.Cells(wsData.Rows.Count, colCompany).End(xlUp).Row

    For r = 2 To lastRow
        companyName = Trim$(CStr(wsData.Cells(r, colCompany).Value))

        If companyName <> "" Then
            If Not keys.Exists(companyName) Then keys.Add companyName, True
            If Not dictPrem.Exists(companyName) Then dictPrem.Add companyName, 0#
            If Not dictComm.Exists(companyName) Then dictComm.Add companyName, 0#

            dictPrem(companyName) = dictPrem(companyName) + ToNumber(wsData.Cells(r, colPremium).Value)
            dictComm(companyName) = dictComm(companyName) + ToNumber(wsData.Cells(r, colCommission).Value)
        End If
    Next r

End Sub

Private Function GetDictNumber(ByVal dict As Object, ByVal key As Variant) As Double
    If dict.Exists(CStr(key)) Then
        GetDictNumber = CDbl(dict(CStr(key)))
    Else
        GetDictNumber = 0#
    End If
End Function

Private Function PercentChange(ByVal baseValue As Double, ByVal currentValue As Double) As Variant
    If baseValue = 0 Then
        PercentChange = vbNullString
    Else
        PercentChange = (currentValue - baseValue) / baseValue
    End If
End Function

Private Function ToNumber(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then
        ToNumber = 0#
    ElseIf Trim$(CStr(v)) = "" Then
        ToNumber = 0#
    ElseIf IsNumeric(v) Then
        ToNumber = CDbl(v)
    Else
        ToNumber = 0#
    End If
End Function
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

            colLetter = Trim$(CStr(ws.Cells(r, "F").Value))
            colLetter = Replace(colLetter, ":", "")
            colLetter = Replace(colLetter, "$", "")

            If colLetter = "" Then
                Err.Raise vbObjectError + 2101, , "Empty column letter for field: " & fieldName
            End If

            GetFieldColumn = ColumnLetterToNumber(colLetter)
            Exit Function

        End If

    Next r

    Err.Raise vbObjectError + 2100, , "Field not found in settings H: " & fieldName

End Function

Private Function ColumnLetterToNumber(ByVal colLetter As String) As Long

    Dim i As Long
    Dim result As Long
    Dim ch As String

    colLetter = UCase$(Trim$(colLetter))
    colLetter = Replace(colLetter, ":", "")
    colLetter = Replace(colLetter, "$", "")

    For i = 1 To Len(colLetter)
        ch = Mid$(colLetter, i, 1)
        If ch >= "A" And ch <= "Z" Then
            result = result * 26 + (Asc(ch) - Asc("A") + 1)
        End If
    Next i

    If result = 0 Then
        Err.Raise vbObjectError + 2110, , "Invalid column letter: " & colLetter
    End If

    ColumnLetterToNumber = result

End Function
Private Function SettingsSheetName() As String
    SettingsSheetName = ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514)
End Function
