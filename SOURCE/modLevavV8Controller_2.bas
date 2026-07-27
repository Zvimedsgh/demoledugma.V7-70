Attribute VB_Name = "modLevavV8Controller"
Option Explicit

Private Const PARAM_NAME_COL As String = "J"
Private Const PARAM_VALUE_COL As String = "K"
Private Const FIELD_COL_COL As String = "F"
Private Const FIELD_ALIAS_COL As String = "H"
Private Const ACTION_COL As String = "N"
Private Const CORRECTION_COL As String = "O"
Private Const TEMPLATE_SUFFIX As String = "_TEMPLATE"
Private Const DATA_START_ROW As Long = 3
Private Const ISSUE_MISSING As String = "MISSING_VALUE"
Private Const ISSUE_HIGH_PREMIUM As String = "HIGH_PREMIUM"
Private Const ISSUE_BAD_NUMBER As String = "BAD_NUMBER"

Public Sub RunLevav_FINAL()
    Dim baseYear As String
    Dim currentYear As String
    Dim folderPath As String
    Dim basePath As String
    Dim currentPath As String
    Dim wsBase As Worksheet
    Dim wsCurrent As Worksheet
    Dim issueCount As Long

    On Error GoTo ErrHandler
    StartFastMode

    baseYear = GetBaseYear()
    currentYear = GetCurrentYear()
    folderPath = GetFilesFolder()

    basePath = FindSourceWorkbook(folderPath, baseYear)
    currentPath = FindSourceWorkbook(folderPath, currentYear)

    If basePath = vbNullString Then Err.Raise vbObjectError + 101, "RunLevav_FINAL", "Base year file was not found: " & baseYear
    If currentPath = vbNullString Then Err.Raise vbObjectError + 102, "RunLevav_FINAL", "Current year file was not found: " & currentYear

    Set wsBase = ImportSourceWorkbook(basePath, baseYear)
    Set wsCurrent = ImportSourceWorkbook(currentPath, currentYear)

    issueCount = BuildDataIssues(wsBase, wsCurrent)

    EndFastMode
    MsgBox "Data check completed. Issues found: " & issueCount, vbInformation
    Exit Sub

ErrHandler:
    EndFastMode
    MsgBox "Data check failed: " & Err.Description, vbExclamation
End Sub

Public Sub BuildReview()
    Dim wsReview As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim actionText As String
    Dim correctionText As String
    Dim fixedCount As Long

    On Error GoTo ErrHandler
    StartFastMode

    Set wsReview = GetReviewSheet()
    lastRow = wsReview.Cells(wsReview.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        actionText = UCase$(Trim$(CStr(wsReview.Cells(r, ACTION_COL).Value2)))
        correctionText = Trim$(CStr(wsReview.Cells(r, CORRECTION_COL).Value2))

        If actionText = "FIX" And correctionText <> vbNullString Then
            ApplyCorrection wsReview, r
            fixedCount = fixedCount + 1
        End If
    Next r

    EndFastMode
    MsgBox "Corrections applied: " & fixedCount, vbInformation
    Exit Sub

ErrHandler:
    EndFastMode
    MsgBox "Apply corrections failed: " & Err.Description, vbExclamation
End Sub

Public Sub BuildReportsAndCharts()
    Dim baseYear As String
    Dim currentYear As String
    Dim wsBase As Worksheet
    Dim wsCurrent As Worksheet
    Dim reportCount As Long

    On Error GoTo ErrHandler
    StartFastMode

    baseYear = GetBaseYear()
    currentYear = GetCurrentYear()

    Set wsBase = ThisWorkbook.Worksheets(baseYear)
    Set wsCurrent = ThisWorkbook.Worksheets(currentYear)

    reportCount = 0
    reportCount = reportCount + BuildOneReport(wsBase, wsCurrent, SheetCompaniesName(), "COMPANY_NAME", baseYear, currentYear)
    reportCount = reportCount + BuildOneReport(wsBase, wsCurrent, SheetMonthsName(), "__MONTH__", baseYear, currentYear)
    reportCount = reportCount + BuildOneReport(wsBase, wsCurrent, SheetAgentsName(), "AGENT", baseYear, currentYear)
    reportCount = reportCount + BuildOneReport(wsBase, wsCurrent, SheetTellersName(), "TELLER", baseYear, currentYear)

    If FieldColumnExists("BRANCH") Then
        reportCount = reportCount + BuildOneReport(wsBase, wsCurrent, SheetBranchName(), "BRANCH", baseYear, currentYear)
    End If

    If FieldColumnExists("MAIN_BRANCH") Then
        reportCount = reportCount + BuildOneReport(wsBase, wsCurrent, SheetMainBranchName(), "MAIN_BRANCH", baseYear, currentYear)
    End If

    SetAllSheetsRightToLeftSafe

    EndFastMode
    MsgBox "Excel reports built. Report sheets: " & reportCount, vbInformation
    Exit Sub

ErrHandler:
    EndFastMode
    MsgBox "Excel reports failed: " & Err.Description, vbExclamation
End Sub

Public Sub BuildLevavPresentationFull()
    On Error GoTo ErrHandler
    StartFastMode

    BuildPresentationFromReportSheets

    EndFastMode
    MsgBox "Presentation created.", vbInformation
    Exit Sub

ErrHandler:
    EndFastMode
    MsgBox "Presentation failed: " & Err.Description, vbExclamation
End Sub

Private Sub StartFastMode()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub EndFastMode()
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Function GetFilesFolder() As String
    Dim folderPath As String
    folderPath = Trim$(GetParamValue("FILES_FOLDER"))
    If folderPath = vbNullString Then Err.Raise vbObjectError + 201, "GetFilesFolder", "FILES_FOLDER is empty"
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    If Dir(folderPath, vbDirectory) = vbNullString Then Err.Raise vbObjectError + 202, "GetFilesFolder", "Folder not found: " & folderPath
    GetFilesFolder = folderPath
End Function

Private Function GetReportsFolder() As String
    Dim folderPath As String
    folderPath = Trim$(GetParamValue("REPORTS_FOLDER"))
    If folderPath = vbNullString Then Err.Raise vbObjectError + 211, "GetReportsFolder", "REPORTS_FOLDER is empty"
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    EnsureFolderExists folderPath
    GetReportsFolder = folderPath
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim parts() As String
    Dim i As Long
    Dim currentPath As String

    If Dir(folderPath, vbDirectory) <> vbNullString Then Exit Sub
    parts = Split(folderPath, "\")
    If UBound(parts) < 0 Then Exit Sub
    currentPath = parts(0) & "\"

    For i = 1 To UBound(parts)
        If parts(i) <> vbNullString Then
            currentPath = currentPath & parts(i) & "\"
            If Dir(currentPath, vbDirectory) = vbNullString Then MkDir currentPath
        End If
    Next i
End Sub

Private Function GetParamValue(ByVal paramName As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SettingsSheetName())
    lastRow = ws.Cells(ws.Rows.Count, PARAM_NAME_COL).End(xlUp).Row

    For r = 1 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, PARAM_NAME_COL).Value2))) = UCase$(paramName) Then
            GetParamValue = Trim$(CStr(ws.Cells(r, PARAM_VALUE_COL).Value2))
            Exit Function
        End If
    Next r

    Err.Raise vbObjectError + 203, "GetParamValue", "Parameter not found: " & paramName
End Function

Private Function GetOptionalParamValue(ByVal paramName As String, ByVal defaultValue As String) As String
    On Error GoTo Fallback
    GetOptionalParamValue = GetParamValue(paramName)
    Exit Function
Fallback:
    GetOptionalParamValue = defaultValue
End Function

Private Function GetBaseYear() As String
    GetBaseYear = GetNamedOrCellValue("rngBaseYear", HomeSheetName(), "B2")
End Function

Private Function GetCurrentYear() As String
    GetCurrentYear = GetNamedOrCellValue("rngCurrentYear", HomeSheetName(), "B3")
End Function

Private Function GetNamedOrCellValue(ByVal rangeName As String, ByVal sheetName As String, ByVal fallbackAddress As String) As String
    On Error GoTo Fallback
    GetNamedOrCellValue = Trim$(CStr(ThisWorkbook.Names(rangeName).RefersToRange.Value2))
    If GetNamedOrCellValue <> vbNullString Then Exit Function
Fallback:
    GetNamedOrCellValue = Trim$(CStr(ThisWorkbook.Worksheets(sheetName).Range(fallbackAddress).Value2))
End Function

Private Function FindSourceWorkbook(ByVal folderPath As String, ByVal yearText As String) As String
    Dim p As String

    p = folderPath & yearText & ".xlsx"
    If Len(Dir(p)) > 0 Then FindSourceWorkbook = p: Exit Function

    p = folderPath & yearText & ".xlsm"
    If Len(Dir(p)) > 0 Then FindSourceWorkbook = p: Exit Function

    p = folderPath & yearText & ".xls"
    If Len(Dir(p)) > 0 Then FindSourceWorkbook = p: Exit Function

    p = folderPath & yearText & "_fixed.xlsx"
    If Len(Dir(p)) > 0 Then FindSourceWorkbook = p: Exit Function

    p = folderPath & yearText & "_fixed.xlsm"
    If Len(Dir(p)) > 0 Then FindSourceWorkbook = p: Exit Function

    FindSourceWorkbook = vbNullString
End Function

Private Function ImportSourceWorkbook(ByVal filePath As String, ByVal targetSheetName As String) As Worksheet
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim srcRange As Range

    DeleteSheetIfExists targetSheetName

    Set wbSrc = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=False)
    Set wsSrc = FirstVisibleWorksheet(wbSrc)
    Set srcRange = wsSrc.UsedRange

    Set wsDest = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsDest.Name = targetSheetName

    srcRange.Copy Destination:=wsDest.Range("A1")
    wsDest.Columns.AutoFit
    wsDest.DisplayRightToLeft = True

    wbSrc.Close SaveChanges:=False
    Set ImportSourceWorkbook = wsDest
End Function

Private Function FirstVisibleWorksheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            Set FirstVisibleWorksheet = ws
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 204, "FirstVisibleWorksheet", "No visible worksheet in source workbook"
End Function

Private Function BuildDataIssues(ByVal wsBase As Worksheet, ByVal wsCurrent As Worksheet) As Long
    Dim wsReview As Worksheet
    Dim outRow As Long
    Dim threshold As Double

    Set wsReview = CreateReviewSheetFromTemplate()
    threshold = CDbl(Val(GetOptionalParamValue("PREMIUM_THRESHOLD", "200000")))
    outRow = 2

    ScanOneSheetForIssues wsReview, outRow, wsBase, threshold
    ScanOneSheetForIssues wsReview, outRow, wsCurrent, threshold
    ApplyReviewValidation wsReview, outRow
    wsReview.Columns.AutoFit
    BuildDataIssues = outRow - 2
End Function

Private Function CreateReviewSheetFromTemplate() As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsReview As Worksheet
    Dim templateName As String

    templateName = ReviewTemplateSheetName()
    If Not SheetExists(templateName) Then Err.Raise vbObjectError + 221, "CreateReviewSheetFromTemplate", "Review template sheet was not found"

    Set wsTemplate = ThisWorkbook.Worksheets(templateName)
    DeleteSheetIfExists ReviewSheetName()
    wsTemplate.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set wsReview = ActiveSheet
    wsReview.Name = ReviewSheetName()
    wsReview.DisplayRightToLeft = True
    ClearReviewData wsReview
    Set CreateReviewSheetFromTemplate = wsReview
End Function

Private Sub ClearReviewData(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2
    ws.Rows("2:" & ws.Rows.Count).ClearContents
End Sub

Private Sub ApplyReviewValidation(ByVal ws As Worksheet, ByVal outRow As Long)
    Dim endRow As Long
    endRow = outRow - 1
    If endRow < 2 Then endRow = 2
    With ws.Range(ACTION_COL & "2:" & ACTION_COL & endRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="FIX,IGNORE,SEND_TO_REVIEW"
    End With
End Sub

Private Sub ScanOneSheetForIssues(ByVal wsReview As Worksheet, ByRef outRow As Long, ByVal wsData As Worksheet, ByVal threshold As Double)
    Dim requiredKeys As Variant
    Dim key As Variant
    Dim colNum As Long
    Dim lastRow As Long
    Dim r As Long
    Dim v As Variant
    Dim n As Double
    Dim premiumCol As Long

    requiredKeys = Array("COMPANY_NAME", "BRANCH", "AGENT", "PREMIUM", "COMPANY_COMMISSION", "ID_NUMBER", "POLICY", "ACTION")
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    For Each key In requiredKeys
        If FieldColumnExists(CStr(key)) Then
            colNum = GetFieldColumn(CStr(key))
            For r = 2 To lastRow
                If Trim$(CStr(wsData.Cells(r, colNum).Value2)) = vbNullString Then
                    AddIssue wsReview, outRow, wsData.Name, r, colNum, CStr(key), vbNullString, ISSUE_MISSING, wsData.Name & " row " & r
                End If
            Next r
        End If
    Next key

    If FieldColumnExists("PREMIUM") Then
        premiumCol = GetFieldColumn("PREMIUM")
        For r = 2 To lastRow
            v = wsData.Cells(r, premiumCol).Value2
            If Trim$(CStr(v)) <> vbNullString Then
                If TryToDouble(v, n) Then
                    If Abs(n) > threshold Then
                        AddIssue wsReview, outRow, wsData.Name, r, premiumCol, "PREMIUM", v, ISSUE_HIGH_PREMIUM, wsData.Name & " row " & r & " value " & CStr(v)
                    End If
                Else
                    AddIssue wsReview, outRow, wsData.Name, r, premiumCol, "PREMIUM", v, ISSUE_BAD_NUMBER, wsData.Name & " row " & r
                End If
            End If
        Next r
    End If
End Sub

Private Sub AddIssue(ByVal ws As Worksheet, ByRef outRow As Long, ByVal targetSheet As String, ByVal targetRow As Long, ByVal targetCol As Long, ByVal fieldKey As String, ByVal currentValue As Variant, ByVal issueCode As String, ByVal detailText As String)
    ws.Cells(outRow, "A").Value = targetSheet
    ws.Cells(outRow, "B").Value = targetSheet
    ws.Cells(outRow, "C").Value = targetRow
    ws.Cells(outRow, "D").Value = targetCol
    ws.Cells(outRow, "E").Value = fieldKey
    ws.Cells(outRow, "F").Value = currentValue
    ws.Cells(outRow, "G").Value = issueCode
    ws.Cells(outRow, "H").Value = detailText
    ws.Cells(outRow, ACTION_COL).Value = "SEND_TO_REVIEW"
    outRow = outRow + 1
End Sub

Private Sub ApplyCorrection(ByVal wsReview As Worksheet, ByVal reviewRow As Long)
    Dim targetSheet As String
    Dim targetRow As Long
    Dim targetCol As Long
    Dim correctionValue As Variant

    targetSheet = Trim$(CStr(wsReview.Cells(reviewRow, "B").Value2))
    targetRow = CLng(Val(wsReview.Cells(reviewRow, "C").Value2))
    targetCol = CLng(Val(wsReview.Cells(reviewRow, "D").Value2))
    correctionValue = wsReview.Cells(reviewRow, CORRECTION_COL).Value

    If targetSheet = vbNullString Then Err.Raise vbObjectError + 301, "ApplyCorrection", "Missing target sheet at review row " & reviewRow
    If targetRow < 1 Then Err.Raise vbObjectError + 302, "ApplyCorrection", "Missing target row at review row " & reviewRow
    If targetCol < 1 Then Err.Raise vbObjectError + 303, "ApplyCorrection", "Missing target column at review row " & reviewRow

    ThisWorkbook.Worksheets(targetSheet).Cells(targetRow, targetCol).Value = correctionValue
End Sub

Private Function BuildOneReport(ByVal wsBase As Worksheet, ByVal wsCurrent As Worksheet, ByVal reportSheetName As String, ByVal groupKey As String, ByVal baseYear As String, ByVal currentYear As String) As Long
    Dim dict As Object
    Dim wsOut As Worksheet

    If groupKey <> "__MONTH__" Then
        If Not FieldColumnExists(groupKey) Then Exit Function
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    AddSheetToReportDict dict, wsBase, groupKey, True
    AddSheetToReportDict dict, wsCurrent, groupKey, False

    Set wsOut = CreateReportSheetFromTemplate(reportSheetName)
    WriteReportData wsOut, dict, baseYear, currentYear
    BuildOneReport = 1
End Function

Private Function CreateReportSheetFromTemplate(ByVal reportSheetName As String) As Worksheet
    Dim templateName As String
    Dim wsTemplate As Worksheet
    Dim wsOut As Worksheet

    templateName = reportSheetName & TEMPLATE_SUFFIX
    DeleteSheetIfExists reportSheetName

    If SheetExists(templateName) Then
        Set wsTemplate = ThisWorkbook.Worksheets(templateName)
        wsTemplate.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Set wsOut = ActiveSheet
        wsOut.Name = reportSheetName
        wsOut.DisplayRightToLeft = True
        ClearReportData wsOut
    Else
        Set wsOut = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsOut.Name = reportSheetName
        wsOut.DisplayRightToLeft = True
        WriteFallbackHeaders wsOut
    End If

    Set CreateReportSheetFromTemplate = wsOut
End Function

Private Sub ClearReportData(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < DATA_START_ROW Then lastRow = DATA_START_ROW
    ws.Rows(DATA_START_ROW & ":" & ws.Rows.Count).ClearContents
End Sub

Private Sub WriteFallbackHeaders(ByVal ws As Worksheet)
    ws.Range("A1").Value = "Category"
    ws.Range("B1").Value = "Base Premium"
    ws.Range("C1").Value = "Current Premium"
    ws.Range("D1").Value = "Premium Change"
    ws.Range("E1").Value = "Base Commission"
    ws.Range("F1").Value = "Current Commission"
    ws.Range("G1").Value = "Commission Change"
    ws.Range("H1").Value = "Base Documents"
    ws.Range("I1").Value = "Current Documents"
    ws.Range("J1").Value = "Documents Change"
    ws.Range("K1").Value = "Base Insured"
    ws.Range("L1").Value = "Current Insured"
    ws.Range("M1").Value = "Insured Change"
    ws.Range("N1").Value = "Base Policies"
    ws.Range("O1").Value = "Current Policies"
    ws.Range("P1").Value = "Policies Change"
    ws.Rows(1).Font.Bold = True
End Sub

Private Sub WriteReportData(ByVal ws As Worksheet, ByVal dict As Object, ByVal baseYear As String, ByVal currentYear As String)
    Dim r As Long
    Dim k As Variant
    Dim arr As Variant
    Dim totalRow As Long

    r = DATA_START_ROW
    For Each k In dict.Keys
        arr = dict(k)
        ws.Cells(r, "A").Value = CStr(k)
        ws.Cells(r, "B").Value = arr(0)
        ws.Cells(r, "C").Value = arr(1)
        ws.Cells(r, "D").Value = SafeChange(arr(1), arr(0))
        ws.Cells(r, "E").Value = arr(2)
        ws.Cells(r, "F").Value = arr(3)
        ws.Cells(r, "G").Value = SafeChange(arr(3), arr(2))
        ws.Cells(r, "H").Value = arr(4)
        ws.Cells(r, "I").Value = arr(5)
        ws.Cells(r, "J").Value = SafeChange(arr(5), arr(4))
        ws.Cells(r, "K").Value = arr(6).Count
        ws.Cells(r, "L").Value = arr(7).Count
        ws.Cells(r, "M").Value = SafeChange(arr(7).Count, arr(6).Count)
        ws.Cells(r, "N").Value = arr(8).Count
        ws.Cells(r, "O").Value = arr(9).Count
        ws.Cells(r, "P").Value = SafeChange(arr(9).Count, arr(8).Count)
        r = r + 1
    Next k

    totalRow = r
    ws.Cells(totalRow, "A").Value = "TOTAL"
    If totalRow > DATA_START_ROW Then
        ws.Range("B" & totalRow & ":P" & totalRow).FormulaR1C1 = "=SUM(R" & DATA_START_ROW & "C:R[-1]C)"
    End If
    ws.Rows(totalRow).Font.Bold = True
    ws.Range("D:D,G:G,J:J,M:M,P:P").NumberFormat = "0.0%"
    ws.Range("B:C,E:F,H:I,K:L,N:O").NumberFormat = "#,##0"
    ws.Columns.AutoFit
End Sub

Private Sub AddSheetToReportDict(ByVal dict As Object, ByVal ws As Worksheet, ByVal groupKey As String, ByVal isBase As Boolean)
    Dim lastRow As Long
    Dim r As Long
    Dim groupValue As String
    Dim arr As Variant
    Dim prem As Double
    Dim comm As Double
    Dim idVal As String
    Dim polVal As String
    Dim polCount As Double

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        groupValue = GetGroupValue(ws, r, groupKey)
        If groupValue <> vbNullString Then
            If Not dict.Exists(groupValue) Then dict.Add groupValue, EmptyReportArray()
            arr = dict(groupValue)

            prem = GetNumericCellByAlias(ws, r, "PREMIUM")
            comm = GetNumericCellByAlias(ws, r, "COMPANY_COMMISSION")
            idVal = GetStringCellByAlias(ws, r, "ID_NUMBER")
            polVal = GetStringCellByAlias(ws, r, "POLICY")
            polCount = GetPolicyCountValue(ws, r)

            If isBase Then
                arr(0) = arr(0) + prem
                arr(2) = arr(2) + comm
                arr(4) = arr(4) + 1
                AddUnique arr(6), idVal
                arr(8) = arr(8) + polCount
            Else
                arr(1) = arr(1) + prem
                arr(3) = arr(3) + comm
                arr(5) = arr(5) + 1
                AddUnique arr(7), idVal
                arr(9) = arr(9) + polCount
            End If

            dict(groupValue) = arr
        End If
    Next r
End Sub

Private Function EmptyReportArray() As Variant
    Dim arr(0 To 9) As Variant
    Set arr(6) = CreateObject("Scripting.Dictionary")
    Set arr(7) = CreateObject("Scripting.Dictionary")
    arr(0) = 0: arr(1) = 0: arr(2) = 0: arr(3) = 0: arr(4) = 0
    arr(5) = 0: arr(8) = 0: arr(9) = 0
    EmptyReportArray = arr
End Function

Private Sub AddUnique(ByVal dict As Object, ByVal keyValue As String)
    If keyValue = vbNullString Then Exit Sub
    If Not dict.Exists(keyValue) Then dict.Add keyValue, True
End Sub

Private Function GetPolicyCountValue(ByVal ws As Worksheet, ByVal r As Long) As Double
    Dim actionText As String
    actionText = UCase$(Trim$(GetStringCellByAlias(ws, r, "ACTION")))
    If actionText = vbNullString Then
        GetPolicyCountValue = 0
    ElseIf InStr(1, actionText, "CANCEL", vbTextCompare) > 0 Then
        GetPolicyCountValue = -1
    ElseIf InStr(1, actionText, "NEW", vbTextCompare) > 0 Or InStr(1, actionText, "RENEW", vbTextCompare) > 0 Then
        GetPolicyCountValue = 1
    Else
        GetPolicyCountValue = 0
    End If
End Function

Private Function GetGroupValue(ByVal ws As Worksheet, ByVal r As Long, ByVal groupKey As String) As String
    Dim dt As Variant
    If groupKey = "__MONTH__" Then
        dt = GetBestDateValue(ws, r)
        If IsDate(dt) Then
            GetGroupValue = Format$(Month(CDate(dt)), "00")
        Else
            GetGroupValue = vbNullString
        End If
    Else
        GetGroupValue = Trim$(GetStringCellByAlias(ws, r, groupKey))
    End If
End Function

Private Function GetBestDateValue(ByVal ws As Worksheet, ByVal r As Long) As Variant
    If FieldColumnExists("BORDEREAU_DATE") Then
        GetBestDateValue = ws.Cells(r, GetFieldColumn("BORDEREAU_DATE")).Value
    ElseIf FieldColumnExists("INSURANCE_START") Then
        GetBestDateValue = ws.Cells(r, GetFieldColumn("INSURANCE_START")).Value
    ElseIf FieldColumnExists("INSURANCE_START_DATE") Then
        GetBestDateValue = ws.Cells(r, GetFieldColumn("INSURANCE_START_DATE")).Value
    Else
        GetBestDateValue = Empty
    End If
End Function

Private Function SafeChange(ByVal newVal As Double, ByVal oldVal As Double) As Double
    If oldVal = 0 Then
        If newVal = 0 Then SafeChange = 0 Else SafeChange = 1
    Else
        SafeChange = (newVal - oldVal) / oldVal
    End If
End Function

Private Function GetNumericCellByAlias(ByVal ws As Worksheet, ByVal r As Long, ByVal fieldKey As String) As Double
    Dim n As Double
    If Not FieldColumnExists(fieldKey) Then Exit Function
    If TryToDouble(ws.Cells(r, GetFieldColumn(fieldKey)).Value2, n) Then GetNumericCellByAlias = n
End Function

Private Function GetStringCellByAlias(ByVal ws As Worksheet, ByVal r As Long, ByVal fieldKey As String) As String
    If Not FieldColumnExists(fieldKey) Then Exit Function
    GetStringCellByAlias = Trim$(CStr(ws.Cells(r, GetFieldColumn(fieldKey)).Value2))
End Function

Private Function FieldColumnExists(ByVal fieldKey As String) As Boolean
    On Error GoTo Nope
    FieldColumnExists = (GetFieldColumn(fieldKey) > 0)
    Exit Function
Nope:
    FieldColumnExists = False
End Function

Private Function GetFieldColumn(ByVal fieldKey As String) As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim aliasText As String
    Dim colText As String

    Set ws = ThisWorkbook.Worksheets(SettingsSheetName())
    lastRow = ws.Cells(ws.Rows.Count, FIELD_ALIAS_COL).End(xlUp).Row

    For r = 1 To lastRow
        aliasText = UCase$(Trim$(CStr(ws.Cells(r, FIELD_ALIAS_COL).Value2)))
        If aliasText = UCase$(fieldKey) Then
            colText = Trim$(CStr(ws.Cells(r, FIELD_COL_COL).Value2))
            GetFieldColumn = ColumnLetterToNumber(colText)
            Exit Function
        End If
    Next r

    Err.Raise vbObjectError + 401, "GetFieldColumn", "Field alias not found in settings: " & fieldKey
End Function

Private Function ColumnLetterToNumber(ByVal colText As String) As Long
    Dim i As Long
    Dim ch As String
    Dim result As Long

    colText = UCase$(Trim$(colText))
    If IsNumeric(colText) Then ColumnLetterToNumber = CLng(colText): Exit Function

    For i = 1 To Len(colText)
        ch = Mid$(colText, i, 1)
        If ch < "A" Or ch > "Z" Then Err.Raise vbObjectError + 402, "ColumnLetterToNumber", "Bad column: " & colText
        result = result * 26 + (Asc(ch) - Asc("A") + 1)
    Next i

    ColumnLetterToNumber = result
End Function

Private Function TryToDouble(ByVal v As Variant, ByRef result As Double) As Boolean
    Dim s As String
    On Error GoTo Nope
    If IsNumeric(v) Then
        result = CDbl(v)
        TryToDouble = True
        Exit Function
    End If
    s = Trim$(CStr(v))
    s = Replace(s, ",", vbNullString)
    s = Replace(s, " ", vbNullString)
    If IsNumeric(s) Then
        result = CDbl(s)
        TryToDouble = True
    End If
    Exit Function
Nope:
    TryToDouble = False
End Function

Private Sub BuildPresentationFromReportSheets()
    Dim pptApp As Object
    Dim pptPres As Object
    Dim reportSheets As Variant
    Dim nm As Variant
    Dim slideCount As Long
    Dim outputPath As String

    reportSheets = Array(SheetCompaniesName(), SheetMonthsName(), SheetAgentsName(), SheetTellersName())
    Set pptApp = CreateObject("PowerPoint.Application")
    pptApp.Visible = True
    Set pptPres = pptApp.Presentations.Add

    For Each nm In reportSheets
        If SheetExists(CStr(nm)) Then
            If SheetHasReportData(ThisWorkbook.Worksheets(CStr(nm))) Then
                AddReportSlide pptPres, ThisWorkbook.Worksheets(CStr(nm)), CStr(nm)
                slideCount = slideCount + 1
            End If
        End If
    Next nm

    If slideCount = 0 Then Err.Raise vbObjectError + 501, "BuildPresentationFromReportSheets", "No report sheets with data were found"

    outputPath = GetReportsFolder() & "Levav_Presentation_" & Format(Now, "yyyymmdd_hhnn") & ".pptx"
    pptPres.SaveAs outputPath
End Sub

Private Function SheetHasReportData(ByVal ws As Worksheet) As Boolean
    SheetHasReportData = (ws.Cells(ws.Rows.Count, "A").End(xlUp).Row >= DATA_START_ROW)
End Function

Private Sub AddReportSlide(ByVal pptPres As Object, ByVal ws As Worksheet, ByVal titleText As String)
    Const ppLayoutBlank As Long = 12
    Const ppPasteEnhancedMetafile As Long = 2
    Const msoTextOrientationHorizontal As Long = 1
    Dim slideObj As Object
    Dim titleShape As Object
    Dim pastedShape As Object
    Dim rng As Range
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow > 35 Then lastRow = 35
    Set rng = ws.Range("A1:P" & lastRow)

    Set slideObj = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutBlank)
    Set titleShape = slideObj.Shapes.AddTextbox(msoTextOrientationHorizontal, 30, 18, 900, 36)
    titleShape.TextFrame.TextRange.Text = titleText
    titleShape.TextFrame.TextRange.Font.Size = 22
    titleShape.TextFrame.TextRange.Font.Bold = True
    titleShape.TextFrame.TextRange.ParagraphFormat.Alignment = 2

    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    Set pastedShape = slideObj.Shapes.PasteSpecial(ppPasteEnhancedMetafile)(1)
    With pastedShape
        .LockAspectRatio = True
        .Left = 30
        .Top = 70
        If .Width > 900 Then .Width = 900
        If .Height > 430 Then .Height = 430
        .Left = (960 - .Width) / 2
    End With
End Sub

Private Function GetReviewSheet() As Worksheet
    Set GetReviewSheet = ThisWorkbook.Worksheets(ReviewSheetName())
End Function

Private Sub DeleteSheetIfExists(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then ws.Delete
End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Sub SetAllSheetsRightToLeftSafe()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.DisplayRightToLeft = True
    Next ws
End Sub

Private Function HomeSheetName() As String
    HomeSheetName = ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)
End Function

Private Function SettingsSheetName() As String
    SettingsSheetName = ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514)
End Function

Private Function ReviewSheetName() As String
    ReviewSheetName = ChrW(1500) & ChrW(1496) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1500)
End Function

Private Function ReviewTemplateSheetName() As String
    ReviewTemplateSheetName = ReviewSheetName() & "_2025"
End Function

Private Function SheetCompaniesName() As String
    SheetCompaniesName = ChrW(1495) & ChrW(1489) & ChrW(1512) & ChrW(1493) & ChrW(1514)
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

Private Function SheetMonthsName() As String
    SheetMonthsName = ChrW(1495) & ChrW(1493) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
End Function
