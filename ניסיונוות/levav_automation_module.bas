Option Explicit

Private Function GetSettingValue(ByVal keyName As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Set ws = ThisWorkbook.Worksheets("הגדרות_הפעלה")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, 1).Value)) = keyName Then
            GetSettingValue = Trim(CStr(ws.Cells(i, 2).Value))
            Exit Function
        End If
    Next i
    GetSettingValue = ""
End Function

Private Sub UpdateStatus(ByVal statusName As String, ByVal statusValue As String)
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Worksheets("הפעלה")
    For i = 17 To 20
        If Trim(CStr(ws.Cells(i, 8).Value)) = statusName Then
            ws.Cells(i, 9).Value = statusValue
            Exit For
        End If
    Next i
End Sub

Private Function PickFile(Optional ByVal titleText As String = "בחר קובץ") As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = titleText
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xlsm;*.xls"
        If .Show <> -1 Then Exit Function
        PickFile = .SelectedItems(1)
    End With
End Function

Private Sub CopyUsedRangeToSheet(ByVal sourcePath As String, ByVal sourceSheetIndex As Long, ByVal targetSheetName As String)
    Dim wbSrc As Workbook, wsSrc As Worksheet, wsTgt As Worksheet
    Set wbSrc = Workbooks.Open(sourcePath, ReadOnly:=True)
    Set wsSrc = wbSrc.Worksheets(sourceSheetIndex)
    Set wsTgt = ThisWorkbook.Worksheets(targetSheetName)
    wsTgt.Cells.Clear
    wsSrc.UsedRange.Copy
    wsTgt.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    wsTgt.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    wbSrc.Close SaveChanges:=False
End Sub

Public Sub טען_קובץ_מקור()
    Dim filePath As String
    filePath = PickFile("בחר קובץ מקור")
    If Len(filePath) = 0 Then Exit Sub
    CopyUsedRangeToSheet filePath, 1, GetSettingValue("גיליון מקור גלם")
    UpdateStatus "קובץ מקור", "נטען"
    MsgBox "קובץ המקור נטען לגיליון מקור_גלם", vbInformation
End Sub

Public Sub צור_קובץ_לטיפול()
    Dim folderPath As String, outPath As String
    folderPath = GetSettingValue("תיקיית לטיפול")
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath
    outPath = folderPath & "\" & "לטיפול_" & Format(Now, "yyyy-mm-dd_hhmm") & ".xlsx"
    ThisWorkbook.Worksheets(GetSettingValue("גיליון לטיפול")).Copy
    With ActiveWorkbook
        .SaveAs Filename:=outPath, FileFormat:=xlOpenXMLWorkbook
        .Close SaveChanges:=False
    End With
    UpdateStatus "קובץ לטיפול", "נוצר"
    MsgBox "נוצר קובץ לטיפול:" & vbCrLf & outPath, vbInformation
End Sub

Public Sub טען_קובץ_תיקונים()
    Dim filePath As String
    filePath = PickFile("בחר קובץ תיקונים")
    If Len(filePath) = 0 Then Exit Sub
    CopyUsedRangeToSheet filePath, 1, GetSettingValue("גיליון יומן יישום")
    UpdateStatus "קובץ תיקונים", "נטען"
    MsgBox "קובץ התיקונים נטען לגיליון יומן_יישום", vbInformation
End Sub

Public Sub הפק_דוח_סופי()
    Dim folderPath As String, outPath As String
    folderPath = GetSettingValue("תיקיית תוצאות")
    If Len(Dir(folderPath, vbDirectory)) = 0 Then MkDir folderPath

    Call סידור_רוחב_עמודות

    outPath = folderPath & "\" & "דוח_חודשי_" & Format(Now, "yyyy-mm-dd_hhmm") & ".xlsx"
    ThisWorkbook.SaveCopyAs outPath
    UpdateStatus "דוח סופי", "נשמר"
    MsgBox "נשמר דוח סופי:" & vbCrLf & outPath, vbInformation
End Sub

Public Sub סידור_רוחב_עמודות()
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim i As Long

    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case "לטיפול", "יומן_יישום", "בסיס_2022_מתוקן", "בסיס_2021_ייחוס", _
                 "חברות", "ענפים", "ענף מרכז", "סוכנים", "טלרים", "חודשים", _
                 "תקציר", "הפעלה"
                lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
                If lastCol < 1 Then lastCol = 1
                ws.Columns("A:" & Split(ws.Cells(1, lastCol).Address, "$")(1)).AutoFit
                For i = 1 To lastCol
                    If ws.Columns(i).ColumnWidth > 25 Then ws.Columns(i).ColumnWidth = 25
                Next i
                If ws.Columns("A").ColumnWidth < 14 Then ws.Columns("A").ColumnWidth = 14
        End Select
    Next ws
End Sub

Public Sub הפעלה_מלאה()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    On Error GoTo SafeExit

    Call טען_קובץ_מקור
    Call צור_קובץ_לטיפול
    Call טען_קובץ_תיקונים
    Call הפק_דוח_סופי

    MsgBox "התהליך הושלם בהצלחה", vbInformation

SafeExit:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    If Err.Number <> 0 Then
        MsgBox "אירעה שגיאה: " & Err.Description, vbExclamation
    End If
End Sub