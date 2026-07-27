Attribute VB_Name = "RebuildParams"
Public Sub RebuildParamsSheet()
    Dim ws As Worksheet
    Dim sheetName As String
    ' "הגדרות_פרמטרים"
    sheetName = ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514) & "_" & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512) & ChrW(1497) & ChrW(1501)
    
    ' Try to create new sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    If Err.Number <> 0 Then
        MsgBox "שגיאה: שם הגיליון כבר קיים. אנא שנה את שם הגיליון הישן קודם."
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Set RTL
    ws.DisplayRightToLeft = True
    
    ' Headers
    ws.Cells(1, 1).Value = ChrW(1513) & ChrW(1501) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512)
    ws.Cells(1, 2).Value = ChrW(1506) & ChrW(1512) & ChrW(1498) & " " & ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512)
    ws.Range("A1:B1").Font.Bold = True
    ws.Range("A1:B1").Interior.Color = RGB(200, 230, 255)
    ws.Columns("A:A").ColumnWidth = 25
    ws.Columns("B:B").ColumnWidth = 40
    
    ' Basic params
    ws.Cells(2, 1).Value = "PREMIUM_THRESHOLD"
    ws.Cells(2, 2).Value = 20
    ws.Cells(3, 1).Value = "ERROR_EMAIL"
    ws.Cells(3, 2).Value = "zvi@gorentech.co.il"
    ws.Cells(4, 1).Value = "BACKUP_PATH"
    ws.Cells(4, 2).Value = "C:\LEVAV PROJECT\BACKUPS"
    ws.Cells(5, 1).Value = "FILES_FOLDER"
    ws.Cells(5, 2).Value = "C:\" & ChrW(1508) & ChrW(1512) & ChrW(1493) & ChrW(1497) & ChrW(1511) & ChrW(1496) & " " & ChrW(1500) & ChrW(1489) & ChrW(1489) & "\SOURCE\"
    ws.Cells(6, 1).Value = "REPORTS_FOLDER"
    ws.Cells(6, 2).Value = "C:\" & ChrW(1508) & ChrW(1512) & ChrW(1493) & ChrW(1497) & ChrW(1511) & ChrW(1496) & " " & ChrW(1500) & ChrW(1489) & ChrW(1489) & "\REPORTS\"
    ws.Cells(7, 1).Value = "DEMO_MODE"
    ws.Cells(7, 2).Value = ChrW(1500) & ChrW(1488) ' לא
    ws.Cells(8, 1).Value = "AGENCY_NAME"
    ws.Cells(8, 2).Value = ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495) & " " & ChrW(1500) & ChrW(1489) & ChrW(1489)
    ws.Cells(9, 1).Value = "DEMO_AGENCY_NAME"
    ws.Cells(9, 2).Value = ChrW(1500) & ChrW(1491) & ChrW(1493) & ChrW(1490) & ChrW(1502) & ChrW(1488) & " " & ChrW(1505) & ChrW(1493) & ChrW(1499) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1500) & ChrW(1489) & ChrW(1497) & ChrW(1496) & ChrW(1493) & ChrW(1495)
    
    ' Add Back Button
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("E2").Left, ws.Range("E2").Top, 100, 30)
    btn.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1494) & ChrW(1512) & ChrW(1492) & " " & ChrW(1500) & ChrW(1492) & ChrW(1490) & ChrW(1491) & ChrW(1512) & ChrW(1493) & ChrW(1514)
    btn.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    btn.TextFrame2.TextRange.Font.Bold = msoTrue
    btn.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    btn.OnAction = "NavSettings_Menu"
    
    ws.Cells(10, 1).Value = "EOD"
    
    MsgBox "הגיליון הוקם בהצלחה!"
End Sub
