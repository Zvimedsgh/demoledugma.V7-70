Public Sub SetupSettingsMenu()
10  Dim wsMgmt As Worksheet
20  Dim shp As Shape
30  Dim s As Shape
40  Dim btnTop As Double
50  Dim btnLeft As Double
60  Dim btnW As Double
70  Dim btnH As Double
80  Dim btnGap As Double
    
90  On Error Resume Next
100 Set wsMgmt = ThisWorkbook.Worksheets(MANAGEMENT_SHEET_NAME())
110 On Error GoTo 0
120 If wsMgmt Is Nothing Then Exit Sub
    
130 Application.ScreenUpdating = False

    ' Make top rows larger to hold floating menu
132 wsMgmt.Rows("1:3").RowHeight = 35
    
    ' Remove old menu buttons
140 On Error Resume Next
150 For Each s In wsMgmt.Shapes
160     If Left(s.Name, 6) = "btnNav" Then s.Delete
170 Next s
180 On Error GoTo 0
    
    ' Freeze panes at row 4 so menu is always visible
182 wsMgmt.Activate
184 ActiveWindow.FreezePanes = False
186 Application.Goto wsMgmt.Range("A4")
188 ActiveWindow.FreezePanes = True
    
    ' Horizontal Floating Menu Dimensions
210 btnW = 140
220 btnH = 26
    ' In RTL, moving LEFT means INCREASING the Left value
    ' We place them starting at 10 points from the right margin
230 btnGap = btnW + 10
    
    ' 1. Field Mapping button (blue) - ROW 1
    btnLeft = 10
    btnTop = 5
280 Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
290 shp.Name = "btnNavFieldMap"
300 shp.Fill.ForeColor.RGB = RGB(0, 70, 140)
310 shp.TextFrame2.TextRange.Text = ChrW(1502) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1497) & " " & ChrW(1513) & ChrW(1491) & ChrW(1493) & ChrW(1514)
320 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
330 shp.TextFrame2.TextRange.Font.Size = 11
340 shp.TextFrame2.TextRange.Font.Bold = msoTrue
350 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
360 shp.OnAction = "NavSettings_FieldMap"
370 btnLeft = btnLeft + btnGap
    
    ' 2. Branch Names button (dark cyan) - ROW 1
380 Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
390 shp.Name = "btnNavBranchName"
400 shp.Fill.ForeColor.RGB = RGB(0, 100, 100)
410 shp.TextFrame2.TextRange.Text = ChrW(1506) & ChrW(1504) & ChrW(1508) & ChrW(1497) & ChrW(1501)
420 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
430 shp.TextFrame2.TextRange.Font.Size = 11
440 shp.TextFrame2.TextRange.Font.Bold = msoTrue
450 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
460 shp.OnAction = "NavSettings_BranchName"
470 btnLeft = btnLeft + btnGap
    
    ' 3. Parameters button (dark green) - ROW 1
480 Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
490 shp.Name = "btnNavParams"
500 shp.Fill.ForeColor.RGB = RGB(0, 100, 0)
510 shp.TextFrame2.TextRange.Text = ChrW(1508) & ChrW(1512) & ChrW(1502) & ChrW(1496) & ChrW(1512) & ChrW(1497) & ChrW(1501)
520 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
530 shp.TextFrame2.TextRange.Font.Size = 11
540 shp.TextFrame2.TextRange.Font.Bold = msoTrue
550 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
560 shp.OnAction = "NavSettings_Params"
570 btnLeft = btnLeft + btnGap
    
    ' 4. Reason Codes button (purple) - ROW 1
580 Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
590 shp.Name = "btnNavReasonCode"
600 shp.Fill.ForeColor.RGB = RGB(100, 60, 140)
610 shp.TextFrame2.TextRange.Text = ChrW(1505) & ChrW(1497) & ChrW(1489) & ChrW(1493) & ChrW(1514)
620 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
630 shp.TextFrame2.TextRange.Font.Size = 11
640 shp.TextFrame2.TextRange.Font.Bold = msoTrue
650 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
660 shp.OnAction = "NavSettings_ReasonCode"
670 btnLeft = btnLeft + btnGap
    
    ' 5. Period Lists button (orange) - ROW 1
680 Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
690 shp.Name = "btnNavPeriodType"
700 shp.Fill.ForeColor.RGB = RGB(200, 100, 0)
710 shp.TextFrame2.TextRange.Text = H_RESHIMOT() & " " & ChrW(1514) & ChrW(1511) & ChrW(1493) & ChrW(1508) & ChrW(1492)
720 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
730 shp.TextFrame2.TextRange.Font.Size = 11
740 shp.TextFrame2.TextRange.Font.Bold = msoTrue
750 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
760 shp.OnAction = "NavSettings_PeriodType"
    
    ' Reset btnLeft to 10 for ROW 2, and increase btnTop
    btnLeft = 10
    btnTop = btnTop + btnH + 5
    
    ' 6. System Messages button (teal) - ROW 2
780 Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
790 shp.Name = "btnNavMessages"
800 shp.Fill.ForeColor.RGB = RGB(0, 128, 128)
810 shp.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1493) & ChrW(1491) & ChrW(1506) & ChrW(1493) & ChrW(1514)
820 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
830 shp.TextFrame2.TextRange.Font.Size = 11
840 shp.TextFrame2.TextRange.Font.Bold = msoTrue
850 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
860 shp.OnAction = "NavSettings_Messages"
870 btnLeft = btnLeft + btnGap
    
    ' 7. Permissions button (dark brown) - ROW 2
880 Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
890 shp.Name = "btnNavPermissions"
900 shp.Fill.ForeColor.RGB = RGB(120, 60, 0)
910 shp.TextFrame2.TextRange.Text = ChrW(1492) & ChrW(1512) & ChrW(1513) & ChrW(1488) & ChrW(1493) & ChrW(1514)
920 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
930 shp.TextFrame2.TextRange.Font.Size = 11
940 shp.TextFrame2.TextRange.Font.Bold = msoTrue
950 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
960 shp.OnAction = "NavSettings_Permissions"
970 btnLeft = btnLeft + btnGap
    
    ' 8. Client List button (steel blue) - ROW 2
980 Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
990 shp.Name = "btnNavClients"
1000 shp.Fill.ForeColor.RGB = RGB(70, 130, 180)
1010 shp.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
1020 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
1030 shp.TextFrame2.TextRange.Font.Size = 11
1040 shp.TextFrame2.TextRange.Font.Bold = msoTrue
1050 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
1060 shp.OnAction = "NavSettings_Clients"
1070 btnLeft = btnLeft + btnGap
    
    ' 9. Back to Home button (red) - ROW 2
1080 Set shp = wsMgmt.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnW, btnH)
1090 shp.Name = "btnNavHome"
1100 shp.Fill.ForeColor.RGB = RGB(180, 0, 0)
1110 shp.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1494) & ChrW(1512) & ChrW(1492) & " " & ChrW(1500) & ChrW(1491) & ChrW(1507) & " " & ChrW(1492) & ChrW(1489) & ChrW(1497) & ChrW(1514)
1120 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
1130 shp.TextFrame2.TextRange.Font.Size = 11
1140 shp.TextFrame2.TextRange.Font.Bold = msoTrue
1150 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
1160 shp.OnAction = "NavToIndex"
    
1180 Application.ScreenUpdating = True
End Sub
