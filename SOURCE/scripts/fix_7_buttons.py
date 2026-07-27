import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.168.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.169.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_168"', 'Attribute VB_Name = "Goren_Claude_V2_169"')
content = content.replace('VERSION: V2.168', 'VERSION: V2.169')
content = content.replace('APP_VERSION As String = "2.168"', 'APP_VERSION As String = "2.169"')

old_buttons_block = """2880 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top, btnW, btnH)
2890 shp.Name = "btnBuildReview"
2900 shp.Fill.ForeColor.RGB = blueClr
2910 shp.TextFrame2.TextRange.Text = "1 - " & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1514) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
2920 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
2930 shp.TextFrame2.TextRange.Font.Size = 12
2940 shp.TextFrame2.TextRange.Font.Bold = msoTrue
2950 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
2960 shp.OnAction = "BuildReview"

2970 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 44, btnW, btnH)
2980 shp.Name = "btnApplyCorrections"
2990 shp.Fill.ForeColor.RGB = RGB(0, 150, 80)
3000 shp.TextFrame2.TextRange.Text = "2 - " & ChrW(1497) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1501) & " " & ChrW(1493) & ChrW(1491) & ChrW(1493) & ChrW(34) & ChrW(1495) & ChrW(1493) & ChrW(1514)
3010 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3020 shp.TextFrame2.TextRange.Font.Size = 12
3030 shp.TextFrame2.TextRange.Font.Bold = msoTrue
3040 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3050 shp.OnAction = "ApplyCorrectionsAndBuildReports"

3060 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 88, btnW, btnH)
3070 shp.Name = "btnBuildPresentation"
3080 shp.Fill.ForeColor.RGB = RGB(200, 100, 0)
3090 shp.TextFrame2.TextRange.Text = "3 - " & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1493) & ChrW(1512) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514)
3100 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3110 shp.TextFrame2.TextRange.Font.Size = 12
3120 shp.TextFrame2.TextRange.Font.Bold = msoTrue
3130 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3140 shp.OnAction = "BuildPresentation"

3150 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 132, btnW, btnH)
3160 shp.Name = "btnSaveReports"
3170 shp.Fill.ForeColor.RGB = blueClr
3180 shp.TextFrame2.TextRange.Text = "4 - " & ChrW(1513) & ChrW(1502) & ChrW(1497) & ChrW(1512) & ChrW(1514) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
3190 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3200 shp.TextFrame2.TextRange.Font.Size = 12
3210 shp.TextFrame2.TextRange.Font.Bold = msoTrue
3220 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3230 shp.OnAction = "SaveReportsToFolder"

3240 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 176, btnW, btnH)
3250 shp.Name = "btnViewReports"
3260 shp.Fill.ForeColor.RGB = RGB(0, 150, 80)
3270 shp.TextFrame2.TextRange.Text = "5 - " & ChrW(1510) & ChrW(1508) & ChrW(1497) & ChrW(1497) & ChrW(1492) & " " & ChrW(1489) & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
3280 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3290 shp.TextFrame2.TextRange.Font.Size = 12
3300 shp.TextFrame2.TextRange.Font.Bold = msoTrue
3310 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3320 shp.OnAction = "ViewReportsFolder"

3330 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 220, btnW, btnH)
3340 shp.Name = "btnNewClients"
3350 shp.Fill.ForeColor.RGB = RGB(200, 100, 0)
3360 shp.TextFrame2.TextRange.Text = "6 - " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
3370 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3380 shp.TextFrame2.TextRange.Font.Size = 12
3390 shp.TextFrame2.TextRange.Font.Bold = msoTrue
3400 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3410 shp.OnAction = "NewClients" """

new_buttons_block = """2880 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top, btnW, btnH)
2890 shp.Name = "btnBuildReview"
2900 shp.Fill.ForeColor.RGB = blueClr
2910 shp.TextFrame2.TextRange.Text = "1 - " & ChrW(1489) & ChrW(1491) & ChrW(1497) & ChrW(1511) & ChrW(1514) & " " & ChrW(1504) & ChrW(1514) & ChrW(1493) & ChrW(1504) & ChrW(1497) & ChrW(1501)
2920 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
2930 shp.TextFrame2.TextRange.Font.Size = 12
2940 shp.TextFrame2.TextRange.Font.Bold = msoTrue
2950 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
2960 shp.OnAction = "BuildReview"

2970 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 44, btnW, btnH)
2980 shp.Name = "btnApplyCorrections"
2990 shp.Fill.ForeColor.RGB = RGB(0, 150, 80)
3000 shp.TextFrame2.TextRange.Text = "2 - " & ChrW(1497) & ChrW(1497) & ChrW(1513) & ChrW(1493) & ChrW(1501) & " " & ChrW(1493) & ChrW(1491) & ChrW(1493) & ChrW(34) & ChrW(1495) & ChrW(1493) & ChrW(1514)
3010 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3020 shp.TextFrame2.TextRange.Font.Size = 12
3030 shp.TextFrame2.TextRange.Font.Bold = msoTrue
3040 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3050 shp.OnAction = "ApplyCorrectionsAndBuildReports"

3060 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 88, btnW, btnH)
3070 shp.Name = "btnShowResults"
3080 shp.Fill.ForeColor.RGB = RGB(0, 150, 80)
3090 shp.TextFrame2.TextRange.Text = "3 - " & ChrW(1492) & ChrW(1510) & ChrW(1490) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1504) & ChrW(1493) & ChrW(1514) & " " & ChrW(1491) & ChrW(1497) & ChrW(1493) & ChrW(1493) & ChrW(1495)
3100 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3110 shp.TextFrame2.TextRange.Font.Size = 12
3120 shp.TextFrame2.TextRange.Font.Bold = msoTrue
3130 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3140 shp.OnAction = "ShowResultSheets"

3150 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 132, btnW, btnH)
3160 shp.Name = "btnBuildPresentation"
3170 shp.Fill.ForeColor.RGB = RGB(200, 100, 0)
3180 shp.TextFrame2.TextRange.Text = "4 - " & ChrW(1497) & ChrW(1497) & ChrW(1510) & ChrW(1493) & ChrW(1512) & " " & ChrW(1502) & ChrW(1510) & ChrW(1490) & ChrW(1514)
3190 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3200 shp.TextFrame2.TextRange.Font.Size = 12
3210 shp.TextFrame2.TextRange.Font.Bold = msoTrue
3220 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3230 shp.OnAction = "BuildPresentation"

3240 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 176, btnW, btnH)
3250 shp.Name = "btnSaveReports"
3260 shp.Fill.ForeColor.RGB = blueClr
3270 shp.TextFrame2.TextRange.Text = "5 - " & ChrW(1513) & ChrW(1502) & ChrW(1497) & ChrW(1512) & ChrW(1514) & " " & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514)
3280 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3290 shp.TextFrame2.TextRange.Font.Size = 12
3300 shp.TextFrame2.TextRange.Font.Bold = msoTrue
3310 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3320 shp.OnAction = "SaveReportsToFolder"

3330 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 220, btnW, btnH)
3340 shp.Name = "btnViewReports"
3350 shp.Fill.ForeColor.RGB = RGB(0, 150, 80)
3360 shp.TextFrame2.TextRange.Text = "6 - " & ChrW(1510) & ChrW(1508) & ChrW(1497) & ChrW(1497) & ChrW(1492) & " " & ChrW(1489) & ChrW(1491) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1513) & ChrW(1502) & ChrW(1493) & ChrW(1512) & ChrW(1497) & ChrW(1501)
3370 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3380 shp.TextFrame2.TextRange.Font.Size = 12
3390 shp.TextFrame2.TextRange.Font.Bold = msoTrue
3400 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3410 shp.OnAction = "ViewReportsFolder"

3411 Set shp = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, wsMain.Range("C2").Top + 264, btnW, btnH)
3412 shp.Name = "btnNewClients"
3413 shp.Fill.ForeColor.RGB = RGB(200, 100, 0)
3414 shp.TextFrame2.TextRange.Text = "7 - " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & ChrW(1493) & ChrW(1514) & " " & ChrW(1495) & ChrW(1491) & ChrW(1513) & ChrW(1497) & ChrW(1501)
3415 shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3416 shp.TextFrame2.TextRange.Font.Size = 12
3417 shp.TextFrame2.TextRange.Font.Bold = msoTrue
3418 shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3419 shp.OnAction = "NewClients" """

# Also update AssignButtonMacros
old_assign = """Case "btnBuildReview": shp.OnAction = "BuildReview"
Case "btnApplyCorrections": shp.OnAction = "ApplyCorrectionsAndBuildReports"
Case "btnBuildPresentation": shp.OnAction = "BuildPresentation"
Case "btnSaveReports": shp.OnAction = "SaveReportsToFolder"
Case "btnViewReports": shp.OnAction = "ViewReportsFolder"
Case "btnNewClients": shp.OnAction = "NewClients" """

new_assign = """Case "btnBuildReview": shp.OnAction = "BuildReview"
Case "btnApplyCorrections": shp.OnAction = "ApplyCorrectionsAndBuildReports"
Case "btnShowResults": shp.OnAction = "ShowResultSheets"
Case "btnBuildPresentation": shp.OnAction = "BuildPresentation"
Case "btnSaveReports": shp.OnAction = "SaveReportsToFolder"
Case "btnViewReports": shp.OnAction = "ViewReportsFolder"
Case "btnNewClients": shp.OnAction = "NewClients" """

content = content.replace(old_buttons_block, new_buttons_block)
content = content.replace(old_assign, new_assign)

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.169 correctly!")
