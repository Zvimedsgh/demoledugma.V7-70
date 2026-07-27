import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.121.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: Move btnSearch and btnAll to H10
target_btns = """3490 Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F11").Left + (wsMain.Range("F11").Width - 60) / 2, wsMain.Range("F11").Top + 5, 60, 25)
3500 btnSearch.Name = "btnSearchClient"
3510 btnSearch.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513)
3520 btnSearch.TextFrame2.TextRange.Font.Size = 10
3530 btnSearch.TextFrame2.TextRange.Font.Bold = msoTrue
3540 btnSearch.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3550 btnSearch.Fill.ForeColor.RGB = RGB(0, 150, 80)
3560 btnSearch.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3570 btnSearch.OnAction = "SearchClientName"

3580 Set btnAll = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G11").Left + (wsMain.Range("G11").Width - 60) / 2, wsMain.Range("G11").Top + 5, 60, 25)
3590 btnAll.Name = "btnAllClients"
3600 btnAll.TextFrame2.TextRange.Text = ChrW(1499) & ChrW(1493) & ChrW(1500) & ChrW(1501)
3610 btnAll.TextFrame2.TextRange.Font.Size = 10
3620 btnAll.TextFrame2.TextRange.Font.Bold = msoTrue
3630 btnAll.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3640 btnAll.Fill.ForeColor.RGB = blueClr
3650 btnAll.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3660 btnAll.OnAction = "ResetClientFilter"

3670 Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnAll.Left + (btnSearch.Left + btnSearch.Width - btnAll.Left - 120) / 2, wsMain.Range("F11").Top + 32, 120, 25)"""

new_btns = """3490 Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("H10").Left + 5, wsMain.Range("H10").Top + 2, 50, 22)
3500 btnSearch.Name = "btnSearchClient"
3510 btnSearch.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513)
3520 btnSearch.TextFrame2.TextRange.Font.Size = 10
3530 btnSearch.TextFrame2.TextRange.Font.Bold = msoTrue
3540 btnSearch.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3550 btnSearch.Fill.ForeColor.RGB = RGB(0, 150, 80)
3560 btnSearch.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3570 btnSearch.OnAction = "SearchClientName"

3580 Set btnAll = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("H10").Left + 60, wsMain.Range("H10").Top + 2, 50, 22)
3590 btnAll.Name = "btnAllClients"
3600 btnAll.TextFrame2.TextRange.Text = ChrW(1499) & ChrW(1493) & ChrW(1500) & ChrW(1501)
3610 btnAll.TextFrame2.TextRange.Font.Size = 10
3620 btnAll.TextFrame2.TextRange.Font.Bold = msoTrue
3630 btnAll.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3640 btnAll.Fill.ForeColor.RGB = blueClr
3650 btnAll.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3660 btnAll.OnAction = "ResetClientFilter"

3670 Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F11").Left + (wsMain.Range("F11:G11").Width - 120) / 2, wsMain.Range("F11").Top + 5, 120, 25)"""

if target_btns in content:
    content = content.replace(target_btns, new_btns)
else:
    # try manual slice replacement
    idx = content.find("3490 Set btnSearch")
    if idx != -1:
        idx2 = content.find("3680 btnReset.Name = ", idx)
        if idx2 != -1:
            content = content[:idx] + new_btns + "\n" + content[idx2:]
            print("manual replace btns done")
        else:
            print("failed manual 2")
    else:
        print("failed manual 1")

# Fix 2: Demo message
target_demo = """1650 If isDemoMode Then
1660 wsMain.Range("G3").Value = 2024
1670 wsMain.Range("G4").Value = 2025
1680 End If"""

new_demo = """1650 If isDemoMode Then
1660 wsMain.Range("G3").Value = 2024
1670 wsMain.Range("G4").Value = 2025
    ' ADD DEMO MESSAGE
    On Error Resume Next
    wsMain.Shapes("shpDemoMsgText").Delete
    On Error GoTo ERR_HANDLER
    Dim shpDemoMsg As Shape
    Set shpDemoMsg = wsMain.Shapes.AddTextbox(msoTextOrientationHorizontal, wsMain.Range("C4").Left, wsMain.Range("C4").Top, 350, 40)
    shpDemoMsg.Name = "shpDemoMsgText"
    shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " Whatsapp " & ChrW(1500) & " 054-6677396"
    shpDemoMsg.TextFrame2.TextRange.Font.Size = 16
    shpDemoMsg.TextFrame2.TextRange.Font.Bold = msoTrue
    shpDemoMsg.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 255)
    shpDemoMsg.Line.Visible = msoFalse
    shpDemoMsg.Fill.Visible = msoFalse
Else
    On Error Resume Next
    wsMain.Shapes("shpDemoMsgText").Delete
    On Error GoTo ERR_HANDLER
1680 End If"""

if target_demo in content:
    content = content.replace(target_demo, new_demo)
else:
    print("Could not find target_demo")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_121"', 'Attribute VB_Name = "Goren_Claude_V2_122"')
content = content.replace('VERSION: V2.121', 'VERSION: V2.122')
content = content.replace('APP_VERSION As String = "2.121"', 'APP_VERSION As String = "2.122"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.122.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.122 created.")
