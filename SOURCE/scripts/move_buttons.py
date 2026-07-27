import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.115.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update A00_SetupMainSheet
old_btns = """3490 Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F11").Left + (wsMain.Range("F11").Width - 60) / 2, wsMain.Range("F11").Top + 5, 60, 25)
3500 btnSearch.Name = "btnSearchClient"
3510 btnSearch.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513)
3520 btnSearch.TextFrame2.TextRange.Font.Size = 11
3530 btnSearch.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3540 btnSearch.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3550 btnSearch.Fill.ForeColor.RGB = RGB(0, 112, 192)
3560 btnSearch.Line.Visible = msoFalse
3570 btnSearch.OnAction = "SearchClientName"

3580 Set btnAll = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G11").Left + (wsMain.Range("G11").Width - 60) / 2, wsMain.Range("G11").Top + 5, 60, 25)
3590 btnAll.Name = "btnAllClients"
3600 btnAll.TextFrame2.TextRange.Text = ChrW(1499) & ChrW(1493) & ChrW(1500) & ChrW(1501)
3610 btnAll.TextFrame2.TextRange.Font.Size = 11
3620 btnAll.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3630 btnAll.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3640 btnAll.Fill.ForeColor.RGB = RGB(0, 176, 80)
3650 btnAll.Line.Visible = msoFalse
3660 btnAll.OnAction = "ResetClientFilter"

3670 Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, btnAll.Left + (btnSearch.Left + btnSearch.Width - btnAll.Left - 120) / 2, wsMain.Range("F11").Top + 32, 120, 25)"""

new_btns = """3490 Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G10").Left, wsMain.Range("G10").Top, wsMain.Range("G10").Width, wsMain.Range("G10").Height)
3500 btnSearch.Name = "btnSearchClient"
3510 btnSearch.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513)
3520 btnSearch.TextFrame2.TextRange.Font.Size = 11
3530 btnSearch.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3540 btnSearch.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3550 btnSearch.Fill.ForeColor.RGB = RGB(0, 112, 192)
3560 btnSearch.Line.Visible = msoFalse
3570 btnSearch.OnAction = "SearchClientName"

3580 Set btnAll = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("H10").Left, wsMain.Range("H10").Top, wsMain.Range("H10").Width, wsMain.Range("H10").Height)
3590 btnAll.Name = "btnAllClients"
3600 btnAll.TextFrame2.TextRange.Text = ChrW(1499) & ChrW(1493) & ChrW(1500) & ChrW(1501)
3610 btnAll.TextFrame2.TextRange.Font.Size = 11
3620 btnAll.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3630 btnAll.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3640 btnAll.Fill.ForeColor.RGB = RGB(0, 176, 80)
3650 btnAll.Line.Visible = msoFalse
3660 btnAll.OnAction = "ResetClientFilter"

3670 Set btnReset = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G11").Left, wsMain.Range("G11").Top, wsMain.Range("G11:H11").Width, wsMain.Range("G11").Height)"""

content = content.replace(old_btns, new_btns)

# 2. Update ConfirmClientSelection to change btnSearch text
old_confirm = """wsMain.Range("rngClientName").Value = selectedName

' Delete search sheet and go back to home"""

new_confirm = """wsMain.Range("rngClientName").Value = selectedName

On Error Resume Next
wsMain.Shapes("btnSearchClient").TextFrame2.TextRange.Text = selectedName
On Error GoTo 0

' Delete search sheet and go back to home"""

content = content.replace(old_confirm, new_confirm)

# 3. Update ResetClientFilter to reset btnSearch text
old_reset = """wsMain.Range("rngClientName").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)  ' "???/?"
Application.EnableEvents = True"""

new_reset = """wsMain.Range("rngClientName").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)  ' "???/?"
On Error Resume Next
wsMain.Shapes("btnSearchClient").TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513) ' "???"
On Error GoTo 0
Application.EnableEvents = True"""

content = content.replace(old_reset, new_reset)

# 4. Update ResetHomeDefaults
old_homereset = """wsMain.Range("G10").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497)
wsMain.Range("G10").Interior.Color = RGB(220, 240, 220)"""

new_homereset = """wsMain.Range("G10").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497)
wsMain.Range("G10").Interior.Color = RGB(220, 240, 220)
On Error Resume Next
wsMain.Shapes("btnSearchClient").TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513) ' "???"
On Error GoTo 0"""

content = content.replace(old_homereset, new_homereset)


# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_115"', 'Attribute VB_Name = "Goren_Claude_V2_116"')
content = content.replace('VERSION: V2.115', 'VERSION: V2.116')
content = content.replace('APP_VERSION As String = "2.115"', 'APP_VERSION As String = "2.116"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.116.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.116 created.")
