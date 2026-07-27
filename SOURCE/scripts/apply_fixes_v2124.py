import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.123.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: SetupMainSheet button sizes
target_btns = """3490 Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("H10").Left + 5, wsMain.Range("H10").Top + 2, 50, 22)
3500 btnSearch.Name = "btnSearchClient"
3510 btnSearch.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513)
3520 btnSearch.TextFrame2.TextRange.Font.Size = 10
3530 btnSearch.TextFrame2.TextRange.Font.Bold = msoTrue
3540 btnSearch.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3550 btnSearch.Fill.ForeColor.RGB = RGB(0, 150, 80)
3560 btnSearch.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3570 btnSearch.OnAction = "SearchClientName"

3580 Set btnAll = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("H10").Left + 60, wsMain.Range("H10").Top + 2, 50, 22)"""

new_btns = """3490 Set btnSearch = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("G10").Left, wsMain.Range("G10").Top, wsMain.Range("G10").Width, wsMain.Range("G10").Height)
3500 btnSearch.Name = "btnSearchClient"
3510 btnSearch.TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513)
3520 btnSearch.TextFrame2.TextRange.Font.Size = 10
3530 btnSearch.TextFrame2.TextRange.Font.Bold = msoTrue
3540 btnSearch.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
3550 btnSearch.Fill.ForeColor.RGB = RGB(0, 150, 80)
3560 btnSearch.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
3570 btnSearch.OnAction = "SearchClientName"

3580 Set btnAll = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("H10").Left + 2, wsMain.Range("H10").Top, wsMain.Range("H10").Width - 4, wsMain.Range("H10").Height)"""

if target_btns in content:
    content = content.replace(target_btns, new_btns)
else:
    print("Could not find target_btns")


# Fix 2: SearchClientName updating button text
target_search = """wsMain.Range("rngClientName").Value = selectedName"""
new_search = """wsMain.Range("rngClientName").Value = selectedName
        On Error Resume Next
        wsMain.Shapes("btnSearchClient").TextFrame2.TextRange.Text = selectedName
        On Error GoTo ERR_HANDLER"""

if target_search in content:
    content = content.replace(target_search, new_search)
else:
    print("Could not find target_search")


# Fix 3: ResetClientFilter updating button text
target_reset_client = """wsMain.Range("rngClientName").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & ChrW(47) & ChrW(1497)  ' "???/?"
End Sub"""
new_reset_client = """wsMain.Range("rngClientName").Value = ""
On Error Resume Next
wsMain.Shapes("btnSearchClient").TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513) ' "?????"
On Error GoTo 0
End Sub"""

if target_reset_client in content:
    content = content.replace(target_reset_client, new_reset_client)
else:
    print("Could not find target_reset_client")


# Fix 4: ResetHomeDefaults
target_reset_home = """wsMain.Range("G10").Value = ChrW(1489) & ChrW(1495) & ChrW(1512) & "/" & ChrW(1497)
wsMain.Range("G10").Interior.Color = RGB(220, 240, 220)"""
new_reset_home = """wsMain.Range("G10").Value = ""
wsMain.Range("G10").Interior.Color = RGB(220, 240, 220)
On Error Resume Next
wsMain.Shapes("btnSearchClient").TextFrame2.TextRange.Text = ChrW(1495) & ChrW(1508) & ChrW(1513) ' "?????"
On Error GoTo 0"""

if target_reset_home in content:
    content = content.replace(target_reset_home, new_reset_home)
else:
    print("Could not find target_reset_home")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_123"', 'Attribute VB_Name = "Goren_Claude_V2_124"')
content = content.replace('VERSION: V2.123', 'VERSION: V2.124')
content = content.replace('APP_VERSION As String = "2.123"', 'APP_VERSION As String = "2.124"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.124.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.124 created.")
