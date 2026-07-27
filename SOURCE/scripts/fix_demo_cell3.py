import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.167.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.168.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_167"', 'Attribute VB_Name = "Goren_Claude_V2_168"')
content = content.replace('VERSION: V2.167', 'VERSION: V2.168')
content = content.replace('APP_VERSION As String = "2.167"', 'APP_VERSION As String = "2.168"')

old_block_demo = """    On Error Resume Next
    wsMain.Shapes("shpDemoMsgText").Delete
    On Error GoTo 0
    Dim shpDemoMsg As Shape
    Dim tLeft As Single, tWidth As Single
    tLeft = wsMain.Range("E19").Left
    tWidth = wsMain.Range("K19").Left + wsMain.Range("K19").Width - tLeft
    Set shpDemoMsg = wsMain.Shapes.AddTextbox(1, tLeft, wsMain.Range("F19").Top, tWidth, 40)
    shpDemoMsg.Name = "shpDemoMsgText"
    shpDemoMsg.TextFrame2.TextRange.Text = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " WhatsApp " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & ": " & ChrW(8207) & "054-6677396"
    shpDemoMsg.TextFrame2.TextRange.Font.Size = 17
    shpDemoMsg.TextFrame2.TextRange.ParagraphFormat.Alignment = 2 ' msoAlignCenter
    shpDemoMsg.TextFrame2.TextRange.ParagraphFormat.TextDirection = 2 ' msoTextDirectionRightToLeft
    shpDemoMsg.TextFrame2.TextRange.Font.Bold = -1
    shpDemoMsg.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 255)
    shpDemoMsg.Line.Visible = 0
    shpDemoMsg.Fill.Visible = 0"""

new_block_demo = """    On Error Resume Next
    wsMain.Shapes("shpDemoMsgText").Delete
    On Error GoTo 0
    With wsMain.Range("E19:K20")
        .Merge
        .Value = ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & " " & ChrW(1513) & ChrW(1500) & ChrW(1495) & " WhatsApp " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & ": 054-6677396"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 255)
    End With"""

old_block_else = """    wsMain.Range("G3:G4").Interior.ColorIndex = xlNone
    On Error Resume Next
    wsMain.Shapes("shpDemoLockG3G4").Delete
    wsMain.Shapes("shpDemoMsgText").Delete
    On Error GoTo 0
End If"""

new_block_else = """    wsMain.Range("G3:G4").Interior.ColorIndex = xlNone
    On Error Resume Next
    wsMain.Shapes("shpDemoLockG3G4").Delete
    wsMain.Shapes("shpDemoMsgText").Delete
    wsMain.Range("E19:K20").ClearContents
    On Error GoTo 0
End If"""


content = content.replace(old_block_demo, new_block_demo)
content = content.replace(old_block_else, new_block_else)

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.168 correctly!")
