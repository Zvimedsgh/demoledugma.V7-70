import re

with open('c:\\LEVAV PROJECT\\SOURCE\\Goren_Claude_V1.68.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Fix 1: Typo
old_typo = 'ChrW(1492) & ChrW(1502) & ChrW(1514) & ChrW(1497) & ChrW(1504) & ChrW(1504) & ChrW(1493)'
new_typo = 'ChrW(1492) & ChrW(1502) & ChrW(1514) & ChrW(1497) & ChrW(1504) & ChrW(1493)'
content = content.replace(old_typo, new_typo)

# Fix 2: White area color index
old_color = 'wsProgress1.Range("G15:K15").Interior.ColorIndex = xlNone'
new_color = 'wsProgress1.Range("G15:K15").Interior.Color = RGB(220, 240, 220)'
content = content.replace(old_color, new_color)

# Fix 3: Exit System button
old_button = '''      ' ---- Add Credit and Version to A19 ----
      wsMain.Range("A19:Z30").ClearContents
      wsMain.Range("A19").Value = Space(25) & ChrW(1504) & ChrW(1489) & ChrW(1504) & ChrW(1492) & " " & ChrW(1506) & ChrW(1500) & " " & ChrW(1497) & ChrW(1491) & ChrW(1497) & " " & ChrW(1490) & ChrW(1493) & ChrW(1512) & ChrW(1504) & ChrW(1496) & ChrW(1511) & " 054-6677396 - v" & APP_VERSION
      wsMain.Range("A19").Font.Color = RGB(0, 0, 0)
      wsMain.Range("A19").Font.Size = 12
      wsMain.Range("A19").Font.Bold = True
      wsMain.Range("A19").HorizontalAlignment = xlRight'''
      
new_button = '''      ' ---- Add Exit System Button (Row 21) ----
      On Error Resume Next
      wsMain.Shapes("btnNavExit").Delete
      On Error GoTo 0
      Dim shpExit As Shape
      Set shpExit = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("A21").Left + 10, wsMain.Range("A21").Top, 120, 25)
      shpExit.Name = "btnNavExit"
      shpExit.Fill.ForeColor.RGB = RGB(180, 0, 0)
      shpExit.TextFrame2.TextRange.Text = ChrW(1497) & ChrW(1513) & ChrW(1497) & ChrW(1488) & ChrW(1492) & " " & ChrW(1502) & ChrW(1492) & ChrW(1502) & ChrW(1506) & ChrW(1512) & ChrW(1499) & ChrW(1514)
      shpExit.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
      shpExit.TextFrame2.TextRange.Font.Size = 10
      shpExit.TextFrame2.TextRange.Font.Bold = msoTrue
      shpExit.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
      shpExit.OnAction = "ExitSystem"'''

content = content.replace(old_button, new_button)

# Fix 4: Unlock G3:G12
old_unlock = 'wsMain.Range("F3:F10").HorizontalAlignment = xlRight'
new_unlock = 'wsMain.Range("F3:F10").HorizontalAlignment = xlRight\n    wsMain.Range("G3:G12").Locked = False'
content = content.replace(old_unlock, new_unlock)

with open('c:\\LEVAV PROJECT\\SOURCE\\Goren_Claude_V1.68.bas', 'w', encoding='utf-8') as f:
    f.write(content)
