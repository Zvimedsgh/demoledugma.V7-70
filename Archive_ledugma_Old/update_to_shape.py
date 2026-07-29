with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re

# Find the block I added previously
pattern = r'    \' Add the 3-line installation message in A2:B4.*?End With'

replacement = """    ' Clear previous mess
    wsMain.Range("B14:K15").Clear
    With wsMain.Range("A2:B4")
        .UnMerge
        .Clear
    End With
    
    ' Add a floating Shape instead
    On Error Resume Next
    wsMain.Shapes("shpInstallMsg").Delete
    On Error GoTo ERR_HANDLER
    
    Dim shpInstall As Shape
    ' Place it in A2, adjust size
    Set shpInstall = wsMain.Shapes.AddShape(msoShapeRectangle, wsMain.Range("A2").Left + 10, wsMain.Range("A2").Top, 220, 65)
    shpInstall.Name = "shpInstallMsg"
    shpInstall.Fill.ForeColor.RGB = RGB(245, 245, 245)
    shpInstall.Line.ForeColor.RGB = RGB(0, 0, 139)
    shpInstall.Line.Weight = 2
    
    With shpInstall.TextFrame2.TextRange
        .Text = ChrW(1512) & ChrW(1488) & ChrW(1492) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & vbCrLf & _
                 ChrW(1500) & ChrW(1506) & ChrW(1494) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & vbCrLf & _
                 ChrW(1513) & ChrW(1500) & ChrW(1495) & " " & "W" & "h" & "a" & "t" & "s" & "A" & "p" & "p" & " " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & " " & "0" & "5" & "4" & "-" & "6" & "6" & "7" & "7" & "3" & "9" & "6"
        .Font.Size = 12
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(0, 0, 139)
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    shpInstall.TextFrame2.VerticalAnchor = msoAnchorMiddle"""

text = re.sub(pattern, replacement, text, flags=re.DOTALL)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated VBA to use a Shape successfully.")
