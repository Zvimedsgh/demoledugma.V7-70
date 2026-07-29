with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'r', encoding='utf-8') as f:
    text = f.read()

import re
pattern = r'With wsMain\.Range\("B14:K15"\)\s*\.Merge.*?End With'

replacement = """    ' Add the 3-line installation message in A2:B4
    wsMain.Range("B14:K15").Clear
    With wsMain.Range("A2:B4")
        .Merge
        .Value = ChrW(1512) & ChrW(1488) & ChrW(1492) & " " & ChrW(1490) & ChrW(1497) & ChrW(1500) & ChrW(1497) & ChrW(1493) & ChrW(1503) & " " & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & vbCrLf & _
                 ChrW(1500) & ChrW(1506) & ChrW(1494) & ChrW(1512) & ChrW(1492) & " " & ChrW(1489) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & " " & ChrW(1493) & ChrW(1492) & ChrW(1491) & ChrW(1512) & ChrW(1499) & ChrW(1492) & vbCrLf & _
                 ChrW(1513) & ChrW(1500) & ChrW(1495) & " " & "W" & "h" & "a" & "t" & "s" & "A" & "p" & "p" & " " & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & " " & "0" & "5" & "4" & "-" & "6" & "6" & "7" & "7" & "3" & "9" & "6"
        .HorizontalAlignment = xlLeft ' Align left so it is close to column B
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 139) ' Dark Blue
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
        .Borders.Color = RGB(0, 0, 139)
        .Interior.Color = RGB(245, 245, 245) ' Light gray background
    End With"""

text = re.sub(pattern, replacement, text, flags=re.DOTALL)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas', 'w', encoding='utf-8') as f:
    f.write(text)
print("Updated VBA for A2:B4 successfully.")
