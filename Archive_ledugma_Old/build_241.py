import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.240.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.240', 'VERSION: V2.241')
content = content.replace('Error in V2.240!', 'Error in V2.241!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_240"', 'Attribute VB_Name = "Goren_Claude_V2_241"')
content = content.replace('APP_VERSION As String = "2.240"', 'APP_VERSION As String = "2.241"')
content = content.replace('APP_VERSION = "2.240"', 'APP_VERSION = "2.241"')

old_with = """        With shpInstall.TextFrame2.TextRange
            .Text = ChrW(1492) & ChrW(1511) & ChrW(1513) & ChrW(32) & ChrW(1499) & ChrW(1488) & ChrW(1503) & ChrW(32) & ChrW(1500) & ChrW(1508) & ChrW(1514) & ChrW(1497) & ChrW(1495) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & vbCrLf & ChrW(1504) & ChrW(1497) & ChrW(1514) & ChrW(1503) & ChrW(32) & ChrW(1500) & ChrW(1513) & ChrW(1500) & ChrW(1493) & ChrW(1495) & ChrW(32) & "WhatsApp" & ChrW(32) & ChrW(1500) & ChrW(1496) & ChrW(1500) & ChrW(1508) & ChrW(1493) & ChrW(1503) & ChrW(32) & "054-6677396" & ChrW(32) & ChrW(1500) & ChrW(1511) & ChrW(1489) & ChrW(1500) & ChrW(1514) & ChrW(32) & ChrW(1506) & ChrW(1494) & ChrW(1512) & ChrW(1492)
            .Font.Size = 10
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 139)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With"""

new_with = """        With shpInstall.TextFrame2.TextRange
            .Text = ChrW(1497) & ChrW(1513) & ChrW(32) & ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & ChrW(32) & ChrW(1499) & ChrW(1488) & ChrW(1503) & ChrW(33) & vbCrLf & ChrW(1492) & ChrW(1511) & ChrW(1513) & ChrW(32) & ChrW(1500) & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & vbCrLf & ChrW(1500) & ChrW(1514) & ChrW(1502) & ChrW(1497) & ChrW(1499) & ChrW(1492) & ChrW(32) & ChrW(1489) & ChrW(1493) & ChrW(1493) & ChrW(1488) & ChrW(1496) & ChrW(1505) & ChrW(1488) & ChrW(1508) & ChrW(58) & ChrW(32) & ChrW(48) & ChrW(53) & ChrW(52) & ChrW(45) & ChrW(54) & ChrW(54) & ChrW(55) & ChrW(55) & ChrW(51) & ChrW(57) & ChrW(54)
            .Font.Size = 14
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 139)
            .ParagraphFormat.Alignment = msoAlignCenter
            .ParagraphFormat.TextDirection = 2
        End With"""

content = content.replace(old_with, new_with)

# Also increase height since there are now 3 lines and size 14
# Set shpInstall = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F15").Left, wsMain.Range("F15").Top, 220, 65)
old_shape = 'Set shpInstall = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F15").Left, wsMain.Range("F15").Top, 220, 65)'
new_shape = 'Set shpInstall = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F15").Left, wsMain.Range("F15").Top, 250, 90)'
content = content.replace(old_shape, new_shape)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.241.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 241')
