import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.245.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.245', 'VERSION: V2.246')
content = content.replace('Error in V2.245!', 'Error in V2.246!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_245"', 'Attribute VB_Name = "Goren_Claude_V2_246"')
content = content.replace('APP_VERSION As String = "2.245"', 'APP_VERSION As String = "2.246"')
content = content.replace('APP_VERSION = "2.245"', 'APP_VERSION = "2.246"')

# Restore native Hyperlink on the Shape so it works in Excel Web!
old_with = """        With shpInstall.TextFrame2.TextRange
            .Text = ChrW(1497) & ChrW(1513) & ChrW(32) & ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & ChrW(32) & ChrW(1499) & ChrW(1488) & ChrW(1503) & ChrW(33) & vbCrLf & ChrW(1492) & ChrW(1511) & ChrW(1513) & ChrW(32) & ChrW(1500) & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & vbCrLf & ChrW(1500) & ChrW(1514) & ChrW(1502) & ChrW(1497) & ChrW(1499) & ChrW(1492) & ChrW(32) & ChrW(1489) & ChrW(1493) & ChrW(1493) & ChrW(1488) & ChrW(1496) & ChrW(1505) & ChrW(1488) & ChrW(1508) & ChrW(58) & ChrW(32) & ChrW(48) & ChrW(53) & ChrW(52) & ChrW(45) & ChrW(54) & ChrW(54) & ChrW(55) & ChrW(55) & ChrW(51) & ChrW(57) & ChrW(54)
            .Font.Size = 14
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 139)
            .ParagraphFormat.Alignment = msoAlignCenter
            .ParagraphFormat.TextDirection = 2
        End With"""

new_with = """        With shpInstall.TextFrame2.TextRange
            .Text = ChrW(1497) & ChrW(1513) & ChrW(32) & ChrW(1500) & ChrW(1492) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & ChrW(32) & ChrW(1499) & ChrW(1488) & ChrW(1503) & ChrW(33) & vbCrLf & ChrW(1492) & ChrW(1511) & ChrW(1513) & ChrW(32) & ChrW(1500) & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1491) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(1493) & ChrW(1512) & ChrW(1488) & ChrW(1493) & ChrW(1514) & ChrW(32) & ChrW(1492) & ChrW(1514) & ChrW(1511) & ChrW(1504) & ChrW(1492) & vbCrLf & ChrW(1500) & ChrW(1514) & ChrW(1502) & ChrW(1497) & ChrW(1499) & ChrW(1492) & ChrW(32) & ChrW(1489) & ChrW(1493) & ChrW(1493) & ChrW(1488) & ChrW(1496) & ChrW(1505) & ChrW(1488) & ChrW(1508) & ChrW(58) & ChrW(32) & ChrW(48) & ChrW(53) & ChrW(52) & ChrW(45) & ChrW(54) & ChrW(54) & ChrW(55) & ChrW(55) & ChrW(51) & ChrW(57) & ChrW(54)
            .Font.Size = 14
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 139)
            .ParagraphFormat.Alignment = msoAlignCenter
            .ParagraphFormat.TextDirection = 2
        End With
        ' Restore Native Hyperlink for Excel Web compatibility!
        wsMain.Hyperlinks.Add Anchor:=shpInstall, Address:="https://gorentec-my.sharepoint.com/:x:/g/personal/zvi_gorentech_co_il/IQBnB0klIujxR4O5fALynO8EAWddhgVppMJI7THxhW3R6fo?e=Z8iVEV"
"""

content = content.replace(old_with, new_with)

# Remove the OnAction assignment for shpInstall ONLY
content = content.replace('shpInstall.OnAction = "OpenInstallInstructions"', '')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.246.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 246')
