import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.238.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.238', 'VERSION: V2.239')
content = content.replace('Error in V2.238!', 'Error in V2.239!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_238"', 'Attribute VB_Name = "Goren_Claude_V2_239"')
content = content.replace('APP_VERSION As String = "2.238"', 'APP_VERSION As String = "2.239"')
content = content.replace('APP_VERSION = "2.238"', 'APP_VERSION = "2.239"')

# Move to F15 and change color
old_pos = 'Set shpInstall = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("A2").Left + 10, wsMain.Range("A2").Top, 220, 65)'
new_pos = 'Set shpInstall = wsMain.Shapes.AddShape(msoShapeRoundedRectangle, wsMain.Range("F15").Left, wsMain.Range("F15").Top, 220, 65)'
content = content.replace(old_pos, new_pos)

old_color = 'shpInstall.Fill.ForeColor.RGB = RGB(245, 245, 245)'
new_color = 'shpInstall.Fill.ForeColor.RGB = RGB(255, 204, 229) \' Pastel Pink'
content = content.replace(old_color, new_color)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.239.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 239')
