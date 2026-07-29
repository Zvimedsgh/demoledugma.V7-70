import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.232.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.232', 'VERSION: V2.233')
content = content.replace('Error in V2.232!', 'Error in V2.233!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_232"', 'Attribute VB_Name = "Goren_Claude_V2_233"')
content = content.replace('APP_VERSION As String = "2.232"', 'APP_VERSION As String = "2.233"')

# Move shpManualMsg to bottom of row 1, middle of column L
old_str = 'Set shpManual = wsMain.Shapes.AddShape(msoShapeOval, wsMain.Range("L1").Left, wsMain.Range("L1").Top, 120, 80)'
new_str = 'Set shpManual = wsMain.Shapes.AddShape(msoShapeOval, wsMain.Range("L1").Left + (wsMain.Range("L1").Width - 120) / 2, wsMain.Range("L1").Top + wsMain.Range("L1").Height - 80, 120, 80)'
content = content.replace(old_str, new_str)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.233.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 233')
