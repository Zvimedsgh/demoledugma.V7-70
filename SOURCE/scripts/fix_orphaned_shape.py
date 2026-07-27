import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.169.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.170.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_169"', 'Attribute VB_Name = "Goren_Claude_V2_170"')
content = content.replace('VERSION: V2.169', 'VERSION: V2.170')
content = content.replace('APP_VERSION As String = "2.169"', 'APP_VERSION As String = "2.170"')

old_delete_block = """360 For Each shpBtn In wsMain.Shapes
370 If shpBtn.Type = 5 Then ' msoShapeRoundedRectangle
380 shpBtn.Delete
390 End If
400 Next shpBtn"""

new_delete_block = """360 For Each shpBtn In wsMain.Shapes
370 If shpBtn.Type = 5 Or shpBtn.Type = 17 Then ' msoShapeRoundedRectangle OR msoTextBox
380 shpBtn.Delete
390 End If
400 Next shpBtn"""

content = content.replace(old_delete_block, new_delete_block)

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.170 correctly!")
