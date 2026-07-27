import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.216.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.217.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

old_shp = 'Set shp = wsInst.Shapes.AddShape(msoShapeRoundedRectangle, wsInst.Range("A2").Left, wsInst.Range("A2").Top, 250, 40)'
new_shp = 'Set shp = wsInst.Shapes.AddShape(msoShapeRoundedRectangle, wsInst.Range("B23").Left, wsInst.Range("B23").Top, 250, 40)'

content = content.replace(old_shp, new_shp)

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_216"', 'Attribute VB_Name = "Goren_Claude_V2_217"')
content = content.replace('VERSION: V2.216', 'VERSION: V2.217')
content = content.replace('APP_VERSION As String = "2.216"', 'APP_VERSION As String = "2.217"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.217")
