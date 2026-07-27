import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.219.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.220.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

old_shp = 'Set shp = wsInst.Shapes.AddShape(msoShapeRoundedRectangle, wsInst.Range("A20").Left, wsInst.Range("A20").Top, 250, 40)'
new_shp = 'Set shp = wsInst.Shapes.AddShape(msoShapeRoundedRectangle, wsInst.Range("C20").Left, wsInst.Range("C20").Top, 250, 40)'

content = content.replace(old_shp, new_shp)

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_219"', 'Attribute VB_Name = "Goren_Claude_V2_220"')
content = content.replace('VERSION: V2.219', 'VERSION: V2.220')
content = content.replace('APP_VERSION As String = "2.219"', 'APP_VERSION As String = "2.220"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.220")
