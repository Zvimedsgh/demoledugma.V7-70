import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.221.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.222.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

old_shp = 'Set shp = wsInst.Shapes.AddShape(msoShapeRoundedRectangle, wsInst.Range("B20").Left, wsInst.Range("B20").Top, 250, 40)'
new_shp = 'Set shp = wsInst.Shapes.AddShape(msoShapeRoundedRectangle, wsInst.Range("A20").Left + 150, wsInst.Range("A20").Top, 250, 40)'

content = content.replace(old_shp, new_shp)

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_221"', 'Attribute VB_Name = "Goren_Claude_V2_222"')
content = content.replace('VERSION: V2.221', 'VERSION: V2.222')
content = content.replace('APP_VERSION As String = "2.221"', 'APP_VERSION As String = "2.222"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.222")
