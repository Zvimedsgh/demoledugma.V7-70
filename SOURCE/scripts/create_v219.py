import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.218.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.219.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

old_shp = 'Set shp = wsInst.Shapes.AddShape(msoShapeRoundedRectangle, wsInst.Range("J20").Left, wsInst.Range("J20").Top, 250, 40)'
new_shp = 'Set shp = wsInst.Shapes.AddShape(msoShapeRoundedRectangle, wsInst.Range("A20").Left, wsInst.Range("A20").Top, 250, 40)'

content = content.replace(old_shp, new_shp)

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_218"', 'Attribute VB_Name = "Goren_Claude_V2_219"')
content = content.replace('VERSION: V2.218', 'VERSION: V2.219')
content = content.replace('APP_VERSION As String = "2.218"', 'APP_VERSION As String = "2.219"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.219")
