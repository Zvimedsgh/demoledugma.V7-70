import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.180.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.181.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace xlCenter with xlRight for the WhatsApp message
content = content.replace('.HorizontalAlignment = xlCenter\n        .VerticalAlignment = xlCenter\n        .Font.Size = 18\n        .Font.Bold = True\n        .Font.Color = RGB(0, 0, 255)', 
                          '.HorizontalAlignment = xlRight\n        .VerticalAlignment = xlCenter\n        .Font.Size = 18\n        .Font.Bold = True\n        .Font.Color = RGB(0, 0, 255)')

# Also replace D19:K20 with just D19:H20 to keep it from stretching too far if they don't want it to
# Actually, keeping D19:K20 with xlRight is safe, it just aligns to the right of D.

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_180"', 'Attribute VB_Name = "Goren_Claude_V2_181"')
content = content.replace('VERSION: V2.180', 'VERSION: V2.181')
content = content.replace('APP_VERSION As String = "2.180"', 'APP_VERSION As String = "2.181"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.181 correctly!")
