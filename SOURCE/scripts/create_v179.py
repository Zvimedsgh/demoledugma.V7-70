import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.178.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.179.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix the MarginRight error
content = content.replace('shp.TextFrame2.TextRange.MarginRight = 10', 'shp.TextFrame2.MarginRight = 10')

# Update version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_178"', 'Attribute VB_Name = "Goren_Claude_V2_179"')
content = content.replace('VERSION: V2.178', 'VERSION: V2.179')
content = content.replace('APP_VERSION As String = "2.178"', 'APP_VERSION As String = "2.179"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.179 correctly!")
