import sys
import re

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.177.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.178.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix the duplicate line numbers for MarginRight
content = re.sub(r'(\d+)(\s+shp\.TextFrame2\.MarginRight = 10)', r'\2', content)

# Update version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_177"', 'Attribute VB_Name = "Goren_Claude_V2_178"')
content = content.replace('VERSION: V2.177', 'VERSION: V2.178')
content = content.replace('APP_VERSION As String = "2.177"', 'APP_VERSION As String = "2.178"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.178 correctly!")
