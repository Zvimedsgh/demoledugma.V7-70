import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.171.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.172.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Update version
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_171"', 'Attribute VB_Name = "Goren_Claude_V2_172"')
content = content.replace('VERSION: V2.171', 'VERSION: V2.172')
content = content.replace('APP_VERSION As String = "2.171"', 'APP_VERSION As String = "2.172"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.172 correctly!")
