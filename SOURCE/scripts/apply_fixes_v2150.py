import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.149.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_145"', 'Attribute VB_Name = "Goren_Claude_V2_150"')
content = content.replace('VERSION: V2.145', 'VERSION: V2.150')
content = content.replace('APP_VERSION As String = "2.145"', 'APP_VERSION As String = "2.150"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.150.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.150 created.")
