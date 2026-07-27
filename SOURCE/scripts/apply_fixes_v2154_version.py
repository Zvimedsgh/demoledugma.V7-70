import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_153"', 'Attribute VB_Name = "Goren_Claude_V2_154"')
content = content.replace('VERSION: V2.153', 'VERSION: V2.154')
content = content.replace('APP_VERSION As String = "2.153"', 'APP_VERSION As String = "2.154"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.154.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.154 created.")
