import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.111.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("pwd = InputBoxU(", "pwd = InputBox(")

# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_111"', 'Attribute VB_Name = "Goren_Claude_V2_112"')
content = content.replace('VERSION: V2.111', 'VERSION: V2.112')
content = content.replace('APP_VERSION As String = "2.111"', 'APP_VERSION As String = "2.112"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.112.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.112 created.")
