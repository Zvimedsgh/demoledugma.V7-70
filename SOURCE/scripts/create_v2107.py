import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.106.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = re.sub(r'\n\s*\d*\s*MsgBoxU ChrW\(1499\) & ChrW\(1500\) & " " & ChrW\(1492\) & ChrW\(1490\).*?vbInformation', '', content)

# Version string
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_106"', 'Attribute VB_Name = "Goren_Claude_V2_107"')
content = content.replace('VERSION: V2.106', 'VERSION: V2.107')
content = content.replace('APP_VERSION As String = "2.106"', 'APP_VERSION As String = "2.107"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.107.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.107 created.")
