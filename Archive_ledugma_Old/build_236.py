import re

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.235.bas', 'r', encoding='utf-8') as f:
    content = f.read()

# Update version and name
content = content.replace('VERSION: V2.235', 'VERSION: V2.236')
content = content.replace('Error in V2.235!', 'Error in V2.236!')
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_235"', 'Attribute VB_Name = "Goren_Claude_V2_236"')
content = content.replace('APP_VERSION As String = "2.235"', 'APP_VERSION As String = "2.236"')
content = content.replace('APP_VERSION = "2.235"', 'APP_VERSION = "2.236"')

# Fix syntax error
content = content.replace('CreateObject("WScript.Shell").Run  & url &', 'CreateObject("WScript.Shell").Run url')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.236.bas', 'w', encoding='utf-8') as f:
    f.write(content)
print('Generated 236')
