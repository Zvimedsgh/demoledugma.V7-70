import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.040_20260702_1500.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Private Const APP_VERSION As String = "2.037"', 'Private Const APP_VERSION As String = "2.040"')
content = content.replace('Private Const APP_VERSION As String = "2.036"', 'Private Const APP_VERSION As String = "2.040"')

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Updated APP_VERSION to 2.040")
