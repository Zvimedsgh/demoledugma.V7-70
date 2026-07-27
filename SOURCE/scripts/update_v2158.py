import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.158.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_157"', 'Attribute VB_Name = "Goren_Claude_V2_158"')
content = content.replace('VERSION: V2.157', 'VERSION: V2.158')
content = content.replace('APP_VERSION As String = "2.157"', 'APP_VERSION As String = "2.158"')

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("Updated version to V2.158.")
