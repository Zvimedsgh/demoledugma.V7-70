import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.212.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.213.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

# Fix syntax errors caused by \ before '
content = content.replace("3 \\' Max", "3 ' Max")
content = content.replace("\\' Show success message", "' Show success message")

# Update versions
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_212"', 'Attribute VB_Name = "Goren_Claude_V2_213"')
content = content.replace('VERSION: V2.212', 'VERSION: V2.213')
content = content.replace('APP_VERSION As String = "2.212"', 'APP_VERSION As String = "2.213"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.213")
