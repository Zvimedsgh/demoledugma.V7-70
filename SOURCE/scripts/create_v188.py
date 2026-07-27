import sys

filepath_in = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.187.bas'
filepath_out = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.188.bas'

with open(filepath_in, 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace("If psw = ADMIN_PASSWORD() Then", 'If psw = "Z961814r" Then')

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_187"', 'Attribute VB_Name = "Goren_Claude_V2_188"')
content = content.replace('VERSION: V2.187', 'VERSION: V2.188')
content = content.replace('APP_VERSION As String = "2.187"', 'APP_VERSION As String = "2.188"')

with open(filepath_out, 'w', encoding='utf-8') as f:
    f.write(content)

print("Created V2.188 correctly!")
