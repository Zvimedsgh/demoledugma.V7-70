import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.150.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target = """    250 If dict.Count = 0 Then Exit Sub"""
new_code = """    If dict.Count = 0 Then
        wsClients.Range("A2:A" & wsClients.Rows.Count).ClearContents
        Exit Sub
    End If"""

if target in content:
    content = content.replace(target, new_code)
    print("Fixed UpdateClientList clear.")
else:
    print("Could not find target string.")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_150"', 'Attribute VB_Name = "Goren_Claude_V2_151"')
content = content.replace('VERSION: V2.150', 'VERSION: V2.151')
content = content.replace('APP_VERSION As String = "2.150"', 'APP_VERSION As String = "2.151"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.151.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.151 created.")
