import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.140.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_exact_match = """If bWildcard Then
If UCase(cName) Like UCase(searchText) Then bMatch = True
Else
If InStr(1, cName, searchText, vbTextCompare) > 0 Then bMatch = True
End If"""

new_exact_match = """If bWildcard Then
If UCase(cName) Like UCase(searchText) Then bMatch = True
Else
If UCase(cName) = UCase(searchText) Then bMatch = True
End If"""

if target_exact_match in content:
    content = content.replace(target_exact_match, new_exact_match)
else:
    print("Could not find exact match logic!")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_140"', 'Attribute VB_Name = "Goren_Claude_V2_141"')
content = content.replace('VERSION: V2.140', 'VERSION: V2.141')
content = content.replace('APP_VERSION As String = "2.140"', 'APP_VERSION As String = "2.141"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.141.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.141 created.")
