import sys, re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.143.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

pattern_match = re.compile(r"(If bWildcard Then\s*If UCase\(cName\) Like UCase\(searchText\) Then bMatch = True\s*Else\s*)If InStr\(1, cName, searchText, vbTextCompare\) > 0 Then bMatch = True")

new_match = r"""\1' Whole word match
Dim paddedName As String, paddedSearch As String
paddedName = " " & Replace(Replace(UCase(cName), "-", " "), "_", " ") & " "
paddedSearch = " " & Replace(Replace(UCase(searchText), "-", " "), "_", " ") & " "
If InStr(1, paddedName, paddedSearch, vbTextCompare) > 0 Then bMatch = True"""

if pattern_match.search(content):
    content = pattern_match.sub(new_match, content)
    print("Updated to Whole Word Match using regex.")
else:
    print("Could not find target match block with regex.")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_143"', 'Attribute VB_Name = "Goren_Claude_V2_144"')
content = content.replace('VERSION: V2.143', 'VERSION: V2.144')
content = content.replace('APP_VERSION As String = "2.143"', 'APP_VERSION As String = "2.144"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.144 updated.")
