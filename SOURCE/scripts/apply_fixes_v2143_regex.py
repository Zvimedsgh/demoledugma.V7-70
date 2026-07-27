import sys, re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.142.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

pattern_exact = re.compile(r"(If bWildcard Then\s*If UCase\(cName\) Like UCase\(searchText\) Then bMatch = True\s*Else\s*)If UCase\(cName\) = UCase\(searchText\) Then bMatch = True")

if pattern_exact.search(content):
    content = pattern_exact.sub(r"\1If InStr(1, cName, searchText, vbTextCompare) > 0 Then bMatch = True", content)
    print("Reverted exact match to InStr with regex.")
else:
    print("Could not find exact match block with regex either.")

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_142"', 'Attribute VB_Name = "Goren_Claude_V2_143"')
content = content.replace('VERSION: V2.142', 'VERSION: V2.143')
content = content.replace('APP_VERSION As String = "2.142"', 'APP_VERSION As String = "2.143"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.143.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.143 updated.")
