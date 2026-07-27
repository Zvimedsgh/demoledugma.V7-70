import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.143.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

target_match = """If bWildcard Then
If UCase(cName) Like UCase(searchText) Then bMatch = True
Else
If InStr(1, cName, searchText, vbTextCompare) > 0 Then bMatch = True
End If"""

new_match = """If bWildcard Then
If UCase(cName) Like UCase(searchText) Then bMatch = True
Else
' Whole word match
Dim paddedName As String, paddedSearch As String
paddedName = " " & Replace(Replace(UCase(cName), "-", " "), "_", " ") & " "
paddedSearch = " " & Replace(Replace(UCase(searchText), "-", " "), "_", " ") & " "
If InStr(1, paddedName, paddedSearch, vbTextCompare) > 0 Then bMatch = True
End If"""

if target_match in content:
    content = content.replace(target_match, new_match)
    print("Updated to Whole Word Match.")
else:
    print("Could not find target match block.")

# Let's fix the instructions again to match Whole Word Match
target_inst1 = 'wsSearch.Cells(6, 5).Value = "- " & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513) & " " & ChrW(1512) & ChrW(1490) & ChrW(1497) & ChrW(1500) & " (" & ChrW(1500) & ChrW(1502) & ChrW(1513) & ChrW(1500) & " 77): " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & " " & ChrW(1513) & ChrW(1492) & ChrW(1513) & ChrW(1501) & " " & ChrW(1513) & ChrW(1500) & ChrW(1493) & " " & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1500) & " 77"'
new_inst1 = 'wsSearch.Cells(6, 5).Value = "- " & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513) & " " & ChrW(1512) & ChrW(1490) & ChrW(1497) & ChrW(1500) & " (" & ChrW(1500) & ChrW(1502) & ChrW(1513) & ChrW(1500) & " 77): " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & " " & ChrW(1513) & ChrW(1492) & ChrW(1502) & ChrW(1497) & ChrW(1500) & ChrW(1492) & " 77 " & ChrW(1502) & ChrW(1493) & ChrW(1508) & ChrW(1497) & ChrW(1506) & ChrW(1514) & " " & ChrW(1489) & ChrW(1513) & ChrW(1501) & " " & ChrW(1513) & ChrW(1500) & ChrW(1493) & " " & ChrW(1489) & ChrW(1513) & ChrW(1500) & ChrW(1502) & ChrW(1493) & ChrW(1514)'

if target_inst1 in content:
    content = content.replace(target_inst1, new_inst1)

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_143"', 'Attribute VB_Name = "Goren_Claude_V2_144"')
content = content.replace('VERSION: V2.143', 'VERSION: V2.144')
content = content.replace('APP_VERSION As String = "2.143"', 'APP_VERSION As String = "2.144"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.144 created.")
