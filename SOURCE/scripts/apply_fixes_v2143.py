import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.142.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Revert exact match back to InStr
target_exact = """If bWildcard Then
If UCase(cName) Like UCase(searchText) Then bMatch = True
Else
If UCase(cName) = UCase(searchText) Then bMatch = True
End If"""

new_instr = """If bWildcard Then
If UCase(cName) Like UCase(searchText) Then bMatch = True
Else
If InStr(1, cName, searchText, vbTextCompare) > 0 Then bMatch = True
End If"""

if target_exact in content:
    content = content.replace(target_exact, new_instr)
    print("Reverted exact match back to InStr.")
else:
    print("Could not find exact match block.")

# Update instructions
target_inst1 = 'wsSearch.Cells(6, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1512) & ChrW(1511) & " 77"'
new_inst1 = 'wsSearch.Cells(6, 5).Value = "- " & ChrW(1495) & ChrW(1497) & ChrW(1508) & ChrW(1493) & ChrW(1513) & " " & ChrW(1512) & ChrW(1490) & ChrW(1497) & ChrW(1500) & " (" & ChrW(1500) & ChrW(1502) & ChrW(1513) & ChrW(1500) & " 77): " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1500) & ChrW(1511) & ChrW(1493) & ChrW(1495) & " " & ChrW(1513) & ChrW(1492) & ChrW(1513) & ChrW(1501) & " " & ChrW(1513) & ChrW(1500) & ChrW(1493) & " " & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1500) & " 77" \' "- חיפוש רגיל (למשל 77): ימצא כל לקוח שהשם שלו מכיל 77"'

target_inst2 = 'wsSearch.Cells(7, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " *77*: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1502) & ChrW(1497) & " " & ChrW(1513) & ChrW(1502) & ChrW(1499) & ChrW(1497) & ChrW(1500) & " 77"'
new_inst2 = 'wsSearch.Cells(7, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " *77: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1502) & ChrW(1497) & " " & ChrW(1513) & ChrW(1502) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & "-77" \' "- הקלד *77: ימצא כל מי שמתחיל ב-77"'

target_inst3 = 'wsSearch.Cells(8, 5).Value = "- " & ChrW(1492) & ChrW(1511) & ChrW(1500) & ChrW(1491) & " 77*: " & ChrW(1497) & ChrW(1502) & ChrW(1510) & ChrW(1488) & " " & ChrW(1499) & ChrW(1500) & " " & ChrW(1502) & ChrW(1497) & " " & ChrW(1513) & ChrW(1502) & ChrW(1514) & ChrW(1495) & ChrW(1497) & ChrW(1500) & " " & ChrW(1489) & "-77"'
new_inst3 = 'wsSearch.Cells(8, 5).Value = ""'

content = content.replace(target_inst1, new_inst1)
content = content.replace(target_inst2, new_inst2)
content = content.replace(target_inst3, new_inst3)

content = content.replace('Attribute VB_Name = "Goren_Claude_V2_142"', 'Attribute VB_Name = "Goren_Claude_V2_143"')
content = content.replace('VERSION: V2.142', 'VERSION: V2.143')
content = content.replace('APP_VERSION As String = "2.142"', 'APP_VERSION As String = "2.143"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.143.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("V2.143 created.")
