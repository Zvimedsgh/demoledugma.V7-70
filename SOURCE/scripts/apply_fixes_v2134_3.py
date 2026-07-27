import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.134.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "If UCase(cName) = UCase(searchText) Then bMatch = True" in line:
        lines[i] = "If InStr(1, cName, searchText, vbTextCompare) > 0 Then bMatch = True\n"
        print("Replaced exact match with InStr.")

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(lines)

