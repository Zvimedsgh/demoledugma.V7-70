import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.124.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for i, line in enumerate(lines):
    if "For r = 2 To lastRow" in line:
        new_lines.append("""    Dim bWildcard As Boolean\n""")
        new_lines.append("""    bWildcard = (InStr(1, searchText, "*") > 0 Or InStr(1, searchText, "?") > 0)\n\n""")
        new_lines.append(line)
    elif "If InStr(1, cName, searchText, vbTextCompare) > 0 Then" in line:
        new_lines.append("""                Dim bMatch As Boolean\n""")
        new_lines.append("""                bMatch = False\n""")
        new_lines.append("""                If bWildcard Then\n""")
        new_lines.append("""                    If UCase(cName) Like UCase(searchText) Then bMatch = True\n""")
        new_lines.append("""                Else\n""")
        new_lines.append("""                    If UCase(cName) = UCase(searchText) Then bMatch = True\n""")
        new_lines.append("""                End If\n""")
        new_lines.append("""                If bMatch Then\n""")
    else:
        new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.125.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Robust replace done.")
