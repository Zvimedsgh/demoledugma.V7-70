import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.124.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
in_do_search = False

for i, line in enumerate(lines):
    if "Public Sub DoClientSearch()" in line:
        in_do_search = True
    if "End Sub" in line and in_do_search:
        in_do_search = False
        
    if in_do_search and "For r = 2 To lastRow" in line:
        new_lines.append("""    Dim bWildcard As Boolean\n""")
        new_lines.append("""    bWildcard = (InStr(1, searchText, "*") > 0 Or InStr(1, searchText, "?") > 0)\n\n""")
        new_lines.append(line)
    elif in_do_search and "If InStr(1, cName, searchText, vbTextCompare) > 0 Then" in line:
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

content = "".join(new_lines)
content = content.replace('Attribute VB_Name = "Goren_Claude_V2_124"', 'Attribute VB_Name = "Goren_Claude_V2_125"')
content = content.replace('VERSION: V2.124', 'VERSION: V2.125')
content = content.replace('APP_VERSION As String = "2.124"', 'APP_VERSION As String = "2.125"')

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.125.bas', 'w', encoding='utf-8') as f:
    f.write(content)

print("Proper robust replace done.")
