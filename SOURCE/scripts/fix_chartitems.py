import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.053_LNUM.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    new_lines.append(line)
    if "chartItems = nItems" in line:
        new_lines.append("    If chartItems = 0 Then\n")
        new_lines.append("        Application.ScreenUpdating = True\n")
        new_lines.append("        Exit Sub\n")
        new_lines.append("    End If\n")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.054.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.054")
