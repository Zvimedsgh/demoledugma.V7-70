import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.054.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for line in lines:
    if "tmpWs.ChartObjects.Delete" in line:
        new_lines.append("        On Error Resume Next\n")
        new_lines.append(line)
        new_lines.append("        On Error GoTo ERR_HANDLER\n")
    else:
        new_lines.append(line)

# Bump version to 2.055
for i, line in enumerate(new_lines):
    if "Attribute VB_Name =" in line:
        new_lines[i] = line.replace("V2_054", "V2_055")
    if "Private Const APP_VERSION As String =" in line:
        new_lines[i] = line.replace("2.054", "2.055")
    if "VERSION: V2.054" in line:
        new_lines[i] = line.replace("2.054", "2.055")

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.055.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.055")
