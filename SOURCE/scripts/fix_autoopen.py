import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.060.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.061.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_061"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.061\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.061"\n')
    elif "Public Sub CheckUserPermissions()" in line:
        new_lines.append(line)
        new_lines.append("    ' --- Failsafe: Reset Application state --- \n")
        new_lines.append("    Application.EnableEvents = True\n")
        new_lines.append("    Application.ScreenUpdating = True\n")
        new_lines.append("    Application.Calculation = xlCalculationAutomatic\n")
        new_lines.append("    ' ----------------------------------------- \n")
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.061")

