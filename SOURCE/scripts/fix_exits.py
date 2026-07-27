import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.068.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.069.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_069"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.069\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.069"\n')
    elif "490     If cnt = 0 Then MsgBoxU " in line and "Exit Sub" in line:
        new_lines.append(line.replace("Exit Sub", "Application.EnableEvents = True: Exit Sub"))
    elif "If MsgBoxU(alreadyMsg3, vbYesNo + vbQuestion) <> vbYes Then" in "".join(lines[max(0, i-2):i]) and line.strip() == "Exit Sub":
        new_lines.append("        Application.EnableEvents = True: Exit Sub\n")
    elif "120         Exit Sub" in line and "BuildPresentation" in "".join(lines[max(0, i-200):i]):
        new_lines.append("    120         Application.EnableEvents = True: Exit Sub\n")
    elif "1060    Exit Sub" in line and "BuildPresentation" in "".join(lines[max(0, i-50):i]):
        new_lines.append("    1060    Application.EnableEvents = True: Exit Sub\n")
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.069")

