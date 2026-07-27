import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.061.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.062.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for line in lines:
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_062"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.062\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.062"\n')
    elif 'If demoParam = ChrW(1499) & ChrW(1503) Or demoParam = "YES" Then isDemoMode = True' in line:
        new_lines.append('        If Not FORCE_DEMO_MODE Then\n')
        new_lines.append('            ' + line.strip() + '\n')
    elif 'If demoParam = ChrW(1500) & ChrW(1488) Or demoParam = "NO" Then isDemoMode = False' in line:
        new_lines.append('            ' + line.strip() + '\n')
        new_lines.append('        End If\n')
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.062")

