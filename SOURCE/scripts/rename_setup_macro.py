import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.077.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.078.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_078"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.078\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.078"\n')
    elif "Public Sub A00_SetupMainSheet()" in line:
        new_lines.append(line.replace("A00_SetupMainSheet", "A00_SetupMainSheet_V2078"))
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("V2.078 created with renamed Setup macro.")

