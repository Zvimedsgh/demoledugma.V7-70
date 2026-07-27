import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.081.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.082.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_082"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.082\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.082"\n')
    elif "Public Sub FixExcel()" in line and i < 7900: # skip first one
        skip = True
        new_lines.append("        ' Old FixExcel removed\n")
    elif skip and "End Sub" in line:
        skip = False
    elif not skip:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("V2.082 created with single FixExcel.")

