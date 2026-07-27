import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.076.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.077.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_077"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.077\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.077"\n')
    elif "Next ws" in line and "On Error GoTo 0" in lines[i+1] and "End If" in lines[i-1] and "ws.Visible = xlSheetVisible" in lines[i-2]:
        new_lines.append("        End If\n")
        new_lines.append(line)
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("V2.077 created with missing End If fixed.")

