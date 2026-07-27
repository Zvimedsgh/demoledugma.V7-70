import sys
import re

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.078.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.079.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_079"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.079\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.079"\n')
    elif 'MsgBoxU "Error in SetupMainSheet!"' in line:
        new_lines.append(line.replace("Error in SetupMainSheet!", "Error in V2.079! (If you don't see this, you are running old code!)"))
    elif "Public Sub A00_SetupMainSheet_V2078()" in line:
        new_lines.append(line.replace("A00_SetupMainSheet_V2078", "A00_SetupMainSheet_V2079"))
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("V2.079 created with updated error message.")

