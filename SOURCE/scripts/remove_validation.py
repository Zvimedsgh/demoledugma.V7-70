import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.075.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.076.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
skip = False
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_076"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.076\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.076"\n')
    elif "1680 On Error Resume Next" in line:
        skip = True
        new_lines.append("        ' Data Validation block removed to avoid Excel bug\n")
    elif skip and "1731 On Error GoTo ERR_HANDLER" in line:
        skip = False
    elif not skip:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("V2.076 created with validation block entirely removed.")

