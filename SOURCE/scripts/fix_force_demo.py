import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.067.bas'
outpath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.068.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "Attribute VB_Name =" in line:
        new_lines.append('Attribute VB_Name = "Goren_Claude_V2_068"\n')
    elif "VERSION: " in line:
        new_lines.append("' VERSION: V2.068\n")
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append('Private Const APP_VERSION As String = "2.068"\n')
    elif "Private Const FORCE_DEMO_MODE As Boolean = True" in line:
        new_lines.append(line.replace("True", "False"))
    else:
        new_lines.append(line)

with open(outpath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.068")

