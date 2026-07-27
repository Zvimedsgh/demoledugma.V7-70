import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.047_20260702_1625.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "Private Const FORCE_DEMO_MODE As Boolean = True" in line:
        new_lines.append(line.replace("True", "False"))
    elif "Attribute VB_Name =" in line:
        new_lines.append(line.replace("V2_047", "V2_048"))
    elif "Private Const APP_VERSION As String =" in line:
        new_lines.append(line.replace("2.047", "2.048"))
    elif "VERSION: V2.047" in line:
        new_lines.append(line.replace("2.047", "2.048"))
    else:
        new_lines.append(line)

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.048.bas', 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Created V2.048 with FORCE_DEMO_MODE = False")
