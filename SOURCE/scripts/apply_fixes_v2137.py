import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.136.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Attribute VB_Name" in line:
        lines[i] = 'Attribute VB_Name = "Goren_Claude_V2_137"\n'
    elif "' VERSION:" in line:
        lines[i] = "' VERSION: V2.137\n"
    elif "Private Const APP_VERSION As String =" in line:
        lines[i] = 'Private Const APP_VERSION As String = "2.137"\n'

with open(r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.137.bas', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("V2.137 created and versions updated properly.")
