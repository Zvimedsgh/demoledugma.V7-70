import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.199.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if 'wsMain.Range("rngFilterValue").Value = selectText' in line:
        lines[i] = '        On Error Resume Next\n' + line + '        On Error GoTo ERR_HANDLER\n'

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Updated V2.199")
