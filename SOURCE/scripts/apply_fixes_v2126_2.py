import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.126.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if 'wsMain.Shapes("shpDemoLockG3G4").Delete' in line and i > 7400:
        new_lines.append(line)
        new_lines.append('wsMain.Shapes("shpDemoMsgText").Delete\n')
    else:
        new_lines.append(line)

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Fixed Else delete.")
