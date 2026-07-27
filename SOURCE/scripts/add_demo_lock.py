import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.041.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

new_lines = []
for i, line in enumerate(lines):
    if "Public Sub ViewReportsFolder()" in line:
        new_lines.append(line)
        new_lines.append("If CheckDemoLock() Then Exit Sub\n")
    else:
        new_lines.append(line)

with open(filepath, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("Added Demo Lock to ViewReportsFolder")
