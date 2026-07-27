import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.045_20260702_1605.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_sub = False
sub_lines = []
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet()" in line:
        in_sub = True
    if in_sub:
        sub_lines.append(f"[{i}] {line}")
    if in_sub and "End Sub" in line:
        break

with open(r'c:\LEVAV PROJECT\SOURCE\A00_SetupMainSheet.txt', 'w', encoding='utf-8') as f:
    f.writelines(sub_lines)
