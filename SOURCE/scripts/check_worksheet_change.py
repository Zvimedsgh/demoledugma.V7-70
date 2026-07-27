import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.196.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_sub = False
for i, line in enumerate(lines):
    if "Private Sub Worksheet_Change" in line:
        in_sub = True
    if in_sub:
        print(f"[{i+1}] {line.strip()}")
    if in_sub and "End Sub" in line:
        break

