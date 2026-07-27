import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.139.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub SaveReportsToFolder()" in line:
        in_func = True
    if in_func:
        if "hiddenCount" in line:
            for j in range(max(0, i-2), min(len(lines), i+15)):
                print(f"[{j+1}] {lines[j].strip()}")
            break

