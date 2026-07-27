import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.139.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub SaveReportsToFolder()" in line:
        in_func = True
    if in_func:
        if "For hiErr = 1 To hiddenCount" in line or "For" in line and "hiddenCount" in line and "hiErr" not in line:
            for j in range(max(0, i-5), min(len(lines), i+10)):
                print(f"[{j+1}] {lines[j].strip()}")

