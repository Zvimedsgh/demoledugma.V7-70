import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.138.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub BuildPresentation()" in line:
        in_func = True
    if in_func:
        if "End Sub" in line:
            for j in range(max(0, i-20), i+1):
                print(f"[{j+1}] {lines[j].strip()}")
            break

