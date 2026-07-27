import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.123.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Sub ResetHomeDefaults()" in line:
        in_func = True
    if in_func:
        if "End Sub" in line:
            for j in range(max(0, i-25), i+2):
                print(f"[{j+1}] {lines[j].strip()}")
            break

