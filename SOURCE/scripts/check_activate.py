import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.137.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub SearchClientName()" in line:
        in_func = True
    if in_func:
        if "wsSearch.Activate" in line:
            for j in range(max(0, i-5), min(len(lines), i+5)):
                print(f"[{j+1}] {lines[j].strip()}")
            break

