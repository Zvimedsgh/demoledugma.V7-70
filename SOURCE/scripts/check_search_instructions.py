import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.142.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub SearchClientName()" in line:
        in_func = True
    if in_func:
        if "wsSearch.Cells(5, 5).Value =" in line:
            for j in range(i, i+5):
                print(f"[{j+1}] {lines[j].strip()}")
            break

