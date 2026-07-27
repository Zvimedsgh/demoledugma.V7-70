import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.127.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub SearchClientName()" in line:
        in_func = True
    if in_func:
        if "Add ""Search"" button" in line:
            for j in range(i, i+30):
                print(f"[{j+1}] {lines[j].strip()}")
            break

