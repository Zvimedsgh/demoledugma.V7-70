import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Function FindSourceFile" in line:
        in_func = True
    if in_func:
        print(f"[{i+1}] {lines[i].strip()}")
        if "End Function" in line:
            in_func = False
            break

