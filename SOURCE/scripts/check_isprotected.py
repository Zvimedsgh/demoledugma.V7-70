import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.150.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Private Function IsProtectedSheet(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 50: break
        print(f"[{i+1}] {line.strip()}")
        if "End Function" in line:
            break

