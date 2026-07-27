import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.034.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "shtHeaders = Array" in line:
        in_func = True
    if in_func and "End Sub" in line:
        in_func = False
    
    if in_func:
        print(f"[{i}] {line.strip()}")
        if i > 3420: break
