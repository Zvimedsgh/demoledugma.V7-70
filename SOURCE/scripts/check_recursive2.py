import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.144.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub BuildReview(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 1000: break
        if "BuildReview" in line and i > 626:
            print(f"[{i+1}] {line.strip()}")
        if "End Sub" in line:
            break

