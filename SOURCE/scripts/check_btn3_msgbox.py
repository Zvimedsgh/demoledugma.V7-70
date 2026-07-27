import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.155.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub BuildPresentation(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 1000: break
        if "Yes" in line and "No" in line and "Cancel" in line:
            for j in range(i-5, i+5):
                print(f"[{j+1}] {lines[j].strip()}")
            break

