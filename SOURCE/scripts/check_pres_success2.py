import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.059.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub BuildPresentation()" in line_strip:
        in_func = True
    elif in_func and "End Sub" in line_strip:
        break
    elif in_func and "bSaved" in line_strip:
        for j in range(i-5, i+15):
            print(f"[{j+1}] {lines[j].strip()}")
        break

