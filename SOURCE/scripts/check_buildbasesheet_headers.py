import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.153.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Private Sub BuildBaseSheet(" in line:
        in_func = True
    if in_func:
        if i - lines.index(line) > 500: break
        if "Dim headers" in line or "Array(" in line:
            for j in range(i, i+5):
                print(f"[{j+1}] {lines[j].strip()}")
            break
        if "End Sub" in line:
            break

