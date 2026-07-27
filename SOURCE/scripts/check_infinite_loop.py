import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.082.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Public Sub CheckUserPermissions()" in line_strip:
        in_func = True
    elif in_func and "End Sub" in line_strip:
        for j in range(i-25, i+1):
            if j >= 0:
                print(f"[{j+1}] {lines[j].strip()}")
        break
    elif in_func:
        pass

