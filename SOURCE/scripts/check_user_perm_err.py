import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.108.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub CheckUserPermissions()" in line:
        in_func = True
    elif in_func and "End Sub" in line:
        break
    elif in_func and "Exit Sub" in line:
        for j in range(max(0, i-5), min(len(lines), i+6)):
            print(f"[{j+1}] {lines[j].strip()}")
        print("-" * 20)

