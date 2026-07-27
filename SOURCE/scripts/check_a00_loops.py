import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.092.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet()" in line:
        in_func = True
        print(f"[{i+1}] {line.strip()}")
    elif in_func and "End Sub" in line:
        print(f"[{i+1}] {line.strip()}")
        break
    elif in_func:
        if "For " in line or "While " in line or "Do " in line or "Loop" in line:
            print(f"[{i+1}] {line.strip()}")

