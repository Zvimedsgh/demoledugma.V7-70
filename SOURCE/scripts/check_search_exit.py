import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.137.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub SearchClientName()" in line:
        in_func = True
    if in_func:
        if "Exit Sub" in line or "Application.ScreenUpdating" in line or "Application.EnableEvents" in line:
            print(f"[{i+1}] {lines[i].strip()}")
        if "End Sub" in line:
            break

