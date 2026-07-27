import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.119.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub ApplyDemoLockOnOpen" in line:
        in_func = True
    if in_func:
        print(f"[{i+1}] {line.strip()}")
        if "End Sub" in line:
            break

