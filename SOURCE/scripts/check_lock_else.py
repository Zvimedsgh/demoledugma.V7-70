import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.126.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub ApplyDemoLockOnOpen" in line:
        in_func = True
    if in_func:
        if "shpDemoLockG3G4" in line:
            for j in range(max(0, i-2), min(len(lines), i+6)):
                print(f"[{j+1}] {lines[j].strip()}")
            print("-" * 20)
        if "End Sub" in line:
            break

