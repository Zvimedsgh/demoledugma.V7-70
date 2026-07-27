import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.059.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for i, line in enumerate(lines):
    if "Private Sub LockDemoUI(" in line or "Public Sub ApplyDemoLockOnOpen(" in line:
        for j in range(i, i+40):
            print(f"[{j+1}] {lines[j].strip()}")
        break

