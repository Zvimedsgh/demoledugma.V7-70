import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.196.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_demo = False
for i, line in enumerate(lines):
    if "Public Sub ApplyDemoLockOnOpen" in line:
        in_demo = True
    if in_demo:
        print(f"[{i+1}] {line.strip()}")
    if in_demo and "End Sub" in line:
        break

