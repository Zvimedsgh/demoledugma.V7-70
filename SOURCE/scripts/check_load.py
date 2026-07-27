import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.193.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_load = False
for i, line in enumerate(lines):
    if "Public Sub LoadCheckedFields" in line:
        in_load = True
    if in_load:
        print(f"[{i+1}] {line.strip()}")
    if in_load and "End Sub" in line:
        break

