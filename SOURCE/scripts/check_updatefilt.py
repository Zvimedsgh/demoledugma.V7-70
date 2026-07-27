import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.192.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_update = False
for i, line in enumerate(lines):
    if "Public Sub UpdateFilterValueDropdown" in line:
        in_update = True
    if in_update:
        print(f"[{i+1}] {line.strip()}")
    if in_update and "End Sub" in line:
        break

