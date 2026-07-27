import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.194.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_check = False
for i, line in enumerate(lines):
    if "Public Sub CheckUserPermissions()" in line:
        in_check = True
    if in_check:
        print(f"[{i+1}] {line.strip()}")
    if in_check and "End Sub" in line:
        break

