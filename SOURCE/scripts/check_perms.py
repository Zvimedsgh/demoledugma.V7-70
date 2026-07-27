import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.072.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    line_strip = line.strip()
    if "Private Sub CheckUserPermissions()" in line_strip or "Sub CheckUserPermissions()" in line_strip:
        in_func = True
        print(f"[{i+1}] {line_strip}")
    elif in_func and "End Sub" in line_strip:
        print(f"[{i+1}] {line_strip}")
        break
    elif in_func:
        print(f"[{i+1}] {line_strip}")

