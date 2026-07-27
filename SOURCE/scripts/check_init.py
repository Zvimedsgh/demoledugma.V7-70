import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.195.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_init = False
for i, line in enumerate(lines):
    if "Private Sub InitRawColumns" in line:
        in_init = True
    if in_init:
        print(f"[{i+1}] {line.strip()}")
    if in_init and "End Sub" in line:
        break

