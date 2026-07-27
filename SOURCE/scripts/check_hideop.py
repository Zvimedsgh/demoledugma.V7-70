import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.194.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_hideop = False
for i, line in enumerate(lines):
    if "Public Sub HideOpInst()" in line:
        in_hideop = True
    if in_hideop:
        print(f"[{i+1}] {line.strip()}")
    if in_hideop and "End Sub" in line:
        break

