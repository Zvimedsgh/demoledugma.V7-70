import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.036_20260702_1257.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Private Sub LoadCorrectionsToDicts" in line:
        in_func = True
    if in_func and "End Sub" in line:
        print(f"[{i}] {lines[i].strip()}")
        break
    if in_func:
        print(f"[{i}] {lines[i].strip()}")
