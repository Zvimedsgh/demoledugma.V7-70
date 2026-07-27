import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.041_20260702_1520.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub ResetHomeDefaults()" in line:
        in_func = True
    if in_func and "End Sub" in line:
        for j in range(i-30, i+2):
            print(f"[{j}] {lines[j].strip()}")
        break
