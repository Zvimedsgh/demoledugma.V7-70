import sys

filepath = r'c:\LEVAV PROJECT\SOURCE\Goren_Claude_V2.038_20260702_1440.bas'

with open(filepath, 'r', encoding='utf-8') as f:
    lines = f.readlines()

in_func = False
for i, line in enumerate(lines):
    if "Public Sub A00_SetupMainSheet" in line:
        in_func = True
    if in_func and "End Sub" in line:
        break
    if in_func and ("AddShape" in line or "Name =" in line):
        if i > 3600 and i < 3800:
            print(f"[{i}] {line.strip()}")
